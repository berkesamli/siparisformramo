import { NextResponse } from "next/server";
import { createHmac, timingSafeEqual } from "node:crypto";
import { bekleyenAl, bekleyenSil } from "@/lib/wa-bekleyen";
import { dbConfigured } from "@/lib/mesaj/db";
import { izKaydet } from "@/lib/mesaj/webhook-iz";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Meta WhatsApp webhook'u.
//   GET  — Meta'nın doğrulama isteği (hub.verify_token = WHATSAPP_VERIFY_TOKEN)
//   POST — teslim durumları: müşteriye giden fiş "failed" olursa SMS'e düşülür.
// Gelen müşteri mesajları (messages) DATABASE_URL tanımlıysa gelen kutusuna
// (lib/mesaj) yazılır; giden mesajların teslim/okundu durumu da oraya işlenir.
//
// Güvenlik: POST istekleri X-Hub-Signature-256 ile WHATSAPP_APP_SECRET (ya da
// META_APP_SECRET) üzerinden doğrulanır. Gizli anahtar tanımlı değilse HİÇBİR
// POST işlenmez (401): gelen mesajlar doğrudan gelen kutusuna yazıldığı için
// imzasız istek kabul etmek sahte müşteri mesajı enjekte edilmesine yol açar.

export async function GET(req: Request) {
  const url = new URL(req.url);
  const mode = url.searchParams.get("hub.mode");
  const token = url.searchParams.get("hub.verify_token");
  const challenge = url.searchParams.get("hub.challenge") || "";
  const beklenen = (process.env.WHATSAPP_VERIFY_TOKEN || "").trim();
  if (mode === "subscribe" && beklenen && token === beklenen) {
    return new NextResponse(challenge, { status: 200, headers: { "Content-Type": "text/plain" } });
  }
  return NextResponse.json({ ok: false, error: "Doğrulama başarısız." }, { status: 403 });
}

interface Durum {
  id?: string;
  status?: string;
  recipient_id?: string;
  errors?: { code?: number; title?: string; message?: string; error_data?: { details?: string } }[];
}

const appSecret = () => (process.env.WHATSAPP_APP_SECRET || process.env.META_APP_SECRET || "").trim();

function imzaGecerli(raw: string, header: string | null): boolean {
  const secret = appSecret();
  if (!secret) return false; // gizli anahtar girilmeden POST kabul edilmez (fail-closed)
  if (!header || !/^sha256=[0-9a-f]{64}$/i.test(header)) return false;
  const beklenen = createHmac("sha256", secret).update(raw, "utf8").digest();
  return timingSafeEqual(Buffer.from(header.slice(7), "hex"), beklenen);
}

export async function POST(req: Request) {
  const raw = await req.text();
  if (!imzaGecerli(raw, req.headers.get("x-hub-signature-256"))) {
    await izKaydet("whatsapp", "imza-red", appSecret()
      ? "Meta'dan olay geldi ama imza doğrulanamadı: Vercel'deki WHATSAPP_APP_SECRET, Meta uygulamasının App Secret'ıyla aynı değil."
      : "Meta'dan olay geldi ama WHATSAPP_APP_SECRET tanımlı olmadığı için reddedildi. Meta → App settings → Basic → App secret değerini Vercel'e girin.");
    return NextResponse.json({ ok: false, error: "İmza geçersiz." }, { status: 401 });
  }
  let body: { entry?: { changes?: { field?: string; value?: { statuses?: Durum[]; messages?: unknown[] } }[] }[] } | null = null;
  try { body = JSON.parse(raw); } catch { return NextResponse.json({ ok: false }, { status: 400 }); }

  // Kurulum teşhisi için son olayın özeti (mesaj / durum sayısı, alan adı)
  {
    let mesajSayisi = 0, durumSayisi = 0; const alanlar = new Set<string>();
    for (const e of body?.entry || []) for (const ch of e.changes || []) { alanlar.add(ch.field || "?"); mesajSayisi += ch.value?.messages?.length || 0; durumSayisi += ch.value?.statuses?.length || 0; }
    await izKaydet("whatsapp", mesajSayisi ? "mesaj" : durumSayisi ? "durum" : "diger", `alan: ${[...alanlar].join(",") || "-"} · ${mesajSayisi} mesaj · ${durumSayisi} durum${dbConfigured() ? "" : " · gelen kutusu kurulu değil (DATABASE_URL yok)"}`);
  }

  let smsGonderilen = 0;
  for (const entry of body?.entry || []) {
    for (const ch of entry.changes || []) {
      const v = ch.value;
      if (!v) continue;
      if (v.messages?.length) {
        if (dbConfigured()) {
          try {
            const { gelenWhatsapp } = await import("@/lib/mesaj/whatsapp");
            await gelenWhatsapp(v as Parameters<typeof gelenWhatsapp>[0]);
          } catch (err) {
            console.error("WhatsApp gelen mesaj kaydedilemedi:", err);
          }
        } else {
          console.log(`WhatsApp: ${v.messages.length} gelen mesaj (gelen kutusu kurulu değil).`);
        }
      }
      if (v.statuses?.length && dbConfigured()) {
        try {
          const { whatsappDurumlar } = await import("@/lib/mesaj/whatsapp");
          await whatsappDurumlar(v.statuses as Parameters<typeof whatsappDurumlar>[0]);
        } catch (err) {
          console.error("WhatsApp durum güncellenemedi:", err);
        }
      }
      for (const st of v.statuses || []) {
        if (!st.id) continue;
        if (st.status === "delivered" || st.status === "read") {
          await bekleyenSil(st.id);
          continue;
        }
        if (st.status !== "failed") continue; // "sent" → sunucuya ulaştı, bekle
        const kayit = await bekleyenAl(st.id);
        if (!kayit) continue; // bizim takip ettiğimiz bir mesaj değil (örn. patron fişi)
        const sebep = (st.errors || []).map((e) => `${e.title || e.message || "hata"} (kod ${e.code ?? "?"})`).join("; ") || "teslim edilemedi";
        console.warn(`WhatsApp teslim edilemedi → SMS: ${kayit.orderId} ${kayit.telefon} — ${sebep}`);
        try {
          const { smsConfigured, sendSms } = await import("@/lib/sms");
          if (smsConfigured()) {
            const r = await sendSms([kayit.telefon], kayit.smsMetni);
            if (r.ok) smsGonderilen++;
            const { saveSmsRecord, newSmsId, istanbulDateKey } = await import("@/lib/sms-log");
            const t = new Date();
            await saveSmsRecord({
              id: newSmsId(t),
              dateKey: istanbulDateKey(t),
              createdAt: t.toISOString(),
              sender: "Sistem (WhatsApp yedeği)",
              message: kayit.smsMetni,
              recipients: r.sent,
              segments: 1,
              credits: r.sent.length,
              ok: r.ok,
              jobId: r.jobId,
              error: r.error ? `${r.error} — WhatsApp: ${sebep}` : `WhatsApp: ${sebep}`,
              iysfilter: "0",
            }).catch(() => {});
          } else {
            console.error("SMS yapılandırılmamış; WhatsApp yedeği gönderilemedi:", kayit.orderId);
          }
        } catch (err) {
          console.error("WhatsApp yedeği SMS gönderilemedi:", err);
        }
        await bekleyenSil(st.id);
      }
    }
  }
  return NextResponse.json({ ok: true, smsGonderilen });
}
