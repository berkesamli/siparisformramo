import { NextResponse } from "next/server";
import { createHmac, timingSafeEqual } from "node:crypto";
import { bekleyenAl, bekleyenSil } from "@/lib/wa-bekleyen";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Meta WhatsApp webhook'u.
//   GET  — Meta'nın doğrulama isteği (hub.verify_token = WHATSAPP_VERIFY_TOKEN)
//   POST — teslim durumları: müşteriye giden fiş "failed" olursa SMS'e düşülür.
// Gelen müşteri mesajları (messages) şimdilik yalnızca loglanır.
//
// Güvenlik: WHATSAPP_APP_SECRET tanımlıysa X-Hub-Signature-256 doğrulanır.
// Tanımlı değilse yalnızca bizim yazdığımız (tahmin edilemez) wamid'lere
// ait kayıtlar işlenir; sahte bir istek en fazla boşa döner.

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

function imzaGecerli(raw: string, header: string | null): boolean {
  const secret = (process.env.WHATSAPP_APP_SECRET || "").trim();
  if (!secret) return true; // imza zorunlu değil (bkz. üstteki not)
  if (!header || !header.startsWith("sha256=")) return false;
  const beklenen = createHmac("sha256", secret).update(raw, "utf8").digest("hex");
  const gelen = header.slice(7);
  if (gelen.length !== beklenen.length) return false;
  return timingSafeEqual(Buffer.from(gelen, "hex"), Buffer.from(beklenen, "hex"));
}

export async function POST(req: Request) {
  const raw = await req.text();
  if (!imzaGecerli(raw, req.headers.get("x-hub-signature-256"))) {
    return NextResponse.json({ ok: false, error: "İmza geçersiz." }, { status: 401 });
  }
  let body: { entry?: { changes?: { field?: string; value?: { statuses?: Durum[]; messages?: unknown[] } }[] }[] } | null = null;
  try { body = JSON.parse(raw); } catch { return NextResponse.json({ ok: false }, { status: 400 }); }

  let smsGonderilen = 0;
  for (const entry of body?.entry || []) {
    for (const ch of entry.changes || []) {
      const v = ch.value;
      if (!v) continue;
      if (v.messages?.length) console.log(`WhatsApp: ${v.messages.length} gelen mesaj (yanıtlanmıyor).`);
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
