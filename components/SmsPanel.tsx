"use client";

// SMS gönderim ekranı — müşteri defterinden çoklu alıcı seçimi, hazır şablonlar,
// canlı kredi sayacı ve gönderim geçmişi.

import { useEffect, useMemo, useState } from "react";
import { normalizePhone, smsSegments, stripTurkish } from "@/lib/sms-format";
import { eslesir } from "@/lib/search-norm";

interface Customer {
  id: string;
  firstName: string;
  lastName: string;
  company: string;
  phone: string;
  city: string;
}

interface SmsRecord {
  id: string;
  createdAt: string;
  sender: string;
  message: string;
  recipients: string[];
  credits: number;
  ok: boolean;
  error?: string;
}

const SABLONLAR: { ad: string; metin: string }[] = [
  {
    ad: "Kargonuz Çıktı",
    metin:
      "Sayin musterimiz, siparisiniz kargoya verilmistir. Olga Cerceve",
  },
  {
    ad: "Siparişiniz Hazır",
    metin:
      "Sayin musterimiz, siparisiniz hazirdir. Teslim alabilirsiniz. Olga Cerceve",
  },
  {
    ad: "Ödeme Hatırlatma",
    metin:
      "Sayin musterimiz, vadesi gelen bakiyeniz bulunmaktadir. Bilginize. Olga Cerceve",
  },
];

function title(c: Customer): string {
  const kisi = `${c.firstName || ""} ${c.lastName || ""}`.trim();
  if (c.company && kisi) return `${c.company} — ${kisi}`;
  return c.company || kisi || "-";
}

export default function SmsPanel() {
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [records, setRecords] = useState<SmsRecord[]>([]);
  const [configured, setConfigured] = useState(true);
  const [yukleniyor, setYukleniyor] = useState(true);

  const [secili, setSecili] = useState<Set<string>>(new Set());
  const [elle, setElle] = useState("");
  const [mesaj, setMesaj] = useState("");
  const [ara, setAra] = useState("");

  // İYS filtresi. Varsayılan "0" (bilgilendirme) — kargo/sipariş mesajları
  // ticari ileti sayılmaz. Kampanya için 11/12 seçilmeli, aksi hâlde mevzuata
  // aykırı gönderim yapılmış olur.
  const [iysfilter, setIysfilter] = useState<"0" | "11" | "12">("0");
  const [gonderiliyor, setGonderiliyor] = useState(false);
  const [sonuc, setSonuc] = useState<{ ok: boolean; text: string } | null>(null);

  useEffect(() => {
    Promise.all([
      fetch("/api/musteriler").then((r) => r.json()).catch(() => ({})),
      fetch("/api/sms").then((r) => r.json()).catch(() => ({})),
    ])
      .then(([m, s]) => {
        setCustomers(Array.isArray(m?.customers) ? m.customers : []);
        if (s?.records) setRecords(s.records);
        if (typeof s?.configured === "boolean") setConfigured(s.configured);
      })
      .finally(() => setYukleniyor(false));
  }, []);

  // Telefonu olmayan müşteriye SMS atılamaz; listeye hiç almıyoruz.
  const telefonlu = useMemo(
    () => customers.filter((c) => normalizePhone(c.phone)),
    [customers]
  );

  const listelenen = useMemo(
    () =>
      !ara.trim()
        ? telefonlu
        : telefonlu.filter((c) =>
            eslesir(ara, title(c), c.phone, c.city)
          ),
    [telefonlu, ara]
  );

  // Elle girilen numaralar: virgül, boşluk veya satır sonuyla ayrılabilir.
  const elleNumaralar = useMemo(
    () => elle.split(/[\s,;]+/).map((s) => s.trim()).filter(Boolean),
    [elle]
  );

  const tumAlicilar = useMemo(() => {
    const out: string[] = [];
    for (const c of telefonlu) if (secili.has(c.id)) out.push(c.phone);
    out.push(...elleNumaralar);
    return out;
  }, [telefonlu, secili, elleNumaralar]);

  const gecerli = tumAlicilar.filter((n) => normalizePhone(n));
  const gecersiz = tumAlicilar.filter((n) => !normalizePhone(n));
  // Aynı numara hem defterden hem elle girilmiş olabilir — kredi hesabı
  // sunucudaki tekilleştirmeyle aynı olsun diye burada da tekilleştiriyoruz.
  const tekilSayi = new Set(gecerli.map((n) => normalizePhone(n))).size;

  const sayim = smsSegments(mesaj);
  const kredi = sayim.segments * tekilSayi;

  function toggle(id: string) {
    setSecili((s) => {
      const n = new Set(s);
      if (n.has(id)) n.delete(id);
      else n.add(id);
      return n;
    });
  }

  async function gonder() {
    if (!tekilSayi || !mesaj.trim() || gonderiliyor) return;
    const onay = window.confirm(
      `${tekilSayi} alıcıya gönderilecek, ${kredi} SMS kredisi harcanacak.\n\nOnaylıyor musunuz?`
    );
    if (!onay) return;

    setGonderiliyor(true);
    setSonuc(null);
    try {
      const r = await fetch("/api/sms", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ numbers: gecerli, message: mesaj, iysfilter }),
      });
      const d = await r.json();
      if (d.ok) {
        setSonuc({
          ok: true,
          text: `${d.sent} alıcıya gönderildi (${d.credits} kredi).`,
        });
        setMesaj("");
        setSecili(new Set());
        setElle("");
        fetch("/api/sms")
          .then((x) => x.json())
          .then((s) => s?.records && setRecords(s.records))
          .catch(() => {});
      } else {
        setSonuc({ ok: false, text: d.error || "Gönderilemedi." });
      }
    } catch {
      setSonuc({ ok: false, text: "Sunucuya ulaşılamadı." });
    } finally {
      setGonderiliyor(false);
    }
  }

  if (yukleniyor) return <p className="subtitle">Yükleniyor…</p>;

  return (
    <>
      {!configured && (
        <div className="card" style={{ borderColor: "var(--danger, #b00)" }}>
          <strong>NETGSM bilgileri tanımlı değil.</strong>
          <p style={{ margin: "6px 0 0", color: "var(--muted)", fontSize: 14 }}>
            Vercel → Settings → Environment Variables içine{" "}
            <code>NETGSM_USERCODE</code>, <code>NETGSM_PASSWORD</code> ve{" "}
            <code>NETGSM_HEADER</code> girip yeniden dağıtım alın. O zamana kadar
            gönderim yapılamaz.
          </p>
        </div>
      )}

      {/* ---------------- Alıcılar ---------------- */}
      <div className="card">
        <h2 style={{ marginTop: 0 }}>Alıcılar</h2>

        <input
          value={ara}
          onChange={(e) => setAra(e.target.value)}
          placeholder="Müşteri ara — isim, firma, şehir veya numara…"
          aria-label="Müşteri ara"
        />

        <div
          style={{
            maxHeight: 260,
            overflowY: "auto",
            marginTop: 12,
            border: "1px solid var(--border, #333)",
            borderRadius: 8,
          }}
        >
          {listelenen.map((c) => (
            <label
              key={c.id}
              style={{
                display: "flex",
                gap: 10,
                alignItems: "center",
                padding: "8px 12px",
                cursor: "pointer",
              }}
            >
              <input
                type="checkbox"
                checked={secili.has(c.id)}
                onChange={() => toggle(c.id)}
                style={{ width: "auto", margin: 0 }}
              />
              <span style={{ flex: 1 }}>{title(c)}</span>
              <span style={{ color: "var(--muted)", fontSize: 13 }}>
                {c.phone}
              </span>
            </label>
          ))}
          {!listelenen.length && (
            <p style={{ padding: 12, margin: 0, color: "var(--muted)" }}>
              {telefonlu.length
                ? "Aramaya uyan müşteri yok."
                : "Müşteri defterinde telefon numarası kayıtlı kimse yok."}
            </p>
          )}
        </div>

        <p style={{ margin: "10px 0 4px", fontSize: 14 }}>
          Listede olmayan numaralar (virgül veya satır ile ayırın):
        </p>
        <textarea
          rows={2}
          value={elle}
          onChange={(e) => setElle(e.target.value)}
          placeholder="0532 123 45 67, 05551234567"
        />
      </div>

      {/* ---------------- Mesaj ---------------- */}
      <div className="card">
        <h2 style={{ marginTop: 0 }}>Mesaj</h2>

        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
          {SABLONLAR.map((s) => (
            <button
              key={s.ad}
              className="btn small secondary"
              onClick={() => setMesaj(s.metin)}
            >
              {s.ad}
            </button>
          ))}
          {sayim.encoding === "TR" && (
            <button
              className="btn small secondary"
              onClick={() => setMesaj((m) => stripTurkish(m))}
              title="Türkçe harfleri kaldırarak SMS başına 70 yerine 160 karakter hakkı kazanırsınız"
            >
              ⚡ Türkçe karakterleri kaldır
            </button>
          )}
        </div>

        <textarea
          rows={4}
          value={mesaj}
          onChange={(e) => setMesaj(e.target.value)}
          placeholder="Mesajınızı yazın…"
        />

        <div style={{ marginTop: 12 }}>
          <label style={{ fontSize: 14, display: "block", marginBottom: 4 }}>
            Mesaj türü
          </label>
          <select
            value={iysfilter}
            onChange={(e) => setIysfilter(e.target.value as "0" | "11" | "12")}
            style={{ width: "auto" }}
          >
            <option value="0">Bilgilendirme — kargo, sipariş, hatırlatma</option>
            <option value="11">Ticari / kampanya — alıcı bireysel</option>
            <option value="12">Ticari / kampanya — alıcı tacir (firma)</option>
          </select>
          <p style={{ margin: "6px 0 0", color: "var(--muted)", fontSize: 13 }}>
            {iysfilter === "0"
              ? "Mevcut alışveriş ilişkisine dair mesaj — İYS onayı aranmaz."
              : "Ticari ileti — İYS'de onayı olmayan numaralara gönderilmez. Kampanya mesajını bilgilendirme olarak göndermek mevzuata aykırıdır."}
          </p>
        </div>

        <p style={{ margin: "8px 0 0", color: "var(--muted)", fontSize: 13 }}>
          {sayim.chars} karakter · <strong>{sayim.segments}</strong> SMS ·{" "}
          {tekilSayi} alıcı ={" "}
          <strong style={{ color: "var(--brand-light)" }}>{kredi} kredi</strong>
          {sayim.encoding === "TR" && (
            <>
              {" "}
              — Türkçe karakter kullanıldığı için SMS başına {sayim.limit} karakter
              (aksi halde 160 olurdu).
            </>
          )}
        </p>

        {gecersiz.length > 0 && (
          <p style={{ margin: "8px 0 0", color: "#e88", fontSize: 13 }}>
            Geçersiz numara atlanacak: {gecersiz.join(", ")}
          </p>
        )}

        {sonuc && (
          <p
            style={{
              margin: "10px 0 0",
              fontSize: 14,
              color: sonuc.ok ? "var(--brand-light)" : "#e88",
            }}
          >
            {sonuc.text}
          </p>
        )}

        <div style={{ marginTop: 14 }}>
          <button
            className="btn"
            onClick={gonder}
            disabled={!tekilSayi || !mesaj.trim() || gonderiliyor || !configured}
          >
            {gonderiliyor ? "Gönderiliyor…" : `Gönder (${kredi} kredi)`}
          </button>
        </div>
      </div>

      {/* ---------------- Geçmiş ---------------- */}
      <div className="card">
        <h2 style={{ marginTop: 0 }}>Gönderim Geçmişi</h2>
        {records.length ? (
          <div style={{ overflowX: "auto" }}>
            <table>
              <thead>
                <tr>
                  <th>Tarih</th>
                  <th>Gönderen</th>
                  <th>Mesaj</th>
                  <th>Alıcı</th>
                  <th>Kredi</th>
                  <th>Durum</th>
                </tr>
              </thead>
              <tbody>
                {records.map((r) => (
                  <tr key={r.id}>
                    <td style={{ whiteSpace: "nowrap" }}>
                      {new Date(r.createdAt).toLocaleString("tr-TR", {
                        timeZone: "Europe/Istanbul",
                        dateStyle: "short",
                        timeStyle: "short",
                      })}
                    </td>
                    <td>{r.sender}</td>
                    <td style={{ maxWidth: 320 }}>{r.message}</td>
                    <td>{r.recipients.length}</td>
                    <td>{r.credits}</td>
                    <td>
                      {r.ok ? (
                        <span style={{ color: "var(--brand-light)" }}>✓</span>
                      ) : (
                        <span style={{ color: "#e88" }} title={r.error}>
                          ✗ {r.error?.slice(0, 40)}
                        </span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <p style={{ color: "var(--muted)", margin: 0 }}>
            Henüz SMS gönderilmemiş.
          </p>
        )}
      </div>
    </>
  );
}
