"use client";

// SMS gönderim ekranı — müşteri defterinden çoklu alıcı seçimi, hazır şablonlar,
// canlı kredi sayacı ve gönderim geçmişi.

import { useEffect, useMemo, useState } from "react";
import Icon from "@/components/shell/Icon";
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

// Alıcı satırı: global <label> stili (küçük, büyük harf, kalın) burada
// istenmez — satır metni normal gövde yazısı olarak kalır.
const aliciSatir: React.CSSProperties = {
  display: "flex",
  gap: 10,
  alignItems: "center",
  padding: "9px 12px",
  cursor: "pointer",
  margin: 0,
  fontSize: 14,
  fontWeight: 500,
  letterSpacing: 0,
  textTransform: "none",
  color: "var(--text)",
  borderBottom: "1px solid var(--hairline)",
};

export default function SmsPanel() {
  const [customers, setCustomers] = useState<Customer[]>([]);
  const [records, setRecords] = useState<SmsRecord[]>([]);
  const [toplam, setToplam] = useState<number | null>(null); // depodaki toplam gönderim
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
        if (typeof s?.total === "number") setToplam(s.total);
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
          .then((s) => {
            if (s?.records) setRecords(s.records);
            if (typeof s?.total === "number") setToplam(s.total);
          })
          .catch(() => {});
      } else {
        // Ham yanıtı da gösteriyoruz: kod çevirileri kesin değil, gerçek
        // sebebi NETGSM'in kendi yanıtı söylüyor.
        const ham = d.raw ? `  ·  NETGSM ham yanıt: ${d.raw}` : "";
        setSonuc({
          ok: false,
          text: `${d.error || "Gönderilemedi."}${ham}`,
        });
      }
    } catch {
      setSonuc({ ok: false, text: "Sunucuya ulaşılamadı." });
    } finally {
      setGonderiliyor(false);
    }
  }

  if (yukleniyor) return <div className="empty">Yükleniyor…</div>;

  return (
    <>
      {!configured && (
        <div className="notice err" style={{ margin: "0 0 16px", display: "flex", gap: 12, alignItems: "flex-start" }}>
          <span style={{ color: "var(--error)", flexShrink: 0, display: "inline-flex", marginTop: 2 }}>
            <Icon name="alert" size={18} />
          </span>
          <div>
            <strong>NETGSM bilgileri tanımlı değil.</strong>
            <p className="text-2" style={{ margin: "6px 0 0", fontSize: 14 }}>
              Vercel → Settings → Environment Variables içine{" "}
              <code>NETGSM_USERCODE</code>, <code>NETGSM_PASSWORD</code> ve{" "}
              <code>NETGSM_HEADER</code> girip yeniden dağıtım alın. O zamana kadar
              gönderim yapılamaz.
            </p>
          </div>
        </div>
      )}

      {/* ---------------- Alıcılar ---------------- */}
      <div className="card no-print">
        <div className="card-head">
          <span className="card-head-icon"><Icon name="users" size={16} /></span>
          <div>
            <h2>Alıcılar</h2>
            <span className="card-head-sub">
              {secili.size ? `${secili.size} seçili · ` : ""}{telefonlu.length} numaralı müşteri
            </span>
          </div>
        </div>

        <input
          value={ara}
          onChange={(e) => setAra(e.target.value)}
          placeholder="Ara: isim / firma / şehir / numara"
          aria-label="Müşteri ara"
        />

        <div
          style={{
            maxHeight: 260,
            overflowY: "auto",
            marginTop: 12,
            border: "1px solid var(--border)",
            borderRadius: "var(--radius-sm)",
            background: "var(--surface-2)",
          }}
        >
          {listelenen.map((c, i) => (
            <label
              key={c.id}
              style={i === listelenen.length - 1 ? { ...aliciSatir, borderBottom: "none" } : aliciSatir}
            >
              <input
                type="checkbox"
                checked={secili.has(c.id)}
                onChange={() => toggle(c.id)}
                style={{ width: "auto", margin: 0, flexShrink: 0 }}
              />
              <span style={{ flex: 1, minWidth: 0, overflowWrap: "anywhere" }}>{title(c)}</span>
              <span className="muted num" style={{ fontSize: 13, whiteSpace: "nowrap" }}>
                {c.phone}
              </span>
            </label>
          ))}
          {!listelenen.length && (
            <div className="empty" style={{ padding: "18px 12px" }}>
              {telefonlu.length
                ? "Aramaya uyan müşteri yok."
                : "Müşteri defterinde telefon numarası kayıtlı kimse yok."}
            </div>
          )}
        </div>

        <label style={{ marginTop: 12 }}>
          Listede olmayan numaralar (virgül veya satır ile ayırın)
        </label>
        <textarea
          rows={2}
          value={elle}
          onChange={(e) => setElle(e.target.value)}
          placeholder="0532 123 45 67, 05551234567"
        />
      </div>

      {/* ---------------- Mesaj ---------------- */}
      <div className="card no-print">
        <div className="card-head">
          <span className="card-head-icon"><Icon name="message" size={16} /></span>
          <div>
            <h2>Mesaj</h2>
          </div>
        </div>

        <div className="row" style={{ gap: 8, marginBottom: 10 }}>
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
              <Icon name="zap" size={14} /> Türkçe karakterleri kaldır
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
          <label>Mesaj türü</label>
          {/* Global select stili width:100% + 34px sağ boşluk (ok işareti) verir;
              width:auto bu boşluğu içsel genişliğe katmadığı için son harf okla
              çakışıyordu. Genişlik 440'ta sınırlanır, ≤680'de kart genişliğine yayılır. */}
          <select
            value={iysfilter}
            onChange={(e) => setIysfilter(e.target.value as "0" | "11" | "12")}
            style={{ maxWidth: 440, textOverflow: "ellipsis" }}
          >
            <option value="0">Bilgilendirme — kargo, sipariş, hatırlatma</option>
            <option value="11">Ticari / kampanya — alıcı bireysel</option>
            <option value="12">Ticari / kampanya — alıcı tacir (firma)</option>
          </select>
          <p className="muted" style={{ margin: "6px 0 0", fontSize: 13 }}>
            {iysfilter === "0"
              ? "Mevcut alışveriş ilişkisine dair mesaj — İYS onayı aranmaz."
              : "Ticari ileti — İYS'de onayı olmayan numaralara gönderilmez. Kampanya mesajını bilgilendirme olarak göndermek mevzuata aykırıdır."}
          </p>
        </div>

        <p className="muted" style={{ margin: "8px 0 0", fontSize: 13 }}>
          {sayim.chars} karakter · <strong>{sayim.segments}</strong> SMS ·{" "}
          {tekilSayi} alıcı ={" "}
          <strong style={{ color: "var(--brand)" }}>{kredi} kredi</strong>
          {sayim.encoding === "TR" && (
            <>
              {" "}
              — Türkçe karakter kullanıldığı için SMS başına {sayim.limit} karakter
              (aksi halde 160 olurdu).
            </>
          )}
        </p>

        {gecersiz.length > 0 && (
          <p style={{ margin: "8px 0 0", color: "var(--error)", fontSize: 13 }}>
            Geçersiz numara atlanacak: {gecersiz.join(", ")}
          </p>
        )}

        {sonuc && (
          <div className={`notice ${sonuc.ok ? "ok" : "err"}`} style={{ margin: "10px 0 0" }}>
            {sonuc.text}
          </div>
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
        <div className="card-head">
          <span className="card-head-icon"><Icon name="clock" size={16} /></span>
          <div>
            <h2>Gönderim Geçmişi</h2>
            <span className="card-head-sub">
              {toplam !== null && toplam > records.length
                ? `Son ${records.length} gönderim · toplam ${toplam.toLocaleString("tr-TR")}`
                : `${records.length} gönderim`}
            </span>
          </div>
        </div>
        {records.length ? (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Tarih</th>
                  <th>Gönderen</th>
                  <th>Mesaj</th>
                  <th className="num">Alıcı</th>
                  <th className="num">Kredi</th>
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
                    <td style={{ whiteSpace: "nowrap" }}>{r.sender}</td>
                    <td style={{ minWidth: 200, maxWidth: 320 }}>{r.message}</td>
                    <td className="num">{r.recipients.length}</td>
                    <td className="num">{r.credits}</td>
                    <td>
                      {r.ok ? (
                        <span className="badge ok">✓</span>
                      ) : (
                        <span className="badge err" title={r.error} style={{ whiteSpace: "normal" }}>
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
          <div className="empty" style={{ padding: "18px 12px" }}>
            Henüz SMS gönderilmemiş.
          </div>
        )}
      </div>
    </>
  );
}
