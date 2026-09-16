"use client";

// Sahiplere özel: patrona WhatsApp fiş bildirimi kurulum durumu + test gönderimi.

import { useEffect, useState } from "react";
import Icon from "@/components/shell/Icon";

interface Durum { api: boolean; alicilar: string[]; sablon: string | null; dil: string; }
interface Sonuc { ok: boolean; gonderilen: string[]; hatalar: string[]; yontem: string; }

export default function BildirimAyarlari() {
  const [durum, setDurum] = useState<Durum | null>(null);
  const [hata, setHata] = useState("");
  const [gonderiyor, setGonderiyor] = useState(false);
  const [sonuc, setSonuc] = useState<Sonuc | null>(null);

  useEffect(() => {
    fetch("/api/whatsapp/test")
      .then((r) => r.json())
      .then((d) => (d.ok ? setDurum(d) : setHata(d.error || "Durum alınamadı.")))
      .catch(() => setHata("Sunucuya ulaşılamadı."));
  }, []);

  async function test() {
    setGonderiyor(true);
    setSonuc(null);
    try {
      const r = await fetch("/api/whatsapp/test", { method: "POST" });
      const d = await r.json();
      if (d.sonuc) setSonuc(d.sonuc);
      else setHata(d.error || "Test gönderilemedi.");
    } catch {
      setHata("Sunucuya ulaşılamadı.");
    } finally {
      setGonderiyor(false);
    }
  }

  const hazir = !!durum?.api && (durum?.alicilar.length || 0) > 0;
  const Satir = ({ ok, baslik, aciklama }: { ok: boolean; baslik: string; aciklama: string }) => (
    <li className="row" style={{ padding: "10px 0", borderBottom: "1px solid var(--border)", alignItems: "flex-start", flexWrap: "nowrap" }}>
      <span className={`badge ${ok ? "ok" : "warn"}`} style={{ marginTop: 2 }}>
        <Icon name={ok ? "check-circle" : "alert"} size={13} /> {ok ? "Hazır" : "Eksik"}
      </span>
      <span style={{ minWidth: 0 }}>
        <strong style={{ display: "block", fontSize: 14 }}>{baslik}</strong>
        <span className="muted" style={{ fontSize: 12.5 }}>{aciklama}</span>
      </span>
    </li>
  );

  return (
    <div className="dash">
      <div className="card">
        <div className="card-head">
          <span className="card-head-icon"><Icon name="message" size={18} /></span>
          <div>
            <h2>Patrona WhatsApp Fiş Bildirimi</h2>
            <span className="card-head-sub">Her sipariş kaydedildiğinde fiş PDF'i dosya olarak gider</span>
          </div>
          <span className="spacer" />
          <button type="button" className="btn" onClick={test} disabled={!hazir || gonderiyor}>
            <Icon name="arrow-up-right" size={16} /> {gonderiyor ? "Gönderiliyor…" : "Test fişi gönder"}
          </button>
        </div>
        {hata && <div className="notice err">{hata}</div>}
        {!durum && !hata && <div className="skeleton" style={{ height: 120 }} />}
        {durum && (
          <ul style={{ listStyle: "none" }}>
            <Satir ok={durum.api} baslik="WhatsApp Cloud API" aciklama={durum.api ? "WHATSAPP_TOKEN ve WHATSAPP_PHONE_ID tanımlı." : "Vercel ortam değişkenlerinde WHATSAPP_TOKEN ve WHATSAPP_PHONE_ID eksik."} />
            <Satir ok={durum.alicilar.length > 0} baslik="Alıcı numaralar (PATRON_WHATSAPP)" aciklama={durum.alicilar.length ? durum.alicilar.join(", ") : "Örn. PATRON_WHATSAPP=05325099442 — birden çok numara virgülle."} />
            <Satir ok={!!durum.sablon} baslik="Onaylı şablon (WHATSAPP_TEMPLATE_SIPARIS)" aciklama={durum.sablon ? `${durum.sablon} · dil: ${durum.dil}` : "Şablon tanımlı değil: mesaj yalnızca alıcı son 24 saatte işletmeye yazdıysa gider. Kalıcı çözüm için Meta'da belge başlıklı şablonu onaylatıp adını girin."} />
          </ul>
        )}
        {sonuc && (
          <div className={`notice ${sonuc.ok ? "ok" : "err"}`} style={{ marginTop: 14 }}>
            <strong>{sonuc.ok ? "Test fişi gönderildi" : "Gönderilemedi"}</strong>
            {" · "}yöntem: {sonuc.yontem === "sablon" ? "onaylı şablon" : sonuc.yontem === "serbest" ? "serbest belge mesajı (24 saat penceresi)" : "—"}
            {sonuc.gonderilen.length > 0 && <div>Gidenler: {sonuc.gonderilen.join(", ")}</div>}
            {sonuc.hatalar.map((h, i) => <div key={i} style={{ marginTop: 4 }}>{h}</div>)}
          </div>
        )}
      </div>

      <div className="card">
        <div className="card-head">
          <span className="card-head-icon"><Icon name="file-text" size={18} /></span>
          <div>
            <h2>Meta'da tanımlanacak şablon</h2>
            <span className="card-head-sub">WhatsApp Manager → Message templates → Create template</span>
          </div>
        </div>
        <div className="grid cols-2">
          <div>
            <label>Ad / Kategori / Dil</label>
            <p><code>siparis_fisi</code> · Utility (Yardımcı) · Türkçe (tr)</p>
            <label style={{ marginTop: 12 }}>Başlık (Header)</label>
            <p>Tür: <strong>Document</strong> — örnek olarak herhangi bir PDF yükleyin.</p>
            <label style={{ marginTop: 12 }}>Alt bilgi (Footer, isteğe bağlı)</label>
            <p>Olga Çerçeve sipariş sistemi</p>
          </div>
          <div>
            <label>Gövde (Body)</label>
            <pre style={{ whiteSpace: "pre-wrap", background: "var(--surface-2)", padding: 12, borderRadius: 10, fontSize: 13.5 }}>{`Yeni sipariş: {{1}}
Müşteri: {{2}}
Tutar: ₺{{3}}
Alan: {{4}}`}</pre>
            <p className="muted" style={{ fontSize: 12.5, marginTop: 8 }}>
              Örnek değerler (onay için istenir): 1 → Toptan OLG-2026-275 · 2 → Ayşe Özyürek · 3 → 4.267,08 · 4 → Alaattin Yıldız
            </p>
          </div>
        </div>
        <div className="notice info" style={{ marginTop: 12 }}>
          Şablon onaylanınca Vercel'de <code>WHATSAPP_TEMPLATE_SIPARIS=siparis_fisi</code> ve <code>PATRON_WHATSAPP=05325099442</code> tanımlayıp yeniden dağıtın; sonra bu sayfadan test fişi gönderin.
        </div>
      </div>
    </div>
  );
}
