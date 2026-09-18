"use client";

// Sahiplere özel: patrona WhatsApp fiş bildirimi kurulum durumu + test gönderimi.

import { useEffect, useState } from "react";
import Icon from "@/components/shell/Icon";

interface Durum { api: boolean; alicilar: string[]; sablon: string | null; dil: string; musteriSablon: string | null; webhook: boolean; imza: boolean; }
interface Sonuc { ok: boolean; gonderilen: string[]; hatalar: string[]; notlar?: string[]; yontem: string; sablon?: string; }

export default function BildirimAyarlari() {
  const [durum, setDurum] = useState<Durum | null>(null);
  const [hata, setHata] = useState("");
  const [gonderiyor, setGonderiyor] = useState(false);
  const [sonuc, setSonuc] = useState<Sonuc | null>(null);
  // Patron denemesi: numara yazılırsa oraya, boşsa PATRON_WHATSAPP'a gider
  const [patronTel, setPatronTel] = useState("");
  // Müşteri fişi denemesi — bir numara girilir, müşteri şablonuyla örnek fiş gider
  const [musteriTel, setMusteriTel] = useState("");
  const [musteriGonderiyor, setMusteriGonderiyor] = useState(false);
  const [musteriSonuc, setMusteriSonuc] = useState<Sonuc | null>(null);

  async function musteriTest() {
    if (!musteriTel.trim()) return;
    setMusteriGonderiyor(true);
    setMusteriSonuc(null);
    try {
      const r = await fetch("/api/whatsapp/test", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ musteriTelefon: musteriTel.trim() }),
      });
      const d = await r.json();
      if (d.sonuc) setMusteriSonuc(d.sonuc);
      else setHata(d.error || "Test gönderilemedi.");
    } catch {
      setHata("Sunucuya ulaşılamadı.");
    } finally {
      setMusteriGonderiyor(false);
    }
  }

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
      const r = await fetch("/api/whatsapp/test", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(patronTel.trim() ? { patronTelefon: patronTel.trim() } : {}),
      });
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
          <input
            style={{ maxWidth: 220 }}
            placeholder="Deneme: 05xx… ya da 05xx…:Ad Bey"
            value={patronTel}
            onChange={(e) => setPatronTel(e.target.value)}
            inputMode="tel"
            title="Boş bırakılırsa PATRON_WHATSAPP numaralarına gider"
          />
          <button type="button" className="btn" onClick={test} disabled={!durum?.api || gonderiyor || (!hazir && !patronTel.trim())}>
            <Icon name="arrow-up-right" size={16} /> {gonderiyor ? "Gönderiliyor…" : "Test fişi gönder"}
          </button>
        </div>
        {hata && <div className="notice err">{hata}</div>}
        {!durum && !hata && <div className="skeleton" style={{ height: 120 }} />}
        {durum && (
          <ul style={{ listStyle: "none" }}>
            <Satir ok={durum.api} baslik="WhatsApp Cloud API" aciklama={durum.api ? "WHATSAPP_TOKEN ve WHATSAPP_PHONE_ID tanımlı." : "Vercel ortam değişkenlerinde WHATSAPP_TOKEN ve WHATSAPP_PHONE_ID eksik."} />
            <Satir ok={durum.alicilar.length > 0} baslik="Alıcı numaralar (PATRON_WHATSAPP)" aciklama={durum.alicilar.length ? durum.alicilar.join(", ") : "Örn. PATRON_WHATSAPP=05325099442:Özgür Bey,05336610287:Gültekin Bey — numara ve hitap, virgülle."} />
            <Satir ok={!!durum.sablon} baslik="Onaylı şablon (WHATSAPP_TEMPLATE_SIPARIS)" aciklama={durum.sablon ? `${durum.sablon} · dil: ${durum.dil} · birden çok ad varsa sırayla denenir` : "Şablon tanımlı değil: mesaj yalnızca alıcı son 24 saatte işletmeye yazdıysa gider. Kalıcı çözüm için Meta'da belge başlıklı şablonu onaylatıp adını girin."} />
          </ul>
        )}
        {sonuc && (
          <div className={`notice ${sonuc.ok ? "ok" : "err"}`} style={{ marginTop: 14 }}>
            <strong>{sonuc.ok ? "Test fişi gönderildi" : "Gönderilemedi"}</strong>
            {" · "}yöntem: {sonuc.yontem === "sablon" ? `onaylı şablon${sonuc.sablon ? ` (${sonuc.sablon})` : ""}` : sonuc.yontem === "serbest" ? "serbest belge mesajı (24 saat penceresi)" : "—"}
            {sonuc.gonderilen.length > 0 && <div>Gidenler: {sonuc.gonderilen.join(", ")}</div>}
            {(sonuc.notlar || []).map((n, i) => <div key={`n${i}`} style={{ marginTop: 4 }}>Not: {n}</div>)}
            {sonuc.hatalar.map((h, i) => <div key={i} style={{ marginTop: 4 }}>{h}</div>)}
          </div>
        )}
      </div>

      <div className="card">
        <div className="card-head">
          <span className="card-head-icon"><Icon name="users" size={18} /></span>
          <div>
            <h2>Müşteriye WhatsApp Fiş Bildirimi</h2>
            <span className="card-head-sub">Sipariş formunda &quot;müşteriye bildir&quot; işaretliyse: önce WhatsApp ile fiş, WhatsApp&apos;ı yoksa SMS</span>
          </div>
        </div>
        {durum && (
          <ul style={{ listStyle: "none" }}>
            <Satir ok={!!durum.musteriSablon} baslik="Müşteri şablonu (WHATSAPP_TEMPLATE_MUSTERI)" aciklama={durum.musteriSablon ? `${durum.musteriSablon} · dil: ${durum.dil}` : "Tanımlı değil: müşteriye yalnızca SMS gider. Meta'da 2 değişkenli, belge başlıklı şablonu (örn. musteri_siparis_fisi) onaylatıp adını girin."} />
            <Satir ok={durum.webhook} baslik="Teslim webhook'u (WHATSAPP_VERIFY_TOKEN)" aciklama={durum.webhook ? "Tanımlı. Meta uygulamasında Callback URL: https://olgasiparis.com/api/whatsapp/webhook · alan: messages" : "Tanımlı değil: numarada WhatsApp yoksa SMS'e düşülemez. Rastgele bir doğrulama metni belirleyip Vercel'e ve Meta'daki webhook ayarına aynı değeri girin."} />
            <Satir ok={durum.imza} baslik="Webhook imzası (WHATSAPP_APP_SECRET)" aciklama={durum.imza ? "Tanımlı: gelen istekler Meta imzasıyla doğrulanıyor." : "İsteğe bağlı. Meta uygulaması → App settings → Basic → App secret değerini girerseniz webhook istekleri imzayla doğrulanır."} />
          </ul>
        )}
        <div className="row" style={{ marginTop: 12, gap: 8 }}>
          <input
            style={{ maxWidth: 220 }}
            placeholder="Deneme numarası: 05xx xxx xx xx"
            value={musteriTel}
            onChange={(e) => setMusteriTel(e.target.value)}
            inputMode="tel"
          />
          <button type="button" className="btn secondary" onClick={musteriTest} disabled={!durum?.musteriSablon || musteriGonderiyor || !musteriTel.trim()}>
            <Icon name="arrow-up-right" size={16} /> {musteriGonderiyor ? "Gönderiliyor…" : "Müşteri fişi dene"}
          </button>
        </div>
        {musteriSonuc && (
          <div className={`notice ${musteriSonuc.ok ? "ok" : "err"}`} style={{ marginTop: 14 }}>
            <strong>{musteriSonuc.ok ? "Örnek müşteri fişi gönderildi" : "Gönderilemedi"}</strong>
            {" · "}yöntem: {musteriSonuc.yontem === "sablon" ? `onaylı şablon${musteriSonuc.sablon ? ` (${musteriSonuc.sablon})` : ""}` : musteriSonuc.yontem === "serbest" ? "serbest belge mesajı (24 saat penceresi)" : "—"}
            {musteriSonuc.gonderilen.length > 0 && <div>Gidenler: {musteriSonuc.gonderilen.join(", ")}</div>}
            {(musteriSonuc.notlar || []).map((n, i) => <div key={`n${i}`} style={{ marginTop: 4 }}>Not: {n}</div>)}
            {musteriSonuc.hatalar.map((h, i) => <div key={i} style={{ marginTop: 4 }}>{h}</div>)}
          </div>
        )}
      </div>

      <div className="card">
        <div className="card-head">
          <span className="card-head-icon"><Icon name="file-text" size={18} /></span>
          <div>
            <h2>Meta'da tanımlı şablonlar</h2>
            <span className="card-head-sub">WhatsApp Manager → Mesaj şablonları · ikisi de Utility (Yardımcı), Türkçe, başlık: Belge</span>
          </div>
        </div>
        <div className="grid cols-2">
          <div>
            <label>Patrona fiş · <code>siparis_fisi_v4</code> · 5 değişken</label>
            <pre style={{ whiteSpace: "pre-wrap", background: "var(--surface-2)", padding: 12, borderRadius: 10, fontSize: 13 }}>{`Sayın {{1}}, sipariş sisteminden yeni bir sipariş fişi geldi.

Sipariş: {{2}}
Müşteri: {{3}}
Toplam tutar: ₺{{4}}
Siparişi alan: {{5}}

Fişin tamamı ekteki PDF dosyasındadır. İyi çalışmalar.`}</pre>
            <p className="muted" style={{ fontSize: 12.5, marginTop: 8 }}>Alt bilgi: Olga Çerçeve sipariş sistemi · {"{{1}}"} hitap PATRON_WHATSAPP&apos;taki addır (Özgür Bey, Gültekin Bey) · Örnekler: Özgür Bey · Toptan OLG-2026-275 · Ayşe Özyürek · 4.267,08 · Alaattin Yıldız</p>
          </div>
          <div>
            <label>Müşteriye fiş · <code>musteri_siparis_fisi_v2</code> · 2 değişken</label>
            <pre style={{ whiteSpace: "pre-wrap", background: "var(--surface-2)", padding: 12, borderRadius: 10, fontSize: 13 }}>{`Sayın {{1}}, {{2}} numaralı siparişiniz alınmıştır. Siparişinizin ayrıntılarını ekteki PDF dosyasında görebilirsiniz. Bizi tercih ettiğiniz için teşekkür ederiz.

www.olgacerceve.com`}</pre>
            <p className="muted" style={{ fontSize: 12.5, marginTop: 8 }}>Alt bilgi: Olga Çerçeve · Örnekler: Ayşe Özyürek · OLG-2026-275</p>
          </div>
        </div>
        <div className="notice info" style={{ marginTop: 12 }}>
          Şablon adları Vercel&apos;de <code>WHATSAPP_TEMPLATE_SIPARIS</code> ve <code>WHATSAPP_TEMPLATE_MUSTERI</code> değişkenlerinde; değiştirince yeniden dağıtın ve bu sayfadan deneyin.
        </div>
      </div>
    </div>
  );
}
