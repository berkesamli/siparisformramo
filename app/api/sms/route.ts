import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import { sendSms, smsConfigured, type IysFilter } from "@/lib/sms";
import { smsSegments } from "@/lib/sms-format";
import {
  listSmsRecords,
  saveSmsRecord,
  newSmsId,
  istanbulDateKey,
  type SmsRecord,
} from "@/lib/sms-log";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

// Gönderim geçmişi — yalnızca çalışanlar.
export async function GET() {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }
  const records = await listSmsRecords();
  return NextResponse.json({ ok: true, configured: smsConfigured(), records });
}

// Tek seferde gönderilebilecek alıcı sayısı — yanlışlıkla yüzlerce kişiye
// gönderim yapılmasını ve tek istekte kredi tüketilmesini sınırlar.
const MAX_ALICI = 200;

export async function POST(req: Request) {
  const user = await getSessionUser();
  if (!user || user.role !== "staff") {
    return NextResponse.json({ ok: false, error: "Yetkisiz." }, { status: 401 });
  }

  const body = (await req.json().catch(() => null)) as {
    numbers?: unknown;
    message?: unknown;
    iysfilter?: unknown;
  } | null;

  const numbers = Array.isArray(body?.numbers)
    ? body!.numbers.map((n) => String(n)).filter(Boolean)
    : [];
  const message = String(body?.message || "").trim();

  // Tanınmayan bir değer gelirse bilgilendirmeye düşüyoruz: "0" İYS kontrolü
  // istemez, yani en dar kapsam. Ticari gönderim ancak açıkça seçilirse yapılır.
  const gelenFiltre = String(body?.iysfilter || "0");
  const iysfilter: IysFilter =
    gelenFiltre === "11" || gelenFiltre === "12" ? gelenFiltre : "0";

  if (!numbers.length) {
    return NextResponse.json(
      { ok: false, error: "En az bir alıcı seçin." },
      { status: 400 }
    );
  }
  if (!message) {
    return NextResponse.json(
      { ok: false, error: "Mesaj boş olamaz." },
      { status: 400 }
    );
  }
  if (numbers.length > MAX_ALICI) {
    return NextResponse.json(
      {
        ok: false,
        error: `Tek seferde en fazla ${MAX_ALICI} alıcıya gönderilebilir. Listeyi bölün.`,
      },
      { status: 400 }
    );
  }

  const result = await sendSms(numbers, message, iysfilter);
  const { segments } = smsSegments(message);

  // Başarılı da olsa başarısız da olsa kaydı tutuyoruz — hata ayıklarken ve
  // kredi tüketimini incelerken bu geçmiş gerekiyor.
  const now = new Date();
  const rec: SmsRecord = {
    id: newSmsId(now),
    dateKey: istanbulDateKey(now),
    createdAt: now.toISOString(),
    sender: user.name || user.username,
    message,
    recipients: result.sent,
    segments,
    credits: segments * result.sent.length,
    ok: result.ok,
    jobId: result.jobId,
    error: result.error,
    iysfilter,
  };
  await saveSmsRecord(rec).catch(() => {
    /* kayıt tutulamazsa gönderimi başarısız sayma */
  });

  if (!result.ok) {
    return NextResponse.json(
      {
        ok: false,
        error: result.error,
        code: result.code,
        invalid: result.invalid,
      },
      { status: 502 }
    );
  }

  return NextResponse.json({
    ok: true,
    jobId: result.jobId,
    sent: result.sent.length,
    credits: rec.credits,
    invalid: result.invalid,
  });
}
