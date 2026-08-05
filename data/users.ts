// Kullanıcılar — çalışanlar (staff) sipariş paneline, müşteriler (customer) portala girer.
// Üretimde USERS_JSON ortam değişkeni ile bu liste tamamen değiştirilebilir:
// USERS_JSON='[{"username":"ali","password":"...","name":"Ali","role":"staff"}]'

export type Role = "staff" | "customer";

export interface User {
  username: string;
  password: string;
  name: string;
  role: Role;
  // Çalışanın varsayılan şubesi (opsiyonel) — sipariş/tahsilat formlarında
  // öneri olarak kullanılır; personel iki şubede de çalışabildiği için
  // kayıt anında her zaman değiştirilebilir.
  branch?: "ankara" | "istanbul";
}

// Gerçek kullanıcılar USERS_JSON ortam değişkeninde tanımlıdır (Vercel →
// Settings → Environment Variables). Aşağıdaki liste yalnızca USERS_JSON hiç
// tanımlanmadığında (ör. ilk kurulum/yerel geliştirme) devreye girer.
const DEFAULT_USERS: User[] = [
  { username: "demo", password: "kurulum-tamamlanmadi", name: "Demo", role: "customer" },
];

export function getUsers(): User[] {
  const raw = process.env.USERS_JSON;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed as User[];
    } catch {
      // hatalı JSON — varsayılan listeye düş
    }
  }
  return DEFAULT_USERS;
}

// Kullanıcı adı eşleştirmesi büyük/küçük harf ve Türkçe karakter duyarsızdır:
// "Özgür", "özgür", "ozgur", "OZGUR" aynı kullanıcıya gider. (JS'in toLowerCase
// fonksiyonu Türkçe İ/ı harflerinde beklenmedik sonuç verdiği için harfler önce
// ASCII karşılığına çevrilir.)
const TR_TO_ASCII: Record<string, string> = {
  ç: "c", Ç: "c", ğ: "g", Ğ: "g", ı: "i", İ: "i",
  ö: "o", Ö: "o", ş: "s", Ş: "s", ü: "u", Ü: "u",
};

function normalizeUsername(s: string): string {
  return String(s || "")
    .replace(/[çÇğĞıİöÖşŞüÜ]/g, (c) => TR_TO_ASCII[c])
    .toLowerCase()
    .replace(/\s+/g, "");
}

export function findUser(username: string, password: string): User | undefined {
  const q = normalizeUsername(username);
  return getUsers().find(
    (u) => normalizeUsername(u.username) === q && u.password === password
  );
}

// Raporlar (ciro, tahsilat, kâr kırılımları) yalnızca firma sahiplerine açıktır.
// Liste OWNER_USERNAMES ortam değişkeniyle değiştirilebilir:
// OWNER_USERNAMES="berke,özgür,eren"
const DEFAULT_OWNERS = ["berke", "özgür"];

export function ownerUsernames(): string[] {
  const raw = process.env.OWNER_USERNAMES;
  const list = raw
    ? raw.split(",").map((s) => s.trim()).filter(Boolean)
    : DEFAULT_OWNERS;
  return list.map(normalizeUsername);
}

export function isOwner(username: string | undefined | null): boolean {
  if (!username) return false;
  return ownerUsernames().includes(normalizeUsername(username));
}

// Finans ekranları (kasa, giderler, çek/senet, tahsilat listeleri) sahiplere ve
// FINANCE_USERNAMES ile eklenen kişilere açıktır. Sahipler otomatik dahildir:
// FINANCE_USERNAMES="tugba" → sahipler + Tuğba görür.
export function financeUsernames(): string[] {
  const raw = process.env.FINANCE_USERNAMES;
  const extra = raw
    ? raw.split(",").map((s) => s.trim()).filter(Boolean).map(normalizeUsername)
    : [];
  return [...new Set([...ownerUsernames(), ...extra])];
}

export function isFinance(username: string | undefined | null): boolean {
  if (!username) return false;
  return financeUsernames().includes(normalizeUsername(username));
}

// Finans ekranları (menü + kasa/gider/çek/personel sayfaları) şimdilik kapalı:
// önce sipariş akışına alışılacak. Aktarılan veriler Blob'da durur; açmak için
// Vercel'e FINANS_AKTIF=1 ekleyip yeniden dağıtmak yeterlidir — kod değişmez.
// (Raporlar sayfası bundan bağımsızdır, finans yetkililerine açık kalır.)
export function finansAktif(): boolean {
  return process.env.FINANS_AKTIF === "1";
}

// Maliyet & kârlılık ekranı — alış fiyatları en dar çevrenin bilgisidir.
// MALIYET_USERNAMES ile değiştirilebilir; varsayılan yalnızca Berke ve Özgür.
const DEFAULT_MALIYET = ["berke", "özgür"];

export function maliyetUsernames(): string[] {
  const raw = process.env.MALIYET_USERNAMES;
  const list = raw
    ? raw.split(",").map((s) => s.trim()).filter(Boolean)
    : DEFAULT_MALIYET;
  return list.map(normalizeUsername);
}

export function isMaliyet(username: string | undefined | null): boolean {
  if (!username) return false;
  return maliyetUsernames().includes(normalizeUsername(username));
}
