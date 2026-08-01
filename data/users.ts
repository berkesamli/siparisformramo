// Kullanıcılar — çalışanlar (staff) sipariş paneline, müşteriler (customer) portala girer.
// Üretimde USERS_JSON ortam değişkeni ile bu liste tamamen değiştirilebilir:
// USERS_JSON='[{"username":"ali","password":"...","name":"Ali","role":"staff"}]'

export type Role = "staff" | "customer";

export interface User {
  username: string;
  password: string;
  name: string;
  role: Role;
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

export function findUser(username: string, password: string): User | undefined {
  return getUsers().find(
    (u) => u.username.toLowerCase() === username.toLowerCase() && u.password === password
  );
}
