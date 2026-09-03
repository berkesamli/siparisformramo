// Olga yönetici hesapları — ortam değişkeninden okunur (bayi tanımlayan kişiler).
export interface AdminUser {
  username: string;
  password: string;
  name: string;
}

export function getAdmins(): AdminUser[] {
  const raw = process.env.ADMIN_USERS_JSON;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length) return parsed as AdminUser[];
    } catch {
      /* hatalı JSON — tekil değişkenlere düş */
    }
  }
  const u = process.env.ADMIN_USERNAME;
  const p = process.env.ADMIN_PASSWORD;
  if (u && p) return [{ username: u, password: p, name: "Olga Yönetici" }];
  return [];
}

export function findAdmin(username: string, password: string): AdminUser | undefined {
  const q = username.trim().toLowerCase();
  return getAdmins().find((a) => a.username.toLowerCase() === q && a.password === password);
}
