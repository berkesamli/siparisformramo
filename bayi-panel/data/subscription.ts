// Abonelik durumları — istemci ve sunucuda ortak (Node modülü içermez).
export type SubscriptionStatus = "aktif" | "muaf" | "odeme_bekliyor" | "askida";

export const SUBSCRIPTION_LABELS: Record<SubscriptionStatus, string> = {
  aktif: "Aktif (aylık ödeme)",
  muaf: "Ücretsiz (alım muafiyeti)",
  odeme_bekliyor: "Ödeme bekliyor",
  askida: "Askıda",
};
