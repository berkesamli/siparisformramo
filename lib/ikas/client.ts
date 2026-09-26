// ikas Admin API istemcisi (özel uygulama, OAuth client_credentials).
//
// Ortam değişkenleri:
//   IKAS_CLIENT_ID, IKAS_CLIENT_SECRET — ikas paneli → Uygulamalar → Uygulamalarım → Özel uygulama
//   IKAS_STORE          — mağaza adı (olgacerceve → https://olgacerceve.myikas.com). Verilirse jeton
//                         mağazaya özel adresten alınır; boşsa genel https://api.myikas.com adresi.
//   IKAS_API_VERSION    — v1 (varsayılan) ya da v2
//   IKAS_WEBHOOK_KEY    — bizim webhook adresimizi koruyan gizli anahtar (?k=…). ikas webhook'u imzalamaz;
//                         gelen olay yalnızca tetikleyicidir, sipariş her zaman API'den yeniden okunur.
//
// Jeton 4 saat geçerlidir; süreç içinde önbelleğe alınır. GraphQL hataları
// "Cannot query field X" derse o alan seçimden düşürülüp bir kez yeniden denenir
// (şema sürümleri arasında createdAt/updatedAt gibi alanlar değişebiliyor).

export interface IkasAdres {
  firstName?: string | null; lastName?: string | null; phone?: string | null; company?: string | null;
  addressLine1?: string | null; addressLine2?: string | null; postalCode?: string | null;
  city?: { id?: string | null; code?: string | null; name?: string | null } | null;
  district?: { id?: string | null; code?: string | null; name?: string | null } | null;
  country?: { id?: string | null; code?: string | null; name?: string | null } | null;
}

export interface IkasSatirSecenek { name: string; type?: string | null; values?: { name?: string | null; price?: number | null; value?: string | null }[] | null }

export interface IkasSatir {
  id: string;
  quantity: number;
  price?: number | null;
  finalPrice?: number | null;
  status?: string | null;
  options?: IkasSatirSecenek[] | null;
  variant: { id?: string | null; productId?: string | null; name: string; sku?: string | null; barcodeList?: string[] | null };
}

export interface IkasOrder {
  id: string;
  orderNumber?: string | null;
  orderedAt?: number | string | null;
  createdAt?: number | string | null;
  updatedAt?: number | string | null;
  dueDate?: number | string | null;
  status?: string | null;
  orderPackageStatus?: string | null;
  orderPaymentStatus?: string | null;
  shippingMethod?: string | null;
  currencyCode?: string | null;
  totalPrice?: number | null;
  totalFinalPrice?: number | null;
  itemCount?: number | null;
  note?: string | null;
  salesChannelId?: string | null;
  salesChannel?: { id?: string | null; name?: string | null; type?: number | string | null } | null;
  customerId?: string | null;
  customer?: { id?: string | null; firstName?: string | null; lastName?: string | null; fullName?: string | null; email?: string | null; phone?: string | null } | null;
  attributes?: { orderAttributeId?: string | null; orderAttributeOptionId?: string | null; value?: string | null }[] | null;
  shippingAddress?: IkasAdres | null;
  billingAddress?: IkasAdres | null;
  orderLineItems: IkasSatir[];
  orderPackages?: { id: string; orderPackageNumber?: string | null; orderPackageFulfillStatus?: string | null; orderLineItemIds?: string[] | null; trackingInfo?: { cargoCompany?: string | null; trackingNumber?: string | null; trackingLink?: string | null } | null }[] | null;
}

export function ikasConfigured(): boolean {
  return Boolean((process.env.IKAS_CLIENT_ID || "").trim() && (process.env.IKAS_CLIENT_SECRET || "").trim());
}

export function ikasWebhookKey(): string {
  return (process.env.IKAS_WEBHOOK_KEY || "").trim();
}

function jetonUrl(): string {
  const store = (process.env.IKAS_STORE || "").trim().replace(/^https?:\/\//, "").replace(/\.myikas\.com.*$/, "");
  return store ? `https://${store}.myikas.com/api/admin/oauth/token` : "https://api.myikas.com/api/admin/oauth/token";
}

function graphqlUrl(): string {
  const v = (process.env.IKAS_API_VERSION || "v1").trim().replace(/[^a-z0-9]/gi, "") || "v1";
  return `https://api.myikas.com/api/${v}/admin/graphql`;
}

let jeton: { token: string; bitis: number } | null = null;

export async function ikasJeton(zorla = false): Promise<string> {
  if (!ikasConfigured()) throw new Error("ikas bağlantısı ayarlanmamış (IKAS_CLIENT_ID / IKAS_CLIENT_SECRET).");
  if (jeton && !zorla && Date.now() < jeton.bitis) return jeton.token;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: (process.env.IKAS_CLIENT_ID || "").trim(),
    client_secret: (process.env.IKAS_CLIENT_SECRET || "").trim(),
  });
  const r = await fetch(jetonUrl(), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded", "User-Agent": "olga-uretim-takvimi" },
    body: body.toString(),
    signal: AbortSignal.timeout(20_000),
  });
  const j = (await r.json().catch(() => null)) as { access_token?: string; expires_in?: number; error?: string; error_description?: string } | null;
  if (!r.ok || !j?.access_token) {
    throw new Error(`ikas jetonu alınamadı (${r.status}): ${j?.error_description || j?.error || "yanıt boş"}`);
  }
  jeton = { token: j.access_token, bitis: Date.now() + (Number(j.expires_in) || 14400) * 1000 - 60_000 };
  return jeton.token;
}

// Şema sürümüne göre var olmayabilecek alanlar hata verince seçimden düşürülür
const kaldirilan = new Set<string>();

function alanlariUygula(sorgu: string): string {
  if (!kaldirilan.size) return sorgu;
  return sorgu
    .split("\n")
    .filter((satir) => {
      const ad = satir.trim().split(/[\s({]/)[0];
      return !kaldirilan.has(ad);
    })
    .join("\n");
}

export interface GqlHata { message?: string; extensions?: Record<string, unknown> }

/** GraphQL isteği; 401'de jeton bir kez tazelenir, bilinmeyen alan bir kez düşürülüp yeniden denenir. */
export async function ikasGql<T = any>(sorgu: string, degiskenler: Record<string, unknown> = {}, deneme = 0): Promise<T> {
  const token = await ikasJeton(deneme > 0);
  const r = await fetch(graphqlUrl(), {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}`, "User-Agent": "olga-uretim-takvimi" },
    body: JSON.stringify({ query: alanlariUygula(sorgu), variables: degiskenler }),
    signal: AbortSignal.timeout(30_000),
  });
  if (r.status === 401 && deneme === 0) return ikasGql<T>(sorgu, degiskenler, 1);
  if (r.status === 429) throw new Error("ikas hız sınırı (429): kısa süre sonra yeniden deneyin.");
  const j = (await r.json().catch(() => null)) as { data?: T; errors?: GqlHata[] } | null;
  if (!r.ok || !j) throw new Error(`ikas API hatası (${r.status}).`);
  if (j.errors?.length) {
    const msg = j.errors.map((e) => e.message || "").join(" | ");
    const m = /Cannot query field "([A-Za-z0-9_]+)"/.exec(msg);
    if (m && deneme < 3 && !kaldirilan.has(m[1])) {
      kaldirilan.add(m[1]);
      return ikasGql<T>(sorgu, degiskenler, deneme + 1);
    }
    throw new Error(`ikas GraphQL: ${msg}`);
  }
  return j.data as T;
}

export const SIPARIS_ALANLARI = `
      id
      orderNumber
      orderedAt
      createdAt
      updatedAt
      dueDate
      status
      orderPackageStatus
      orderPaymentStatus
      shippingMethod
      currencyCode
      totalPrice
      totalFinalPrice
      itemCount
      note
      salesChannelId
      salesChannel { id name type }
      customerId
      customer { id firstName lastName fullName email phone }
      attributes { orderAttributeId orderAttributeOptionId value }
      shippingAddress {
        firstName lastName phone company addressLine1 addressLine2 postalCode
        city { id code name }
        district { id code name }
        country { id code name }
      }
      billingAddress {
        firstName lastName phone company addressLine1 addressLine2 postalCode
        city { id code name }
        district { id code name }
        country { id code name }
      }
      orderLineItems {
        id
        quantity
        price
        finalPrice
        status
        options { name type values { name price value } }
        variant { id productId name sku barcodeList }
      }
      orderPackages {
        id
        orderPackageNumber
        orderPackageFulfillStatus
        orderLineItemIds
        trackingInfo { cargoCompany trackingNumber trackingLink }
      }`;

const LISTE = `
query OlgaSiparisler($pagination: PaginationInput, $orderedAt: DateFilterInput, $updatedAt: DateFilterInput, $id: StringFilterInput, $orderNumber: StringFilterInput, $sort: String) {
  listOrder(pagination: $pagination, orderedAt: $orderedAt, updatedAt: $updatedAt, id: $id, orderNumber: $orderNumber, sort: $sort) {
    count
    page
    limit
    hasNext
    data {${SIPARIS_ALANLARI}
    }
  }
}`;

export interface IkasListeSonuc { data: IkasOrder[]; hasNext: boolean; count: number; page: number }

export async function ikasSiparisListesi(v: {
  page?: number; limit?: number; orderedAtGte?: number; updatedAtGte?: number; id?: string; orderNumber?: string; sort?: string;
}): Promise<IkasListeSonuc> {
  const degiskenler: Record<string, unknown> = {
    pagination: { page: v.page || 1, limit: Math.min(200, Math.max(1, v.limit || 50)) },
    sort: v.sort || "-orderedAt",
  };
  if (v.orderedAtGte) degiskenler.orderedAt = { gte: v.orderedAtGte };
  if (v.updatedAtGte) degiskenler.updatedAt = { gte: v.updatedAtGte };
  if (v.id) degiskenler.id = { eq: v.id };
  if (v.orderNumber) degiskenler.orderNumber = { eq: v.orderNumber };
  const d = await ikasGql<{ listOrder: IkasListeSonuc }>(LISTE, degiskenler);
  const l = d?.listOrder;
  return { data: l?.data || [], hasNext: Boolean(l?.hasNext), count: Number(l?.count) || 0, page: Number(l?.page) || 1 };
}

export async function ikasSiparis(id: string): Promise<IkasOrder | null> {
  const r = await ikasSiparisListesi({ id, limit: 1 });
  return r.data[0] || null;
}

export async function ikasSiparisNo(orderNumber: string): Promise<IkasOrder | null> {
  const r = await ikasSiparisListesi({ orderNumber: String(orderNumber).replace(/^#/, ""), limit: 1 });
  return r.data[0] || null;
}

export async function ikasBen(): Promise<{ id: string } | null> {
  const d = await ikasGql<{ me: { id: string } }>("query { me { id } }");
  return d?.me || null;
}

export interface IkasWebhook { id: string; scope: string; endpoint: string }

export async function ikasWebhookListesi(): Promise<IkasWebhook[]> {
  const d = await ikasGql<{ listWebhook: IkasWebhook[] }>("query { listWebhook { id scope endpoint } }");
  return d?.listWebhook || [];
}

export async function ikasWebhookKur(endpoint: string, scopes: string[] = ["store/order/created", "store/order/updated"]): Promise<IkasWebhook[]> {
  const d = await ikasGql<{ saveWebhook: IkasWebhook[] }>(
    "mutation OlgaWebhook($input: WebhookInput!) { saveWebhook(input: $input) { id scope endpoint } }",
    { input: { scopes, endpoint } }
  );
  return d?.saveWebhook || [];
}

export async function ikasWebhookSil(scopes: string[] = ["store/order/created", "store/order/updated"]): Promise<boolean> {
  const d = await ikasGql<{ deleteWebhook: boolean }>("mutation OlgaWebhookSil($scopes: [String!]!) { deleteWebhook(scopes: $scopes) }", { scopes });
  return Boolean(d?.deleteWebhook);
}

/** Zaman damgası (ms sayı ya da ISO metin) → ISO metin; yoksa null. */
export function ikasZaman(v: number | string | null | undefined): string | null {
  if (v === null || v === undefined || v === "") return null;
  const n = typeof v === "number" ? v : /^\d+$/.test(String(v)) ? Number(v) : Date.parse(String(v));
  return Number.isFinite(n) && n > 0 ? new Date(n).toISOString() : null;
}
