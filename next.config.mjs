/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  experimental: {
    // PDF üretiminde kullanılan fontların serverless pakete dahil edilmesi
    outputFileTracingIncludes: {
      "/api/orders": ["./assets/**"],
      "/api/orders/pdf": ["./assets/**"],
      "/api/perakende/orders": ["./assets/**"],
      "/api/perakende/orders/pdf": ["./assets/**"],
      "/api/etiket/pdf": ["./assets/**"],
      "/api/uretim/isler/[id]/foy": ["./assets/**"],
      "/api/ikas/webhook": ["./assets/**"],
      "/api/ikas/senk": ["./assets/**"],
      "/api/uretim/cron": ["./assets/**"],
    },
  },
};

export default nextConfig;
