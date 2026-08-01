/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  experimental: {
    // PDF üretiminde kullanılan fontların serverless pakete dahil edilmesi
    outputFileTracingIncludes: {
      "/api/orders": ["./assets/**"],
      "/api/orders/pdf": ["./assets/**"],
    },
  },
};

export default nextConfig;
