/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  experimental: {
    // PDF üretiminde kullanılan fontların serverless pakete dahil edilmesi
    outputFileTracingIncludes: {
      "/api/siparisler": ["./assets/**"],
      "/api/siparisler/pdf": ["./assets/**"],
    },
  },
};

export default nextConfig;
