/** @type {import('next').NextConfig} */
const nextConfig = {
  experimental: {
    // Garante que exceljs e jszip sejam resolvidos no ambiente do server
    serverComponentsExternalPackages: ['exceljs', 'jszip']
  }
};

export default nextConfig;
