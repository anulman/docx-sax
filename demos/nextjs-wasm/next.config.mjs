/** @type {import('next').NextConfig} */
const nextConfig = {
  output: 'standalone',
  transpilePackages: ['@docx-sax/browser'],
};

export default nextConfig;
