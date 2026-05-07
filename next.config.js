/** @type {import('next').NextConfig} */
const nextConfig = {
  experimental: {
    serverComponentsExternalPackages: ['ws', 'bufferutil', 'utf-8-validate'],
    instrumentationHook: true,
  },
};

module.exports = nextConfig;
