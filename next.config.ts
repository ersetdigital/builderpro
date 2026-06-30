import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  eslint: {
    // Lint is run separately via `npm run lint`
    ignoreDuringBuilds: true,
  },
  typescript: {
    // Type checking is run separately via `tsc --noEmit`
    ignoreBuildErrors: true,
  },
};

export default nextConfig;
