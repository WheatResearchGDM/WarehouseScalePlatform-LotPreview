import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // pages/app.js is the standalone GitHub Pages client, not a Next route.
  pageExtensions: ["tsx", "ts"],
};

export default nextConfig;
