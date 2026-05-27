import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  webpack: (config, { isServer }) => {
    if (!isServer) {
      config.resolve.fallback = {
        ...config.resolve.fallback,
        buffer: require.resolve('buffer/'),
        stream: require.resolve('stream-browserify'),
        path: require.resolve('path-browserify'),
        assert: require.resolve('assert/'),
        zlib: require.resolve('browserify-zlib'),
        constants: require.resolve('constants-browserify'),
      };
    }
    return config;
  },
};

export default nextConfig;
