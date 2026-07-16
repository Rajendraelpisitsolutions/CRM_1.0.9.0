const webpack = require("webpack");

module.exports = {
  webpack: {
    configure: (webpackConfig) => {
      // Ensure fallback object exists
      if (!webpackConfig.resolve) {
        webpackConfig.resolve = {};
      }
      if (!webpackConfig.resolve.fallback) {
        webpackConfig.resolve.fallback = {};
      }

      // Add Node.js module polyfills for browser environment
      webpackConfig.resolve.fallback = {
        ...webpackConfig.resolve.fallback,
        crypto: require.resolve("crypto-browserify"),
        stream: require.resolve("stream-browserify"),
        buffer: require.resolve("buffer/"),
        util: require.resolve("util/"),
        vm: require.resolve("vm-browserify"),
        path: false,
        fs: false,
        net: false,
        tls: false,
      };

      // Add ProvidePlugin for global polyfills
      if (!webpackConfig.plugins) {
        webpackConfig.plugins = [];
      }
      webpackConfig.plugins.push(
        new webpack.ProvidePlugin({
          process: "process",
          Buffer: ["buffer", "Buffer"],
        })
      );

      // Terser minifies in a worker thread that does NOT inherit the parent's
      // --max-old-space-size, so a large bundle OOMs the worker ("JS heap out of memory")
      // even when the build is given plenty of heap. Run it in-process (parallel:false) so the
      // raised NODE_OPTIONS heap applies to minification too.
      if (webpackConfig.optimization && Array.isArray(webpackConfig.optimization.minimizer)) {
        webpackConfig.optimization.minimizer.forEach((m) => {
          if (m && m.constructor && m.constructor.name === "TerserPlugin") {
            m.options = m.options || {};
            m.options.parallel = false;
          }
        });
      }

      return webpackConfig;
    },
  },
};

