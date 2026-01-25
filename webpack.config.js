const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const fs = require("fs");
const os = require("os");

// Check if we're in production mode
const isProduction = process.env.NODE_ENV === 'production' || process.argv.includes('--mode=production') || process.argv.includes('production');

// GitHub Pages URL - update this to your repo
const GITHUB_PAGES_URL = "https://lancedesk.github.io/ms-office-ai-helper/";

// Generate a unique build ID to bust cache
const BUILD_ID = Date.now();

// Get the Office add-in dev certificates (only for development)
let httpsOptions = true;
if (!isProduction) {
  const certPath = path.join(os.homedir(), ".office-addin-dev-certs");
  if (fs.existsSync(path.join(certPath, "localhost.crt"))) {
    httpsOptions = {
      key: fs.readFileSync(path.join(certPath, "localhost.key")),
      cert: fs.readFileSync(path.join(certPath, "localhost.crt")),
      ca: fs.readFileSync(path.join(certPath, "ca.crt"))
    };
  }
}

module.exports = {
  mode: isProduction ? 'production' : 'development',
  // Disable caching entirely in development
  cache: isProduction,
  entry: {
    taskpane: "./src/taskpane/taskpane.js",
    commands: "./src/commands/commands.js"
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    // Use contenthash for production, timestamp for development
    filename: isProduction ? "[name].[contenthash].js" : `[name].${BUILD_ID}.js`,
    clean: true,
    // Use GitHub Pages URL for production, localhost for development
    publicPath: isProduction ? GITHUB_PAGES_URL : "https://localhost:3001/"
  },
  devServer: {
    port: 3001,
    server: {
      type: "https",
      options: httpsOptions
    },
    // Disable HMR - Office WebView doesn't support it well
    hot: false,
    // Enable live reload instead
    liveReload: true,
    // Watch for file changes
    watchFiles: ["src/**/*"],
    headers: {
      "Access-Control-Allow-Origin": "*",
      "Cache-Control": "no-store, no-cache, must-revalidate, proxy-revalidate, max-age=0",
      "Pragma": "no-cache",
      "Expires": "-1",
      "Surrogate-Control": "no-store"
    },
    devMiddleware: {
      // Write files to disk for Office to read
      writeToDisk: true,
      // Disable in-memory caching
      stats: 'minimal'
    },
    // Disable client-side caching
    client: {
      overlay: true,
      progress: true
    }
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/taskpane/taskpane.html",
      chunks: ["taskpane"]
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["commands"]
    })
  ],
  module: {
    rules: [
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"]
      },
      {
        test: /\.(png|jpg|jpeg|gif|svg)$/,
        type: "asset/resource",
        generator: {
          filename: "assets/[name][ext]"
        }
      }
    ]
  },
  resolve: {
    extensions: [".js", ".json"]
  }
};
