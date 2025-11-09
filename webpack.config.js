// webpack.config.js — version ESM compatible Node "type": "module"
import path from "path";
import fs from "fs";
import os from "os";
import { fileURLToPath } from "url";
import CopyWebpackPlugin from "copy-webpack-plugin";
import HtmlWebpackPlugin from "html-webpack-plugin";
import webpack from "webpack";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 🔹 Détermination automatique du dossier utilisateur (Windows/macOS/Linux)
const homeDir =
  process.env.USERPROFILE || // Windows
  process.env.HOME || // macOS / Linux
  os.homedir(); // fallback universel

const certDir = path.resolve(homeDir, ".office-addin-dev-certs");

// 🔹 Chargement conditionnel des certificats HTTPS
let serverOptions;
if (
  fs.existsSync(path.join(certDir, "localhost.key")) &&
  fs.existsSync(path.join(certDir, "localhost.crt")) &&
  fs.existsSync(path.join(certDir, "ca.crt"))
) {
  console.log(`✅ Certificats trouvés dans : ${certDir}`);
  serverOptions = {
    type: "https",
    options: {
      key: fs.readFileSync(path.join(certDir, "localhost.key")),
      cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
      ca: fs.readFileSync(path.join(certDir, "ca.crt")),
    },
  };
} else {
  console.warn(`⚠️ Certificats introuvables dans ${certDir} — passage en HTTP`);
  serverOptions = { type: "http" };
}

export default {
  mode: "development",

  entry: {
    taskpane: "./src/code/taskpane.js",
    compose: "./src/code/compose.js",
    read: "./src/code/read.js",
    options: "./src/code/options.js",
  },

  output: {
    clean: true,
    path: path.resolve(__dirname, "dist"),
    filename: "[name].bundle.js",
    // publicPath: 'auto', - inutile car déjà mis à auto par défaut
  },

  devServer: {
    port: 3000,
    hot: true,
    server: serverOptions,
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    static: {
      directory: path.join(__dirname, "assets"),
    },
    devMiddleware: {
      writeToDisk: true,
    },
  },

  // Génère les fichiers HTML pour chaque entrée
  plugins: [
    // Compatibilité pour les vieux outlook avant 2024
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/pages/commands.html",
      chunks: ["read", "compose"],
    }),
    new HtmlWebpackPlugin({
      filename: "taskpane.html",
      template: "./src/pages/taskpane.html",
      chunks: ["taskpane"],
    }),
    new HtmlWebpackPlugin({
      filename: "compose.html",
      template: "./src/pages/compose.html",
      chunks: ["compose"],
    }),
    new HtmlWebpackPlugin({
      filename: "read.html",
      template: "./src/pages/read.html",
      chunks: ["read"],
    }),
    new HtmlWebpackPlugin({
      filename: "options.html",
      template: "./src/pages/options.html",
      chunks: [],
      inject: false,
    }),
    new HtmlWebpackPlugin({
      filename: "password.html",
      template: "./src/pages/password.html",
      chunks: [], // pas besoin d'injecter de JS
      inject: false,
    }),
    new HtmlWebpackPlugin({
      filename: "dialog-alert.html",
      template: "./src/pages/dialog-alert.html",
      chunks: [],
      inject: false,
    }),

    // Copie le manifeste et les assets
    new CopyWebpackPlugin({
      patterns: [
        { from: "manifest.xml", to: "[name][ext]" },
        { from: "assets", to: "assets" },
      ],
    }),

    // Fournit un polyfill de `process` pour le browser
    new webpack.ProvidePlugin({
      process: "process/browser",
    }),
  ],

  resolve: {
    extensions: [".js"],
    fallback: {
      fs: false,
      path: false,
      os: false,
    },
  },
};
