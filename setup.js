#!/usr/bin/env node
/**
 * ZipMail - Script d'installation initiale
 * ----------------------------------------
 * Configure le certificat HTTPS local et installe les dépendances nécessaires
 * pour le développement de l'add-in Outlook.
 */

import { execSync } from "child_process";
import fs from "fs";
import os from "os";

function run(cmd, description) {
  console.log(`\n🔧 ${description}...`);
  try {
    execSync(cmd, { stdio: "inherit" });
  } catch (err) {
    console.error(`❌ Erreur lors de l'exécution de : ${cmd}`);
    process.exit(1);
  }
}

// Vérifie si les dépendances sont installées
if (!fs.existsSync("node_modules")) {
  run("npm install", "Installation des dépendances npm");
} else {
  console.log("✅ Dépendances déjà installées.");
}

// Installe les certificats de développement Office
run("npx office-addin-dev-certs install --days 365", "Installation des certificats HTTPS de développement");

// Vérifie ou crée le dossier 'dist' pour éviter les erreurs webpack
if (!fs.existsSync("dist")) {
  fs.mkdirSync("dist");
  console.log("📁 Dossier 'dist' créé.");
} else {
  console.log("📁 Dossier 'dist' déjà présent.");
}

// Message de fin
let message = `
🎉 Installation terminée !

Prochaines étapes :
1️⃣  Lancer le serveur de développement :
      npm run dev-server

2️⃣  Ouvrir Outlook et charger le complément via :
      manifest.json
`;

// Ajoute un conseil spécifique pour macOS
if (process.platform === "darwin") {
  message += `
💡 Astuce macOS :
Si Outlook ne parvient pas à se connecter à https://localhost:3000/,
tu peux approuver manuellement le certificat de développement :

    sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain ~/.office-addin-dev-certs/localhost.crt

(Cette opération est à faire une seule fois.)
`;
}

console.log(message);
