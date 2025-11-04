/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// =============================================
// INITIALISATION
// =============================================
Office.onReady(() => {
  console.log("ZipMail commands.js chargé");
});

window.addEventListener("message", (event) => {
  if (event.origin !== "https://localhost:3000") return;
  const [type, key, value] = event.data.split(":");
  if (type === "update") {
    if (key === "level") localStorage.setItem("zipLevel", value);
    if (key === "password") localStorage.setItem("zipPassword", value);
  }
});

Office.actions.associate("onMessageSend", onMessageSend);
Office.actions.associate("onMessageRead", onMessageRead);
Office.actions.associate("disableZip", disableZip);
Office.actions.associate("enableZip", enableZip);
Office.actions.associate("enableZipEncrypted", enableZipEncrypted);

// =============================================
// 0️⃣ CONFIGURATION GLOBALE
// =============================================
const ZIP_MODE_KEY = "zipMode";
const ZIP_LEVEL_KEY = "zipLevel";
const ZIP_PASSWORD_KEY = "zipPassword";

// =============================================
// Aide au déboguage
// =============================================
// Usage: in an async func: await showDialogAlert("Compression terminée !");
// Usage: in an non- async func: showDialogAlert("Compression terminée !");
function showDialogAlert(message) {
  const url = `https://localhost:3000/dialog-alert.html?msg=${encodeURIComponent(message)}`;

  Office.context.ui.displayDialogAsync(url, { height: 30, width: 40 }, (result) => {
    if (result.status === "succeeded") {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => dialog.close());
    }
  });
}

// =============================================
// 1️⃣ Boutons du ruban
// =============================================

function getZipLevel() {
  return localStorage.getItem(ZIP_LEVEL_KEY) || "6";
}

function getZipPassword() {
  return localStorage.getItem(ZIP_PASSWORD_KEY) || "";
}

function getZipMode() {
  return localStorage.getItem(ZIP_MODE_KEY) || "none";
}

// Mise à jour de l'icone du split-button ZipMail en fonction de l'état (non, zip, encrypted)
async function updateRibbonIcon() {
  let icons = {};
  const currentMode = getZipMode();

  if (currentMode === "zip") {
    icons = {
      ZipMailMenu: [
        { size: 16, resid: "Icon.16" },
        { size: 32, resid: "Icon.32" },
        { size: 64, resid: "Icon.64" },
        { size: 80, resid: "Icon.80" },
        { size: 128, resid: "Icon.128" },
      ],
    };
  } else if (currentMode === "encrypted") {
    icons = {
      ZipMailMenu: [
        { size: 16, resid: "IconLocked.16" },
        { size: 32, resid: "IconLocked.32" },
        { size: 64, resid: "IconLocked.64" },
        { size: 80, resid: "IconLocked.80" },
        { size: 128, resid: "IconLocked.128" },
      ],
    };
  } else {
    icons = {
      ZipMailMenu: [
        { size: 16, resid: "IconGreyed.16" },
        { size: 32, resid: "IconGreyed.32" },
        { size: 64, resid: "IconGreyed.64" },
        { size: 80, resid: "IconGreyed.80" },
        { size: 128, resid: "IconGreyed.128" },
      ],
    };
  }
  try {
    await Office.ribbon.requestUpdate({ icons });
  } catch (error) {
    console.error("ZipMail: Échec mise à jour icône:", error);
  }
}

// mode: "none", "zip", "encrypted"
function setZipMode(mode) {
  localStorage.setItem(ZIP_MODE_KEY, mode);
  // Refleter l'état du mode Zip sur l'icone du split button ZipMail du bandeau.
  updateRibbonIcon();
}

function disableZip(event) {
  setZipMode("none");
  showNotification("ZIP désactivé");

  // C'est tout bon.
  event.completed({ allowEvent: true });
}

function enableZip(event) {
  const current = getZipMode();
  const newMode = current === "zip" ? "none" : "zip";

  setZipMode(newMode);
  showNotification(newMode === "zip" ? "ZIP activé" : "ZIP désactivé");

  event.completed({ allowEvent: true });
}

function enableZipEncrypted(event) {
  const current = getZipMode();
  let newMode = "none";

  if (current === "encrypted") {
    newMode = "none"; // désactiver
  } else {
    newMode = "encrypted"; // activer chiffré
  }

  setZipMode(newMode);
  showNotification(newMode === "encrypted" ? "ZIP chiffré activé" : "ZIP désactivé");

  event.completed({ allowEvent: true });
}

// =============================================
// 2️⃣ Envoi du message
// =============================================
async function onMessageSend(event) {
  const mode = getZipMode();
  const zipLevel = getZipLevel();
  const zipPassword = getZipPassword();

  if (mode === "none") {
    event.completed({ allowEvent: true });
    return;
  }

  const isEncrypted = mode === "encrypted";
  const item = Office.context.mailbox.item;

  try {
    // récupère le corps HTML du message (TODO: gérer le cas text mode only du message)
    let bodyHtml = await new Promise((resolve) =>
      item.body.getAsync("html", (res) => resolve(res.value))
    );

    // Insère le <meta name="zipmail"> dans le <head>
    const metaContent = buildZipMailMeta({
      version: "1.0",
      encrypted: isEncrypted,
      timestamp: new Date().toISOString(),
    });
    const metaTag = `<meta name="zipmail" content="${metaContent}">`;
    bodyHtml = bodyHtml.includes("<head>")
      ? bodyHtml.replace("<head>", `<head>${metaTag}`)
      : `<head>${metaTag}</head>${bodyHtml}`;

    // Crée le writer ZIP
    const blobWriter = new zip.BlobWriter("application/zip");
    const zipWriter = new zip.ZipWriter(blobWriter);

    // Options de bases toujours appliquées
    let options = {
      compression: "DEFLATE",
      compressionOptions: { level: parseInt(zipLevel) },
    };

    // Si chiffrement activé, demande l'ajout du mot de passe à la config zip.
    if (isEncrypted) {
      const result = await getPasswordFromDialog(zipPassword);
      if (!result || !result.password) {
        showNotification("Mot de passe requis — envoi annulé");
        await zipWriter.close();
        event.completed({ allowEvent: false });
        return;
      }

      const { password, save } = result;

      // Si coche pour sauver le mot de passe activée; le sauver.
      if (save) {
        localStorage.setItem(ZIP_PASSWORD_KEY, password);
      }

      options = { ...options, password, encryptionStrength: 3 };
    }

    // Ajoute le corps du message dans le Zip.
    await zipWriter.add("message.htm", new zip.TextReader(bodyHtml), options);

    // Ajoutes toutes les pièces jointes dans le Zip.
    let attachments = [];
    try {
      const result = await new Promise((resolve, reject) => {
        item.getAttachmentsAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            resolve(res.value);
          } else {
            reject(new Error("getAttachmentsAsync failed: " + (res.error?.message || "unknown")));
          }
        });
      });
      attachments = result;
    } catch (e) {
      console.error("ZipMail: Impossible de lire les pièces jointes:", e);
      showNotification("Erreur critique : pièces jointes inaccessibles. Envoi bloqué.");
      await zipWriter.close();
      event.completed({ allowEvent: false });
      return;
    }
    for (const att of attachments) {
      const content = await getAttachmentContent(att.id);
      if (content.format === "base64") {
        const bytes = base64ToUint8Array(content.content);
        const blob = new Blob([bytes], { type: content.contentType || "application/octet-stream" });
        await zipWriter.add(att.name, new zip.BlobReader(blob), options);
      }
    }

    // Ferme le zip
    const zipBlob = await zipWriter.close();
    const base64Zip = await blobToBase64(zipBlob);

    // Suppression des anciennes pièces jointes
    for (const att of attachments) {
      try {
        await removeAttachment(att.id);
      } catch (e) {
        console.error("ZipMail: Échec suppression pièce jointe:", att.id, e);
        showNotification("Erreur critique : impossible de supprimer une pièce jointe. Envoi bloqué.");
        await zipWriter.close();
        event.completed({ allowEvent: false });
        return;
      }
    }

    // Ajoute le msg.zip
    try {
      await addAttachmentFromBase64("msg.zip", base64Zip);
    } catch (e) {
      console.error("ZipMail: Échec ajout msg.zip:", e);
      showNotification("Erreur : impossible d’ajouter le ZIP. Envoi bloqué.");
      event.completed({ allowEvent: false });
      return;
    }

    // Remplace le corps du mail par le message générique
    try {
      // 1. Tentative de chargement du modèle
      const response = await fetch("https://localhost:3000/assets/ZipMailMessage.html");
      if (!response.ok) throw new Error(`HTTP ${response.status}`);

      const genericHTML = await response.text();

      // 2. Injection du corps
      await new Promise((resolve) =>
        item.body.setAsync(genericHTML, { coercionType: "html" }, resolve)
      );

      // 3. Envoi autorisé
      event.completed({ allowEvent: true });
    } catch (err) {
      // En cas d'erreur, on bloque l'envoi.
      const errorMsg = `Erreur modèle : ${err.message}`;
      showNotification(errorMsg);
      console.error(errorMsg, err);
      event.completed({ allowEvent: false });
    }
  } catch (err) {
    console.error("Erreur ZipMail:", err);
    showNotification("Erreur ZipMail : " + err.message);
    event.completed({ allowEvent: false });
  }
}

// =============================================
// 3️⃣ Lecture du message
// =============================================
async function onMessageRead(event) {
  const item = Office.context.mailbox.item;
  const msgZip = item.attachments?.find((a) => a.name === "msg.zip");
  if (!msgZip) {
    event.completed();
    return;
  }

  try {
    const zipContent = await getAttachmentContent(msgZip.id);
    const zipBytes = base64ToUint8Array(zipContent.content);

    // Essai sans mot de passe
    let reader = new zip.ZipReader(new zip.BlobReader(new Blob([zipBytes])));
    let entries;
    try {
      entries = await reader.getEntries();
    } catch {
      // probablement chiffré
      const password = await getPasswordFromDialog();
      if (!password) {
        await reader.close();
        event.completed();
        return;
      }
      reader = new zip.ZipReader(new zip.BlobReader(new Blob([zipBytes])), { password });
      entries = await reader.getEntries();
    }

    // Cherches message.htm
    const messageEntry = entries.find((e) => e.filename.toLowerCase() === "message.htm");
    if (!messageEntry) {
      console.log("ZipMail: Pas de message.htm — zip ignoré.");
      await reader.close();
      event.completed();
      return;
    }

    const htmlBody = await messageEntry.getData(new zip.TextWriter());

    // Vérifie le tag ZipMail
    const meta = parseZipMailMeta(htmlBody);
    if (!meta) {
      await reader.close();
      event.completed();
      return;
    }

    // Réinjecte le corps original
    await new Promise((resolve) => item.body.setAsync(htmlBody, { coercionType: "html" }, resolve));

    // Ajoute les autres fichiers comme pièce jointes
    for (const entry of entries) {
      if (entry.filename.toLowerCase() === "message.htm") continue;
      const blob = await entry.getData(new zip.BlobWriter());
      const base64 = await blobToBase64(blob);
      await addAttachmentFromBase64(entry.filename, base64);
    }

    // Supprime le msg.zip
    await removeAttachment(msgZip.id);

    // Active le bon mode
    if (meta.encrypted) enableZipEncrypted({ completed: () => {} });
    else enableZip({ completed: () => {} });

    await reader.close();
    event.completed();
  } catch (err) {
    console.error("ZipMail: Erreur lecture message ZipMail:", err);
    showNotification("Erreur lecture message ZipMail : " + err.message);
    event.completed();
  }
}

// =============================================
// 4️⃣ Helpers
// =============================================

// Parse le tag <meta name="zipmail" ...>
function parseZipMailMeta(html) {
  try {
    const metaMatch = html.match(/<meta\s+name=["']zipmail["']\s+content=["']([^"']+)["']/i);
    if (!metaMatch) return null;

    const metaContent = metaMatch[1];
    const parts = metaContent.split(";").map((p) => p.trim());
    const result = {};

    // Premier élément = version si numérique
    if (parts.length > 0 && /^[0-9.]+$/.test(parts[0])) {
      result.version = parts[0];
      parts.shift();
    }

    for (const part of parts) {
      const [key, value] = part.split("=").map((s) => s.trim());
      if (!key) continue;
      result[key] = value === "true" ? true : value === "false" ? false : value;
    }

    return result;
  } catch {
    return null;
  }
}

// Construit la chaine de meta zipmail
function buildZipMailMeta(obj) {
  const entries = [];
  if (obj.version) entries.push(obj.version);
  for (const [key, value] of Object.entries(obj)) {
    if (key === "version") continue;
    entries.push(`${key}=${value}`);
  }
  return entries.join(";");
}

// --- Pièces jointes & conversions utilitaires ---
function base64ToUint8Array(base64) {
  const raw = atob(base64);
  const array = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) array[i] = raw.charCodeAt(i);
  return array;
}

function blobToBase64(blob) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(",")[1]);
    reader.readAsDataURL(blob);
  });
}

function showNotification(msg) {
  console.log(msg);
  if (Office.context.mailbox.item?.notificationMessages) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("ZipMail", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: msg,
      icon: "icon-zip-16",
      persistent: false,
    });
  }
}

// --- API Outlook pour pièces jointes ---
async function getAttachmentContent(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(id, (res) => {
      res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value) : reject(res.error);
    });
  });
}

async function removeAttachment(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.removeAttachmentAsync(id, (res) => {
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  });
}

async function addAttachmentFromBase64(name, base64) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(base64, name, (res) => {
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  });
}

// --- Boîte de dialogue mot de passe ---
async function getPasswordFromDialog(defaultPassword = "") {
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/password.html",
      { height: 35, width: 25 },
      (asyncResult) => {
        const dialog = asyncResult.value;

        // Envoie le mot de passe par postmessage
        const readyHandler = (arg) => {
          if (arg.message === "ready") {
            dialog.postMessage({ type: "defaultPassword", value: defaultPassword });
            dialog.removeEventHandler(Office.EventType.DialogMessageReceived, readyHandler);
          }
        };
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, readyHandler);

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          try {
            const data = JSON.parse(arg.message);
            resolve({ password: data.password, save: data.save });
          } catch {
            resolve({ password: arg.message, save: false });
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => resolve(null));
      }
    );
  });
}
