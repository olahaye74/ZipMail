/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

let zipEnabled = false;
let zipEncrypted = false;

// =============================================
// INITIALISATION
// =============================================
Office.onReady(() => {
  console.log("ZipMail commands.js chargé");
});

Office.actions.associate("onMessageSend", onMessageSend);
Office.actions.associate("onMessageRead", onMessageRead);
Office.actions.associate("enableZip", enableZip);
Office.actions.associate("enableZipEncrypted", enableZipEncrypted);

// =============================================
// 1️⃣ Boutons du ruban
// =============================================

function updateRibbonState() {
  const control = Office.context.ui.getControl("ZipMailMenu");
  if (!control) return;

  // Bouton enfoncé
  control.isPressed = zipEnabled;

  // Icône : cadenas si chiffré
  const iconSet = zipEncrypted
    ? {
        16: "IconLocked.16",
        32: "IconLocked.32",
        64: "IconLocked.64",
        80: "IconLocked.80",
        128: "IconLocked.128"
      }
    : {
        16: "Icon.16",
        32: "Icon.32",
        64: "Icon.64",
        80: "Icon.80",
        128: "Icon.128"
      };

  control.setIcon(iconSet);
}

function enableZip(event) {
  zipEnabled = !zipEnabled;
  zipEncrypted = false;
  updateRibbonState();
  showNotification(zipEnabled ? "Zip activé" : "Zip désactivé");
  event.completed({ allowEvent: true });
}

function enableZipEncrypted(event) {
  zipEnabled = true;
  zipEncrypted = !zipEncrypted;
  updateRibbonState();
  showNotification(zipEncrypted ? "Zip chiffré activé 🔒" : "Zip normal activé");
  event.completed({ allowEvent: true });
}

// =============================================
// 2️⃣ Envoi du message
// =============================================
async function onMessageSend(event) {
  if (!zipEnabled) {
    event.completed({ allowEvent: true });
    return;
  }

  const item = Office.context.mailbox.item;

  try {
    // Récupère le corps HTML du message
    let bodyHtml = await new Promise((resolve) =>
      item.body.getAsync("html", (res) => resolve(res.value))
    );

    // Insère le <meta name="zipmail"> dans le <head>
    const metaContent = buildZipMailMeta({
      version: "1.0",
      encrypted: zipEncrypted,
      timestamp: new Date().toISOString(),
    });

    const metaTag = `<meta name="zipmail" content="${metaContent}">`;
    if (bodyHtml.includes("<head>")) {
      bodyHtml = bodyHtml.replace("<head>", `<head>${metaTag}`);
    } else {
      bodyHtml = `<head>${metaTag}</head>${bodyHtml}`;
    }

    // Crée le writer ZIP
    const blobWriter = new zip.BlobWriter("application/zip");
    const zipWriter = new zip.ZipWriter(blobWriter);

    // Si chiffrement activé, demande du mot de passe
    let password = null;
    if (zipEncrypted) {
      password = await getPasswordFromDialog();
      if (!password) {
        showNotification("Mot de passe non fourni — envoi annulé");
        event.completed({ allowEvent: false });
        await zipWriter.close();
        return;
      }
    }

    const encryptionOptions = password ? { password, encryptionStrength: 3 } : {};

    // Ajoute le corps dans le zip
    await zipWriter.add("message.htm", new zip.TextReader(bodyHtml), encryptionOptions);

    // Ajoute toutes les pièces jointes existantes
    const attachments = item.attachments || [];
    for (const att of attachments) {
      const content = await getAttachmentContent(att.id);
      if (content.format === "base64") {
        const bytes = base64ToUint8Array(content.content);
        const blob = new Blob([bytes], { type: content.contentType || "application/octet-stream" });
        await zipWriter.add(att.name, new zip.BlobReader(blob), encryptionOptions);
      }
    }

    // Ferme le ZIP
    const zipBlob = await zipWriter.close();
    const base64Zip = await blobToBase64(zipBlob);

    // Supprime les anciennes pièces jointes
    for (const att of attachments) {
      await removeAttachment(att.id);
    }

    // Ajoute msg.zip
    await addAttachmentFromBase64("msg.zip", base64Zip);

    // Remplace le corps du mail par le message générique
    const genericHTML = await (
      await fetch("https://localhost:3000/assets/ZipMailMessage.html")
    ).text();
    await new Promise((resolve) =>
      item.body.setAsync(genericHTML, { coercionType: "html" }, resolve)
    );

    event.completed({ allowEvent: true });
  } catch (err) {
    console.error("Erreur ZipMail (onMessageSend):", err);
    showNotification("Erreur ZipMail : " + err.message);
    event.completed({ allowEvent: false });
  }
}

// =============================================
// 3️⃣ Lecture du message (ouverture d’un mail)
// =============================================
async function onMessageRead(event) {
  const item = Office.context.mailbox.item;
  const msgZip = item.attachments.find((a) => a.name === "msg.zip");
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
      reader = new zip.ZipReader(new zip.BlobReader(new Blob([zipBytes])), { password });
      entries = await reader.getEntries();
    }

    // Cherche message.htm
    const messageEntry = entries.find((e) => e.filename.toLowerCase() === "message.htm");
    if (!messageEntry) {
      console.log("Pas de message.htm — zip ignoré.");
      await reader.close();
      event.completed();
      return;
    }

    const htmlBody = await messageEntry.getData(new zip.TextWriter());

    // Vérifie le tag ZipMail
    const meta = parseZipMailMeta(htmlBody);
    if (!meta) {
      console.log("Pas de tag ZipMail — zip ignoré.");
      await reader.close();
      event.completed();
      return;
    }

    // Réinjecte le corps original
    await new Promise((resolve) => item.body.setAsync(htmlBody, { coercionType: "html" }, resolve));

    // Ajoute les autres fichiers comme pièces jointes
    for (const entry of entries) {
      if (entry.filename.toLowerCase() === "message.htm") continue;
      const blob = await entry.getData(new zip.BlobWriter());
      const base64 = await blobToBase64(blob);
      await addAttachmentFromBase64(entry.filename, base64);
    }

    // Supprime msg.zip
    await removeAttachment(msgZip.id);

    // Active le bon mode
    if (meta.encrypted) enableZipEncrypted({ completed: () => {} });
    else enableZip({ completed: () => {} });

    await reader.close();
    event.completed();
  } catch (err) {
    console.error("Erreur ZipMail (onMessageRead):", err);
    showNotification("Erreur lecture ZipMail : " + err.message);
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
      if (value === "true") result[key] = true;
      else if (value === "false") result[key] = false;
      else result[key] = value;
    }

    return result;
  } catch (e) {
    console.error("Erreur parseZipMailMeta:", e);
    return null;
  }
}

// Construit la chaîne de meta zipmail
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
  const array = new Uint8Array(new ArrayBuffer(raw.length));
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
  if (Office.context.mailbox.item.notificationMessages)
    Office.context.mailbox.item.notificationMessages.replaceAsync("ZipMail", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: msg,
      icon: "icon-zip-16",
      persistent: false,
    });
}

// --- API Outlook pour pièces jointes ---
async function getAttachmentContent(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(id, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(res.error);
    });
  });
}

async function removeAttachment(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.removeAttachmentAsync(id, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

async function addAttachmentFromBase64(name, base64) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(base64, name, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

// --- Boîte de dialogue mot de passe ---
async function getPasswordFromDialog() {
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/password.html",
      { height: 30, width: 30 },
      (asyncResult) => {
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close();
          resolve(arg.message);
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => resolve(null));
      }
    );
  });
}
