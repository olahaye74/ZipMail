/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// =============================================
// INITIALIZATION
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
Office.actions.associate("onItemChanged", onItemChanged);
Office.actions.associate("disableZip", disableZip);
Office.actions.associate("enableZip", enableZip);
Office.actions.associate("enableZipEncrypted", enableZipEncrypted);

// =============================================
// 0️⃣ GLOBAL CONFIG
// =============================================
const ZIP_MODE_KEY = "zipMode";
const ZIP_LEVEL_KEY = "zipLevel";
const ZIP_PASSWORD_KEY = "zipPassword";

// ============================================================================
// Volatile in-memory encrypted password storage
// - The cipher (Uint8Array) and iv are kept in volatileEncryptedPassword
// - The AES-GCM key (CryptoKey) is kept in volatileKey
// - Nothing is persisted to disk/localStorage
// ============================================================================
let volatileEncryptedPassword = null; // { cipher: Uint8Array, iv: Uint8Array } or null
let volatileKey = null;               // CryptoKey or null

// Generate a volatile AES-GCM key (kept in memory)
async function generateVolatileKey() {
  volatileKey = await crypto.subtle.generateKey(
    { name: "AES-GCM", length: 256 },
    true,
    ["encrypt", "decrypt"]
  );
  return volatileKey;
}

// Encrypt plaintext password into volatileEncryptedPassword
async function encryptVolatile(password) {
  if (!volatileKey) await generateVolatileKey();
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const enc = new TextEncoder().encode(password);
  const cipherBuffer = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, volatileKey, enc);
  volatileEncryptedPassword = {
    cipher: new Uint8Array(cipherBuffer),
    iv: new Uint8Array(iv),
  };
  return volatileEncryptedPassword;
}

// Decrypt from volatileEncryptedPassword; returns plaintext string or null
async function decryptVolatile() {
  try {
    if (!volatileEncryptedPassword || !volatileKey) return null;
    const plainBuf = await crypto.subtle.decrypt(
      { name: "AES-GCM", iv: volatileEncryptedPassword.iv },
      volatileKey,
      volatileEncryptedPassword.cipher
    );
    const dec = new TextDecoder().decode(plainBuf);
    return dec;
  } catch (e) {
    console.warn("decryptVolatile failed", e);
    return null;
  }
}

// Securely erase volatile encrypted password and key
function clearVolatilePassword() {
  try {
    if (volatileEncryptedPassword && volatileEncryptedPassword.cipher instanceof Uint8Array) {
      volatileEncryptedPassword.cipher.fill(0);
    }
    if (volatileEncryptedPassword && volatileEncryptedPassword.iv instanceof Uint8Array) {
      volatileEncryptedPassword.iv.fill(0);
    }
  } catch (e) {}
  volatileEncryptedPassword = null;
  volatileKey = null;
}

// =============================================
// DEBUG HELPERS
// =============================================
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
// 1️⃣ RIBBON BUTTONS
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

// Update the ZipMail ribbon icon according to current mode (none, zip, encrypted)
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
  // Reflect mode on ribbon icon
  updateRibbonIcon();
}

function disableZip(event) {
  setZipMode("none");
  // clear volatile password too
  clearVolatilePassword();
  showNotification("ZIP désactivé");

  // All done.
  event.completed({ allowEvent: true });
}

function enableZip(event) {
  // Simplified: set mode to "zip"
  setZipMode("zip");
  showNotification("ZIP activé");
  event.completed({ allowEvent: true });
}

async function enableZipEncrypted(event) {
  // Prompt the user for a password (dialog always shows empty field)
  const result = await getPasswordFromDialog();
  if (!result || !result.password) {
    const mode = getZipMode();
    showNotification("Sans mot de passe ZIP chiffré non activé. Mode actuel: " + mode);
    event.completed({ allowEvent: true });
    return;
  }

  // Store password volatile & encrypted (in-memory)
  try {
    await encryptVolatile(result.password);
  } catch (e) {
    console.error("Failed to encrypt volatile password", e);
    showNotification("Erreur stockage du mot de passe — ZIP chiffré non activé");
    event.completed({ allowEvent: true });
    return;
  } finally {
    // wipe plaintext
    try { result.password = null; } catch (e) {}
  }

  setZipMode("encrypted");
  showNotification("ZIP chiffré activé " + result.password);
  event.completed({ allowEvent: true });
}

// =============================================
// 2️⃣ Sending the message
// =============================================
async function onMessageSend(event) {

  const mode = getZipMode();
  const zipLevel = getZipLevel();
  // zipPassword is no longer read from localStorage for encrypted mode
  // const zipPassword = getZipPassword();

  if (mode === "none") {
    event.completed({ allowEvent: true });
    return;
  }

  const isEncrypted = mode === "encrypted";
  const item = Office.context.mailbox.item;

  try {
    // get HTML body
    let bodyHtml = await new Promise((resolve) =>
      item.body.getAsync("html", (res) => resolve(res.value))
    );

    // insert <meta name="zipmail">
    const metaContent = buildZipMailMeta({
      version: "1.0",
      encrypted: isEncrypted,
      timestamp: new Date().toISOString(),
    });
    const metaTag = `<meta name="zipmail" content="${metaContent}">`;
    bodyHtml = bodyHtml.includes("<head>")
      ? bodyHtml.replace("<head>", `<head>${metaTag}`)
      : `<head>${metaTag}</head>${bodyHtml}`;

    // Create zip writer
    const blobWriter = new zip.BlobWriter("application/zip");
    const zipWriter = new zip.ZipWriter(blobWriter);

    // base options
    let options = {
      compression: "DEFLATE",
      compressionOptions: { level: parseInt(zipLevel) },
    };

    // If encryption enabled, retrieve volatile password (do not open dialog here)
    if (isEncrypted) {
      const passwordPlain = await decryptVolatile();
      if (!passwordPlain) {
        showNotification("Mot de passe requis — envoi annulé " + passwordPlain);
        try { await zipWriter.close(); } catch (e) {}
        event.completed({ allowEvent: false });
        return;
      }

      options = { ...options, password: passwordPlain, encryptionStrength: 3 };

      // After setting options, we will ensure to wipe passwordPlain after zipWriter.close()
    }

    // add message body
    await zipWriter.add("message.htm", new zip.TextReader(bodyHtml), options);

    // collect attachments
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
      const message = "ZipMail: Impossible de lire les pièces jointes: " + e;
      console.error(message, e);
      showNotification("Erreur critique : pièces jointes inaccessibles. Envoi bloqué.");
      try { await zipWriter.close(); } catch (e) {}
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }

    // add attachments content
    for (const att of attachments) {
      const content = await getAttachmentContent(att.id);
      if (content.format === "base64") {
        const bytes = base64ToUint8Array(content.content);
        const blob = new Blob([bytes], { type: content.contentType || "application/octet-stream" });
        await zipWriter.add(att.name, new zip.BlobReader(blob), options);
      }
    }

    // close zip
    const zipBlob = await zipWriter.close();
    const base64Zip = await blobToBase64(zipBlob);

    // remove old attachments
    for (const att of attachments) {
      try {
        await removeAttachment(att.id);
      } catch (e) {
        const message = "ZipMail: Échec suppression pièce jointe: " + att.name;
        console.error(message, e);
        showNotification(
          "Erreur critique : impossible de supprimer une pièce jointe. Envoi bloqué."
        );
        event.completed({ allowEvent: false, errorMessage: message });
        return;
      }
    }

    // add msg.zip
    try {
      await addAttachmentFromBase64("msg.zip", base64Zip);
    } catch (e) {
      const message = "ZipMail: Échec ajout msg.zip:" + e;
      console.error(message, e);
      showNotification("Erreur : impossible d’ajouter le ZIP. Envoi bloqué.");
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }

    // replace body with generic message
    try {
      const response = await fetch("https://localhost:3000/assets/ZipMailMessage.html");
      if (!response.ok) throw new Error(`HTTP ${response.status}`);

      const genericHTML = await response.text();

      await new Promise((resolve) =>
        item.body.setAsync(genericHTML, { coercionType: "html" }, resolve)
      );

      // clean sensitive traces AFTER successful zip creation and attachments replaced
      if (isEncrypted) {
        // erase plaintext password from options and volatile storage
        try {
          if (options && options.password) {
            options.password = null;
          }
        } catch (e) {}
        clearVolatilePassword();
      }

      event.completed({ allowEvent: true });
    } catch (err) {
      // On template error, block send and attempt to leave message intact
      const errorMsg = `Erreur modèle : ${err.message}`;
      showNotification(errorMsg);
      console.error(errorMsg, err);
      event.completed({ allowEvent: false, errorMessage: errorMsg });
    }
  } catch (err) {
    const message = "ZipMail: " + err.message;
    console.error("Erreur ZipMail:", err);
    showNotification("Erreur ZipMail : " + err.message);
    event.completed({ allowEvent: false, errorMessage: message });
  }
}

// =============================================
// 3️⃣ Reading the message
// =============================================
async function onItemChanged(event) {
  // Wait for the item to be loaded
  await Office.context.mailbox.item;

  // Trigger automatic unzip/restore
  await onMessageRead(event);
}

// =============================================
// 3️⃣ Reading the message (AUTOMATIC)
// =============================================
async function onMessageRead(event) {
  const item = Office.context.mailbox.item;

  // =============================================
  // 1️⃣ Check HTML body + meta tag
  // =============================================
  let bodyHtml = "";
  try {
    bodyHtml = await new Promise((resolve, reject) => {
      item.body.getAsync("html", (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          resolve(res.value);
        } else {
          reject(new Error("Échec lecture corps HTML"));
        }
      });
    });
  } catch (e) {
    console.log("ZipMail: Message non HTML ou erreur lecture corps → ignoré.", e);
    event.completed?.();
    return;
  }

  const meta = parseZipMailMeta(bodyHtml);
  if (!meta) {
    console.log("ZipMail: Pas de <meta name='zipmail'> → ignoré.");
    event.completed?.();
    return;
  }

  // =============================================
  // 2️⃣ Check presence of msg.zip
  // =============================================
  let attachments = [];
  try {
    const result = await new Promise((resolve) => {
      item.getAttachmentsAsync((res) => {
        resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : []);
      });
    });
    attachments = result;
  } catch (e) {
    console.warn("ZipMail: getAttachmentsAsync échoué → ignoré.", e);
    event.completed?.();
    return;
  }

  const msgZip = attachments.find((a) => a.name === "msg.zip");
  if (!msgZip) {
    showNotification("Erreur ZipMail : msg.zip manquant malgré meta tag.");
    console.error("ZipMail: Meta tag présent mais pas de msg.zip");
    event.completed?.();
    return;
  }

  // =============================================
  // 3️⃣ Read ZIP + password if needed
  // =============================================
  let zipBytes;
  try {
    const zipContent = await getAttachmentContent(msgZip.id);
    zipBytes = base64ToUint8Array(zipContent.content);
  } catch (e) {
    showNotification("Erreur : impossible de lire msg.zip");
    console.error("ZipMail: Échec lecture msg.zip", e);
    event.completed?.();
    return;
  }

  let reader = new zip.ZipReader(new zip.BlobReader(new Blob([zipBytes])));
  let entries = [];

  try {
    entries = await reader.getEntries();
  } catch {
    // Encrypted ZIP -> ask for password (NO saving)
    const result = await getPasswordFromDialog(); // Pas de mot de passe par défaut
    if (!result?.password) {
      showNotification("Mot de passe requis pour décompresser.");
      await reader.close();
      event.completed?.();
      return;
    }

    reader = new zip.ZipReader(new zip.BlobReader(new Blob([zipBytes])), {
      password: result.password,
    });
    try {
      entries = await reader.getEntries();
    } catch (err) {
      showNotification("Mot de passe incorrect.");
      console.error("ZipMail: Déchiffrement échoué", err);
      await reader.close();
      event.completed?.();
      return;
    }
  }

  // =============================================
  // 4️⃣ Check message.htm + extraction
  // =============================================
  const messageEntry = entries.find((e) => e.filename.toLowerCase() === "message.htm");
  if (!messageEntry) {
    showNotification("Erreur : message.htm manquant dans msg.zip");
    console.error("ZipMail: message.htm introuvable");
    await reader.close();
    event.completed?.();
    return;
  }

  let originalHtml;
  try {
    originalHtml = await messageEntry.getData(new zip.TextWriter());
  } catch (e) {
    showNotification("Erreur : impossible de lire message.htm");
    console.error("ZipMail: Échec lecture message.htm", e);
    await reader.close();
    event.completed?.();
    return;
  }

  // =============================================
  // 5️⃣ Reinject body + attachments
  // =============================================
  try {
    await new Promise((resolve) => {
      item.body.setAsync(originalHtml, { coercionType: "html" }, resolve);
    });
  } catch (e) {
    showNotification("Erreur : impossible de restaurer le corps");
    console.error("ZipMail: Échec setAsync corps", e);
    await reader.close();
    event.completed?.();
    return;
  }

  // Re-add attachments
  for (const entry of entries) {
    if (entry.filename.toLowerCase() === "message.htm") continue;

    try {
      const blob = await entry.getData(new zip.BlobWriter());
      const base64 = await blobToBase64(blob);
      await addAttachmentFromBase64(entry.filename, base64);
    } catch (e) {
      console.warn(`ZipMail: Échec ajout pièce jointe ${entry.filename}`, e);
      // Continue — do not block
    }
  }

  // =============================================
  // 6️⃣ Remove msg.zip
  // =============================================
  try {
    await removeAttachment(msgZip.id);
  } catch (e) {
    console.warn("ZipMail: Échec suppression msg.zip", e);
    // Non critical
  }

  // =============================================
  // Finalization
  // =============================================
  await reader.close();

  // Activate correct mode based on meta
  if (meta.encrypted) {
    setZipMode("encrypted");
  } else {
    setZipMode("zip");
  }

  showNotification("Message ZipMail restauré automatiquement.");
  event.completed?.();
}

// =============================================
// 4️⃣ Helpers
// =============================================

// Parse the <meta name="zipmail"> tag
function parseZipMailMeta(html) {
  try {
    const metaMatch = html.match(/<meta\s+name=["']zipmail["']\s+content=["']([^"']+)["']/i);
    if (!metaMatch) return null;

    const metaContent = metaMatch[1];
    const parts = metaContent.split(";").map((p) => p.trim());
    const result = {};

    // First element may be version
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

// Build zipmail meta string
function buildZipMailMeta(obj) {
  const entries = [];
  if (obj.version) entries.push(obj.version);
  for (const [key, value] of Object.entries(obj)) {
    if (key === "version") continue;
    entries.push(`${key}=${value}`);
  }
  return entries.join(";");
}

// --- Attachments & conversions ---
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

// --- Outlook attachment APIs ---
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

// --- Password dialog (SEND & READ) ---
// Opens a dialog and returns { password: string } or null if cancelled/empty.
async function getPasswordFromDialog() {
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/password.html",
      { height: 15, width: 12 },
      (asyncResult) => {
        if (asyncResult.status !== "succeeded") {
          resolve(null);
          return;
        }

        const dialog = asyncResult.value;

        // Response handler: parse JSON or fallback to raw string.
        const responseHandler = (arg) => {
          // arg.message est déjà un objet si envoyé avec { password }
          const data = typeof arg.message === "object" ? arg.message : null;

          // If data.password is present and non-empty string -> return it
          if (data && typeof data.password === "string" && data.password.length > 0) {
            cleanupAndResolve({ password: data.password });
          } else {
            // Empty password (or missing) => treat as cancel/no-password
            cleanupAndResolve(null);
          }
        };

        // Dialog close/cancel handler
        const closeHandler = (arg) => {
          // 12006 = dialog closed by user/host
          if (arg && arg.error === 12006) {
            cleanupAndResolve(null);
          }
        };

        // Clean-up helper: close dialog, remove handlers, resolve promise
        function cleanupAndResolve(value) {
          try {
            dialog.removeEventHandler(Office.EventType.DialogMessageReceived, responseHandler);
          } catch (e) {}
          try {
            dialog.removeEventHandler(Office.EventType.DialogEventReceived, closeHandler);
          } catch (e) {}
          try {
            dialog.close();
          } catch (e) {}
          resolve(value);
        }

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, responseHandler);
        dialog.addEventHandler(Office.EventType.DialogEventReceived, closeHandler);
      }
    );
  });
}

