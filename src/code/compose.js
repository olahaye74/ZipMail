/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as zmutils from "./zmutils.js";

// =============================================
// INITIALIZATION
// =============================================
Office.onReady(() => {
  console.log("ZipMail commands.js chargé");
});

// Get values from taskpane (options)
window.addEventListener("message", (event) => {
  if (event.origin !== "https://localhost:3000") return;
  const [type, key, value] = event.data.split(":");
  if (type === "update") {
    if (key === "level") localStorage.setItem("zipLevel", value);
  }
});

Office.actions.associate("onMessageSend", onMessageSend);
Office.actions.associate("disableZip", disableZip);
Office.actions.associate("enableZip", enableZip);
Office.actions.associate("enableZipEncrypted", enableZipEncrypted);

// =============================================
// 0️⃣ GLOBAL CONFIG
// =============================================
const ZIP_MODE_KEY = "zipMode";
const ZIP_LEVEL_KEY = "zipLevel";
const ZIP_PASSWORD_KEY = "zipmailPassword"; // persistent encrypted password storage

// =============================================
// CUSTOM PROPERTIES HELPERS
// =============================================

/**
 * Load custom properties object for current item.
 * Returns the props object or null on failure.
 */
async function getCustomProperties() {
  return new Promise((resolve) => {
    try {
      Office.context.mailbox.item.loadCustomPropertiesAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve(asyncResult.value);
        } else {
          console.error("ZipMail: failed to load customProperties:", asyncResult.error);
          resolve(null);
        }
      });
    } catch (e) {
      console.error("ZipMail: exception in getCustomProperties:", e);
      resolve(null);
    }
  });
}

/**
 * Get a single custom property value for the current item.
 * Returns the value (as stored) or null if absent / on error.
 */
async function getCustomProperty(key) {
  try {
    const props = await getCustomProperties();
    if (!props) return null;
    return props.get(key) ?? null;
  } catch (e) {
    console.error("ZipMail: getCustomProperty error:", e);
    return null;
  }
}

/**
 * Set a single custom property value for the current item and save it.
 * Returns true on success, false on failure.
 */
async function setCustomProperty(key, value) {
  try {
    const props = await getCustomProperties();
    if (!props) return false;
    props.set(key, value);
    return await new Promise((resolve) => {
      props.saveAsync((saveResult) => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) resolve(true);
        else {
          console.error("ZipMail: setCustomProperty saveAsync failed:", saveResult.error);
          resolve(false);
        }
      });
    });
  } catch (e) {
    console.error("ZipMail: setCustomProperty exception:", e);
    return false;
  }
}

/**
 * Remove a custom property (key) and persist change.
 * Returns true on success, false on failure.
 */
async function removeCustomProperty(key) {
  try {
    const props = await getCustomProperties();
    if (!props) return false;
    props.remove(key);
    return await new Promise((resolve) => {
      props.saveAsync((saveResult) => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) resolve(true);
        else {
          console.error("ZipMail: removeCustomProperty saveAsync failed:", saveResult.error);
          resolve(false);
        }
      });
    });
  } catch (e) {
    console.error("ZipMail: removeCustomProperty exception:", e);
    return false;
  }
}

// =============================================
// 1️⃣ RIBBON BUTTONS
// =============================================
function getZipLevel() {
  return localStorage.getItem(ZIP_LEVEL_KEY) || "6";
}

/**
 * Get the current Zip mode for this mail.
 * Returns: "none" | "zip" | "encrypted"
 */
async function getZipMode() {
  try {
    const mode = await getCustomProperty(ZIP_MODE_KEY);
    return mode || "none";
  } catch (e) {
    console.error("ZipMail: getZipMode error:", e);
    return "none";
  }
}

/**
 * Set the current Zip mode for this mail and update icon only if save succeeds.
 * Returns true if saved successfully, false otherwise.
 */
async function setZipMode(mode) {
  try {
    const ok = await setCustomProperty(ZIP_MODE_KEY, mode);
    if (ok) {
      try {
        await updateRibbonIcon(mode);
      } catch (e) {
        // updateRibbonIcon failing should not prevent mode being considered saved
        console.warn("ZipMail: updateRibbonIcon failed after setZipMode:", e);
      }
      return true;
    } else {
      return false;
    }
  } catch (e) {
    console.error("ZipMail: setZipMode exception:", e);
    return false;
  }
}

async function updateRibbonIcon(mode) {
  let icons = {};
  if (mode === "zip") {
    icons = {
      ZipMailMenu: [
        { size: 16, resid: "Icon.16" },
        { size: 32, resid: "Icon.32" },
        { size: 64, resid: "Icon.64" },
        { size: 80, resid: "Icon.80" },
        { size: 128, resid: "Icon.128" },
      ],
    };
  } else if (mode === "encrypted") {
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

async function disableZip(event) {
  try {
    await setZipMode("none");
    await clearPasswordStorage();
    zmutils.showNotification("ZIP désactivé. Mode: " + await getZipMode());
    event.completed({ allowEvent: true });
  } catch (e) {
    console.error("Erreur disableZip:", e);
    event.completed({ allowEvent: true, errorMessage: e.message });
  }
}

async function enableZip(event) {
  try {
    await setZipMode("zip");
    zmutils.showNotification("ZIP activé. Mode: " + await getZipMode());
    await clearPasswordStorage(); // We don't need a password. Clear it if we switched from encryted to normal
    event.completed({ allowEvent: true });
  } catch (e) {
    console.error("Erreur enableZip:", e);
    event.completed({ allowEvent: true, errorMessage: e.message });
  }
}

// =============================================
// 2️⃣ PASSWORD STORAGE
// Password storage (encrypted) in customProperties
// =============================================
/*
 Design:
 - Generate a random AES-GCM key (CryptoKey).
 - Encrypt password with AES-GCM (random IV).
 - Export key as JWK and store together with cipher+iv in customProperties under ZIP_PASSWORD_KEY.
 - To retrieve: import JWK, decrypt cipher with IV.
 - To clear: remove the custom property.
 
 Security note:
 - The key is stored alongside the cipher in customProperties to allow decrypting in other add-in contexts.
 - This provides per-item isolation (no global mixing), and avoids keeping plaintext in JS memory longer than necessary.
 - However, storing key+cipher in item properties means anyone with access to the item properties (or mailbox) could extract them.
 - For stronger security, use a server-side KMS or user-derived secret.
*/

// Helper to convert Uint8Array to plain JS array for JSON serialization
function u8ToArray(u8) {
  return Array.from(u8);
}

// Helper to convert plain JS array back to Uint8Array
function arrayToU8(arr) {
  return new Uint8Array(arr);
}

/**
 * Save password encrypted into customProperties for current item.
 * Returns true on success, false on failure.
 */
async function savePasswordToStorage(password) {
  try {
    if (!password) return false;

    // generate AES-GCM key
    const key = await crypto.subtle.generateKey({ name: "AES-GCM", length: 256 }, true, ["encrypt", "decrypt"]);
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const encoded = new TextEncoder().encode(password);
    const cipherBuffer = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, encoded);

    // export key as JWK (JSON) so we can import it back in another context
    const jwk = await crypto.subtle.exportKey("jwk", key);

    const stored = {
      cipher: u8ToArray(new Uint8Array(cipherBuffer)),
      iv: u8ToArray(iv),
      key: jwk, // JWK object
    };

    // store JSON string in customProperties
    const ok = await setCustomProperty(ZIP_PASSWORD_KEY, JSON.stringify(stored));

    // wipe local references (best-effort)
    // (can't zero out subtle CryptoKey directly; allow GC)
    return !!ok;
  } catch (e) {
    console.error("ZipMail: savePasswordToStorage failed:", e);
    return false;
  }
}

/**
 * Retrieve and decrypt password from customProperties for current item.
 * Returns plaintext password string or null on failure.
 */
async function getPasswordFromStorage() {
  try {
    const raw = await getCustomProperty(ZIP_PASSWORD_KEY);
    if (!raw) return null;

    let stored;
    try {
      stored = typeof raw === "string" ? JSON.parse(raw) : raw;
    } catch (e) {
      console.error("ZipMail: stored password parse error:", e);
      return null;
    }

    if (!stored || !stored.key || !stored.cipher || !stored.iv) return null;

    // import key from JWK
    const key = await crypto.subtle.importKey("jwk", stored.key, { name: "AES-GCM" }, true, ["decrypt"]);
    const cipher = arrayToU8(stored.cipher).buffer;
    const iv = arrayToU8(stored.iv);

    const plainBuf = await crypto.subtle.decrypt({ name: "AES-GCM", iv }, key, cipher);
    const password = new TextDecoder().decode(plainBuf);

    // best-effort wipe of intermediate buffers (allow GC)
    // no direct secure wipe possible for CryptoKey or ArrayBuffers in JS
    return password;
  } catch (e) {
    console.error("ZipMail: getPasswordFromStorage failed:", e);
    return null;
  }
}

/**
 * Remove encrypted password from customProperties for current item.
 * Returns true on success, false otherwise.
 */
async function clearPasswordStorage() {
  try {
    const ok = await removeCustomProperty(ZIP_PASSWORD_KEY);
    return !!ok;
  } catch (e) {
    console.error("ZipMail: clearPasswordStorage failed:", e);
    return false;
  }
}

// =============================================
// 4️⃣ ENABLE ENCRYPTED ZIP
// =============================================
async function enableZipEncrypted(event) {
  try {
    let password = await zmutils.getPasswordFromDialog();
    if (!password) {
      zmutils.showNotification("Activation ZIP chiffré annulée.");
      event.completed({ allowEvent: true });
      return;
    }

    await savePasswordToStorage(password);
    await setZipMode("encrypted");
    zmutils.showNotification("ZIP chiffré activé. Mode: " + await getZipMode());
    event.completed({ allowEvent: true });
  } catch (e) {
    console.error("enableZipEncrypted error:", e);
    zmutils.showNotification("Erreur activation ZIP chiffré.");
    event.completed({ allowEvent: true, errorMessage: e.message });
  }
}

// =============================================
// 5️⃣ ON SEND
// =============================================
async function onMessageSend(event) {
  const mode = await getZipMode();
  zmutils.showNotification("ZipMode: " + mode);
  if (mode === "none") {
    event.completed({ allowEvent: true });
    return;
  }

  const isEncrypted = mode === "encrypted";
  const zipLevel = parseInt(getZipLevel());
  const item = Office.context.mailbox.item;

  try {
    let bodyHtml = await new Promise((resolve) =>
      item.body.getAsync("html", (res) => resolve(res.value))
    );

    const metaTag = `<meta name="zipmail" content="${buildZipMailMeta({ version: "1.0", encrypted: isEncrypted, timestamp: new Date().toISOString() })}">`;
    bodyHtml = bodyHtml.includes("<head>") ? bodyHtml.replace("<head>", `<head>${metaTag}`) : `<head>${metaTag}</head>${bodyHtml}`;

    const blobWriter = new zip.BlobWriter("application/zip");
    const zipWriter = new zip.ZipWriter(blobWriter);

    let options = { compression: "DEFLATE", compressionOptions: { level: zipLevel } };

    if (isEncrypted) {
      const password = await getPasswordFromStorage();
      if (!password) {
        zmutils.showNotification("Mot de passe requis — envoi annulé.");
        await zipWriter.close();
        event.completed({ allowEvent: false });
        return;
      }
      options = { ...options, password, encryptionStrength: 3 };
    }

    await zipWriter.add("message.htm", new zip.TextReader(bodyHtml), options);
    const attachments = await new Promise((resolve, reject) => {
      item.getAttachmentsAsync((res) => res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value) : reject(res.error));
    });

    for (const att of attachments) {
      const content = await zmutils.getAttachmentContent(att.id);
      const bytes = zmutils.base64ToUint8Array(content.content);
      const blob = new Blob([bytes], { type: content.contentType || "application/octet-stream" });
      await zipWriter.add(att.name, new zip.BlobReader(blob), options);
    }

    const zipBlob = await zipWriter.close();
    const base64Zip = await zmutils.blobToBase64(zipBlob);

    for (const att of attachments) {
      await zmutils.removeAttachment(att.id);
    }

    await zmutils.addAttachmentFromBase64("msg.zip", base64Zip);

    // Clear password only after ZIP successfully created
    if (isEncrypted) {
      await clearPasswordStorage();
    }

    // Clear ZipMode from custom properties to prevent reencoding an aborted send
    await removeCustomProperty(ZIP_MODE_KEY);

    const response = await fetch("https://localhost:3000/assets/ZipMailMessage.html");
    const genericHTML = await response.text();
    await new Promise((resolve) => item.body.setAsync(genericHTML, { coercionType: "html" }, resolve));

    event.completed({ allowEvent: true });
  } catch (e) {
    console.error("onMessageSend failed:", e);
    zmutils.showNotification("Erreur ZipMail : " + e.message);
    event.completed({ allowEvent: false, errorMessage: e.message });
  }
}

// =============================================
// 7️⃣ HELPERS
// =============================================
function buildZipMailMeta(obj) {
  const entries = [];
  if (obj.version) entries.push(obj.version);
  for (const [key, value] of Object.entries(obj)) {
    if (key === "version") continue;
    entries.push(`${key}=${value}`);
  }
  return entries.join(";");
}

