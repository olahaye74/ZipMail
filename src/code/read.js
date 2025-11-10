
import * as zmutils from "./zmutils.js";


// src/code/read.js
Office.initialize = () => {};

Office.actions.associate("onItemChanged", onItemChanged);
Office.actions.associate("onItemRead", onItemRead);

// =============================================
// 6️⃣ ON READ
// =============================================
async function onItemChanged(event) {
  await Office.context.mailbox.item;
  await onItemRead(event);
  event.completed();
}

async function onItemRead(event) {
  const item = Office.context.mailbox.item;

  let bodyHtml = "";
  try {
    bodyHtml = await new Promise((resolve, reject) => {
      item.body.getAsync("html", (res) => res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value) : reject(new Error("Échec lecture corps HTML")));
    });
  } catch {
    event.completed?.();
    return;
  }

  const meta = zmutils.parseZipMailMeta(body.value);
  if (!meta) {
    event.completed?.();
    return;
  }
  if (!shouldAutoUnzip(meta)) {
    event.completed?.();
    return;
  }

  let attachments = [];
  try {
    attachments = await new Promise((resolve) => {
      item.getAttachmentsAsync((res) => resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : []));
    });
  } catch {}

  const msgZip = attachments.find((a) => a.name === "msg.zip");
  if (!msgZip) {
    zmutils.showNotification("msg.zip manquant malgré meta tag. Filtré par antivirus? => Impossible d'afficher le mail. ");
    event.completed?.();
    return;
  }

  let zipBytes;
  try {
    const zipContent = await zmutils.getAttachmentContent(msgZip.id);
    zipBytes = zmutils.base64ToUint8Array(zipContent.content);
  } catch {
    zmutils.showNotification("Impossible de lire msg.zip. Fichier corrompu ou absent.");
    event.completed?.();
    return;
  }

  let reader = new zip.ZipReader(new zip.BlobReader(new Blob([zipBytes])));
  let entries = [];
  try {
    entries = await reader.getEntries();
  } catch {
    const result = await zmutils.getPasswordFromDialog();
    if (!result?.password) {
      zmutils.showNotification("Mot de passe requis. Friendly display aborted.");
      await reader.close();
      event.completed?.();
      return;
    }
    reader = new zip.ZipReader(new zip.BlobReader(new Blob([zipBytes])), { password: result.password });
    entries = await reader.getEntries().catch((err) => {
      zmutils.showNotification("Mot de passe incorrect. Friendly display aborted.");
      console.error(err);
      return [];
    });
  }

  const messageEntry = entries.find((e) => e.filename.toLowerCase() === "message.htm");
  if (!messageEntry) {
    zmutils.showNotification("message.htm manquant dans msg.zip");
    await reader.close();
    event.completed?.();
    return;
  }

  const originalHtml = await messageEntry.getData(new zip.TextWriter());
  await item.body.setAsync(originalHtml, { coercionType: "html" }, () => {});

  for (const entry of entries) {
    if (entry.filename.toLowerCase() === "message.htm") continue;
    const blob = await entry.getData(new zip.BlobWriter());
    const base64 = await zmutils.blobToBase64(blob);
    await zmutils.addAttachmentFromBase64(entry.filename, base64);
  }

  await zmutils.removeAttachment(msgZip.id);
  await reader.close();

  // TODO: add this login in copose: if replying, setzipmod accrding to meta infos
  // Here, it has no means
  // if (meta.encrypted) setZipMode("encrypted");
  // else setZipMode("zip");

  zmutils.showNotification("Message ZipMail restauré automatiquement.");

  event.completed?.();
  return;
}

function shouldAutoUnzip(meta) {
  const setting = localStorage.getItem("zipmail_auto_unzip") || "if-unencrypted";
  if (setting === "never") return false;
  if (setting === "always") return true;
  return !meta.encrypted;
}
