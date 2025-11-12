import * as zmutils from "./zmutils.js";

// src/code/read.js
Office.initialize = () => {};

Office.actions.associate("onItemChanged", onItemChanged);
Office.actions.associate("onItemRead", onItemRead);
Office.actions.associate("openOptionsDialog", openOptionsDialog);

const DIALOG_KEY = "_zipMailOptionsDialog";

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
      item.body.getAsync("html", (res) =>
        res.status === Office.AsyncResultStatus.Succeeded
          ? resolve(res.value)
          : reject(new Error("Échec lecture corps HTML"))
      );
    });
  } catch {
    event.completed?.();
    return;
  }

  const meta = zmutils.parseZipMailMeta(bodyHtml.value);
  if (!meta) {
    event.completed?.();
    return;
  }
  if (!shouldAutoUnzip(meta)) {
    event.completed?.();
    return;
  }

  let msgZip = null;

  const attachments = await new Promise((resolve) => {
    item.getAttachmentsAsync((res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : []);
    });
  });

  msgZip = attachments.find((a) => a.name === "msg.zip");

  if (!msgZip) {
    zmutils.showNotification(
      "msg.zip manquant malgré meta tag. Filtré par antivirus? => Impossible d'afficher le mail."
    );
    console.error("Failed to find msg.zip. Antivirus filtered? Can't decode mail.");
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
    reader = new zip.ZipReader(new zip.BlobReader(new Blob([zipBytes])), {
      password: result.password,
    });
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

// Settings
export function getAutoUnzipMode() {
  return Office.context.roamingSettings.get("zipmail_auto_unzip") || "if-unencrypted";
}

export function isPermanentExtract() {
  return Office.context.roamingSettings.get("zipmail_permanent_extract") !== false;


function openOptionsDialog_simple() {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/options.html",
      { height: 15, width: 15, displayInIframe: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      }
    );
  });
}

// =============================================
// OPEN OPTIONS DIALOG (stable)
// =============================================
function openOptionsDialog(event) {
  // Vérifie si le dialogue est déjà ouvert
  if (Office.context.mailbox[DIALOG_KEY]?.isOpen) {
    zmutils.showNotification("Les options sont déjà ouvertes.");
    event.completed(); // on complète immédiatement pour le menu
    return;
  }

  // Initialise l'objet
  Office.context.mailbox[DIALOG_KEY] = { isOpen: true, dialog: null };

  // Ouvre le dialogue
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/options.html",
    { height: 15, width: 15, displayInIframe: true },
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        zmutils.showNotification("Impossible d'ouvrir la fenêtre d'options.");
        Office.context.mailbox[DIALOG_KEY] = null;
        event.completed(); // complète la commande
        return;
      }

      const dialog = result.value;
      Office.context.mailbox[DIALOG_KEY].dialog = dialog;

      // Appelé lorsque le dialogue est fermé par l'utilisateur
      dialog.addEventHandler(Office.EventType.DialogClosed, () => {
        Office.context.mailbox[DIALOG_KEY] = null;
      });

      // Appelé lorsque le dialogue envoie un message (ex: sauvegarde)
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg.message === "saveDone") {
          // fermeture propre après sauvegarde
          dialog.close();
          Office.context.mailbox[DIALOG_KEY] = null;
        }
      });

      // Complète la commande pour le menu immédiatement
      event.completed();
    }
  );
}

