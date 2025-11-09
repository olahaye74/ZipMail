// src/code/read.js
Office.initialize = () => {};

Office.actions.associate("onItemChanged", onItemChanged);
Office.actions.associate("onItemRead", onItemRead);

async function onItemChanged(event) {
  await onItemRead(event);
  event.completed();
}

async function onItemRead(event) {
  // TON CODE DE DÉCOMPRESSION ICI
  const body = await Office.context.mailbox.item.body.getAsync("html");
  const meta = parseZipMailMeta(body.value);
  if (meta && shouldAutoUnzip(meta)) {
    await unzipAndDisplay(body.value, meta);
  }
  event.completed();
}

function shouldAutoUnzip(meta) {
  const setting = localStorage.getItem("zipmail_auto_unzip") || "if-unencrypted";
  if (setting === "never") return false;
  if (setting === "always") return true;
  return !meta.encrypted;
}

// parseZipMailMeta, unzipAndDisplay → à déplacer ici depuis compose.js
