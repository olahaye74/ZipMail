// =============================================
// ZipMail HELPERS
// =============================================
export function showNotification(msg) {
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

// =============================================
// 3️⃣ PASSWORD DIALOG
// =============================================
export async function getPasswordFromDialog() {
  return new Promise((resolve) => {
    try {
      Office.context.ui.displayDialogAsync(
        `${window.location.origin}/password.html`,
        { height: 12, width: 12, displayInIframe: true },
        (asyncResult) => {
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            resolve(null);
            return;
          }
          const dialog = asyncResult.value;
          const handler = (arg) => {
            let password = null;
            try {
              let data = JSON.parse(arg.message);
              password = (data.password || "").trim();
            } catch {
              password = (arg.message || "").trim();
            }
            dialog.close();
            resolve(password || null);
          };
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, handler);
          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => resolve(null));
        }
      );
    } catch {
      resolve(null);
    }
  });
}

// =============================================
// Get meta information from html message
// =============================================
export function parseZipMailMeta(html) {
  try {
    const metaMatch = html.match(/<meta\s+name=["']zipmail["']\s+content=["']([^"']+)["']/i);
    if (!metaMatch) return null;
    const metaContent = metaMatch[1];
    const parts = metaContent.split(";").map((p) => p.trim());
    const result = {};
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

export async function getAttachmentContent(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(id, (res) =>
      res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value) : reject(res.error)
    );
  });
}

export async function removeAttachment(id) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.removeAttachmentAsync(id, (res) =>
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error)
    );
  });
}

export async function addAttachmentFromBase64(name, base64) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(base64, name, (res) =>
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error)
    );
  });
}

// -----------------------------
// base64 / array helpers
// -----------------------------
export function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (let i = 0; i < bytes.byteLength; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

export function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

export function base64ToUint8Array(base64) {
  const raw = atob(base64);
  const array = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) array[i] = raw.charCodeAt(i);
  return array;
}

export function blobToBase64(blob) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(",")[1]);
    reader.readAsDataURL(blob);
  });
}

// Not used, just in case
export function showDialogAlert(message) {
  const url = `https://localhost:3000/dialog-alert.html?msg=${encodeURIComponent(message)}`;

  Office.context.ui.displayDialogAsync(url, { height: 30, width: 40 }, (result) => {
    if (result.status === "succeeded") {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => dialog.close());
    }
  });
}
