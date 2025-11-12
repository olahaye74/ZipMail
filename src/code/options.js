// options.js
Office.onReady(() => {
  const autoUnzip = document.getElementById("autoUnzip");
  const permanentExtract = document.getElementById("permanentExtract");
  const saveBtn = document.getElementById("save");
  const savedMsg = document.getElementById("saved");

  const settings = Office.context.roamingSettings;

  // Chargement
  autoUnzip.value = settings.get("zipmail_auto_unzip") || "if-unencrypted";
  permanentExtract.checked = settings.get("zipmail_permanent_extract") !== false;

  saveBtn.onclick = () => {
    settings.set("zipmail_auto_unzip", autoUnzip.value);
    settings.set("zipmail_permanent_extract", permanentExtract.checked);
    settings.saveAsync(() => {
      savedMsg.style.opacity = 1;
      setTimeout(() => (savedMsg.style.opacity = 0), 2000);
    });
    // Notifie le runtime parent que la sauvegarde est finie
    Office.context.ui.messageParent("saveDone");
  };
});
