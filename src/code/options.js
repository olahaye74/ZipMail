// options.js
Office.onReady(() => {
  const autoUnzip = document.getElementById("autoUnzip");
  const permanentExtract = document.getElementById("permanentExtract");
  const saveBtn = document.getElementById("save");
  const savedMsg = document.getElementById("saved");

  // Chargement
  autoUnzip.value = localStorage.getItem("zipmail_auto_unzip") || "if-unencrypted";
  permanentExtract.checked = localStorage.getItem("zipmail_permanent_extract") !== "false";

  saveBtn.onclick = () => {
    localStorage.setItem("zipmail_auto_unzip", autoUnzip.value);
    localStorage.setItem("zipmail_permanent_extract", permanentExtract.checked);
    savedMsg.style.opacity = 1;
    setTimeout(() => (savedMsg.style.opacity = 0), 2000);
  };
});
