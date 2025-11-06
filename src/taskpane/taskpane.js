/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(() => {
  const levelSelect = document.getElementById("compressionLevel");
  const previewBtn = document.getElementById("previewZip");
  const previewDiv = document.getElementById("preview");
  const status = document.getElementById("status");

  // Charge les valeurs
  levelSelect.value = localStorage.getItem("zipLevel") || "6";
  levelSelect.value = localStorage.getItem("zipLevel") || "6";

  // Sauvegarde en temps réel
  levelSelect.onchange = () => {
    localStorage.setItem("zipLevel", levelSelect.value);
  };

  // Prévisualisation (simulée)
  previewBtn.onclick = async () => {
    status.textContent = "Thinking...";
    previewDiv.innerHTML = `
      <strong>Etat actuel</strong><br>
      Niveau: <strong>${levelSelect.value}</strong><br>
      Mode: <strong>${localStorage.getItem("zipMode")}</strong>
    `;
    setTimeout(() => (status.textContent = "Prêt"), 1000);
  };
});
