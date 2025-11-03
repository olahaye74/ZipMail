/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(() => {
  const levelSelect = document.getElementById("compressionLevel");
  const pwdInput = document.getElementById("zipPassword");
  const toggleBtn = document.getElementById("togglePassword");
  const previewBtn = document.getElementById("previewZip");
  const previewDiv = document.getElementById("preview");
  const status = document.getElementById("status");

  // Charge les valeurs
  levelSelect.value = localStorage.getItem("zipLevel") || "6";
  pwdInput.value = localStorage.getItem("zipPassword") || "";

  // Sauvegarde en temps réel
  levelSelect.onchange = () => {
    localStorage.setItem("zipLevel", levelSelect.value);
    Office.context.ui.messageParent(`update:level:${levelSelect.value}`);
  };
  pwdInput.onchange = () => {
    localStorage.setItem("zipPassword", pwdInput.value);
    Office.context.ui.messageParent(`update:password:${pwdInput.value}`);
  };

  // Toggle visibilité
  toggleBtn.onclick = () => {
    const isText = pwdInput.type === "text";
    pwdInput.type = isText ? "password" : "text";
    toggleBtn.textContent = isText ? "Voir" : "Cacher";
  };

  // Prévisualisation (simulée)
  previewBtn.onclick = async () => {
    status.textContent = "Simulation...";
    previewDiv.innerHTML = `
      <strong>Prévisualisation</strong><br>
      Niveau: <strong>${levelSelect.value}</strong><br>
      Chiffrement: <strong>${pwdInput.value ? "AES-256" : "désactivé"}</strong>
    `;
    setTimeout(() => (status.textContent = "Prêt"), 1000);
  };
});
