Office.onReady(() => {
  const select = document.getElementById('autoUnzip');
  const btn = document.getElementById('save');
  const msg = document.getElementById('saved');

  select.value = localStorage.getItem('zipmail_auto_unzip') || 'if-unencrypted';

  btn.onclick = () => {
    localStorage.setItem('zipmail_auto_unzip', select.value);
    msg.style.opacity = 1;
    setTimeout(() => msg.style.opacity = 0, 2000);
  };
});
