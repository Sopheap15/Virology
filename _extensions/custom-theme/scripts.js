// Simple night mode toggle
(function () {
  'use strict';
  function createToggle() {
    console.log('createToggle called');
    if (document.getElementById('night-mode-toggle')) {
      return;
    }

    const btn = document.createElement('button');
    btn.id = 'night-mode-toggle';
    btn.innerHTML = '🌙';
    btn.title = 'Toggle Night Mode';
    btn.style.cssText = 'position:fixed;top:10px;right:10px;width:50px;height:50px;border-radius:50%;background:#007bff;color:white;border:3px solid white;font-size:20px;cursor:pointer;z-index:100000;';

    btn.onclick = function () {
      document.body.classList.toggle('night-mode');
      btn.innerHTML = document.body.classList.contains('night-mode') ? '☀️' : '🌙';
    };

    document.body.appendChild(btn);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () {
      createToggle();
    });
  } else {
    createToggle();
  }
})();