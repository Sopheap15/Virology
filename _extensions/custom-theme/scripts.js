// Simple night mode toggle
console.log('Custom theme extension loaded');
(function () {
  'use strict';
  console.log('IIFE executing');

  function createToggle() {
    console.log('createToggle called');
    if (document.getElementById('night-mode-toggle')) {
      console.log('Button already exists');
      return;
    }

    console.log('Creating toggle button');
    const btn = document.createElement('button');
    btn.id = 'night-mode-toggle';
    btn.innerHTML = '🌙';
    btn.title = 'Toggle Night Mode';
    btn.style.cssText = 'position:fixed;top:10px;right:10px;width:50px;height:50px;border-radius:50%;background:#007bff;color:white;border:3px solid white;font-size:20px;cursor:pointer;z-index:100000;';

    btn.onclick = function () {
      console.log('Button clicked');
      document.body.classList.toggle('night-mode');
      btn.innerHTML = document.body.classList.contains('night-mode') ? '☀️' : '🌙';
    };

    document.body.appendChild(btn);
    console.log('Button added to DOM');
  }

  if (document.readyState === 'loading') {
    console.log('Document loading, adding DOMContentLoaded listener');
    document.addEventListener('DOMContentLoaded', function () {
      console.log('DOMContentLoaded fired');
      createToggle();
    });
  } else {
    console.log('Document ready, creating toggle immediately');
    createToggle();
  }
})();