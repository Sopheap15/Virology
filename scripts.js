document.addEventListener('DOMContentLoaded', function() {
  // Create toggle button immediately
  const toggleBtn = document.createElement('button');
  toggleBtn.id = 'night-mode-toggle';
  toggleBtn.innerHTML = '🌙';
  toggleBtn.title = 'Toggle Night Mode';
  toggleBtn.style.cssText = `
    position: fixed;
    top: 10px;
    right: 10px;
    width: 30px;
    height: 30px;
    border-radius: 50%;
    background: #007bff;
    color: white;
    border: 1px solid white;
    font-size: 15px;
    cursor: pointer;
    z-index: 10000;
    transition: all 0.3s ease;
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 3px 6px rgba(0,0,0,0.3);
  `;

  // Add hover effect
  toggleBtn.onmouseover = function() {
    this.style.background = '#0056b3';
    this.style.transform = 'scale(1.1)';
  };
  toggleBtn.onmouseout = function() {
    this.style.background = '#007bff';
    this.style.transform = 'scale(1)';
  };

  // Add click handler
  toggleBtn.onclick = function() {
    document.body.classList.toggle('night-mode');
    
    // Update button icon based on current mode
    setTimeout(function() {
      if (document.body.classList.contains('night-mode')) {
        toggleBtn.innerHTML = '☀️';
        toggleBtn.title = 'Switch to Light Mode';
      } else {
        toggleBtn.innerHTML = '🌙';
        toggleBtn.title = 'Switch to Night Mode';
      }
    }, 100);
  };

  // Add to page
  document.body.appendChild(toggleBtn);
  console.log('Button added to body');
  
});

