// X-ERP Dokumentation - App JavaScript

document.addEventListener('DOMContentLoaded', () => {
  // ===== MOBILE MENU =====
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('overlay');
  const mobileMenuBtn = document.getElementById('mobileMenuBtn');
  
  if (mobileMenuBtn && sidebar && overlay) {
    mobileMenuBtn.addEventListener('click', () => {
      sidebar.classList.toggle('open');
      overlay.classList.toggle('open');
    });
    
    overlay.addEventListener('click', () => {
      sidebar.classList.remove('open');
      overlay.classList.remove('open');
    });
  }
  
  // ===== Strg+K für Suche =====
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key === 'k') {
      e.preventDefault();
      const searchInput = document.getElementById('toc-search');
      if (searchInput) {
        searchInput.focus();
      }
    }
  });
  
  // ===== SIDEBAR LINKS - Mobile schließen =====
  const tocLinks = document.querySelectorAll('.toc-tree a');
  tocLinks.forEach(link => {
    link.addEventListener('click', () => {
      if (window.innerWidth <= 900 && sidebar) {
        sidebar.classList.remove('open');
        if (overlay) overlay.classList.remove('open');
      }
    });
  });

  // ===== BILD-LIGHTBOX =====
  const lbOverlay = document.createElement('div');
  lbOverlay.className = 'lightbox-overlay';
  document.body.appendChild(lbOverlay);

  const lbImg = document.createElement('img');
  lbOverlay.appendChild(lbImg);

  document.querySelectorAll('.doc-content img').forEach(img => {
    img.addEventListener('click', () => {
      lbImg.src = img.src;
      lbImg.alt = img.alt;
      lbOverlay.classList.add('active');
    });
  });

  lbOverlay.addEventListener('click', () => {
    lbOverlay.classList.remove('active');
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') lbOverlay.classList.remove('active');
  });

  const pageName = (window.location.pathname.split('/').pop() || '').toLowerCase();
  if (pageName.endsWith('list.html')) {
    setTimeout(() => {
      document.querySelectorAll('.help-image-annotated').forEach((img) => img.remove());
      document.querySelectorAll('.help-view-hero').forEach((panel) => panel.remove());
      document.querySelectorAll('.help-view-panel').forEach((panel) => panel.remove());
      document.querySelectorAll('.help-annotation-callout').forEach((panel) => panel.remove());
      document.querySelectorAll('.help-field-props').forEach((panel) => panel.remove());
      document.querySelectorAll('.help-view-meta').forEach((panel) => panel.remove());
      document.querySelectorAll('.help-view-section-head').forEach((panel) => panel.remove());
      document.querySelectorAll('.help-view-shot-panel').forEach((panel) => panel.remove());
    }, 0);
  }
});
