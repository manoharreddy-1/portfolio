(function () {
  const sidebarLinks = document.querySelectorAll('.cs-sidebar__link');
  const anchors = document.querySelectorAll('.cs-case-anchor');

  if (!sidebarLinks.length || !anchors.length) return;

  function updateSidebar() {
    const scrollPos = window.scrollY + 160;
    let currentId = '';

    anchors.forEach((anchor) => {
      const top = anchor.offsetTop;
      if (scrollPos >= top) {
        currentId = anchor.getAttribute('id');
      }
    });

    if (currentId) {
      sidebarLinks.forEach((link) => {
        const href = link.getAttribute('href');
        link.classList.toggle('active', href === `#${currentId}`);
      });
    }
  }

  window.addEventListener('scroll', updateSidebar, { passive: true });
  updateSidebar();
})();
