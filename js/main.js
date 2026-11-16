(function () {
  const navbar = document.getElementById('navbar');
  const navToggle = document.getElementById('navToggle');
  const navLinks = document.getElementById('navLinks');
  const themeToggle = document.getElementById('themeToggle');
  const sections = document.querySelectorAll('section[id], header[id]');
  const skillCards = document.querySelectorAll('.skill-card');
  const modal = document.getElementById('certModal');
  const modalClose = document.getElementById('modalClose');
  const modalBackdrop = document.getElementById('modalBackdrop');
  const contactForm = document.getElementById('contactForm');
  const formStatus = document.getElementById('formStatus');
  const submitBtn = document.getElementById('submitBtn');
  const RECIPIENT_EMAIL = 'bommareddymanoharreddy999@gmail.com';
  const FORM_ENDPOINT = 'https://formsubmit.co/ajax/bommareddymanoharreddy999@gmail.com';
  const GMAIL_COMPOSE = `https://mail.google.com/mail/?view=cm&fs=1&to=${RECIPIENT_EMAIL}&su=Portfolio%20Inquiry`;

  // Hero email — open Gmail compose (works without a desktop mail app)
  const heroEmail = document.getElementById('heroEmail');
  if (heroEmail) {
    heroEmail.addEventListener('click', (e) => {
      e.preventDefault();
      window.open(GMAIL_COMPOSE, '_blank', 'noopener,noreferrer');
    });
  }

  // Theme toggle
  function getTheme() {
    return document.documentElement.getAttribute('data-theme') || 'dark';
  }

  function setTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
  }

  themeToggle.addEventListener('click', () => {
    setTheme(getTheme() === 'dark' ? 'light' : 'dark');
  });

  // Navbar scroll effect
  window.addEventListener('scroll', () => {
    navbar.classList.toggle('scrolled', window.scrollY > 50);
    updateActiveNav();
  });

  // Mobile menu toggle
  navToggle.addEventListener('click', () => {
    navLinks.classList.toggle('open');
    navToggle.classList.toggle('active');
  });

  navLinks.querySelectorAll('a').forEach((link) => {
    link.addEventListener('click', () => {
      navLinks.classList.remove('open');
      navToggle.classList.remove('active');
    });
  });

  function updateActiveNav() {
    const scrollPos = window.scrollY + 100;

    sections.forEach((section) => {
      const top = section.offsetTop;
      const height = section.offsetHeight;
      const id = section.getAttribute('id');

      if (scrollPos >= top && scrollPos < top + height) {
        navLinks.querySelectorAll('a').forEach((link) => {
          link.classList.toggle('active', link.getAttribute('href') === `#${id}`);
        });
      }
    });
  }

  const skillObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          const level = entry.target.dataset.level;
          entry.target.style.setProperty('--level', `${level}%`);
          entry.target.classList.add('visible');
          skillObserver.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.3 }
  );

  skillCards.forEach((card) => skillObserver.observe(card));

  const fadeObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry, i) => {
        if (entry.isIntersecting) {
          setTimeout(() => entry.target.classList.add('visible'), i * 80);
          fadeObserver.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.1 }
  );

  document.querySelectorAll('.about__card, .timeline__item, .cert-card, .project-card, .case-study-card').forEach((el) => {
    el.classList.add('fade-up');
    fadeObserver.observe(el);
  });

  const modalImage = document.getElementById('modalImage');

  document.querySelectorAll('[data-modal]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const imgSrc = btn.dataset.img;
      if (imgSrc && modalImage) modalImage.src = imgSrc;
      modal.classList.add('open');
      modal.setAttribute('aria-hidden', 'false');
      document.body.style.overflow = 'hidden';
    });
  });

  function closeModal() {
    modal.classList.remove('open');
    modal.setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';
  }

  modalClose.addEventListener('click', closeModal);
  modalBackdrop.addEventListener('click', closeModal);

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && modal.classList.contains('open')) closeModal();
  });

  function showFormStatus(message, type) {
    formStatus.textContent = message;
    formStatus.className = `form-status ${type}`;
  }

  function openGmailCompose(name, email, subject, message) {
    const params = new URLSearchParams({
      view: 'cm',
      fs: '1',
      to: RECIPIENT_EMAIL,
      su: `Portfolio: ${subject}`,
      body: `Name: ${name}\nEmail: ${email}\n\n${message}`,
    });
    window.open(`https://mail.google.com/mail/?${params.toString()}`, '_blank', 'noopener,noreferrer');
  }

  contactForm.addEventListener('submit', async (e) => {
    e.preventDefault();

    const name = document.getElementById('name').value.trim();
    const email = document.getElementById('email').value.trim();
    const subject = document.getElementById('subject').value.trim();
    const message = document.getElementById('message').value.trim();

    if (!name || !email || !subject || !message) {
      showFormStatus('Please fill in all fields.', 'error');
      return;
    }

    submitBtn.disabled = true;
    submitBtn.textContent = 'Sending...';
    showFormStatus('', '');

    const formData = new FormData(contactForm);
    formData.set('_subject', `Portfolio Contact: ${subject}`);

    try {
      const response = await fetch(FORM_ENDPOINT, {
        method: 'POST',
        body: formData,
        headers: { Accept: 'application/json' },
      });

      const result = await response.json().catch(() => ({}));
      const sent = response.ok && (result.success === 'true' || result.success === true);

      if (sent) {
        showFormStatus('Message sent successfully! I will get back to you soon.', 'success');
        contactForm.reset();
      } else {
        throw new Error(result.message || 'Unable to send message.');
      }
    } catch (error) {
      showFormStatus('Could not send automatically. Opening Gmail for you...', 'error');
      setTimeout(() => openGmailCompose(name, email, subject, message), 500);
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = 'Send Message';
    }
  });

  updateActiveNav();
})();
