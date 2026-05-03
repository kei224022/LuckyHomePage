/* ============================================================
   株式会社ラッキー — main.js
   ============================================================ */

(function () {
  'use strict';

  /* --- Sticky header shadow --- */
  const header = document.querySelector('.site-header');
  if (header) {
    const onScroll = () => {
      header.classList.toggle('scrolled', window.scrollY > 10);
    };
    window.addEventListener('scroll', onScroll, { passive: true });
    onScroll();
  }

  /* --- Mobile nav toggle --- */
  const toggle = document.querySelector('.nav-toggle');
  const nav = document.querySelector('.site-nav');
  if (toggle && nav) {
    toggle.addEventListener('click', () => {
      const open = toggle.getAttribute('aria-expanded') === 'true';
      toggle.setAttribute('aria-expanded', String(!open));
      nav.classList.toggle('open', !open);
      document.body.style.overflow = open ? '' : 'hidden';
    });

    /* Close on nav link click (mobile) */
    nav.querySelectorAll('a').forEach(a => {
      a.addEventListener('click', () => {
        toggle.setAttribute('aria-expanded', 'false');
        nav.classList.remove('open');
        document.body.style.overflow = '';
      });
    });

    /* Close on Escape */
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape' && nav.classList.contains('open')) {
        toggle.setAttribute('aria-expanded', 'false');
        nav.classList.remove('open');
        document.body.style.overflow = '';
        toggle.focus();
      }
    });
  }

  /* --- Active nav link --- */
  const page = document.body.dataset.page;
  if (page) {
    document.querySelectorAll('.site-nav a[data-nav]').forEach(a => {
      if (a.dataset.nav === page) {
        a.setAttribute('aria-current', 'page');
      }
    });
  }

  /* --- Scroll reveal --- */
  const revealEls = document.querySelectorAll('.reveal');
  if (revealEls.length) {
    const observer = new IntersectionObserver(
      entries => {
        entries.forEach(entry => {
          if (entry.isIntersecting) {
            entry.target.classList.add('visible');
            observer.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.12, rootMargin: '0px 0px -40px 0px' }
    );
    revealEls.forEach(el => observer.observe(el));
  }

  /* --- Counter animation (trust strip) --- */
  function animateCounter(el, target, suffix) {
    const duration = 1200;
    const start = performance.now();
    const isFloat = target % 1 !== 0;
    const step = ts => {
      const elapsed = ts - start;
      const progress = Math.min(elapsed / duration, 1);
      const eased = 1 - Math.pow(1 - progress, 3);
      const current = target * eased;
      el.textContent = (isFloat ? current.toFixed(1) : Math.round(current)).toLocaleString('ja-JP') + suffix;
      if (progress < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }

  const trustStrip = document.querySelector('.trust-strip');
  if (trustStrip) {
    const counters = trustStrip.querySelectorAll('[data-count]');
    const obs = new IntersectionObserver(entries => {
      if (entries[0].isIntersecting) {
        counters.forEach(el => {
          const target = parseFloat(el.dataset.count);
          const suffix = el.dataset.suffix || '';
          animateCounter(el, target, suffix);
        });
        obs.disconnect();
      }
    }, { threshold: 0.5 });
    obs.observe(trustStrip);
  }

})();
