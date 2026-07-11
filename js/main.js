/* ============================================
   Studio Dotbox - Main JavaScript
   ============================================ */

(function () {
  'use strict';

  var header = document.querySelector('.site-header');
  if (header) {
    window.addEventListener('scroll', function () {
      if (window.scrollY > 20) header.classList.add('scrolled');
      else header.classList.remove('scrolled');
    });
  }

  var menuToggle = document.querySelector('.site-header__menu-toggle');
  var mobileMenu = document.querySelector('.mobile-menu');
  var mobileMenuOverlay = document.querySelector('.mobile-menu__overlay');
  var mobileMenuClose = document.querySelector('.mobile-menu__close');

  function openMobileMenu() {
    if (!mobileMenu) return;
    mobileMenu.classList.add('open');
    if (mobileMenuOverlay) mobileMenuOverlay.classList.add('visible');
    document.body.style.overflow = 'hidden';
  }
  function closeMobileMenu() {
    if (!mobileMenu) return;
    mobileMenu.classList.remove('open');
    if (mobileMenuOverlay) mobileMenuOverlay.classList.remove('visible');
    document.body.style.overflow = '';
  }
  if (menuToggle) menuToggle.addEventListener('click', openMobileMenu);
  if (mobileMenuClose) mobileMenuClose.addEventListener('click', closeMobileMenu);
  if (mobileMenuOverlay) mobileMenuOverlay.addEventListener('click', closeMobileMenu);
  if (mobileMenu) {
    mobileMenu.querySelectorAll('a').forEach(function (link) {
      link.addEventListener('click', closeMobileMenu);
    });
  }
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') closeMobileMenu();
  });

  var revealElements = document.querySelectorAll('.reveal');
  if (revealElements.length && 'IntersectionObserver' in window) {
    var revealObserver = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          entry.target.classList.add('visible');
          revealObserver.unobserve(entry.target);
        }
      });
    }, { threshold: 0.15, rootMargin: '0px 0px -50px 0px' });
    revealElements.forEach(function (el) { revealObserver.observe(el); });
  } else {
    revealElements.forEach(function (el) { el.classList.add('visible'); });
  }

  var currentPath = window.location.pathname.replace(/\/$/, '') || '/';
  document.querySelectorAll('.site-header__nav a, .mobile-menu__primary a').forEach(function (link) {
    var href = (link.getAttribute('href') || '').replace(/\/$/, '') || '/';
    var fileName = currentPath.split('/').pop() || 'index.html';
    var linkFileName = href.split('/').pop() || 'index.html';
    if (fileName === linkFileName ||
        (currentPath === '/' && (href === '/' || href === '/index.html' || href === 'index.html'))) {
      link.classList.add('active');
    }
  });

  // Work archive filtering
  var filterButtons = document.querySelectorAll('.filters__btn');
  var workCards = document.querySelectorAll('.work-card');
  if (filterButtons.length && workCards.length) {
    filterButtons.forEach(function (btn) {
      btn.addEventListener('click', function () {
        var group = btn.dataset.group;
        var value = btn.dataset.value;
        document.querySelectorAll('.filters__btn[data-group="' + group + '"]').forEach(function (b) {
          b.classList.remove('active');
        });
        btn.classList.add('active');
        var sectorFilter = document.querySelector('.filters__btn[data-group="sector"].active');
        var stageFilter = document.querySelector('.filters__btn[data-group="stage"].active');
        var sectorVal = sectorFilter ? sectorFilter.dataset.value : 'all';
        var stageVal = stageFilter ? stageFilter.dataset.value : 'all';
        workCards.forEach(function (card) {
          var matchSector = sectorVal === 'all' || card.dataset.sector === sectorVal;
          var matchStage = stageVal === 'all' || card.dataset.stage === stageVal;
          card.style.display = (matchSector && matchStage) ? '' : 'none';
        });
      });
    });
  }
})();
