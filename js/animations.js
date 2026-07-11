/* ============================================
   Studio Dotbox - Animations
   ============================================ */

(function () {
  'use strict';

  var prefersReducedMotion = window.matchMedia &&
    window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  function drawPath(el, duration) {
    if (!el) return;
    var length;
    try { length = el.getTotalLength(); } catch (e) { length = 300; }
    el.style.strokeDasharray = length;
    el.style.strokeDashoffset = length;
    el.style.opacity = '1';
    el.style.transition = 'stroke-dashoffset ' + duration + 'ms ease-out';
    el.getBoundingClientRect();
    el.style.strokeDashoffset = '0';
  }

  function runSignatureIntro() {
    var overlay = document.getElementById('intro-overlay');
    if (!overlay) return;

    if (prefersReducedMotion) {
      overlay.remove();
      document.body.classList.remove('intro-active');
      return;
    }

    try {
      if (sessionStorage.getItem('studiodotbox_intro_seen')) {
        overlay.remove();
        document.body.classList.remove('intro-active');
        return;
      }
    } catch (e) {}

    document.body.classList.add('intro-active');

    var dot = overlay.querySelector('#sig-dot');
    var line = overlay.querySelector('#sig-line');
    var square = overlay.querySelector('#sig-square');
    var depthTL = overlay.querySelector('#sig-depth-tl');
    var depthTR = overlay.querySelector('#sig-depth-tr');
    var depthBR = overlay.querySelector('#sig-depth-br');
    var back = overlay.querySelector('#sig-back');
    var cubeShapes = overlay.querySelectorAll('#sig-cube-group *');
    var wordmark = overlay.querySelector('#sig-wordmark');
    var skipBtn = overlay.querySelector('.intro-overlay__skip');

    function fadeOutAndRemove() {
      overlay.classList.add('fade-out');
      setTimeout(function () {
        if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
        document.body.classList.remove('intro-active');
      }, 600);
      try { sessionStorage.setItem('studiodotbox_intro_seen', '1'); } catch (e) {}
    }
    if (skipBtn) skipBtn.addEventListener('click', fadeOutAndRemove);

    setTimeout(function () { drawPath(line, 800); }, 600);
    setTimeout(function () {
      if (dot) { dot.style.transition = 'opacity 400ms'; dot.style.opacity = '0'; }
    }, 1500);
    setTimeout(function () { drawPath(square, 1400); }, 1900);
    setTimeout(function () {
      drawPath(depthTL, 700);
      drawPath(depthTR, 700);
      drawPath(depthBR, 700);
    }, 3400);
    setTimeout(function () { drawPath(back, 1200); }, 4200);
    setTimeout(function () {
      cubeShapes.forEach(function (s) { s.style.transition = 'opacity 600ms'; s.style.opacity = '0'; });
    }, 5500);
    setTimeout(function () {
      if (wordmark) {
        wordmark.style.transition = 'opacity 700ms ease-out, transform 700ms ease-out';
        wordmark.style.opacity = '1';
        wordmark.style.transform = 'translateY(0)';
      }
    }, 5900);
    setTimeout(fadeOutAndRemove, 7500);
  }

  function runPhilosophyOpening() {
    var container = document.getElementById('philosophy-opening');
    if (!container) return;

    var dot = container.querySelector('#phil-dot');
    var line = container.querySelector('#phil-line');
    var square = container.querySelector('#phil-square');
    var depthTL = container.querySelector('#phil-depth-tl');
    var depthTR = container.querySelector('#phil-depth-tr');
    var depthBR = container.querySelector('#phil-depth-br');
    var back = container.querySelector('#phil-back');
    var cubeShapes = container.querySelectorAll('#phil-cube-group *');
    var text = container.querySelector('#philosophy-text');

    if (prefersReducedMotion) {
      if (text) text.style.opacity = '1';
      cubeShapes.forEach(function (s) { s.style.opacity = '1'; });
      return;
    }

    var hasStarted = false;
    function start() {
      if (hasStarted) return;
      hasStarted = true;

      setTimeout(function () { if (dot) { dot.style.transition = 'opacity 700ms ease-out'; dot.style.opacity = '1'; } }, 200);
      setTimeout(function () { drawPath(line, 900); }, 1100);
      setTimeout(function () { if (dot) { dot.style.transition = 'opacity 600ms'; dot.style.opacity = '0'; } }, 2200);
      setTimeout(function () { drawPath(square, 1500); }, 2600);
      setTimeout(function () {
        drawPath(depthTL, 800); drawPath(depthTR, 800); drawPath(depthBR, 800);
      }, 4300);
      setTimeout(function () { drawPath(back, 1300); }, 5200);
      setTimeout(function () {
        if (text) {
          text.style.transition = 'opacity 800ms ease-out, transform 800ms ease-out';
          text.style.opacity = '1';
          text.style.transform = 'translateY(0)';
        }
      }, 6700);
    }

    if ('IntersectionObserver' in window) {
      var observer = new IntersectionObserver(function (entries) {
        entries.forEach(function (entry) {
          if (entry.isIntersecting) { start(); observer.unobserve(entry.target); }
        });
      }, { threshold: 0.3 });
      observer.observe(container);
    } else { start(); }
  }

  function setupStageMarks() {
    var marks = document.querySelectorAll('.stage__mark');
    if (!marks.length) return;
    if (prefersReducedMotion || !('IntersectionObserver' in window)) {
      marks.forEach(function (m) {
        m.classList.add('animate-in');
        m.querySelectorAll('svg *').forEach(function (c) { c.style.opacity = '1'; });
      });
      return;
    }
    var observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          entry.target.classList.add('animate-in');
          observer.unobserve(entry.target);
        }
      });
    }, { threshold: 0.4 });
    marks.forEach(function (m) { observer.observe(m); });
  }

  function run404Animation() {
    var line = document.getElementById('failed-line');
    if (!line) return;
    if (prefersReducedMotion) { line.setAttribute('x2', '90'); return; }
    setTimeout(function () {
      var startX = 60, endX = 88, current = startX;
      var anim = setInterval(function () {
        current += 1.5;
        if (current >= endX) {
          current = endX; clearInterval(anim);
          var flickerCount = 0;
          var flicker = setInterval(function () {
            line.style.opacity = (flickerCount % 2 === 0) ? '0.3' : '1';
            flickerCount++;
            if (flickerCount >= 5) { clearInterval(flicker); line.style.opacity = '1'; }
          }, 90);
        }
        line.setAttribute('x2', current);
      }, 16);
    }, 300);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () {
      runSignatureIntro(); runPhilosophyOpening(); setupStageMarks(); run404Animation();
    });
  } else {
    runSignatureIntro(); runPhilosophyOpening(); setupStageMarks(); run404Animation();
  }
})();
