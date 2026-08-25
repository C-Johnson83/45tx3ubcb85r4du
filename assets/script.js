 (function () {
      var links = document.querySelectorAll('#sideNav a');
      var panels = document.querySelectorAll('.panel');
      var main = document.getElementById('mainContent');

      function activate(targetId, updateHash) {
        panels.forEach(function (p) { p.classList.toggle('active', p.id === targetId); });
        links.forEach(function (a) { a.classList.toggle('active', a.dataset.target === targetId); });
        main.scrollTop = 0;
        if (updateHash) { history.replaceState(null, '', '#' + targetId); }
      }

      links.forEach(function (a) {
        a.addEventListener('click', function (e) {
          e.preventDefault();
          activate(a.dataset.target, true);
        });
      });

      var initial = window.location.hash ? window.location.hash.substring(1) : 'panel-overview';
      if (document.getElementById(initial)) { activate(initial, false); }

      /* ============================================================
   THEME SWITCHER
   ============================================================ */

var themeButtons = document.querySelectorAll('.theme-btn');

function setTheme(theme) {

  document.body.classList.remove(
    'theme-blue',
    'theme-dark'
  );

  if (theme === 'blue') {
    document.body.classList.add('theme-blue');
  }

  if (theme === 'dark') {
    document.body.classList.add('theme-dark');
  }

  themeButtons.forEach(function (button) {
    button.classList.toggle(
      'active',
      button.dataset.theme === theme
    );
  });

  localStorage.setItem('t556-theme', theme);
}


/* Restore saved theme */
var savedTheme =
  localStorage.getItem('t556-theme') || 'default';

setTheme(savedTheme);


/* Theme button events */
themeButtons.forEach(function (button) {

  button.addEventListener('click', function () {
    setTheme(button.dataset.theme);
  });

});

/* Tutorial screenshot viewer */
var tutorialViewer = document.getElementById('tutorialViewer');
if (tutorialViewer) {
  var tutorialSlides = Array.from(tutorialViewer.querySelectorAll('.tutorial-slide'));
  var tutorialPrev = document.getElementById('tutorialPrev');
  var tutorialNext = document.getElementById('tutorialNext');
  var tutorialDots = document.getElementById('tutorialDots');
  var tutorialCounter = document.getElementById('tutorialCounter');
  var tutorialLightbox = document.getElementById('tutorialLightbox');
  var tutorialLightboxImage = document.getElementById('tutorialLightboxImage');
  var tutorialLightboxClose = document.getElementById('tutorialLightboxClose');
  var tutorialIndex = 0;

  tutorialSlides.forEach(function (_, index) {
    var dot = document.createElement('button');
    dot.type = 'button';
    dot.className = 'tutorial-dot';
    dot.textContent = index + 1;
    dot.setAttribute('aria-label', 'Show tutorial step ' + (index + 1));
    dot.addEventListener('click', function () { showTutorialSlide(index); });
    tutorialDots.appendChild(dot);
  });

  function showTutorialSlide(index) {
    tutorialIndex = Math.max(0, Math.min(index, tutorialSlides.length - 1));
    tutorialSlides.forEach(function (slide, i) {
      slide.classList.toggle('active', i === tutorialIndex);
    });
    Array.from(tutorialDots.children).forEach(function (dot, i) {
      dot.classList.toggle('active', i === tutorialIndex);
      dot.setAttribute('aria-current', i === tutorialIndex ? 'step' : 'false');
    });
    tutorialPrev.disabled = tutorialIndex === 0;
    tutorialNext.disabled = tutorialIndex === tutorialSlides.length - 1;
    tutorialCounter.textContent = 'Step ' + (tutorialIndex + 1) + ' of ' + tutorialSlides.length;
  }

  tutorialPrev.addEventListener('click', function () { showTutorialSlide(tutorialIndex - 1); });
  tutorialNext.addEventListener('click', function () { showTutorialSlide(tutorialIndex + 1); });

  tutorialViewer.addEventListener('click', function (event) {
    var img = event.target.closest('.tutorial-image-frame img');
    if (!img) return;
    tutorialLightboxImage.src = img.src;
    tutorialLightboxImage.alt = img.alt || 'Tutorial screenshot';
    tutorialLightbox.classList.add('open');
    tutorialLightbox.setAttribute('aria-hidden', 'false');
  });

  function closeTutorialLightbox() {
    tutorialLightbox.classList.remove('open');
    tutorialLightbox.setAttribute('aria-hidden', 'true');
    tutorialLightboxImage.src = '';
  }

  tutorialLightboxClose.addEventListener('click', closeTutorialLightbox);
  tutorialLightbox.addEventListener('click', function (event) {
    if (event.target === tutorialLightbox) closeTutorialLightbox();
  });
  document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape' && tutorialLightbox.classList.contains('open')) {
      closeTutorialLightbox();
    }
  });

  showTutorialSlide(0);
}

    })();