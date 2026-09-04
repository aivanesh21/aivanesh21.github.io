/* Sawant & Kadrekar LLP — site behaviour.
   Progressive enhancement only: every page works with this file absent. */
(function () {
  'use strict';

  var reduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var $  = function (s, c) { return (c || document).querySelector(s); };
  var $$ = function (s, c) { return Array.prototype.slice.call((c || document).querySelectorAll(s)); };

  /* --- Sticky header shadow -------------------------------------------- */
  var hdr = $('[data-header]');
  if (hdr) {
    var onScroll = function () { hdr.dataset.stuck = window.scrollY > 8 ? 'true' : 'false'; };
    onScroll();
    window.addEventListener('scroll', onScroll, { passive: true });

    // The wide mega panel is fixed to the viewport and sits directly under the header,
    // so it needs the header's real height rather than a guess.
    var setHdrH = function () {
      document.documentElement.style.setProperty('--hdr-h', hdr.offsetHeight + 'px');
    };
    setHdrH();
    window.addEventListener('resize', setHdrH, { passive: true });
    if (window.ResizeObserver) new ResizeObserver(setHdrH).observe(hdr);
  }

  /* --- Mobile navigation ------------------------------------------------ */
  var mnav = $('[data-mnav]');
  var burger = $('[data-burger]');
  var burgerClose = $('[data-burger-close]');

  function setNav(open) {
    if (!mnav || !burger) return;
    mnav.dataset.open = open ? 'true' : 'false';
    burger.setAttribute('aria-expanded', open ? 'true' : 'false');
    burger.setAttribute('aria-label', open ? 'Close menu' : 'Open menu');
    document.body.style.overflow = open ? 'hidden' : '';
    if (open) {
      mnav.hidden = false;
      var first = mnav.querySelector('a, button');
      if (first) first.focus();
    } else {
      // keep it out of the tab order once the transition has finished
      window.setTimeout(function () { if (mnav.dataset.open === 'false') mnav.hidden = true; }, 340);
      burger.focus();
    }
  }

  if (burger) burger.addEventListener('click', function () { setNav(mnav.dataset.open !== 'true'); });
  if (burgerClose) burgerClose.addEventListener('click', function () { setNav(false); });
  if (mnav) {
    mnav.hidden = true;
    $$('a', mnav).forEach(function (a) { a.addEventListener('click', function () { setNav(false); }); });
  }
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && mnav && mnav.dataset.open === 'true') setNav(false);
  });

  // Trap focus inside the open mobile menu
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Tab' || !mnav || mnav.dataset.open !== 'true') return;
    // Links inside a collapsed sub-section still have an offsetParent, so measure instead:
    // a collapsed grid row gives them zero height.
    var f = $$('a[href], button:not([disabled])', mnav).filter(function (el) {
      return el.getBoundingClientRect().height > 0;
    });
    if (!f.length) return;
    var first = f[0], last = f[f.length - 1];
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
  });

  /* --- Desktop dropdowns -------------------------------------------------
     Opens on hover and on click, closes on Escape, outside click or blur.
     The button is a real <button aria-expanded>, so it works from the keyboard
     with no pointer at all. */
  var openPanel = null, closeTimer = null;

  function setPanel(item, open) {
    var btn = $('.navitem__btn', item);
    var panel = $('[data-navpanel]', item);
    if (!btn || !panel) return;
    if (open) {
      if (openPanel && openPanel !== item) setPanel(openPanel, false);
      panel.hidden = false;
      // let the browser register the un-hide before transitioning
      requestAnimationFrame(function () { panel.dataset.open = 'true'; });
      btn.setAttribute('aria-expanded', 'true');
      openPanel = item;
    } else {
      panel.dataset.open = 'false';
      btn.setAttribute('aria-expanded', 'false');
      window.setTimeout(function () {
        if (panel.dataset.open === 'false') panel.hidden = true;
      }, reduced ? 0 : 210);
      if (openPanel === item) openPanel = null;
    }
  }

  $$('[data-navitem]').forEach(function (item) {
    var btn = $('.navitem__btn', item);
    if (!btn) return;

    btn.addEventListener('click', function (e) {
      e.preventDefault();
      setPanel(item, btn.getAttribute('aria-expanded') !== 'true');
    });

    item.addEventListener('mouseenter', function () {
      if (window.matchMedia('(hover: none)').matches) return;
      window.clearTimeout(closeTimer);
      setPanel(item, true);
    });
    item.addEventListener('mouseleave', function () {
      if (window.matchMedia('(hover: none)').matches) return;
      window.clearTimeout(closeTimer);
      closeTimer = window.setTimeout(function () { setPanel(item, false); }, 140);
    });

    // Close once focus leaves the whole item
    item.addEventListener('focusout', function (e) {
      if (!item.contains(e.relatedTarget)) setPanel(item, false);
    });

    btn.addEventListener('keydown', function (e) {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setPanel(item, true);
        var first = $('.navpanel__link', item);
        if (first) window.setTimeout(function () { first.focus(); }, 20);
      }
    });

    item.addEventListener('keydown', function (e) {
      if (e.key !== 'Escape') return;
      setPanel(item, false);
      btn.focus();
    });
  });

  document.addEventListener('click', function (e) {
    if (openPanel && !openPanel.contains(e.target)) setPanel(openPanel, false);
  });

  /* --- Mobile nav sub-sections ------------------------------------------ */
  $$('.mnav__btn').forEach(function (btn) {
    btn.addEventListener('click', function () {
      var open = btn.getAttribute('aria-expanded') === 'true';
      var sub = document.getElementById(btn.getAttribute('aria-controls'));
      btn.setAttribute('aria-expanded', open ? 'false' : 'true');
      if (sub) sub.dataset.open = open ? 'false' : 'true';
    });
  });

  /* --- Hero scroller -----------------------------------------------------
     The slides advance with the page, not on a timer. The stage is sticky inside a tall
     track (see .heroscroll in site.css) and the active slide is derived from how far
     through that track the page has scrolled.

     Deliberately NOT a wheel handler. Calling preventDefault on wheel to fake a
     scroller breaks trackpad momentum, the scrollbar, Page Down, Home/End and anything
     driving the page programmatically. Mapping native scroll position costs nothing and
     keeps all of that working. */
  var scroller = $('[data-hero-scroll]');
  if (scroller) {
    var track = $('[data-hero-track]', scroller);
    var stage = $('.heroscroll__stage', scroller);
    var slides = $$('[data-slide]', scroller);
    var dots = $$('[data-dot]', scroller);
    var statusEl = $('[data-slide-status]', scroller);
    var index = -1;
    var top = 0, travel = 1;

    /* Videos ship with preload="none" so an off-screen slide costs nothing. That means
       the first play() can land before any data exists and reject, so wait for data
       before starting, and only seek once metadata is actually available. */
    function playVideo(vid) {
      if (vid.preload === 'none') { vid.preload = 'auto'; vid.load(); }
      var go = function () {
        try { if (vid.readyState >= 1) vid.currentTime = 0; } catch (e) { /* not seekable yet */ }
        var pr = vid.play();
        if (pr && pr.catch) pr.catch(function () { /* blocked by policy; the poster stands in */ });
      };
      if (vid.readyState >= 2) go();
      else vid.addEventListener('canplay', go, { once: true });
    }

    function apply(next, announce) {
      if (next === index) return;
      index = next;
      slides.forEach(function (sl, i) {
        var on = i === index;
        sl.dataset.active = on ? 'true' : 'false';
        if (on) sl.removeAttribute('aria-hidden'); else sl.setAttribute('aria-hidden', 'true');
        // keep off-screen slides out of the tab order
        $$('a, button', sl).forEach(function (el) { el.tabIndex = on ? 0 : -1; });

        var vid = sl.querySelector('[data-hero-video]');
        if (vid) { if (on && !reduced) playVideo(vid); else vid.pause(); }
      });
      dots.forEach(function (d, i) { d.setAttribute('aria-selected', i === index ? 'true' : 'false'); });
      if (announce && statusEl) statusEl.textContent = 'Slide ' + (index + 1) + ' of ' + slides.length;
    }

    function measure() {
      // The stage sticks `top: var(--hdr-h)` below the viewport top, so it starts pinning
      // one header-height BEFORE the track reaches the top of the screen. Measuring from
      // the track's own offset instead would push the whole mapping down by that amount,
      // and the last slide would never quite be reached before the hero scrolled away.
      var stickyTop = parseFloat(window.getComputedStyle(stage).top) || 0;
      top = track.getBoundingClientRect().top + window.pageYOffset - stickyTop;
      // How far the page scrolls while the stage stays pinned.
      travel = Math.max(track.offsetHeight - stage.offsetHeight, 1);
    }

    function sync() {
      if (!isDriven()) return;
      var p = (window.pageYOffset - top) / travel;
      if (p < 0) p = 0; else if (p > 1) p = 1;
      apply(Math.round(p * (slides.length - 1)), true);
    }

    // Scroll-driving is desktop only — see the .heroscroll media query in site.css.
    var driven = window.matchMedia('(min-width: 761px)');
    var isDriven = function () { return driven.matches; };

    /* On desktop the dots and arrows move the page, so scroll position stays the single
       source of truth rather than the slide being swapped behind its back. On mobile
       there is no scroll track to move through, so they switch the slide directly. */
    function goTo(i) {
      if (i < 0) i = 0; else if (i > slides.length - 1) i = slides.length - 1;
      if (!isDriven()) { apply(i, true); return; }
      var y = top + travel * (i / (slides.length - 1));
      if (reduced) window.scrollTo(0, y);
      else window.scrollTo({ top: y, behavior: 'smooth' });
    }

    dots.forEach(function (d) {
      d.addEventListener('click', function () { goTo(parseInt(d.dataset.dot, 10)); });
    });
    var nextBtn = $('[data-next]', scroller), prevBtn = $('[data-prev]', scroller);
    if (nextBtn) nextBtn.addEventListener('click', function () { goTo(index + 1); });
    if (prevBtn) prevBtn.addEventListener('click', function () { goTo(index - 1); });

    scroller.addEventListener('keydown', function (e) {
      if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
      e.preventDefault();
      goTo(index + (e.key === 'ArrowRight' ? 1 : -1));
    });

    // A clearly horizontal swipe still jumps a slide; vertical is left to the page,
    // because vertical is now how the hero advances in the first place.
    var x0 = null, y0 = null;
    scroller.addEventListener('touchstart', function (e) {
      x0 = e.changedTouches[0].clientX; y0 = e.changedTouches[0].clientY;
    }, { passive: true });
    scroller.addEventListener('touchend', function (e) {
      if (x0 === null) return;
      var dx = e.changedTouches[0].clientX - x0;
      var dy = e.changedTouches[0].clientY - y0;
      if (Math.abs(dx) > 45 && Math.abs(dx) > Math.abs(dy) * 1.5) goTo(index + (dx < 0 ? 1 : -1));
      x0 = y0 = null;
    }, { passive: true });

    // The track is only given its height once this runs, so with scripting off the hero
    // stays a normal static first slide rather than several blank viewports.
    if (slides.length > 1) {
      track.setAttribute('data-scroll-ready', 'true');
      measure();
      apply(0, false);
      sync();
      var ticking = false;
      window.addEventListener('scroll', function () {
        if (ticking) return;
        ticking = true;
        window.requestAnimationFrame(function () { sync(); ticking = false; });
      }, { passive: true });
      window.addEventListener('resize', function () { measure(); sync(); });
      // Crossing the breakpoint changes which mode is in force; re-measure and, when
      // dropping to the tap-driven layout, leave whatever slide is showing in place.
      var onModeChange = function () { measure(); sync(); };
      if (driven.addEventListener) driven.addEventListener('change', onModeChange);
      else if (driven.addListener) driven.addListener(onModeChange);
      window.addEventListener('load', function () { measure(); sync(); });
      // requestAnimationFrame is throttled in a background tab, so a scroll that happened
      // while hidden may not have been applied. Resync on the way back.
      // requestAnimationFrame does not run in a background tab, so a queued frame can
      // leave `ticking` stuck true. Clear it on the way back, or scrolling stays dead.
      document.addEventListener('visibilitychange', function () {
        if (document.hidden) return;
        ticking = false; measure(); sync();
      });
      // bfcache restores the scroll position without firing a scroll event.
      window.addEventListener('pageshow', function () { ticking = false; measure(); sync(); });
    } else {
      apply(0, false);
    }
  }

  /* --- Reveal on scroll -------------------------------------------------- */
  var revealTargets = $$('[data-reveal-group] > *').concat($$('[data-reveal]'));
  if (!reduced && 'IntersectionObserver' in window && revealTargets.length) {
    // Shorter stagger, capped at three steps. At 55ms x 6 the last card in a row was
    // still arriving a third of a second after the first, so a fast scroller met a
    // half-empty grid. Fires earlier too: rootMargin now pre-empts the viewport
    // instead of waiting until an element is 8% inside it.
    revealTargets.forEach(function (el, i) {
      el.setAttribute('data-reveal', '');
      el.style.transitionDelay = Math.min(i % 8, 3) * 40 + 'ms';
    });
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) return;
        entry.target.dataset.shown = 'true';
        io.unobserve(entry.target);
      });
    }, { rootMargin: '0px 0px 12% 0px', threshold: 0 });
    revealTargets.forEach(function (el) { io.observe(el); });
  } else {
    revealTargets.forEach(function (el) { el.dataset.shown = 'true'; });
  }

  /* --- Animated counters ------------------------------------------------- */
  var counters = $$('[data-count-to]');
  if (counters.length && !reduced && 'IntersectionObserver' in window) {
    var cio = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) return;
        var el = entry.target;
        cio.unobserve(el);
        var raw = String(el.dataset.countTo).replace(/[^0-9.]/g, '');
        var target = parseFloat(raw);
        if (!isFinite(target)) return;
        var suffix = el.dataset.suffix || '';
        var decimals = (raw.split('.')[1] || '').length;
        var start = performance.now(), dur = 1100, done = false;
        // The true figure is server-rendered; this animation temporarily replaces it
        // with a partial one. requestAnimationFrame stops in a hidden or throttled
        // tab, which left a frozen, WRONG number on screen - "2 Partners" for a firm
        // with 8. A timer that cannot be throttled away snaps to the real value, so
        // an interrupted animation degrades to the fact rather than to a misstatement.
        var settle = function () {
          if (done) return;
          done = true;
          el.textContent = raw + suffix;
        };
        var step = function (t) {
          if (done) return;
          var p = Math.min((t - start) / dur, 1);
          if (p >= 1) return settle();
          var eased = 1 - Math.pow(1 - p, 3);
          el.textContent = (target * eased).toFixed(decimals) + suffix;
          requestAnimationFrame(step);
        };
        setTimeout(settle, dur + 500);
        requestAnimationFrame(step);
      });
    }, { threshold: 0.4 });
    counters.forEach(function (el) { cio.observe(el); });
  }

  /* --- Accordion --------------------------------------------------------- */
  $$('.acc__btn').forEach(function (btn) {
    btn.addEventListener('click', function () {
      var open = btn.getAttribute('aria-expanded') === 'true';
      var panel = document.getElementById(btn.getAttribute('aria-controls'));
      btn.setAttribute('aria-expanded', open ? 'false' : 'true');
      if (panel) panel.dataset.open = open ? 'false' : 'true';
    });
  });

  /* --- Insights filter + search ------------------------------------------ */
  var list = $('[data-insight-list]');
  if (list) {
    var items = $$('[data-article]', list);
    var emptyBox = $('[data-insight-empty]');
    var search = $('[data-insight-search]');
    var active = 'all';

    var apply = function () {
      var q = (search && search.value || '').trim().toLowerCase();
      var shown = 0;
      items.forEach(function (it) {
        var okCat = active === 'all' || it.dataset.category === active;
        var okQ = !q || (it.dataset.search || '').indexOf(q) !== -1;
        var show = okCat && okQ;
        it.hidden = !show;
        if (show) shown++;
      });
      if (emptyBox) emptyBox.hidden = shown !== 0;
      list.hidden = shown === 0;
    };

    $$('[data-filter]').forEach(function (btn) {
      btn.addEventListener('click', function () {
        active = btn.dataset.filter;
        $$('[data-filter]').forEach(function (b) {
          b.setAttribute('aria-pressed', b === btn ? 'true' : 'false');
        });
        apply();
      });
    });
    if (search) search.addEventListener('input', apply);
  }

  /* --- Forms -------------------------------------------------------------
     Validation runs fully client-side. Submission is deliberately honest:
     with no endpoint configured the user is told the form is not connected
     rather than being shown a fake success message. See BACKEND-SPEC.md. */

  // Injected by the build from forms.endpoint in content/site.json.
  var ENDPOINT = (window.SK_CONFIG && window.SK_CONFIG.formsEndpoint) || '';

  var messages = {
    required: 'This field is required.',
    email: 'Enter a valid email address.',
    url: 'Enter a full web address, starting with http:// or https://',
    consent: 'We need your consent before we can proceed.',
    file: 'That file is too large. The limit is 5 MB.',
  };

  function showError(form, name, msg) {
    var holder = form.querySelector('[data-error-for="' + name + '"]');
    var input = form.querySelector('[name="' + name + '"]');
    if (holder) { holder.textContent = msg; holder.dataset.show = msg ? 'true' : 'false'; }
    if (input) {
      if (msg) input.setAttribute('aria-invalid', 'true');
      else input.removeAttribute('aria-invalid');
    }
  }

  function validate(form) {
    var firstBad = null;
    $$('[data-error-for]', form).forEach(function (p) { p.dataset.show = 'false'; p.textContent = ''; });
    $$('[aria-invalid]', form).forEach(function (i) { i.removeAttribute('aria-invalid'); });

    $$('input, select, textarea', form).forEach(function (el) {
      if (el.type === 'hidden' || el.closest('[aria-hidden="true"]')) return;
      var name = el.name, msg = '';

      if (el.type === 'checkbox') {
        if (el.required && !el.checked) msg = messages.consent;
      } else if (el.required && !el.value.trim()) {
        msg = messages.required;
      } else if (el.type === 'email' && el.value && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(el.value)) {
        msg = messages.email;
      } else if (el.type === 'url' && el.value && !/^https?:\/\/.+/.test(el.value)) {
        msg = messages.url;
      } else if (el.type === 'file' && el.files && el.files[0] && el.files[0].size > 5 * 1024 * 1024) {
        msg = messages.file;
      }

      if (msg) {
        showError(form, name, msg);
        if (!firstBad) firstBad = el;
      }
    });

    if (firstBad) {
      firstBad.focus();
      firstBad.scrollIntoView({ block: 'center', behavior: reduced ? 'auto' : 'smooth' });
    }
    return !firstBad;
  }

  $$('[data-form]').forEach(function (form) {
    var status = $('[data-status]', form);

    form.addEventListener('submit', function (e) {
      e.preventDefault();
      if (!validate(form)) return;

      // Honeypot — a real person never fills this in.
      var hp = form.querySelector('[name="company_website"]');
      if (hp && hp.value) return;

      var btn = form.querySelector('button[type="submit"]');

      if (!ENDPOINT) {
        if (status) {
          status.dataset.state = 'info';
          status.innerHTML = '<strong>Your details have not been sent.</strong> This form is validated and ready, ' +
            'but it is not yet connected to a mail or storage service, so nothing would reach the firm. ' +
            'Please contact us directly in the meantime. (Developers: set <code>forms.endpoint</code> in ' +
            '<code>content/site.json</code>, or build the service described in BACKEND-SPEC.md.)';
          status.scrollIntoView({ block: 'nearest', behavior: reduced ? 'auto' : 'smooth' });
        }
        return;
      }

      if (btn) { btn.disabled = true; btn.dataset.label = btn.textContent; btn.textContent = 'Sending…'; }
      if (status) { status.dataset.state = 'info'; status.textContent = 'Sending…'; }

      fetch(ENDPOINT, { method: 'POST', body: new FormData(form), headers: { Accept: 'application/json' } })
        .then(function (r) {
          if (!r.ok) throw new Error('Request failed: ' + r.status);
          form.reset();
          if (status) {
            status.dataset.state = 'ok';
            status.textContent = 'Thank you. We have received your message and will be in touch.';
          }
        })
        .catch(function () {
          if (status) {
            status.dataset.state = 'err';
            status.textContent = 'Sorry — something went wrong sending that. Please try again, or contact us directly.';
          }
        })
        .finally(function () {
          if (btn) { btn.disabled = false; btn.textContent = btn.dataset.label || 'Submit'; }
          if (status) status.scrollIntoView({ block: 'nearest', behavior: reduced ? 'auto' : 'smooth' });
        });
    });

    // Clear an error as soon as the person starts fixing it
    $$('input, select, textarea', form).forEach(function (el) {
      el.addEventListener('input', function () { showError(form, el.name, ''); });
    });
  });

  /* --- Generic Carousel Helper ------------------------------------------ */
  function initCarousel(cfg) {
    var container = $(cfg.container);
    if (!container) return;
    var slides = $$(cfg.slide, container);
    if (!slides.length) return;
    var dots = $$(cfg.dot);
    var prevBtn = $(cfg.prev);
    var nextBtn = $(cfg.next);
    var counter = $(cfg.counter);
    var curr = 0;
    var autoTimer = null;
    var delay = cfg.delay || 6500;
    var userPaused = false;

    function render(idx) {
      curr = (idx + slides.length) % slides.length;
      slides.forEach(function (s, i) {
        var on = i === curr;
        s.dataset.active = on ? 'true' : 'false';
        if (on) s.removeAttribute('aria-hidden'); else s.setAttribute('aria-hidden', 'true');
        $$('a, button', s).forEach(function (el) { el.tabIndex = on ? 0 : -1; });
      });
      dots.forEach(function (d, i) {
        d.classList.toggle(cfg.activeDotClass || 'partner-dot--active', i === curr);
        d.setAttribute('aria-selected', i === curr ? 'true' : 'false');
      });
      if (counter) counter.textContent = (curr + 1) + ' / ' + slides.length;
    }

    function stopAuto() { if (autoTimer) { clearInterval(autoTimer); autoTimer = null; } }
    function startAuto() {
      stopAuto();
      if (reduced || userPaused || document.hidden || slides.length < 2) return;
      autoTimer = setInterval(function () { render(curr + 1); }, delay);
    }

    if (prevBtn) prevBtn.addEventListener('click', function () { userPaused = true; stopAuto(); render(curr - 1); });
    if (nextBtn) nextBtn.addEventListener('click', function () { userPaused = true; stopAuto(); render(curr + 1); });

    dots.forEach(function (d, i) {
      d.addEventListener('click', function () { userPaused = true; stopAuto(); render(i); });
    });

    container.addEventListener('mouseenter', stopAuto);
    container.addEventListener('mouseleave', function () { if (!userPaused) startAuto(); });
    container.addEventListener('focusin', stopAuto);
    container.addEventListener('focusout', function (e) {
      if (!container.contains(e.relatedTarget) && !userPaused) startAuto();
    });

    // Touch swipe
    var touchX = null, touchY = null;
    container.addEventListener('touchstart', function (e) {
      touchX = e.changedTouches[0].clientX;
      touchY = e.changedTouches[0].clientY;
    }, { passive: true });
    container.addEventListener('touchend', function (e) {
      if (touchX === null) return;
      var dx = e.changedTouches[0].clientX - touchX;
      var dy = e.changedTouches[0].clientY - touchY;
      if (Math.abs(dx) > 40 && Math.abs(dx) > Math.abs(dy) * 1.4) {
        userPaused = true;
        stopAuto();
        render(curr + (dx < 0 ? 1 : -1));
      }
      touchX = null; touchY = null;
    }, { passive: true });

    render(0);
    startAuto();
  }

  /* --- Partner Spotlight Carousel ---------------------------------------- */
  initCarousel({
    container: '[data-partner-carousel]',
    slide: '[data-partner-slide]',
    dot: '[data-partner-dot]',
    prev: '[data-partner-prev]',
    next: '[data-partner-next]',
    counter: '[data-partner-counter]',
    delay: 7000,
    activeDotClass: 'partner-dot--active'
  });

  /* --- Client Testimonial Carousel --------------------------------------- */
  initCarousel({
    container: '[data-test-carousel]',
    slide: '[data-test-slide]',
    dot: '[data-test-dot]',
    prev: '[data-test-prev]',
    next: '[data-test-next]',
    counter: '[data-test-counter]',
    delay: 8500,
    activeDotClass: 'partner-dot--active'
  });

  /* --- 360° Ecosystem Interactive Widget --------------------------------- */
  var ecoWidget = $('[data-eco-widget]');
  if (ecoWidget) {
    var ecoTriggers = $$('[data-eco-trigger]', ecoWidget);
    var ecoPanels = $$('[data-eco-panel]', ecoWidget);

    function selectEco(id) {
      ecoTriggers.forEach(function (btn) {
        var on = btn.dataset.ecoTrigger === id;
        btn.classList.toggle('eco-nav__btn--active', on);
        btn.setAttribute('aria-selected', on ? 'true' : 'false');
      });
      ecoPanels.forEach(function (p) {
        var on = p.dataset.ecoPanel === id;
        p.classList.toggle('eco-panel--active', on);
      });
    }

    ecoTriggers.forEach(function (btn) {
      btn.addEventListener('click', function () {
        selectEco(btn.dataset.ecoTrigger);
      });
    });
  }

  /* --- Client Growth Navigator Interactive Widget ----------------------- */
  var navWidget = $('[data-navigator-widget]');
  if (navWidget) {
    var navTriggers = $$('[data-nav-trigger]', navWidget);
    var navPanels = $$('[data-nav-panel]', navWidget);

    function selectNav(id) {
      navTriggers.forEach(function (btn) {
        var on = btn.dataset.navTrigger === id;
        btn.classList.toggle('navigator-tab--active', on);
        btn.setAttribute('aria-selected', on ? 'true' : 'false');
      });
      navPanels.forEach(function (p) {
        var on = p.dataset.navPanel === id;
        p.classList.toggle('navigator-panel--active', on);
      });
    }

    navTriggers.forEach(function (btn) {
      btn.addEventListener('click', function () {
        selectNav(btn.dataset.navTrigger);
      });
    });
  }

})();

