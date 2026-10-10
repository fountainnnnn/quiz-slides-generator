(function () {
  const ready = (callback) => {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", callback, { once: true });
      return;
    }
    callback();
  };

  ready(() => {
    const gsap = window.gsap;
    if (!gsap) {
      document.documentElement.classList.add("motion-unavailable");
      return;
    }

    const prefersReducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    document.documentElement.classList.add("gsap-enhanced");

    if (prefersReducedMotion) {
      document.documentElement.classList.add("reduced-motion");
      return;
    }

    if (window.ScrollTrigger) {
      gsap.registerPlugin(window.ScrollTrigger);
    }

    const lenis = initLenis(gsap);

    gsap.defaults({
      ease: "power3.out",
      duration: 0.72,
    });

    animateChrome(gsap);
    animatePageIntro(gsap);
    animateScrollReveals(gsap);
    animateProjectPreviews(gsap);
    wireMicroInteractions(gsap);
    wireCertificateModal(gsap);

    if (lenis && window.ScrollTrigger) {
      window.ScrollTrigger.refresh();
    }
  });

  function initLenis(gsap) {
    if (!window.Lenis) return null;

    const lenis = new window.Lenis({
      anchors: {
        offset: 92,
      },
      lerp: 0.14,
      wheelMultiplier: 1,
      touchMultiplier: 1.16,
      smoothWheel: true,
      gestureOrientation: "vertical",
      prevent: (node) => Boolean(node.closest?.(
        "[data-lenis-prevent], .modal, .modal-dialog, .chat-panel, textarea, select, input",
      )),
    });

    document.documentElement.classList.add("lenis-enhanced");
    window.portfolioLenis = lenis;
    wireDirectionRelease(lenis);

    if (window.ScrollTrigger) {
      lenis.on("scroll", window.ScrollTrigger.update);
    }

    gsap.ticker.add((time) => {
      lenis.raf(time * 1000);
    });
    gsap.ticker.lagSmoothing(0);

    return lenis;
  }

  function wireDirectionRelease(lenis) {
    let lastInputDirection = 0;
    let lastTouchY = null;

    window.addEventListener("wheel", (event) => {
      const direction = Math.sign(event.deltaY);
      if (!direction) return;
      if (lastInputDirection && direction !== lastInputDirection) {
        releaseScrollMomentum(lenis);
      }
      lastInputDirection = direction;
    }, { passive: true, capture: true });

    window.addEventListener("touchstart", (event) => {
      lastTouchY = event.touches[0]?.clientY ?? null;
    }, { passive: true });

    window.addEventListener("touchmove", (event) => {
      if (lastTouchY == null) return;
      const nextY = event.touches[0]?.clientY;
      if (nextY == null) return;
      const direction = Math.sign(lastTouchY - nextY);
      if (direction && lastInputDirection && direction !== lastInputDirection) {
        releaseScrollMomentum(lenis);
      }
      if (direction) lastInputDirection = direction;
      lastTouchY = nextY;
    }, { passive: true, capture: true });
  }

  function releaseScrollMomentum(lenis) {
    const currentScroll = window.scrollY;
    lenis.animate?.stop?.();
    lenis.targetScroll = currentScroll;
    lenis.animatedScroll = currentScroll;
    lenis.velocity = 0;
    lenis.lastVelocity = 0;
    lenis.setScroll?.(currentScroll);
  }

  function animateChrome(gsap) {
    gsap.from(".site-navbar, .project-app .navbar", {
      y: -18,
      autoAlpha: 0,
      duration: 0.58,
      clearProps: "transform,opacity,visibility",
    });

    gsap.from(".site-footer", {
      scrollTrigger: trigger(".site-footer", "top 96%"),
      y: 18,
      autoAlpha: 0,
      duration: 0.55,
      clearProps: "transform,opacity,visibility",
    });
  }

  function animatePageIntro(gsap) {
    if (document.body.classList.contains("project-app")) {
      revealNow(gsap, ".project-hero-copy > *, .project-product-preview", {
        y: 22,
        stagger: 0.08,
        delay: 0.12,
      });
      return;
    }

    if (document.body.classList.contains("listing-page")) {
      revealNow(gsap, ".listing-hero > *, .project-index-heading > *, .cert-section-heading > *", {
        y: 18,
        stagger: 0.07,
        delay: 0.08,
      });
      return;
    }

    revealNow(gsap, ".hero-text > *, .hero-photo", {
      y: 24,
      stagger: 0.08,
      delay: 0.22,
    });
  }

  function animateScrollReveals(gsap) {
    revealBatch(gsap, [
      ".section-title",
      ".about-container p",
      ".skill-category",
      ".project-card",
      ".repo-card",
      ".cert-card",
      ".timeline-item",
      ".certifications-heading > *",
      ".project-section-heading > *",
      ".project-form-card",
      ".project-chat-grid > *",
      ".accordion-item",
      ".listing-actions > *",
    ].join(", "));
  }

  function animateProjectPreviews(gsap) {
    if (!window.ScrollTrigger) return;

    gsap.utils.toArray(".preview-window, .project-card-thumb img, .cert-preview img").forEach((element) => {
      gsap.to(element, {
        yPercent: element.matches(".preview-window") ? -4 : -3,
        ease: "none",
        scrollTrigger: {
          trigger: element,
          start: "top bottom",
          end: "bottom top",
          scrub: 0.55,
        },
      });
    });
  }

  function wireMicroInteractions(gsap) {
    const canHover = window.matchMedia("(hover: hover) and (pointer: fine)").matches;
    if (!canHover) return;

    gsap.utils.toArray(".project-card, .cert-card, .repo-card").forEach((card) => {
      card.classList.add("motion-card");
      card.addEventListener("pointerenter", () => {
        gsap.to(card, { y: -6, scale: 1.01, duration: 0.24, overwrite: "auto" });
      });
      card.addEventListener("pointerleave", () => {
        gsap.to(card, { y: 0, scale: 1, duration: 0.34, overwrite: "auto" });
      });
    });

    gsap.utils.toArray(".btn, .chat-toggle, .view-more-link").forEach((element) => {
      const moveX = gsap.quickTo(element, "x", { duration: 0.28, ease: "power3.out" });
      const moveY = gsap.quickTo(element, "y", { duration: 0.28, ease: "power3.out" });

      element.addEventListener("pointermove", (event) => {
        const rect = element.getBoundingClientRect();
        moveX((event.clientX - rect.left - rect.width / 2) * 0.08);
        moveY((event.clientY - rect.top - rect.height / 2) * 0.12);
      });
      element.addEventListener("pointerleave", () => {
        moveX(0);
        moveY(0);
      });
    });
  }

  function wireCertificateModal(gsap) {
    const modal = document.getElementById("certModal");
    if (!modal) return;

    modal.addEventListener("shown.bs.modal", () => {
      gsap.fromTo(
        modal.querySelector(".modal-content"),
        { y: 28, scale: 0.98, autoAlpha: 0 },
        { y: 0, scale: 1, autoAlpha: 1, duration: 0.38, clearProps: "transform,opacity,visibility" },
      );
    });
  }

  function revealNow(gsap, selector, options = {}) {
    const elements = gsap.utils.toArray(selector);
    if (!elements.length) return;

    gsap.from(elements, {
      autoAlpha: 0,
      y: options.y ?? 20,
      stagger: options.stagger ?? 0.06,
      delay: options.delay ?? 0,
      duration: options.duration ?? 0.68,
      clearProps: "transform,opacity,visibility",
    });
  }

  function revealBatch(gsap, selector) {
    const elements = gsap.utils.toArray(selector);
    if (!elements.length) return;

    gsap.set(elements, { autoAlpha: 0, y: 24 });

    if (!window.ScrollTrigger) {
      gsap.to(elements, {
        autoAlpha: 1,
        y: 0,
        stagger: 0.045,
        clearProps: "transform,opacity,visibility",
      });
      return;
    }

    window.ScrollTrigger.batch(elements, {
      start: "top 88%",
      once: true,
      onEnter: (batch) => {
        gsap.to(batch, {
          autoAlpha: 1,
          y: 0,
          stagger: 0.055,
          duration: 0.62,
          clearProps: "transform,opacity,visibility",
        });
      },
    });
  }

  function trigger(selector, start) {
    return window.ScrollTrigger && document.querySelector(selector)
      ? { trigger: selector, start, once: true }
      : undefined;
  }
})();
