(() => {
  const title = document.getElementById("title-page");
  const mapSection = document.getElementById("map-section");
  if (!title || !mapSection) return;

  function clamp01(x) {
    if (x < 0) return 0;
    if (x > 1) return 1;
    return x;
  }

  let rafId = 0;
  function schedule() {
    if (rafId) return;
    rafId = requestAnimationFrame(() => {
      rafId = 0;
      update();
    });
  }

  function update() {
    const vh = window.innerHeight || 1;
    const rect = title.getBoundingClientRect();
    const total = Math.max(1, title.offsetHeight - vh);
    const progressed = clamp01((-rect.top) / total);

    document.documentElement.style.setProperty("--intro-progress", progressed.toFixed(4));

    // Toggle pointer events so you don't accidentally interact with the map mid-transition.
    mapSection.style.pointerEvents = progressed > 0.9 ? "auto" : "none";
  }

  update();
  window.addEventListener("scroll", schedule, { passive: true });
  window.addEventListener("resize", schedule);
})();

(() => {
  const card = document.querySelector(".title-card");
  if (!card) return;

  const h1 = card.querySelector("h1");
  const h2 = card.querySelector("h2");
  const subtitle = card.querySelector(".title-subtitle");
  const cta = card.querySelector(".title-cta");
  if (!h1 || !h2 || !subtitle || !cta) return;

  const reduceMotion = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  const original = {
    h1: h1.textContent || "",
    h2: h2.textContent || "",
    subtitle: subtitle.textContent || ""
  };

  // Always start with CTA hidden; we'll reveal it immediately if we skip typing.
  cta.classList.add("is-hidden");
  cta.classList.remove("is-visible");

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function makeCursor() {
    const cursor = document.createElement("span");
    cursor.className = "type-cursor";
    cursor.textContent = "▍";
    cursor.setAttribute("aria-hidden", "true");
    return cursor;
  }

  async function typeInto(el, text, opts) {
    const speed = opts?.speed ?? 22;
    const startDelay = opts?.startDelay ?? 0;
    const afterDelay = opts?.afterDelay ?? 240;

    el.textContent = "";
    const cursor = makeCursor();
    el.appendChild(cursor);
    if (startDelay) await sleep(startDelay);

    for (let i = 0; i < text.length; i++) {
      cursor.insertAdjacentText("beforebegin", text[i]);
      await sleep(speed);
    }

    cursor.remove();
    if (afterDelay) await sleep(afterDelay);
  }

  async function runTyping() {
    if (reduceMotion) {
      h1.textContent = original.h1;
      h2.textContent = original.h2;
      subtitle.textContent = original.subtitle;
      cta.classList.remove("is-hidden");
      cta.classList.add("is-visible");
      return;
    }

    // Clear text before the first paint of typing (scripts are loaded at end of body).
    h1.textContent = "";
    h2.textContent = "";
    subtitle.textContent = "";

    await typeInto(h1, original.h1, { speed: 18, startDelay: 220, afterDelay: 220 });
    await typeInto(h2, original.h2, { speed: 14, startDelay: 80, afterDelay: 190 });
    await typeInto(subtitle, original.subtitle, { speed: 12, startDelay: 60, afterDelay: 260 });

    cta.classList.remove("is-hidden");
    cta.classList.add("is-visible");
  }

  function revealWithoutTyping() {
    h1.textContent = original.h1;
    h2.textContent = original.h2;
    subtitle.textContent = original.subtitle;
    cta.classList.remove("is-hidden");
    cta.classList.add("is-visible");
  }

  let started = false;
  function startTypingIfVisible() {
    if (started) return;
    started = true;

    const rect = card.getBoundingClientRect();
    const vh = window.innerHeight || 1;
    // Visible if a meaningful chunk of the intro card is on screen.
    const visible = rect.bottom > vh * 0.2 && rect.top < vh * 0.8;

    if (!visible) {
      revealWithoutTyping();
      return;
    }

    runTyping();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      // Wait a beat so CSS/layout is settled before we clear+type.
      setTimeout(startTypingIfVisible, 120);
    }, { once: true });
  } else {
    setTimeout(startTypingIfVisible, 120);
  }
})();
