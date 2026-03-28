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

