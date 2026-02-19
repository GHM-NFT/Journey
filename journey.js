<script>
(() => {
  const TOTAL = 6;
  const KEY   = "ghm_collected";
  const START_ID = "journey-start";
  const PROGRESS_TOP_ID = "progress-top";

  // ===== EDITOR OVERRIDE =====
  const EDIT_MODE =
    document.body.classList.contains('elementor-editor-active') ||
    document.body.classList.contains('elementor-editor-preview');

  if (EDIT_MODE) {
    document.body.classList.remove('ghm-complete');
    localStorage.setItem(KEY, "0");
  }

  const START = () => document.getElementById(START_ID);
  const clamp = (n,a,b)=>Math.max(a,Math.min(b,n));
  const $ = (sel) => (START() ? START().querySelector(sel) : null);

  function getCount(){
    const n = parseInt(localStorage.getItem(KEY) || "0", 10);
    return Number.isFinite(n) ? clamp(n,0,TOTAL) : 0;
  }

  function setCount(n){
    localStorage.setItem(KEY, String(clamp(n,0,TOTAL)));
    apply();
  }

  function setSectionSwap(collected){
    document.body.classList.toggle("ghm-complete", collected >= TOTAL);
  }

function setBeatState(i, collected){
  const beat = $(`.journey-beat-${i}`); // your existing beat card wrapper
  const cardDefault = $(`.journey-card-${i}.journey-card-default`);
  const cardCollected = $(`.journey-card-${i}.journey-card-collected`);

  const isCollected = i <= collected;
  const isNext      = i === collected + 1;
  const isLocked    = i > collected + 1;

  // keep original beat state if you still use it anywhere
  if (beat){
    beat.classList.toggle("journey-collected", isCollected);
    beat.classList.toggle("journey-next", isNext);
    beat.classList.toggle("journey-locked", isLocked);
  }

  // apply state directly onto both columns
  [cardDefault, cardCollected].forEach(el => {
    if(!el) return;
    el.classList.toggle("journey-collected", isCollected);
    el.classList.toggle("journey-next", isNext);
    el.classList.toggle("journey-locked", isLocked);
  });

  // button logic unchanged
  const btn  = $(`.journey-btn-${i}`);
  if (btn){
    btn.classList.toggle("journey-collected", isCollected);
    btn.classList.toggle("journey-next", isNext);
    btn.classList.toggle("journey-locked", isLocked);

    btn.style.pointerEvents = isNext ? "auto" : "none";
    btn.onclick = isNext ? () => setCount(collected + 1) : null;
  }
}



  function applyTopProgress(collected){
    const top = document.getElementById(PROGRESS_TOP_ID);
    if (!top) return;

    const t = top.querySelector(".journey-progress-text");
    if (t) t.textContent = `COLLECTED ${collected}/${TOTAL}`;

    for (let i=1;i<=TOTAL;i++){
      const icon = top.querySelector(`.progress-icon-${i}`);
      if(!icon) continue;

      const isCollected = i <= collected;
      const isNext      = i === collected + 1;
      const isLocked    = i > collected + 1;

      icon.classList.toggle("is-collected", isCollected);
      icon.classList.toggle("is-next", isNext);
      icon.classList.toggle("is-locked", isLocked);
    }
  }

  function apply(){
    const collected = getCount();
    setSectionSwap(collected);

    for (let i=1;i<=TOTAL;i++){
      setBeatState(i, collected);
    }

    applyTopProgress(collected);
  }

  document.addEventListener("DOMContentLoaded", apply);

  window.GHM = {
    get: getCount,
    set: setCount,
    reset: () => setCount(0),
    complete: () => setCount(TOTAL)
  };
})();
</script>
