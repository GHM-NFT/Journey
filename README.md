# Journey / Beats — Module README

This folder contains the canonical source for the Elementor “Journey / Beats” UI logic:
- `journey.js` — state + rendering + button wiring
- `journey.css` — visibility + locked/collected styles
- (optional) additional files as the project grows

The live site may paste these files into Elementor (Custom Code / HTML widget), but **the repo is the source of truth**.

---

## 1) Goal

Render 6 “beats” (cards) with reliable states:
- **default** (not collected)
- **collected**
- **locked** (optional, inert UI)

Bug to avoid: Beat 1 updates but beats 2–5 don’t render correctly due to selector scope or duplicate IDs inside Elementor templates.

---

## 2) Markup contract (required HTML structure)

### Preferred approach: `data-beat` + state wrappers

Each beat card must have:

- An outer wrapper:
  - class: `.journey-card`
  - attribute: `data-beat="N"` (N = 1..6)

- Two inner state wrappers (both present in DOM):
  - `.journey-card-default`
  - `.journey-card-collected`

Optional:
- `.journey-card-locked` wrapper or a locked class on the outer card (recommended).

**Example (one card):**
```html
<div class="journey-card" data-beat="1">
  <div class="journey-card-default">
    <!-- default UI -->
  </div>

  <div class="journey-card-collected">
    <!-- collected UI -->
  </div>
</div>
