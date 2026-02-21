Practical JOURNEY'S CODEX template FOR Story/beat handover
Page:
Journey total beats:
Wrapper IDs:
- start:
- complete:
- progress:

Beat classes present:
- journey-beat-N: yes/no
- journey-card-N default/collected: yes/no
- journey-btn-N: yes/no

Progress icons:
- progress-icon-N present: yes/no

Differences from standard:
- ...


What I need from you (minimum viable handoff)
For each Journey page/template, send:
One screenshot of Navigator structure
Showing the wrapper section and beat blocks.
The exact wrapper IDs/classes used
e.g. #journey-start, #progress-top, #journey-complete.
Beat class naming for each index
Confirm you have:
.journey-beat-N
.journey-card-N.journey-card-default
.journey-card-N.journey-card-collected
.journey-btn-N
These are the runtime expectations. 
Journey length per page
Set data-total-beats="X" on #journey-start (1..12). 
Progress icon naming (if used)
.progress-icon-N for the top strip. 
Any per-page differences
Example: some pages hide progress bar, custom button labels, different lock visuals.
Optional but very helpful
Elementor Template JSON export (only when layout is complex or buggy).
Published page URL (staging) + temporary access if needed.
A short mapping table:
NFT stage -> beat number -> story label -> reward trigger.
Recommended working model (best for speed)
You (Elementor side)
Build/adjust visuals and structure.
Keep naming contract intact.
Share screenshots + IDs/classes + any exceptions.
Me (code side)
Keep JS/CSS robust and generic.
Update logic for edge cases.
Return exact copy/paste snippets + what class hooks must exist.
This is exactly aligned with the repo design: logic in JS/CSS, Elementor as rendering shell. 
