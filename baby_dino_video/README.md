# Baby Dino Surprise Egg (10 s, 9:16, loopable)

`baby_dino_surprise_egg.mp4` — 1080×1920, 30 fps, H.264 + AAC, exactly 10 s.

| Time | What happens | Sound |
|---|---|---|
| 0.00–0.75 | Rainbow polka-dot egg wobbles in the center | tick-tock + rising notes |
| 0.75 | Egg pops open, confetti burst, baby dino springs out | POP + sparkle glissando |
| 1.0–2.2 | Dino giggles, gentle zoom-in on its face, then looks at camera | "hee-hee" giggle |
| 2 / 4 / 6 s | Same dance move ×3: lean left, lean right, hop, land + twirl | boing on each hop |
| 3.5 / 5.5 / 7.5 | Red, then blue, then yellow balloon pops in on the last beat | bloop |
| 8.0–9.3 | Balloons float up, dino waves goodbye (gentle zoom-in) | whistle + giggle |
| 9.3–10 | Egg closes back around the dino, so the last frame meets the first | click |

Music: original 120 BPM nursery-rhyme loop in C (mallet melody, ukulele, soft bass), synthesized in code, so it's royalty-free. The camera stays centered with no cuts and nothing flashes.

## Files
- `scene.js` — Three.js scene; every frame is a pure function of time `t`.
- `index.html` — live preview (serve the folder, open the page, press ▶).
- `soundtrack.py` — generates `soundtrack.wav` (needs numpy).
- `render.mjs` — renders frames in headless Chromium and muxes the MP4 (needs Playwright + ffmpeg).

```sh
python3 soundtrack.py
node render.mjs 30 1080 1920   # -> baby_dino_surprise_egg.mp4
```
