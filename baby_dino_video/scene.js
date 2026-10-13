// Baby Dino Surprise Egg — 10 s loopable 9:16 scene.
// Everything is a pure function of time t in [0, 10), so the same code drives
// the live preview (index.html) and the frame-by-frame video render (render.mjs).
import * as THREE from 'three';
import { RoomEnvironment } from './vendor/RoomEnvironment.js';

export const DURATION = 10;

// ---------- timeline (seconds, 120 BPM => 1 beat = 0.5 s) ----------
const POP = 0.75;                        // egg bursts open
const DANCE = [2, 4, 6];                 // three dance reps, one bar each
const BALLOON_AT = DANCE.map(d => d + 1.5); // red, blue, yellow on beat 4
const FLOAT = 8.0;                       // balloons float up, wave goodbye
const CLOSE = 9.3;                       // egg reassembles for a seamless loop

// ---------- helpers ----------
const clamp = (x, a = 0, b = 1) => Math.min(b, Math.max(a, x));
const lerp = (a, b, k) => a + (b - a) * k;
const smooth = (a, b, x) => { const k = clamp((x - a) / (b - a)); return k * k * (3 - 2 * k); };
const easeOutBack = (k, s = 2.2) => { k = clamp(k) - 1; return 1 + k * k * ((s + 1) * k + s); };
const easeOutElastic = k => {
  k = clamp(k);
  if (k === 0 || k === 1) return k;
  return Math.pow(2, -9 * k) * Math.sin((k * 10 - 0.75) * (2 * Math.PI) / 3) + 1;
};
const easeInBack = (k, s = 1.8) => { k = clamp(k); return k * k * ((s + 1) * k - s); };
function rng(seed) { // mulberry32
  return () => {
    seed |= 0; seed = (seed + 0x6D2B79F5) | 0;
    let r = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    r = (r + Math.imul(r ^ (r >>> 7), 61 | r)) ^ r;
    return ((r ^ (r >>> 14)) >>> 0) / 4294967296;
  };
}

// glossy, soft "vinyl toy" material
function toy(color, opts = {}) {
  return new THREE.MeshPhysicalMaterial({
    color, roughness: 0.38, clearcoat: 0.8, clearcoatRoughness: 0.18,
    sheen: 0.12, sheenRoughness: 0.6, sheenColor: new THREE.Color(0xffffff), ...opts,
  });
}
function mesh(geo, mat, { pos = [0, 0, 0], scale = [1, 1, 1], rot = [0, 0, 0] } = {}) {
  const m = new THREE.Mesh(geo, mat);
  m.position.set(...pos); m.scale.set(...scale); m.rotation.set(...rot);
  m.castShadow = true; m.receiveShadow = true;
  return m;
}
const sphere = r => new THREE.SphereGeometry(r, 48, 32);

// ---------- textures ----------
function backgroundTexture() {
  const c = document.createElement('canvas'); c.width = 540; c.height = 960;
  const g = c.getContext('2d');
  const grad = g.createLinearGradient(0, 0, 0, c.height);
  grad.addColorStop(0, '#cfe9ff');   // pastel sky blue
  grad.addColorStop(0.55, '#ffe3f1'); // pastel pink
  grad.addColorStop(1, '#fff3d6');   // pastel cream
  g.fillStyle = grad; g.fillRect(0, 0, c.width, c.height);
  // soft bokeh circles for depth
  const r = rng(7);
  const tints = ['#ffffff', '#fff6b8', '#d9f7e3', '#e6dcff'];
  for (let i = 0; i < 26; i++) {
    const x = r() * c.width, y = r() * c.height * 0.75, rad = 18 + r() * 60;
    const rg = g.createRadialGradient(x, y, 0, x, y, rad);
    rg.addColorStop(0, tints[i % tints.length] + 'aa'); rg.addColorStop(1, tints[i % tints.length] + '00');
    g.fillStyle = rg; g.beginPath(); g.arc(x, y, rad, 0, Math.PI * 2); g.fill();
  }
  const t = new THREE.CanvasTexture(c); t.colorSpace = THREE.SRGBColorSpace; return t;
}
function eggTexture() {
  const c = document.createElement('canvas'); c.width = 1024; c.height = 512;
  const g = c.getContext('2d');
  const bands = ['#ff3b3b', '#ff9f1c', '#ffd60a', '#2ec84a', '#1e90ff', '#9b5de5'];
  const bh = c.height / bands.length;
  bands.forEach((col, i) => { g.fillStyle = col; g.fillRect(0, i * bh, c.width, bh + 1); });
  // white polka dots
  g.fillStyle = '#ffffff';
  for (let row = 0; row < bands.length; row++) {
    for (let k = 0; k < 10; k++) {
      const x = (k + (row % 2) * 0.5) * (c.width / 10), y = (row + 0.5) * bh;
      g.beginPath(); g.arc(x, y, bh * 0.2, 0, Math.PI * 2); g.fill();
    }
  }
  // zig-zag "crack line" at the equator
  g.strokeStyle = '#ffffff'; g.lineWidth = 10; g.beginPath();
  for (let x = 0; x <= c.width; x += 32) g.lineTo(x, c.height / 2 + ((x / 32) % 2 ? -12 : 12));
  g.stroke();
  const t = new THREE.CanvasTexture(c); t.colorSpace = THREE.SRGBColorSpace; t.anisotropy = 8; return t;
}

// ---------- build ----------
export function createScene(canvas, width, height) {
  const renderer = new THREE.WebGLRenderer({ canvas, antialias: true, preserveDrawingBuffer: true });
  renderer.setPixelRatio(1);
  renderer.setSize(width, height, false);
  renderer.toneMapping = THREE.NeutralToneMapping;
  renderer.toneMappingExposure = 1.0;
  renderer.shadowMap.enabled = true;
  renderer.shadowMap.type = THREE.PCFSoftShadowMap;

  const scene = new THREE.Scene();
  scene.background = backgroundTexture();
  const pmrem = new THREE.PMREMGenerator(renderer);
  scene.environment = pmrem.fromScene(new RoomEnvironment(), 0.04).texture;
  scene.environmentIntensity = 0.4;

  const camera = new THREE.PerspectiveCamera(38, width / height, 0.1, 100);

  // lights
  scene.add(new THREE.HemisphereLight(0xfff4fb, 0xbfe3ff, 0.7));
  const key = new THREE.DirectionalLight(0xffffff, 2.4);
  key.position.set(3, 7, 6); key.castShadow = true;
  key.shadow.mapSize.set(2048, 2048); key.shadow.radius = 6; key.shadow.bias = -0.0005;
  Object.assign(key.shadow.camera, { left: -4, right: 4, top: 6, bottom: -2, near: 1, far: 20 });
  scene.add(key);
  const rim = new THREE.DirectionalLight(0xffd1f0, 1.4); rim.position.set(-4, 4, -5); scene.add(rim);

  // round pastel stage
  const stage = mesh(new THREE.CylinderGeometry(2.3, 2.45, 0.35, 96), toy(0xbfe7ff, { roughness: 0.5 }), { pos: [0, -0.175, 0] });
  stage.castShadow = false; scene.add(stage);
  const rimRing = mesh(new THREE.TorusGeometry(2.38, 0.07, 16, 120), toy(0xffffff), { pos: [0, 0, 0], rot: [Math.PI / 2, 0, 0] });
  scene.add(rimRing);

  // ---- baby dino ----
  const dino = new THREE.Group(); scene.add(dino);
  const green = toy(0x1fc94a), belly = toy(0xffe14d), white = toy(0xffffff, { roughness: 0.15 });
  const pupil = toy(0x1b1340, { roughness: 0.1, clearcoat: 1 });
  const blush = toy(0xff7fb0, { roughness: 0.6, clearcoat: 0.2 });
  const shine = new THREE.MeshBasicMaterial({ color: 0xffffff });

  const body = new THREE.Group(); dino.add(body); // squash/stretch pivot at feet
  body.add(mesh(sphere(0.78), green, { pos: [0, 0.85, 0], scale: [1, 1.04, 0.92] }));
  body.add(mesh(sphere(0.6), belly, { pos: [0, 0.8, 0.42], scale: [0.85, 0.95, 0.5] }));

  // legs + feet
  const legs = [];
  for (const s of [-1, 1]) {
    const leg = new THREE.Group(); leg.position.set(0.36 * s, 0.32, 0.08); body.add(leg);
    leg.add(mesh(new THREE.CapsuleGeometry(0.2, 0.18, 8, 24), green, { pos: [0, -0.06, 0] }));
    leg.add(mesh(sphere(0.23), green, { pos: [0, -0.2, 0.1], scale: [1, 0.6, 1.25] }));
    for (const tx of [-0.09, 0, 0.09]) leg.add(mesh(sphere(0.05), white, { pos: [tx, -0.22, 0.36] }));
    legs.push(leg);
  }

  // arms (pivot at shoulder)
  const arms = [];
  for (const s of [-1, 1]) {
    const arm = new THREE.Group(); arm.position.set(0.6 * s, 1.12, 0.25); body.add(arm);
    const a = mesh(new THREE.CapsuleGeometry(0.11, 0.26, 8, 20), green, { pos: [0.12 * s, -0.15, 0.04], rot: [0, 0, 0.75 * s] });
    arm.add(a);
    arm.add(mesh(sphere(0.12), green, { pos: [0.26 * s, -0.28, 0.06] }));
    arms.push(arm);
  }

  // tail curling off to the right so it reads from the front
  const tail = new THREE.Group(); tail.position.set(0.25, 0.55, -0.55); body.add(tail);
  [[0.33, 0, 0, 0], [0.26, 0.28, 0.08, -0.3], [0.2, 0.52, 0.2, -0.5], [0.14, 0.7, 0.36, -0.55]]
    .forEach(([r, x, y, z]) => tail.add(mesh(sphere(r), green, { pos: [x, y, z] })));

  // head
  const head = new THREE.Group(); head.position.set(0, 1.72, 0.05); body.add(head);
  head.add(mesh(sphere(0.72), green, { scale: [1.12, 0.95, 1] }));
  head.add(mesh(sphere(0.44), green, { pos: [0, -0.2, 0.48], scale: [1.15, 0.72, 0.9] }));
  const eyes = [];
  for (const s of [-1, 1]) {
    const eye = new THREE.Group(); eye.position.set(0.3 * s, 0.13, 0.56); head.add(eye);
    eye.add(mesh(sphere(0.25), white, { scale: [1, 1.08, 0.7] }));
    eye.add(mesh(sphere(0.18), pupil, { pos: [0.02 * s, -0.01, 0.1], scale: [1, 1.08, 0.75] }));
    const s1 = mesh(sphere(0.065), shine, { pos: [0.05 * s + 0.04, 0.07, 0.24] }); s1.castShadow = false;
    const s2 = mesh(sphere(0.03), shine, { pos: [-0.04 * s - 0.02, -0.06, 0.25] }); s2.castShadow = false;
    eye.add(s1, s2);
    eyes.push({ eye, sparkle: [s1, s2] });
  }
  for (const s of [-1, 1]) head.add(mesh(sphere(0.11), blush, { pos: [0.52 * s, -0.17, 0.52], scale: [1, 0.6, 0.4] }));
  const mouth = mesh(new THREE.TorusGeometry(0.13, 0.032, 12, 32, Math.PI), toy(0x7a1c3a), { pos: [0, -0.23, 0.86], rot: [0.25, 0, Math.PI] });
  head.add(mouth);
  for (const s of [-1, 1]) head.add(mesh(sphere(0.03), pupil, { pos: [0.12 * s, -0.07, 0.88] }));

  // rainbow spikes along head + back
  const spikeCols = [0xff3b3b, 0xff9f1c, 0xffd60a, 0x1e90ff, 0x9b5de5];
  const cone = new THREE.ConeGeometry(0.15, 0.36, 24);
  [[0, 0.72, 0.12, 0.05], [0, 0.66, -0.25, -0.45], [0, 0.5, -0.52, -1.0], [0, 1.5, -0.62, -1.3], [0, 1.15, -0.8, -1.8]]
    .forEach(([x, y, z, rx], i) => {
      const m = mesh(cone, toy(spikeCols[i]), { pos: [x, y, z], rot: [rx, 0, 0] });
      (i < 3 ? head : body).add(m);
    });

  // ---- surprise egg (two halves, scaled into an egg shape) ----
  const eggMat = toy(0xffffff, { map: eggTexture(), side: THREE.DoubleSide, roughness: 0.25, clearcoat: 1 });
  const egg = new THREE.Group(); egg.position.y = 1.45; scene.add(egg);
  const eggShape = new THREE.Group(); eggShape.scale.set(1, 1.3, 1); egg.add(eggShape);
  const eggTop = new THREE.Group(), eggBot = new THREE.Group(); eggShape.add(eggTop, eggBot);
  eggTop.add(mesh(new THREE.SphereGeometry(1.08, 64, 32, 0, Math.PI * 2, 0, Math.PI / 2), eggMat));
  eggBot.add(mesh(new THREE.SphereGeometry(1.08, 64, 32, 0, Math.PI * 2, Math.PI / 2, Math.PI / 2), eggMat));

  // ---- confetti ----
  const N = 220;
  const confetti = new THREE.InstancedMesh(new THREE.PlaneGeometry(0.11, 0.06),
    new THREE.MeshStandardMaterial({ side: THREE.DoubleSide, roughness: 0.5 }), N);
  confetti.frustumCulled = false;
  const rainbow = [0xff3b3b, 0xff9f1c, 0xffd60a, 0x2ec84a, 0x1e90ff, 0x9b5de5, 0xff5fb0];
  const R = rng(42), parts = [];
  for (let i = 0; i < N; i++) {
    const a = R() * Math.PI * 2, sp = 3 + R() * 6;
    parts.push({
      v: [Math.cos(a) * sp * 0.55, Math.abs(Math.sin(a)) * sp + 2.5, (R() - 0.4) * 2.5],
      spin: [R() * 14 - 7, R() * 14 - 7, R() * 14 - 7], ph: R() * 6.28,
    });
    confetti.setColorAt(i, new THREE.Color(rainbow[i % rainbow.length]));
  }
  scene.add(confetti);

  // ---- balloons ----
  const balloonDefs = [
    { col: 0xff2d3d, home: [-1.15, 3.0, -0.15] },
    { col: 0x1e7bff, home: [1.15, 3.15, -0.15] },
    { col: 0xffcc00, home: [-0.05, 3.95, -0.7] },
  ];
  const balloons = balloonDefs.map(d => {
    const g = new THREE.Group(); scene.add(g);
    const mat = toy(d.col, { roughness: 0.15, clearcoat: 1, clearcoatRoughness: 0.05 });
    g.add(mesh(sphere(0.42), mat, { scale: [1, 1.18, 1] }));
    g.add(mesh(new THREE.ConeGeometry(0.07, 0.12, 16), mat, { pos: [0, -0.53, 0], rot: [Math.PI, 0, 0] }));
    const hl = mesh(sphere(0.08), new THREE.MeshBasicMaterial({ color: 0xffffff, transparent: true, opacity: 0.75 }), { pos: [-0.16, 0.22, 0.33], scale: [1, 1.6, 0.5] });
    hl.castShadow = false; g.add(hl);
    g.add(mesh(new THREE.CylinderGeometry(0.008, 0.008, 1.3, 6), new THREE.MeshBasicMaterial({ color: 0xffffff }), { pos: [0, -1.2, 0] }));
    return { g, ...d };
  });

  const tmp = new THREE.Object3D();

  // ---------- pose the whole scene at time t ----------
  function update(t) {
    t = ((t % DURATION) + DURATION) % DURATION;

    // --- egg ---
    const wob = smooth(0, 0.5, t) * (1 - smooth(0.62, POP, t));
    egg.rotation.z = Math.sin(t * Math.PI * 2 * 3.4) * 0.2 * wob;
    egg.position.y = 1.45 + Math.abs(Math.sin(t * Math.PI * 2 * 1.7)) * 0.12 * wob;
    const ant = smooth(0.55, POP - 0.02, t) * (t < POP ? 1 : 0); // anticipation squash
    eggShape.scale.set(1 + 0.08 * ant, 1.3 * (1 - 0.1 * ant), 1 + 0.08 * ant);

    let topVis = true, botScale = 1;
    eggTop.position.set(0, 0, 0); eggTop.rotation.set(0, 0, 0);
    if (t >= POP && t < CLOSE) {
      const k = t - POP;
      eggTop.position.set(1.2 * k, 9 * k - 6 * k * k, 0.4 * k);
      eggTop.rotation.set(0, 0, -5 * k);
      topVis = k < 1.2;
      botScale = 1 - easeInBack(k / 0.35);
    } else if (t >= CLOSE) { // egg comes back together around the dino
      const kb = smooth(CLOSE, CLOSE + 0.35, t);
      botScale = easeOutBack(kb, 1.6) * (kb > 0 ? 1 : 0);
      const kt = smooth(CLOSE + 0.25, 9.85, t);
      eggTop.position.set(0, (1 - kt) * 4.5, 0);
      eggTop.rotation.set(0, 0, (1 - kt) * 1.5);
    }
    eggTop.visible = topVis;
    eggBot.visible = botScale > 0.001;
    eggBot.scale.setScalar(Math.max(botScale, 0.0001));
    // settle wobble right as it closes
    if (t > 9.85) egg.rotation.z = Math.sin((t - 9.85) * 40) * 0.05 * (1 - smooth(9.85, 10, t));

    // --- dino scale (pop out / tuck back in) ---
    let ds;
    if (t < POP) ds = 0;
    else if (t < CLOSE + 0.1) ds = easeOutElastic((t - POP) / 0.7);
    else ds = 1 - smooth(CLOSE + 0.1, CLOSE + 0.45, t);
    dino.visible = ds > 0.001;
    dino.scale.setScalar(Math.max(ds, 0.0001));

    // reset pose
    dino.position.set(0, 0, 0); dino.rotation.set(0, 0, 0);
    body.scale.set(1, 1, 1); head.rotation.set(0, 0, 0);
    arms[0].rotation.set(0, 0, 0); arms[1].rotation.set(0, 0, 0);
    legs[0].rotation.set(0, 0, 0); legs[1].rotation.set(0, 0, 0);
    tail.rotation.set(0, Math.sin(t * 6) * 0.25, 0);
    let eyeOpen = 1;

    // pop-out hop
    if (t >= POP && t < 1.6) {
      const k = (t - POP) / 0.85;
      dino.position.y = Math.max(0, Math.sin(Math.PI * k)) * 0.7;
      arms[0].rotation.z = -1.6 * Math.sin(Math.PI * k); arms[1].rotation.z = 1.6 * Math.sin(Math.PI * k);
    }
    // giggle: happy squinty eyes, bouncy shoulders, head tilt
    if (t >= 1.15 && t < 1.95) {
      const g = smooth(1.15, 1.25, t) * (1 - smooth(1.8, 1.95, t));
      const b = Math.sin((t - 1.15) * Math.PI * 2 * 7);
      body.scale.set(1 + 0.03 * b * g, 1 - 0.04 * b * g, 1);
      head.rotation.z = 0.18 * g;
      head.rotation.x = -0.08 * g;
      eyeOpen = 1 - 0.45 * g;
      arms[0].rotation.z = 0.5 * g; arms[1].rotation.z = -0.5 * g;
    }
    // look at camera with a cute head tilt before dancing
    if (t >= 1.95 && t < 2.0) head.rotation.z = 0;

    // --- dance: lean L, lean R, hop (boing!), land + twirl — x3 ---
    for (const T0 of DANCE) {
      if (t < T0 || t >= T0 + 2) continue;
      const u = (t - T0) / 0.5; // beats 0..4
      if (u < 2) {
        const s = Math.sin(Math.PI * u);
        dino.rotation.z = 0.22 * s;
        dino.position.x = -0.18 * s;
        head.rotation.z = 0.15 * s;
        if (s > 0) arms[0].rotation.z = -1.8 * s; else arms[1].rotation.z = -1.8 * s;
        legs[0].rotation.x = 0.3 * Math.max(0, -s); legs[1].rotation.x = 0.3 * Math.max(0, s);
        body.scale.y = 1 - 0.05 * Math.abs(Math.sin(Math.PI * 2 * u));
      } else if (u < 3) {
        const k = u - 2;
        const crouch = 1 - smooth(0, 0.18, k);
        dino.position.y = Math.sin(Math.PI * smooth(0.12, 0.95, k)) * 1.0;
        const air = Math.sin(Math.PI * clamp((k - 0.12) / 0.83));
        body.scale.set(1 + 0.12 * crouch - 0.05 * air, 1 - 0.18 * crouch + 0.1 * air, 1);
        arms[0].rotation.z = -1.85 * air; arms[1].rotation.z = 1.85 * air;
        legs[0].rotation.x = -0.4 * air; legs[1].rotation.x = -0.4 * air;
      } else {
        const k = u - 3;
        const land = Math.exp(-k * 7) * Math.cos(k * 22);
        body.scale.set(1 + 0.14 * land, 1 - 0.2 * land, 1);
        dino.rotation.y = Math.sin(Math.PI * 2 * smooth(0.1, 0.95, k)) * 0.55;
        arms[0].rotation.z = -1.2 * Math.sin(Math.PI * k); arms[1].rotation.z = 1.2 * Math.sin(Math.PI * k);
        head.rotation.x = -0.1 * Math.sin(Math.PI * k);
      }
      tail.rotation.y = Math.sin(t * 12) * 0.45;
    }

    // --- goodbye wave ---
    if (t >= FLOAT && t < CLOSE + 0.1) {
      const w = smooth(FLOAT, FLOAT + 0.3, t) * (1 - smooth(CLOSE - 0.15, CLOSE + 0.1, t));
      arms[1].rotation.z = 1.75 * w + Math.sin((t - FLOAT) * Math.PI * 2 * 2.4) * 0.4 * w;
      arms[0].rotation.z = -0.25 * w;
      dino.rotation.z = Math.sin((t - FLOAT) * Math.PI * 2) * 0.08 * w;
      head.rotation.z = -0.12 * w;
      head.rotation.x = -0.1 * w; // gaze up a little toward the balloons
      body.scale.y = 1 + Math.sin((t - FLOAT) * Math.PI * 4) * 0.02 * w;
    }

    // blinks
    for (const b of [2.75, 6.75, 8.9]) {
      const d = Math.abs(t - b);
      if (d < 0.08) eyeOpen = Math.min(eyeOpen, 0.1 + 0.9 * (d / 0.08));
    }
    for (const { eye, sparkle } of eyes) {
      eye.scale.set(1, eyeOpen, 1);
      sparkle[0].scale.setScalar(1 + 0.2 * Math.sin(t * 9));
      sparkle[1].scale.setScalar(1 + 0.3 * Math.sin(t * 11 + 1));
    }
    mouth.scale.set(1, t >= 1.15 && t < 1.95 ? 1.4 : 1, 1);

    // --- confetti ---
    const kc = t - POP;
    const kd = 2.2, gr = 3.2; // drag + gravity -> slow fluttery fall
    for (let i = 0; i < N; i++) {
      const p = parts[i];
      if (kc < 0 || kc > 3.6) { tmp.scale.setScalar(0); }
      else {
        const e = (1 - Math.exp(-kd * kc)) / kd;
        tmp.position.set(
          p.v[0] * e + 0.12 * Math.sin(kc * 4 + p.ph),
          1.5 + p.v[1] * e - (gr / kd) * (kc - e),
          p.v[2] * e);
        tmp.rotation.set(p.spin[0] * kc + p.ph, p.spin[1] * kc, p.spin[2] * kc);
        tmp.scale.setScalar(smooth(0, 0.05, kc) * (1 - smooth(2.9, 3.6, kc)));
      }
      tmp.updateMatrix();
      confetti.setMatrixAt(i, tmp.matrix);
    }
    confetti.instanceMatrix.needsUpdate = true;

    // --- balloons ---
    balloons.forEach((b, i) => {
      const k = t - BALLOON_AT[i];
      if (k < 0) { b.g.visible = false; return; }
      const fk = t - (FLOAT + 0.12 * i);
      const rise = fk > 0 ? 0.5 * 5 * fk * fk + 0.3 * fk : 0;
      b.g.visible = rise < 8;
      b.g.position.set(
        b.home[0] + Math.sin(t * 1.6 + i * 2) * 0.06 + (fk > 0 ? Math.sin(fk * 3 + i) * 0.15 * fk : 0),
        b.home[1] + Math.sin(t * 2.1 + i) * 0.07 + rise,
        b.home[2]);
      b.g.rotation.z = Math.sin(t * 1.8 + i) * 0.08;
      b.g.scale.setScalar(Math.max(0.0001, easeOutElastic(k / 0.6)));
    });

    // --- camera: centered, gentle zoom-ins on the face ---
    const z1 = smooth(1.0, 1.5, t) * (1 - smooth(1.75, 2.25, t));
    const z2 = 0.7 * smooth(8.35, 8.95, t) * (1 - smooth(9.05, 9.7, t));
    const z = Math.max(z1, z2);
    const faceY = 1.85 + dino.position.y * 0.5;
    camera.position.set(0, lerp(2.15, faceY + 0.15, z), lerp(9.2, 5.4, z));
    camera.lookAt(0, lerp(2.15, faceY, z), 0);
  }

  function render(t) { update(t); renderer.render(scene, camera); }
  function resize(w, h) { renderer.setSize(w, h, false); camera.aspect = w / h; camera.updateProjectionMatrix(); }
  return { render, resize, renderer };
}
