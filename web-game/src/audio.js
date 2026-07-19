const TRACKS = [
  '/sounds/n31.mp3',   // 0: menu (loop)
  '/sounds/n11.mp3',   // 1: game clear (once)
  '/sounds/c26.mp3',   // 2: game over (once)
  '/sounds/c1.mp3',    // 3: dungeon (loop)
];

const audioElements = [];
let currentTrack = -1;

export function initAudio() {
  for (let i = 0; i < TRACKS.length; i++) {
    const a = new Audio(TRACKS[i]);
    a.loop = (i === 0 || i === 3);
    audioElements[i] = a;
  }
}

// ── DLC: synthesized sound effects (WebAudio, no asset files) ─────────────────

let actx = null;

// Call once on first keydown to satisfy autoplay policy
export function onUserInteraction() {
  if (typeof window === 'undefined') return;
  const AC = window.AudioContext || window.webkitAudioContext;
  if (!AC) return;
  if (!actx) actx = new AC();
  if (actx.state === 'suspended') actx.resume();
}

function tone(c, { type = 'square', f0 = 440, f1 = f0, t = 0.1, g = 0.05, delay = 0 }) {
  const o = c.createOscillator(), gn = c.createGain();
  const t0 = c.currentTime + delay;
  o.type = type;
  o.frequency.setValueAtTime(f0, t0);
  if (f1 !== f0) o.frequency.exponentialRampToValueAtTime(Math.max(1, f1), t0 + t);
  gn.gain.setValueAtTime(g, t0);
  gn.gain.exponentialRampToValueAtTime(0.0005, t0 + t);
  o.connect(gn).connect(c.destination);
  o.start(t0); o.stop(t0 + t + 0.02);
}

// per-sound minimum interval (ms) so the 20 ticks/s game loop can't spam
const THROTTLE = { hit: 130, break: 150, pickup: 150, bad: 300 };
const _last = {};

export function sfx(name) {
  if (typeof window === 'undefined' || !actx || actx.state !== 'running') return;
  const now = performance.now();
  if (_last[name] && now - _last[name] < (THROTTLE[name] || 80)) return;
  _last[name] = now;
  const c = actx;
  switch (name) {
    case 'hit':      tone(c, { type: 'square',   f0: 220,  f1: 120,  t: 0.06, g: 0.03 }); break;
    case 'break':    tone(c, { type: 'triangle', f0: 520,  f1: 70,   t: 0.12, g: 0.05 }); break;
    case 'pickup':   tone(c, { type: 'sine',     f0: 660,  f1: 1040, t: 0.12, g: 0.05 }); break;
    case 'bad':      tone(c, { type: 'sawtooth', f0: 180,  f1: 60,   t: 0.25, g: 0.05 }); break;
    case 'stair':    tone(c, { type: 'sine',     f0: 392,  f1: 196,  t: 0.18, g: 0.04 }); break;
    case 'death':    tone(c, { type: 'sawtooth', f0: 220,  f1: 35,   t: 0.8,  g: 0.07 }); break;
    case 'levelup':
      tone(c, { type: 'sine', f0: 523, t: 0.09, g: 0.05 });
      tone(c, { type: 'sine', f0: 659, t: 0.09, g: 0.05, delay: 0.08 });
      tone(c, { type: 'sine', f0: 784, t: 0.14, g: 0.05, delay: 0.16 });
      break;
    case 'boss':
      tone(c, { type: 'sawtooth', f0: 70,  f1: 55, t: 0.6, g: 0.07 });
      tone(c, { type: 'sawtooth', f0: 105, f1: 80, t: 0.6, g: 0.05 });
      break;
    case 'bossdown':
      [523, 659, 784, 1046].forEach((f, i) => tone(c, { type: 'square', f0: f, t: 0.12, g: 0.05, delay: i * 0.09 }));
      break;
  }
}

// ── music ─────────────────────────────────────────────────────────────────────

export function playMusic(idx) {
  if (currentTrack === idx) {
    const a = audioElements[idx];
    if (a && !a.paused) return; // already playing
  }
  stopAll();
  currentTrack = idx;
  const a = audioElements[idx];
  if (a) { a.currentTime = 0; a.play().catch(() => {}); }
}

export function stopAll() {
  for (const a of audioElements) {
    if (a) { a.pause(); a.currentTime = 0; }
  }
  currentTrack = -1;
}
