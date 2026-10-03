// 効果音。ブラウザ版独自（VB6 版は BGM のみ）。素材を使わず Web Audio で合成する。
// エンジンが積んだ g.events を読んで鳴らすだけで、ゲームの進行には関与しない。

import { BLUE_BOX, RED_BOX, YELLOW_BOX, GREEN_BOX, PURPLE_BOX } from './constants.js';

const MIN_INTERVAL = 0.05; // 同じ音を続けて鳴らす最短間隔（秒）。オートプレイの倍速で音が重なりすぎないように

export function createSfx() {
  let ctx = null;
  let master = null;
  let muted = false;
  const lastPlayed = {};

  function tone(freq, dur, { type = 'square', vol = 0.3, to = freq, delay = 0 } = {}) {
    const t = ctx.currentTime + delay;
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.type = type;
    osc.frequency.setValueAtTime(freq, t);
    if (to !== freq) osc.frequency.exponentialRampToValueAtTime(to, t + dur);
    gain.gain.setValueAtTime(vol, t);
    gain.gain.exponentialRampToValueAtTime(0.001, t + dur);
    osc.connect(gain).connect(master);
    osc.start(t);
    osc.stop(t + dur + 0.02);
  }

  function noise(dur, { vol = 0.3, cutoff = 2000, to = cutoff, delay = 0 } = {}) {
    const t = ctx.currentTime + delay;
    const len = Math.max(1, Math.floor(ctx.sampleRate * dur));
    const buffer = ctx.createBuffer(1, len, ctx.sampleRate);
    const data = buffer.getChannelData(0);
    for (let i = 0; i < len; i++) data[i] = Math.random() * 2 - 1;
    const src = ctx.createBufferSource();
    src.buffer = buffer;
    const filter = ctx.createBiquadFilter();
    filter.type = 'lowpass';
    filter.frequency.setValueAtTime(cutoff, t);
    if (to !== cutoff) filter.frequency.exponentialRampToValueAtTime(to, t + dur);
    const gain = ctx.createGain();
    gain.gain.setValueAtTime(vol, t);
    gain.gain.exponentialRampToValueAtTime(0.001, t + dur);
    src.connect(filter).connect(gain).connect(master);
    src.start(t);
  }

  const arpeggio = (notes, step, dur, opts) =>
    notes.forEach((f, i) => tone(f, dur, { ...opts, delay: i * step }));

  const BOX_SOUNDS = {
    [BLUE_BOX]: () => arpeggio([660, 880], 0.07, 0.18, { type: 'sine', vol: 0.3 }),
    [RED_BOX]: () => { tone(220, 0.3, { type: 'sawtooth', vol: 0.25, to: 70 }); noise(0.2, { vol: 0.15, cutoff: 600 }); },
    [YELLOW_BOX]: () => arpeggio([700, 1050, 880], 0.05, 0.1, { type: 'square', vol: 0.18 }),
    [GREEN_BOX]: () => arpeggio([523, 659, 784], 0.06, 0.14, { type: 'triangle', vol: 0.35 }),
    [PURPLE_BOX]: () => arpeggio([880, 1175, 1568, 2093], 0.05, 0.16, { type: 'sine', vol: 0.25 }),
  };

  const SOUNDS = {
    hit: (e) => (e.tile
      ? tone(150, 0.06, { type: 'triangle', vol: 0.35, to: 90 })
      : tone(260, 0.06, { type: 'square', vol: 0.15, to: 150 })),
    kill: () => tone(500, 0.1, { type: 'square', vol: 0.15, to: 950 }),
    box: (e) => (BOX_SOUNDS[e.box] ?? (() => noise(0.1, { vol: 0.25, cutoff: 900 })))(),
    hurt: () => { noise(0.14, { vol: 0.35, cutoff: 500 }); tone(110, 0.14, { type: 'sine', vol: 0.4, to: 60 }); },
    level: () => arpeggio([523, 659, 784, 1047], 0.07, 0.16, { type: 'square', vol: 0.16 }),
    floor: () => arpeggio([392, 330, 262], 0.06, 0.1, { type: 'triangle', vol: 0.3 }),
    cmd: (e) => {
      if (e.key === 'Z') { noise(0.35, { vol: 0.35, cutoff: 3000, to: 200 }); tone(120, 0.35, { type: 'sawtooth', vol: 0.2, to: 40 }); }
      else if (e.key === 'X') tone(440, 0.35, { type: 'sine', vol: 0.3, to: 1320 });
      else if (e.key === 'C') { noise(0.6, { vol: 0.4, cutoff: 200, to: 6000 }); tone(80, 0.6, { type: 'sine', vol: 0.3, to: 30 }); }
      else if (e.key === 'D') arpeggio([330, 247, 415], 0.06, 0.12, { type: 'square', vol: 0.18 });
    },
    revive: () => arpeggio([392, 523, 659, 784, 1047, 1319], 0.08, 0.3, { type: 'sine', vol: 0.3 }),
    bless: () => arpeggio([523, 784, 1047], 0.12, 0.5, { type: 'sine', vol: 0.25 }),
    blessed: () => arpeggio([659, 784, 1047, 1319], 0.06, 0.3, { type: 'triangle', vol: 0.3 }),
    select: () => tone(880, 0.04, { type: 'square', vol: 0.12 }),
  };

  return {
    // ブラウザの自動再生制限: 最初のキー入力/クリックで呼ぶ
    unlock() {
      if (!ctx) {
        const AC = window.AudioContext || window.webkitAudioContext;
        if (!AC) return;
        ctx = new AC();
        master = ctx.createGain();
        master.gain.value = 0.35;
        master.connect(ctx.destination);
      }
      if (ctx.state === 'suspended') ctx.resume();
    },
    setMuted(m) {
      muted = m;
    },
    play(events) {
      if (!ctx || muted || ctx.state !== 'running') return;
      for (const e of events) {
        const sound = SOUNDS[e.t];
        if (!sound) continue;
        const key = e.t === 'box' ? `box${e.box}` : e.t === 'cmd' ? `cmd${e.key}` : e.t;
        if (ctx.currentTime - (lastPlayed[key] ?? -1) < MIN_INTERVAL) continue;
        lastPlayed[key] = ctx.currentTime;
        sound(e);
      }
    },
  };
}
