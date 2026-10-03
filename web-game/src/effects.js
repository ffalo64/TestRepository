// 画面の演出（ダメージ数字・破片・揺れ・箱の効果のバナーなど）。ブラウザ版独自。
// エンジンが積んだ g.events を読むだけで、ゲームの進行には関与しない。

import { LAND_NUMBER, SC, PURPLE_BOX } from './constants.js';

const MAP = LAND_NUMBER * SC;
const BOX_COLORS = ['#4aa3ff', '#ff5a4a', '#ffd84a', '#5ad66a', '#c07aff'];
const WALL_COLOR = '#9a9a9a';
const GOLD = '#ffd84a';
const MAX_POPUPS = 40;
const MAX_PARTICLES = 300;
const FONT_FAMILY = '"Meiryo", "メイリオ", "Hiragino Kaku Gothic ProN", "Noto Sans JP", sans-serif';

const center = (v) => v * SC + SC / 2;

export function createEffects() {
  let popups = [];
  let particles = [];
  let rings = [];
  let shake = 0;
  let hurt = 0; // 被弾時の赤い縁取りの濃さ
  let flash = null; // 画面全体の閃光 { color, life, max, alpha }
  let banner = null; // 箱の効果 { text, color, life }
  let intro = null; // 階の表示 { floor, life }
  let clock = 0;

  function popup(x, y, text, color, size = 11, life = 700) {
    if (popups.length >= MAX_POPUPS) popups.shift();
    // 同じマスに続けて出るときは少しずらして重ならないようにする
    const dx = ((popups.length % 3) - 1) * 4;
    popups.push({ x: x + dx, y, text, color, size, life, max: life });
  }

  function burst(x, y, color, count, speed) {
    for (let k = 0; k < count && particles.length < MAX_PARTICLES; k++) {
      const a = Math.random() * Math.PI * 2;
      const v = speed * (0.4 + Math.random() * 0.6);
      const life = 250 + Math.random() * 300;
      particles.push({
        x, y, vx: Math.cos(a) * v, vy: Math.sin(a) * v - speed * 0.3,
        color, life, max: life, size: 1 + Math.random() * 2,
      });
    }
  }

  function ring(x, y, color, radius, life) {
    rings.push({ x, y, color, radius, life, max: life });
  }

  function setFlash(color, alpha, life) {
    flash = { color, alpha, life, max: life };
  }

  function handle(e, g) {
    const px = center(g.player.x);
    const py = center(g.player.y);
    switch (e.t) {
      case 'hit':
        popup(center(e.x), e.y * SC, String(e.dmg), e.tile ? '#ffffff' : '#ffe066');
        burst(center(e.x), center(e.y), '#ffffff', 3, 0.05);
        break;
      case 'kill':
        burst(center(e.x), center(e.y), '#9be8ff', 12, 0.09);
        break;
      case 'box': {
        const isBox = e.box <= PURPLE_BOX;
        const color = isBox ? BOX_COLORS[e.box] : WALL_COLOR;
        burst(center(e.x), center(e.y), color, isBox ? 18 : 8, 0.11);
        if (isBox) {
          ring(center(e.x), center(e.y), color, 26, 350);
          if (e.text) banner = { text: e.text, color, life: 2400 };
        }
        break;
      }
      case 'hurt':
        popup(px, py - SC, `-${e.dmg}`, '#ff5a4a', 12);
        shake = Math.min(7, 2 + e.ratio * 14);
        hurt = Math.min(0.7, 0.25 + e.ratio);
        break;
      case 'level':
        popup(px, py - SC * 1.5, 'LEVEL UP!', GOLD, 12, 1100);
        ring(px, py, GOLD, 30, 500);
        break;
      case 'floor':
        // 階が変わると座標の意味が変わるので、前の階の演出は消す
        popups = [];
        particles = [];
        rings = [];
        intro = { floor: e.floor, life: 900 };
        break;
      case 'cmd':
        if (e.key === 'Z') { ring(px, py, '#ffffff', 320, 450); setFlash('#ffffff', 0.25, 200); }
        else if (e.key === 'X') { ring(px, py, '#5ad66a', 40, 500); popup(px, py - SC, 'Hp全快', '#5ad66a', 12, 900); }
        else if (e.key === 'C') { setFlash('#ffffff', 0.85, 500); ring(px, py, '#ffffff', 420, 500); }
        else if (e.key === 'D') { setFlash('#c07aff', 0.4, 350); }
        break;
      case 'revive':
        setFlash(GOLD, 0.7, 700);
        ring(px, py, GOLD, 120, 700);
        popup(px, py - SC * 1.5, '復活！', GOLD, 14, 1400);
        break;
      case 'blessed':
        setFlash(GOLD, 0.35, 500);
        banner = { text: `${e.name}を受けた`, color: GOLD, life: 2400 };
        break;
    }
  }

  return {
    update(events, g, dt) {
      for (const e of events) handle(e, g);
      clock += dt;
      for (const p of popups) p.life -= dt;
      for (const p of particles) {
        p.life -= dt;
        p.x += p.vx * dt;
        p.y += p.vy * dt;
        p.vy += 0.0004 * dt;
      }
      for (const r of rings) r.life -= dt;
      popups = popups.filter((p) => p.life > 0);
      particles = particles.filter((p) => p.life > 0);
      rings = rings.filter((r) => r.life > 0);
      shake = Math.max(0, shake - dt * 0.03);
      hurt = Math.max(0, hurt - dt * 0.002);
      if (flash && (flash.life -= dt) <= 0) flash = null;
      if (banner && (banner.life -= dt) <= 0) banner = null;
      if (intro && (intro.life -= dt) <= 0) intro = null;
    },

    // マップを揺らす量（ピクセル）
    shakeOffset() {
      if (shake <= 0) return [0, 0];
      return [(Math.random() * 2 - 1) * shake, (Math.random() * 2 - 1) * shake];
    },

    // マップと同じ座標系（揺れの内側）に描くもの
    drawWorld(ctx) {
      for (const r of rings) {
        const k = 1 - r.life / r.max;
        ctx.globalAlpha = 1 - k;
        ctx.strokeStyle = r.color;
        ctx.lineWidth = 2;
        ctx.beginPath();
        ctx.arc(r.x, r.y, r.radius * k, 0, Math.PI * 2);
        ctx.stroke();
      }
      for (const p of particles) {
        ctx.globalAlpha = Math.min(1, (p.life / p.max) * 2);
        ctx.fillStyle = p.color;
        ctx.fillRect(p.x - p.size / 2, p.y - p.size / 2, p.size, p.size);
      }
      ctx.textAlign = 'center';
      ctx.textBaseline = 'bottom';
      ctx.lineJoin = 'round';
      for (const p of popups) {
        const k = 1 - p.life / p.max;
        ctx.globalAlpha = Math.min(1, (p.life / p.max) * 3);
        ctx.font = `bold ${p.size}px ${FONT_FAMILY}`;
        const x = Math.max(12, Math.min(MAP - 12, p.x));
        const y = Math.max(12, p.y - k * 16);
        ctx.lineWidth = 3;
        ctx.strokeStyle = '#000000';
        ctx.strokeText(p.text, x, y);
        ctx.fillStyle = p.color;
        ctx.fillText(p.text, x, y);
      }
      ctx.globalAlpha = 1;
    },

    // マップの上に重ねるもの（揺れの外側）
    drawOverlay(ctx, g) {
      if (flash) {
        ctx.globalAlpha = flash.alpha * (flash.life / flash.max);
        ctx.fillStyle = flash.color;
        ctx.fillRect(0, 0, MAP, MAP);
      }
      // 被弾と、残りターンが少ないときの警告は赤い縁取りで知らせる
      const lowTurn = g.turn <= 30 ? 0.25 + 0.2 * Math.sin(clock / 150) : 0;
      const edge = Math.max(hurt, lowTurn);
      if (edge > 0) {
        const grad = ctx.createRadialGradient(MAP / 2, MAP / 2, MAP * 0.35, MAP / 2, MAP / 2, MAP * 0.75);
        grad.addColorStop(0, 'rgba(255,0,0,0)');
        grad.addColorStop(1, 'rgba(255,0,0,1)');
        ctx.globalAlpha = edge;
        ctx.fillStyle = grad;
        ctx.fillRect(0, 0, MAP, MAP);
      }
      if (intro) {
        ctx.globalAlpha = Math.min(1, intro.life / 400) * 0.9;
        ctx.font = `bold 64px ${FONT_FAMILY}`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.lineWidth = 6;
        ctx.strokeStyle = '#000000';
        ctx.strokeText(`${intro.floor}F`, MAP / 2, MAP / 2);
        ctx.fillStyle = '#ffffff';
        ctx.fillText(`${intro.floor}F`, MAP / 2, MAP / 2);
      }
      if (banner) {
        const alpha = Math.min(1, banner.life / 400, (2400 - banner.life) / 120 + 0.2);
        const w = MAP - 16;
        const h = 24;
        // プレイヤーが上の方に居るときは下に出して、足元を隠さない
        const y = g.player.y < 6 ? MAP - h - 8 : 8;
        ctx.globalAlpha = alpha * 0.85;
        ctx.fillStyle = '#000000';
        ctx.fillRect(8, y, w, h);
        ctx.globalAlpha = alpha;
        ctx.fillStyle = banner.color;
        ctx.fillRect(8, y, 5, h);
        ctx.fillRect(8 + w - 5, y, 5, h);
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        let size = 14;
        ctx.font = `bold ${size}px ${FONT_FAMILY}`;
        const tw = ctx.measureText(banner.text).width;
        if (tw > w - 24) {
          size = Math.floor((size * (w - 24)) / tw);
          ctx.font = `bold ${size}px ${FONT_FAMILY}`;
        }
        ctx.fillStyle = '#ffffff';
        ctx.fillText(banner.text, 8 + w / 2, y + h / 2 + 1);
      }
      ctx.globalAlpha = 1;
    },
  };
}
