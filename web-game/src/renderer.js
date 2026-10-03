// 描画。VB6 の Timer1_Timer の描画部分（BitBlt / DrawText）を Canvas 2D に置き換えたもの。
// 論理座標は元のフォームと同じ 600x600 ピクセル。

import {
  ENTRANCE, DUNGEON, HOWTO_PLAY, OPTIONS, GAME_OVER, GAME_CLEAR, MUSEUM, BLESSING,
  PURPLE_BOX, STAIR, WALL, ROOM, ENEMY, LAND_NUMBER, SC, SCREEN, FONT_SIZE, LINE_HEIGHT,
} from './constants.js';
import { blessingInfo } from './engine.js';
import { createEffects } from './effects.js';

const FONT_PX = 19; // 14.25pt
const FONT_FAMILY = '"Meiryo", "メイリオ", "Hiragino Kaku Gothic ProN", "Noto Sans JP", sans-serif';

// フォームと同じ矩形（LandSet / FloorSet の MessageBar 等）
const RECT_SPEC = { left: LAND_NUMBER * SC, right: SCREEN, top: 0, bottom: LAND_NUMBER * SC };
const RECT_STATUS = { left: 0, right: SCREEN, top: LAND_NUMBER * SC, bottom: SCREEN };
const RECT_MESSAGE = { left: 0, right: SCREEN, top: LAND_NUMBER * SC + FONT_SIZE * 2, bottom: SCREEN };
const MAP = LAND_NUMBER * SC;
const GOLD = '#ffd84a';
const HP_BAR = { left: SCREEN - 400, top: LAND_NUMBER * SC + FONT_SIZE * 1.5, width: 400, height: 5 };

export function createRenderer(canvas, sheets) {
  const ctx = canvas.getContext('2d');
  let scale = 1;
  const fx = createEffects();

  function resize() {
    const dpr = window.devicePixelRatio || 1;
    const css = canvas.getBoundingClientRect().width || SCREEN;
    scale = (css / SCREEN) * dpr;
    const px = Math.round(SCREEN * scale);
    if (canvas.width !== px) {
      canvas.width = px;
      canvas.height = px;
    }
  }

  function sprite(sheet, row, left, top) {
    ctx.drawImage(sheets[sheet], 0, row * SC, SC, SC, left, top, SC, SC);
  }

  // DrawText 相当。text は '\n' 区切り、align は 'center' | 'left'
  function drawText(text, rect, align, color) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top);
    ctx.clip();
    ctx.fillStyle = color;
    ctx.textBaseline = 'top';
    ctx.textAlign = align;
    const width = rect.right - rect.left;
    const x = align === 'center' ? (rect.left + rect.right) / 2 : rect.left;
    text.split('\n').forEach((line, k) => {
      if (!line) return;
      ctx.font = `bold ${FONT_PX}px ${FONT_FAMILY}`;
      const w = ctx.measureText(line).width;
      if (w > width) {
        // 画面幅に収まらない行だけ縮める
        ctx.font = `bold ${Math.floor((FONT_PX * width) / w)}px ${FONT_FAMILY}`;
      }
      ctx.fillText(line, x, rect.top + k * LINE_HEIGHT);
    });
    ctx.restore();
  }

  const join = (words, from, to) => words.slice(from, to + 1).join('\n');

  function drawHpBar(p) {
    const { left, top, width, height } = HP_BAR;
    ctx.fillStyle = '#000';
    ctx.fillRect(left, top, width, height);
    const hp = Math.max(0, p.hp);
    if (p.maxHp <= width - 1) {
      ctx.fillStyle = 'rgb(34,177,76)';
      ctx.fillRect(left, top, Math.min(hp, width), height);
      ctx.fillStyle = '#ff0000';
      ctx.fillRect(left + hp, top, Math.max(0, p.maxHp - hp), height);
    } else {
      const g = Math.floor((hp / p.maxHp) * width) + 1;
      ctx.fillStyle = 'rgb(34,177,76)';
      ctx.fillRect(left, top, Math.min(g, width), height);
      ctx.fillStyle = '#ff0000';
      ctx.fillRect(left + g, top, Math.max(0, width - g), height);
    }
  }

  // 削れている箱・壁・モンスターの残り Hp（マスの下端に細い線で出す）
  function damageBar(x, y, hp, maxHp, color) {
    if (hp <= 0 || hp >= maxHp) return;
    ctx.fillStyle = '#000';
    ctx.fillRect(x * SC + 1, y * SC + SC - 2, SC - 2, 2);
    ctx.fillStyle = color;
    ctx.fillRect(x * SC + 1, y * SC + SC - 2, Math.max(1, Math.round((SC - 2) * (hp / maxHp))), 2);
  }

  // 右下の小さなステータス（ブラウザ版独自）
  function drawInfo(g) {
    const p = g.player;
    const m = g.monsters[0];
    const lines = [
      `攻 ${p.atk}  守 ${p.def}`,
      `次のLvまで ${Math.max(0, p.level ** 3 - p.exp)}`,
      `敵Lv${m.level} 攻${m.atk} 守${m.def}`,
    ];
    ctx.fillStyle = '#b8b8b8';
    ctx.textBaseline = 'top';
    ctx.textAlign = 'left';
    lines.forEach((line, k) => {
      ctx.font = `12px ${FONT_FAMILY}`;
      const w = ctx.measureText(line).width;
      const max = SCREEN - MAP - 4;
      if (w > max) ctx.font = `${Math.floor((12 * max) / w)}px ${FONT_FAMILY}`;
      ctx.fillText(line, MAP, 348 + k * 17);
    });
  }

  function drawDungeon(g) {
    const frame = Math.floor((g.floor % 200) / 20); // 20 階ごとに床と壁の絵が変わる
    const [sx, sy] = fx.shakeOffset();
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, 0, MAP, MAP);
    ctx.clip();
    ctx.translate(sx, sy);
    for (let x = 0; x < LAND_NUMBER; x++) {
      for (let y = 0; y < LAND_NUMBER; y++) {
        const t = g.land[x][y];
        const c = t.condition;
        if (c === WALL) sprite('wall', frame, x * SC, y * SC);
        else if (c === ROOM || c === ENEMY) sprite('room', frame, x * SC, y * SC);
        else if (c <= STAIR) sprite('box', c, x * SC, y * SC);
        if (c <= PURPLE_BOX || c === WALL) damageBar(x, y, t.hp, t.maxHp, GOLD);
      }
    }
    const p = g.player;
    if (p.alive) sprite('player', p.condition, p.x * SC, p.y * SC);
    for (let i = 0; i < g.monsterNumber; i++) {
      const m = g.monsters[i];
      if (!m.alive) continue;
      sprite('monster', m.ability, m.x * SC, m.y * SC);
      damageBar(m.x, m.y, m.hp, m.maxHp, '#ff5a4a');
    }
    fx.drawWorld(ctx);
    ctx.translate(-sx, -sy);
    fx.drawOverlay(ctx, g);
    ctx.restore();
    drawInfo(g);
    drawHpBar(p);
    const w = g.words;
    drawText(w[9], RECT_STATUS, 'center', g.foreColor);
    drawText(join(w, 0, 3), RECT_MESSAGE, 'center', g.foreColor);
    drawText(join(w, 4, 8), RECT_SPEC, 'left', g.foreColor);
  }

  // 祝福の選択画面（アレンジルール）。ダンジョンを暗くした上に 3 枚のカードを並べる
  function drawBlessing(g) {
    const b = g.blessing;
    ctx.fillStyle = 'rgba(0,0,0,0.8)';
    ctx.fillRect(0, 0, SCREEN, SCREEN);
    const full = { left: 0, right: SCREEN, top: 0, bottom: SCREEN };
    drawText(`${b.floor}F 突破！`, { ...full, top: 70 }, 'center', GOLD);
    drawText('祝福を1つ選んでください', { ...full, top: 110 }, 'center', '#ffffff');
    const w = 180;
    const h = 170;
    b.options.forEach((id, i) => {
      const info = blessingInfo(id);
      const selected = i === b.cursor;
      const left = 20 + i * (w + 10);
      const top = selected ? 180 : 190;
      ctx.fillStyle = selected ? '#3a3210' : '#1c1c1c';
      ctx.fillRect(left, top, w, h);
      ctx.strokeStyle = selected ? GOLD : '#666666';
      ctx.lineWidth = selected ? 3 : 1;
      ctx.strokeRect(left + 0.5, top + 0.5, w - 1, h - 1);
      const rect = { left: left + 6, right: left + w - 6, top: top + 20, bottom: top + h };
      drawText(info.name, rect, 'center', selected ? GOLD : '#cccccc');
      drawText(info.desc, { ...rect, top: top + 75 }, 'center', selected ? '#ffffff' : '#999999');
    });
    drawText('左右キーで選択、Enterキーで決定', { ...full, top: 400 }, 'center', b.lock > 0 ? '#666666' : '#ffffff');
  }

  function render(g, events = [], dt = 0) {
    fx.update(events, g, dt);
    resize();
    ctx.setTransform(scale, 0, 0, scale, 0, 0);
    ctx.imageSmoothingEnabled = false;
    ctx.fillStyle = g.backColor;
    ctx.fillRect(0, 0, SCREEN, SCREEN);

    const w = g.words;
    const top = FONT_SIZE * 2;
    switch (g.mode) {
      case ENTRANCE: {
        const bottom = FONT_SIZE * 32;
        drawText(join(w, 0, 4), { left: 0, right: SCREEN, top, bottom }, 'center', g.foreColor);
        drawText(join(w, 5, 8), { left: 0, right: SCREEN, top: bottom, bottom: SCREEN }, 'center', g.foreColor);
        break;
      }
      case OPTIONS:
      case GAME_OVER:
      case GAME_CLEAR:
      case MUSEUM:
        drawText(join(w, 0, 8), { left: 0, right: SCREEN, top, bottom: SCREEN }, 'center', g.foreColor);
        break;
      case HOWTO_PLAY: {
        const bottom = FONT_SIZE * 22;
        drawText(join(w, 0, 4), { left: 0, right: SCREEN, top, bottom }, 'center', g.foreColor);
        drawText(join(w, 5, 8), { left: 0, right: SCREEN, top: bottom, bottom: SCREEN }, 'center', g.foreColor);
        for (const d of g.demo) sprite(d.sheet, d.row, d.left, d.top);
        break;
      }
      case DUNGEON:
        drawDungeon(g);
        break;
      case BLESSING:
        drawDungeon(g);
        drawBlessing(g);
        break;
    }
  }

  return { render };
}
