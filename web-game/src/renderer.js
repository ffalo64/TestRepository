import { SC, LAND_NUMBER, SPEC_X, STATUS_Y, MSG_Y, CANVAS_W, CANVAS_H, GameMode, T, A } from './constants.js';
import { state } from './state.js';
import { BANDS, bandIndex } from './lore.js';

const BOX_COLORS = ['#4488ff', '#ff4444', '#ffff00', '#44cc44', '#aa44ff'];
const TILE_COLORS = { wall: '#555555', room: '#222222', stair: '#ffffff', enemy: '#222222' };

// ── main draw ─────────────────────────────────────────────────────────────────

export function draw(ctx) {
  ctx.fillStyle = '#000';
  ctx.fillRect(0, 0, CANVAS_W, CANVAS_H);

  const gm = state.gameMode;

  if (gm === GameMode.Dungeon) {
    drawDungeon(ctx);
  } else if (gm === GameMode.HowtoPlay) {
    drawHowtoPlay(ctx);
  } else {
    drawTextScreen(ctx);
  }
}

// ── dungeon ───────────────────────────────────────────────────────────────────

function drawDungeon(ctx) {
  const { landsquare, player, monsters, floor } = state;
  const animFrame = Math.floor((floor % 200) / 20);

  // tile layer (cached)
  if (state.tileLayerDirty) {
    renderTileLayer(animFrame);
    state.tileLayerDirty = false;
  }
  ctx.drawImage(state.tileLayerCanvas, 0, 0);

  // monsters (DLC: depth-band tint; boss drawn larger with a red aura)
  const mSpr = monsterSprite(bandIndex(floor));
  for (let i = 0; i < state.monsterNumber; i++) {
    const m = monsters[i];
    if (!m.alive) continue;
    if (m.ability === A.Stealth) continue;
    if (i === state.bossIdx && mSpr) {
      ctx.strokeStyle = '#ff3333';
      ctx.strokeRect(m.left - 2.5, m.top - 2.5, SC + 5, SC + 5);
      ctx.drawImage(mSpr, 0, m.ability * SC, mSpr.width / 2, SC, m.left - 3, m.top - 3, SC + 6, SC + 6);
    } else {
      drawSprite(ctx, mSpr, m.ability, m.left, m.top);
    }
  }

  // player
  if (player.alive) {
    const cond = (player.condition <= 4) ? player.condition : 0;
    drawSprite(ctx, state.sprites.player, cond, player.left, player.top);
  }

  // separator line
  ctx.strokeStyle = '#444';
  ctx.beginPath(); ctx.moveTo(0, STATUS_Y); ctx.lineTo(SPEC_X, STATUS_Y); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(SPEC_X, 0); ctx.lineTo(SPEC_X, CANVAS_H); ctx.stroke();

  // HP bar
  drawHpBar(ctx);

  // minimap
  drawMinimap(ctx);

  // SpecBar: ability counts
  drawSpecBar(ctx);

  // StatusBar: floor/level/hp/turn
  ctx.fillStyle = '#fff';
  ctx.font = '12px monospace';
  ctx.textAlign = 'center';
  ctx.fillText(state.words[9] || '', SPEC_X / 2, STATUS_Y + 18);

  // MessageBar: messages 0-3
  ctx.font = '13px "MS Gothic", monospace';
  ctx.textAlign = 'left';
  ctx.fillStyle = player.hp > 0 && player.hp <= player.maxHp / 4 ? '#ff7f27' : '#fff';
  for (let k = 0; k <= 3; k++) {
    if (state.words[k]) drawWrapped(ctx, state.words[k], 4, MSG_Y + k * 30, SPEC_X - 8, 13);
  }

  // Auto-play indicator
  if (state.autoPlay) {
    ctx.font = 'bold 11px monospace';
    ctx.textAlign = 'left';
    ctx.fillStyle = '#000';
    ctx.fillRect(2, 2, 38, 16);
    ctx.fillStyle = '#ffcc00';
    ctx.fillText('AUTO', 4, 14);
  }
}

function renderTileLayer(animFrame) {
  const tileCtx = state.tileLayerCanvas.getContext('2d');
  tileCtx.clearRect(0, 0, LAND_NUMBER * SC, LAND_NUMBER * SC);
  const { landsquare } = state;

  for (let i = 0; i < LAND_NUMBER; i++) {
    for (let j = 0; j < LAND_NUMBER; j++) {
      const ls = landsquare[i][j];
      const x = i * SC, y = j * SC;
      const c = ls.condition;

      if (c === T.Wall) {
        drawSprite(tileCtx, state.sprites.wall, animFrame, x, y);
        if (!state.sprites.wall) { tileCtx.fillStyle = TILE_COLORS.wall; tileCtx.fillRect(x, y, SC, SC); }
      } else if (c === T.Room || c === T.Enemy) {
        drawSprite(tileCtx, state.sprites.room, animFrame, x, y);
        if (!state.sprites.room) { tileCtx.fillStyle = TILE_COLORS.room; tileCtx.fillRect(x, y, SC, SC); }
      } else if (c <= T.Stair) {
        drawSprite(tileCtx, state.sprites.box, c, x, y);
        if (!state.sprites.box) {
          tileCtx.fillStyle = c === T.Stair ? TILE_COLORS.stair : (BOX_COLORS[c] || '#888');
          tileCtx.fillRect(x, y, SC, SC);
        }
      }
    }
  }
}

function drawSprite(ctx, canvas, row, dx, dy) {
  if (!canvas) return;
  const sprW = canvas.width / 2;
  const sprH = SC;
  ctx.drawImage(canvas, 0, row * sprH, sprW, sprH, dx, dy, SC, SC);
}

// DLC: per-depth-band hue-rotated monster sprite (cached per band)
const _tintedMonster = {};
function monsterSprite(bi) {
  const base = state.sprites.monster;
  if (!base || bi === 0) return base;
  if (!_tintedMonster[bi]) {
    const c = new OffscreenCanvas(base.width, base.height);
    const cx = c.getContext('2d');
    cx.filter = 'hue-rotate(' + BANDS[bi].hue + 'deg)';
    cx.drawImage(base, 0, 0);
    _tintedMonster[bi] = c;
  }
  return _tintedMonster[bi];
}

function drawHpBar(ctx) {
  const p = state.player;
  if (!p.maxHp) return;
  const BAR_X = 402, BAR_Y = STATUS_Y + 15, BAR_W = 194, BAR_H = 8;
  const ratio = Math.max(0, Math.min(1, p.hp / p.maxHp));
  ctx.fillStyle = '#333';
  ctx.fillRect(BAR_X, BAR_Y, BAR_W, BAR_H);
  ctx.fillStyle = ratio > 0.25 ? '#22b14c' : '#ff4444';
  ctx.fillRect(BAR_X, BAR_Y, Math.round(BAR_W * ratio), BAR_H);
  ctx.strokeStyle = '#666';
  ctx.strokeRect(BAR_X, BAR_Y, BAR_W, BAR_H);
}

function drawMinimap(ctx) {
  const { landsquare, player, floor } = state;
  const MX = SPEC_X + 10, MY = STATUS_Y + 30;
  const CELL = 2;
  const mapW = LAND_NUMBER * CELL;
  const animFrame = Math.floor((floor % 200) / 20);

  for (let i = 0; i < LAND_NUMBER; i++) {
    for (let j = 0; j < LAND_NUMBER; j++) {
      const c = landsquare[i][j].condition;
      if (c === T.Wall) ctx.fillStyle = '#555';
      else if (c === T.Room || c === T.Enemy) ctx.fillStyle = '#222';
      else if (c === T.Stair) ctx.fillStyle = '#fff';
      else ctx.fillStyle = BOX_COLORS[c] || '#888';
      ctx.fillRect(MX + i*CELL, MY + j*CELL, CELL, CELL);
    }
  }
  // boss dot (DLC)
  if (state.bossIdx >= 0) {
    const b = state.monsters[state.bossIdx];
    if (b.alive) {
      ctx.fillStyle = '#ff2222';
      ctx.fillRect(MX + Math.floor(b.left/SC)*CELL - 1, MY + Math.floor(b.top/SC)*CELL - 1, CELL + 2, CELL + 2);
    }
  }
  // player dot
  ctx.fillStyle = '#ffff00';
  ctx.fillRect(MX + Math.floor(player.left/SC)*CELL, MY + Math.floor(player.top/SC)*CELL, CELL, CELL);

  ctx.strokeStyle = '#666';
  ctx.strokeRect(MX - 1, MY - 1, mapW + 2, mapW + 2);
}

function drawSpecBar(ctx) {
  ctx.font = '11px monospace';
  ctx.textAlign = 'left';
  ctx.fillStyle = '#aaa';
  const lines = [
    state.words[4], state.words[5], state.words[6], state.words[7], state.words[8]
  ];
  let y = 10;
  for (const line of lines) {
    if (!line) continue;
    for (const part of line.split('\n')) {
      ctx.fillStyle = part.match(/^\d+$/) ? '#ffff44' : '#aaa';
      ctx.fillText(part, SPEC_X + 4, y);
      y += 14;
    }
    y += 2;
  }
}

// ── HowtoPlay ─────────────────────────────────────────────────────────────────

function drawHowtoPlay(ctx) {
  const { player, landsquare, monsters } = state;
  const pg = player.condition;

  // draw demo sprites for page 1
  if (pg === 1) {
    for (let i = 0; i <= 25; i++) {
      const ls = landsquare[i][0];
      if (!ls.alive) continue;
      const c = ls.condition;
      if (c === T.Wall) drawSprite(ctx, state.sprites.wall, ls.ability, ls.left, ls.top);
      else if (c === T.Room) drawSprite(ctx, state.sprites.room, ls.ability, ls.left, ls.top);
      else if (c === T.Mine) drawSprite(ctx, state.sprites.player, ls.ability, ls.left, ls.top);
      else if (c <= T.Stair) drawSprite(ctx, state.sprites.box, c, ls.left, ls.top);
    }
    // player demo
    drawSprite(ctx, state.sprites.player, 0, player.left, player.top);
    // monsters
    for (let i = 0; i <= 4; i++) drawSprite(ctx, state.sprites.monster, i, monsters[i].left, monsters[i].top);

  } else if (pg >= 2 && pg <= 5) {
    // boxes only
    for (let i = 0; i <= 4; i++) {
      const ls = landsquare[i][0];
      if (ls.alive) drawSprite(ctx, state.sprites.box, ls.condition, ls.left, ls.top);
    }
  }

  // text
  drawTextScreen(ctx);
}

// ── text screens (Entrance, Options, GameOver, GameClear, Museum) ─────────────

function drawTextScreen(ctx) {
  ctx.font = 'bold 15px "MS Gothic", sans-serif';
  ctx.textAlign = 'center';
  ctx.fillStyle = state.gameMode === GameMode.GameClear ? '#0000ff' : '#fff';

  const lines = [];
  for (let k = 0; k <= 8; k++) {
    if (state.words[k]) lines.push(...state.words[k].split('\n'));
    else lines.push('');
  }

  let y = 60;
  for (const line of lines) {
    ctx.fillText(line, CANVAS_W / 2, y);
    y += 22;
  }
}

// ── utility ───────────────────────────────────────────────────────────────────

function drawWrapped(ctx, text, x, y, maxW, lineH) {
  for (const line of text.split('\n')) {
    ctx.fillText(line, x, y);
    y += lineH + 2;
  }
}
