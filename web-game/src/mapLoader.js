import { SC, LAND_NUMBER, T } from './constants.js';
import { state } from './state.js';

export async function loadSprites() {
  const defs = [
    ['player',  '/images/Player.bmp'],
    ['box',     '/images/Box.bmp'],
    ['monster', '/images/Monster.bmp'],
    ['room',    '/images/Room.bmp'],
    ['wall',    '/images/Wall.bmp'],
  ];
  await Promise.all(defs.map(([key, src]) => loadSpriteCanvas(key, src)));

  state.tileLayerCanvas = new OffscreenCanvas(LAND_NUMBER * SC, LAND_NUMBER * SC);
}

async function loadSpriteCanvas(key, src) {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => {
      const c = new OffscreenCanvas(img.naturalWidth, img.naturalHeight);
      const ctx = c.getContext('2d');
      ctx.drawImage(img, 0, 0);
      const id = ctx.getImageData(0, 0, c.width, c.height);
      const d = id.data;
      // black → transparent (color-key)
      for (let i = 0; i < d.length; i += 4)
        if (d[i] === 0 && d[i+1] === 0 && d[i+2] === 0) d[i+3] = 0;
      ctx.putImageData(id, 0, 0);
      state.sprites[key] = c;
      resolve();
    };
    img.onerror = () => { state.sprites[key] = null; resolve(); };
    img.src = src;
  });
}

export async function loadMap() {
  return new Promise(resolve => {
    const img = new Image();
    img.onload = () => {
      const c = new OffscreenCanvas(img.naturalWidth, img.naturalHeight);
      c.getContext('2d').drawImage(img, 0, 0);
      const id = c.getContext('2d').getImageData(0, 0, c.width, c.height);
      state.mapPixels = id.data;
      state.mapW = c.width;
      state.mapH = c.height;
      colorSet();
      resolve();
    };
    img.onerror = () => resolve();
    img.src = '/images/Map40.bmp';
  });
}

function colorSet() {
  const { mapPixels, mapW } = state;
  if (!mapPixels) return;
  for (let x = 1; x <= 6; x++) {
    const idx = (1 * mapW + x) * 4;
    state.colorRef[x] = { r: mapPixels[idx], g: mapPixels[idx+1], b: mapPixels[idx+2] };
  }
}

function getMapPixel(x, y) {
  const { mapPixels, mapW } = state;
  const i = (y * mapW + x) * 4;
  return { r: mapPixels[i], g: mapPixels[i+1], b: mapPixels[i+2] };
}

function colorsMatch(c, ref) {
  return Math.abs(c.r - ref.r) <= 10 && Math.abs(c.g - ref.g) <= 10 && Math.abs(c.b - ref.b) <= 10;
}

export function mapSet() {
  if (state.mapPixels) {
    bitmapMapSet();
  } else {
    proceduralMapSet();
  }
}

function bitmapMapSet() {
  const { landsquare, colorRef, monsters, mapW, mapH } = state;
  const ox = Math.floor(Math.random() * Math.floor(mapW / LAND_NUMBER)) * LAND_NUMBER;
  const oy = Math.floor(Math.random() * Math.floor(mapH / LAND_NUMBER)) * LAND_NUMBER;

  let roomCount = 0;
  for (let i = 0; i < LAND_NUMBER; i++) {
    for (let j = 0; j < LAND_NUMBER; j++) {
      const ls = landsquare[i][j];
      const c = getMapPixel(ox + i, oy + j);
      let cond = T.Room;
      if      (colorRef[1] && colorsMatch(c, colorRef[1])) cond = T.BlueBox;
      else if (colorRef[2] && colorsMatch(c, colorRef[2])) cond = T.RedBox;
      else if (colorRef[3] && colorsMatch(c, colorRef[3])) cond = T.YellowBox;
      else if (colorRef[4] && colorsMatch(c, colorRef[4])) cond = T.GreenBox;
      else if (colorRef[5] && colorsMatch(c, colorRef[5])) cond = T.Wall;
      else if (colorRef[6] && colorsMatch(c, colorRef[6])) cond = T.Room;

      ls.condition = cond;
      ls.maxHp = monsters[0].maxHp;
      ls.hp = ls.maxHp;
      ls.def = monsters[0].def;
      ls.alive = true;
      if (cond !== T.Room) roomCount++;
    }
  }
  state.randomPosition.hp = LAND_NUMBER * LAND_NUMBER - roomCount;
}

function proceduralMapSet() {
  const { landsquare, monsters } = state;

  for (let i = 0; i < LAND_NUMBER; i++)
    for (let j = 0; j < LAND_NUMBER; j++) {
      const ls = landsquare[i][j];
      ls.condition = T.Wall;
      ls.maxHp = monsters[0].maxHp; ls.hp = ls.maxHp;
      ls.def = monsters[0].def; ls.alive = true;
    }

  const rooms = [];
  for (let attempt = 0; attempt < 50; attempt++) {
    const rw = 3 + Math.floor(Math.random() * 6);
    const rh = 3 + Math.floor(Math.random() * 6);
    const rx = 1 + Math.floor(Math.random() * (LAND_NUMBER - rw - 2));
    const ry = 1 + Math.floor(Math.random() * (LAND_NUMBER - rh - 2));
    let overlap = false;
    for (const r of rooms) {
      if (rx < r.x+r.w+1 && rx+rw > r.x-1 && ry < r.y+r.h+1 && ry+rh > r.y-1) { overlap = true; break; }
    }
    if (!overlap) {
      rooms.push({x:rx, y:ry, w:rw, h:rh});
      for (let i = rx; i < rx+rw; i++)
        for (let j = ry; j < ry+rh; j++)
          landsquare[i][j].condition = T.Room;
    }
  }

  for (let k = 1; k < rooms.length; k++) {
    const a = rooms[k-1], b = rooms[k];
    const ax = Math.floor(a.x + a.w/2), ay = Math.floor(a.y + a.h/2);
    const bx = Math.floor(b.x + b.w/2), by = Math.floor(b.y + b.h/2);
    for (let i = Math.min(ax,bx); i <= Math.max(ax,bx); i++)
      if (i > 0 && i < LAND_NUMBER-1) landsquare[i][ay].condition = T.Room;
    for (let j = Math.min(ay,by); j <= Math.max(ay,by); j++)
      if (j > 0 && j < LAND_NUMBER-1) landsquare[bx][j].condition = T.Room;
  }

  // scatter some boxes
  for (let t = 0; t < 20; t++) {
    const bi = 1 + Math.floor(Math.random() * (LAND_NUMBER-2));
    const bj = 1 + Math.floor(Math.random() * (LAND_NUMBER-2));
    if (landsquare[bi][bj].condition === T.Room)
      landsquare[bi][bj].condition = Math.floor(Math.random() * 4);
  }

  let roomCount = 0;
  for (let i = 0; i < LAND_NUMBER; i++)
    for (let j = 0; j < LAND_NUMBER; j++)
      if (landsquare[i][j].condition === T.Room) roomCount++;
  state.randomPosition.hp = roomCount;
}
