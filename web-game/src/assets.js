// 画像の読み込み。VB6 版は「左半分=絵、右半分=マスク(黒=不透明)」の BMP を
// AND/OR の BitBlt で重ねていたので、それをアルファ付きのキャンバスに変換する。

import { decodeBmp } from './bmp.js';

const BASE = import.meta.env.BASE_URL;

async function fetchBmp(name) {
  const res = await fetch(`${BASE}images/${name}`);
  if (!res.ok) throw new Error(`failed to load ${name}`);
  return decodeBmp(await res.arrayBuffer());
}

function maskedSheet(img) {
  const w = img.width / 2;
  const h = img.height;
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext('2d');
  const out = ctx.createImageData(w, h);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const s = (y * img.width + x) * 4;
      const m = (y * img.width + x + w) * 4;
      const o = (y * w + x) * 4;
      const maskLum = img.data[m] + img.data[m + 1] + img.data[m + 2];
      out.data[o] = img.data[s];
      out.data[o + 1] = img.data[s + 1];
      out.data[o + 2] = img.data[s + 2];
      out.data[o + 3] = maskLum < 384 ? 255 : 0;
    }
  }
  ctx.putImageData(out, 0, 0);
  return canvas;
}

export async function loadAssets() {
  const [player, monster, box, room, wall, map] = await Promise.all(
    ['Player.bmp', 'Monster.bmp', 'Box.bmp', 'Room.bmp', 'Wall.bmp', 'Map40.bmp'].map(fetchBmp),
  );
  return {
    sheets: {
      player: maskedSheet(player),
      monster: maskedSheet(monster),
      box: maskedSheet(box),
      room: maskedSheet(room),
      wall: maskedSheet(wall),
    },
    map,
  };
}
