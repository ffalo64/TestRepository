// 最小限の BMP デコーダ（24bpp / 8bpp, 非圧縮）。
// ブラウザの画像デコードに頼らず、マップ判定用のピクセル値を正確に取るために使う。
// ブラウザ・Node 両方で動く（ArrayBuffer を受け取る）。

export function decodeBmp(buffer) {
  const v = new DataView(buffer);
  if (v.getUint8(0) !== 0x42 || v.getUint8(1) !== 0x4d) throw new Error('not a BMP');
  const offset = v.getUint32(10, true);
  const headerSize = v.getUint32(14, true);
  const width = v.getInt32(18, true);
  const rawHeight = v.getInt32(22, true);
  const bpp = v.getUint16(28, true);
  const compression = v.getUint32(30, true);
  if (compression !== 0) throw new Error('compressed BMP is not supported');
  if (bpp !== 24 && bpp !== 8) throw new Error(`unsupported bpp: ${bpp}`);

  const height = Math.abs(rawHeight);
  const bottomUp = rawHeight > 0;
  const rowSize = Math.ceil((width * bpp) / 32) * 4;

  let palette = null;
  if (bpp === 8) {
    const colors = v.getUint32(46, true) || 256;
    palette = [];
    const pOff = 14 + headerSize;
    for (let i = 0; i < colors; i++) {
      const p = pOff + i * 4;
      palette.push([v.getUint8(p + 2), v.getUint8(p + 1), v.getUint8(p)]);
    }
  }

  const data = new Uint8ClampedArray(width * height * 4);
  for (let y = 0; y < height; y++) {
    const row = offset + (bottomUp ? height - 1 - y : y) * rowSize;
    for (let x = 0; x < width; x++) {
      let r, g, b;
      if (bpp === 24) {
        const p = row + x * 3;
        b = v.getUint8(p);
        g = v.getUint8(p + 1);
        r = v.getUint8(p + 2);
      } else {
        [r, g, b] = palette[v.getUint8(row + x)];
      }
      const o = (y * width + x) * 4;
      data[o] = r;
      data[o + 1] = g;
      data[o + 2] = b;
      data[o + 3] = 255;
    }
  }
  return { width, height, data };
}

// GetPixel 相当。0xRRGGBB を返す。
export function getPixel(img, x, y) {
  const o = (y * img.width + x) * 4;
  return (img.data[o] << 16) | (img.data[o + 1] << 8) | img.data[o + 2];
}
