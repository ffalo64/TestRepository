// VB6 の数値変換の挙動を再現するヘルパ

// Double → Long 代入（暗黙の CLng）。VB6 は偶数丸め（銀行丸め）。
export function cLng(x) {
  const f = Math.floor(x);
  const d = x - f;
  if (d > 0.5) return f + 1;
  if (d < 0.5) return f;
  return f % 2 === 0 ? f : f + 1;
}

// Int()
export const vbInt = Math.floor;

let rng = Math.random;

// Rnd()
export function rnd() {
  return rng();
}

// テスト用に乱数源を差し替える
export function setRandom(fn) {
  rng = fn;
}

// 再現可能な乱数（テスト用）
export function mulberry32(seed) {
  let a = seed >>> 0;
  return function () {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
