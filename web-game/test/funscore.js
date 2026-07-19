// FUN score: scalar "fun" comparison between two headless.js reports.
//   node test/funscore.js <baseline.txt> <variant.txt>
// Five equally-weighted categories (user decision 2026-07-19). Each metric is
// scored as its %-change vs baseline, signed so that the target direction is
// positive, clipped to ±100%; metrics within a category are averaged, then the
// five category scores are averaged into FUN. Adopt when FUN > 0 and the
// guardrails hold (CLEAR 45-65%, sudden-death 0%).
import { readFileSync } from 'node:fs';

function parse(path) {
  const t = readFileSync(path, 'utf8');
  const g = (re) => { const m = t.match(re); return m ? parseFloat(m[1]) : NaN; };
  return {
    clear: g(/1000F reach rate: ([\d.]+)%/),
    nearDeath: g(/nearDeath\/100F=([\d.]+)/),
    revive: g(/蘇生発動\/100F=([\d.]+)/),
    dull: g(/無イベント階=([\d.]+)%/),
    stagnation: g(/最大Lv停滞=([\d.]+)階/),
    lvlRate: g(/Lv\/100F=([\d.]+)/),
    bigHit: g(/bigHit\/100F=([\d.]+)/),
    sudden: t.includes('即死率(HP死のうち)=-%') ? 0 : g(/即死率\(HP死のうち\)=([\d.]+)%/),
    ticks: g(/ticks\/floor=([\d.]+)/),
    maxFloor: g(/max1floor=([\d.]+)/),
  };
}

const [bFile, vFile] = process.argv.slice(2);
if (!bFile || !vFile) { console.error('usage: node test/funscore.js <baseline.txt> <variant.txt>'); process.exit(1); }
const b = parse(bFile), v = parse(vFile);

const clip = x => Math.max(-100, Math.min(100, x));
// signed %-change: dir=+1 means "higher is better", dir=-1 "lower is better"
const pct = (key, dir) => {
  if (!isFinite(b[key]) || !isFinite(v[key]) || b[key] === 0) return 0;
  return clip(dir * 100 * (v[key] - b[key]) / b[key]);
};

const cats = [
  ['緊張感', (pct('nearDeath', +1) + pct('revive', -1)) / 2],
  ['単調さ', pct('dull', -1)],
  ['成長',   (pct('stagnation', -1) + pct('lvlRate', +1)) / 2],
  ['理不尽', pct('bigHit', -1)],
  ['テンポ', (pct('ticks', -1) + pct('maxFloor', -1)) / 2],
];
const fun = cats.reduce((a, [, s]) => a + s, 0) / cats.length;

const guardClear = v.clear >= 45 && v.clear <= 65;
const guardSudden = v.sudden === 0;

console.log('═══ FUN score ═══  (baseline: ' + bFile + ')');
for (const [name, s] of cats) console.log(`  ${name}: ${s >= 0 ? '+' : ''}${s.toFixed(1)}`);
console.log(`  FUN合計: ${fun >= 0 ? '+' : ''}${fun.toFixed(1)}  (>0で採用)`);
console.log(`  ガードレール: CLEAR=${v.clear}% ${guardClear ? 'OK' : 'NG'} / 即死率=${v.sudden}% ${guardSudden ? 'OK' : 'NG'}`);
console.log(`  判定: ${fun > 0 && guardClear && guardSudden ? '✅ 採用' : '❌ 棄却'}`);
