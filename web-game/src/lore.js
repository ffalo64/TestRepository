// DLC: depth-themed monster bands, floor guardians (bosses), and player titles.

// One band per 100 floors. `hue` rotates the monster sprite palette so each
// depth zone looks distinct without new image assets.
export const BANDS = [
  { name: 'ヘキサスライム',     boss: 'キングスライム',       hue: 0 },
  { name: 'ドクコウモリ',       boss: '蝙蝠女王ノクターナ',   hue: 45 },
  { name: 'ゴブリンソルジャー', boss: 'ゴブリンロード',       hue: 90 },
  { name: 'さまようスケルトン', boss: '死霊将軍デスペル',     hue: 135 },
  { name: 'オークウォーリア',   boss: 'オークの大王',         hue: 180 },
  { name: 'リビングアーマー',   boss: '呪鎧ヴァルグレイヴ',   hue: 215 },
  { name: 'ガーゴイル',         boss: '石翼公ガルガンチュア', hue: 250 },
  { name: 'アイスゴーレム',     boss: '氷帝グラキエス',       hue: 285 },
  { name: 'ヘルハウンド',       boss: '獄炎犬ケルベロス',     hue: 320 },
  { name: 'カオスドラゴン',     boss: '深淵竜アビスヴァーン', hue: 350 },
];

export function bandIndex(floor) {
  return Math.min(BANDS.length - 1, Math.max(0, Math.floor((floor - 1) / 100)));
}

// Player title, derived from persistent records (see state.records).
export function titleFor(r) {
  if (r.clears >= 5) return '伝説の勇者';
  if (r.clears >= 1) return '深淵を制した勇者';
  if (r.deepestFloor >= 800) return '深淵を覗く者';
  if (r.deepestFloor >= 500) return '中層の覇者';
  if (r.deepestFloor >= 300) return '迷宮の狩人';
  if (r.deepestFloor >= 100) return 'ダンジョン探索者';
  return '駆け出しの冒険者';
}
