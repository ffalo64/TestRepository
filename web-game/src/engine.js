import { SC, LAND_NUMBER, PLAYER_CONST, GameMode, T, A, TURN_CONST } from './constants.js';
import { state, makeStatus } from './state.js';
import { mapSet } from './mapLoader.js';
import { playMusic, stopAll, sfx } from './audio.js';
import { BANDS, bandIndex, titleFor } from './lore.js';

// ── helpers ──────────────────────────────────────────────────────────────────

const rnd = () => Math.random();
const rint = n => Math.floor(n);

// ── monsterSet ────────────────────────────────────────────────────────────────

export function monsterSet() {
  const gm = state.gameMode;
  if (gm === GameMode.Entrance || gm === GameMode.Dungeon) {
    for (let i = 0; i < 1000; i++) {
      const m = state.monsters[i];
      m.name = 'ヘキサスライム' + i;
      m.explanation = '';
      m.hp = 10; m.maxHp = 10; m.level = 1; m.exp = 10;
      m.ability = A.Noability;
      m.atk = 5; m.def = 1;
      m.height = SC; m.width = SC;
      m.direction = 1;
    }
  } else if (gm === GameMode.HowtoPlay) {
    for (let i = 0; i <= 4; i++) {
      const m = state.monsters[i];
      m.name = 'ヘキサスライム' + i;
      m.left = 300 + SC * 2 * i;
      m.top = state.player.top + 38;
      m.explanation = '';
      m.ability = i;
      m.height = SC; m.width = SC;
    }
  }
}

// ── landSet ───────────────────────────────────────────────────────────────────

export function landSet() {
  const { player, abilityHp, landsquare, words, oword } = state;
  const gm = state.gameMode;

  if (gm === GameMode.Entrance) {
    state.abilityHp.fill(0);
    player.alive = false; player.height = SC; player.width = SC;
    player.left = 150; player.top = 50;
    player.level = 1; player.maxHp = 20; player.hp = 20;
    player.condition = 0; player.direction = 1;
    state.floor = 0;
    for (let i = 0; i < LAND_NUMBER; i++)
      for (let j = 0; j < LAND_NUMBER; j++) {
        const ls = landsquare[i][j];
        ls.width = SC; ls.height = SC;
        ls.left = SC * i; ls.top = SC * j;
        ls.condition = T.Room; ls.alive = true;
      }

  } else if (gm === GameMode.Dungeon) {
    for (let k = 0; k < state.words.length; k++) state.words[k] = '';
    player.hp = player.maxHp; player.alive = true;

  } else if (gm === GameMode.HowtoPlay) {
    player.width = SC; player.height = SC;
    player.left = 300; player.top = 199;
    // resize landsquare demo array (use indices 0..25 of row 0)
    for (let i = 0; i <= 25; i++) {
      landsquare[i][0].width = SC; landsquare[i][0].height = SC;
    }

  } else if (gm === GameMode.GameOver) {
    stopAll();
    playMusic(2);
    sfx('death');
    saveRecords();

  } else if (gm === GameMode.GameClear) {
    stopAll();
    playMusic(1);
    saveRecords();

  } else if (gm === GameMode.Museum) {
    const save = loadSave();
    if (save) {
      player.hp = save[1]; player.maxHp = save[2]; player.atk = save[3];
      player.def = save[4]; player.exp = save[5]; player.level = save[6];
      player.ability = save[15];
      for (let k = 0; k <= 5; k++) state.abilityHp[k] = save[7 + k];
      state.floor = save[13]; state.turn = save[14];
      state.monsterNumber = save['m7'];
      // monster base stats
      const m0 = state.monsters[0];
      m0.maxHp = save['m1']; m0.hp = m0.maxHp;
      m0.atk = save['m2']; m0.def = save['m3'];
      m0.exp = save['m4']; m0.level = save['m5'];
      m0.ability = save['m6'];
    }
    statusCheck(false);
  }
}

// ── wordSet ───────────────────────────────────────────────────────────────────

export function wordSet() {
  const { player, landsquare, abilityHp, words, oword } = state;
  const gm = state.gameMode;

  if (gm === GameMode.Entrance) {
    words[0] = 'ダンジョンと不思議の箱';
    words[1] = 'ダンジョンに入る (Enterキー)\n不思議の箱を使ってチャレンジが始まります。';
    words[2] = 'How to play (zキー)\n操作の概要を知りたいなら';
    words[3] = 'Option (cキー)\n難易度設定';
    words[4] = 'Museum (dキー)\n記録の閲覧\n\n(矢印キーは範囲を外に使いません)';
    words[5] = '\n音楽データ提供';
    words[6] = 'フリー音楽 H/MIX GALLERY\n管理者：H/T';
    words[7] = 'http://www.hmix.net/';
    words[8] = state.records.runs > 0
      ? '\n称号: ' + titleFor(state.records) + '   最深到達: ' + state.records.deepestFloor + 'F'
      : '';

  } else if (gm === GameMode.Dungeon) {
    words[4] = '全体攻撃(z)\n' + abilityHp[0];
    words[5] = 'Hp全回復(x)\n' + abilityHp[1];
    words[6] = '全消去(c)\n' + abilityHp[2];
    words[7] = 'モンスター除去(d)\n' + abilityHp[3];
    words[8] = '次の階へ(Enter)\n' + abilityHp[4] + '\n蘇りの術\n' + abilityHp[5];
    words[9] = state.floor + 'F  Lv' + player.level
      + '  HP' + rint(player.hp) + '/' + player.maxHp + '  Turn ' + state.turn;

    // message fade via oword
    for (let i = 0; i <= 3; i++) {
      const ow = oword[i];
      if (ow.explanation !== words[i]) {
        ow.explanation = words[i]; ow.hp = 40;
      } else if (ow.hp > 0) {
        ow.hp--;
      } else {
        words[i] = '';
      }
    }
    // set land names
    const nameMap = [T.BlueBox, T.RedBox, T.YellowBox, T.GreenBox, T.PurpleBox, T.Wall];
    const names = ['青箱', '赤箱', '黄箱', '緑箱', '紫箱', '壁'];
    for (let i = 0; i < LAND_NUMBER; i++)
      for (let j = 0; j < LAND_NUMBER; j++) {
        const ls = landsquare[i][j];
        const idx = nameMap.indexOf(ls.condition);
        if (idx >= 0) ls.name = names[idx];
      }

  } else if (gm === GameMode.HowtoPlay) {
    const pg = player.condition;
    const pages = [
      ['How to play\n1/7', '\nこのゲームは気楽に遊べるゲームです。', 'なので、説明を読むのが嫌な人は、\n説明は読み飛ばしてしまって構いです。', '', '', '', '', '', '\n(xキーでメニュー画面に戻る,右左キーでページ選択)'],
      ['How to play\n2/7', 'このゲームはモンスターを倒しながら、\n次の階を目指して階段を下りていくゲームです。', 'プレイヤー                      \nモンスター                      ', '階段                      \n(階段はマウスポインタと見た目が紛らわしいので注意してください。)', '操作は基本的に十字キーによる移動だけです。\n攻撃は隣の敵マスに進もうとするだけで出せます。', '\n壁                     ', '道                     \n', 'ダンジョンの地形は主に2種類あります。\nそのうち壁の上には進むことができません。', '\n(xキーでメニュー画面に戻る,右左キーでページ選択)'],
      ['How to play\n3/7', 'ダンジョンに出てくる5つの箱\n', '青…受けるとHPを回復し、モンスターの状態異常が解除される。\n赤…受けるとマイナス効果が与えられる。', '黄…受けると色々な効果が与えられる\n緑…受けるとプラス効果が与えられる。', '紫…受けるとコマンドキーの使用回数を増加。\n', '箱について\n', 'この5つの箱のうち、紫以外は消えてしまいます。\nその箱は効果などで別の箱に変わらず、', '難易度によって効果が変わります。\n', '\n(xキーでメニュー画面に戻る,右左キーでページ選択)'],
      ['How to play\n4/7', 'ダンジョン内で使える5つのコマンド\n', 'zキー...敵全体と青/緑/紫の箱に攻撃\nxキー...Hpを全回復させる', 'cキー...階段以外の全ての箱、壁、モンスターを消去する。\ndキー...モンスターを箱に変化させる', '(zやcで壊した青/緑/紫の箱からは、直接壊すより\n弱い効果を得られます。1回につき3個まで。)', 'Enterキー...次のフロアに降りる\n', '(これらのことはゲーム画面でも表示されているので、\nどのキーがどんな効果かは覚えなくて大丈夫です。)', '', '\n(xキーでメニュー画面に戻る,右左キーでページ選択)'],
      ['How to play\n5/7', 'セーブ/ロードについて\n', 'ゲームを中断したいなら階段の上でSボタンを押すとセーブで出来ます。\n再開したい時は記録の閲覧からロードできます。', 'セーブすると、以前のデータは消えてしまうので、ご注意してください。\n', 'ターンについて\n', 'このゲームでは1ターンごとにターン数が1減っていきます。\nターン数が0になるとゲームオーバーです。', 'そのことを踏まえると、ターン回数は回復します。\n', 'これ以降の説明は、壁を通れるようになったものに入ったようなものです。\n蘇りの術で復活します。', '\n(xキーでメニュー画面に戻る,右左キーでページ選択)'],
      ['How to play\n6/7', 'ダンジョン最大の6つのステータス\n', 'レベル\nこれが上がると、全体的に強くなります。', '攻撃力\nこの数値が大きいほど、与えるダメージが大きくなります。', '守備力\nこの数値が大きいほど、受けるダメージが少なくなります。', 'Hp\n体力です。これが0になると倒れます。', '最大Hp\nHpの最大値です。Hpはこれ以上に回復しません。', '経験値\nこれが溜まるとレベルアップしていきます。', 'モンスターを倒すと、その経験値が自分のものになります。\n(xキーでメニュー画面に戻る,右左キーでページ選択)'],
      ['How to play\n7/7', 'DLC: 深層の守護者\n', '100階ごとに、そのフロアの守護者(ボス)が待ち構えています。\n赤いオーラをまとった大きなモンスターが目印です。', '倒せば紫の箱と追加ターンが手に入りますが、\n通常より遥かに強敵です。コマンドで切り抜けるのも手です。', 'モンスターは100階ごとに姿と名前を変え、強くなっていきます。\n', '実績と称号\n', '最深到達階・討伐数・守護者討伐などの記録は\nMuseum(記録の閲覧)でいつでも確認できます。', '記録を伸ばすと、あなたの称号が変わっていきます。', '\n(xキーでメニュー画面に戻る,右左キーでページ選択)'],
    ];
    if (pg >= 0 && pg <= 6) {
      const p = pages[pg];
      for (let k = 0; k <= 8; k++) words[k] = p[k] || '';
    }

  } else if (gm === GameMode.Options) {
    const diffNames = ['Very Easy', 'Easy', 'Normal', 'Hard', 'Very Hard'];
    words[0] = 'Options\n';
    words[1] = (diffNames[player.ability] || 'Normal') + '\n';
    words[2] = '違う難易度ほど蘇りの術が出やすいです。\n';
    words[3] = '\n';
    words[4] = 'xキーでメニュー画面に戻る,右左キーで難易度選択\nEnterキーでダンジョンに入る。';

  } else if (gm === GameMode.GameOver) {
    words[0] = 'GameOver';
    words[1] = '';
    words[2] = 'プレイヤーは力尽きた';
    words[3] = '';
    words[4] = 'xキーでメニュー画面に戻る';

  } else if (gm === GameMode.GameClear) {
    words[0] = 'GameClear\nついに最深部1000Fです。';
    words[1] = '\nこのゲームを最後まで遊んでくれたあなたは勇者です。';
    words[2] = 'クリア時のステータス\n';
    words[3] = 'HP ' + rint(player.hp) + '/' + player.maxHp + '\nLv ' + player.level;
    words[4] = '称号: ' + titleFor(state.records) + '\n(クリア' + state.records.clears + '回目)';
    words[5] = '\nxキーでメニュー画面に戻る';

  } else if (gm === GameMode.Museum) {
    words[0] = '現在の記録\n';
    words[1] = (state.floor + 1) + 'F  Lv' + player.level + '\nHP' + rint(player.hp) + '/' + player.maxHp + '  Turn ' + state.turn;
    words[2] = '全体攻撃 ' + state.abilityHp[0] + '\nHp全回復 ' + state.abilityHp[1] + '\n全消去 ' + state.abilityHp[2];
    words[3] = 'モンスター除去 ' + state.abilityHp[3] + '\n次の階へ ' + state.abilityHp[4] + '\n蘇りの術 ' + state.abilityHp[5];
    words[4] = 'Lボタンを押すと、このセーブデータから始められます。\nxキーでメニュー画面に戻る';
    const r = state.records;
    words[5] = '\n― 実績 ―';
    words[6] = '称号: ' + titleFor(r) + '\n最深到達 ' + r.deepestFloor + 'F   クリア ' + r.clears + '回   挑戦 ' + r.runs + '回';
    words[7] = 'モンスター討伐 ' + r.totalKills + '   箱破壊 ' + r.totalBoxes + '   守護者討伐 ' + r.bossKills;
    const seen = Math.min(10, Math.max(1, Math.ceil(r.deepestFloor / 100)));
    words[8] = 'モンスター図鑑: ' + BANDS.slice(0, seen).map(b => b.name).join('、') + (seen < 10 ? '、???' : '  (全種発見!)');
  }
}

// ── statusCheck ───────────────────────────────────────────────────────────────

export function statusCheck(levelup) {
  const { player, monsters, landsquare, abilityHp, words } = state;

  // Player level up
  if (player.exp >= Math.pow(player.level, 3) && player.level < 999 && player.alive) {
    const dHp = rint(5 + rnd() * 3);
    player.level++;
    player.maxHp += dHp;
    player.hp += dHp;
    player.atk = rint(player.atk * 1.1) + 1;
    player.def = player.def + 1;
    sfx('levelup');
  }

  // Player stat caps
  if (player.atk > 1e9 - 1) player.atk = 1e9 - 1;
  if (player.def > 1e9 - 1) player.def = 1e9 - 1;
  if (player.def <= 0) player.def = 1;
  if (state.turn > 1e4 - 1) state.turn = 1e4 - 1;
  if (player.exp > 1e9 - 1) player.exp = 1e9 - 1;
  if (player.maxHp > 1e7 - 1) player.maxHp = 1e7 - 1;
  if (player.hp > player.maxHp) player.hp = player.maxHp;

  // Game over check
  if ((player.hp <= 0 || state.turn <= 0) && player.alive) {
    if (abilityHp[5] === 0) {
      player.hp = 0; player.direction = 1; player.alive = false;
      state.gameMode = GameMode.GameOver;
      landSet(); wordSet();
    } else {
      player.alive = false;
      abilityEffect('KeyD');
    }
  }

  // Monster stat updates
  // Balance (JS port): growth 1.12→1.06 (atk), 1.1→1.08 (hp), atk cap 1e9→3e5
  // — tuned so autoplay reaches 1000F ~50% of the time (VB6 was unwinnable:
  // monster atk hit the 1e9 cap near floor 170 and one-shot the player).
  const cap = levelup ? 1000 : state.monsterNumber;
  for (let i = 0; i < cap; i++) {
    const m = monsters[i];
    if (m.level < 999 && levelup) {
      m.level++;
      m.maxHp = rint(m.maxHp * 1.08) + 1;
      m.hp = m.maxHp;
      m.atk = rint(m.atk * 1.06) + 1;
      m.def = m.def + 1;
      m.exp = m.exp + 10;
      if (i === 0) words[0] = 'モンスター達はLevel' + m.level + 'になった。';
    }
    if (m.atk > 3e5-1) m.atk = 3e5-1;
    if (m.def > 1e9-1) m.def = 1e9-1;
    if (m.def <= 0) m.def = 1;
    if (m.exp > 1e4-1) m.exp = 1e4-1;
    if (m.hp > m.maxHp) m.hp = m.maxHp;
    if (m.maxHp > 1e6-1) m.maxHp = 1e6-1;
  }

  // Landsquare stat caps
  for (let i = 0; i < LAND_NUMBER; i++)
    for (let j = 0; j < LAND_NUMBER; j++) {
      const ls = landsquare[i][j];
      if (ls.def > 1e9-1) ls.def = 1e9-1;
      if (ls.def <= 0) ls.def = 1;
      if (ls.hp > ls.maxHp) ls.hp = ls.maxHp;
      if (ls.maxHp > 1e6-1) ls.maxHp = 1e6-1;
    }
}

// ── battleCheck ───────────────────────────────────────────────────────────────

// Shared player-kill bookkeeping. DLC: a slain floor guardian leaves a purple
// box and bonus turns where it stood.
function _monsterKilled(idx) {
  const m = state.monsters[idx];
  const li = rint(m.left / SC), lj = rint(m.top / SC);
  state.landsquare[li][lj].condition = T.Room;
  state.player.exp += m.exp;
  state.records.totalKills++;
  if (idx === state.bossIdx) {
    state.landsquare[li][lj].condition = T.PurpleBox;
    if (state.player.condition !== TURN_CONST) state.turn += 100;
    state.records.bossKills++;
    state.bossIdx = -1;
    state.words[1] = '守護者「' + m.name + '」を倒した! 紫の箱が残された。(+100ターン)';
    sfx('bossdown');
  }
  state.tileLayerDirty = true;
}

export function battleCheck(attacker, i, j) {
  const { player, monsters, landsquare, words } = state;

  if (attacker === PLAYER_CONST) {
    if (j < LAND_NUMBER) {
      // attack a tile (box/wall)
      const ls = landsquare[i][j];
      let dmg = rint(player.atk * (rnd() * 0.2 + 0.9) / ls.def) + 1;
      if (dmg > 1e7-1) dmg = 1e7-1;
      ls.hp -= dmg;
      words[2] = ls.name + 'に' + dmg + 'ダメージを与えた';
      sfx('hit');
      if (ls.hp <= 0) {
        ls.hp = 0;
        ls.ability = rint(rnd() * 15);
        if (player.condition !== TURN_CONST) state.turn += 15; // balance: was +10 (deep floors died to turn exhaustion)
        state.records.totalBoxes++;
        sfx('break');
        boxEffect(i, j);
      }
    } else {
      // attack a monster
      const m = monsters[i];
      if (m.alive) {
        let dmg = rint(player.atk * (rnd() * 0.2 + 0.9) / m.def) + 1;
        if (dmg > 1e7-1) dmg = 1e7-1;
        m.hp -= dmg;
        words[2] = m.name + 'に' + dmg + 'ダメージを与えた';
        sfx('hit');
        if (m.hp <= 0) {
          m.hp = 0; m.alive = false;
          _monsterKilled(i);
        }
      }
    }
  } else {
    // monster attacks
    if (j < LAND_NUMBER) {
      const ls = landsquare[i][j];
      if (ls.condition !== T.Stair) {
        const m = monsters[attacker];
        let dmg = rint(m.atk * (rnd() * 0.2 + 0.9) / ls.def) + 1;
        if (dmg > 1e7-1) dmg = 1e7-1;
        ls.hp -= dmg;
        words[2] = ls.name + 'に' + dmg + 'ダメージを与えた';
        if (ls.hp <= 0) { ls.hp = 0; ls.condition = T.Room; state.tileLayerDirty = true; }
      }
    } else {
      const m = monsters[attacker];
      let dmg = rint(m.atk * (rnd() * 0.2 + 0.9) / player.def) + 1;
      if (dmg > 1e7-1) dmg = 1e7-1;
      player.hp -= dmg;
      state.sumDamage += dmg;
    }
  }
}

// ── positionCheck ─────────────────────────────────────────────────────────────

export function positionCheck(character) {
  const { player, monsters, landsquare } = state;

  if (character === PLAYER_CONST) {
    const p = player;
    const pi = rint(p.left / SC);
    const pj = rint(p.top / SC);
    const ls = landsquare[pi][pj];

    if (ls.condition === T.Stair) {
      if (state.abilityHp[6] === 0) {
        floorSet();
      } else {
        abilityEffect('KeyS');
      }
    } else if (ls.condition <= T.PurpleBox) {
      battleCheck(PLAYER_CONST, pi, pj);
      p.left = p.oleft; p.top = p.otop;
    } else if (ls.condition === T.Wall) {
      if (p.condition === A.WallBreak && p.left !== 0 && p.top !== 0
          && p.left !== (LAND_NUMBER-1)*SC && p.top !== (LAND_NUMBER-1)*SC) {
        battleCheck(PLAYER_CONST, pi, pj);
      }
      p.left = p.oleft; p.top = p.otop;
    }

    // check monster collisions
    for (let i = 0; i < 1000; i++) {
      const m = monsters[i];
      if (m.alive && m.left === p.left && m.top === p.top) {
        battleCheck(PLAYER_CONST, i, LAND_NUMBER);
        p.left = p.oleft; p.top = p.otop;
      }
    }
    p.oleft = p.left; p.otop = p.top;

  } else {
    const m = monsters[character];
    if (!m.alive) return;

    const mi = rint(m.left / SC);
    const mj = rint(m.top / SC);

    if (m.left === player.left && m.top === player.top && player.alive) {
      m.left = m.oleft; m.top = m.otop;
      battleCheck(character, PLAYER_CONST, LAND_NUMBER);
    } else if (landsquare[mi][mj].condition !== T.Room) {
      if (landsquare[mi][mj].condition === T.Wall && m.ability === A.WallBreak
          && m.left !== 0 && m.top !== 0
          && m.left !== (LAND_NUMBER-1)*SC && m.top !== (LAND_NUMBER-1)*SC) {
        battleCheck(character, mi, mj);
      } else if (landsquare[mi][mj].condition <= T.PurpleBox && m.ability === A.Boxattack) {
        battleCheck(character, mi, mj);
      }

      // try alternate movement based on direction
      const oi = rint(m.oleft / SC);
      const oj = rint(m.otop / SC);
      const dir = m.direction;
      const tryAlt = (dx1, dy1, dx2, dy2) => {
        const safe = (x, y) => x >= 0 && x < LAND_NUMBER && y >= 0 && y < LAND_NUMBER;
        if (safe(oi+dx1, oj+dy1) && landsquare[oi+dx1][oj+dy1].condition === T.Room) {
          m.left = m.oleft + dx1*SC; m.top = m.otop + dy1*SC;
        } else if (safe(oi+dx2, oj+dy2) && landsquare[oi+dx2][oj+dy2].condition === T.Room) {
          m.left = m.oleft + dx2*SC; m.top = m.otop + dy2*SC;
        } else {
          m.left = m.oleft; m.top = m.otop;
        }
      };
      if (dir === 2) tryAlt(-1,-1, 1,-1);
      else if (dir === 3) tryAlt(1,1, -1,1);
      else if (dir === 5) tryAlt(1,-1, 1,1);
      else if (dir === 7) tryAlt(-1,1, -1,-1);
      else if (dir === 10) tryAlt(0,-1, 1,0);
      else if (dir === 15) tryAlt(1,0, 0,1);
      else if (dir === 21) tryAlt(0,1, -1,0);
      else if (dir === 14) tryAlt(-1,0, 0,-1);
      else { m.left = m.oleft; m.top = m.otop; }
    }

    const ni = rint(m.left / SC);
    const nj = rint(m.top / SC);
    landsquare[rint(m.oleft/SC)][rint(m.otop/SC)].condition = T.Room;
    landsquare[ni][nj].condition = T.Enemy;
    m.oleft = m.left; m.otop = m.top;
    // perf: no tileLayerDirty here — Room/Enemy tiles render identically, and
    // real tile changes (box/wall destruction) set the flag where they happen.
  }
}

// ── movement ──────────────────────────────────────────────────────────────────

export function movement() {
  const { player, monsters } = state;
  const gm = state.gameMode;

  if (gm === GameMode.Entrance) {
    player.direction = 1;

  } else if (gm === GameMode.Dungeon) {
    if (player.direction > 1) {
      const p = player;
      const maxPos = (LAND_NUMBER - 1) * SC;

      if (p.direction % 2 === 0 && p.top > 0)      p.top -= SC;
      if (p.direction % 3 === 0 && p.top < maxPos)  p.top += SC;
      if (p.direction % 5 === 0 && p.left < maxPos) p.left += SC;
      if (p.direction % 7 === 0 && p.left > 0)      p.left -= SC;

      positionCheck(PLAYER_CONST); // perf: movement never changes tile visuals; no blanket tileLayerDirty

      if (p.condition !== A.Slow) {
        p.direction = 1; state.turn--;
      } else if (p.condition === A.Slow && state.turn % 2 === 0) {
        p.direction = 11; state.turn--;
      } else {
        p.direction = 1; state.turn--;
      }

      // monster movement
      for (let i = 0; i < state.monsterNumber; i++) {
        const m = monsters[i];
        if (!m.alive) continue;
        if (player.condition === A.Stealth) continue;
        if (m.ability === A.Slow && state.turn % 2 !== 0) continue;

        if (m.top > player.top)  { m.top -= SC;  m.direction *= 2; }
        if (m.top < player.top)  { m.top += SC;  m.direction *= 3; }
        if (m.left < player.left){ m.left += SC; m.direction *= 5; }
        if (m.left > player.left){ m.left -= SC; m.direction *= 7; }

        positionCheck(i);
        m.direction = 1;
      }

      if (state.sumDamage > 0) {
        state.words[3] = 'プレイヤーは' + state.sumDamage + 'ダメージを受けた';
        state.sumDamage = 0;
      }
    }

  } else if (gm === GameMode.HowtoPlay) {
    if (player.direction % 5 === 0) player.condition = (player.condition + 1) % 7;
    else if (player.direction % 7 === 0) player.condition = ((player.condition - 1) + 7) % 7;
    player.direction = 1;
    _updateHowtoPlaySprites();

  } else if (gm === GameMode.Options) {
    if (player.direction % 5 === 0) player.ability = (player.ability + 1) % 5;
    else if (player.direction % 7 === 0) player.ability = ((player.ability - 1) + 5) % 5;
    player.direction = 1;

  } else if (gm === GameMode.Museum) {
    if (player.direction % 5 === 0) player.condition = (player.condition + 1) % 6;
    else if (player.direction % 7 === 0) player.condition = ((player.condition - 1) + 6) % 6;
    player.direction = 1;
  }
}

function _updateHowtoPlaySprites() {
  const { player, landsquare, monsters } = state;
  const pg = player.condition;

  for (let i = 0; i <= 25; i++) landsquare[i][0].alive = false;

  if (pg === 0) {
    // nothing shown

  } else if (pg === 1) {
    for (let i = 0; i <= 25; i++) {
      const ls = landsquare[i][0];
      ls.alive = true;
      if (i <= 4) {
        ls.left = 300 + SC*2*i; ls.top = player.top;
        ls.condition = T.Mine; ls.ability = i;
      } else if (i === 5) {
        ls.left = 300; ls.top = player.top + 76;
        ls.condition = T.Stair;
      } else if (i <= 15) {
        ls.left = 300 + SC*2*(i-10); ls.top = player.top + 266;
        ls.condition = T.Wall; ls.ability = i-6;
      } else {
        ls.left = 300 + SC*2*(i-20); ls.top = player.top + 304;
        ls.condition = T.Room; ls.ability = i-16;
      }
    }
    for (let i = 0; i <= 4; i++) {
      monsters[i].left = 300 + SC*2*i;
      monsters[i].top = player.top + 38;
    }

  } else if (pg >= 2 && pg <= 5) {
    for (let i = 0; i <= 4; i++) {
      const ls = landsquare[i][0];
      ls.alive = true; ls.condition = i;
      ls.left = 300 + SC*2*(i-2); ls.top = player.top - 38;
    }
    for (let i = 5; i <= 25; i++) landsquare[i][0].alive = false;
  }
}

// ── floorSet ──────────────────────────────────────────────────────────────────

export function floorSet() {
  state.floor++;
  state.randomPosition.hp = LAND_NUMBER * LAND_NUMBER;
  const { player, monsters, landsquare, words } = state;

  // balance: per-floor turn stipend — deep floors cost ~110 turns to cross but
  // box income is ~30, so without this the run always ends in turn exhaustion.
  if (state.floor >= 2) { state.turn += 150; sfx('stair'); }

  // DLC: undo the previous boss's buffs before this floor's statusCheck(true)
  // re-levels everyone, or the multipliers compound run-long. monsters[2] is a
  // never-buffed copy of the same shared progression.
  if (state.bossIdx >= 0) {
    const b = monsters[state.bossIdx], t = monsters[2];
    b.maxHp = t.maxHp; b.hp = t.hp; b.atk = t.atk; b.def = t.def;
    b.exp = t.exp; b.level = t.level;
    state.bossIdx = -1;
  }
  if (state.floor > state.records.deepestFloor) state.records.deepestFloor = state.floor;

  if (state.floor === 1) {
    for (let k = 0; k < words.length; k++) words[k] = '';
    player.hp = 30; player.maxHp = 30; player.level = 1; player.exp = 0; // balance: was 20 (early swarm deaths)
    player.height = SC; player.width = SC;
    player.left = 0; player.top = 0;
    player.atk = 10; player.def = 4; // balance: def was 3
    player.direction = 1; player.condition = 0;
    state.turn = 200;
    state.monsterNumber = 10;
    state.records.runs++;
  } else if ([20,40,60,80,100].includes(state.floor)) {
    state.monsterNumber += 5;
  } else if ([200,300,400,500,600,700,800].includes(state.floor)) {
    state.monsterNumber += 5; // balance: was +50 (late floors had 235-385 monsters)
  } else if (state.floor === 900) {
    state.monsterNumber = 100; // balance: was 1000
  } else if (state.floor === 1000) {
    state.records.clears++;
    player.direction = 1; player.alive = false;
    state.gameMode = GameMode.GameClear;
    landSet(); wordSet(); return;
  }

  // DLC: depth-themed monster names (band changes every 100 floors)
  const band = BANDS[bandIndex(state.floor)];
  for (let i = 0; i < 1000; i++) monsters[i].name = band.name + i;

  mapSet();
  state.tileLayerDirty = true;

  // place player
  let placed = false;
  let tries = 0;
  while (!placed && tries++ < LAND_NUMBER*LAND_NUMBER) {
    const px = rint(rnd() * LAND_NUMBER) * SC;
    const py = rint(rnd() * LAND_NUMBER) * SC;
    if (landsquare[rint(px/SC)][rint(py/SC)].condition === T.Room) {
      player.left = px; player.top = py;
      player.oleft = px; player.otop = py;
      player.alive = true;
      landsquare[rint(px/SC)][rint(py/SC)].alive = false;
      state.randomPosition.hp--;
      player.condition = A.Noability;
      placed = true;
    }
  }

  // place stair
  tries = 0;
  placed = false;
  while (!placed && tries++ < LAND_NUMBER*LAND_NUMBER*4) {
    const sx = rint(rnd() * LAND_NUMBER) * SC;
    const sy = rint(rnd() * LAND_NUMBER) * SC;
    const si = rint(sx/SC), sj = rint(sy/SC);
    if (landsquare[si][sj].alive && landsquare[si][sj].condition === T.Room) {
      landsquare[si][sj].condition = T.Stair;
      state.randomPosition.hp--;
      placed = true;
    }
  }

  // place purple box
  tries = 0;
  let pboxPlaced = false;
  let pboxTries = 0;
  while (!pboxPlaced && pboxTries < LAND_NUMBER*LAND_NUMBER) {
    const bx = rint(rnd() * LAND_NUMBER) * SC;
    const by = rint(rnd() * LAND_NUMBER) * SC;
    const bi = rint(bx/SC), bj = rint(by/SC);
    if (landsquare[bi][bj].alive && landsquare[bi][bj].condition === T.Room) {
      landsquare[bi][bj].condition = T.PurpleBox;
      state.randomPosition.hp--;
      pboxPlaced = true;
    }
    pboxTries++;
    if (pboxTries > LAND_NUMBER*LAND_NUMBER) break;
  }

  // place monsters
  for (let k = 0; k < 1000; k++) {
    const m = monsters[k];
    if (k >= state.monsterNumber) {
      m.alive = false; m.condition = 0; continue;
    }
    let mplaced = false;
    let mt = 0;
    while (!mplaced && mt++ < LAND_NUMBER*LAND_NUMBER*4) {
      const mx = rint(rnd() * LAND_NUMBER) * SC;
      const my = rint(rnd() * LAND_NUMBER) * SC;
      const mi = rint(mx/SC), mj = rint(my/SC);
      if (landsquare[mi][mj].alive && landsquare[mi][mj].condition === T.Room) {
        m.alive = true; m.hp = m.maxHp;
        m.left = mx; m.top = my; m.oleft = mx; m.otop = my;
        landsquare[mi][mj].condition = T.Enemy;
        state.randomPosition.hp--;
        m.condition = 0;
        mplaced = true;
      }
      if (state.randomPosition.hp <= 0) { m.alive = false; break; }
    }
    if (!mplaced) m.alive = false;
  }

  if (state.floor >= 2) statusCheck(true);

  // DLC: every 100th floor is guarded by a boss. Buffed AFTER statusCheck(true)
  // so the per-floor levelup doesn't compound the boss multipliers.
  if (state.floor >= 100 && state.floor % 100 === 0 && monsters[1].alive) {
    const b = monsters[1];
    state.bossIdx = 1;
    b.maxHp = Math.min(rint(b.maxHp * 6), 1e6 - 1); b.hp = b.maxHp;
    b.atk = Math.min(rint(b.atk * 1.5), 3e5 - 1);
    b.exp = Math.min(b.exp * 15, 1e4 - 1);
    b.name = band.boss;
    words[0] = '⚠ ボスフロア! 守護者「' + band.boss + '」が待ち構えている!';
    sfx('boss');
  }
  state.tileLayerDirty = true;
}

// ── boxEffect ─────────────────────────────────────────────────────────────────

export function boxEffect(i, j) {
  const { player, monsters, landsquare, abilityHp } = state;
  const ls = landsquare[i][j];
  const ab = ls.ability;

  if (ls.condition === T.BlueBox || ls.condition === T.GreenBox || ls.condition === T.PurpleBox) sfx('pickup');
  else if (ls.condition === T.RedBox) sfx('bad');

  switch (ls.condition) {
    case T.BlueBox:
      player.hp += rint(player.maxHp / 4); // balance: was /10 (early attrition deaths before heal charges exist)
      for (let x = 0; x < 1000; x++) monsters[x].ability = A.Noability;
      ls.explanation = 'HPが少し回復し、モンスターの状態異常が解除された。';
      break;

    case T.RedBox:
      switch (ab) {
        case 0:
          // fun: iter8 — 〜50Fは非即死化(Hp1→maxHpの25%)。序盤の理不尽な即死だけを削り、
          // 箱そのもの・破壊ターン収入・瀕死体験(緊張感)は残す。rnd()非消費。
          if (state.floor <= 50) { player.hp = Math.max(rint(player.maxHp * 0.25), 1); ls.explanation = 'Hpが大きく減った'; }
          else { player.hp = 1; ls.explanation = 'Hpが1になってしまった'; }
          break;
        case 1: player.atk = rint(player.atk * 0.9); ls.explanation = '攻撃力が下がった'; break;
        case 2: player.condition = A.Slow; ls.explanation = 'プレイヤーはこのフロアにいる間、動きが遅くなった。'; break;
        case 3: player.condition = TURN_CONST; ls.explanation = 'このフロアにいる間、ターン数が減らなくなった。'; break;
        case 4: player.maxHp = rint(player.maxHp * 0.9); ls.explanation = '最大Hpが下がった'; break;
        case 5: player.hp = rint(player.hp / 2) + 1; ls.explanation = 'Hpが半分になった'; break;
        case 6:
          for (let x = 0; x < 1000; x++) if (monsters[x].alive) { monsters[x].hp = monsters[x].maxHp; monsters[x].condition = 0; }
          ls.explanation = 'モンスターが全回復した。';
          break;
        case 7:
          for (let x = 0; x < 1000; x++) monsters[x].atk = rint(monsters[x].atk * 1.1);
          ls.explanation = '全てのモンスターの攻撃力が上がった';
          break;
        case 8: statusCheck(true); ls.explanation = '全てのモンスターのレベルが上がった'; break;
        case 9:
          for (let x = 0; x < LAND_NUMBER; x++) for (let y = 0; y < LAND_NUMBER; y++)
            if (landsquare[x][y].condition <= T.PurpleBox) landsquare[x][y].def *= 2;
          ls.explanation = 'このフロアの全ての箱の守備力が2倍になった。';
          break;
        case 10:
          for (let x = 0; x < 1000; x++) monsters[x].def = rint(monsters[x].def * 1.1);
          ls.explanation = '全てのモンスターの守備力が上がった';
          break;
        case 11:
          for (let x = 0; x < LAND_NUMBER; x++) for (let y = 0; y < LAND_NUMBER; y++)
            if (landsquare[x][y].condition <= T.GreenBox) landsquare[x][y].condition = T.RedBox;
          ls.explanation = '全ての箱が赤色になった'; state.tileLayerDirty = true;
          break;
        case 12:
          for (let x = 0; x < 1000; x++) monsters[x].maxHp = rint(monsters[x].maxHp * 1.1);
          ls.explanation = '全てのモンスターの最大Hpが上がった';
          break;
        case 13:
          for (let x = 0; x < 1000; x++) monsters[x].exp = rint(monsters[x].exp * 0.9);
          ls.explanation = '全てのモンスターの経験値が下がった';
          break;
        case 14:
          for (let x = 0; x < LAND_NUMBER; x++) for (let y = 0; y < LAND_NUMBER; y++)
            if (landsquare[x][y].condition <= T.PurpleBox) { landsquare[x][y].hp *= 2; landsquare[x][y].maxHp *= 2; }
          ls.explanation = 'このフロアの全ての箱のHpが2倍になった。';
          break;
      }
      break;

    case T.YellowBox:
      switch (ab) {
        case 0:
          for (let x = 0; x < 1000; x++) monsters[x].ability = A.Stealth;
          ls.explanation = 'モンスターが透明になった。';
          break;
        case 1:
          for (let x = 0; x < 6; x++) abilityHp[x]++;
          ls.explanation = '全てのコマンドの使用回数を1増やした。';
          break;
        case 2: player.condition = TURN_CONST; ls.explanation = 'このフロアにいる間、ターン数が減らなくなった。'; break;
        case 3:
          for (let x = 0; x < state.monsterNumber; x++) if (monsters[x].alive) { monsters[x].alive=false; monsters[x].hp=0; monsters[x].condition=0; landsquare[rint(monsters[x].left/SC)][rint(monsters[x].top/SC)].condition=T.Room; }
          ls.explanation = 'モンスターが全て死んだ。'; state.tileLayerDirty = true;
          break;
        case 4: player.def *= 2; ls.explanation = '守備力が2倍になった'; break;
        case 5: player.atk *= 2; ls.explanation = '攻撃力が2倍になった'; break;
        case 6: state.turn = 1000; ls.explanation = 'ターンが残り1000になった。'; break;
        case 7: player.atk = rint(player.atk / 2); ls.explanation = '攻撃力が半分になった'; break;
        case 8:
          for (let x = 0; x < 1000; x++) monsters[x].ability = A.WallBreak;
          ls.explanation = 'モンスターが壁を通れるようになった。(端の壁は除く)';
          break;
        case 9:
          for (let x = 0; x < 6; x++) abilityHp[x] = 5;
          ls.explanation = '全てのコマンドの使用回数が5になった。';
          break;
        case 10:
          for (let x = 0; x < LAND_NUMBER; x++) for (let y = 0; y < LAND_NUMBER; y++)
            if (landsquare[x][y].condition <= T.GreenBox) landsquare[x][y].condition = T.Room;
          ls.explanation = '全ての箱が消えた。'; state.tileLayerDirty = true;
          break;
        case 11:
          for (let x = 0; x < LAND_NUMBER; x++) for (let y = 0; y < LAND_NUMBER; y++)
            if (landsquare[x][y].condition <= T.GreenBox) landsquare[x][y].condition = T.BlueBox;
          ls.explanation = '全ての箱が青色になった'; state.tileLayerDirty = true;
          break;
        case 12:
          for (let x = 0; x < LAND_NUMBER; x++) for (let y = 0; y < LAND_NUMBER; y++)
            if (landsquare[x][y].condition <= T.GreenBox) landsquare[x][y].condition = rint(rnd()*4);
          ls.explanation = '全ての箱がランダムに変化した。'; state.tileLayerDirty = true;
          break;
        case 13:
          for (let x = 0; x < LAND_NUMBER; x++) for (let y = 0; y < LAND_NUMBER; y++)
            if (landsquare[x][y].condition <= T.GreenBox) landsquare[x][y].condition = T.YellowBox;
          ls.explanation = '全ての箱が黄色になった'; state.tileLayerDirty = true;
          break;
        case 14:
          for (let x = 0; x < 1000; x++) monsters[x].ability = A.Boxattack;
          ls.explanation = 'モンスターが箱を壊せるようになった。';
          break;
      }
      break;

    case T.GreenBox:
      switch (ab) {
        case 0: player.hp = player.maxHp; ls.explanation = 'Hpが全回復した'; break;
        case 1: player.atk = rint(player.atk * 1.1) + 1; ls.explanation = '攻撃力が上がった'; break;
        case 2:
          for (let x = 0; x < state.monsterNumber; x++) if (monsters[x].alive) { monsters[x].alive=false; monsters[x].hp=0; monsters[x].condition=0; landsquare[rint(monsters[x].left/SC)][rint(monsters[x].top/SC)].condition=T.Room; }
          ls.explanation = 'モンスターが全て死んだ。'; state.tileLayerDirty = true;
          break;
        case 3:
          for (let x = 0; x < 1000; x++) monsters[x].ability = A.Slow;
          ls.explanation = 'モンスターの動きが遅くなった。';
          break;
        case 4: player.maxHp = rint(player.maxHp * 1.1) + 1; ls.explanation = '最大Hpが上がった'; break;
        case 5:
          for (let x = 1; x < LAND_NUMBER-1; x++) for (let y = 1; y < LAND_NUMBER-1; y++)
            if (landsquare[x][y].condition === T.Wall) landsquare[x][y].condition = T.Room;
          ls.explanation = '端の壁以外の壁が全て壊れた。'; state.tileLayerDirty = true;
          break;
        case 6: player.condition = A.Stealth; ls.explanation = 'プレイヤーは透明になった。(このフロアのみ)'; break;
        case 7:
          for (let x = 0; x < 1000; x++) monsters[x].hp = 1;
          ls.explanation = '全てのモンスターのHpが残り1になった';
          break;
        case 8:
          for (let x = 0; x < 1000; x++) monsters[x].atk = rint(monsters[x].atk * 0.9);
          ls.explanation = '全てのモンスターの攻撃力が下がった';
          break;
        case 9:
          for (let x = 0; x < 1000; x++) monsters[x].exp += 5;
          ls.explanation = '全てのモンスターの経験値が上がった';
          break;
        case 10:
          for (let x = 0; x < 1000; x++) monsters[x].maxHp = rint(monsters[x].maxHp * 0.9);
          ls.explanation = '全てのモンスターの最大Hpが下がった';
          break;
        case 11:
          for (let x = 0; x < LAND_NUMBER; x++) for (let y = 0; y < LAND_NUMBER; y++)
            if (landsquare[x][y].condition <= T.PurpleBox) landsquare[x][y].hp = 1;
          ls.explanation = '全ての箱のHpが残り1になった。';
          break;
        case 12:
          for (let x = 0; x < LAND_NUMBER; x++) for (let y = 0; y < LAND_NUMBER; y++)
            if (landsquare[x][y].condition === T.RedBox) landsquare[x][y].condition = T.Room;
          ls.explanation = '赤箱が消えた。'; state.tileLayerDirty = true;
          break;
        case 13: player.def++; ls.explanation = '守備力が上がった'; break;
        case 14: player.condition = A.WallBreak; ls.explanation = 'このフロアにいる間、壁を通れるようになった(端の壁は除く)'; break;
      }
      break;

    case T.PurpleBox: {
      const y = rint(rnd() * (5 - player.ability)) + 1;
      if (ab <= 2)       { abilityHp[0] += y; ls.explanation = '全体攻撃の使用回数を' + y + '増やした。'; }
      else if (ab <= 6)  { abilityHp[1] += y; ls.explanation = 'Hp全回復の使用回数を' + y + '増やした。'; }
      else if (ab === 7) { abilityHp[2] += y; ls.explanation = '全消去の使用回数を' + y + '増やした。'; }
      else if (ab === 8) { abilityHp[3] += y; ls.explanation = 'モンスター除去の使用回数を' + y + '増やした。'; }
      else if (ab <= 10) { abilityHp[4] += y; ls.explanation = '次の階へ行く術の使用回数を' + y + '増やした。'; } // balance: was ab<=11 (蘇生 3/15→4/15)
      else               { abilityHp[5] += y; ls.explanation = '蘇りの術が' + y + '回にまで増えた。'; }
      break;
    }

    case T.Wall:
      ls.explanation = '';
      break;
  }

  statusCheck(false);
  state.words[1] = ls.explanation;
  ls.ability = 0;
  ls.condition = T.Room;
  state.tileLayerDirty = true;
}

// ── abilityEffect ─────────────────────────────────────────────────────────────

// fun: 遠隔破壊(全体攻撃/全消去)した青/緑/紫箱から得る弱体化ボーナス。
// 直接壊す場合(フル効果+ターン+15)より恩恵を絞り、テンポと引き換えにする。
// 青=回復maxHp/16(通常はmaxHp/4)・緑=atk+1%(通常は抽選で全回復やx1.1等)・
// 紫=コマンド1回分のみ(通常1〜5回、蘇生は対象外)。ターンボーナスなし。
// 1回の使用で吸収できるのはREMOTE_BONUS_CAP個まで(初回計測CLEAR80%の
// ガードレール超過を受けた調整。超過分の箱は壊れるだけ)。
const REMOTE_BONUS_CAP = 3;
function remoteBoxBonus(cond) {
  const { player, abilityHp } = state;
  if (cond === T.BlueBox) player.hp += rint(player.maxHp / 16);
  else if (cond === T.GreenBox) player.atk = rint(player.atk * 1.01) + 1;
  else if (cond === T.PurpleBox) abilityHp[rint(rnd() * 5)]++;
}

// fun: 遠隔破壊された箱のターン収入は通常破壊(+15)の1/3。ゼロにすると
// メレー農業で成立していたターン経済に穴が開き、ターン切れ死が増える
// (iter9b: 13→17件, CLEAR40%)ことが計測で判明したための補償。個数無制限。
function remoteBoxTurn() {
  if (state.player.condition !== TURN_CONST) state.turn += 5;
  state.records.totalBoxes++;
}

export function abilityEffect(key) {
  const { player, monsters, landsquare, abilityHp } = state;

  if (key === 'KeyA') {
    for (let i = 0; i < LAND_NUMBER; i++)
      for (let j = 0; j < LAND_NUMBER; j++) {
        if (landsquare[i][j].condition === T.GreenBox) {
          landsquare[i][j].ability = rint(rnd() * 15);
          if (player.condition !== TURN_CONST) state.turn += 10;
          boxEffect(i, j);
        }
      }
    floorSet();

  } else if (key === 'KeyZ') {
    if (abilityHp[0] > 0) {
      abilityHp[0]--;
      const roll = rnd() * 0.2 + 0.9;
      const dmg = rint(player.atk * roll / (monsters[0].def || 1));
      const cappedDmg = Math.min(dmg, 1e7-1);
      for (let i = 0; i < state.monsterNumber; i++) {
        const m = monsters[i];
        if (m.alive) {
          m.hp -= cappedDmg;
          if (m.hp <= 0) {
            m.hp = 0; m.alive = false;
            _monsterKilled(i);
          }
        }
      }
      // fun: 全体攻撃は青/緑/紫の箱にも届く(テンポ改善)。壊れた箱は
      // remoteBoxBonusの弱体化ボーナスのみで、フル効果もターン+15もない。
      let boxBroken = 0;
      for (let i = 0; i < LAND_NUMBER; i++)
        for (let j = 0; j < LAND_NUMBER; j++) {
          const ls = landsquare[i][j];
          const c = ls.condition;
          if (c !== T.BlueBox && c !== T.GreenBox && c !== T.PurpleBox) continue;
          ls.hp -= Math.min(rint(player.atk * roll / Math.max(1, ls.def)) + 1, 1e7-1);
          if (ls.hp <= 0) {
            ls.hp = 0; ls.condition = T.Room; ls.ability = 0;
            if (boxBroken < REMOTE_BONUS_CAP) remoteBoxBonus(c);
            remoteBoxTurn();
            boxBroken++;
          }
        }
      state.words[2] = '全てのモンスターに' + cappedDmg + 'ダメージを与えた'
        + (boxBroken ? '。箱' + boxBroken + '個を砕き、微かな力を得た' : '');
      player.direction = 11;
      state.tileLayerDirty = true;
    }

  } else if (key === 'KeyX') {
    if (abilityHp[1] > 0) { abilityHp[1]--; player.hp = player.maxHp; }

  } else if (key === 'KeyC') {
    if (abilityHp[2] > 0) {
      abilityHp[2]--;
      // fun: 全消去は紫箱も消去対象に加え、青/緑/紫の箱からは
      // remoteBoxBonusの弱体化ボーナスを取り込む(テンポ改善)。
      let absorbed = 0;
      for (let i = 0; i < LAND_NUMBER; i++)
        for (let j = 0; j < LAND_NUMBER; j++) {
          const c = landsquare[i][j].condition;
          if (c <= T.PurpleBox || c === T.Wall) {
            if (c === T.BlueBox || c === T.GreenBox || c === T.PurpleBox) {
              if (absorbed < REMOTE_BONUS_CAP) remoteBoxBonus(c);
              remoteBoxTurn();
              absorbed++;
            }
            landsquare[i][j].condition = T.Room;
          }
        }
      for (let x = 0; x < state.monsterNumber; x++)
        if (monsters[x].alive) { monsters[x].alive=false; monsters[x].hp=0; monsters[x].condition=0; landsquare[rint(monsters[x].left/SC)][rint(monsters[x].top/SC)].condition=T.Room; }
      if (absorbed) state.words[2] = '消した箱' + absorbed + '個から微かな力を取り込んだ';
      state.tileLayerDirty = true;
    }

  } else if (key === 'KeyD') {
    if (abilityHp[3] > 0 && player.alive) {
      abilityHp[3]--;
      for (let i = 0; i < state.monsterNumber; i++) {
        const m = monsters[i];
        if (m.alive) {
          m.alive = false; m.hp = 0; m.condition = 0;
          landsquare[rint(m.left/SC)][rint(m.top/SC)].condition = rint(rnd() * 4);
        }
      }
      state.tileLayerDirty = true;
    } else if (abilityHp[5] > 0 && !player.alive) {
      abilityHp[5]--;
      player.alive = true; player.hp = player.maxHp;
      if (player.condition !== TURN_CONST) state.turn += 100;
      player.condition = A.Noability;
      state.words[3] = '蘇りの術の効果で復活した。';
      statusCheck(false);
    }

  } else if (key === 'KeyS') {
    if (abilityHp[6] === 0) {
      abilityHp[6] = 1;
      state.words[3] = '次に階段に乗ったら、セーブしてメニュー画面に戻ります。';
    } else if (abilityHp[6] > 0 && landsquare[rint(player.left/SC)][rint(player.top/SC)].condition === T.Stair) {
      statusCheck(false);
      writeSave();
      state.gameMode = GameMode.Entrance;
      stopAll();
      landSet(); wordSet();
    }

  } else if (key === 'KeyL') {
    state.gameMode = GameMode.Dungeon;
    landSet(); monsterSet();
    const save = loadSave();
    if (save) {
      for (let i = 0; i < 1000; i++) {
        const m = monsters[i];
        m.maxHp = save['m1']; m.hp = m.maxHp;
        m.atk = save['m2']; m.def = save['m3'];
        m.exp = save['m4']; m.level = save['m5'];
        m.ability = save['m6'];
      }
    }
    floorSet();

  } else if (key === 'Enter') {
    if (abilityHp[4] > 0) { abilityHp[4]--; floorSet(); }
  }
}

// ── musicCheck ────────────────────────────────────────────────────────────────

export function musicCheck() {
  const gm = state.gameMode;
  if ([GameMode.Entrance, GameMode.HowtoPlay, GameMode.Options, GameMode.Museum].includes(gm)) {
    playMusic(0);
  } else if (gm === GameMode.Dungeon) {
    playMusic(3);
  }
}

// ── save / load ───────────────────────────────────────────────────────────────

function writeSave() {
  const { player, monsters, abilityHp } = state;
  const data = {
    1: player.hp, 2: player.maxHp, 3: player.atk, 4: player.def,
    5: player.exp, 6: player.level,
    7: abilityHp[0], 8: abilityHp[1], 9: abilityHp[2],
    10: abilityHp[3], 11: abilityHp[4], 12: abilityHp[5],
    13: state.floor, 14: state.turn, 15: player.ability,
    m1: monsters[0].maxHp, m2: monsters[0].atk, m3: monsters[0].def,
    m4: monsters[0].exp, m5: monsters[0].level, m6: monsters[0].ability,
    m7: state.monsterNumber
  };
  try { localStorage.setItem('dungeon_save', JSON.stringify(data)); } catch(e) {}
}

function loadSave() {
  try {
    const raw = localStorage.getItem('dungeon_save');
    return raw ? JSON.parse(raw) : null;
  } catch(e) { return null; }
}

// DLC: persistent achievement records (independent of the run save)
function saveRecords() {
  try { localStorage.setItem('dungeon_records', JSON.stringify(state.records)); } catch(e) {}
}

export function loadRecords() {
  try {
    const raw = localStorage.getItem('dungeon_records');
    if (raw) Object.assign(state.records, JSON.parse(raw));
  } catch(e) {}
}
