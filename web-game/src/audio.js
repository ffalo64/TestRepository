// BGM。VB6 版の ImscLib12（mciSendString）を HTMLAudioElement に置き換え。
// エンジンが g.music / g.musicToken を更新し、ここが毎フレームそれに合わせる。

const BASE = import.meta.env.BASE_URL;
const FILES = ['n31.mp3', 'n11.mp3', 'c26.mp3', 'c1.mp3']; // メニュー / クリア / ゲームオーバー / ダンジョン

export function createAudio() {
  const tracks = FILES.map((f) => {
    const a = new Audio(`${BASE}sounds/${f}`);
    a.preload = 'auto';
    a.volume = 0.6;
    return a;
  });
  let current = -1;
  let token = -1;
  let unlocked = false;
  let muted = false;

  function stopAll() {
    for (const a of tracks) {
      a.pause();
      a.currentTime = 0;
    }
  }

  function start() {
    if (!unlocked || muted || current < 0) return;
    tracks[current].play().catch(() => {});
  }

  return {
    // ブラウザの自動再生制限: 最初のキー入力/クリックで呼ぶ
    unlock() {
      if (unlocked) return;
      unlocked = true;
      start();
    },
    toggleMute() {
      muted = !muted;
      if (muted) stopAll();
      else start();
      return muted;
    },
    sync(g) {
      if (g.music === current && g.musicToken === token) return;
      stopAll();
      current = g.music;
      token = g.musicToken;
      if (current >= 0) {
        tracks[current].loop = g.musicLoop;
        start();
      }
    },
  };
}
