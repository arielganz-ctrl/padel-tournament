// Shared tournament logic: game list, times, standings and the knockout bracket.
// Results come from the Google Sheet through Apps Script (JSONP, see apps-script/Code.gs).
(function () {
  const T = window.TOURNAMENT;

  const toMin = (t) => { const [h, m] = t.split(':').map(Number); return h * 60 + m; };
  const fmt = (m) => `${String(Math.floor(m / 60) % 24).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;

  const roundStart = (r) => toMin(T.start) + r * T.slot;
  const koStart = (stage) => roundStart(T.rounds.length) + T.koBreak + stage * T.slot;

  // group stage: every game gets a stable id from its round and court
  const games = T.rounds.flatMap((round, r) => round.map(([house, x, y], c) => ({
    id: `R${r + 1}C${c + 1}`, stage: 'group', round: r, court: c + 1, house, x, y, time: fmt(roundStart(r)),
  })));

  // knockout: side 1 = QF1 + QF2 -> SF1, side 2 = QF3 + QF4 -> SF2
  const KO = [
    { id: 'QF1', stage: 0, court: 1, name: 'רבע גמר 1', side: 1, from: [['seed', 'A', 1], ['seed', 'B', 2]] },
    { id: 'QF2', stage: 0, court: 2, name: 'רבע גמר 2', side: 1, from: [['seed', 'A', 3], ['seed', 'A', 6]] },
    { id: 'QF3', stage: 0, court: 3, name: 'רבע גמר 3', side: 2, from: [['seed', 'A', 2], ['seed', 'B', 1]] },
    { id: 'QF4', stage: 0, court: 4, name: 'רבע גמר 4', side: 2, from: [['seed', 'A', 4], ['seed', 'A', 5]] },
    { id: 'SF1', stage: 1, court: 1, name: 'חצי גמר 1', side: 1, from: [['win', 'QF1'], ['win', 'QF2']] },
    { id: 'SF2', stage: 1, court: 2, name: 'חצי גמר 2', side: 2, from: [['win', 'QF3'], ['win', 'QF4']] },
    { id: 'F', stage: 2, court: 1, name: 'גמר 🏆', from: [['win', 'SF1'], ['win', 'SF2']] },
    { id: 'P3', stage: 2, court: 2, name: 'מקום 3', from: [['lose', 'SF1'], ['lose', 'SF2']] },
  ].map((g) => ({ ...g, time: fmt(koStart(g.stage)) }));
  const KO_STAGES = ['רבע גמר', 'חצי גמר', 'גמר ומקום 3'];
  const QUALIFY = { A: 6, B: 2 };

  const num = (v) => (v === '' || v == null || isNaN(Number(v)) ? null : Number(v));
  // a game counts once both scores are in and differ
  function score(results, id) {
    const r = results[id];
    if (!r) return null;
    const a = num(r.a), b = num(r.b);
    return a == null || b == null || a === b ? null : [a, b];
  }

  function standings(house, results) {
    const rows = new Map(T[house].map((p) => [p, { pair: p, played: 0, w: 0, l: 0, gf: 0, ga: 0 }]));
    const h2h = {};
    for (const g of games) {
      if (g.house !== house) continue;
      const s = score(results, g.id);
      if (!s) continue;
      const [a, b] = s, X = rows.get(g.x), Y = rows.get(g.y);
      X.played++; Y.played++;
      X.gf += a; X.ga += b; Y.gf += b; Y.ga += a;
      const [win, lose] = a > b ? [X, Y] : [Y, X];
      win.w++; lose.l++;
      h2h[`${win.pair}|${lose.pair}`] = true;
    }
    const list = [...rows.values()].map((r) => ({ ...r, diff: r.gf - r.ga, pts: r.w }));
    list.sort((p, q) => q.pts - p.pts || q.diff - p.diff || q.gf - p.gf || p.pair.localeCompare(q.pair, 'he'));
    // head-to-head decides between exactly two pairs level on points and game difference
    for (let i = 0; i + 1 < list.length; i++) {
      const p = list[i], q = list[i + 1];
      const tied = p.pts === q.pts && p.diff === q.diff;
      const third = (list[i + 2] && list[i + 2].pts === p.pts && list[i + 2].diff === p.diff) ||
        (list[i - 1] && list[i - 1].pts === p.pts && list[i - 1].diff === p.diff);
      if (tied && !third && h2h[`${q.pair}|${p.pair}`]) { list[i] = q; list[i + 1] = p; }
    }
    const done = games.filter((g) => g.house === house).every((g) => score(results, g.id));
    return { rows: list.map((r, i) => ({ ...r, pos: i + 1 })), done };
  }

  // who plays in each knockout game: manual override > result of earlier game > final standings
  function bracket(results) {
    const st = { A: standings('A', results), B: standings('B', results) };
    const out = {};
    const resolve = (src) => {
      if (src[0] === 'seed') {
        const [, house, pos] = src;
        return { label: `${house}${pos}`, pair: st[house].done ? st[house].rows[pos - 1].pair : null };
      }
      const prev = out[src[1]], s = score(results, src[1]);
      const label = `${src[0] === 'win' ? 'מנצחת' : 'מפסידת'} ${KO.find((k) => k.id === src[1]).name}`;
      if (!prev || !s || !prev.x.pair || !prev.y.pair) return { label, pair: null };
      const xWon = s[0] > s[1];
      return { label, pair: (src[0] === 'win') === xWon ? prev.x.pair : prev.y.pair };
    };
    for (const k of KO) {
      const r = results[k.id] || {};
      const x = resolve(k.from[0]), y = resolve(k.from[1]);
      if (r.p1) x.pair = r.p1;
      if (r.p2) y.pair = r.p2;
      out[k.id] = { ...k, x, y };
    }
    return out;
  }

  // JSONP: works even where the browser blocks reading cross-site fetch replies
  let jsonpId = 0;
  function api(params) {
    return new Promise((resolve, reject) => {
      const cb = `__t${Date.now()}_${jsonpId++}`;
      const el = document.createElement('script');
      const done = () => { clearTimeout(timer); delete window[cb]; el.remove(); };
      const timer = setTimeout(() => { done(); reject(new Error('timeout')); }, 45000);
      window[cb] = (data) => { done(); resolve(data); };
      el.onerror = () => { done(); reject(new Error('script load failed')); };
      el.src = `${T.scriptUrl}?${new URLSearchParams({ ...params, callback: cb })}`;
      document.head.append(el);
    });
  }
  const fetchResults = () => api({ action: 'results' }).then((d) => {
    if (!d.ok) throw new Error(d.error || 'results');
    return d.results || {};
  });

  window.TourneyKit = { T, games, KO, KO_STAGES, QUALIFY, fmt, roundStart, koStart, score, standings, bracket, api, fetchResults };
})();
