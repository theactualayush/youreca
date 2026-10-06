/* FF Meeting Calculator engine.
 * A line-by-line port of the "SR3-SR1" sheet and the Graphs strategy blocks of
 * FF_Meeting_Calculator_OptimizeTest.xlsx. Dates are Excel serial numbers (day 0 = 1899-12-30)
 * and may carry fractional times, exactly like the workbook. Excel errors are represented as NaN
 * so they propagate through sums, products and averages the same way. */
(function (root) {
  "use strict";
  const DAY = 864e5, OFF = 25569;
  const MN = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  /* ---------- date helpers (Excel semantics) ---------- */
  const ser = (y, m, d) => Date.UTC(y, m - 1, d) / DAY + OFF;
  const parts = x => { const t = new Date((Math.floor(x) - OFF) * DAY); return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate(), w: t.getUTCDay() }; };
  const lab = x => { const p = parts(x); return MN[p.m - 1] + String(p.y % 100).padStart(2, "0"); };
  const labYM = (y, m) => { const t = new Date(Date.UTC(y, m - 1, 1)); return MN[t.getUTCMonth()] + String(t.getUTCFullYear() % 100).padStart(2, "0"); };
  const parseLab = l => { const i = MN.indexOf(String(l).slice(0, 3)); const yy = +String(l).slice(3, 5); return i < 0 || isNaN(yy) ? null : { y: 2000 + yy, m: i + 1 }; };
  const eom = x => { const p = parts(x); return ser(p.y, p.m + 1, 0); };
  const first = x => { const p = parts(x); return ser(p.y, p.m, 1); };
  const dim = x => { const p = parts(x); return ser(p.y, p.m + 1, 0) - ser(p.y, p.m, 1) + 1; };
  const edate = (x, k) => { const p = parts(x); const last = new Date(Date.UTC(p.y, p.m - 1 + k + 1, 0)).getUTCDate(); return ser(p.y, p.m + k, Math.min(p.d, last)); };
  const thirdWed = (y, m) => { const f = ser(y, m, 1), w = parts(f).w; return f + ((3 - w + 7) % 7) + 14; };
  function fromText(s) {
    s = String(s).trim(); let m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(s))) {
      let v = ser(+m[1], +m[2], +m[3]); if (m[4]) v += (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400; return v;
    }
    if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(s))) { let y = +m[3]; if (y < 100) y += 2000; return ser(y, +m[1], +m[2]); }
    if (/^\d{5}(\.\d+)?$/.test(s)) return +s; // raw Excel serial
    return null;
  }
  const iso = x => { const p = parts(x); return `${p.y}-${String(p.m).padStart(2, "0")}-${String(p.d).padStart(2, "0")}`; };
  const isoTime = x => { const f = x - Math.floor(x); if (f < 1e-9) return iso(x); const s = Math.round(f * 86400); return `${iso(x)} ${String(Math.floor(s / 3600)).padStart(2, "0")}:${String(Math.floor(s / 60) % 60).padStart(2, "0")}`; };
  const num = v => { if (v === null || v === undefined || v === "") return undefined; const f = typeof v === "number" ? v : Number(String(v).trim()); return isNaN(f) ? NaN : f; };
  const isErr = v => typeof v === "number" && isNaN(v);

  function bisectRight(a, x) { let lo = 0, hi = a.length; while (lo < hi) { const mid = (lo + hi) >> 1; if (a[mid] <= x) lo = mid + 1; else hi = mid; } return lo; }

  /* ---------- meeting list (Raw Data AU + Graphs B42:B66) ---------- */
  function meetingList(confirmed, scenario) {
    const sc = scenario.slice().sort((a, b) => a - b);
    if (!sc.length) return confirmed.slice().sort((a, b) => a - b);
    const B = sc.slice(), A = [];
    for (let i = 0; i + 1 < B.length; i++) A.push(B[i + 1] - B[i]);
    while (B.length < 24 && A.length) {           // rolling average of the last 5 gaps (Graphs A48:A66)
      const w = A.slice(-5); A.push(w.reduce((s, v) => s + v, 0) / w.length);
      B.push(B[B.length - 1] + A[B.length - 1]);
    }
    return confirmed.filter(d => d < sc[0]).sort((a, b) => a - b).concat(B);
  }

  /* ---------- market table helpers ---------- */
  function table(rows) { const m = new Map(); for (const r of rows || []) if (!m.has(r.label)) m.set(r.label, r); return m; }
  const vw = (t, l) => { const r = t.get(l); return r ? num(r.vwap) : undefined; };
  const st = (t, l) => { const r = t.get(l); return r ? num(r.settle) : undefined; };
  const need = v => (v === undefined || v === null ? NaN : v);

  /* ---------- main calculation ---------- */
  function run(inp, scenarioDates) {
    const TODAY = inp.today;
    const zq = table(inp.mkt.ZQ_Out), sr1 = table(inp.mkt.SR1_Out), sr3 = table(inp.mkt.SR3_Out),
      sr1zq = table(inp.mkt.SR1ZQ), zqms = table(inp.mkt.ZQ_MS), zq1ms = table(inp.mkt.ZQ_1MS), sr33ms = inp.mkt.SR3_3MS || [];
    const sofrMap = new Map(inp.sofr.map(r => [r[0], r[1]])), effrMap = new Map(inp.effr.map(r => [r[0], r[1]]));
    const sofrDates = [...sofrMap.keys()].sort((a, b) => a - b);
    const lastSOFR = sofrDates[sofrDates.length - 1];
    const lastFix = Math.max(...effrMap.keys());   // rate!D2: last EFFR date, also used by the SOFR month block
    const prevSofr = d => { const i = bisectRight(sofrDates, d) - 1; return i >= 0 ? sofrMap.get(sofrDates[i]) : undefined; };
    const hol = new Set(inp.holidays.map(Math.floor));

    /* meetings */
    const AU = meetingList(inp.confirmed, scenarioDates);
    const i0 = AU.findIndex(x => x >= TODAY);
    const MC = i0 < 0 ? [] : AU.slice(i0, i0 + 12), U = i0 < 0 ? [] : AU.slice(i0, i0 + 24);

    /* weight matrix (Graphs U5:BE29) */
    const cols = [], colIdx = new Map();
    if (U.length) {
      let s = ser(parts(U[0]).y, Math.min(parts(TODAY).m, parts(U[0]).m), 1);
      for (let j = 0; j < 36; j++) { const e = eom(s); cols.push({ label: lab(s), s, e, n: Math.round(e - s + 1) }); colIdx.set(lab(s), j); s = e + 1; }
    }
    const W = (m, label) => {
      const j = colIdx.get(label); if (j === undefined) return NaN;
      const c = cols[j]; let q = (Math.floor(c.e) - Math.floor(m)) / c.n;
      if (q > 1) q = 1; if (q < 0) return 0; return -q;
    };

    /* implied cuts (meetingCuts) */
    const rows = MC.map(m => {
      const d = parts(m).d;
      const l1 = d < 15 ? lab(edate(m, -1)) : lab(m), l2 = d < 15 ? lab(m) : lab(edate(m, 1));
      const con = `${l1}-${l2} Calendar`;
      const E = zqms.has(con) ? need(vw(zqms, con)) : 0;
      return { month: lab(m), date: m, con, l1, l2, E };
    });
    const F = new Array(rows.length).fill(NaN);
    for (let i = rows.length - 1; i >= 0; i--) {
      const r = rows[i]; const den = W(r.date, r.l1) - W(r.date, r.l2);
      let f = den === 0 ? NaN : r.E / den;
      if (i < rows.length - 1) {
        const nx = rows[i + 1], l2m = parseLab(r.l2);
        if (l2m && l2m.m === parts(nx.date).m) f = f + F[i + 1] * W(nx.date, nx.l1);
      }
      F[i] = f; r.cut = f;
    }
    const cutsByMonth = new Map(); rows.forEach(r => { if (!cutsByMonth.has(r.month)) cutsByMonth.set(r.month, r); });
    const cutByDate = new Map(rows.map(r => [r.date, r.cut]));

    /* current-month SOFR estimate block (SR3-SR1!AY1:BC36) */
    const AZ2 = first(TODAY), AZ3 = eom(TODAY), m0 = lab(TODAY);
    const BA3 = 100 - (need(vw(zq, m0)) + need(vw(sr1zq, m0)) / 100);
    const BA = new Map();
    {
      const mc = cutsByMonth.get(m0);
      const meet = mc ? mc.date + 1 : 1, hike = mc ? (isErr(mc.cut) ? NaN : mc.cut) / 100 : 0;
      const total = AZ3 - AZ2 + 1, remaining = AZ3 - lastFix;
      let fixed = 0; for (let x = AZ2; x <= Math.min(AZ3, lastFix); x++) { const v = prevSofr(x); fixed += v === undefined ? 0 : v; }
      const post = meet === 1 ? 0 : Math.max(0, AZ3 - Math.max(lastFix + 1, meet) + 1);
      const pre = (BA3 * total - fixed - hike * post) / remaining;
      let prev = NaN;
      for (let d = AZ2; d <= AZ3; d++) {
        const w = parts(d).w, fd = w === 6 ? d - 1 : w === 0 ? d - 2 : d;
        let v;
        if (fd <= lastFix) v = sofrMap.has(fd) ? sofrMap.get(fd) : prev;
        else v = meet === 1 ? pre : (d < meet ? pre : pre + hike);
        BA.set(d, v); prev = v;
      }
    }

    /* live SOFR path (G on C grid, W on S grid) */
    const pathRate = d => {
      let base;
      if (sofrMap.has(d)) base = sofrMap.get(d);
      else if (d < TODAY) base = need(prevSofr(d));
      else if (BA.has(d)) base = BA.get(d);
      else { const mm = lab(d); base = 100 - (need(vw(zq, mm)) + need(vw(sr1zq, mm)) / 100); }
      const mc = cutsByMonth.get(lab(d));
      if (!mc) return base;
      const md = mc.date + 1, h = mc.cut / 100, days = dim(d), before = md - first(d), after = days - before;
      return d < md ? base - h * after / days : base + h * before / days;
    };

    /* grids */
    const S0 = TODAY - 403, N = 1009;               // workbook: 2025-08-29 .. 2028-06-02 for TODAY = 2026-10-06
    const C = [S0 + 3]; for (let x = S0 + 3; C.length < N;) { x++; const w = parts(x).w; if (w !== 0 && w !== 6 && !hol.has(x)) C.push(x); }
    const S = []; for (let k = 0; k < N; k++) S.push(S0 + k);
    const L = C.map((d, i) => (i + 1 < C.length ? C[i + 1] - d : NaN));
    let e0 = 4.33; { let best = -Infinity; for (const [d, v] of effrMap) if (d <= S0 && d > best) { best = d; e0 = v; } }
    const carry = dates => { let prev = e0; return dates.map(d => { const v = effrMap.has(d) ? effrMap.get(d) : prev; prev = v; return v; }); };
    const Dc = carry(C), T = carry(S);
    const G = C.map(pathRate), Wp = S.map(pathRate);

    /* case inputs (Graphs B32:D40) */
    // "auto" Case 1 cells split Total Hikes (P8) with front-loaded weights n..1 (Graphs C34:C39)
    const P8 = totalHikes(inp, { rows });
    const raw = [];
    (inp.cases.rows || []).forEach((r, k) => { if (k < MC.length) raw.push({ date: MC[k], c: [r.c1, r.c2] }); });
    (inp.cases.extra || []).forEach(r => { if (r.date != null) raw.push({ date: r.date, c: [r.c1, r.c2] }); });
    const autos = []; raw.forEach(r => r.c.forEach((v, j) => { if (v === "auto") autos.push([r, j]); }));
    const wsum = autos.length * (autos.length + 1) / 2;
    autos.forEach(([r, j], i) => { r.c[j] = P8 * (autos.length - i) / wsum; });
    const caseRows = raw;
    const caseVal = (d, k) => { for (const r of caseRows) if (r.date === d) return r.c[k]; return undefined; };
    const cum = vals => { let s = 0; return vals.map(v => { if (v !== undefined) s += v; return s; }); };
    const adj = (grid, k) => grid.map(d => { const v = caseVal(d - 1, k); return v === undefined || v === null || v === "" ? undefined : -num(v) / 100; });
    const Vs = S.map((d, i) => { if (i === 0) return undefined; const c = cutByDate.get(S[i - 1]); return c === undefined ? undefined : c / 100; });
    const cumV = cum(Vs), AD = S.map((d, i) => 100 - (T[i] + cumV[i]));
    const basis = d => need(vw(sr1zq, lab(d))) / 100;
    const caseI = k => { const ch = cum(adj(C, k)); return C.map((d, i) => d <= lastSOFR ? G[i] : Dc[i] - ch[i] - basis(d)); };
    const caseY = k => { const cx = cum(adj(S, k)); return S.map((d, i) => d < TODAY ? Wp[i] : T[i] - cx[i] - basis(d)); };
    const caseAB = k => { const cx = cum(adj(S, k)); return S.map((d, i) => d <= lastSOFR ? T[i] : T[i] - cx[i]); };
    const I = [caseI(0), caseI(1)], Y = [caseY(0), caseY(1)], AB = [caseAB(0), caseAB(1)];

    /* contract fair values (SR3SR1_Fair, SR3_Cases, SR1_Cases, ZQ_Cases) */
    const sr3fair = (rates, y, m) => {
      const s = thirdWed(y, m), en = thirdWed(y + Math.floor((m + 2) / 12), ((m + 2) % 12) + 1) - 1;
      let p = 1, sl = 0;
      for (let i = 0; i < C.length; i++) if (C[i] >= s && C[i] <= en) { p *= 1 + rates[i] * L[i] / 36000; sl += L[i]; }
      return sl === 0 ? NaN : 100 - (p - 1) * 100 * 360 / sl;
    };
    const mavg = (vals, y, m) => {
      const s = ser(y, m, 1), e = ser(y, m + 1, 1); let t = 0, n = 0;
      for (let i = 0; i < S.length; i++) if (S[i] >= s && S[i] < e) { t += vals[i]; n++; }
      return n ? t / n : NaN;
    };
    const fair = new Map(), months = [];
    const p0 = parts(C[0]);
    for (let k = 0; k < 36; k++) {
      const y = p0.y + Math.floor((p0.m - 1 + k) / 12), m = ((p0.m - 1 + k) % 12) + 1, l = labYM(y, m);
      months.push(l);
      fair.set(l, {
        sr3: sr3fair(G, y, m), sr1: 100 - mavg(Wp, y, m), zq: mavg(AD, y, m),
        sr3c: [sr3fair(I[0], y, m), sr3fair(I[1], y, m)],
        sr1c: [100 - mavg(Y[0], y, m), 100 - mavg(Y[1], y, m)],
        zqc: [100 - mavg(AB[0], y, m), 100 - mavg(AB[1], y, m)]
      });
    }
    const F_ = (l, key, k) => { const r = fair.get(l); if (!r) return NaN; return k === undefined ? r[key] : r[key][k]; };

    /* Meeting-neutral groups (Raw Data BA4:BJ42, single-meeting case) */
    const pm = parts(TODAY), qEnd = ser(pm.y, pm.m - ((pm.m - 1) % 3), 0);  // EOMONTH(TODAY,-MOD(MONTH-1,3)-1)
    const after = AU.filter(x => x > qEnd).slice(1, 7);                     // index starts at BH5: first meeting skipped
    const mn = after.map(m => {
      const d = parts(m).d, D = dim(m);
      return { meeting: m, legs: [lab(edate(m, -1)), lab(m), lab(edate(m, 1))], ratios: [d, -D, D - d] };
    });

    return {
      AU, MC, rows, cols, fair, months, mn, BA3, lastSOFR, lastFix, caseRows, P8,
      src: {
        zqLive: l => need(vw(zq, l)), zqSettle: l => need(st(zq, l)),
        sr3Live: l => need(vw(sr3, l)), sr3Settle: l => need(st(sr3, l)),
        sr1Live: l => need(vw(sr1, l)), sr1Settle: l => need(st(sr1, l)),
        zq1ms: l => need(vw(zq1ms, l)),
        zqFair: l => F_(l, "zq"), sr3Fair: l => F_(l, "sr3"), sr1Fair: l => F_(l, "sr1"),
        zqCase: k => l => F_(l, "zqc", k), sr3Case: k => l => F_(l, "sr3c", k), sr1Case: k => l => F_(l, "sr1c", k)
      },
      path: { C, S, G, Wp, AD }
    };
  }

  /* Graphs P8 "Total Hikes 2027": SR3 3M spreads up to the first negative, less implied cuts before a cutoff */
  function totalHikes(inp, res) {
    const v = (inp.mkt.SR3_3MS || []).map(r => num(r.vwap));
    const k = v.findIndex(x => x !== undefined && x < 0);
    if (k < 0) return NaN;
    let s = 0; for (let i = 0; i < k; i++) s += v[i] === undefined ? 0 : v[i];
    for (const r of res.rows) if (r.date < inp.cases.cutoff) s -= r.cut;
    return s;
  }

  /* ---------- strategy formulas (Graphs) ---------- */
  function comb(n, k) { let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return r; }
  function binom(label, f) {
    const legs = label.split("-").map(s => s.trim()); const n = legs.length;
    if (n === 1) return 100 * f(legs[0]);
    let s = 0; legs.forEach((l, k) => { s += (k % 2 ? -1 : 1) * comb(n - 1, k) * f(l); }); return 100 * s;
  }
  const legsOf = label => label.split("-").map(s => s.trim());
  const flyX = (label, f3, fo) => { const l = legsOf(label); return 100 * (2 * f3(l[0]) - fo(l[1]) - fo(l[2])); };
  const four = (label, f3, fz) => { const l = legsOf(label); return 100 * (f3(l[0]) - f3(l[1]) - fz(l[2]) + fz(l[3])); };
  const ratioSum = (legs, ratios, f) => legs.reduce((s, l, i) => s + ratios[i] * f(l), 0);

  root.FFEngine = { run, totalHikes, binom, flyX, four, ratioSum, legsOf, meetingList,
    util: { ser, parts, lab, labYM, parseLab, fromText, iso, isoTime, num, isErr, thirdWed, dim, edate } };
})(typeof window !== "undefined" ? window : globalThis);
