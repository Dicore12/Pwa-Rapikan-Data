/* Rapikan Data — logika inti (tanpa DOM, bisa diuji di Node) */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.RD = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const EPOCH = Date.UTC(1899, 11, 30);
  const isDate = v => v instanceof Date && !isNaN(v.getTime());
  const isEmpty = v => v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  const pad = n => (n < 10 ? '0' : '') + n;
  const round2 = x => Math.round(x * 100) / 100;

  /* ---------- tanggal ---------- */
  function serialToDate(s) { return new Date(Math.round((EPOCH + s * 86400000) / 1000) * 1000); }
  function dateToSerial(d) { return (d.getTime() - EPOCH) / 86400000; }
  function hasTime(d) { return !!(d.getUTCHours() || d.getUTCMinutes() || d.getUTCSeconds()); }
  function fmtDate(d) {
    let out = d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
    if (hasTime(d)) out += ' ' + pad(d.getUTCHours()) + ':' + pad(d.getUTCMinutes()) + (d.getUTCSeconds() ? ':' + pad(d.getUTCSeconds()) : '');
    return out;
  }

  /* ---------- angka (format Indonesia & internasional) ---------- */
  function parseNum(s) {
    if (typeof s === 'number') return isFinite(s) ? s : null;
    if (typeof s !== 'string') return null;
    let t = s.trim().replace(/^(rp\.?|idr)\s*/i, '').replace(/\s+/g, '');
    if (!t) return null;
    let sign = 1;
    if (/^\(.*\)$/.test(t)) { sign = -1; t = t.slice(1, -1); }
    if (t[0] === '-') { sign = -sign; t = t.slice(1); } else if (t[0] === '+') t = t.slice(1);
    if (/^0\d/.test(t)) return null; // kode/nomor telepon berawalan 0 tidak diubah
    let n;
    if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) n = parseFloat(t.replace(/\./g, '').replace(',', '.')); // 1.250.000,50
    else if (/^\d{1,3}(,\d{3})+\.\d+$/.test(t) || /^\d{1,3}(,\d{3}){2,}$/.test(t)) n = parseFloat(t.replace(/,/g, '')); // 1,250,000.50
    else if (/^\d+,\d+$/.test(t)) n = parseFloat(t.replace(',', '.')); // 12,5
    else if (/^\d+(\.\d+)?$/.test(t)) n = parseFloat(t); // 1250 / 12.5
    else return null;
    return isFinite(n) ? sign * n : null;
  }
  function numOf(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v === 'string') return parseNum(v);
    return null;
  }

  /* ---------- tampilan nilai ---------- */
  function fmtVal(v, o) {
    o = o || {};
    if (v === null || v === undefined) return '';
    if (isDate(v)) return fmtDate(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') {
      if (o.decimalComma) return String(v).replace('.', ',');
      if (o.thousand) return v.toLocaleString('id-ID', { maximumFractionDigits: 2 });
      return String(v);
    }
    return String(v);
  }
  const keyOf = v => (isEmpty(v) ? '(kosong)' : isDate(v) ? fmtDate(v) : String(v));
  const disp = v => (isEmpty(v) ? '(kosong)' : isDate(v) ? fmtDate(v) : v);

  /* ---------- baca sheet ---------- */
  function sheetToAOA(XLSX, ws) {
    if (!ws || !ws['!ref']) return [];
    const rg = XLSX.utils.decode_range(ws['!ref']);
    const aoa = [];
    for (let r = rg.s.r; r <= rg.e.r; r++) {
      const row = [];
      for (let c = rg.s.c; c <= rg.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        let v = '';
        if (cell) {
          if (cell.t === 'n' && cell.z && XLSX.SSF && XLSX.SSF.is_date(cell.z)) v = serialToDate(cell.v);
          else if (cell.t === 'd' && cell.v instanceof Date) v = cell.v;
          else if (cell.t === 'e') v = cell.w || '';
          else if (cell.v === undefined || cell.v === null) v = '';
          else v = cell.v;
        }
        row.push(v);
      }
      aoa.push(row);
    }
    // sel gabungan vertikal (satu kolom, beberapa baris): salin nilai ke seluruh baris yang digabung
    (ws['!merges'] || []).forEach(m => {
      if (m.s.c !== m.e.c || m.e.r <= m.s.r) return;
      const c = m.s.c - rg.s.c, r0 = m.s.r - rg.s.r;
      if (c < 0 || r0 < 0 || !aoa[r0]) return;
      const v = aoa[r0][c];
      if (isEmpty(v)) return;
      for (let r = r0 + 1; r <= m.e.r - rg.s.r && r < aoa.length; r++) if (isEmpty(aoa[r][c])) aoa[r][c] = v;
    });
    // buang baris kosong di bagian bawah (rentang sheet sering lebih panjang dari datanya)
    while (aoa.length && aoa[aoa.length - 1].every(isEmpty)) aoa.pop();
    return aoa;
  }

  function autoHeaderRow(aoa) {
    const lim = Math.min(aoa.length, 30);
    const counts = [];
    for (let i = 0; i < lim; i++) counts.push(aoa[i].filter(v => !isEmpty(v)).length);
    const max = Math.max(0, ...counts);
    if (max < 2) return 0;
    const need = Math.ceil(max * 0.6);
    for (let i = 0; i < lim; i++) {
      if (counts[i] < need) continue;
      const ne = aoa[i].filter(v => !isEmpty(v));
      if (ne.filter(v => typeof v === 'string').length >= ne.length / 2) return i;
    }
    return 0;
  }

  function uniqueHeaders(arr) {
    const seen = new Map();
    return arr.map((h, i) => {
      let name = String(h == null ? '' : h).replace(/[ ​]/g, ' ').replace(/\s+/g, ' ').trim();
      if (!name) name = 'Kolom ' + (i + 1);
      const k = name.toLowerCase();
      const n = (seen.get(k) || 0) + 1;
      seen.set(k, n);
      return n > 1 ? name + ' (' + n + ')' : name;
    });
  }

  function buildDS(aoa, h) {
    if (!aoa.length) return { headers: [], rows: [] };
    h = Math.max(0, Math.min(h | 0, aoa.length - 1));
    const ncol = aoa.reduce((m, r) => Math.max(m, r.length), 0);
    const headers = uniqueHeaders(Array.from({ length: ncol }, (_, i) => (isEmpty(aoa[h][i]) ? '' : fmtVal(aoa[h][i]))));
    const rows = aoa.slice(h + 1).map(r => Array.from({ length: ncol }, (_, i) => (r[i] === undefined || r[i] === null ? '' : r[i])));
    return { headers, rows };
  }

  /* ---------- merapikan ---------- */
  function applyCase(s, mode) {
    if (mode === 'upper') return s.toUpperCase();
    if (mode === 'lower') return s.toLowerCase();
    if (mode === 'proper') return s.toLowerCase().replace(/(^|[\s\-\/(.])(\p{L})/gu, (m, a, b) => a + b.toUpperCase());
    return s;
  }

  /* baris subtotal: ada sel berakhiran "Total", "Subtotal", "Grand Total" */
  const SUBTOTAL_RE = /(^|\s)((sub|grand)\s?)?total$/i;
  const isSubtotalRow = r => r.some(v => typeof v === 'string' && SUBTOTAL_RE.test(v.trim()));

  /* teks tanggal -> Date (UTC). order: 'dmy' | 'mdy' untuk bentuk a/b/tahun */
  function mkDate(y, mo, d, h, mi, s) {
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    const dt = new Date(Date.UTC(y, mo - 1, d, (h | 0), (mi | 0), (s | 0)));
    return dt.getUTCMonth() === mo - 1 ? dt : null;
  }
  const RE_YMD = /^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
  const RE_ABY = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
  function parseDateText(s, order) {
    const t = String(s).trim();
    let m = RE_YMD.exec(t);
    if (m) return mkDate(+m[1], +m[2], +m[3], m[4], m[5], m[6]);
    m = RE_ABY.exec(t);
    if (m) return order === 'mdy' ? mkDate(+m[3], +m[1], +m[2], m[4], m[5], m[6]) : mkDate(+m[3], +m[2], +m[1], m[4], m[5], m[6]);
    return null;
  }

  /* tebak kolom yang perlu diisi ke bawah (sel kosong akibat merge) */
  function suggestFillCols(ds) {
    const rows = ds.rows.filter(r => r.some(v => !isEmpty(v)) && !isSubtotalRow(r));
    if (!rows.length) return [];
    const st = ds.headers.map((_, c) => {
      let empty = 0, ne = 0, num = 0;
      const set = new Set();
      for (const r of rows) {
        const v = r[c];
        if (isEmpty(v)) { empty++; continue; }
        ne++;
        if (numOf(v) !== null) num++;
        set.add(keyOf(v));
      }
      return { empty, ne, numeric: ne ? num / ne >= 0.8 : true, distinct: set.size };
    });
    // kolom kunci = kolom teks paling kanan yang tidak punya sel kosong (mis. nama customer);
    // kolom teks di sebelah kirinya yang punya sel kosong dianggap label kelompok (hasil merge)
    let key = -1;
    for (let c = st.length - 1; c >= 0; c--) if (!st[c].numeric && st[c].ne && st[c].empty === 0) { key = c; break; }
    const out = [];
    for (let c = 0; c < st.length; c++) {
      const s = st[c];
      if (s.numeric || !s.ne || !s.empty) continue;
      if (!(key >= 0 ? c < key : s.empty / rows.length >= 0.2)) continue;
      // sel kosong jarang + isinya hampir unik per baris (email, ID): bukan label kelompok
      if (s.ne > 10 && s.empty / rows.length < 0.3 && s.distinct / s.ne > 0.9) continue;
      out.push(c);
    }
    return out;
  }

  /* blok kolom yang dipisahkan kolom kosong */
  function findBlocks(aoa) {
    const ncol = aoa.reduce((m, r) => Math.max(m, r.length), 0);
    const blocks = [];
    let cur = null;
    for (let c = 0; c < ncol; c++) {
      let n = 0;
      for (const r of aoa) if (!isEmpty(r[c])) n++;
      if (n) { if (!cur) { cur = { c0: c, c1: c, count: 0 }; blocks.push(cur); } cur.c1 = c; cur.count += n; }
      else cur = null;
    }
    return blocks;
  }
  const cropCols = (aoa, c0, c1) => aoa.map(r => r.slice(c0, c1 + 1));

  function tidy(ds, o) {
    o = o || {};
    const rep = { cells: 0, emptyRows: 0, emptyCols: 0, dups: 0, filled: 0, numCols: [], subtotals: 0, dateCols: [], filledNames: [] };
    const ncol = ds.headers.length;
    let headers = o.headers ? uniqueHeaders(ds.headers) : ds.headers.slice();
    const caseCols = o.caseCols && o.caseCols.length ? new Set(o.caseCols) : null;

    let rows = ds.rows.map(r => r.map((v, c) => {
      if (typeof v !== 'string') return v;
      let s = v;
      if (o.nbsp) s = s.replace(/[ ​  ﻿]/g, ' ');
      if (o.breaks) s = s.replace(/[\r\n\t]+/g, ' ');
      if (o.trim) s = s.trim();
      if (o.collapse) s = s.replace(/\s{2,}/g, ' ');
      if (o.case && o.case !== 'none' && (!caseCols || caseCols.has(c))) s = applyCase(s, o.case);
      if (s !== v) rep.cells++;
      return s;
    }));

    if (o.subtotals) {
      const before = rows.length;
      rows = rows.filter(r => !isSubtotalRow(r));
      rep.subtotals = before - rows.length;
    }

    if (o.emptyRows) {
      const before = rows.length;
      rows = rows.filter(r => r.some(v => !isEmpty(v)));
      rep.emptyRows = before - rows.length;
    }

    (o.fillCols || []).forEach(c => {
      let last = '', n = 0;
      rows.forEach(r => {
        if (isEmpty(r[c])) { if (!isEmpty(last)) { r[c] = last; rep.filled++; n++; } } else last = r[c];
      });
      if (n) rep.filledNames.push(headers[c]);
    });

    if (o.toNumber) {
      const convertible = s => !/^\d{12,}$/.test(s.trim()) && parseNum(s) !== null;
      for (let c = 0; c < ncol; c++) {
        let n = 0, p = 0, strs = 0;
        for (const r of rows) {
          const v = r[c];
          if (isEmpty(v)) continue;
          n++;
          if (typeof v === 'number') p++;
          else if (typeof v === 'string') { strs++; if (convertible(v)) p++; }
        }
        if (n && strs && p > 0 && p / n >= 0.9) {
          let cv = 0;
          for (const r of rows) {
            const v = r[c];
            if (typeof v === 'string' && !isEmpty(v) && convertible(v)) { r[c] = parseNum(v); cv++; }
          }
          rep.numCols.push(headers[c]);
          rep.cells += cv;
        }
      }
    }

    if (o.dates) {
      for (let c = 0; c < ncol; c++) {
        let n = 0, p = 0, maxA = 0, maxB = 0, slash = 0;
        for (const r of rows) {
          const v = r[c];
          if (isEmpty(v)) continue;
          n++;
          if (isDate(v)) { p++; continue; }
          if (typeof v !== 'string') continue;
          const t = v.trim();
          if (RE_YMD.test(t)) p++;
          else { const m = RE_ABY.exec(t); if (m && parseDateText(t, 'dmy') || m && parseDateText(t, 'mdy')) { p++; slash++; maxA = Math.max(maxA, +m[1]); maxB = Math.max(maxB, +m[2]); } }
        }
        if (!n || !p || p / n < 0.9) continue;
        let order = o.dates, ambiguous = false;
        if (order === 'auto') {
          if (maxA > 12) order = 'dmy';
          else if (maxB > 12) order = 'mdy';
          else { order = 'dmy'; ambiguous = slash > 0; }
        }
        let cv = 0;
        for (const r of rows) {
          const v = r[c];
          if (typeof v === 'string' && !isEmpty(v)) { const d = parseDateText(v, order); if (d) { r[c] = d; cv++; } }
        }
        if (cv) { rep.dateCols.push({ name: headers[c], order, ambiguous }); rep.cells += cv; }
      }
    }

    if (o.dedupe) {
      const cols = o.dedupeCols && o.dedupeCols.length ? o.dedupeCols : Array.from({ length: ncol }, (_, i) => i);
      const seen = new Set();
      const before = rows.length;
      rows = rows.filter(r => {
        const k = cols.map(c => {
          const v = r[c];
          return isDate(v) ? 'd' + v.getTime() : typeof v === 'string' ? 's' + v.trim().toLowerCase() : typeof v + String(v);
        }).join('\u0001');
        if (seen.has(k)) return false;
        seen.add(k);
        return true;
      });
      rep.dups = before - rows.length;
    }

    if (o.emptyCols && rows.length) {
      const keep = [];
      for (let c = 0; c < ncol; c++) {
        const auto = /^Kolom \d+$/.test(headers[c]);
        if (auto && !rows.some(r => !isEmpty(r[c]))) rep.emptyCols++; else keep.push(c);
      }
      if (keep.length !== ncol) {
        headers = keep.map(c => headers[c]);
        rows = rows.map(r => keep.map(c => r[c]));
      }
    }
    return { headers, rows, report: rep };
  }

  /* ---------- urut & filter ---------- */
  const rank = v => (typeof v === 'number' ? 0 : isDate(v) ? 1 : 2);
  function cmpVal(a, b) {
    const ra = rank(a), rb = rank(b);
    if (ra !== rb) return ra - rb;
    if (ra === 0) return a - b;
    if (ra === 1) return a.getTime() - b.getTime();
    return String(a).localeCompare(String(b), 'id', { numeric: true, sensitivity: 'base' });
  }
  function sortRows(rows, keys) {
    if (!keys || !keys.length) return rows.slice();
    return rows.slice().sort((x, y) => {
      for (const k of keys) {
        const a = x[k.col], b = y[k.col];
        const ea = isEmpty(a), eb = isEmpty(b);
        if (ea || eb) { if (ea && eb) continue; return ea ? 1 : -1; }
        const c = cmpVal(a, b);
        if (c) return k.dir === 'desc' ? -c : c;
      }
      return 0;
    });
  }

  function matchFilter(v, f) {
    const op = f.op;
    if (op === 'empty') return isEmpty(v);
    if (op === 'notempty') return !isEmpty(v);
    if (isEmpty(v)) return op === 'notcontains' || op === 'notequals';
    const q = String(f.val == null ? '' : f.val).trim();
    const s = fmtVal(v).trim();
    const sl = s.toLowerCase(), ql = q.toLowerCase();
    switch (op) {
      case 'contains': return sl.includes(ql);
      case 'notcontains': return !sl.includes(ql);
      case 'starts': return sl.startsWith(ql);
      case 'ends': return sl.endsWith(ql);
      case 'equals': case 'notequals': {
        const nv = numOf(v), nq = parseNum(q);
        const r = nv !== null && nq !== null ? nv === nq : sl === ql;
        return op === 'equals' ? r : !r;
      }
      case 'gt': case 'gte': case 'lt': case 'lte': {
        let c;
        if (isDate(v)) { const d = fmtDate(v).slice(0, q.length); c = d < q ? -1 : d > q ? 1 : 0; }
        else {
          const nv = numOf(v), nq = parseNum(q);
          c = nv !== null && nq !== null ? nv - nq : s.localeCompare(q, 'id', { numeric: true, sensitivity: 'base' });
        }
        return op === 'gt' ? c > 0 : op === 'gte' ? c >= 0 : op === 'lt' ? c < 0 : c <= 0;
      }
    }
    return true;
  }
  function filterRows(rows, filters, search) {
    const fs = (filters || []).filter(f => f && f.col >= 0 && (f.op === 'empty' || f.op === 'notempty' || String(f.val == null ? '' : f.val).trim() !== ''));
    const q = (search || '').trim().toLowerCase();
    if (!fs.length && !q) return rows;
    return rows.filter(r => {
      for (const f of fs) if (!matchFilter(r[f.col], f)) return false;
      if (q && !r.some(v => fmtVal(v).toLowerCase().includes(q))) return false;
      return true;
    });
  }

  /* ---------- agregasi, kelompok, pivot ---------- */
  const FN_LABEL = { rows: 'Jumlah baris', count: 'Hitung', distinct: 'Unik', sum: 'Jumlah', avg: 'Rata-rata', min: 'Min', max: 'Maks' };

  function aggregate(fn, vals) {
    switch (fn) {
      case 'rows': return vals.length;
      case 'count': return vals.filter(v => !isEmpty(v)).length;
      case 'distinct': return new Set(vals.filter(v => !isEmpty(v)).map(keyOf)).size;
      case 'sum': case 'avg': {
        let s = 0, n = 0;
        for (const v of vals) { const x = numOf(v); if (x !== null) { s += x; n++; } }
        if (!n) return fn === 'sum' ? 0 : '';
        return round2(fn === 'sum' ? s : s / n);
      }
      case 'min': case 'max': {
        const ne = vals.filter(v => !isEmpty(v));
        if (!ne.length) return '';
        const nums = ne.map(numOf).filter(x => x !== null);
        if (nums.length) return round2(nums.reduce((a, b) => (fn === 'min' ? (b < a ? b : a) : (b > a ? b : a))));
        return ne.reduce((a, b) => ((fn === 'min' ? cmpVal(b, a) < 0 : cmpVal(b, a) > 0) ? b : a));
      }
    }
    return '';
  }

  function groupBy(ds, groupCols, aggs) {
    const map = new Map();
    for (const r of ds.rows) {
      const k = groupCols.map(c => keyOf(r[c])).join('\u0001');
      let g = map.get(k);
      if (!g) { g = { keys: groupCols.map(c => disp(r[c])), rows: [] }; map.set(k, g); }
      g.rows.push(r);
    }
    const headers = groupCols.map(c => ds.headers[c]).concat(aggs.map(a => (a.fn === 'rows' ? FN_LABEL.rows : FN_LABEL[a.fn] + ' ' + ds.headers[a.col])));
    let out = [...map.values()].map(g => g.keys.concat(aggs.map(a => aggregate(a.fn, a.fn === 'rows' ? g.rows : g.rows.map(r => r[a.col])))));
    out = sortRows(out, groupCols.map((_, i) => ({ col: i, dir: 'asc' })));
    return { headers, rows: out };
  }

  function pivot(ds, o) {
    const rc = o.rows || [];
    const cc = o.col == null || o.col < 0 ? null : o.col;
    let vc = o.val == null || o.val < 0 ? null : o.val;
    let fn = o.fn || 'sum';
    if (vc === null) fn = 'rows';
    const rmap = new Map(), colKeys = new Map(), colTot = new Map(), grand = [];
    for (const r of ds.rows) {
      const rk = rc.map(c => keyOf(r[c])).join('\u0001');
      let ent = rmap.get(rk);
      if (!ent) { ent = { keys: rc.map(c => disp(r[c])), cells: new Map(), all: [] }; rmap.set(rk, ent); }
      const ck = cc === null ? '' : keyOf(r[cc]);
      if (cc !== null && !colKeys.has(ck)) colKeys.set(ck, disp(r[cc]));
      const val = vc === null ? 1 : r[vc];
      if (!ent.cells.has(ck)) ent.cells.set(ck, []);
      ent.cells.get(ck).push(val);
      ent.all.push(val);
      if (!colTot.has(ck)) colTot.set(ck, []);
      colTot.get(ck).push(val);
      grand.push(val);
    }
    const colList = [...colKeys.entries()].sort((a, b) => cmpVal(a[1], b[1]));
    const ents = [...rmap.values()].sort((a, b) => {
      for (let i = 0; i < rc.length; i++) { const c = cmpVal(a.keys[i], b.keys[i]); if (c) return c; }
      return 0;
    });
    const valLabel = FN_LABEL[fn] + (vc !== null ? ' ' + ds.headers[vc] : '');
    const showTotals = !!o.totals;
    const headers = rc.map(c => ds.headers[c]);
    if (cc === null) headers.push(valLabel);
    else { colList.forEach(([, label]) => headers.push(String(label))); if (showTotals) headers.push('Total'); }

    const rows = ents.map(e => {
      const row = e.keys.slice();
      if (cc === null) row.push(aggregate(fn, e.all));
      else {
        colList.forEach(([ck]) => row.push(e.cells.has(ck) ? aggregate(fn, e.cells.get(ck)) : ''));
        if (showTotals) row.push(aggregate(fn, e.all));
      }
      return row;
    });
    if (showTotals) {
      const t = rc.map((_, i) => (i === 0 ? 'Total' : ''));
      if (cc === null) t.push(aggregate(fn, grand));
      else { colList.forEach(([ck]) => t.push(aggregate(fn, colTot.get(ck)))); t.push(aggregate(fn, grand)); }
      rows.push(t);
    }
    return { headers, rows, hasTotalRow: showTotals };
  }

  return {
    isDate, isEmpty, parseNum, numOf, fmtVal, fmtDate, hasTime, serialToDate, dateToSerial,
    sheetToAOA, autoHeaderRow, buildDS, uniqueHeaders, tidy, isSubtotalRow, parseDateText, suggestFillCols, findBlocks, cropCols, sortRows, filterRows, matchFilter,
    aggregate, groupBy, pivot, cmpVal, FN_LABEL
  };
});
