/* Rapikan Data — antarmuka */
(() => {
  'use strict';
  const R = window.RD;
  const $ = s => document.querySelector(s);
  const $$ = s => [...document.querySelectorAll(s)];
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const nf = n => Number(n).toLocaleString('id-ID');

  const S = {
    wb: null, name: '', sheet: '', aoa: [], headerRow: 0,
    orig: null, cur: null, undo: [], view: [],
    filters: [], sorts: [], search: '',
    sel: { caseCols: [], fillCols: [], dedupeCols: [], groupCols: [], pvRows: [] },
    aggs: [{ fn: 'rows', col: 0 }],
    results: { group: null, pivot: null }
  };

  const OPS = [
    ['contains', 'mengandung'], ['notcontains', 'tidak mengandung'], ['equals', 'sama dengan'], ['notequals', 'tidak sama dengan'],
    ['starts', 'diawali'], ['ends', 'diakhiri'], ['gt', 'lebih besar dari'], ['gte', 'lebih besar atau sama'],
    ['lt', 'lebih kecil dari'], ['lte', 'lebih kecil atau sama'], ['empty', 'kosong'], ['notempty', 'tidak kosong']
  ];
  const FNS = [
    ['rows', 'Jumlah baris'], ['sum', 'Jumlah (Sum)'], ['count', 'Hitung (Count)'], ['avg', 'Rata-rata'],
    ['min', 'Minimum'], ['max', 'Maksimum'], ['distinct', 'Hitung unik']
  ];

  let toastTimer;
  function toast(msg) {
    const t = $('#toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 3200);
  }

  /* ---------- memuat file ---------- */
  async function loadFile(file) {
    try {
      const ext = (file.name.split('.').pop() || '').toLowerCase();
      let wb;
      if (['csv', 'tsv', 'txt'].includes(ext)) {
        const text = (await file.text()).replace(/^﻿/, '');
        wb = XLSX.read(text, { type: 'string', raw: true, cellNF: true });
      } else {
        const buf = file.arrayBuffer ? await file.arrayBuffer() : await new Promise((ok, no) => {
          const fr = new FileReader(); fr.onload = () => ok(fr.result); fr.onerror = no; fr.readAsArrayBuffer(file);
        });
        wb = XLSX.read(buf, { type: 'array', cellNF: true });
      }
      openWorkbook(wb, file.name);
    } catch (e) {
      console.error(e);
      toast('Gagal membaca file. Pastikan formatnya Excel atau CSV.');
    }
  }

  function openWorkbook(wb, name) {
    S.wb = wb; S.name = name;
    $('#fileName').textContent = name;
    $('#sheetSel').innerHTML = wb.SheetNames.map(n => `<option>${esc(n)}</option>`).join('');
    useSheet(wb.SheetNames[0]);
    $('#welcome').hidden = true;
    $('#app').hidden = false;
  }

  function useSheet(name) {
    S.sheet = name;
    S.fullAOA = R.sheetToAOA(XLSX, S.wb.Sheets[name]);
    S.blocks = R.findBlocks(S.fullAOA);
    let best = 0;
    S.blocks.forEach((b, i) => { if (b.count > S.blocks[best].count) best = i; });
    const col = c => XLSX.utils.encode_col(c);
    $('#blockSel').innerHTML = S.blocks.map((b, i) => {
      const part = R.cropCols(S.fullAOA, b.c0, b.c1);
      const names = (part[R.autoHeaderRow(part)] || []).filter(v => !R.isEmpty(v)).slice(0, 3).map(v => R.fmtVal(v)).join(', ');
      return `<option value="${i}"${i === best ? ' selected' : ''}>Blok ${i + 1}: kolom ${col(b.c0)}–${col(b.c1)} (${esc(names)}…)</option>`;
    }).join('') + '<option value="-1">Semua kolom</option>';
    $('#blockWrap').hidden = S.blocks.length < 2;
    applyBlock(S.blocks.length ? best : -1);
  }

  function applyBlock(i) {
    const b = S.blocks[i];
    S.aoa = b ? R.cropCols(S.fullAOA, b.c0, b.c1) : S.fullAOA;
    S.headerRow = R.autoHeaderRow(S.aoa);
    $('#headerRow').value = S.headerRow + 1;
    rebuild();
  }

  function rebuild() {
    S.orig = R.buildDS(S.aoa, S.headerRow);
    S.cur = S.orig;
    S.undo = [];
    resetView(true);
    S.sel = { caseCols: [], fillCols: [], dedupeCols: [], groupCols: [], pvRows: [] };
    S.aggs = [{ fn: 'rows', col: 0 }];
    S.results = { group: null, pivot: null };
    $('#groupResult').innerHTML = '';
    $('#pivotResult').innerHTML = '';
    $('#tidyReport').hidden = true;
    refreshAll();
  }

  function setCur(ds) {
    const sameCols = ds.headers.length === S.cur.headers.length;
    S.cur = ds;
    if (!sameCols) {
      resetView(true);
      S.sel = { caseCols: [], fillCols: [], dedupeCols: [], groupCols: [], pvRows: [] };
      S.aggs = [{ fn: 'rows', col: 0 }];
    }
    S.results = { group: null, pivot: null };
    $('#groupResult').innerHTML = '';
    $('#pivotResult').innerHTML = '';
    refreshAll();
  }

  function resetView(silent) {
    S.filters = []; S.sorts = []; S.search = '';
    $('#search').value = '';
    if (!silent) { renderFilters(); renderSorts(); renderDataTable(); }
  }

  /* ---------- helpers UI ---------- */
  function colOptions(selected, withNone, noneLabel) {
    let h = withNone ? `<option value="-1">${esc(noneLabel || '(tidak ada)')}</option>` : '';
    h += S.cur.headers.map((t, i) => `<option value="${i}"${i === selected ? ' selected' : ''}>${esc(t)}</option>`).join('');
    return h;
  }

  function renderChips(key) {
    const box = $('#chips-' + key);
    const arr = S.sel[key];
    box.innerHTML = S.cur.headers.map((t, i) => {
      const pos = arr.indexOf(i);
      return `<button type="button" class="chip${pos >= 0 ? ' on' : ''}" data-key="${key}" data-i="${i}">${pos >= 0 && (key === 'groupCols' || key === 'pvRows') ? `<span class="ord">${pos + 1}</span>` : ''}${esc(t)}</button>`;
    }).join('');
  }
  const CHIP_KEYS = ['caseCols', 'fillCols', 'dedupeCols', 'groupCols', 'pvRows'];

  function refreshAll() {
    const n = S.cur.rows.length, m = S.cur.headers.length;
    $('#statLine').textContent = `${nf(n)} baris · ${m} kolom`;
    CHIP_KEYS.forEach(renderChips);
    renderFilters(); renderSorts(); renderAggs();
    $('#pvCol').innerHTML = colOptions(-1, true, '(tidak ada)');
    $('#pvVal').innerHTML = colOptions(-1, true, '(tidak ada — hitung baris)');
    renderDataTable();
    $('#btnUndo').disabled = !S.undo.length;
  }

  /* ---------- tabel ---------- */
  function tableHTML(headers, rows, o) {
    o = o || {};
    const lim = o.limit || 500;
    const sortOf = c => (o.sorts || []).findIndex(s => s.col === c);
    let h = '<table><thead><tr>';
    if (o.rownum) h += '<th class="rn">#</th>';
    headers.forEach((t, i) => {
      let arrow = '';
      const k = sortOf(i);
      if (k >= 0) arrow = (o.sorts[k].dir === 'desc' ? ' ▼' : ' ▲') + (o.sorts.length > 1 ? (k + 1) : '');
      h += `<th${o.sortable ? ` class="sortable" data-c="${i}" title="Klik untuk mengurutkan"` : ''}>${esc(t)}${arrow}</th>`;
    });
    h += '</tr></thead><tbody>';
    rows.slice(0, lim).forEach((r, ri) => {
      const isTotal = o.totalRow && ri === rows.length - 1;
      h += `<tr${isTotal ? ' class="total"' : ''}>`;
      if (o.rownum) h += `<td class="rn">${ri + 1}</td>`;
      r.forEach(v => {
        const txt = esc(R.fmtVal(v, { thousand: o.thousand }));
        h += `<td${typeof v === 'number' ? ' class="num"' : ''} title="${txt}">${txt}</td>`;
      });
      h += '</tr>';
    });
    h += '</tbody></table>';
    if (rows.length > lim) h += `<div class="hint" style="padding:8px 10px">Menampilkan ${nf(lim)} dari ${nf(rows.length)} baris. Unduh file untuk melihat semuanya.</div>`;
    if (!rows.length) h += '<div class="hint" style="padding:8px 10px">Tidak ada baris yang cocok.</div>';
    return h;
  }

  function renderDataTable() {
    S.view = R.sortRows(R.filterRows(S.cur.rows, S.filters, S.search), S.sorts);
    $('#viewMeta').textContent = `${nf(S.view.length)} dari ${nf(S.cur.rows.length)} baris tampil`;
    $('#dataTable').innerHTML = tableHTML(S.cur.headers, S.view, { rownum: true, sortable: true, sorts: S.sorts });
  }

  /* ---------- filter & urut ---------- */
  function renderFilters() {
    $('#filters').innerHTML = S.filters.map((f, i) => `
      <div class="frow" data-i="${i}">
        <select data-k="col">${colOptions(f.col)}</select>
        <select data-k="op">${OPS.map(([v, l]) => `<option value="${v}"${f.op === v ? ' selected' : ''}>${l}</option>`).join('')}</select>
        <input type="text" data-k="val" value="${esc(f.val)}" placeholder="nilai"${f.op === 'empty' || f.op === 'notempty' ? ' disabled' : ''}>
        <button class="btn small x" data-act="rmFilter" aria-label="Hapus filter">✕</button>
      </div>`).join('');
  }
  function renderSorts() {
    $('#sorts').innerHTML = S.sorts.map((s, i) => `
      <div class="frow" data-i="${i}">
        <select data-k="col">${colOptions(s.col)}</select>
        <select data-k="dir"><option value="asc"${s.dir === 'asc' ? ' selected' : ''}>A → Z / kecil → besar</option><option value="desc"${s.dir === 'desc' ? ' selected' : ''}>Z → A / besar → kecil</option></select>
        <button class="btn small x" data-act="rmSort" aria-label="Hapus urutan">✕</button>
      </div>`).join('');
  }

  /* ---------- perhitungan ---------- */
  function renderAggs() {
    $('#aggs').innerHTML = S.aggs.map((a, i) => `
      <div class="frow" data-i="${i}">
        <select data-k="fn">${FNS.map(([v, l]) => `<option value="${v}"${a.fn === v ? ' selected' : ''}>${l}</option>`).join('')}</select>
        ${a.fn === 'rows' ? '' : `<select data-k="col">${colOptions(a.col)}</select>`}
        <button class="btn small x" data-act="rmAgg" aria-label="Hapus">✕</button>
      </div>`).join('');
  }

  function viewDS() { return { headers: S.cur.headers, rows: S.view }; }

  function showResult(kind, ds, boxSel, opts) {
    S.results[kind] = ds;
    const box = $(boxSel);
    box.innerHTML = `
      <div class="muted meta">${nf(ds.rows.length)} baris hasil</div>
      <div class="tablewrap">${tableHTML(ds.headers, ds.rows, Object.assign({ thousand: true }, opts))}</div>
      <div class="row wrap gap">
        <button class="btn" data-export="${kind}" data-kind="xlsx">Unduh Excel</button>
        <button class="btn" data-export="${kind}" data-kind="csv">Unduh CSV</button>
        <button class="btn ghost" data-promote="${kind}">Jadikan data utama</button>
      </div>`;
  }

  function runGroup() {
    const gc = S.sel.groupCols;
    if (!gc.length) return toast('Pilih minimal satu kolom untuk dikelompokkan.');
    const aggs = S.aggs.filter(a => a.fn === 'rows' || (a.col >= 0 && a.col < S.cur.headers.length));
    if (!aggs.length) return toast('Tambahkan minimal satu perhitungan.');
    showResult('group', R.groupBy(viewDS(), gc, aggs), '#groupResult');
  }

  function runPivot() {
    const rows = S.sel.pvRows;
    if (!rows.length) return toast('Pilih minimal satu kolom untuk Baris.');
    const col = +$('#pvCol').value, val = +$('#pvVal').value;
    let fn = $('#pvFn').value;
    if (val < 0 && fn !== 'rows') toast('Tanpa kolom Nilai, pivot menghitung jumlah baris.');
    const totals = $('#pvTotals').checked;
    const ds = R.pivot(viewDS(), { rows, col, val, fn, totals });
    showResult('pivot', ds, '#pivotResult', { totalRow: totals });
  }

  /* ---------- rapikan ---------- */
  function runTidy() {
    const o = {
      headers: $('#oHeaders').checked, trim: $('#oTrim').checked, collapse: $('#oCollapse').checked,
      nbsp: $('#oNbsp').checked, breaks: $('#oBreaks').checked, emptyRows: $('#oEmptyRows').checked,
      emptyCols: $('#oEmptyCols').checked, toNumber: $('#oToNumber').checked,
      subtotals: $('#oSubtotals').checked, dates: $('#oDates').value,
      case: $('#oCase').value, caseCols: S.sel.caseCols.slice(),
      fillCols: S.sel.fillCols.slice(), dedupe: $('#oDedupe').checked, dedupeCols: S.sel.dedupeCols.slice()
    };
    const res = R.tidy(S.cur, o);
    S.undo.push(S.cur);
    if (S.undo.length > 10) S.undo.shift();
    const rep = res.report;
    setCur({ headers: res.headers, rows: res.rows });
    const lines = [];
    if (rep.cells) lines.push(`${nf(rep.cells)} sel dirapikan`);
    if (rep.numCols.length) lines.push(`Kolom diubah jadi angka: ${esc(rep.numCols.join(', '))}`);
    if (rep.subtotals) lines.push(`${nf(rep.subtotals)} baris subtotal/Total dihapus`);
    rep.dateCols.forEach(d => lines.push(`Kolom tanggal: ${esc(d.name)} (dibaca sebagai ${d.order === 'mdy' ? 'bulan/hari/tahun' : 'hari/bulan/tahun'})${d.ambiguous ? ' — <b>urutan tidak bisa dipastikan</b>, ubah pilihan "Teks tanggal" jika salah' : ''}`));
    if (rep.emptyRows) lines.push(`${nf(rep.emptyRows)} baris kosong dihapus`);
    if (rep.emptyCols) lines.push(`${nf(rep.emptyCols)} kolom kosong dihapus`);
    if (rep.filled) lines.push(`${nf(rep.filled)} sel kosong diisi dari atasnya (kolom: ${esc(rep.filledNames.join(', '))})`);
    if (rep.dups) lines.push(`${nf(rep.dups)} baris duplikat dihapus`);
    const box = $('#tidyReport');
    box.hidden = false;
    box.innerHTML = lines.length ? `<b>Selesai.</b><ul>${lines.map(l => `<li>${l}</li>`).join('')}</ul>` : '<b>Tidak ada yang perlu diubah.</b> Data sudah rapi untuk pilihan ini.';
    $('#btnUndo').disabled = !S.undo.length;
  }

  function runAuto() {
    $('#oSubtotals').checked = true;
    $('#oDates').value = 'auto';
    ['oHeaders', 'oTrim', 'oCollapse', 'oNbsp', 'oEmptyRows', 'oEmptyCols', 'oToNumber'].forEach(id => { $('#' + id).checked = true; });
    S.sel.fillCols = R.suggestFillCols(S.cur);
    renderChips('fillCols');
    runTidy();
  }

  /* ---------- ekspor ---------- */
  function toSheet(ds) {
    const dateCells = [];
    const aoa = [ds.headers].concat(ds.rows.map((r, ri) => r.map((v, ci) => {
      if (R.isDate(v)) { dateCells.push([ri + 1, ci, R.hasTime(v)]); return R.dateToSerial(v); }
      return v === '' ? null : v;
    })));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    dateCells.forEach(([r, c, t]) => {
      const a = XLSX.utils.encode_cell({ r, c });
      if (ws[a]) ws[a].z = t ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd';
    });
    ws['!cols'] = ds.headers.map((h, i) => {
      let w = String(h).length;
      for (let k = 0; k < Math.min(ds.rows.length, 200); k++) w = Math.max(w, R.fmtVal(ds.rows[k][i]).length);
      return { wch: Math.min(45, w + 2) };
    });
    if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
    return ws;
  }
  function toCSV(ds) {
    const q = v => {
      const s = R.fmtVal(v, { decimalComma: true });
      return /[";\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    return '﻿' + [ds.headers.map(q).join(';')].concat(ds.rows.map(r => r.map(q).join(';'))).join('\r\n');
  }
  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  function exportDS(which, kind) {
    const ds = which === 'view' ? viewDS() : S.results[which];
    if (!ds) return;
    const base = S.name.replace(/\.[^.]+$/, '') || 'data';
    const suffix = { view: 'rapi', group: 'kelompok', pivot: 'pivot' }[which];
    if (kind === 'xlsx') {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, toSheet(ds), (which === 'pivot' ? 'Pivot' : which === 'group' ? 'Kelompok' : 'Data'));
      XLSX.writeFile(wb, `${base}_${suffix}.xlsx`);
    } else {
      download(new Blob([toCSV(ds)], { type: 'text/csv;charset=utf-8' }), `${base}_${suffix}.csv`);
    }
  }

  /* ---------- data contoh ---------- */
  function sampleAOA() {
    const money = n => 'Rp ' + String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    const wil = ['Bandung', 'Jakarta', 'Surabaya'];
    const prod = ['Vitamin C 500mg', 'Masker Medis', 'Hand Sanitizer', 'Minyak Kayu Putih'];
    const harga = [45000, 28500, 19000, 32000];
    const kr = ['@sari.id', '@BUDI_OFC', '@rina.store', '@dewi.sehat'];
    const rows = [['Laporan Penjualan Bulanan', '', '', '', '', '', '', ''], new Array(8).fill('')];
    rows.push(['Tanggal', 'Wilayah', 'Produk ', 'Kreator', 'Qty', 'Harga', 'Status', '']);
    for (let i = 0; i < 30; i++) {
      const p = (i * 7 + 3) % 4;
      rows.push([
        R.serialToDate(46000 + i * 2),
        i % 10 === 0 ? wil[i / 10] : '',
        i % 3 === 0 ? '  ' + prod[p] + '  ' : prod[p],
        i % 4 === 0 ? kr[i % 4] + ' ' : kr[(i + 1) % 4],
        ((i * 5) % 9 + 1) + ' ',
        money(harga[p]),
        i % 5 === 0 ? 'pending ' : 'Selesai',
        ''
      ]);
      if (i === 10) rows.push(new Array(8).fill(''));
    }
    rows.push(rows[rows.length - 1].slice()); // baris duplikat untuk demo
    return rows;
  }
  function loadSample() {
    const ws = XLSX.utils.aoa_to_sheet(sampleAOA().map(r => r.map(v => (R.isDate(v) ? R.dateToSerial(v) : v))));
    const ref = XLSX.utils.decode_range(ws['!ref']);
    for (let r = 3; r <= ref.e.r; r++) { const c = ws[XLSX.utils.encode_cell({ r, c: 0 })]; if (c && c.t === 'n') c.z = 'yyyy-mm-dd'; }
    const wb = { SheetNames: ['Penjualan'], Sheets: { Penjualan: ws } };
    openWorkbook(wb, 'contoh-penjualan.xlsx');
  }

  /* ---------- event ---------- */
  function bind() {
    const fi = $('#fileInput');
    ['#btnOpen', '#btnOpen2'].forEach(s => $(s).addEventListener('click', () => fi.click()));
    fi.addEventListener('change', () => { if (fi.files[0]) loadFile(fi.files[0]); fi.value = ''; });
    $('#btnSample').addEventListener('click', loadSample);
    ['dragover', 'dragenter'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); $('#welcome').classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => document.addEventListener(ev, e => { e.preventDefault(); $('#welcome').classList.remove('over'); }));
    document.addEventListener('drop', e => { const f = e.dataTransfer && e.dataTransfer.files[0]; if (f) loadFile(f); });

    $('#sheetSel').addEventListener('change', e => useSheet(e.target.value));
    $('#blockSel').addEventListener('change', e => applyBlock(+e.target.value));
    $('#btnAuto').addEventListener('click', runAuto);
    $('#btnSuggestFill').addEventListener('click', () => {
      S.sel.fillCols = R.suggestFillCols(S.cur);
      renderChips('fillCols');
      toast(S.sel.fillCols.length ? `${S.sel.fillCols.length} kolom dipilih. Periksa dulu, lalu tekan Terapkan.` : 'Tidak ada kolom yang tampak perlu diisi.');
    });
    $('#headerRow').addEventListener('change', e => {
      const v = Math.max(1, Math.min(S.aoa.length || 1, parseInt(e.target.value, 10) || 1));
      e.target.value = v; S.headerRow = v - 1; rebuild();
    });

    $$('.tab-btn').forEach(b => b.addEventListener('click', () => {
      $$('.tab-btn').forEach(x => x.classList.toggle('active', x === b));
      $$('.tab').forEach(t => { t.hidden = t.id !== 'tab-' + b.dataset.tab; });
    }));

    let st;
    $('#search').addEventListener('input', e => { clearTimeout(st); st = setTimeout(() => { S.search = e.target.value; renderDataTable(); }, 200); });

    $('#addFilter').addEventListener('click', () => { S.filters.push({ col: 0, op: 'contains', val: '' }); renderFilters(); });
    $('#addSort').addEventListener('click', () => { if (S.sorts.length < 4) { S.sorts.push({ col: 0, dir: 'asc' }); renderSorts(); renderDataTable(); } });
    $('#resetView').addEventListener('click', () => resetView(false));

    $('#filters').addEventListener('input', e => {
      const row = e.target.closest('.frow'); if (!row) return;
      const f = S.filters[+row.dataset.i], k = e.target.dataset.k; if (!f || !k) return;
      f[k] = k === 'col' ? +e.target.value : e.target.value;
      if (k === 'op') renderFilters();
      clearTimeout(st); st = setTimeout(renderDataTable, 150);
    });
    $('#filters').addEventListener('click', e => {
      if (e.target.dataset.act !== 'rmFilter') return;
      S.filters.splice(+e.target.closest('.frow').dataset.i, 1); renderFilters(); renderDataTable();
    });
    $('#sorts').addEventListener('change', e => {
      const row = e.target.closest('.frow'); const s = row && S.sorts[+row.dataset.i]; const k = e.target.dataset.k; if (!s || !k) return;
      s[k] = k === 'col' ? +e.target.value : e.target.value; renderDataTable();
    });
    $('#sorts').addEventListener('click', e => {
      if (e.target.dataset.act !== 'rmSort') return;
      S.sorts.splice(+e.target.closest('.frow').dataset.i, 1); renderSorts(); renderDataTable();
    });
    $('#dataTable').addEventListener('click', e => {
      const th = e.target.closest('th[data-c]'); if (!th) return;
      const c = +th.dataset.c, cur = S.sorts[0];
      if (S.sorts.length === 1 && cur.col === c) { if (cur.dir === 'asc') cur.dir = 'desc'; else S.sorts = []; }
      else S.sorts = [{ col: c, dir: 'asc' }];
      renderSorts(); renderDataTable();
    });

    document.addEventListener('click', e => {
      const chip = e.target.closest('.chip');
      if (chip) {
        const arr = S.sel[chip.dataset.key], i = +chip.dataset.i, p = arr.indexOf(i);
        if (p >= 0) arr.splice(p, 1); else arr.push(i);
        renderChips(chip.dataset.key);
        return;
      }
      const ex = e.target.closest('[data-export]');
      if (ex) return exportDS(ex.dataset.export, ex.dataset.kind);
      const pr = e.target.closest('[data-promote]');
      if (pr && S.results[pr.dataset.promote]) {
        const ds = S.results[pr.dataset.promote];
        S.undo.push(S.cur);
        setCur({ headers: ds.headers.slice(), rows: ds.rows.map(r => r.slice()) });
        $$('.tab-btn')[0].click();
        toast('Hasil dijadikan data utama. Gunakan Urungkan di tab Rapikan untuk kembali.');
      }
    });

    $('#btnTidy').addEventListener('click', runTidy);
    $('#btnUndo').addEventListener('click', () => {
      if (!S.undo.length) return;
      setCur(S.undo.pop()); $('#tidyReport').hidden = true; $('#btnUndo').disabled = !S.undo.length;
    });
    $('#btnRestore').addEventListener('click', () => { rebuild(); toast('Data dikembalikan seperti file asli.'); });

    $('#addAgg').addEventListener('click', () => { S.aggs.push({ fn: 'sum', col: 0 }); renderAggs(); });
    $('#aggs').addEventListener('change', e => {
      const row = e.target.closest('.frow'); const a = row && S.aggs[+row.dataset.i]; const k = e.target.dataset.k; if (!a || !k) return;
      a[k] = k === 'col' ? +e.target.value : e.target.value; if (k === 'fn') renderAggs();
    });
    $('#aggs').addEventListener('click', e => {
      if (e.target.dataset.act !== 'rmAgg') return;
      S.aggs.splice(+e.target.closest('.frow').dataset.i, 1); renderAggs();
    });
    $('#btnGroup').addEventListener('click', runGroup);
    $('#btnPivot').addEventListener('click', runPivot);

    let deferred;
    window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); deferred = e; $('#btnInstall').hidden = false; });
    $('#btnInstall').addEventListener('click', async () => {
      if (!deferred) return; deferred.prompt(); await deferred.userChoice; deferred = null; $('#btnInstall').hidden = true;
    });
    window.addEventListener('appinstalled', () => { $('#btnInstall').hidden = true; });

    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
      navigator.serviceWorker.register('sw.js').catch(() => {});
    }
  }

  bind();
})();
