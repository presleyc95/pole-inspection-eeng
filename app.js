/* ============================================================
   Entergy Pole Inspection – Progressive Web App
   ============================================================ */

// ── Constants ────────────────────────────────────────────────

const HEADER_CELLS = {
  projectId:       { c: 'E2'  },
  completionDate:  { c: 'L2'  },
  jurisdiction:    { c: 'R2'  },
  region:          { c: 'X2'  },
  network:         { c: 'AE2' },
  substation:      { c: 'AK2' },
  inspectors:      { c: 'AO2' },
  feeder:          { c: 'AH4' },
  deviceId:        { c: 'AK4' },
  deviceType:      { c: 'AO4' },
  reconductorCost: { c: 'M4'  },
  relocationCost:  { c: 'V4'  },
  sectCost:        { c: 'AD4' },
};

const REGIONS = {
  Central:   ['West Markham','Baseline','Jacksonville'],
  Northeast: ['Batesville','Newport','Walnut Ridge','Blytheville','Forrest City','Harrisburg','Marion','Searcy'],
  Northwest: ['Flippin','Harrison','Mt. View','Conway','Russellville'],
  Southeast: ['Helena','Lonoke','Stuttgart','McGehee','Warren','Pine Bluff'],
  Southwest: ['El Dorado','Magnolia','Arkadelphia','Glenwood','Hot Springs','Malvern'],
};

const DATA_START_ROW = 9;
const MAX_ROW = 509;
const SHEET_NAME = 'Reliability Form';

function colIdx(letter) {
  let n = 0;
  for (let i = 0; i < letter.length; i++) n = n * 26 + letter.charCodeAt(i) - 64;
  return n - 1;
}
function colLetter(idx) {
  let s = ''; idx++;
  while (idx > 0) { idx--; s = String.fromCharCode(65 + (idx % 26)) + s; idx = Math.floor(idx / 26); }
  return s;
}

const POLE_SECTIONS = [
  { id: 'location', title: 'Location & Access', fields: [
    { key: 'C', label: 'DLOC# / GPS',           type: 'text',    gps: true },
    { key: 'D', label: 'Pictures',               type: 'text',    placeholder: 'e.g. 1,2' },
    { key: 'E', label: 'Truck Access',            type: 'yn' },
    { key: 'F', label: 'Traffic Control',          type: 'segment', options: ['H','M','L'], labels: ['Heavy','Med','Light'] },
    { key: 'G', label: 'Additional Cover-Up',     type: 'yn' },
  ]},
  { id: 'structure', title: 'Structure', fields: [
    { key: 'B', label: 'BIL Value (kV)',          type: 'number' },
    { key: 'H', label: 'Pole Type',               type: 'segment', options: ['W','C','S'], labels: ['Wood','Conc','Steel'] },
    { key: 'I', label: 'Wire Size ≥336 ACSR',    type: 'yn' },
    { key: 'J', label: 'Structure Type',           type: 'segment', options: ['DE','DDE','T','A','V'] },
  ]},
  { id: 'pole-arm', title: 'Pole & Crossarm Issues', fields: [
    { key: 'K', label: 'Bad Pole',                type: 'qty' },
    { key: 'L', label: 'Bad Pole (Top/Btm)',      type: 'segment', options: ['T','B','ALL'] },
    { key: 'M', label: 'Bad Crossarm',            type: 'qty' },
    { key: 'N', label: 'Bad Crossarm Brace',      type: 'qty' },
    { key: 'O', label: 'Fiberglass Standoff',     type: 'qty' },
  ]},
  { id: 'hardware', title: 'Hardware & Connections', fields: [
    { key: 'P', label: 'Damaged Insulators',      type: 'qty' },
    { key: 'Q', label: 'Loose Guys',              type: 'qty' },
    { key: 'R', label: 'Bad Anchors',             type: 'qty' },
    { key: 'S', label: 'Guy Strain Insulator',    type: 'qty' },
    { key: 'T', label: 'Lightning Arrester',       type: 'qty' },
    { key: 'U', label: 'Arrester Action',          type: 'select', options: ['','RMV-L','IN-E','RPL-E','REL-E','RMV-E'] },
    { key: 'V', label: 'Fuse Switches',            type: 'qty' },
    { key: 'W', label: 'Fuse Switch Type',         type: 'select', options: ['','3P-IN-FG','1P-IN-FG','3P-RPL-FG','1P-RPL-FG','3P-REL-FG','1P-REL-FG'] },
  ]},
  { id: 'grounding', title: 'Grounding', fields: [
    { key: 'X', label: 'Remove Grounds (Primary Zone)', type: 'qty' },
    { key: 'Y', label: 'Install Hendrix Ground',  type: 'qty' },
    { key: 'Z', label: 'Missing/Damaged Pole Ground', type: 'qty' },
  ]},
  { id: 'conductor', title: 'Conductor & Line', fields: [
    { key: 'AA', label: 'Unfused Lateral/Xfrmr',  type: 'qty' },
    { key: 'AB', label: 'Animal Guard/Avian Cover',type: 'qty' },
    { key: 'AC', label: 'Slack Conductor',         type: 'qty' },
    { key: 'AD', label: 'Missing Neutral (Spans)', type: 'qty' },
    { key: 'AE', label: 'Conductor Damage',        type: 'qty' },
    { key: 'AF', label: 'AAAC Sleeve on 336ACSR',  type: 'qty' },
  ]},
  { id: 'switches', title: 'Switches & Other', fields: [
    { key: 'AG', label: 'Disconnect Switch Dmg',   type: 'qty' },
    { key: 'AH', label: 'GOAB Switch Dmg',         type: 'qty' },
    { key: 'AI', label: 'Vegetation Issues',        type: 'qty' },
    { key: 'AJ', label: 'Other Issues',             type: 'qty' },
    { key: 'AK', label: 'Other Issues Cost ($)',    type: 'number', prefix: '$' },
  ]},
  { id: 'classification', title: 'Classification', fields: [
    { key: 'AM', label: 'Reliability Override',     type: 'select', options: ['','NO WORK','URGENT','POLE-3rd PARTY'] },
    { key: 'AN', label: 'Improvement Options',      type: 'select', options: ['','RECON','SECT','RELOC'] },
  ]},
  { id: 'comments', title: 'Comments', fields: [
    { key: 'AO', label: 'Safety Assessment & Comments', type: 'textarea' },
  ]},
];

const DATA_COLS = [];
POLE_SECTIONS.forEach(s => s.fields.forEach(f => DATA_COLS.push(f.key)));
DATA_COLS.unshift('A');

// ── App State ────────────────────────────────────────────────

let state = {
  rawFile: null,         // original file bytes (ArrayBuffer) for export
  workbook: null,        // SheetJS workbook (for reading)
  header: {},
  poles: [],             // poles with DLOC numbers
  emptyRows: [],         // rows with item # but no DLOC (available for new poles)
  currentPole: 0,
  fileName: '',
};

// ── Screen Management ────────────────────────────────────────

function showScreen(id) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById('screen-' + id).classList.add('active');
  if (id === 'pole') renderPole();
  if (id === 'summary') renderSummary();
}

// ── File Import ──────────────────────────────────────────────

const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const importStatus = document.getElementById('import-status');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.style.borderColor = 'var(--blue)'; });
dropZone.addEventListener('dragleave', () => { dropZone.style.borderColor = ''; });
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.style.borderColor = '';
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => { if (fileInput.files.length) handleFile(fileInput.files[0]); });

function handleFile(file) {
  importStatus.hidden = false;
  importStatus.className = 'import-status';
  importStatus.textContent = 'Reading file...';
  state.fileName = file.name;

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const buf = e.target.result;
      state.rawFile = buf.slice(0);  // keep a copy of original bytes
      const data = new Uint8Array(buf);
      state.workbook = XLSX.read(data, { type: 'array', cellFormula: true, cellStyles: false });
      parseWorkbook();
      importStatus.textContent = `Loaded ${state.poles.length} poles (${state.emptyRows.length} empty rows available)`;
      populateHeaderForm();
      document.getElementById('pole-count-label').textContent = state.poles.length;
      _doAutoSave(); // persist rawFile + data to IndexedDB immediately
      setTimeout(() => showScreen('header'), 400);
    } catch (err) {
      importStatus.className = 'import-status error';
      importStatus.textContent = 'Error: ' + err.message;
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseWorkbook() {
  const ws = state.workbook.Sheets[SHEET_NAME];
  if (!ws) throw new Error('Sheet "' + SHEET_NAME + '" not found');

  // Read header cells
  for (const [key, def] of Object.entries(HEADER_CELLS)) {
    const cell = ws[def.c];
    if (cell) {
      if (key === 'completionDate' && cell.t === 'd') {
        state.header[key] = formatDateForInput(cell.v);
      } else if (key === 'completionDate' && cell.t === 'n') {
        state.header[key] = formatDateForInput(excelDateToJS(cell.v));
      } else {
        state.header[key] = cell.v != null ? String(cell.v) : '';
      }
    }
  }

  // Read pole rows – only count poles that have a DLOC number (column C)
  state.poles = [];
  state.emptyRows = [];

  for (let r = DATA_START_ROW; r <= MAX_ROW; r++) {
    const cellA = ws[`A${r}`];
    if (!cellA || cellA.v == null) continue;

    const cellC = ws[`C${r}`];
    const hasDLOC = cellC && cellC.v != null && String(cellC.v).trim() !== '';

    if (hasDLOC) {
      const pole = {};
      for (const col of DATA_COLS) {
        const addr = col + r;
        const cell = ws[addr];
        if (cell && cell.v != null) pole[col] = cell.v;
      }
      pole._row = r;
      state.poles.push(pole);
    } else {
      state.emptyRows.push({ itemNum: cellA.v, row: r });
    }
  }
}

function excelDateToJS(serial) {
  const epoch = new Date(1899, 11, 30);
  return new Date(epoch.getTime() + serial * 86400000);
}

function jsDateToExcelSerial(dateStr) {
  const d = new Date(dateStr);
  const epoch = new Date(1899, 11, 30);
  return Math.round((d.getTime() - epoch.getTime()) / 86400000);
}

function formatDateForInput(d) {
  if (typeof d === 'string') d = new Date(d);
  if (!(d instanceof Date) || isNaN(d)) return '';
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

// ── Header Form ──────────────────────────────────────────────

function populateHeaderForm() {
  for (const [key, val] of Object.entries(state.header)) {
    const el = document.querySelector(`[data-header="${key}"]`);
    if (el) {
      el.value = val || '';
      if (key === 'jurisdiction') updateRegionOptions(val);
      if (key === 'region') setTimeout(() => { updateNetworkOptions(val); }, 0);
    }
  }
}

document.querySelector('[data-header="jurisdiction"]').addEventListener('change', function() {
  updateRegionOptions(this.value);
});
document.getElementById('sel-region').addEventListener('change', function() {
  updateNetworkOptions(this.value);
});

function updateRegionOptions(jurisdiction) {
  const sel = document.getElementById('sel-region');
  sel.innerHTML = '<option value="">-- Select --</option>';
  Object.keys(REGIONS).forEach(r => {
    const opt = document.createElement('option');
    opt.value = r.toLowerCase();
    opt.textContent = r;
    if (r.toLowerCase() === (state.header.region || '').toLowerCase()) opt.selected = true;
    sel.appendChild(opt);
  });
}

function updateNetworkOptions(regionVal) {
  const sel = document.getElementById('sel-network');
  sel.innerHTML = '<option value="">-- Select --</option>';
  const key = Object.keys(REGIONS).find(k => k.toLowerCase() === (regionVal || '').toLowerCase());
  if (key) {
    REGIONS[key].forEach(n => {
      const opt = document.createElement('option');
      opt.value = n.toLowerCase();
      opt.textContent = n;
      if (n.toLowerCase() === (state.header.network || '').toLowerCase()) opt.selected = true;
      sel.appendChild(opt);
    });
  }
}

function saveHeader() {
  document.querySelectorAll('[data-header]').forEach(el => {
    state.header[el.dataset.header] = el.value;
  });
}

function startInspection() {
  saveHeader();
  state.currentPole = 0;
  showScreen('pole');
}

// ── Pole Form ────────────────────────────────────────────────

function renderPole() {
  const idx = state.currentPole;
  const pole = state.poles[idx];
  if (!pole) return;

  document.getElementById('pole-counter').textContent = `${idx + 1} / ${state.poles.length}`;
  document.getElementById('btn-prev').disabled = idx === 0;
  document.getElementById('btn-next').disabled = idx === state.poles.length - 1;
  document.getElementById('btn-prev-bottom').disabled = idx === 0;
  const isLast = idx === state.poles.length - 1;
  document.getElementById('btn-next-bottom').textContent = isLast ? 'Finish' : 'Next Pole';

  updateProgress();

  const container = document.getElementById('pole-form-container');
  container.innerHTML = '';

  const badge = document.createElement('div');
  badge.className = 'form-card';
  badge.innerHTML = `<div class="form-section" style="padding:12px 16px;display:flex;align-items:center;justify-content:space-between">
    <div><span style="font-size:13px;color:var(--text-tertiary)">Pole Item #</span>
    <strong style="font-size:22px;margin-left:8px">${pole.A || (idx+1)}</strong></div>
    <div style="font-size:13px;color:var(--text-tertiary)">Row ${pole._row}</div>
  </div>`;
  container.appendChild(badge);

  POLE_SECTIONS.forEach(section => {
    const card = document.createElement('div');
    card.className = 'form-card';

    const filledCount = section.fields.filter(f => {
      const v = pole[f.key];
      return v != null && v !== '' && v !== 0;
    }).length;

    const header = document.createElement('div');
    header.className = 'section-header';
    header.innerHTML = `<h3>${section.title}<span class="section-badge ${filledCount ? '' : 'empty'}">${filledCount}/${section.fields.length}</span></h3>
      <span class="section-arrow">&#9660;</span>`;

    const body = document.createElement('div');
    body.className = 'section-body';

    header.addEventListener('click', () => {
      header.classList.toggle('collapsed');
      body.classList.toggle('collapsed');
    });

    section.fields.forEach(field => {
      const row = document.createElement('div');
      row.className = 'pole-row';
      row.innerHTML = `<div class="pole-row-label">${field.label}</div><div class="pole-row-input" id="field-${field.key}"></div>`;
      body.appendChild(row);
      setTimeout(() => renderFieldInput(field, pole), 0);
    });

    card.appendChild(header);
    card.appendChild(body);
    container.appendChild(card);
  });

  container.scrollTop = 0;
}

function renderFieldInput(field, pole) {
  const wrap = document.getElementById('field-' + field.key);
  if (!wrap) return;
  const val = pole[field.key] != null ? pole[field.key] : '';

  switch (field.type) {
    case 'text': {
      if (field.gps) {
        wrap.innerHTML = `<div class="gps-row">
          <input type="text" value="${escHtml(val)}" data-col="${field.key}" placeholder="${field.placeholder || ''}"
                 onchange="savePoleField('${field.key}', this.value)">
          <button class="gps-btn" onclick="fillGPS('${field.key}')" title="Get GPS">&#128205;</button>
        </div>`;
      } else {
        wrap.innerHTML = `<input type="text" value="${escHtml(val)}" data-col="${field.key}" placeholder="${field.placeholder || ''}"
               onchange="savePoleField('${field.key}', this.value)">`;
      }
      break;
    }
    case 'number': {
      wrap.innerHTML = `<input type="number" inputmode="decimal" value="${val}" data-col="${field.key}"
             onchange="savePoleField('${field.key}', this.value ? Number(this.value) : '')">`;
      break;
    }
    case 'yn': {
      wrap.innerHTML = `<div class="segment-control">${
        ['Y','N'].map(o => `<button class="segment-btn ${val && String(val).toUpperCase() === o ? (o === 'Y' ? 'active-yes' : 'active-no') : ''}"
          onclick="savePoleField('${field.key}','${o}'); renderPole()">${o === 'Y' ? 'Yes' : 'No'}</button>`).join('')
      }</div>`;
      break;
    }
    case 'segment': {
      const labels = field.labels || field.options;
      wrap.innerHTML = `<div class="segment-control">${
        field.options.map((o, i) => `<button class="segment-btn ${String(val).toUpperCase() === o.toUpperCase() ? 'active' : ''}"
          onclick="savePoleField('${field.key}','${o}'); renderPole()">${labels[i]}</button>`).join('')
      }</div>`;
      break;
    }
    case 'qty': {
      const numVal = val === '' || val == null ? '' : Number(val);
      wrap.innerHTML = `<div class="qty-field">
        <button class="qty-btn" onclick="adjustQty('${field.key}', -1)">−</button>
        <input class="qty-input" type="number" inputmode="numeric" value="${numVal}" data-col="${field.key}"
               onchange="savePoleField('${field.key}', this.value ? Number(this.value) : '')">
        <button class="qty-btn" onclick="adjustQty('${field.key}', 1)">+</button>
      </div>`;
      break;
    }
    case 'select': {
      wrap.innerHTML = `<select onchange="savePoleField('${field.key}', this.value)" style="width:100%;border:1.5px solid var(--gray-4);border-radius:8px;padding:10px 12px;font-size:16px;background:white;height:44px">
        ${field.options.map(o => `<option value="${o}" ${String(val).toUpperCase() === o.toUpperCase() ? 'selected' : ''}>${o || '-- None --'}</option>`).join('')}
      </select>`;
      break;
    }
    case 'textarea': {
      wrap.innerHTML = `<textarea rows="3" onchange="savePoleField('${field.key}', this.value)"
        style="width:100%;border:1.5px solid var(--gray-4);border-radius:8px;padding:10px 12px;font-size:16px;font-family:inherit;resize:vertical">${escHtml(val)}</textarea>`;
      break;
    }
  }
}

function savePoleField(col, value) {
  state.poles[state.currentPole][col] = value;
  autoSave();
}

function adjustQty(col, delta) {
  const pole = state.poles[state.currentPole];
  let cur = Number(pole[col]) || 0;
  cur = Math.max(0, cur + delta);
  pole[col] = cur || '';
  const input = document.querySelector(`[data-col="${col}"]`);
  if (input) input.value = cur || '';
  savePoleField(col, cur || '');
}

function escHtml(s) {
  if (s == null) return '';
  return String(s).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ── Add Pole ─────────────────────────────────────────────────

function addPole() {
  if (state.emptyRows.length === 0) {
    alert('No empty rows available. All 500 rows are in use.');
    return;
  }
  const nextRow = state.emptyRows.shift();
  const newPole = { A: nextRow.itemNum, _row: nextRow.row, _isNew: true };
  state.poles.push(newPole);
  state.currentPole = state.poles.length - 1;
  document.getElementById('pole-count-label').textContent = state.poles.length;
  autoSave();
  renderPole();
}

// ── Pole Navigation ──────────────────────────────────────────

function navigatePole(dir) {
  const next = state.currentPole + dir;
  if (next < 0 || next >= state.poles.length) {
    if (dir > 0 && state.currentPole === state.poles.length - 1) showScreen('summary');
    return;
  }
  state.currentPole = next;
  renderPole();
}

function openPolePicker() {
  document.getElementById('pole-picker-modal').hidden = false;
  renderPolePicker();
  document.getElementById('pole-search').value = '';
  document.getElementById('pole-search').focus();
}

function closePolePicker() {
  document.getElementById('pole-picker-modal').hidden = true;
}

function renderPolePicker(filter) {
  const list = document.getElementById('pole-picker-list');
  list.innerHTML = '';
  const q = (filter || '').toLowerCase();
  state.poles.forEach((pole, i) => {
    const num = String(pole.A || (i+1));
    const dloc = String(pole.C || '');
    if (q && !num.includes(q) && !dloc.toLowerCase().includes(q)) return;
    const hasData = poleHasData(pole);
    const item = document.createElement('div');
    item.className = 'pole-picker-item' + (i === state.currentPole ? ' current' : '');
    item.innerHTML = `<div><span class="pole-num">#${num}</span> <span class="pole-dloc">${dloc}</span></div>
      <div class="pole-status-dot ${hasData ? 'has-data' : 'empty'}"></div>`;
    item.onclick = () => { state.currentPole = i; closePolePicker(); renderPole(); };
    list.appendChild(item);
  });
}

function filterPoles() { renderPolePicker(document.getElementById('pole-search').value); }

function poleHasData(pole) {
  return POLE_SECTIONS.some(s => s.fields.some(f => {
    if (f.key === 'C') return false; // DLOC doesn't count as "inspected"
    const v = pole[f.key];
    return v != null && v !== '' && v !== 0;
  }));
}

function poleFilledCount(pole) {
  let count = 0;
  POLE_SECTIONS.forEach(s => s.fields.forEach(f => {
    const v = pole[f.key];
    if (v != null && v !== '' && v !== 0) count++;
  }));
  return count;
}

function updateProgress() {
  const total = state.poles.length;
  const done = state.poles.filter(p => poleHasData(p)).length;
  const pct = total ? Math.round(done / total * 100) : 0;
  document.getElementById('progress-fill').style.width = pct + '%';
  document.getElementById('progress-text').textContent = `${done}/${total} poles (${pct}%)`;
}

// ── GPS ──────────────────────────────────────────────────────

function fillGPS(col) {
  if (!navigator.geolocation) { alert('GPS not available'); return; }
  navigator.geolocation.getCurrentPosition(
    pos => {
      const coords = pos.coords.latitude.toFixed(6) + ', ' + pos.coords.longitude.toFixed(6);
      savePoleField(col, coords);
      const input = document.querySelector(`[data-col="${col}"]`);
      if (input) input.value = coords;
    },
    err => alert('GPS error: ' + err.message),
    { enableHighAccuracy: true, timeout: 10000 }
  );
}

// ── Export (JSZip – preserves formatting) ─────────────────────

async function exportFile() {
  if (!state.rawFile) {
    alert('Original file not available. Please re-import the .xlsx file before exporting.');
    return;
  }

  saveHeader();

  try {
    const zip = await JSZip.loadAsync(state.rawFile);
    const sheetPath = await findSheetPath(zip, SHEET_NAME);
    const xml = await zip.file(sheetPath).async('string');

    const parser = new DOMParser();
    const doc = parser.parseFromString(xml, 'application/xml');
    const ns = doc.documentElement.namespaceURI;

    // Write header cells
    for (const [key, def] of Object.entries(HEADER_CELLS)) {
      const val = state.header[key];
      if (val == null || val === '') continue;
      const match = def.c.match(/^([A-Z]+)(\d+)$/);
      const col = match[1], row = parseInt(match[2]);
      if (key === 'completionDate') {
        setCellValue(doc, ns, row, col, jsDateToExcelSerial(val), true);
      } else if (['reconductorCost','relocationCost','sectCost'].includes(key)) {
        setCellValue(doc, ns, row, col, Number(val) || 0, true);
      } else {
        setCellValue(doc, ns, row, col, String(val), false);
      }
    }

    // Write pole data
    state.poles.forEach(pole => {
      const r = pole._row;
      DATA_COLS.forEach(col => {
        if (col === 'A') return;
        const val = pole[col];
        if (val == null || val === '') return;
        const isNum = typeof val === 'number' || (typeof val === 'string' && /^-?\d+\.?\d*$/.test(val.trim()) && val.trim() !== '');
        setCellValue(doc, ns, r, col, val, isNum);
      });
    });

    // Serialize back to XML
    const serializer = new XMLSerializer();
    let newXml = serializer.serializeToString(doc);
    if (!newXml.startsWith('<?xml')) {
      newXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + newXml;
    }
    zip.file(sheetPath, newXml);

    // Generate file
    const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    downloadBlob(blob, state.fileName);

  } catch (err) {
    alert('Export error: ' + err.message);
    console.error(err);
  }
}

async function findSheetPath(zip, sheetName) {
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const wbDoc = new DOMParser().parseFromString(wbXml, 'application/xml');
  const wbNs = wbDoc.documentElement.namespaceURI;
  const rNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  let rId = null;
  const sheets = wbDoc.getElementsByTagNameNS(wbNs, 'sheet');
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getAttribute('name') === sheetName) {
      rId = sheets[i].getAttributeNS(rNs, 'id') || sheets[i].getAttribute('r:id');
      break;
    }
  }
  if (!rId) throw new Error('Sheet not found: ' + sheetName);

  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const relsDoc = new DOMParser().parseFromString(relsXml, 'application/xml');
  const rels = relsDoc.getElementsByTagName('Relationship');
  for (let i = 0; i < rels.length; i++) {
    if (rels[i].getAttribute('Id') === rId) {
      const target = rels[i].getAttribute('Target');
      return target.startsWith('/') ? target.substring(1) : 'xl/' + target;
    }
  }
  throw new Error('Sheet file not found for rId: ' + rId);
}

function setCellValue(doc, ns, rowNum, colRef, value, isNumber) {
  const sheetData = doc.getElementsByTagNameNS(ns, 'sheetData')[0];
  const cellRef = colRef + rowNum;

  // Find or create the row
  let rowEl = null;
  const rows = sheetData.getElementsByTagNameNS(ns, 'row');
  for (let i = 0; i < rows.length; i++) {
    if (parseInt(rows[i].getAttribute('r')) === rowNum) { rowEl = rows[i]; break; }
  }
  if (!rowEl) {
    rowEl = doc.createElementNS(ns, 'row');
    rowEl.setAttribute('r', String(rowNum));
    let inserted = false;
    for (let i = 0; i < rows.length; i++) {
      if (parseInt(rows[i].getAttribute('r')) > rowNum) {
        sheetData.insertBefore(rowEl, rows[i]);
        inserted = true; break;
      }
    }
    if (!inserted) sheetData.appendChild(rowEl);
  }

  // Find or create the cell
  let cellEl = null;
  const cells = rowEl.getElementsByTagNameNS(ns, 'c');
  for (let i = 0; i < cells.length; i++) {
    if (cells[i].getAttribute('r') === cellRef) { cellEl = cells[i]; break; }
  }
  if (!cellEl) {
    cellEl = doc.createElementNS(ns, 'c');
    cellEl.setAttribute('r', cellRef);
    let inserted = false;
    for (let i = 0; i < cells.length; i++) {
      const existRef = cells[i].getAttribute('r');
      const existCol = existRef.replace(/\d+/, '');
      if (colIdx(existCol) > colIdx(colRef)) {
        rowEl.insertBefore(cellEl, cells[i]);
        inserted = true; break;
      }
    }
    if (!inserted) rowEl.appendChild(cellEl);
  }

  // Skip formula cells
  const formula = cellEl.getElementsByTagNameNS(ns, 'f')[0];
  if (formula) return;

  // Preserve the style attribute (s="...") – it stays on the element
  // Clear existing value/inline-string children
  const toRemove = [];
  for (let i = 0; i < cellEl.childNodes.length; i++) {
    const child = cellEl.childNodes[i];
    if (child.nodeType === 1) {
      const tag = child.localName;
      if (tag === 'v' || tag === 'is') toRemove.push(child);
    }
  }
  toRemove.forEach(c => cellEl.removeChild(c));

  if (isNumber) {
    cellEl.removeAttribute('t');
    const vEl = doc.createElementNS(ns, 'v');
    vEl.textContent = String(value);
    cellEl.appendChild(vEl);
  } else {
    cellEl.setAttribute('t', 'inlineStr');
    const isEl = doc.createElementNS(ns, 'is');
    const tEl = doc.createElementNS(ns, 't');
    tEl.textContent = String(value);
    isEl.appendChild(tEl);
    cellEl.appendChild(isEl);
  }
}

function downloadBlob(blob, filename) {
  // Re-wrap blob with explicit xlsx MIME type so Safari doesn't save as .zip
  const xlsxBlob = new Blob([blob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(xlsxBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  // Delay revoke so Safari has time to start the download
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ── Summary Screen ───────────────────────────────────────────

function renderSummary() {
  saveHeader();

  const total = state.poles.length;
  const withData = state.poles.filter(p => poleHasData(p)).length;
  const totalFields = POLE_SECTIONS.reduce((a, s) => a + s.fields.length, 0);
  const filledFields = state.poles.reduce((a, p) => a + poleFilledCount(p), 0);

  let badPoles = 0, workItems = 0, urgentCount = 0;
  state.poles.forEach(p => {
    if (Number(p.K) > 0) badPoles += Number(p.K);
    if (p.AM === 'URGENT') urgentCount++;
    ['M','N','O','P','Q','R','S','T','V','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ'].forEach(c => {
      if (Number(p[c]) > 0) workItems += Number(p[c]);
    });
  });

  document.getElementById('summary-stats').innerHTML = `
    <div class="stat-card"><div class="stat-value">${total}</div><div class="stat-label">Total Poles</div></div>
    <div class="stat-card"><div class="stat-value">${withData}</div><div class="stat-label">Poles Inspected</div></div>
    <div class="stat-card"><div class="stat-value">${badPoles}</div><div class="stat-label">Bad Poles</div></div>
    <div class="stat-card"><div class="stat-value">${workItems}</div><div class="stat-label">Work Items</div></div>
    <div class="stat-card"><div class="stat-value">${urgentCount}</div><div class="stat-label">Urgent</div></div>
    <div class="stat-card"><div class="stat-value">${Math.round(filledFields / (total * totalFields) * 100) || 0}%</div><div class="stat-label">Fields Filled</div></div>
  `;

  const grid = document.getElementById('pole-status-grid');
  grid.innerHTML = '';
  state.poles.forEach((pole, i) => {
    const filled = poleFilledCount(pole);
    const totalF = POLE_SECTIONS.reduce((a, s) => a + s.fields.length, 0);
    let cls = 'empty';
    if (filled > totalF * 0.5) cls = 'complete';
    else if (filled > 0) cls = 'partial';

    const dot = document.createElement('div');
    dot.className = 'pole-dot ' + cls;
    dot.textContent = pole.A || (i+1);
    dot.title = `Pole ${pole.A || (i+1)}: ${filled}/${totalF} fields`;
    dot.onclick = () => { state.currentPole = i; showScreen('pole'); };
    grid.appendChild(dot);
  });
}

// ── IndexedDB persistence (stores rawFile + form data) ───────

const DB_NAME = 'PoleInspectionDB';
const DB_VERSION = 1;
const STORE_NAME = 'sessions';
const SESSION_KEY = 'current';

function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE_NAME);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbSave(data) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).put(data, SESSION_KEY);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}

async function dbLoad() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const req = tx.objectStore(STORE_NAME).get(SESSION_KEY);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

async function dbClear() {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).delete(SESSION_KEY);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}

// ── Auto-save (IndexedDB – includes raw file) ───────────────

let _saveTimeout = null;

function autoSave() {
  // Debounce: save at most once per second
  if (_saveTimeout) clearTimeout(_saveTimeout);
  _saveTimeout = setTimeout(() => _doAutoSave(), 1000);
}

async function _doAutoSave() {
  const statusEl = document.getElementById('save-status');
  try {
    if (statusEl) { statusEl.textContent = 'Saving...'; statusEl.className = 'save-status saving'; }
    await dbSave({
      header: state.header,
      poles: state.poles,
      emptyRows: state.emptyRows,
      currentPole: state.currentPole,
      fileName: state.fileName,
      rawFile: state.rawFile,  // ArrayBuffer stored in IndexedDB (no size limit)
      savedAt: new Date().toISOString(),
    });
    if (statusEl) {
      statusEl.textContent = 'Auto-saved';
      statusEl.className = 'save-status';
      setTimeout(() => { statusEl.textContent = ''; }, 3000);
    }
  } catch(e) {
    if (statusEl) { statusEl.textContent = 'Save failed'; statusEl.className = 'save-status error'; }
  }
}

async function checkResume() {
  try {
    const data = await dbLoad();
    if (data && data.poles && data.poles.length > 0) {
      const btn = document.getElementById('btn-resume');
      btn.hidden = false;
      const ago = timeSince(new Date(data.savedAt));
      btn.textContent = `Resume: ${data.fileName} (${data.poles.length} poles, saved ${ago})`;
      btn.onclick = () => {
        state.header = data.header;
        state.poles = data.poles;
        state.emptyRows = data.emptyRows || [];
        state.currentPole = data.currentPole || 0;
        state.fileName = data.fileName;
        state.rawFile = data.rawFile || null;
        populateHeaderForm();
        document.getElementById('pole-count-label').textContent = state.poles.length;
        showScreen('header');
      };
    }
  } catch(e) {}
}

function timeSince(date) {
  const secs = Math.floor((new Date() - date) / 1000);
  if (secs < 60) return 'just now';
  const mins = Math.floor(secs / 60);
  if (mins < 60) return mins + 'm ago';
  const hrs = Math.floor(mins / 60);
  if (hrs < 24) return hrs + 'h ago';
  return Math.floor(hrs / 24) + 'd ago';
}

// ── Manual Save (download .xlsx with same filename) ──────────

async function saveFile() {
  saveHeader();
  await exportFile();  // reuses the export logic, now with original filename
}

// ── Service Worker ───────────────────────────────────────────

if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}

// ── Init ─────────────────────────────────────────────────────

checkResume();
