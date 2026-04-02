// ============================================
// Anomaly Quantix - Spreadsheet Engine
// ============================================

const ROWS = 1000;
const COLS = 26;
const VISIBLE_BUFFER = 10; // extra rows to render above/below viewport
const COL_LETTERS = Array.from({ length: COLS }, (_, i) => String.fromCharCode(65 + i));

// State
let sheets = [];
let activeSheet = 0;
let selectedCell = { row: 0, col: 0 };
let selectionRange = null; // { startRow, startCol, endRow, endCol }
let isSelecting = false;
let editingCell = null;
let undoStack = [];
let redoStack = [];
let clipboardData = null;
let resizingCol = null;

// ============================================
// Initialization
// ============================================

function createSheetData(name) {
  const cells = {};
  const colWidths = {};
  for (let c = 0; c < COLS; c++) colWidths[c] = 80;
  return { name, cells, colWidths };
}

function init() {
  sheets = [createSheetData('Sheet 1')];
  activeSheet = 0;
  renderSheetTabs();
  renderSheet();
  renderTemplates();
  selectCell(0, 0);
  autoSave();

  // Try restore from localStorage
  const saved = localStorage.getItem('quantix-autosave');
  if (saved) {
    try {
      const data = JSON.parse(saved);
      sheets = data.sheets;
      activeSheet = data.activeSheet || 0;
      document.getElementById('file-name').value = data.fileName || 'Untitled Spreadsheet';
      renderSheetTabs();
      renderSheet();
      selectCell(0, 0);
    } catch (e) { /* ignore */ }
  }
}

// ============================================
// Rendering
// ============================================

const ROW_HEIGHT = 27;
let visibleStart = 0;
let visibleEnd = 60;
let renderScheduled = false;

function renderSheet() {
  const sheet = sheets[activeSheet];
  const thead = document.getElementById('sheet-head');
  const tbody = document.getElementById('sheet-body');
  const container = document.getElementById('sheet-container');

  // Header
  let headHTML = '<tr><th></th>';
  for (let c = 0; c < COLS; c++) {
    const w = sheet.colWidths[c] || 80;
    headHTML += `<th style="width:${w}px;min-width:${w}px;max-width:${w}px" data-col="${c}">
      ${COL_LETTERS[c]}
      <div class="col-resize" data-col="${c}"></div>
    </th>`;
  }
  headHTML += '</tr>';
  thead.innerHTML = headHTML;

  // Calculate visible range
  const scrollTop = container.scrollTop;
  const viewHeight = container.clientHeight;
  visibleStart = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - VISIBLE_BUFFER);
  visibleEnd = Math.min(ROWS, Math.ceil((scrollTop + viewHeight) / ROW_HEIGHT) + VISIBLE_BUFFER);

  // Body with spacers
  let bodyHTML = '';

  // Top spacer
  if (visibleStart > 0) {
    bodyHTML += `<tr style="height:${visibleStart * ROW_HEIGHT}px"><td colspan="${COLS + 1}"></td></tr>`;
  }

  for (let r = visibleStart; r < visibleEnd; r++) {
    bodyHTML += `<tr><td>${r + 1}</td>`;
    for (let c = 0; c < COLS; c++) {
      const key = cellKey(r, c);
      const cell = sheet.cells[key] || {};

      // Skip cells merged into another
      if (cell._mergedInto) {
        bodyHTML += `<td data-row="${r}" data-col="${c}" style="display:none"></td>`;
        continue;
      }

      const display = getDisplayValue(cell);
      const style = getCellStyle(cell);
      const type = detectType(cell);
      const dropdownClass = cell.validation && cell.validation.type === 'list' ? ' has-dropdown' : '';

      // Merge attributes
      let mergeAttr = '';
      if (cell.merge) {
        const rs = cell.merge.r2 - cell.merge.r1 + 1;
        const cs = cell.merge.c2 - cell.merge.c1 + 1;
        mergeAttr = ` rowspan="${rs}" colspan="${cs}"`;
      }

      bodyHTML += `<td data-row="${r}" data-col="${c}"${mergeAttr}>
        <div class="cell${dropdownClass}" data-type="${type}" style="${style}">${escapeHTML(display)}</div>
      </td>`;
    }
    bodyHTML += '</tr>';
  }

  // Bottom spacer
  const remaining = ROWS - visibleEnd;
  if (remaining > 0) {
    bodyHTML += `<tr style="height:${remaining * ROW_HEIGHT}px"><td colspan="${COLS + 1}"></td></tr>`;
  }

  tbody.innerHTML = bodyHTML;

  attachCellEvents();
  attachResizeEvents();

  // Remove old scroll listener and re-add
  container.removeEventListener('scroll', onSheetScroll);
  container.addEventListener('scroll', onSheetScroll);
}

function onSheetScroll() {
  if (renderScheduled) return;
  renderScheduled = true;
  requestAnimationFrame(() => {
    renderScheduled = false;
    const container = document.getElementById('sheet-container');
    const scrollTop = container.scrollTop;
    const viewHeight = container.clientHeight;
    const newStart = Math.max(0, Math.floor(scrollTop / ROW_HEIGHT) - VISIBLE_BUFFER);
    const newEnd = Math.min(ROWS, Math.ceil((scrollTop + viewHeight) / ROW_HEIGHT) + VISIBLE_BUFFER);

    // Only re-render if range changed significantly
    if (Math.abs(newStart - visibleStart) > 5 || Math.abs(newEnd - visibleEnd) > 5) {
      visibleStart = newStart;
      visibleEnd = newEnd;
      renderSheet();
      // Re-highlight selection
      if (selectionRange) highlightRange();
      else {
        const td = getCellTd(selectedCell.row, selectedCell.col);
        if (td) td.classList.add('selected');
      }
    }
  });
}

function renderSheetTabs() {
  const list = document.getElementById('tabs-list');
  list.innerHTML = sheets.map((s, i) => `
    <div class="sheet-tab ${i === activeSheet ? 'active' : ''}" onclick="switchSheet(${i})" ondblclick="renameSheet(${i})">
      ${escapeHTML(s.name)}
      ${sheets.length > 1 ? `<span class="tab-close" onclick="event.stopPropagation();deleteSheet(${i})">&times;</span>` : ''}
    </div>
  `).join('');
}

function renderTemplates() {
  const grid = document.getElementById('templates-grid');
  grid.innerHTML = TEMPLATES.map((t, i) => `
    <div class="template-card" onclick="applyTemplate(${i})">
      <h3>${t.name}</h3>
      <p>${t.description}</p>
    </div>
  `).join('');
}

// ============================================
// Cell Events
// ============================================

function attachCellEvents() {
  const tbody = document.getElementById('sheet-body');

  tbody.addEventListener('mousedown', (e) => {
    const td = e.target.closest('td[data-row]');
    if (!td) return;

    const row = +td.dataset.row;
    const col = +td.dataset.col;

    if (editingCell) commitEdit();

    if (e.shiftKey) {
      selectionRange = {
        startRow: selectedCell.row, startCol: selectedCell.col,
        endRow: row, endCol: col
      };
      highlightRange();
    } else {
      selectCell(row, col);
      isSelecting = true;
      selectionRange = { startRow: row, startCol: col, endRow: row, endCol: col };
    }
  });

  tbody.addEventListener('mousemove', (e) => {
    if (!isSelecting) return;
    const td = e.target.closest('td[data-row]');
    if (!td) return;
    selectionRange.endRow = +td.dataset.row;
    selectionRange.endCol = +td.dataset.col;
    highlightRange();
    updateStatusBar();
  });

  document.addEventListener('mouseup', () => { isSelecting = false; });

  tbody.addEventListener('dblclick', (e) => {
    const td = e.target.closest('td[data-row]');
    if (!td) return;
    const row = +td.dataset.row;
    const col = +td.dataset.col;
    const cell = sheets[activeSheet].cells[cellKey(row, col)];
    // Show dropdown if cell has list validation
    if (cell && cell.validation && cell.validation.type === 'list') {
      showCellDropdown(td, row, col, cell.validation.options);
      return;
    }
    startEdit(row, col);
  });

  // Row header selection
  tbody.querySelectorAll('td:first-child').forEach(td => {
    td.addEventListener('click', (e) => {
      if (td.dataset.row !== undefined) return;
      const row = td.parentElement.rowIndex - 1;
      selectRow(row);
    });
  });

  // Context menu
  tbody.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    const td = e.target.closest('td[data-row]');
    if (td) showContextMenu(e.clientX, e.clientY, +td.dataset.row, +td.dataset.col);
  });
}

function attachResizeEvents() {
  document.querySelectorAll('.col-resize').forEach(handle => {
    handle.addEventListener('mousedown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const col = +handle.dataset.col;
      const th = handle.parentElement;
      const startX = e.clientX;
      const startW = th.offsetWidth;

      handle.classList.add('resizing');

      const onMove = (e2) => {
        const newW = Math.max(40, startW + (e2.clientX - startX));
        sheets[activeSheet].colWidths[col] = newW;
        th.style.width = newW + 'px';
        th.style.minWidth = newW + 'px';
        th.style.maxWidth = newW + 'px';
        // Update body cells
        document.querySelectorAll(`td[data-col="${col}"]`).forEach(td => {
          td.style.width = newW + 'px';
          td.style.minWidth = newW + 'px';
          td.style.maxWidth = newW + 'px';
        });
      };

      const onUp = () => {
        handle.classList.remove('resizing');
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
      };

      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
  });
}

// ============================================
// Selection
// ============================================

function selectCell(row, col) {
  selectedCell = { row, col };
  selectionRange = null;

  // Update UI
  document.querySelectorAll('td.selected').forEach(td => td.classList.remove('selected'));
  document.querySelectorAll('td.in-range').forEach(td => td.classList.remove('in-range'));
  document.querySelectorAll('th.col-selected').forEach(th => th.classList.remove('col-selected'));
  document.querySelectorAll('td.row-selected').forEach(td => td.classList.remove('row-selected'));

  const td = getCellTd(row, col);
  if (td) td.classList.add('selected');

  // Update cell ref
  document.getElementById('cell-ref').textContent = COL_LETTERS[col] + (row + 1);

  // Update formula bar
  const key = cellKey(row, col);
  const cell = sheets[activeSheet].cells[key] || {};
  document.getElementById('formula-input').value = cell.formula || cell.value || '';

  // Update toolbar state
  updateToolbarState(cell);
  updateStatusBar();

  // Scroll into view
  if (td) td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

function highlightRange() {
  document.querySelectorAll('td.in-range').forEach(td => td.classList.remove('in-range'));
  document.querySelectorAll('td.selected').forEach(td => td.classList.remove('selected'));

  if (!selectionRange) return;

  const r1 = Math.min(selectionRange.startRow, selectionRange.endRow);
  const r2 = Math.max(selectionRange.startRow, selectionRange.endRow);
  const c1 = Math.min(selectionRange.startCol, selectionRange.endCol);
  const c2 = Math.max(selectionRange.startCol, selectionRange.endCol);

  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const td = getCellTd(r, c);
      if (td) td.classList.add(r === selectedCell.row && c === selectedCell.col ? 'selected' : 'in-range');
    }
  }
}

function selectRow(row) {
  selectionRange = { startRow: row, startCol: 0, endRow: row, endCol: COLS - 1 };
  selectCell(row, 0);
  highlightRange();
}

// ============================================
// Editing
// ============================================

function startEdit(row, col) {
  if (editingCell) commitEdit();
  editingCell = { row, col };

  const td = getCellTd(row, col);
  if (!td) return;

  const key = cellKey(row, col);
  const cell = sheets[activeSheet].cells[key] || {};
  const raw = cell.formula || cell.value || '';

  const cellDiv = td.querySelector('.cell');
  cellDiv.innerHTML = `<input class="cell-input" value="${escapeAttr(raw)}" />`;
  const input = cellDiv.querySelector('input');
  input.focus();
  input.setSelectionRange(raw.length, raw.length);

  input.addEventListener('blur', () => commitEdit());
}

function commitEdit() {
  if (!editingCell) return;
  const { row, col } = editingCell;
  const td = getCellTd(row, col);
  editingCell = null;
  if (!td) return;

  const input = td.querySelector('input');
  if (!input) return;

  const raw = input.value;
  setCellValue(row, col, raw);
}

function setCellValue(row, col, raw, skipUndo) {
  const key = cellKey(row, col);
  const sheet = sheets[activeSheet];
  const prev = { ...sheet.cells[key] };

  // Validate if needed
  if (raw && sheet.cells[key]?.validation && !validateCell(row, col, raw)) {
    document.getElementById('status-info').textContent = 'Invalid value for this cell';
    return;
  }

  if (!skipUndo) {
    undoStack.push({ sheet: activeSheet, key, prev: { ...prev }, action: 'edit' });
    redoStack = [];
  }

  if (!raw && raw !== 0) {
    delete sheet.cells[key];
  } else if (typeof raw === 'string' && raw.startsWith('=')) {
    sheet.cells[key] = { ...(sheet.cells[key] || {}), formula: raw, value: evaluateFormula(raw, activeSheet) };
  } else {
    const parsed = autoDetect(raw);
    sheet.cells[key] = { ...(sheet.cells[key] || {}), value: parsed.value, formula: undefined, detectedType: parsed.type };
  }

  refreshCell(row, col);
  recalcDependents();
  updateFormulaBar();
  updateStatusBar();
  triggerAutoSave();
}

function refreshCell(row, col) {
  const td = getCellTd(row, col);
  if (!td) return;
  const key = cellKey(row, col);
  const cell = sheets[activeSheet].cells[key] || {};
  const display = getDisplayValue(cell);
  const style = getCellStyle(cell);
  const type = detectType(cell);
  td.querySelector('.cell').outerHTML = `<div class="cell" data-type="${type}" style="${style}">${escapeHTML(display)}</div>`;
}

function updateFormulaBar() {
  const key = cellKey(selectedCell.row, selectedCell.col);
  const cell = sheets[activeSheet].cells[key] || {};
  document.getElementById('formula-input').value = cell.formula || cell.value || '';
}

// Formula bar input
document.getElementById('formula-input').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    const raw = e.target.value;
    setCellValue(selectedCell.row, selectedCell.col, raw);
    moveSelection(1, 0);
    e.target.blur();
  } else if (e.key === 'Escape') {
    updateFormulaBar();
    e.target.blur();
  }
});

// ============================================
// Formula Engine
// ============================================

function evaluateFormula(formula, sheetIdx) {
  try {
    const expr = formula.substring(1).toUpperCase();
    const sheet = sheets[sheetIdx];

    // Handle VLOOKUP specially before resolving references
    const vlookupMatch = expr.match(/^VLOOKUP\((.+)$/);
    if (vlookupMatch) {
      return evalVlookup(expr, sheet, sheetIdx);
    }

    // Replace cell references with values
    const resolved = expr.replace(/\b([A-Z])(\d+):([A-Z])(\d+)\b/g, (_, c1, r1, c2, r2) => {
      const values = getRangeValues(c1.charCodeAt(0) - 65, +r1 - 1, c2.charCodeAt(0) - 65, +r2 - 1, sheetIdx);
      return '[' + values.map(v => typeof v === 'number' ? v : `"${v}"`).join(',') + ']';
    }).replace(/\b([A-Z])(\d+)\b/g, (_, c, r) => {
      const key = cellKey(+r - 1, c.charCodeAt(0) - 65);
      const cell = sheet.cells[key];
      if (!cell) return '0';
      const v = cell.formula ? evaluateFormula(cell.formula, sheetIdx) : cell.value;
      return typeof v === 'number' ? v : `"${v || 0}"`;
    });

    // Built-in functions
    const funcs = {
      // Math
      SUM: (arr) => flat(arr).reduce((a, b) => a + toNum(b), 0),
      AVERAGE: (arr) => { const f = flat(arr).map(toNum); return f.reduce((a, b) => a + b, 0) / f.length; },
      MEDIAN: (arr) => { const s = flat(arr).map(toNum).sort((a,b) => a-b); const m = Math.floor(s.length/2); return s.length % 2 ? s[m] : (s[m-1]+s[m])/2; },
      COUNT: (arr) => flat(arr).filter(v => typeof v === 'number' || !isNaN(+v)).length,
      COUNTA: (arr) => flat(arr).filter(v => v !== '' && v !== null && v !== undefined).length,
      COUNTIF: (args) => { const [range, crit] = args; return flat(Array.isArray(range) ? range : [range]).filter(v => matchCriteria(v, crit)).length; },
      SUMIF: (args) => { const [range, crit, sumRange] = args; const r = flat(Array.isArray(range) ? range : [range]); const s = sumRange ? flat(Array.isArray(sumRange) ? sumRange : [sumRange]) : r; let total = 0; r.forEach((v, i) => { if (matchCriteria(v, crit)) total += toNum(s[i] ?? 0); }); return total; },
      AVERAGEIF: (args) => { const [range, crit, avgRange] = args; const r = flat(Array.isArray(range) ? range : [range]); const s = avgRange ? flat(Array.isArray(avgRange) ? avgRange : [avgRange]) : r; let total = 0, cnt = 0; r.forEach((v, i) => { if (matchCriteria(v, crit)) { total += toNum(s[i] ?? 0); cnt++; } }); return cnt ? total / cnt : 0; },
      MIN: (arr) => Math.min(...flat(arr).map(toNum).filter(n => !isNaN(n))),
      MAX: (arr) => Math.max(...flat(arr).map(toNum).filter(n => !isNaN(n))),
      ABS: (args) => Math.abs(toNum(args[0])),
      SQRT: (args) => Math.sqrt(toNum(args[0])),
      POWER: (args) => Math.pow(toNum(args[0]), toNum(args[1])),
      MOD: (args) => toNum(args[0]) % toNum(args[1]),
      ROUND: (args) => { const [n, d = 0] = args; return +toNum(n).toFixed(toNum(d)); },
      ROUNDUP: (args) => { const [n, d = 0] = args; const f = Math.pow(10, toNum(d)); return Math.ceil(toNum(n) * f) / f; },
      ROUNDDOWN: (args) => { const [n, d = 0] = args; const f = Math.pow(10, toNum(d)); return Math.floor(toNum(n) * f) / f; },
      CEILING: (args) => { const [n, s = 1] = args; return Math.ceil(toNum(n) / toNum(s)) * toNum(s); },
      FLOOR: (args) => { const [n, s = 1] = args; return Math.floor(toNum(n) / toNum(s)) * toNum(s); },
      RAND: () => Math.random(),
      RANDBETWEEN: (args) => { const [lo, hi] = args.map(toNum); return Math.floor(Math.random() * (hi - lo + 1)) + lo; },
      PI: () => Math.PI,

      // Logic
      IF: (args) => args[0] ? args[1] : args[2],
      AND: (arr) => flat(arr).every(Boolean),
      OR: (arr) => flat(arr).some(Boolean),
      NOT: (args) => !args[0],
      IFERROR: (args) => { try { return args[0] !== '#ERROR' && args[0] !== '#ERR' ? args[0] : args[1]; } catch { return args[1]; } },

      // Lookup
      VLOOKUP: (args) => '#USE_SPECIAL',
      INDEX: (args) => {
        const [range, rowIdx, colIdx] = args;
        if (Array.isArray(range)) return range[toNum(rowIdx) - 1] ?? '#REF';
        return range;
      },
      MATCH: (args) => {
        const [needle, range] = args;
        const arr = flat(Array.isArray(range) ? range : [range]);
        const idx = arr.findIndex(v => v == needle || String(v).toLowerCase() === String(needle).toLowerCase());
        return idx >= 0 ? idx + 1 : '#N/A';
      },

      // Text
      CONCATENATE: (args) => flat(args).join(''),
      CONCAT: (args) => flat(args).join(''),
      TEXTJOIN: (args) => { const [delim, skipEmpty, ...rest] = args; const vals = flat(rest); return (skipEmpty ? vals.filter(v => v !== '' && v != null) : vals).join(delim); },
      LEFT: (args) => String(args[0]).substring(0, toNum(args[1] ?? 1)),
      RIGHT: (args) => { const s = String(args[0]); const n = toNum(args[1] ?? 1); return s.substring(s.length - n); },
      MID: (args) => String(args[0]).substring(toNum(args[1]) - 1, toNum(args[1]) - 1 + toNum(args[2])),
      LEN: (args) => String(args[0]).length,
      UPPER: (args) => String(args[0]).toUpperCase(),
      LOWER: (args) => String(args[0]).toLowerCase(),
      PROPER: (args) => String(args[0]).replace(/\b\w/g, c => c.toUpperCase()),
      TRIM: (args) => String(args[0]).trim(),
      SUBSTITUTE: (args) => { const [text, old, rep, nth] = args; if (nth) { let i = 0; return String(text).replace(new RegExp(escapeRegex(String(old)), 'g'), m => (++i === toNum(nth)) ? rep : m); } return String(text).split(String(old)).join(String(rep)); },
      FIND: (args) => { const idx = String(args[1]).indexOf(String(args[0]), toNum(args[2] ?? 1) - 1); return idx >= 0 ? idx + 1 : '#VALUE'; },
      SEARCH: (args) => { const idx = String(args[1]).toLowerCase().indexOf(String(args[0]).toLowerCase(), toNum(args[2] ?? 1) - 1); return idx >= 0 ? idx + 1 : '#VALUE'; },
      REPT: (args) => String(args[0]).repeat(toNum(args[1])),
      TEXT: (args) => formatText(args[0], String(args[1])),

      // Date
      NOW: () => new Date().toLocaleString(),
      TODAY: () => new Date().toLocaleDateString(),
      YEAR: (args) => new Date(args[0]).getFullYear(),
      MONTH: (args) => new Date(args[0]).getMonth() + 1,
      DAY: (args) => new Date(args[0]).getDate(),
      DAYS: (args) => Math.round((new Date(args[0]) - new Date(args[1])) / 86400000),
      EDATE: (args) => { const d = new Date(args[0]); d.setMonth(d.getMonth() + toNum(args[1])); return d.toLocaleDateString(); },
    };

    function matchCriteria(val, criteria) {
      const s = String(criteria);
      if (s.startsWith('>=')) return toNum(val) >= toNum(s.slice(2));
      if (s.startsWith('<=')) return toNum(val) <= toNum(s.slice(2));
      if (s.startsWith('<>')) return String(val) !== s.slice(2);
      if (s.startsWith('>')) return toNum(val) > toNum(s.slice(1));
      if (s.startsWith('<')) return toNum(val) < toNum(s.slice(1));
      if (s.includes('*') || s.includes('?')) {
        const regex = new RegExp('^' + s.replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
        return regex.test(String(val));
      }
      return String(val).toLowerCase() === s.toLowerCase() || toNum(val) === toNum(s);
    }

    function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

    function formatText(val, fmt) {
      if (fmt === '0%') return (toNum(val) * 100).toFixed(0) + '%';
      if (fmt === '0.00') return toNum(val).toFixed(2);
      if (fmt.includes('$')) return '$' + toNum(val).toFixed(2);
      return String(val);
    }

    // Parse function calls
    let processed = resolved;
    for (const [name, fn] of Object.entries(funcs)) {
      const regex = new RegExp(`${name}\\(`, 'g');
      while (regex.test(processed)) {
        processed = processed.replace(new RegExp(`${name}\\(([^)]*?)\\)`), (_, args) => {
          try {
            const parsed = Function(`"use strict"; return [${args}]`)();
            return JSON.stringify(fn(parsed));
          } catch {
            return '#ERR';
          }
        });
        regex.lastIndex = 0;
      }
    }

    // Evaluate the expression
    const result = Function(`"use strict"; return (${processed})`)();
    return typeof result === 'number' ? (Math.round(result * 1e10) / 1e10) : result;
  } catch (e) {
    return '#ERROR';
  }
}

function evalVlookup(expr, sheet, sheetIdx) {
  // Parse: VLOOKUP(value, A1:D10, 3, FALSE)
  const inner = expr.slice(8, -1); // strip VLOOKUP( ... )
  const parts = splitTopLevel(inner);
  if (parts.length < 3) return '#ERR';

  // Resolve the lookup value
  const needle = resolveValue(parts[0].trim(), sheet, sheetIdx);
  const colIdx = toNum(resolveValue(parts[2].trim(), sheet, sheetIdx));
  const approx = parts[3] ? resolveValue(parts[3].trim(), sheet, sheetIdx) : true;

  // Parse the range
  const rangeMatch = parts[1].trim().match(/^([A-Z])(\d+):([A-Z])(\d+)$/);
  if (!rangeMatch) return '#ERR';

  const c1 = rangeMatch[1].charCodeAt(0) - 65;
  const r1 = +rangeMatch[2] - 1;
  const c2 = rangeMatch[3].charCodeAt(0) - 65;
  const r2 = +rangeMatch[4] - 1;

  // Search first column for needle
  for (let r = r1; r <= r2; r++) {
    const lookupCell = sheet.cells[cellKey(r, c1)];
    const lookupVal = lookupCell ? (lookupCell.formula ? evaluateFormula(lookupCell.formula, sheetIdx) : lookupCell.value) : '';

    if (String(lookupVal).toLowerCase() === String(needle).toLowerCase() || lookupVal == needle) {
      const resultCell = sheet.cells[cellKey(r, c1 + colIdx - 1)];
      if (!resultCell) return '';
      return resultCell.formula ? evaluateFormula(resultCell.formula, sheetIdx) : resultCell.value;
    }
  }
  return '#N/A';
}

function splitTopLevel(str) {
  const parts = [];
  let depth = 0, current = '';
  for (const ch of str) {
    if (ch === '(') depth++;
    else if (ch === ')') depth--;
    if (ch === ',' && depth === 0) { parts.push(current); current = ''; }
    else current += ch;
  }
  parts.push(current);
  return parts;
}

function resolveValue(token, sheet, sheetIdx) {
  token = token.trim();
  if (token === 'TRUE') return true;
  if (token === 'FALSE') return false;
  if (token.startsWith('"') && token.endsWith('"')) return token.slice(1, -1);
  if (!isNaN(+token)) return +token;
  // Cell reference
  const cellMatch = token.match(/^([A-Z])(\d+)$/);
  if (cellMatch) {
    const cell = sheet.cells[cellKey(+cellMatch[2] - 1, cellMatch[1].charCodeAt(0) - 65)];
    if (!cell) return '';
    return cell.formula ? evaluateFormula(cell.formula, sheetIdx) : cell.value;
  }
  return token;
}

function getRangeValues(c1, r1, c2, r2, sheetIdx) {
  const values = [];
  const sheet = sheets[sheetIdx];
  const minR = Math.min(r1, r2), maxR = Math.max(r1, r2);
  const minC = Math.min(c1, c2), maxC = Math.max(c1, c2);
  for (let r = minR; r <= maxR; r++) {
    for (let c = minC; c <= maxC; c++) {
      const cell = sheet.cells[cellKey(r, c)];
      if (cell) {
        const v = cell.formula ? evaluateFormula(cell.formula, sheetIdx) : cell.value;
        values.push(v);
      } else {
        values.push(0);
      }
    }
  }
  return values;
}

function recalcDependents() {
  const sheet = sheets[activeSheet];
  for (const [key, cell] of Object.entries(sheet.cells)) {
    if (cell.formula) {
      cell.value = evaluateFormula(cell.formula, activeSheet);
      const { row, col } = parseKey(key);
      refreshCell(row, col);
    }
  }
}

// ============================================
// Auto-detect Types
// ============================================

function autoDetect(raw) {
  if (raw === '' || raw === null || raw === undefined) return { value: '', type: 'text' };

  const s = String(raw).trim();

  // Number
  if (/^-?\d{1,3}(,\d{3})*(\.\d+)?$/.test(s)) {
    return { value: parseFloat(s.replace(/,/g, '')), type: 'number' };
  }
  if (/^-?\d+\.?\d*$/.test(s)) {
    return { value: parseFloat(s), type: 'number' };
  }

  // Currency
  if (/^\$-?\d{1,3}(,\d{3})*(\.\d{1,2})?$/.test(s)) {
    return { value: parseFloat(s.replace(/[$,]/g, '')), type: 'currency' };
  }

  // Percent
  if (/^-?\d+\.?\d*%$/.test(s)) {
    return { value: parseFloat(s) / 100, type: 'percent' };
  }

  // Date
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(s) || /^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d)) return { value: s, type: 'date' };
  }

  return { value: raw, type: 'text' };
}

function detectType(cell) {
  if (!cell || (!cell.value && cell.value !== 0)) return 'text';
  if (cell.formula) return 'formula';
  if (cell.detectedType) return cell.detectedType;
  if (cell.formatType && cell.formatType !== 'auto') return cell.formatType;
  return typeof cell.value === 'number' ? 'number' : 'text';
}

function getDisplayValue(cell) {
  if (!cell) return '';
  const val = cell.value;
  if (val === undefined || val === null || val === '') return '';

  const fmt = cell.formatType || cell.detectedType || 'auto';

  if (fmt === 'currency' || (fmt === 'auto' && cell.detectedType === 'currency')) {
    return '$' + (typeof val === 'number' ? val.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : val);
  }
  if (fmt === 'percent' || (fmt === 'auto' && cell.detectedType === 'percent')) {
    return (typeof val === 'number' ? (val * 100).toFixed(1) : val) + '%';
  }
  if (fmt === 'number' && typeof val === 'number') {
    return val.toLocaleString('en-US');
  }

  return String(val);
}

function getCellStyle(cell) {
  if (!cell) return '';
  const parts = [];
  if (cell.bold) parts.push('font-weight:700');
  if (cell.italic) parts.push('font-style:italic');
  if (cell.underline) parts.push('text-decoration:underline');
  if (cell.textColor) parts.push(`color:${cell.textColor}`);
  if (cell._condColor) parts.push(`background:${cell._condColor}33;border-left:3px solid ${cell._condColor}`);
  else if (cell.fillColor && cell.fillColor !== '#1a1a2e') parts.push(`background:${cell.fillColor}`);
  if (cell.align) parts.push(`justify-content:${cell.align === 'left' ? 'flex-start' : cell.align === 'right' ? 'flex-end' : 'center'}`);
  return parts.join(';');
}

// ============================================
// Formatting
// ============================================

function applyToSelection(fn) {
  forEachSelected((r, c) => {
    const key = cellKey(r, c);
    if (!sheets[activeSheet].cells[key]) sheets[activeSheet].cells[key] = {};
    fn(sheets[activeSheet].cells[key], r, c);
    refreshCell(r, c);
  });
  triggerAutoSave();
}

function toggleBold() {
  const key = cellKey(selectedCell.row, selectedCell.col);
  const cell = sheets[activeSheet].cells[key] || {};
  const newVal = !cell.bold;
  applyToSelection(c => c.bold = newVal);
  updateToolbarState(sheets[activeSheet].cells[key] || {});
}

function toggleItalic() {
  const key = cellKey(selectedCell.row, selectedCell.col);
  const cell = sheets[activeSheet].cells[key] || {};
  const newVal = !cell.italic;
  applyToSelection(c => c.italic = newVal);
  updateToolbarState(sheets[activeSheet].cells[key] || {});
}

function toggleUnderline() {
  const key = cellKey(selectedCell.row, selectedCell.col);
  const cell = sheets[activeSheet].cells[key] || {};
  const newVal = !cell.underline;
  applyToSelection(c => c.underline = newVal);
  updateToolbarState(sheets[activeSheet].cells[key] || {});
}

function setAlign(align) {
  applyToSelection(c => c.align = align);
  document.querySelectorAll('[id^="btn-align"]').forEach(b => b.classList.remove('active'));
  document.getElementById('btn-align-' + align).classList.add('active');
}

function setTextColor(color) {
  applyToSelection(c => c.textColor = color);
  document.getElementById('text-color-bar').style.background = color;
}

function setFillColor(color) {
  applyToSelection(c => c.fillColor = color);
  document.getElementById('fill-color-bar').style.background = color;
}

function setFormatType(type) {
  applyToSelection(c => c.formatType = type);
}

function updateToolbarState(cell) {
  document.getElementById('btn-bold').classList.toggle('active', !!cell.bold);
  document.getElementById('btn-italic').classList.toggle('active', !!cell.italic);
  document.getElementById('btn-underline').classList.toggle('active', !!cell.underline);

  document.querySelectorAll('[id^="btn-align"]').forEach(b => b.classList.remove('active'));
  document.getElementById('btn-align-' + (cell.align || 'left')).classList.add('active');

  if (cell.textColor) {
    document.getElementById('text-color').value = cell.textColor;
    document.getElementById('text-color-bar').style.background = cell.textColor;
  }
  if (cell.fillColor) {
    document.getElementById('fill-color').value = cell.fillColor;
    document.getElementById('fill-color-bar').style.background = cell.fillColor;
  }
  document.getElementById('format-type').value = cell.formatType || 'auto';
}

// ============================================
// Keyboard
// ============================================

document.addEventListener('keydown', (e) => {
  // Global shortcuts
  if (e.ctrlKey && e.key === 's') { e.preventDefault(); saveFile(); return; }
  if (e.ctrlKey && e.key === 'o') { e.preventDefault(); document.getElementById('load-input').click(); return; }
  if (e.ctrlKey && e.key === 'f') { e.preventDefault(); openFindBar(); return; }
  if (e.ctrlKey && e.key === 'h') { e.preventDefault(); openFindBar(true); return; }
  if (e.key === 'Escape' && document.getElementById('find-replace-bar').style.display !== 'none') { closeFindBar(); return; }
  if (e.ctrlKey && e.key === 'z') { e.preventDefault(); undo(); return; }
  if (e.ctrlKey && e.key === 'y') { e.preventDefault(); redo(); return; }
  if (e.ctrlKey && e.key === 'b') { e.preventDefault(); toggleBold(); return; }
  if (e.ctrlKey && e.key === 'i') { e.preventDefault(); toggleItalic(); return; }
  if (e.ctrlKey && e.key === 'u') { e.preventDefault(); toggleUnderline(); return; }
  if (e.ctrlKey && e.key === 'c') { copySelection(); return; }
  if (e.ctrlKey && e.key === 'v') { pasteSelection(); return; }
  if (e.ctrlKey && e.key === 'x') { cutSelection(); return; }

  // Skip if focused on any input/select outside the sheet
  const ae = document.activeElement;
  if (ae && (ae.tagName === 'INPUT' || ae.tagName === 'SELECT' || ae.tagName === 'TEXTAREA') && !ae.classList.contains('cell-input')) return;

  // If editing, let the input handle it
  if (editingCell) {
    if (e.key === 'Enter') { e.preventDefault(); commitEdit(); moveSelection(1, 0); }
    else if (e.key === 'Escape') { cancelEdit(); }
    else if (e.key === 'Tab') { e.preventDefault(); commitEdit(); moveSelection(0, e.shiftKey ? -1 : 1); }
    return;
  }

  // Navigation
  if (e.key === 'ArrowUp') { e.preventDefault(); moveSelection(-1, 0); }
  else if (e.key === 'ArrowDown') { e.preventDefault(); moveSelection(1, 0); }
  else if (e.key === 'ArrowLeft') { e.preventDefault(); moveSelection(0, -1); }
  else if (e.key === 'ArrowRight') { e.preventDefault(); moveSelection(0, 1); }
  else if (e.key === 'Tab') { e.preventDefault(); moveSelection(0, e.shiftKey ? -1 : 1); }
  else if (e.key === 'Enter') { e.preventDefault(); startEdit(selectedCell.row, selectedCell.col); }
  else if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    deleteSelection();
  }
  else if (e.key === 'F2') { e.preventDefault(); startEdit(selectedCell.row, selectedCell.col); }
  else if (e.key === 'Home') { e.preventDefault(); selectCell(selectedCell.row, 0); }
  else if (e.key === 'End') { e.preventDefault(); selectCell(selectedCell.row, COLS - 1); }
  else if (e.key.length === 1 && !e.ctrlKey && !e.altKey && !e.metaKey) {
    // Start typing directly
    startEdit(selectedCell.row, selectedCell.col);
    const input = getCellTd(selectedCell.row, selectedCell.col)?.querySelector('input');
    if (input) { input.value = e.key; }
  }
});

function moveSelection(dr, dc) {
  const newRow = Math.max(0, Math.min(ROWS - 1, selectedCell.row + dr));
  const newCol = Math.max(0, Math.min(COLS - 1, selectedCell.col + dc));
  selectCell(newRow, newCol);
}

function cancelEdit() {
  if (!editingCell) return;
  const { row, col } = editingCell;
  editingCell = null;
  refreshCell(row, col);
}

function deleteSelection() {
  forEachSelected((r, c) => {
    const key = cellKey(r, c);
    const sheet = sheets[activeSheet];
    if (sheet.cells[key]) {
      undoStack.push({ sheet: activeSheet, key, prev: { ...sheet.cells[key] }, action: 'delete' });
      delete sheet.cells[key];
      refreshCell(r, c);
    }
  });
  redoStack = [];
  triggerAutoSave();
}

// ============================================
// Copy/Paste
// ============================================

function copySelection() {
  clipboardData = [];
  forEachSelected((r, c) => {
    const key = cellKey(r, c);
    const cell = sheets[activeSheet].cells[key];
    clipboardData.push({ row: r, col: c, cell: cell ? { ...cell } : null });
  });
}

function cutSelection() {
  copySelection();
  deleteSelection();
}

function pasteSelection() {
  if (!clipboardData || clipboardData.length === 0) return;

  const minR = Math.min(...clipboardData.map(d => d.row));
  const minC = Math.min(...clipboardData.map(d => d.col));

  clipboardData.forEach(({ row, col, cell }) => {
    const newR = selectedCell.row + (row - minR);
    const newC = selectedCell.col + (col - minC);
    if (newR < ROWS && newC < COLS) {
      const key = cellKey(newR, newC);
      if (cell) {
        sheets[activeSheet].cells[key] = { ...cell };
      }
      refreshCell(newR, newC);
    }
  });
  recalcDependents();
  triggerAutoSave();
}

// ============================================
// Undo/Redo
// ============================================

function undo() {
  if (undoStack.length === 0) return;
  const action = undoStack.pop();
  const sheet = sheets[action.sheet];
  const current = sheet.cells[action.key] ? { ...sheet.cells[action.key] } : null;

  if (action.prev && Object.keys(action.prev).length > 0) {
    sheet.cells[action.key] = action.prev;
  } else {
    delete sheet.cells[action.key];
  }

  redoStack.push({ ...action, prev: current });

  if (action.sheet === activeSheet) {
    const { row, col } = parseKey(action.key);
    refreshCell(row, col);
    recalcDependents();
  }
  updateFormulaBar();
  triggerAutoSave();
}

function redo() {
  if (redoStack.length === 0) return;
  const action = redoStack.pop();
  const sheet = sheets[action.sheet];
  const current = sheet.cells[action.key] ? { ...sheet.cells[action.key] } : null;

  if (action.prev && Object.keys(action.prev).length > 0) {
    sheet.cells[action.key] = action.prev;
  } else {
    delete sheet.cells[action.key];
  }

  undoStack.push({ ...action, prev: current });

  if (action.sheet === activeSheet) {
    const { row, col } = parseKey(action.key);
    refreshCell(row, col);
    recalcDependents();
  }
  updateFormulaBar();
  triggerAutoSave();
}

// ============================================
// Context Menu
// ============================================

function showContextMenu(x, y, row, col) {
  removeContextMenu();
  const menu = document.createElement('div');
  menu.className = 'context-menu';
  menu.style.left = x + 'px';
  menu.style.top = y + 'px';
  menu.innerHTML = `
    <button onclick="cutSelection();removeContextMenu()">Cut</button>
    <button onclick="copySelection();removeContextMenu()">Copy</button>
    <button onclick="pasteSelection();removeContextMenu()">Paste</button>
    <div class="separator"></div>
    <button onclick="insertRowAbove(${row});removeContextMenu()">Insert Row Above</button>
    <button onclick="insertRowBelow(${row});removeContextMenu()">Insert Row Below</button>
    <div class="separator"></div>
    <button onclick="deleteSelection();removeContextMenu()">Clear Contents</button>
    <button onclick="sortColumn(${col}, true);removeContextMenu()">Sort A → Z</button>
    <button onclick="sortColumn(${col}, false);removeContextMenu()">Sort Z → A</button>
    <div class="separator"></div>
    <button onclick="freezeAt(${row}, ${col});removeContextMenu()">Freeze Above & Left</button>
    <button onclick="unfreeze();removeContextMenu()">Unfreeze Panes</button>
    <div class="separator"></div>
    <button onclick="showConditionalFormat(${row}, ${col});removeContextMenu()">Conditional Format...</button>
    <button onclick="showDataValidation(${row}, ${col});removeContextMenu()">Data Validation...</button>
  `;
  document.body.appendChild(menu);

  // Adjust if off-screen
  const rect = menu.getBoundingClientRect();
  if (rect.right > window.innerWidth) menu.style.left = (x - rect.width) + 'px';
  if (rect.bottom > window.innerHeight) menu.style.top = (y - rect.height) + 'px';

  setTimeout(() => document.addEventListener('click', removeContextMenu, { once: true }), 10);
}

function removeContextMenu() {
  document.querySelectorAll('.context-menu').forEach(m => m.remove());
}

function insertRowAbove(row) {
  const sheet = sheets[activeSheet];
  const newCells = {};
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const { row: r, col: c } = parseKey(key);
    if (r >= row) {
      newCells[cellKey(r + 1, c)] = cell;
    } else {
      newCells[key] = cell;
    }
  }
  sheet.cells = newCells;
  renderSheet();
  selectCell(row, selectedCell.col);
}

function insertRowBelow(row) {
  insertRowAbove(row + 1);
}

function sortColumn(col, ascending) {
  const sheet = sheets[activeSheet];
  // Find rows with data in this column
  const rows = [];
  for (let r = 0; r < ROWS; r++) {
    const key = cellKey(r, col);
    if (sheet.cells[key] && sheet.cells[key].value !== '') {
      rows.push(r);
    }
  }
  if (rows.length === 0) return;

  const minRow = Math.min(...rows);
  const maxRow = Math.max(...rows);

  // Collect all data in the range
  const rowData = [];
  for (let r = minRow; r <= maxRow; r++) {
    const rowCells = {};
    for (let c = 0; c < COLS; c++) {
      const key = cellKey(r, c);
      if (sheet.cells[key]) rowCells[c] = { ...sheet.cells[key] };
    }
    rowData.push({ sortVal: sheet.cells[cellKey(r, col)]?.value, cells: rowCells });
  }

  rowData.sort((a, b) => {
    const va = a.sortVal ?? '';
    const vb = b.sortVal ?? '';
    const cmp = typeof va === 'number' && typeof vb === 'number' ? va - vb : String(va).localeCompare(String(vb));
    return ascending ? cmp : -cmp;
  });

  // Write back
  for (let i = 0; i < rowData.length; i++) {
    const r = minRow + i;
    for (let c = 0; c < COLS; c++) {
      const key = cellKey(r, c);
      if (rowData[i].cells[c]) {
        sheet.cells[key] = rowData[i].cells[c];
      } else {
        delete sheet.cells[key];
      }
    }
  }

  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
  triggerAutoSave();
}

// ============================================
// Sheets
// ============================================

function switchSheet(index) {
  if (editingCell) commitEdit();
  activeSheet = index;
  renderSheetTabs();
  renderSheet();
  selectCell(0, 0);
}

function addSheet() {
  sheets.push(createSheetData('Sheet ' + (sheets.length + 1)));
  switchSheet(sheets.length - 1);
}

function deleteSheet(index) {
  if (sheets.length <= 1) return;
  sheets.splice(index, 1);
  if (activeSheet >= sheets.length) activeSheet = sheets.length - 1;
  renderSheetTabs();
  renderSheet();
  selectCell(0, 0);
}

function renameSheet(index) {
  const name = prompt('Sheet name:', sheets[index].name);
  if (name && name.trim()) {
    sheets[index].name = name.trim();
    renderSheetTabs();
  }
}

// ============================================
// File Operations
// ============================================

function saveFile() {
  const data = {
    type: 'quantix',
    version: 1,
    fileName: document.getElementById('file-name').value,
    sheets,
    activeSheet
  };
  const blob = new Blob([JSON.stringify(data)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (data.fileName || 'spreadsheet') + '.qx';
  a.click();
  URL.revokeObjectURL(a.href);
  document.getElementById('status-info').textContent = 'Saved!';
  setTimeout(() => document.getElementById('status-info').textContent = 'Ready', 2000);
}

function loadFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();

  if (ext === 'xlsx' || ext === 'xls') {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        importExcel(e.target.result);
        document.getElementById('file-name').value = file.name.replace(/\.(xlsx|xls)$/, '');
        document.getElementById('status-info').textContent = 'Loaded: ' + file.name;
      } catch (err) {
        alert('Failed to load Excel file: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  } else {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        if (ext === 'csv') {
          importCSV(e.target.result);
        } else {
          const data = JSON.parse(e.target.result);
          sheets = data.sheets;
          activeSheet = data.activeSheet || 0;
          document.getElementById('file-name').value = data.fileName || file.name.replace('.qx', '');
          renderSheetTabs();
          renderSheet();
          selectCell(0, 0);
        }
        document.getElementById('status-info').textContent = 'Loaded: ' + file.name;
      } catch (err) {
        alert('Failed to load file: ' + err.message);
      }
    };
    reader.readAsText(file);
  }
  event.target.value = '';
}

function importExcel(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true, cellFormula: true });

  sheets = workbook.SheetNames.map(name => {
    const ws = workbook.Sheets[name];
    const sheetData = createSheetData(name);
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

    // Import column widths
    if (ws['!cols']) {
      ws['!cols'].forEach((col, i) => {
        if (col && col.wpx) sheetData.colWidths[i] = Math.max(40, col.wpx);
        else if (col && col.wch) sheetData.colWidths[i] = Math.max(40, col.wch * 8);
      });
    }

    for (let r = range.s.r; r <= Math.min(range.e.r, ROWS - 1); r++) {
      for (let c = range.s.c; c <= Math.min(range.e.c, COLS - 1); c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const wsCell = ws[addr];
        if (!wsCell) continue;

        const key = cellKey(r, c);
        const cell = {};

        // Value
        if (wsCell.f) {
          cell.formula = '=' + wsCell.f;
          cell.value = wsCell.v ?? '';
        } else if (wsCell.t === 'n') {
          cell.value = wsCell.v;
          cell.detectedType = 'number';
        } else if (wsCell.t === 'd') {
          cell.value = wsCell.w || String(wsCell.v);
          cell.detectedType = 'date';
        } else {
          cell.value = wsCell.v ?? '';
          cell.detectedType = 'text';
        }

        // Number format detection
        if (wsCell.z) {
          if (wsCell.z.includes('$') || wsCell.z.includes('€') || wsCell.z.includes('£')) {
            cell.detectedType = 'currency';
          } else if (wsCell.z.includes('%')) {
            cell.detectedType = 'percent';
            if (typeof cell.value === 'number' && cell.value <= 1) {
              // Already in decimal form, keep as-is for percent display
            }
          } else if (wsCell.z.includes('d') || wsCell.z.includes('m') || wsCell.z.includes('y')) {
            cell.detectedType = 'date';
            if (wsCell.w) cell.value = wsCell.w;
          }
        }

        // Style extraction
        if (wsCell.s) {
          if (wsCell.s.font) {
            if (wsCell.s.font.bold) cell.bold = true;
            if (wsCell.s.font.italic) cell.italic = true;
            if (wsCell.s.font.underline) cell.underline = true;
            if (wsCell.s.font.color && wsCell.s.font.color.rgb) {
              cell.textColor = '#' + wsCell.s.font.color.rgb.slice(-6);
            }
          }
          if (wsCell.s.fill && wsCell.s.fill.fgColor && wsCell.s.fill.fgColor.rgb) {
            const fill = '#' + wsCell.s.fill.fgColor.rgb.slice(-6);
            if (fill !== '#000000') cell.fillColor = fill;
          }
          if (wsCell.s.alignment) {
            if (wsCell.s.alignment.horizontal) cell.align = wsCell.s.alignment.horizontal;
          }
        }

        sheetData.cells[key] = cell;
      }
    }

    return sheetData;
  });

  activeSheet = 0;
  renderSheetTabs();
  renderSheet();
  selectCell(0, 0);
  autoOrganize(0);
  triggerAutoSave();
}

function exportExcel() {
  const wb = XLSX.utils.book_new();
  sheets.forEach(sheet => {
    let maxR = 0, maxC = 0;
    for (const key of Object.keys(sheet.cells)) {
      const { row, col } = parseKey(key);
      maxR = Math.max(maxR, row);
      maxC = Math.max(maxC, col);
    }

    const data = [];
    for (let r = 0; r <= maxR; r++) {
      const row = [];
      for (let c = 0; c <= maxC; c++) {
        const cell = sheet.cells[cellKey(r, c)];
        row.push(cell ? (cell.value ?? '') : '');
      }
      data.push(row);
    }

    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  });

  const fileName = (document.getElementById('file-name').value || 'spreadsheet') + '.xlsx';
  XLSX.writeFile(wb, fileName);
}

function exportCSV() {
  const sheet = sheets[activeSheet];
  let maxR = 0, maxC = 0;
  for (const key of Object.keys(sheet.cells)) {
    const { row, col } = parseKey(key);
    maxR = Math.max(maxR, row);
    maxC = Math.max(maxC, col);
  }

  let csv = '';
  for (let r = 0; r <= maxR; r++) {
    const row = [];
    for (let c = 0; c <= maxC; c++) {
      const cell = sheet.cells[cellKey(r, c)];
      let val = cell ? String(cell.value ?? '') : '';
      if (val.includes(',') || val.includes('"') || val.includes('\n')) {
        val = '"' + val.replace(/"/g, '""') + '"';
      }
      row.push(val);
    }
    csv += row.join(',') + '\n';
  }

  const blob = new Blob([csv], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (document.getElementById('file-name').value || 'spreadsheet') + '.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

function importCSV(text) {
  sheets = [createSheetData('Sheet 1')];
  activeSheet = 0;
  const lines = text.split('\n').filter(l => l.trim());

  lines.forEach((line, r) => {
    // Simple CSV parsing
    const cols = parseCSVLine(line);
    cols.forEach((val, c) => {
      if (c < COLS && val.trim()) {
        const parsed = autoDetect(val.trim());
        sheets[0].cells[cellKey(r, c)] = { value: parsed.value, detectedType: parsed.type };
      }
    });
  });

  renderSheetTabs();
  renderSheet();
  selectCell(0, 0);
  autoOrganize(0);
}

function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"' && line[i + 1] === '"') { current += '"'; i++; }
      else if (ch === '"') { inQuotes = false; }
      else { current += ch; }
    } else {
      if (ch === '"') { inQuotes = true; }
      else if (ch === ',') { result.push(current); current = ''; }
      else { current += ch; }
    }
  }
  result.push(current);
  return result;
}

// Autosave
let autoSaveTimer;
function triggerAutoSave() {
  clearTimeout(autoSaveTimer);
  autoSaveTimer = setTimeout(autoSave, 1000);
}

function autoSave() {
  localStorage.setItem('quantix-autosave', JSON.stringify({
    sheets,
    activeSheet,
    fileName: document.getElementById('file-name').value
  }));
}

// ============================================
// Charts
// ============================================

function insertChart() {
  // Pre-fill range if selection exists
  if (selectionRange) {
    const r = selectionRange;
    const range = COL_LETTERS[Math.min(r.startCol, r.endCol)] + (Math.min(r.startRow, r.endRow) + 1) + ':' +
                  COL_LETTERS[Math.max(r.startCol, r.endCol)] + (Math.max(r.startRow, r.endRow) + 1);
    document.getElementById('chart-range').value = range;
  }
  document.getElementById('chart-modal').style.display = 'flex';
}

function createChart() {
  const type = document.getElementById('chart-type').value;
  const rangeStr = document.getElementById('chart-range').value.toUpperCase();
  const title = document.getElementById('chart-title').value || 'Chart';
  const canvas = document.getElementById('chart-canvas');
  const ctx = canvas.getContext('2d');

  // Parse range
  const match = rangeStr.match(/^([A-Z])(\d+):([A-Z])(\d+)$/);
  if (!match) { alert('Invalid range. Use format like A1:B5'); return; }

  const c1 = match[1].charCodeAt(0) - 65;
  const r1 = +match[2] - 1;
  const c2 = match[3].charCodeAt(0) - 65;
  const r2 = +match[4] - 1;

  // Get data
  const labels = [];
  const datasets = [];
  const numCols = c2 - c1 + 1;

  if (numCols >= 2) {
    // First column = labels, rest = data
    for (let r = r1; r <= r2; r++) {
      const cell = sheets[activeSheet].cells[cellKey(r, c1)];
      labels.push(cell ? String(cell.value) : '');
    }
    for (let c = c1 + 1; c <= c2; c++) {
      const data = [];
      for (let r = r1; r <= r2; r++) {
        const cell = sheets[activeSheet].cells[cellKey(r, c)];
        data.push(cell ? toNum(cell.value) : 0);
      }
      datasets.push(data);
    }
  } else {
    // Single column = data, row numbers = labels
    for (let r = r1; r <= r2; r++) {
      labels.push('Row ' + (r + 1));
      const cell = sheets[activeSheet].cells[cellKey(r, c1)];
      if (!datasets[0]) datasets[0] = [];
      datasets[0].push(cell ? toNum(cell.value) : 0);
    }
  }

  drawChart(ctx, canvas, type, title, labels, datasets);
}

function drawChart(ctx, canvas, type, title, labels, datasets) {
  const W = canvas.width;
  const H = canvas.height;
  const pad = { top: 50, right: 30, bottom: 60, left: 60 };

  ctx.clearRect(0, 0, W, H);
  ctx.fillStyle = '#1a1a2e';
  ctx.fillRect(0, 0, W, H);

  // Title
  ctx.fillStyle = '#e0e0ff';
  ctx.font = 'bold 16px Segoe UI';
  ctx.textAlign = 'center';
  ctx.fillText(title, W / 2, 30);

  const colors = ['#6c5ce7', '#00cec9', '#fdcb6e', '#ff6b6b', '#a29bfe', '#55efc4', '#fab1a0'];

  if (type === 'pie') {
    drawPieChart(ctx, W, H, labels, datasets[0] || [], colors);
    return;
  }

  const chartW = W - pad.left - pad.right;
  const chartH = H - pad.top - pad.bottom;
  const allVals = datasets.flat();
  const maxVal = Math.max(...allVals, 0) * 1.1 || 1;
  const minVal = Math.min(0, ...allVals);

  // Grid
  ctx.strokeStyle = '#2a2a4a';
  ctx.lineWidth = 0.5;
  const gridLines = 5;
  for (let i = 0; i <= gridLines; i++) {
    const y = pad.top + (chartH / gridLines) * i;
    ctx.beginPath();
    ctx.moveTo(pad.left, y);
    ctx.lineTo(pad.left + chartW, y);
    ctx.stroke();

    const val = maxVal - (maxVal - minVal) * (i / gridLines);
    ctx.fillStyle = '#8888bb';
    ctx.font = '10px Segoe UI';
    ctx.textAlign = 'right';
    ctx.fillText(val.toFixed(1), pad.left - 8, y + 4);
  }

  if (type === 'bar') {
    const totalBars = labels.length * datasets.length;
    const groupW = chartW / labels.length;
    const barW = Math.min(groupW * 0.7 / datasets.length, 40);

    datasets.forEach((data, di) => {
      data.forEach((val, i) => {
        const x = pad.left + groupW * i + (groupW - barW * datasets.length) / 2 + barW * di;
        const barH = (val / maxVal) * chartH;
        const y = pad.top + chartH - barH;

        ctx.fillStyle = colors[di % colors.length];
        ctx.beginPath();
        roundRect(ctx, x, y, barW - 2, barH, 3);
        ctx.fill();
      });
    });

    // Labels
    ctx.fillStyle = '#8888bb';
    ctx.font = '10px Segoe UI';
    ctx.textAlign = 'center';
    labels.forEach((label, i) => {
      const x = pad.left + groupW * i + groupW / 2;
      ctx.fillText(label.length > 10 ? label.substring(0, 10) + '…' : label, x, H - pad.bottom + 20);
    });
  }

  if (type === 'line') {
    datasets.forEach((data, di) => {
      ctx.strokeStyle = colors[di % colors.length];
      ctx.lineWidth = 2;
      ctx.beginPath();
      data.forEach((val, i) => {
        const x = pad.left + (chartW / (data.length - 1 || 1)) * i;
        const y = pad.top + chartH - (val / maxVal) * chartH;
        if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
      });
      ctx.stroke();

      // Dots
      ctx.fillStyle = colors[di % colors.length];
      data.forEach((val, i) => {
        const x = pad.left + (chartW / (data.length - 1 || 1)) * i;
        const y = pad.top + chartH - (val / maxVal) * chartH;
        ctx.beginPath();
        ctx.arc(x, y, 4, 0, Math.PI * 2);
        ctx.fill();
      });
    });

    // Labels
    ctx.fillStyle = '#8888bb';
    ctx.font = '10px Segoe UI';
    ctx.textAlign = 'center';
    labels.forEach((label, i) => {
      const x = pad.left + (chartW / (labels.length - 1 || 1)) * i;
      ctx.fillText(label.length > 10 ? label.substring(0, 10) + '…' : label, x, H - pad.bottom + 20);
    });
  }
}

function drawPieChart(ctx, W, H, labels, data, colors) {
  const cx = W / 2;
  const cy = H / 2 + 10;
  const radius = Math.min(W, H) / 2 - 70;
  const total = data.reduce((a, b) => a + b, 0) || 1;

  let startAngle = -Math.PI / 2;
  data.forEach((val, i) => {
    const slice = (val / total) * Math.PI * 2;
    ctx.fillStyle = colors[i % colors.length];
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, radius, startAngle, startAngle + slice);
    ctx.closePath();
    ctx.fill();

    // Label
    const mid = startAngle + slice / 2;
    const lx = cx + Math.cos(mid) * (radius + 20);
    const ly = cy + Math.sin(mid) * (radius + 20);
    ctx.fillStyle = '#e0e0ff';
    ctx.font = '11px Segoe UI';
    ctx.textAlign = Math.cos(mid) > 0 ? 'left' : 'right';
    ctx.fillText(`${labels[i]} (${((val / total) * 100).toFixed(0)}%)`, lx, ly);

    startAngle += slice;
  });
}

function roundRect(ctx, x, y, w, h, r) {
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h);
  ctx.lineTo(x, y + h);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
}

// ============================================
// Status Bar
// ============================================

function updateStatusBar() {
  const vals = [];
  forEachSelected((r, c) => {
    const cell = sheets[activeSheet].cells[cellKey(r, c)];
    if (cell && typeof cell.value === 'number') vals.push(cell.value);
  });

  if (vals.length > 1) {
    const sum = vals.reduce((a, b) => a + b, 0);
    const avg = sum / vals.length;
    document.getElementById('status-sum').textContent =
      `Sum: ${sum.toLocaleString()}  |  Avg: ${avg.toFixed(2)}  |  Count: ${vals.length}`;
  } else {
    document.getElementById('status-sum').textContent = '';
  }
}

// ============================================
// Templates
// ============================================

const TEMPLATES = [
  {
    name: '💰 Budget Tracker',
    description: 'Track monthly income and expenses',
    data: {
      'A0': { value: 'Category', bold: true, fillColor: '#2a2a5a' },
      'B0': { value: 'Budget', bold: true, fillColor: '#2a2a5a' },
      'C0': { value: 'Actual', bold: true, fillColor: '#2a2a5a' },
      'D0': { value: 'Difference', bold: true, fillColor: '#2a2a5a' },
      'A1': { value: 'Housing' }, 'B1': { value: 1500, detectedType: 'currency' }, 'C1': { value: 1500, detectedType: 'currency' },
      'A2': { value: 'Food' }, 'B2': { value: 500, detectedType: 'currency' }, 'C2': { value: 450, detectedType: 'currency' },
      'A3': { value: 'Transport' }, 'B3': { value: 200, detectedType: 'currency' }, 'C3': { value: 180, detectedType: 'currency' },
      'A4': { value: 'Entertainment' }, 'B4': { value: 150, detectedType: 'currency' }, 'C4': { value: 200, detectedType: 'currency' },
      'A5': { value: 'Utilities' }, 'B5': { value: 100, detectedType: 'currency' }, 'C5': { value: 95, detectedType: 'currency' },
      'A6': { value: 'Total', bold: true }, 'B6': { formula: '=SUM(B2:B6)', value: 0 }, 'C6': { formula: '=SUM(C2:C6)', value: 0 },
      'D1': { formula: '=B2-C2', value: 0 }, 'D2': { formula: '=B3-C3', value: 0 }, 'D3': { formula: '=B4-C4', value: 0 },
      'D4': { formula: '=B5-C5', value: 0 }, 'D5': { formula: '=B6-C6', value: 0 }, 'D6': { formula: '=B7-C7', value: 0 },
    }
  },
  {
    name: '📊 Sales Dashboard',
    description: 'Track monthly sales by product',
    data: {
      'A0': { value: 'Month', bold: true, fillColor: '#2a2a5a' },
      'B0': { value: 'Product A', bold: true, fillColor: '#2a2a5a' },
      'C0': { value: 'Product B', bold: true, fillColor: '#2a2a5a' },
      'D0': { value: 'Total', bold: true, fillColor: '#2a2a5a' },
      'A1': { value: 'January' }, 'B1': { value: 5000 }, 'C1': { value: 3200 }, 'D1': { formula: '=B2+C2', value: 0 },
      'A2': { value: 'February' }, 'B2': { value: 6200 }, 'C2': { value: 4100 }, 'D2': { formula: '=B3+C3', value: 0 },
      'A3': { value: 'March' }, 'B3': { value: 7800 }, 'C3': { value: 3800 }, 'D3': { formula: '=B4+C4', value: 0 },
      'A4': { value: 'April' }, 'B4': { value: 5500 }, 'C4': { value: 4500 }, 'D4': { formula: '=B5+C5', value: 0 },
    }
  },
  {
    name: '✅ Task Tracker',
    description: 'Simple project task management',
    data: {
      'A0': { value: 'Task', bold: true, fillColor: '#2a2a5a' },
      'B0': { value: 'Priority', bold: true, fillColor: '#2a2a5a' },
      'C0': { value: 'Status', bold: true, fillColor: '#2a2a5a' },
      'D0': { value: 'Due Date', bold: true, fillColor: '#2a2a5a' },
      'A1': { value: 'Design mockups' }, 'B1': { value: 'High' }, 'C1': { value: 'In Progress' }, 'D1': { value: '2026-04-10', detectedType: 'date' },
      'A2': { value: 'Setup database' }, 'B2': { value: 'High' }, 'C2': { value: 'Done' }, 'D2': { value: '2026-04-05', detectedType: 'date' },
      'A3': { value: 'Write API docs' }, 'B3': { value: 'Medium' }, 'C3': { value: 'Not Started' }, 'D3': { value: '2026-04-15', detectedType: 'date' },
      'A4': { value: 'User testing' }, 'B4': { value: 'Low' }, 'C4': { value: 'Not Started' }, 'D4': { value: '2026-04-20', detectedType: 'date' },
    }
  },
  {
    name: '📈 Grade Calculator',
    description: 'Calculate weighted grades',
    data: {
      'A0': { value: 'Assignment', bold: true, fillColor: '#2a2a5a' },
      'B0': { value: 'Score', bold: true, fillColor: '#2a2a5a' },
      'C0': { value: 'Weight', bold: true, fillColor: '#2a2a5a' },
      'D0': { value: 'Weighted', bold: true, fillColor: '#2a2a5a' },
      'A1': { value: 'Homework' }, 'B1': { value: 92 }, 'C1': { value: 0.2, detectedType: 'percent' },
      'A2': { value: 'Midterm' }, 'B2': { value: 85 }, 'C2': { value: 0.3, detectedType: 'percent' },
      'A3': { value: 'Project' }, 'B3': { value: 95 }, 'C3': { value: 0.2, detectedType: 'percent' },
      'A4': { value: 'Final Exam' }, 'B4': { value: 88 }, 'C4': { value: 0.3, detectedType: 'percent' },
      'D1': { formula: '=B2*C2', value: 0 }, 'D2': { formula: '=B3*C3', value: 0 },
      'D3': { formula: '=B4*C4', value: 0 }, 'D4': { formula: '=B5*C5', value: 0 },
      'A6': { value: 'Final Grade', bold: true }, 'D6': { formula: '=SUM(D2:D5)', value: 0, bold: true },
    }
  },
];

function showTemplates() {
  document.getElementById('templates-modal').style.display = 'flex';
}

function applyTemplate(index) {
  const template = TEMPLATES[index];
  sheets = [createSheetData(template.name)];
  activeSheet = 0;

  for (const [ref, cell] of Object.entries(template.data)) {
    const col = ref.charCodeAt(0) - 65;
    const row = parseInt(ref.substring(1));
    const key = cellKey(row, col);
    sheets[0].cells[key] = { ...cell };
  }

  // Recalc formulas
  for (const [key, cell] of Object.entries(sheets[0].cells)) {
    if (cell.formula) {
      cell.value = evaluateFormula(cell.formula, 0);
    }
  }

  document.getElementById('file-name').value = template.name;
  document.getElementById('templates-modal').style.display = 'none';
  renderSheetTabs();
  renderSheet();
  selectCell(0, 0);
  triggerAutoSave();
}

// ============================================
// Utility Functions
// ============================================

function cellKey(row, col) { return `${col}_${row}`; }
function parseKey(key) { const [col, row] = key.split('_').map(Number); return { row, col }; }
function getCellTd(row, col) { return document.querySelector(`td[data-row="${row}"][data-col="${col}"]`); }

function forEachSelected(fn) {
  if (selectionRange) {
    const r1 = Math.min(selectionRange.startRow, selectionRange.endRow);
    const r2 = Math.max(selectionRange.startRow, selectionRange.endRow);
    const c1 = Math.min(selectionRange.startCol, selectionRange.endCol);
    const c2 = Math.max(selectionRange.startCol, selectionRange.endCol);
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) fn(r, c);
    }
  } else {
    fn(selectedCell.row, selectedCell.col);
  }
}

function escapeHTML(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(str) {
  return String(str).replace(/"/g, '&quot;');
}

function toNum(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function flat(arr) {
  const result = [];
  for (const item of arr) {
    if (Array.isArray(item)) result.push(...flat(item));
    else result.push(item);
  }
  return result;
}

// ============================================
// ============================================
// Conditional Formatting
// ============================================

let condRules = []; // { range: {r1,c1,r2,c2}, rule, value1, value2, color }
let selectedCondColor = '#ff6b6b';

function showConditionalFormat(row, col) {
  document.getElementById('cond-format-modal').style.display = 'flex';
}

function selectCondColor(color, btn) {
  selectedCondColor = color;
  document.querySelectorAll('.cond-color-btn').forEach(b => b.classList.remove('selected'));
  btn.classList.add('selected');
}

document.getElementById('cond-rule-type').addEventListener('change', function() {
  document.getElementById('cond-value2-label').style.display = this.value === 'between' ? '' : 'none';
});

function applyConditionalFormat() {
  const ruleType = document.getElementById('cond-rule-type').value;
  const value1 = document.getElementById('cond-value1').value;
  const value2 = document.getElementById('cond-value2').value;

  let r1, c1, r2, c2;
  if (selectionRange) {
    r1 = Math.min(selectionRange.startRow, selectionRange.endRow);
    r2 = Math.max(selectionRange.startRow, selectionRange.endRow);
    c1 = Math.min(selectionRange.startCol, selectionRange.endCol);
    c2 = Math.max(selectionRange.startCol, selectionRange.endCol);
  } else {
    r1 = r2 = selectedCell.row;
    c1 = c2 = selectedCell.col;
  }

  condRules.push({ range: { r1, c1, r2, c2 }, rule: ruleType, value1, value2, color: selectedCondColor, sheet: activeSheet });
  applyAllCondRules();
  document.getElementById('cond-format-modal').style.display = 'none';
  document.getElementById('status-info').textContent = 'Conditional format applied';
}

function clearConditionalFormats() {
  condRules = condRules.filter(r => r.sheet !== activeSheet);
  // Clear conditional fill colors
  const sheet = sheets[activeSheet];
  for (const [key, cell] of Object.entries(sheet.cells)) {
    if (cell._condColor) {
      delete cell._condColor;
      const { row, col } = parseKey(key);
      refreshCell(row, col);
    }
  }
  document.getElementById('cond-format-modal').style.display = 'none';
  document.getElementById('status-info').textContent = 'Conditional formats cleared';
}

function applyAllCondRules() {
  const sheet = sheets[activeSheet];
  // Clear previous
  for (const [key, cell] of Object.entries(sheet.cells)) {
    if (cell._condColor) { delete cell._condColor; }
  }

  for (const rule of condRules) {
    if (rule.sheet !== activeSheet) continue;
    const { r1, c1, r2, c2 } = rule.range;
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const key = cellKey(r, c);
        const cell = sheet.cells[key];
        if (!cell) continue;
        if (testCondRule(cell.value, rule)) {
          cell._condColor = rule.color;
        }
      }
    }
  }
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
}

function testCondRule(val, rule) {
  const num = toNum(val);
  const v1 = toNum(rule.value1);
  const v2 = toNum(rule.value2);
  switch (rule.rule) {
    case 'greater': return num > v1;
    case 'less': return num < v1;
    case 'equal': return String(val) === rule.value1 || num === v1;
    case 'not_equal': return String(val) !== rule.value1 && num !== v1;
    case 'between': return num >= v1 && num <= v2;
    case 'contains': return String(val).toLowerCase().includes(rule.value1.toLowerCase());
    case 'empty': return val === '' || val === null || val === undefined;
    case 'not_empty': return val !== '' && val !== null && val !== undefined;
    default: return false;
  }
}

// ============================================
// Data Validation / Dropdowns
// ============================================

function showDataValidation(row, col) {
  document.getElementById('validation-modal').style.display = 'flex';
}

function updateValidationUI() {
  const type = document.getElementById('validation-type').value;
  document.getElementById('validation-list-opts').style.display = type === 'list' ? '' : 'none';
  document.getElementById('validation-range-opts').style.display = type === 'number' || type === 'text_length' ? '' : 'none';
}

function applyValidation() {
  const type = document.getElementById('validation-type').value;
  const validation = { type };

  if (type === 'list') {
    validation.options = document.getElementById('validation-list').value.split(',').map(s => s.trim()).filter(Boolean);
  } else {
    validation.min = +document.getElementById('validation-min').value || undefined;
    validation.max = +document.getElementById('validation-max').value || undefined;
  }

  forEachSelected((r, c) => {
    const key = cellKey(r, c);
    if (!sheets[activeSheet].cells[key]) sheets[activeSheet].cells[key] = {};
    sheets[activeSheet].cells[key].validation = validation;
    refreshCell(r, c);
  });

  document.getElementById('validation-modal').style.display = 'none';
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
  document.getElementById('status-info').textContent = 'Validation applied';
}

function clearValidation() {
  forEachSelected((r, c) => {
    const key = cellKey(r, c);
    const cell = sheets[activeSheet].cells[key];
    if (cell) delete cell.validation;
  });
  document.getElementById('validation-modal').style.display = 'none';
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
}

function showCellDropdown(td, row, col, options) {
  closeCellDropdown();
  const rect = td.getBoundingClientRect();
  const dropdown = document.createElement('div');
  dropdown.className = 'cell-dropdown';
  dropdown.id = 'active-cell-dropdown';
  dropdown.style.left = rect.left + 'px';
  dropdown.style.top = rect.bottom + 'px';

  options.forEach(opt => {
    const btn = document.createElement('button');
    btn.textContent = opt;
    btn.onclick = () => {
      setCellValue(row, col, opt);
      closeCellDropdown();
    };
    dropdown.appendChild(btn);
  });

  document.body.appendChild(dropdown);
  setTimeout(() => document.addEventListener('click', closeCellDropdown, { once: true }), 10);
}

function closeCellDropdown() {
  const dd = document.getElementById('active-cell-dropdown');
  if (dd) dd.remove();
}

function validateCell(row, col, value) {
  const cell = sheets[activeSheet].cells[cellKey(row, col)];
  if (!cell || !cell.validation) return true;
  const v = cell.validation;
  if (v.type === 'list') return v.options.includes(String(value));
  if (v.type === 'number') {
    const n = toNum(value);
    if (v.min !== undefined && n < v.min) return false;
    if (v.max !== undefined && n > v.max) return false;
    return true;
  }
  if (v.type === 'text_length') {
    const len = String(value).length;
    if (v.min !== undefined && len < v.min) return false;
    if (v.max !== undefined && len > v.max) return false;
    return true;
  }
  return true;
}

// ============================================
// Merge Cells
// ============================================

function mergeCells() {
  if (!selectionRange) return;
  const r1 = Math.min(selectionRange.startRow, selectionRange.endRow);
  const r2 = Math.max(selectionRange.startRow, selectionRange.endRow);
  const c1 = Math.min(selectionRange.startCol, selectionRange.endCol);
  const c2 = Math.max(selectionRange.startCol, selectionRange.endCol);

  if (r1 === r2 && c1 === c2) return; // Single cell

  const sheet = sheets[activeSheet];
  // Use top-left cell value
  const mainKey = cellKey(r1, c1);
  if (!sheet.cells[mainKey]) sheet.cells[mainKey] = { value: '' };
  sheet.cells[mainKey].merge = { r1, c1, r2, c2 };

  // Mark other cells as merged-hidden
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      if (r === r1 && c === c1) continue;
      const key = cellKey(r, c);
      sheet.cells[key] = { _mergedInto: mainKey, value: '' };
    }
  }

  renderSheet();
  selectCell(r1, c1);
  triggerAutoSave();
  document.getElementById('status-info').textContent = 'Cells merged';
}

function unmergeCells() {
  const sheet = sheets[activeSheet];
  const key = cellKey(selectedCell.row, selectedCell.col);
  const cell = sheet.cells[key];
  if (!cell || !cell.merge) return;

  const { r1, c1, r2, c2 } = cell.merge;
  delete cell.merge;

  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      if (r === r1 && c === c1) continue;
      const k = cellKey(r, c);
      if (sheet.cells[k] && sheet.cells[k]._mergedInto) {
        delete sheet.cells[k];
      }
    }
  }

  renderSheet();
  selectCell(r1, c1);
  triggerAutoSave();
  document.getElementById('status-info').textContent = 'Cells unmerged';
}

// ============================================
// Freeze Panes
// ============================================

let freezeRow = -1;
let freezeCol = -1;

function freezeAt(row, col) {
  freezeRow = row;
  freezeCol = col;
  applyFreeze();
  document.getElementById('status-info').textContent = `Frozen at row ${row + 1}, column ${COL_LETTERS[col]}`;
}

function unfreeze() {
  freezeRow = -1;
  freezeCol = -1;
  // Remove all freeze styles
  document.querySelectorAll('.freeze-row').forEach(el => el.classList.remove('freeze-row'));
  document.querySelectorAll('.freeze-col').forEach(el => el.classList.remove('freeze-col'));
  document.querySelectorAll('.freeze-border-bottom').forEach(el => el.classList.remove('freeze-border-bottom'));
  document.querySelectorAll('.freeze-border-right').forEach(el => el.classList.remove('freeze-border-right'));
  document.querySelectorAll('td[style*="position"], th[style*="position"]').forEach(el => {
    if (!el.matches('td:first-child, thead th')) {
      el.style.position = '';
      el.style.top = '';
      el.style.left = '';
      el.style.zIndex = '';
    }
  });
  document.getElementById('status-info').textContent = 'Panes unfrozen';
}

function applyFreeze() {
  const tbody = document.getElementById('sheet-body');
  const thead = document.getElementById('sheet-head');

  // Freeze rows - make them sticky
  if (freezeRow >= 0) {
    const rows = tbody.querySelectorAll('tr');
    let topOffset = thead.offsetHeight;
    for (let r = 0; r <= freezeRow && r < rows.length; r++) {
      const tds = rows[r].querySelectorAll('td');
      tds.forEach((td, c) => {
        td.style.position = 'sticky';
        td.style.top = topOffset + 'px';
        td.style.zIndex = c === 0 ? '4' : '2';
        td.style.background = 'var(--bg-cell)';
      });
      // Add border to last frozen row
      if (r === freezeRow) {
        tds.forEach(td => td.classList.add('freeze-border-bottom'));
      }
      topOffset += rows[r].offsetHeight;
    }
  }

  // Freeze columns
  if (freezeCol >= 0) {
    const allRows = [thead.querySelector('tr'), ...tbody.querySelectorAll('tr')];
    allRows.forEach(row => {
      if (!row) return;
      const cells = row.querySelectorAll('td, th');
      let leftOffset = cells[0] ? cells[0].offsetWidth : 40; // row number col
      for (let c = 1; c <= freezeCol + 1 && c < cells.length; c++) {
        const cell = cells[c];
        cell.style.position = 'sticky';
        cell.style.left = leftOffset + 'px';
        cell.style.zIndex = cell.style.top ? '5' : '3';
        if (!cell.matches('thead th')) cell.style.background = 'var(--bg-cell)';
        if (c === freezeCol + 1) cell.classList.add('freeze-border-right');
        leftOffset += cell.offsetWidth;
      }
    });
  }
}

// ============================================
// Find & Replace
// ============================================

let findMatches = [];
let findIndex = -1;

function openFindBar(showReplace) {
  const bar = document.getElementById('find-replace-bar');
  bar.style.display = 'flex';
  document.getElementById('find-input').focus();
  if (showReplace) document.getElementById('replace-row').style.display = 'flex';
}

function closeFindBar() {
  document.getElementById('find-replace-bar').style.display = 'none';
  document.getElementById('replace-row').style.display = 'none';
  document.getElementById('find-input').value = '';
  document.getElementById('replace-input').value = '';
  clearFindHighlights();
  findMatches = [];
  findIndex = -1;
}

function toggleReplace() {
  const row = document.getElementById('replace-row');
  row.style.display = row.style.display === 'none' ? 'flex' : 'none';
}

function doFind() {
  clearFindHighlights();
  const query = document.getElementById('find-input').value.toLowerCase();
  if (!query) { findMatches = []; updateFindCount(); return; }

  findMatches = [];
  const sheet = sheets[activeSheet];
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const val = String(cell.value ?? '').toLowerCase();
    const formula = String(cell.formula ?? '').toLowerCase();
    if (val.includes(query) || formula.includes(query)) {
      const { row, col } = parseKey(key);
      findMatches.push({ row, col });
      const td = getCellTd(row, col);
      if (td) td.classList.add('find-highlight');
    }
  }
  updateFindCount();
}

function findNext() {
  doFind();
  if (findMatches.length === 0) return;
  findIndex = (findIndex + 1) % findMatches.length;
  const m = findMatches[findIndex];
  selectCell(m.row, m.col);
  updateFindCount();
}

function findPrev() {
  doFind();
  if (findMatches.length === 0) return;
  findIndex = (findIndex - 1 + findMatches.length) % findMatches.length;
  const m = findMatches[findIndex];
  selectCell(m.row, m.col);
  updateFindCount();
}

function replaceCurrent() {
  if (findIndex < 0 || findIndex >= findMatches.length) { findNext(); return; }
  const m = findMatches[findIndex];
  const key = cellKey(m.row, m.col);
  const cell = sheets[activeSheet].cells[key];
  if (!cell) return;

  const query = document.getElementById('find-input').value;
  const replacement = document.getElementById('replace-input').value;
  const newVal = String(cell.value).split(query).join(replacement);
  setCellValue(m.row, m.col, newVal);
  findNext();
}

function replaceAll() {
  doFind();
  const query = document.getElementById('find-input').value;
  const replacement = document.getElementById('replace-input').value;
  if (!query || findMatches.length === 0) return;

  let count = 0;
  for (const m of findMatches) {
    const key = cellKey(m.row, m.col);
    const cell = sheets[activeSheet].cells[key];
    if (!cell) continue;
    const newVal = String(cell.value).split(query).join(replacement);
    setCellValue(m.row, m.col, newVal, true);
    count++;
  }
  document.getElementById('status-info').textContent = `Replaced ${count} occurrences`;
  doFind();
}

function clearFindHighlights() {
  document.querySelectorAll('td.find-highlight').forEach(td => td.classList.remove('find-highlight'));
}

function updateFindCount() {
  const el = document.getElementById('find-count');
  if (findMatches.length === 0) {
    el.textContent = document.getElementById('find-input').value ? 'No matches' : '';
  } else {
    el.textContent = `${findIndex + 1} of ${findMatches.length}`;
  }
}

// Live search as you type
document.getElementById('find-input').addEventListener('input', () => { findIndex = -1; doFind(); });
document.getElementById('find-input').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); findNext(); }
});

// ============================================
// Smart Auto-Organize
// ============================================

function autoOrganize(sheetIdx) {
  const sheet = sheets[sheetIdx ?? activeSheet];
  const changes = [];

  // Find data bounds
  let maxR = 0, maxC = 0;
  for (const key of Object.keys(sheet.cells)) {
    const { row, col } = parseKey(key);
    maxR = Math.max(maxR, row);
    maxC = Math.max(maxC, col);
  }
  if (maxR === 0 && maxC === 0 && Object.keys(sheet.cells).length === 0) return;

  // Analyze each column
  const colAnalysis = [];
  for (let c = 0; c <= maxC; c++) {
    const types = { number: 0, currency: 0, percent: 0, date: 0, text: 0, empty: 0 };
    const values = [];
    let hasHeader = false;

    for (let r = 0; r <= maxR; r++) {
      const cell = sheet.cells[cellKey(r, c)];
      if (!cell || cell.value === '' || cell.value === undefined) {
        types.empty++;
        continue;
      }

      const val = cell.value;
      const detected = cell.detectedType || (typeof val === 'number' ? 'number' : 'text');
      types[detected] = (types[detected] || 0) + 1;
      values.push({ row: r, value: val, type: detected });

      // First row is likely a header if it's text and the rest are numbers
      if (r === 0 && typeof val === 'string' && isNaN(+val)) hasHeader = true;
    }

    const dominant = Object.entries(types)
      .filter(([k]) => k !== 'empty')
      .sort((a, b) => b[1] - a[1])[0];

    colAnalysis.push({
      col: c,
      types,
      dominant: dominant ? dominant[0] : 'text',
      count: values.length,
      hasHeader,
      values
    });
  }

  // 1. Auto-detect and style headers (row 0)
  const headerRow = detectHeaderRow(sheet, maxC, colAnalysis);
  if (headerRow !== null) {
    for (let c = 0; c <= maxC; c++) {
      const key = cellKey(headerRow, c);
      if (sheet.cells[key]) {
        sheet.cells[key].bold = true;
        sheet.cells[key].fillColor = '#2a2a5a';
        changes.push('Styled header row');
      }
    }
  }

  // 2. Auto-format columns by detected type
  const startRow = headerRow !== null ? headerRow + 1 : 0;
  for (const col of colAnalysis) {
    for (let r = startRow; r <= maxR; r++) {
      const key = cellKey(r, col.col);
      const cell = sheet.cells[key];
      if (!cell || cell.value === '' || cell.value === undefined) continue;

      // Auto-align: numbers right, text left
      if (col.dominant === 'number' || col.dominant === 'currency' || col.dominant === 'percent') {
        cell.align = 'right';
        // Try to parse text that looks like numbers
        if (typeof cell.value === 'string' && !isNaN(+cell.value.replace(/[,$%]/g, ''))) {
          const parsed = autoDetect(cell.value);
          cell.value = parsed.value;
          cell.detectedType = parsed.type;
        }
      }

      if (col.dominant === 'currency' && !cell.detectedType) {
        cell.detectedType = 'currency';
      }
      if (col.dominant === 'percent' && !cell.detectedType) {
        cell.detectedType = 'percent';
      }
      if (col.dominant === 'date') {
        cell.detectedType = 'date';
      }
    }
    if (col.dominant !== 'text') {
      changes.push(`Column ${COL_LETTERS[col.col]}: formatted as ${col.dominant}`);
    }
  }

  // 3. Auto-fit column widths based on content
  for (let c = 0; c <= maxC; c++) {
    let maxLen = 0;
    for (let r = 0; r <= maxR; r++) {
      const cell = sheet.cells[cellKey(r, c)];
      if (cell && cell.value !== undefined) {
        const display = getDisplayValue(cell);
        maxLen = Math.max(maxLen, display.length);
      }
    }
    if (maxLen > 0) {
      sheet.colWidths[c] = Math.min(250, Math.max(60, maxLen * 8.5 + 20));
    }
  }
  changes.push('Auto-fitted column widths');

  // 4. Add SUM row if numeric columns detected
  const numericCols = colAnalysis.filter(c => (c.dominant === 'number' || c.dominant === 'currency') && c.count >= 3);
  if (numericCols.length > 0 && headerRow !== null) {
    const sumRow = maxR + 2; // Leave a gap
    // Label
    const labelKey = cellKey(sumRow, 0);
    if (!sheet.cells[labelKey]) {
      sheet.cells[labelKey] = { value: 'Total', bold: true, fillColor: '#1f3a2a' };
    }
    for (const col of numericCols) {
      if (col.col === 0) continue;
      const sumKey = cellKey(sumRow, col.col);
      if (!sheet.cells[sumKey]) {
        const colLetter = COL_LETTERS[col.col];
        const formula = `=SUM(${colLetter}${startRow + 1}:${colLetter}${maxR + 1})`;
        sheet.cells[sumKey] = {
          formula,
          value: evaluateFormula(formula, sheetIdx ?? activeSheet),
          bold: true,
          fillColor: '#1f3a2a',
          detectedType: col.dominant
        };
      }
    }
    changes.push('Added totals row');
  }

  // 5. Zebra-stripe rows for readability
  for (let r = startRow; r <= maxR; r++) {
    if ((r - startRow) % 2 === 1) {
      for (let c = 0; c <= maxC; c++) {
        const key = cellKey(r, c);
        if (sheet.cells[key] && !sheet.cells[key].fillColor) {
          sheet.cells[key].fillColor = '#15152e';
        }
      }
    }
  }
  changes.push('Added zebra striping');

  renderSheet();
  selectCell(0, 0);
  triggerAutoSave();

  // Show summary
  const unique = [...new Set(changes)];
  document.getElementById('status-info').textContent = 'Auto-organized: ' + unique.slice(0, 3).join(', ');
  return unique;
}

function detectHeaderRow(sheet, maxC, colAnalysis) {
  // Check if row 0 looks like headers
  let textCount = 0;
  let hasContent = false;
  for (let c = 0; c <= maxC; c++) {
    const cell = sheet.cells[cellKey(0, c)];
    if (!cell || cell.value === '' || cell.value === undefined) continue;
    hasContent = true;
    if (typeof cell.value === 'string' && isNaN(+cell.value)) textCount++;
  }
  // If most of row 0 is text and columns below have different types, it's a header
  if (hasContent && textCount >= Math.ceil((maxC + 1) * 0.5)) {
    const hasNumericBelow = colAnalysis.some(c => c.dominant === 'number' || c.dominant === 'currency');
    if (hasNumericBelow || textCount >= maxC) return 0;
  }
  return null;
}

function showAutoOrganizeToast(changes) {
  if (!changes || changes.length === 0) return;
}

// ============================================
// Init
// ============================================

init();
