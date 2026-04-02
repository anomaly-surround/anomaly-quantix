// ============================================
// Anomaly Quantix - Spreadsheet Engine
// ============================================

const ROWS = 100;
const COLS = 26;
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

function renderSheet() {
  const sheet = sheets[activeSheet];
  const thead = document.getElementById('sheet-head');
  const tbody = document.getElementById('sheet-body');

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

  // Body
  let bodyHTML = '';
  for (let r = 0; r < ROWS; r++) {
    bodyHTML += `<tr><td>${r + 1}</td>`;
    for (let c = 0; c < COLS; c++) {
      const key = cellKey(r, c);
      const cell = sheet.cells[key] || {};
      const display = getDisplayValue(cell);
      const style = getCellStyle(cell);
      const type = detectType(cell);
      bodyHTML += `<td data-row="${r}" data-col="${c}">
        <div class="cell" data-type="${type}" style="${style}">${escapeHTML(display)}</div>
      </td>`;
    }
    bodyHTML += '</tr>';
  }
  tbody.innerHTML = bodyHTML;

  attachCellEvents();
  attachResizeEvents();
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
    startEdit(+td.dataset.row, +td.dataset.col);
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
      SUM: (arr) => flat(arr).reduce((a, b) => a + toNum(b), 0),
      AVERAGE: (arr) => { const f = flat(arr).map(toNum); return f.reduce((a, b) => a + b, 0) / f.length; },
      COUNT: (arr) => flat(arr).filter(v => typeof v === 'number' || !isNaN(+v)).length,
      COUNTA: (arr) => flat(arr).filter(v => v !== '' && v !== null && v !== undefined).length,
      MIN: (arr) => Math.min(...flat(arr).map(toNum).filter(n => !isNaN(n))),
      MAX: (arr) => Math.max(...flat(arr).map(toNum).filter(n => !isNaN(n))),
      IF: (args) => args[0] ? args[1] : args[2],
      ROUND: (args) => { const [n, d = 0] = args; return +toNum(n).toFixed(toNum(d)); },
      ABS: (args) => Math.abs(toNum(args[0])),
      SQRT: (args) => Math.sqrt(toNum(args[0])),
      POWER: (args) => Math.pow(toNum(args[0]), toNum(args[1])),
      CONCATENATE: (args) => flat(args).join(''),
      LEN: (args) => String(args[0]).length,
      UPPER: (args) => String(args[0]).toUpperCase(),
      LOWER: (args) => String(args[0]).toLowerCase(),
      TRIM: (args) => String(args[0]).trim(),
      NOW: () => new Date().toLocaleString(),
      TODAY: () => new Date().toLocaleDateString(),
      PI: () => Math.PI,
    };

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
  if (cell.fillColor && cell.fillColor !== '#1a1a2e') parts.push(`background:${cell.fillColor}`);
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
  if (e.ctrlKey && e.key === 'z') { e.preventDefault(); undo(); return; }
  if (e.ctrlKey && e.key === 'y') { e.preventDefault(); redo(); return; }
  if (e.ctrlKey && e.key === 'b') { e.preventDefault(); toggleBold(); return; }
  if (e.ctrlKey && e.key === 'i') { e.preventDefault(); toggleItalic(); return; }
  if (e.ctrlKey && e.key === 'u') { e.preventDefault(); toggleUnderline(); return; }
  if (e.ctrlKey && e.key === 'c') { copySelection(); return; }
  if (e.ctrlKey && e.key === 'v') { pasteSelection(); return; }
  if (e.ctrlKey && e.key === 'x') { cutSelection(); return; }

  // Skip if focused on formula bar or file name
  if (document.activeElement.id === 'formula-input' || document.activeElement.id === 'file-name') return;

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
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      if (file.name.endsWith('.csv')) {
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
  event.target.value = '';
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
// Init
// ============================================

init();
