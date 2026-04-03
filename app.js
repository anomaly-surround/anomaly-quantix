// ============================================
// Anomaly Quantix - Spreadsheet Engine
// ============================================

const ROWS = 1000;
const COLS = 26;
const VISIBLE_BUFFER = 10; // extra rows to render above/below viewport
const COL_LETTERS = Array.from({ length: COLS }, (_, i) => String.fromCharCode(65 + i));
const PRO_CHECKOUT_URL = 'https://aanomaly.lemonsqueezy.com/checkout/buy/c37c05a0-89a5-47a5-a1e6-bf90310f5090';
const WORKER_URL = 'https://quantix-pro.xpropics.workers.dev';

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
let multiSelection = []; // [{row, col}, ...] for Ctrl+Click

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
      if (data.sheets && Array.isArray(data.sheets)) {
        sheets = data.sheets.map(s => sanitizeSheetData(s));
        activeSheet = Math.max(0, Math.min(toNum(data.activeSheet || 0), sheets.length - 1));
        document.getElementById('file-name').value = String(data.fileName || 'Untitled Spreadsheet');
        renderSheetTabs();
        renderSheet();
        selectCell(0, 0);
      }
    } catch (e) { /* ignore */ }
  }

  // Request notification permission and start deadline checker
  if ('Notification' in window && Notification.permission === 'default') {
    Notification.requestPermission();
  }
  checkDeadlines();
  setInterval(checkDeadlines, 60000); // check every minute
  handleAuthCallback();
  updateProUI();
  checkProStatus();
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

  const hiddenCols = sheet.hiddenCols || [];
  const hiddenRows = sheet.hiddenRows || [];
  const hiddenFilterRows = sheet.hiddenFilterRows || [];
  const allHiddenRows = new Set([...hiddenRows, ...hiddenFilterRows]);

  // Header
  let headHTML = '<tr><th></th>';
  for (let c = 0; c < COLS; c++) {
    if (hiddenCols.includes(c)) {
      headHTML += `<th data-col="${c}" style="display:none"></th>`;
      continue;
    }
    const w = sheet.colWidths[c] || 80;
    const hiddenIndicator = (c > 0 && hiddenCols.includes(c - 1)) ? ' hidden-col-indicator' : '';
    const filterArrow = filtersActive ? `<span class="col-filter-arrow" data-filter-col="${c}">&#9660;</span>` : '';
    headHTML += `<th style="width:${w}px;min-width:${w}px;max-width:${w}px" data-col="${c}" class="${hiddenIndicator}">
      ${COL_LETTERS[c]}${filterArrow}
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
    if (allHiddenRows.has(r)) {
      bodyHTML += `<tr style="display:none"><td>${r + 1}</td>`;
      for (let c = 0; c < COLS; c++) {
        bodyHTML += `<td data-row="${r}" data-col="${c}"></td>`;
      }
      bodyHTML += '</tr>';
      continue;
    }

    const hiddenRowIndicator = (r > 0 && allHiddenRows.has(r - 1)) ? ' class="hidden-row-indicator"' : '';
    bodyHTML += `<tr${hiddenRowIndicator}><td data-rownumber="${r}">${r + 1}</td>`;
    for (let c = 0; c < COLS; c++) {
      if (hiddenCols.includes(c)) {
        bodyHTML += `<td data-row="${r}" data-col="${c}" style="display:none"></td>`;
        continue;
      }

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
      const dropdownClass = (cell.validation && cell.validation.type === 'list' ? ' has-dropdown' : '') +
                            (cell.comment ? ' has-comment' : '');

      // Merge attributes
      let mergeAttr = '';
      if (cell.merge) {
        const rs = cell.merge.r2 - cell.merge.r1 + 1;
        const cs = cell.merge.c2 - cell.merge.c1 + 1;
        mergeAttr = ` rowspan="${rs}" colspan="${cs}"`;
      }

      const colHiddenIndicator = (c > 0 && hiddenCols.includes(c - 1)) ? ' hidden-col-indicator' : '';

      bodyHTML += `<td data-row="${r}" data-col="${c}"${mergeAttr} class="${colHiddenIndicator}">
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
  attachFilterEvents();
  attachCommentHover();
  attachRowDragEvents();
  repositionAllCharts();

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

let cellEventsAttached = false;

function attachCellEvents() {
  const tbody = document.getElementById('sheet-body');

  if (cellEventsAttached) return;
  cellEventsAttached = true;

  tbody.addEventListener('mousedown', (e) => {
    const td = e.target.closest('td[data-row]');
    if (!td) return;

    const row = +td.dataset.row;
    const col = +td.dataset.col;

    // Right-click: don't reset selection, let contextmenu handler deal with it
    if (e.button === 2) return;

    // If picking a cell for quick formula, place it
    if (pendingFormula) {
      e.preventDefault();
      e.stopPropagation();
      placePendingFormula(row, col);
      return;
    }

    // If editing a formula, clicking another cell inserts its reference
    if (editingCell) {
      const editTd = getCellTd(editingCell.row, editingCell.col);
      const input = editTd?.querySelector('input');
      if (input && input.value.startsWith('=')) {
        e.preventDefault();
        e.stopPropagation();
        const ref = COL_LETTERS[col] + (row + 1);
        insertAtCursor(input, ref);
        return;
      }
      commitEdit();
    }

    // Also check if formula bar is focused and has a formula
    const formulaInput = document.getElementById('formula-input');
    if (document.activeElement === formulaInput && formulaInput.value.startsWith('=')) {
      e.preventDefault();
      const ref = COL_LETTERS[col] + (row + 1);
      insertAtCursor(formulaInput, ref);
      return;
    }

    // Block drag: if clicking inside an existing multi-cell selection, start block drag
    if (!e.shiftKey && !e.ctrlKey && !e.metaKey && selectionRange && isInSelectionRange(row, col)) {
      const r1 = Math.min(selectionRange.startRow, selectionRange.endRow);
      const r2 = Math.max(selectionRange.startRow, selectionRange.endRow);
      const c1 = Math.min(selectionRange.startCol, selectionRange.endCol);
      const c2 = Math.max(selectionRange.startCol, selectionRange.endCol);
      // Only if it's a real range (not a single cell)
      if (r1 !== r2 || c1 !== c2) {
        e.preventDefault();
        startBlockDrag(r1, c1, r2, c2, row, col, e);
        return;
      }
    }

    if (e.ctrlKey || e.metaKey) {
      // Ctrl+Click: add/remove cell from multi-selection
      const idx = multiSelection.findIndex(s => s.row === row && s.col === col);
      if (idx >= 0) {
        multiSelection.splice(idx, 1);
      } else {
        multiSelection.push({ row, col });
      }
      selectedCell = { row, col };
      selectionRange = null;
      highlightMultiSelection();
    } else if (e.shiftKey) {
      multiSelection = [];
      selectionRange = {
        startRow: selectedCell.row, startCol: selectedCell.col,
        endRow: row, endCol: col
      };
      highlightRange();
    } else {
      multiSelection = [];
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

function attachFilterEvents() {
  if (!filtersActive) return;
  document.querySelectorAll('.col-filter-arrow').forEach(arrow => {
    arrow.addEventListener('click', (e) => {
      e.stopPropagation();
      const colIdx = +arrow.dataset.filterCol;
      const th = arrow.closest('th');
      showFilterDropdown(colIdx, th);
    });
  });
}

let commentHoverAttached = false;

function attachCommentHover() {
  if (commentHoverAttached) return;
  commentHoverAttached = true;
  const tbody = document.getElementById('sheet-body');
  tbody.addEventListener('mouseover', (e) => {
    const cellDiv = e.target.closest('.cell.has-comment');
    if (!cellDiv) { hideCommentTooltip(); return; }
    const td = cellDiv.closest('td[data-row]');
    if (!td) return;
    const row = +td.dataset.row;
    const col = +td.dataset.col;
    const cell = sheets[activeSheet].cells[cellKey(row, col)];
    if (cell && cell.comment) showCommentTooltip(td, cell.comment);
  });
  tbody.addEventListener('mouseout', (e) => {
    const cellDiv = e.target.closest('.cell.has-comment');
    if (cellDiv) {
      const related = e.relatedTarget?.closest('.cell.has-comment');
      if (related !== cellDiv) hideCommentTooltip();
    }
  });
}

function attachRowDragEvents() {
  const tbody = document.getElementById('sheet-body');
  tbody.querySelectorAll('td[data-rownumber]').forEach(td => {
    td.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      const row = +td.dataset.rownumber;
      // Delay slightly to distinguish click from drag
      const startY = e.clientY;
      const onMove = (e2) => {
        if (Math.abs(e2.clientY - startY) > 5) {
          document.removeEventListener('mousemove', onMove);
          document.removeEventListener('mouseup', onUp);
          startRowDrag(row, e);
        }
      };
      const onUp = () => {
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

  // Highlight formula references
  highlightFormulaRefs(cell);

  // Update toolbar state
  updateToolbarState(cell);
  updateStatusBar();

  // Scroll into view
  if (td) td.scrollIntoView({ block: 'nearest', inline: 'nearest' });

  // Show autofill handle
  showAutofillHandle();
}

const FORMULA_REF_COLORS = [
  { outline: '#ff6b6b', bg: 'rgba(255,107,107,0.15)' },
  { outline: '#74b9ff', bg: 'rgba(116,185,255,0.15)' },
  { outline: '#00cec9', bg: 'rgba(0,206,201,0.15)' },
  { outline: '#fdcb6e', bg: 'rgba(253,203,110,0.15)' },
  { outline: '#a29bfe', bg: 'rgba(162,155,254,0.15)' },
  { outline: '#fd79a8', bg: 'rgba(253,121,168,0.15)' },
  { outline: '#ffa502', bg: 'rgba(255,165,2,0.15)' },
  { outline: '#7bed9f', bg: 'rgba(123,237,159,0.15)' },
];

function highlightFormulaRefs(cell) {
  // Clear previous highlights
  document.querySelectorAll('td.formula-ref').forEach(td => {
    td.classList.remove('formula-ref');
    td.querySelector('.cell')?.removeAttribute('style');
  });

  if (!cell || !cell.formula) return;

  const formula = cell.formula;
  // Match ranges like A1:B10 and single refs like C3
  const rangeRegex = /([A-Z]+)(\d+):([A-Z]+)(\d+)/gi;
  const singleRegex = /([A-Z]+)(\d+)/gi;

  const refs = []; // each entry: { cells: [{row, col}, ...] }

  // First extract ranges
  const usedPositions = new Set();
  let match;
  while ((match = rangeRegex.exec(formula)) !== null) {
    const c1 = match[1].toUpperCase().charCodeAt(0) - 65;
    const r1 = parseInt(match[2]) - 1;
    const c2 = match[3].toUpperCase().charCodeAt(0) - 65;
    const r2 = parseInt(match[4]) - 1;
    const cells = [];
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
        cells.push({ row: r, col: c });
        usedPositions.add(match.index + '-' + (match.index + match[0].length));
      }
    }
    refs.push({ cells });
  }

  // Then extract single refs not part of a range
  while ((match = singleRegex.exec(formula)) !== null) {
    let isPartOfRange = false;
    for (const pos of usedPositions) {
      const [start, end] = pos.split('-').map(Number);
      if (match.index >= start && match.index < end) { isPartOfRange = true; break; }
    }
    if (isPartOfRange) continue;
    const c = match[1].toUpperCase().charCodeAt(0) - 65;
    const r = parseInt(match[2]) - 1;
    if (c >= 0 && c < COLS && r >= 0 && r < ROWS) {
      refs.push({ cells: [{ row: r, col: c }] });
    }
  }

  // Apply colored highlights
  refs.forEach((ref, i) => {
    const color = FORMULA_REF_COLORS[i % FORMULA_REF_COLORS.length];
    ref.cells.forEach(({ row, col }) => {
      const td = getCellTd(row, col);
      if (td) {
        td.classList.add('formula-ref');
        const cellDiv = td.querySelector('.cell');
        if (cellDiv) {
          cellDiv.style.outline = `2px solid ${color.outline}`;
          cellDiv.style.outlineOffset = '-1px';
          cellDiv.style.backgroundColor = color.bg;
        }
      }
    });
  });
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

function highlightMultiSelection() {
  document.querySelectorAll('td.in-range').forEach(td => td.classList.remove('in-range'));
  document.querySelectorAll('td.selected').forEach(td => td.classList.remove('selected'));

  multiSelection.forEach(({ row, col }) => {
    const td = getCellTd(row, col);
    if (td) td.classList.add('in-range');
  });

  // Update cell ref and formula bar for last selected
  if (multiSelection.length > 0) {
    const last = multiSelection[multiSelection.length - 1];
    document.getElementById('cell-ref').textContent = COL_LETTERS[last.col] + (last.row + 1);
    const key = cellKey(last.row, last.col);
    const cell = sheets[activeSheet].cells[key] || {};
    document.getElementById('formula-input').value = cell.formula || cell.value || '';
  }

  updateStatusBar();
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

  input.addEventListener('blur', () => { hideAutocompleteSuggestions(); commitEdit(); });
  input.addEventListener('input', () => showAutocompleteSuggestions(input, row, col));
  input.addEventListener('keydown', (e) => {
    if (handleAutocompleteKey(e, input)) return;
  });
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

// Safe expression tokenizer & parser (no eval/Function)
let _formulaDepth = 0;
const _formulaVisited = new Set();
const MAX_FORMULA_DEPTH = 100;

function evaluateFormula(formula, sheetIdx) {
  if (_formulaDepth >= MAX_FORMULA_DEPTH) return '#CIRCULAR';
  _formulaDepth++;
  try {
    const expr = formula.substring(1);
    const sheet = sheets[sheetIdx];

    // Handle VLOOKUP/HLOOKUP specially
    const upper = expr.toUpperCase();
    if (upper.startsWith('VLOOKUP(')) return evalVlookup(upper, sheet, sheetIdx);
    if (upper.startsWith('HLOOKUP(')) return evalHlookup(upper, sheet, sheetIdx);

    const tokens = tokenize(expr);
    const resolved = resolveTokens(tokens, sheet, sheetIdx);
    const result = parseExpression(resolved, { pos: 0 });
    return typeof result === 'number' ? (Math.round(result * 1e10) / 1e10) : result;
  } catch (e) {
    return '#ERROR';
  } finally {
    _formulaDepth--;
  }
}

function tokenize(expr) {
  const tokens = [];
  let i = 0;
  while (i < expr.length) {
    const ch = expr[i];
    if (/\s/.test(ch)) { i++; continue; }
    // String literal
    if (ch === '"') {
      let s = ''; i++;
      while (i < expr.length && expr[i] !== '"') { s += expr[i]; i++; }
      i++; // skip closing "
      tokens.push({ type: 'string', value: s });
    }
    // Number
    else if (/[0-9.]/.test(ch) || (ch === '-' && (tokens.length === 0 || ['op', 'lparen', 'comma'].includes(tokens[tokens.length - 1]?.type)))) {
      let n = ch; i++;
      while (i < expr.length && /[0-9.eE]/.test(expr[i])) { n += expr[i]; i++; }
      tokens.push({ type: 'number', value: parseFloat(n) });
    }
    // Cell ref or function name or boolean
    else if (/[A-Za-z_]/.test(ch)) {
      let name = '';
      while (i < expr.length && /[A-Za-z0-9_]/.test(expr[i])) { name += expr[i]; i++; }
      const upper = name.toUpperCase();
      if (upper === 'TRUE') tokens.push({ type: 'boolean', value: true });
      else if (upper === 'FALSE') tokens.push({ type: 'boolean', value: false });
      // Range ref: A1:B10
      else if (i < expr.length && expr[i] === ':') {
        i++; // skip :
        let end = '';
        while (i < expr.length && /[A-Za-z0-9]/.test(expr[i])) { end += expr[i]; i++; }
        tokens.push({ type: 'range', value: upper + ':' + end.toUpperCase() });
      }
      // Function call
      else if (i < expr.length && expr[i] === '(') {
        tokens.push({ type: 'func', value: upper });
      }
      // Cell reference (like A1, B23)
      else if (/^[A-Z]\d+$/i.test(name)) {
        tokens.push({ type: 'cellref', value: upper });
      }
      else {
        tokens.push({ type: 'string', value: name });
      }
    }
    else if (ch === '(') { tokens.push({ type: 'lparen' }); i++; }
    else if (ch === ')') { tokens.push({ type: 'rparen' }); i++; }
    else if (ch === ',') { tokens.push({ type: 'comma' }); i++; }
    else if (ch === '+' || ch === '-' || ch === '*' || ch === '/' || ch === '%' || ch === '^') {
      tokens.push({ type: 'op', value: ch }); i++;
    }
    else if (ch === '>' || ch === '<' || ch === '!' || ch === '=') {
      let op = ch; i++;
      if (i < expr.length && (expr[i] === '=' || (ch === '<' && expr[i] === '>'))) { op += expr[i]; i++; }
      tokens.push({ type: 'compare', value: op });
    }
    else if (ch === '&') { tokens.push({ type: 'op', value: '&' }); i++; }
    else { i++; } // skip unknown
  }
  return tokens;
}

function resolveTokens(tokens, sheet, sheetIdx) {
  return tokens.map(t => {
    if (t.type === 'cellref') {
      const col = t.value.charCodeAt(0) - 65;
      const row = parseInt(t.value.substring(1)) - 1;
      const key = sheetIdx + ':' + cellKey(row, col);
      if (_formulaVisited.has(key)) return { type: 'string', value: '#CIRCULAR' };
      const cell = sheet.cells[cellKey(row, col)];
      if (!cell) return { type: 'number', value: 0 };
      _formulaVisited.add(key);
      const v = cell.formula ? evaluateFormula(cell.formula, sheetIdx) : cell.value;
      _formulaVisited.delete(key);
      return typeof v === 'number' ? { type: 'number', value: v } : { type: 'string', value: String(v ?? '') };
    }
    if (t.type === 'range') {
      const m = t.value.match(/^([A-Z])(\d+):([A-Z])(\d+)$/);
      if (!m) return { type: 'array', value: [] };
      const values = getRangeValues(m[1].charCodeAt(0) - 65, +m[2] - 1, m[3].charCodeAt(0) - 65, +m[4] - 1, sheetIdx);
      return { type: 'array', value: values };
    }
    return t;
  });
}

// Recursive descent parser for safe expression evaluation
function parseExpression(tokens, ctx) {
  return parseCompare(tokens, ctx);
}

function parseCompare(tokens, ctx) {
  let left = parseConcat(tokens, ctx);
  while (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'compare') {
    const op = tokens[ctx.pos].value; ctx.pos++;
    const right = parseConcat(tokens, ctx);
    const l = toNum(left), r = toNum(right);
    if (op === '>') left = l > r;
    else if (op === '<') left = l < r;
    else if (op === '>=') left = l >= r;
    else if (op === '<=') left = l <= r;
    else if (op === '=' || op === '==') left = left == right;
    else if (op === '<>' || op === '!=') left = left != right;
  }
  return left;
}

function parseConcat(tokens, ctx) {
  let left = parseAddSub(tokens, ctx);
  while (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'op' && tokens[ctx.pos].value === '&') {
    ctx.pos++;
    const right = parseAddSub(tokens, ctx);
    left = String(left) + String(right);
  }
  return left;
}

function parseAddSub(tokens, ctx) {
  let left = parseMulDiv(tokens, ctx);
  while (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'op' && (tokens[ctx.pos].value === '+' || tokens[ctx.pos].value === '-')) {
    const op = tokens[ctx.pos].value; ctx.pos++;
    const right = parseMulDiv(tokens, ctx);
    left = op === '+' ? toNum(left) + toNum(right) : toNum(left) - toNum(right);
  }
  return left;
}

function parseMulDiv(tokens, ctx) {
  let left = parsePower(tokens, ctx);
  while (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'op' && (tokens[ctx.pos].value === '*' || tokens[ctx.pos].value === '/' || tokens[ctx.pos].value === '%')) {
    const op = tokens[ctx.pos].value; ctx.pos++;
    const right = parsePower(tokens, ctx);
    if (op === '*') left = toNum(left) * toNum(right);
    else if (op === '/') left = toNum(right) !== 0 ? toNum(left) / toNum(right) : '#DIV/0';
    else left = toNum(left) % toNum(right);
  }
  return left;
}

function parsePower(tokens, ctx) {
  let left = parseUnary(tokens, ctx);
  while (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'op' && tokens[ctx.pos].value === '^') {
    ctx.pos++;
    const right = parseUnary(tokens, ctx);
    left = Math.pow(toNum(left), toNum(right));
  }
  return left;
}

function parseUnary(tokens, ctx) {
  if (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'op' && tokens[ctx.pos].value === '-') {
    ctx.pos++;
    return -toNum(parseAtom(tokens, ctx));
  }
  if (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'op' && tokens[ctx.pos].value === '+') {
    ctx.pos++;
  }
  return parseAtom(tokens, ctx);
}

function parseAtom(tokens, ctx) {
  if (ctx.pos >= tokens.length) return 0;
  const t = tokens[ctx.pos];

  if (t.type === 'number') { ctx.pos++; return t.value; }
  if (t.type === 'string') { ctx.pos++; return t.value; }
  if (t.type === 'boolean') { ctx.pos++; return t.value; }
  if (t.type === 'array') { ctx.pos++; return t.value; }

  // Parenthesized expression
  if (t.type === 'lparen') {
    ctx.pos++;
    const val = parseExpression(tokens, ctx);
    if (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'rparen') ctx.pos++;
    return val;
  }

  // Function call
  if (t.type === 'func') {
    const name = t.value;
    ctx.pos++; // func name
    if (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'lparen') ctx.pos++; // (

    // Parse arguments
    const args = [];
    while (ctx.pos < tokens.length && tokens[ctx.pos]?.type !== 'rparen') {
      args.push(parseExpression(tokens, ctx));
      if (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'comma') ctx.pos++;
    }
    if (ctx.pos < tokens.length && tokens[ctx.pos]?.type === 'rparen') ctx.pos++; // )

    return callFunction(name, args);
  }

  ctx.pos++;
  return 0;
}

// All formula functions - safe, no eval
const FORMULA_FUNCS = {
  // Math
  SUM: (a) => flat(a).reduce((s, v) => s + toNum(v), 0),
  AVERAGE: (a) => { const f = flat(a).map(toNum); return f.reduce((s, v) => s + v, 0) / f.length; },
  MEDIAN: (a) => { const s = flat(a).map(toNum).sort((a,b)=>a-b); const m = Math.floor(s.length/2); return s.length%2 ? s[m] : (s[m-1]+s[m])/2; },
  COUNT: (a) => flat(a).filter(v => typeof v === 'number' || !isNaN(+v)).length,
  COUNTA: (a) => flat(a).filter(v => v !== '' && v !== null && v !== undefined).length,
  COUNTIF: (a) => { const [range, crit] = a; return flat(Array.isArray(range)?range:[range]).filter(v => matchCriteria(v, crit)).length; },
  SUMIF: (a) => { const [range, crit, sr] = a; const r = flat(Array.isArray(range)?range:[range]); const s = sr ? flat(Array.isArray(sr)?sr:[sr]) : r; let t = 0; r.forEach((v,i) => { if(matchCriteria(v,crit)) t += toNum(s[i]??0); }); return t; },
  AVERAGEIF: (a) => { const [range, crit, ar] = a; const r = flat(Array.isArray(range)?range:[range]); const s = ar ? flat(Array.isArray(ar)?ar:[ar]) : r; let t = 0, c = 0; r.forEach((v,i) => { if(matchCriteria(v,crit)){t += toNum(s[i]??0); c++;} }); return c ? t/c : 0; },
  MIN: (a) => Math.min(...flat(a).map(toNum).filter(n => !isNaN(n))),
  MAX: (a) => Math.max(...flat(a).map(toNum).filter(n => !isNaN(n))),
  ABS: (a) => Math.abs(toNum(a[0])),
  SQRT: (a) => Math.sqrt(toNum(a[0])),
  POWER: (a) => Math.pow(toNum(a[0]), toNum(a[1])),
  MOD: (a) => toNum(a[0]) % toNum(a[1]),
  ROUND: (a) => { const [n, d=0] = a; return +toNum(n).toFixed(toNum(d)); },
  ROUNDUP: (a) => { const [n, d=0] = a; const f = Math.pow(10, toNum(d)); return Math.ceil(toNum(n)*f)/f; },
  ROUNDDOWN: (a) => { const [n, d=0] = a; const f = Math.pow(10, toNum(d)); return Math.floor(toNum(n)*f)/f; },
  CEILING: (a) => { const [n, s=1] = a; return Math.ceil(toNum(n)/toNum(s))*toNum(s); },
  FLOOR: (a) => { const [n, s=1] = a; return Math.floor(toNum(n)/toNum(s))*toNum(s); },
  RAND: () => Math.random(),
  RANDBETWEEN: (a) => { const [lo, hi] = a.map(toNum); return Math.floor(Math.random()*(hi-lo+1))+lo; },
  PI: () => Math.PI,
  PRODUCT: (a) => flat(a).map(toNum).reduce((p, v) => p*v, 1),
  INT: (a) => Math.trunc(toNum(a[0])),
  SIGN: (a) => Math.sign(toNum(a[0])),
  LOG: (a) => a.length > 1 ? Math.log(toNum(a[0]))/Math.log(toNum(a[1])) : Math.log10(toNum(a[0])),
  LOG10: (a) => Math.log10(toNum(a[0])),
  LN: (a) => Math.log(toNum(a[0])),
  EXP: (a) => Math.exp(toNum(a[0])),
  SIN: (a) => Math.sin(toNum(a[0])),
  COS: (a) => Math.cos(toNum(a[0])),
  TAN: (a) => Math.tan(toNum(a[0])),
  ASIN: (a) => Math.asin(toNum(a[0])),
  ACOS: (a) => Math.acos(toNum(a[0])),
  ATAN: (a) => Math.atan(toNum(a[0])),
  DEGREES: (a) => toNum(a[0]) * (180/Math.PI),
  RADIANS: (a) => toNum(a[0]) * (Math.PI/180),
  LARGE: (a) => { const s = flat(Array.isArray(a[0])?a[0]:[a[0]]).map(toNum).sort((x,y)=>y-x); return s[toNum(a[1])-1] ?? '#NUM'; },
  SMALL: (a) => { const s = flat(Array.isArray(a[0])?a[0]:[a[0]]).map(toNum).sort((x,y)=>x-y); return s[toNum(a[1])-1] ?? '#NUM'; },
  RANK: (a) => { const [v, r] = a; const arr = flat(Array.isArray(r)?r:[r]).map(toNum).sort((x,y)=>y-x); return arr.indexOf(toNum(v))+1 || '#N/A'; },
  STDEV: (a) => { const f = flat(a).map(toNum); const avg = f.reduce((s,v)=>s+v,0)/f.length; return Math.sqrt(f.reduce((s,v)=>s+(v-avg)**2,0)/(f.length-1)); },
  VAR: (a) => { const f = flat(a).map(toNum); const avg = f.reduce((s,v)=>s+v,0)/f.length; return f.reduce((s,v)=>s+(v-avg)**2,0)/(f.length-1); },
  MODE: (a) => { const f = flat(a).map(toNum); const freq = {}; f.forEach(v => freq[v]=(freq[v]||0)+1); return +Object.entries(freq).sort((a,b)=>b[1]-a[1])[0][0]; },
  PERCENTILE: (a) => { const arr = flat(Array.isArray(a[0])?a[0]:[a[0]]).map(toNum).sort((x,y)=>x-y); const k = toNum(a[1]); const i = k*(arr.length-1); const lo = Math.floor(i); const hi = Math.ceil(i); return lo===hi ? arr[lo] : arr[lo]+(arr[hi]-arr[lo])*(i-lo); },

  // Logic
  IF: (a) => a[0] ? a[1] : a[2],
  AND: (a) => flat(a).every(Boolean),
  OR: (a) => flat(a).some(Boolean),
  NOT: (a) => !a[0],
  IFERROR: (a) => { return (a[0] !== '#ERROR' && a[0] !== '#ERR' && !String(a[0]).startsWith('#')) ? a[0] : a[1]; },
  IFS: (a) => { for (let i = 0; i < a.length; i += 2) { if (a[i]) return a[i+1]; } return '#N/A'; },
  SWITCH: (a) => { const v = a[0]; for (let i = 1; i < a.length-1; i += 2) { if (v == a[i]) return a[i+1]; } return a.length%2===0 ? a[a.length-1] : '#N/A'; },
  CHOOSE: (a) => a[toNum(a[0])] ?? '#VALUE',
  ISBLANK: (a) => a[0] === '' || a[0] === null || a[0] === undefined || a[0] === 0,
  ISNUMBER: (a) => typeof a[0] === 'number' || !isNaN(+a[0]),
  ISTEXT: (a) => typeof a[0] === 'string' && isNaN(+a[0]),
  ISERROR: (a) => String(a[0]).startsWith('#'),
  ISEVEN: (a) => toNum(a[0]) % 2 === 0,
  ISODD: (a) => toNum(a[0]) % 2 !== 0,

  // Lookup
  VLOOKUP: () => '#USE_SPECIAL',
  HLOOKUP: () => '#USE_SPECIAL',
  INDEX: (a) => { if (Array.isArray(a[0])) return a[0][toNum(a[1])-1] ?? '#REF'; return a[0]; },
  MATCH: (a) => { const arr = flat(Array.isArray(a[1])?a[1]:[a[1]]); const idx = arr.findIndex(v => v == a[0] || String(v).toLowerCase()===String(a[0]).toLowerCase()); return idx >= 0 ? idx+1 : '#N/A'; },
  ROW: () => selectedCell.row + 1,
  COLUMN: () => selectedCell.col + 1,
  ROWS: (a) => Array.isArray(a[0]) ? a[0].length : 1,
  COLUMNS: () => 1,

  // Text
  CONCATENATE: (a) => flat(a).join(''),
  CONCAT: (a) => flat(a).join(''),
  TEXTJOIN: (a) => { const [d, skip, ...rest] = a; const vals = flat(rest); return (skip ? vals.filter(v => v !== '' && v != null) : vals).join(d); },
  LEFT: (a) => String(a[0]).substring(0, toNum(a[1] ?? 1)),
  RIGHT: (a) => { const s = String(a[0]); return s.substring(s.length - toNum(a[1] ?? 1)); },
  MID: (a) => String(a[0]).substring(toNum(a[1])-1, toNum(a[1])-1+toNum(a[2])),
  LEN: (a) => String(a[0]).length,
  UPPER: (a) => String(a[0]).toUpperCase(),
  LOWER: (a) => String(a[0]).toLowerCase(),
  PROPER: (a) => String(a[0]).replace(/\b\w/g, c => c.toUpperCase()),
  TRIM: (a) => String(a[0]).trim(),
  CLEAN: (a) => String(a[0]).replace(/[\x00-\x1F]/g, ''),
  EXACT: (a) => String(a[0]) === String(a[1]),
  REPLACE: (a) => { const s = String(a[0]); const st = toNum(a[1])-1; const len = toNum(a[2]); return s.substring(0,st)+String(a[3])+s.substring(st+len); },
  SUBSTITUTE: (a) => { const [text, old, rep, nth] = a; if (nth) { let i = 0; return String(text).replace(new RegExp(escapeRegex(String(old)), 'g'), m => (++i===toNum(nth)) ? rep : m); } return String(text).split(String(old)).join(String(rep)); },
  FIND: (a) => { const idx = String(a[1]).indexOf(String(a[0]), toNum(a[2]??1)-1); return idx >= 0 ? idx+1 : '#VALUE'; },
  SEARCH: (a) => { const idx = String(a[1]).toLowerCase().indexOf(String(a[0]).toLowerCase(), toNum(a[2]??1)-1); return idx >= 0 ? idx+1 : '#VALUE'; },
  REPT: (a) => { const s = String(a[0]); const n = Math.min(toNum(a[1]), Math.floor(10000 / Math.max(s.length, 1))); return s.repeat(Math.max(0, n)); },
  TEXT: (a) => formatText(a[0], String(a[1])),
  VALUE: (a) => toNum(String(a[0]).replace(/[^0-9.\-]/g, '')),
  DOLLAR: (a) => '$' + toNum(a[0]).toFixed(toNum(a[1] ?? 2)),
  FIXED: (a) => toNum(a[0]).toFixed(toNum(a[1] ?? 2)),
  CHAR: (a) => String.fromCharCode(toNum(a[0])),
  CODE: (a) => String(a[0]).charCodeAt(0),
  T: (a) => typeof a[0] === 'string' ? a[0] : '',
  N: (a) => typeof a[0] === 'number' ? a[0] : 0,

  // Date
  NOW: () => new Date().toLocaleString(),
  TODAY: () => new Date().toLocaleDateString(),
  YEAR: (a) => new Date(a[0]).getFullYear(),
  MONTH: (a) => new Date(a[0]).getMonth() + 1,
  DAY: (a) => new Date(a[0]).getDate(),
  HOUR: (a) => new Date(a[0]).getHours(),
  MINUTE: (a) => new Date(a[0]).getMinutes(),
  SECOND: (a) => new Date(a[0]).getSeconds(),
  DAYS: (a) => Math.round((new Date(a[0]) - new Date(a[1])) / 86400000),
  EDATE: (a) => { const d = new Date(a[0]); d.setMonth(d.getMonth()+toNum(a[1])); return d.toLocaleDateString(); },
  EOMONTH: (a) => { const d = new Date(a[0]); d.setMonth(d.getMonth()+toNum(a[1])+1, 0); return d.toLocaleDateString(); },
  DATE: (a) => new Date(toNum(a[0]), toNum(a[1])-1, toNum(a[2])).toLocaleDateString(),
  WEEKDAY: (a) => new Date(a[0]).getDay() + 1,
  WEEKNUM: (a) => { const d = new Date(a[0]); const s = new Date(d.getFullYear(),0,1); return Math.ceil(((d-s)/86400000+s.getDay()+1)/7); },
  NETWORKDAYS: (a) => { const s = new Date(a[0]); const e = new Date(a[1]); if (isNaN(s) || isNaN(e)) return '#VALUE'; const days = Math.round((e - s) / 86400000); if (Math.abs(days) > 3660) return '#VALUE'; let c = 0; const d = new Date(s); while(d<=e){const dy=d.getDay(); if(dy!==0&&dy!==6) c++; d.setDate(d.getDate()+1);} return c; },
  DATEDIF: (a) => { const s = new Date(a[0]); const e = new Date(a[1]); const u = String(a[2]).toUpperCase(); if(u==='D') return Math.round((e-s)/86400000); if(u==='M') return (e.getFullYear()-s.getFullYear())*12+e.getMonth()-s.getMonth(); if(u==='Y') return e.getFullYear()-s.getFullYear(); return '#VALUE'; },

  // Financial
  PMT: (a) => { const [rate,nper,pv,fv=0] = a.map(toNum); if(rate===0) return -(pv+fv)/nper; return -(pv*rate*Math.pow(1+rate,nper)+fv*rate)/(Math.pow(1+rate,nper)-1); },
  FV: (a) => { const [rate,nper,pmt,pv=0] = a.map(toNum); if(rate===0) return -(pv+pmt*nper); return -(pv*Math.pow(1+rate,nper)+pmt*(Math.pow(1+rate,nper)-1)/rate); },
  PV: (a) => { const [rate,nper,pmt,fv=0] = a.map(toNum); if(rate===0) return -(fv+pmt*nper); return -(fv/Math.pow(1+rate,nper)+pmt*(1-Math.pow(1+rate,-nper))/rate); },
  NPV: (a) => { const rate = toNum(a[0]); return flat(a.slice(1)).map(toNum).reduce((s,cf,i) => s+cf/Math.pow(1+rate,i+1), 0); },
  IRR: (a) => { const flows = flat(Array.isArray(a[0])?a[0]:a).map(toNum); let g = 0.1; for(let i=0;i<100;i++){let npv=0,dn=0; flows.forEach((cf,j)=>{npv+=cf/Math.pow(1+g,j); dn-=j*cf/Math.pow(1+g,j+1);}); const n=g-npv/dn; if(Math.abs(n-g)<1e-10) return Math.round(n*1e8)/1e8; g=n;} return '#NUM'; },
};

function callFunction(name, args) {
  const fn = FORMULA_FUNCS[name];
  if (!fn) return '#NAME?';
  return fn(args);
}

function matchCriteria(val, criteria) {
  const s = String(criteria);
  if (s.startsWith('>=')) return toNum(val) >= toNum(s.slice(2));
  if (s.startsWith('<=')) return toNum(val) <= toNum(s.slice(2));
  if (s.startsWith('<>')) return String(val) !== s.slice(2);
  if (s.startsWith('>')) return toNum(val) > toNum(s.slice(1));
  if (s.startsWith('<')) return toNum(val) < toNum(s.slice(1));
  if (s.includes('*') || s.includes('?')) {
    const escaped = s.replace(/[.+^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp('^' + escaped.replace(/\\\*/g, '.*').replace(/\\\?/g, '.') + '$', 'i');
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

function evalHlookup(expr, sheet, sheetIdx) {
  const inner = expr.slice(8, -1);
  const parts = splitTopLevel(inner);
  if (parts.length < 3) return '#ERR';

  const needle = resolveValue(parts[0].trim(), sheet, sheetIdx);
  const rowIdx = toNum(resolveValue(parts[2].trim(), sheet, sheetIdx));

  const rangeMatch = parts[1].trim().match(/^([A-Z])(\d+):([A-Z])(\d+)$/);
  if (!rangeMatch) return '#ERR';

  const c1 = rangeMatch[1].charCodeAt(0) - 65;
  const r1 = +rangeMatch[2] - 1;
  const c2 = rangeMatch[3].charCodeAt(0) - 65;
  const r2 = +rangeMatch[4] - 1;

  // Search first row for needle
  for (let c = c1; c <= c2; c++) {
    const lookupCell = sheet.cells[cellKey(r1, c)];
    const lookupVal = lookupCell ? (lookupCell.formula ? evaluateFormula(lookupCell.formula, sheetIdx) : lookupCell.value) : '';
    if (String(lookupVal).toLowerCase() === String(needle).toLowerCase() || lookupVal == needle) {
      const resultCell = sheet.cells[cellKey(r1 + rowIdx - 1, c)];
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

  // Deadline coloring takes priority
  const dl = getDeadlineColor(cell);
  if (dl) {
    parts.push(`background:${dl.bg}`);
    parts.push(`color:${dl.text}`);
    parts.push(`border-left:3px solid ${dl.border}`);
  } else {
    const tc = sanitizeColor(cell.textColor);
    if (tc) parts.push(`color:${tc}`);
    const cc = sanitizeColor(cell._condColor);
    if (cc) parts.push(`background:${cc}33;border-left:3px solid ${cc}`);
    else { const fc = sanitizeColor(cell.fillColor); if (fc && fc !== '#1a1a2e') parts.push(`background:${fc}`); }
  }

  if (cell.align && /^(left|center|right)$/.test(cell.align)) parts.push(`justify-content:${cell.align === 'left' ? 'flex-start' : cell.align === 'right' ? 'flex-end' : 'center'}`);
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

  // Cancel pending formula pick
  if (pendingFormula && e.key === 'Escape') { cancelPendingFormula(); return; }

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
    // Start typing directly - let the keystroke naturally go into the input
    startEdit(selectedCell.row, selectedCell.col);
    const input = getCellTd(selectedCell.row, selectedCell.col)?.querySelector('input');
    if (input) { input.value = ''; }
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

  // Check if we have a multi-cell selection to show quick formulas
  const hasRange = selectionRange &&
    (selectionRange.startRow !== selectionRange.endRow || selectionRange.startCol !== selectionRange.endCol);
  const hasMulti = multiSelection.length >= 2;

  let formulaRef = '';
  if (hasMulti) {
    formulaRef = multiSelection.map(s => COL_LETTERS[s.col] + (s.row + 1)).join(',');
  } else if (hasRange) {
    formulaRef = getRangeRef();
  }

  let html = `
    <button onclick="cutSelection();removeContextMenu()">Cut</button>
    <button onclick="copySelection();removeContextMenu()">Copy</button>
    <button onclick="pasteSelection();removeContextMenu()">Paste</button>`;

  if (hasRange || hasMulti) {
    html += `
    <div class="separator"></div>
    <div class="context-submenu">
      <button class="submenu-trigger" onclick="toggleSubmenu(event)">Quick Formula &rarr;</button>
      <div class="submenu">
        <button onclick="quickFormula('SUM','${formulaRef}');removeContextMenu()">SUM</button>
        <button onclick="quickFormula('AVERAGE','${formulaRef}');removeContextMenu()">AVERAGE</button>
        <button onclick="quickFormula('COUNT','${formulaRef}');removeContextMenu()">COUNT</button>
        <button onclick="quickFormula('MIN','${formulaRef}');removeContextMenu()">MIN</button>
        <button onclick="quickFormula('MAX','${formulaRef}');removeContextMenu()">MAX</button>
        <button onclick="quickFormula('MEDIAN','${formulaRef}');removeContextMenu()">MEDIAN</button>
        <button onclick="quickFormula('PRODUCT','${formulaRef}');removeContextMenu()">PRODUCT</button>
        <button onclick="quickFormula('STDEV','${formulaRef}');removeContextMenu()">STDEV</button>
        <button onclick="quickFormula('COUNTIF','${formulaRef}');removeContextMenu()">COUNTIF...</button>
        <button onclick="quickFormula('SUMIF','${formulaRef}');removeContextMenu()">SUMIF...</button>
      </div>
    </div>`;
  }

  html += `
    <div class="separator"></div>
    <button onclick="insertToday(${row}, ${col});removeContextMenu()">Insert Today's Date</button>
    <button onclick="showDatePicker(${row}, ${col});removeContextMenu()">Pick a Date...</button>
    <div class="separator"></div>
    <button onclick="toggleDeadline(${row}, ${col});removeContextMenu()">Toggle Deadline</button>
    <div class="separator"></div>
    <button onclick="insertRowAbove(${row});removeContextMenu()">Insert Row Above</button>
    <button onclick="insertRowBelow(${row});removeContextMenu()">Insert Row Below</button>
    <div class="separator"></div>
    <button onclick="deleteSelection();removeContextMenu()">Clear Contents</button>
    <button onclick="sortColumn(${col}, true);removeContextMenu()">Sort A &rarr; Z</button>
    <button onclick="sortColumn(${col}, false);removeContextMenu()">Sort Z &rarr; A</button>
    <div class="separator"></div>
    <button onclick="freezeAt(${row}, ${col});removeContextMenu()">Freeze Above & Left</button>
    <button onclick="unfreeze();removeContextMenu()">Unfreeze Panes</button>
    <div class="separator"></div>
    <button onclick="showConditionalFormat(${row}, ${col});removeContextMenu()">Conditional Format...</button>
    <button onclick="showDataValidation(${row}, ${col});removeContextMenu()">Data Validation...</button>
    <div class="separator"></div>
    <button onclick="hideRow(${row});removeContextMenu()">Hide Row ${row + 1}</button>
    <button onclick="hideColumn(${col});removeContextMenu()">Hide Column ${COL_LETTERS[col]}</button>`;

  // Show unhide options if adjacent rows/cols are hidden
  const sheet = sheets[activeSheet];
  const hRows = sheet.hiddenRows || [];
  const hCols = sheet.hiddenCols || [];
  if (hRows.includes(row - 1) || hRows.includes(row + 1)) {
    html += `<button onclick="unhideRow(${row});removeContextMenu()">Unhide Adjacent Rows</button>`;
  }
  if (hCols.includes(col - 1) || hCols.includes(col + 1)) {
    html += `<button onclick="unhideColumn(${col});removeContextMenu()">Unhide Adjacent Columns</button>`;
  }

  // Comment options
  const commentCell = sheet.cells[cellKey(row, col)] || {};
  if (commentCell.comment) {
    html += `
    <div class="separator"></div>
    <button onclick="addComment(${row}, ${col});removeContextMenu()">Edit Comment...</button>
    <button onclick="deleteComment(${row}, ${col});removeContextMenu()">Delete Comment</button>`;
  } else {
    html += `
    <div class="separator"></div>
    <button onclick="addComment(${row}, ${col});removeContextMenu()">Add Comment...</button>`;
  }
  html += '';

  menu.innerHTML = html;
  document.body.appendChild(menu);

  // Adjust if off-screen
  menu.style.maxHeight = (window.innerHeight - 20) + 'px';
  menu.style.overflowY = 'auto';
  const rect = menu.getBoundingClientRect();
  if (rect.right > window.innerWidth) menu.style.left = Math.max(5, x - rect.width) + 'px';
  if (rect.bottom > window.innerHeight) menu.style.top = Math.max(5, window.innerHeight - rect.height - 10) + 'px';

  setTimeout(() => document.addEventListener('click', removeContextMenu, { once: true }), 10);
}

function getRangeRef() {
  if (!selectionRange) return '';
  const r1 = Math.min(selectionRange.startRow, selectionRange.endRow);
  const r2 = Math.max(selectionRange.startRow, selectionRange.endRow);
  const c1 = Math.min(selectionRange.startCol, selectionRange.endCol);
  const c2 = Math.max(selectionRange.startCol, selectionRange.endCol);
  return COL_LETTERS[c1] + (r1 + 1) + ':' + COL_LETTERS[c2] + (r2 + 1);
}

function toggleDeadline(row, col) {
  const key = cellKey(row, col);
  if (!sheets[activeSheet].cells[key]) sheets[activeSheet].cells[key] = {};
  const cell = sheets[activeSheet].cells[key];
  cell.deadline = !cell.deadline;
  if (cell.deadline) {
    document.getElementById('status-info').textContent = `Deadline set on ${COL_LETTERS[col]}${row + 1}`;
  } else {
    document.getElementById('status-info').textContent = `Deadline removed from ${COL_LETTERS[col]}${row + 1}`;
  }
  renderSheet();
  selectCell(row, col);
  triggerAutoSave();
}

function getDeadlineColor(cell) {
  if (!cell || !cell.deadline) return null;

  // Get the display value (resolve formula if needed)
  const val = cell.formula ? getDisplayValue(cell) : (cell.value ?? '');
  const dateVal = new Date(val);
  if (isNaN(dateVal.getTime())) return null;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  dateVal.setHours(0, 0, 0, 0);
  const daysLeft = Math.round((dateVal - today) / (1000 * 60 * 60 * 24));

  if (daysLeft < 0) return { bg: '#ff4757', text: '#fff', border: '#ff2234' };       // overdue - red
  if (daysLeft === 0) return { bg: '#ff6348', text: '#fff', border: '#ff4520' };      // today - deep orange
  if (daysLeft <= 2) return { bg: '#ff7f50', text: '#fff', border: '#ff6330' };       // 1-2 days - orange
  if (daysLeft <= 5) return { bg: '#ffa502', text: '#1a1a2e', border: '#e89400' };    // 3-5 days - amber
  if (daysLeft <= 7) return { bg: '#fdcb6e', text: '#1a1a2e', border: '#f0b830' };    // 6-7 days - yellow
  if (daysLeft <= 14) return { bg: '#7bed9f', text: '#1a1a2e', border: '#55d97e' };   // 1-2 weeks - light green
  return { bg: '#2ed573', text: '#1a1a2e', border: '#17b558' };                        // 2+ weeks - green
}

const notifiedDeadlines = new Set();

function checkDeadlines() {
  if (!('Notification' in window) || Notification.permission !== 'granted') return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  sheets.forEach((sheet, si) => {
    for (const key of Object.keys(sheet.cells)) {
      const cell = sheet.cells[key];
      if (!cell || !cell.deadline) continue;

      const val = cell.formula ? getDisplayValue(cell) : (cell.value ?? '');
      const dateVal = new Date(val);
      if (isNaN(dateVal.getTime())) continue;

      dateVal.setHours(0, 0, 0, 0);
      const daysLeft = Math.round((dateVal - today) / (1000 * 60 * 60 * 24));

      // Only notify for overdue, today, or within 2 days
      if (daysLeft > 2) continue;

      // Don't re-notify the same cell on the same day
      const notifKey = `${si}-${key}-${today.toDateString()}`;
      if (notifiedDeadlines.has(notifKey)) continue;
      notifiedDeadlines.add(notifKey);

      const { row, col } = parseKey(key);
      const cellRef = COL_LETTERS[col] + (row + 1);
      let msg;
      if (daysLeft < 0) msg = `Overdue by ${Math.abs(daysLeft)} day${Math.abs(daysLeft) > 1 ? 's' : ''}!`;
      else if (daysLeft === 0) msg = `Due today!`;
      else msg = `Due in ${daysLeft} day${daysLeft > 1 ? 's' : ''}`;

      new Notification(`📅 Deadline: ${cellRef} (${sheet.name})`, {
        body: `${val} — ${msg}`,
        icon: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="8" fill="%236c5ce7"/><text x="16" y="23" text-anchor="middle" fill="white" font-size="20" font-weight="800">Q</text></svg>'
      });
    }
  });
}

function insertToday(row, col) {
  selectCell(row, col);
  setCellValue(row, col, '=TODAY()');
  renderSheet();
  selectCell(row, col);
}

function showDatePicker(row, col) {
  // Remove any existing picker
  document.querySelector('.date-picker-popup')?.remove();

  const td = getCellTd(row, col);
  const rect = td ? td.getBoundingClientRect() : { left: 100, bottom: 100 };

  const picker = document.createElement('div');
  picker.className = 'date-picker-popup';
  picker.style.left = rect.left + 'px';
  picker.style.top = rect.bottom + 4 + 'px';

  const input = document.createElement('input');
  input.type = 'date';
  input.valueAsDate = new Date();
  input.addEventListener('change', () => {
    if (input.value) {
      const parts = input.value.split('-');
      const val = parseInt(parts[1]) + '/' + parseInt(parts[2]) + '/' + parts[0];
      selectCell(row, col);
      setCellValue(row, col, val);
      const key = cellKey(row, col);
      sheets[activeSheet].cells[key].detectedType = 'date';
      renderSheet();
      selectCell(row, col);
    }
    picker.remove();
  });

  const cancel = document.createElement('button');
  cancel.textContent = 'Cancel';
  cancel.onclick = () => picker.remove();

  picker.appendChild(input);
  picker.appendChild(cancel);
  document.body.appendChild(picker);

  // Adjust if off-screen
  const pr = picker.getBoundingClientRect();
  if (pr.right > window.innerWidth) picker.style.left = (window.innerWidth - pr.width - 10) + 'px';
  if (pr.bottom > window.innerHeight) picker.style.top = (rect.top - pr.height - 4) + 'px';

  input.showPicker?.();
  input.focus();

  setTimeout(() => {
    document.addEventListener('mousedown', function handler(e) {
      if (!picker.contains(e.target)) { picker.remove(); document.removeEventListener('mousedown', handler); }
    });
  }, 10);
}

let pendingFormula = null;

function quickFormula(func, rangeRef) {
  // For COUNTIF/SUMIF, prompt for criteria first
  let formula;
  if (func === 'COUNTIF') {
    const criteria = prompt('Count cells where value:', '>0');
    if (criteria === null) return;
    formula = `=${func}(${rangeRef},"${criteria}")`;
  } else if (func === 'SUMIF') {
    const criteria = prompt('Sum cells where value:', '>0');
    if (criteria === null) return;
    formula = `=${func}(${rangeRef},"${criteria}")`;
  } else {
    formula = `=${func}(${rangeRef})`;
  }

  // Enter "pick cell" mode
  pendingFormula = { formula, func };
  document.getElementById('status-info').textContent = `Click a cell to place ${func} result (Esc to cancel)`;
  document.getElementById('sheet-container').classList.add('picking-cell');
}

function placePendingFormula(row, col) {
  if (!pendingFormula) return;
  const { formula, func } = pendingFormula;
  pendingFormula = null;
  document.getElementById('sheet-container').classList.remove('picking-cell');
  selectCell(row, col);
  setCellValue(row, col, formula);
  document.getElementById('status-info').textContent = `${func} inserted at ${COL_LETTERS[col]}${row + 1}`;
}

function cancelPendingFormula() {
  if (!pendingFormula) return;
  pendingFormula = null;
  document.getElementById('sheet-container').classList.remove('picking-cell');
  document.getElementById('status-info').textContent = 'Ready';
}

function toggleSubmenu(e) {
  e.stopPropagation();
  const trigger = e.target;
  const submenu = trigger.nextElementSibling;
  submenu.classList.toggle('show');

  if (submenu.classList.contains('show')) {
    const triggerRect = trigger.getBoundingClientRect();
    // Try to open to the right
    let left = triggerRect.right + 2;
    let top = triggerRect.top;

    // If it would go off-screen right, open to the left
    submenu.style.left = left + 'px';
    submenu.style.top = top + 'px';

    const subRect = submenu.getBoundingClientRect();
    if (subRect.right > window.innerWidth) {
      left = triggerRect.left - subRect.width - 2;
    }
    if (subRect.bottom > window.innerHeight) {
      top = window.innerHeight - subRect.height - 5;
    }
    submenu.style.left = Math.max(5, left) + 'px';
    submenu.style.top = Math.max(5, top) + 'px';
  }
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
  if (!isPro() && sheets.length >= FREE_MAX_SHEETS) {
    showProModal();
    document.getElementById('status-info').textContent = 'Free plan: max 3 sheets. Upgrade to Pro for unlimited.';
    return;
  }
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
          if (!data.sheets || !Array.isArray(data.sheets)) throw new Error('Invalid file');
          sheets = data.sheets.map(s => sanitizeSheetData(s));
          activeSheet = Math.max(0, Math.min(toNum(data.activeSheet || 0), sheets.length - 1));
          document.getElementById('file-name').value = String(data.fileName || file.name.replace('.qx', ''));
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
        row.push('');
      }
      data.push(row);
    }

    const ws = XLSX.utils.aoa_to_sheet(data);

    // Write values and formulas into worksheet cells
    for (let r = 0; r <= maxR; r++) {
      for (let c = 0; c <= maxC; c++) {
        const cell = sheet.cells[cellKey(r, c)];
        if (!cell) continue;
        const addr = XLSX.utils.encode_cell({ r, c });
        if (cell.formula) {
          // Strip leading '=' for XLSX format
          ws[addr] = { t: 'n', f: cell.formula.replace(/^=/, ''), v: cell.value ?? 0 };
        } else if (cell.value !== undefined && cell.value !== '') {
          const v = cell.value;
          if (typeof v === 'number' || (typeof v === 'string' && !isNaN(v) && v.trim() !== '')) {
            ws[addr] = { t: 'n', v: Number(v) };
          } else {
            ws[addr] = { t: 's', v: String(v) };
          }
        }
      }
    }

    // Free users: add branding row
    if (!isPro()) {
      const brandRow = maxR + 2;
      const addr = XLSX.utils.encode_cell({ r: brandRow, c: 0 });
      ws[addr] = { t: 's', v: 'Created with Anomaly Quantix - anomaly-surround.github.io/anomaly-quantix' };
      ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: brandRow, c: maxC } });
    }

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
      // Prevent CSV formula injection
      if (/^[=+\-@\t\r]/.test(val)) { val = "'" + val; }
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

  // Place chart on the sheet
  const targetStr = document.getElementById('chart-target').value.toUpperCase().trim();
  placeChartOnSheet(type, title, labels, datasets, rangeStr, targetStr);
  document.getElementById('chart-modal').style.display = 'none';
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
// Floating Charts on Sheet
// ============================================

let sheetCharts = [];

function placeChartOnSheet(type, title, labels, datasets, rangeStr, targetStr) {
  const container = document.getElementById('sheet-container');
  const id = 'chart-' + Date.now();

  // Parse target cell range
  let anchorCells = null;
  const targetMatch = targetStr.match(/^([A-Z])(\d+):([A-Z])(\d+)$/);
  if (targetMatch) {
    anchorCells = {
      c1: targetMatch[1].charCodeAt(0) - 65,
      r1: +targetMatch[2] - 1,
      c2: targetMatch[3].charCodeAt(0) - 65,
      r2: +targetMatch[4] - 1
    };
  }

  const wrapper = document.createElement('div');
  wrapper.className = 'floating-chart' + (anchorCells ? ' anchored-chart' : '');
  wrapper.id = id;

  const header = document.createElement('div');
  header.className = 'floating-chart-header';
  header.innerHTML = `<span class="floating-chart-title">${escapeHTML(title)}</span>
    <div class="floating-chart-actions">
      <button class="floating-chart-btn" onclick="refreshSheetChart('${id}')" title="Refresh data">&#x21bb;</button>
      <button class="floating-chart-btn" onclick="removeSheetChart('${id}')" title="Close">&times;</button>
    </div>`;

  const canvas = document.createElement('canvas');
  canvas.width = 500;
  canvas.height = 300;
  canvas.className = 'floating-chart-canvas';

  wrapper.appendChild(header);
  wrapper.appendChild(canvas);

  const chartAC = new AbortController();

  if (!anchorCells) {
    // Floating mode: add resize handle and drag/resize logic
    const resizeHandle = document.createElement('div');
    resizeHandle.className = 'floating-chart-resize';
    wrapper.appendChild(resizeHandle);
    wrapper.style.left = (container.scrollLeft + 60) + 'px';
    wrapper.style.top = (container.scrollTop + 30) + 'px';

    setupDrag(wrapper, header, chartAC);
    setupResize(wrapper, canvas, resizeHandle, type, title, labels, datasets, chartAC);
  }

  container.appendChild(wrapper);

  const chartData = { id, type, title, labels, datasets, rangeStr, wrapper, canvas, anchorCells, ac: chartAC };
  sheetCharts.push(chartData);

  if (anchorCells) {
    positionAnchoredChart(chartData);
  }

  const ctx = canvas.getContext('2d');
  drawChart(ctx, canvas, type, title, labels, datasets);
}

function positionAnchoredChart(chart) {
  if (!chart.anchorCells) return;
  const { c1, r1, c2, r2 } = chart.anchorCells;

  // Find bounding box from actual cell positions
  const topLeftTd = getCellTd(r1, c1);
  const bottomRightTd = getCellTd(r2, c2);
  const container = document.getElementById('sheet-container');
  const table = document.getElementById('sheet');

  if (topLeftTd && bottomRightTd) {
    const tableRect = table.getBoundingClientRect();
    const tlRect = topLeftTd.getBoundingClientRect();
    const brRect = bottomRightTd.getBoundingClientRect();

    const left = tlRect.left - tableRect.left;
    const top = tlRect.top - tableRect.top;
    const width = brRect.right - tlRect.left;
    const height = brRect.bottom - tlRect.top;

    chart.wrapper.style.left = left + 'px';
    chart.wrapper.style.top = top + 'px';
    chart.wrapper.style.width = width + 'px';

    const headerH = chart.wrapper.querySelector('.floating-chart-header').offsetHeight || 28;
    const canvasH = Math.max(100, height - headerH);
    chart.canvas.width = width;
    chart.canvas.height = canvasH;
    chart.canvas.style.height = canvasH + 'px';

    drawChart(chart.canvas.getContext('2d'), chart.canvas, chart.type, chart.title, chart.labels, chart.datasets);
  } else {
    // Cells not visible yet — estimate from row/col sizes
    const sheet = sheets[activeSheet];
    let left = 40; // row header width
    for (let c = 0; c < c1; c++) left += (sheet.colWidths[c] || 80);
    let width = 0;
    for (let c = c1; c <= c2; c++) width += (sheet.colWidths[c] || 80);
    const top = r1 * ROW_HEIGHT;
    const height = (r2 - r1 + 1) * ROW_HEIGHT;

    chart.wrapper.style.left = left + 'px';
    chart.wrapper.style.top = top + 'px';
    chart.wrapper.style.width = width + 'px';

    const headerH = 28;
    const canvasH = Math.max(100, height - headerH);
    chart.canvas.width = width;
    chart.canvas.height = canvasH;
    chart.canvas.style.height = canvasH + 'px';

    drawChart(chart.canvas.getContext('2d'), chart.canvas, chart.type, chart.title, chart.labels, chart.datasets);
  }
}

function repositionAllCharts() {
  sheetCharts.forEach(chart => {
    if (chart.anchorCells) positionAnchoredChart(chart);
  });
}

function setupDrag(wrapper, header, ac) {
  let dragging = false, dragX, dragY, startLeft, startTop;
  const onStart = (x, y) => { dragging = true; dragX = x; dragY = y; startLeft = wrapper.offsetLeft; startTop = wrapper.offsetTop; };
  header.addEventListener('mousedown', (e) => { onStart(e.clientX, e.clientY); e.preventDefault(); });
  header.addEventListener('touchstart', (e) => { onStart(e.touches[0].clientX, e.touches[0].clientY); e.preventDefault(); });
  const onMove = (x, y) => { if (!dragging) return; wrapper.style.left = (startLeft + x - dragX) + 'px'; wrapper.style.top = (startTop + y - dragY) + 'px'; };
  const sig = ac ? { signal: ac.signal } : {};
  document.addEventListener('mousemove', (e) => onMove(e.clientX, e.clientY), sig);
  document.addEventListener('touchmove', (e) => onMove(e.touches[0].clientX, e.touches[0].clientY), sig);
  document.addEventListener('mouseup', () => { dragging = false; }, sig);
  document.addEventListener('touchend', () => { dragging = false; }, sig);
}

function setupResize(wrapper, canvas, handle, type, title, labels, datasets, ac) {
  let resizing = false, resizeX, resizeY, startW, startH;
  const onStart = (x, y) => { resizing = true; resizeX = x; resizeY = y; startW = canvas.width; startH = canvas.height; };
  handle.addEventListener('mousedown', (e) => { onStart(e.clientX, e.clientY); e.preventDefault(); e.stopPropagation(); });
  handle.addEventListener('touchstart', (e) => { onStart(e.touches[0].clientX, e.touches[0].clientY); e.preventDefault(); e.stopPropagation(); });
  const onMove = (x, y) => {
    if (!resizing) return;
    canvas.width = Math.max(250, startW + x - resizeX);
    canvas.height = Math.max(180, startH + y - resizeY);
    wrapper.style.width = canvas.width + 'px';
    drawChart(canvas.getContext('2d'), canvas, type, title, labels, datasets);
  };
  const sig = ac ? { signal: ac.signal } : {};
  document.addEventListener('mousemove', (e) => onMove(e.clientX, e.clientY), sig);
  document.addEventListener('touchmove', (e) => onMove(e.touches[0].clientX, e.touches[0].clientY), sig);
  document.addEventListener('mouseup', () => { resizing = false; }, sig);
  document.addEventListener('touchend', () => { resizing = false; }, sig);
}

function removeSheetChart(id) {
  const chart = sheetCharts.find(c => c.id === id);
  if (chart && chart.ac) chart.ac.abort();
  const el = document.getElementById(id);
  if (el) el.remove();
  sheetCharts = sheetCharts.filter(c => c.id !== id);
}

function refreshSheetChart(id) {
  const chart = sheetCharts.find(c => c.id === id);
  if (!chart) return;

  // Re-read data from the range
  const match = chart.rangeStr.match(/^([A-Z])(\d+):([A-Z])(\d+)$/);
  if (!match) return;

  const c1 = match[1].charCodeAt(0) - 65;
  const r1 = +match[2] - 1;
  const c2 = match[3].charCodeAt(0) - 65;
  const r2 = +match[4] - 1;

  const labels = [];
  const datasets = [];
  const numCols = c2 - c1 + 1;

  if (numCols >= 2) {
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
    for (let r = r1; r <= r2; r++) {
      labels.push('Row ' + (r + 1));
      const cell = sheets[activeSheet].cells[cellKey(r, c1)];
      if (!datasets[0]) datasets[0] = [];
      datasets[0].push(cell ? toNum(cell.value) : 0);
    }
  }

  chart.labels = labels;
  chart.datasets = datasets;
  const ctx = chart.canvas.getContext('2d');
  drawChart(ctx, chart.canvas, chart.type, chart.title, labels, datasets);
  document.getElementById('status-info').textContent = 'Chart refreshed';
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
  if (multiSelection.length >= 2) {
    multiSelection.forEach(s => fn(s.row, s.col));
  } else if (selectionRange) {
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

function insertAtCursor(input, text) {
  const start = input.selectionStart;
  const end = input.selectionEnd;
  const val = input.value;
  // If cursor is right after a letter/number, add the ref; otherwise just insert
  input.value = val.substring(0, start) + text + val.substring(end);
  const newPos = start + text.length;
  input.setSelectionRange(newPos, newPos);
  input.focus();
}

function escapeHTML(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(str) {
  return String(str).replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/'/g,'&#39;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function sanitizeSheetData(sheet) {
  const clean = createSheetData(String(sheet.name || 'Sheet').substring(0, 50));
  if (sheet.colWidths) {
    for (const [k, v] of Object.entries(sheet.colWidths)) {
      const idx = toNum(k);
      if (idx >= 0 && idx < COLS) clean.colWidths[idx] = Math.max(40, Math.min(500, toNum(v)));
    }
  }
  if (sheet.cells && typeof sheet.cells === 'object') {
    const allowed = ['value', 'formula', 'bold', 'italic', 'underline', 'textColor', 'fillColor', 'align', 'detectedType', 'formatType', 'validation', 'merge', '_mergedInto', '_condColor', 'comment', 'deadline'];
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (!cell || typeof cell !== 'object') continue;
      if (!/^\d+_\d+$/.test(key)) continue;
      const c = {};
      for (const prop of allowed) {
        if (cell[prop] !== undefined) {
          if ((prop === 'textColor' || prop === 'fillColor' || prop === '_condColor') && !isValidColor(cell[prop])) continue;
          if (prop === 'align' && !/^(left|center|right)$/.test(cell[prop])) continue;
          if (prop === 'formula' && typeof cell[prop] === 'string' && !cell[prop].startsWith('=')) continue;
          if (prop === 'bold' || prop === 'italic' || prop === 'underline') { c[prop] = !!cell[prop]; continue; }
          c[prop] = cell[prop];
        }
      }
      clean.cells[key] = c;
    }
  }
  // Preserve hiddenRows, hiddenCols, hiddenFilterRows
  if (Array.isArray(sheet.hiddenRows)) clean.hiddenRows = sheet.hiddenRows.filter(r => typeof r === 'number' && r >= 0 && r < ROWS);
  if (Array.isArray(sheet.hiddenCols)) clean.hiddenCols = sheet.hiddenCols.filter(c => typeof c === 'number' && c >= 0 && c < COLS);
  if (Array.isArray(sheet.hiddenFilterRows)) clean.hiddenFilterRows = sheet.hiddenFilterRows.filter(r => typeof r === 'number' && r >= 0 && r < ROWS);
  return clean;
}

function isValidColor(c) {
  return /^#[0-9a-fA-F]{3,8}$/.test(c);
}

function sanitizeColor(c) {
  return isValidColor(c) ? c : '';
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
// Feature: Auto-fill drag handle
// ============================================

let autofillDragging = false;
let autofillStart = null;
let autofillEnd = null;

function showAutofillHandle() {
  removeAutofillHandle();
  const td = getCellTd(selectedCell.row, selectedCell.col);
  if (!td) return;
  const handle = document.createElement('div');
  handle.className = 'autofill-handle';
  handle.addEventListener('mousedown', startAutofill);
  td.style.position = 'relative';
  td.appendChild(handle);
}

function removeAutofillHandle() {
  document.querySelectorAll('.autofill-handle').forEach(h => h.remove());
}

function startAutofill(e) {
  e.preventDefault();
  e.stopPropagation();
  autofillDragging = true;
  autofillStart = { row: selectedCell.row, col: selectedCell.col };
  autofillEnd = { row: selectedCell.row, col: selectedCell.col };

  const onMove = (e2) => {
    const td = e2.target.closest('td[data-row]');
    if (!td) return;
    autofillEnd = { row: +td.dataset.row, col: +td.dataset.col };
  };

  const onUp = () => {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    autofillDragging = false;
    performAutofill();
  };

  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

function performAutofill() {
  if (!autofillStart || !autofillEnd) return;
  const sr = autofillStart.row, sc = autofillStart.col;
  const er = autofillEnd.row, ec = autofillEnd.col;
  if (sr === er && sc === ec) return;

  const sheet = sheets[activeSheet];
  const srcKey = cellKey(sr, sc);
  const srcCell = sheet.cells[srcKey] || {};
  const srcVal = srcCell.value;
  const srcFormula = srcCell.formula;

  // Determine fill direction
  const fillDown = er > sr;
  const fillRight = ec > sc;

  // Detect sequences
  const DAYS = ['sunday','monday','tuesday','wednesday','thursday','friday','saturday'];
  const DAYS_SHORT = ['sun','mon','tue','wed','thu','fri','sat'];
  const MONTHS = ['january','february','march','april','may','june','july','august','september','october','november','december'];
  const MONTHS_SHORT = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];

  function detectSequenceIndex(val) {
    if (val === undefined || val === null) return null;
    const v = String(val).toLowerCase();
    let idx;
    idx = DAYS.indexOf(v); if (idx >= 0) return { seq: DAYS, idx, origCase: String(val) };
    idx = DAYS_SHORT.indexOf(v); if (idx >= 0) return { seq: DAYS_SHORT, idx, origCase: String(val) };
    idx = MONTHS.indexOf(v); if (idx >= 0) return { seq: MONTHS, idx, origCase: String(val) };
    idx = MONTHS_SHORT.indexOf(v); if (idx >= 0) return { seq: MONTHS_SHORT, idx, origCase: String(val) };
    return null;
  }

  function matchCase(str, ref) {
    if (ref === ref.toUpperCase()) return str.toUpperCase();
    if (ref[0] === ref[0].toUpperCase()) return str[0].toUpperCase() + str.slice(1);
    return str;
  }

  function adjustFormula(formula, rowOff, colOff) {
    return formula.replace(/([A-Z])(\d+)/gi, (m, col, row) => {
      const newCol = col.toUpperCase().charCodeAt(0) - 65 + colOff;
      const newRow = parseInt(row) + rowOff;
      if (newCol < 0 || newCol >= COLS || newRow < 1 || newRow > ROWS) return m;
      return String.fromCharCode(65 + newCol) + newRow;
    });
  }

  if (fillDown) {
    for (let r = sr + 1; r <= er; r++) {
      const offset = r - sr;
      const key = cellKey(r, sc);
      if (srcFormula) {
        const newFormula = adjustFormula(srcFormula, offset, 0);
        sheet.cells[key] = { ...(sheet.cells[key] || {}), formula: newFormula, value: evaluateFormula(newFormula, activeSheet) };
      } else {
        const seqInfo = detectSequenceIndex(srcVal);
        if (seqInfo) {
          const newIdx = (seqInfo.idx + offset) % seqInfo.seq.length;
          sheet.cells[key] = { ...(sheet.cells[key] || {}), value: matchCase(seqInfo.seq[newIdx], seqInfo.origCase), formula: undefined };
        } else if (typeof srcVal === 'number') {
          sheet.cells[key] = { ...(sheet.cells[key] || {}), value: srcVal + offset, formula: undefined, detectedType: 'number' };
        } else if (typeof srcVal === 'string' && (/^\d{4}-\d{2}-\d{2}$/.test(srcVal) || /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(srcVal))) {
          const d = new Date(srcVal);
          d.setDate(d.getDate() + offset);
          const ds = d.toISOString().split('T')[0];
          sheet.cells[key] = { ...(sheet.cells[key] || {}), value: ds, formula: undefined, detectedType: 'date' };
        } else {
          sheet.cells[key] = { ...(sheet.cells[key] || {}), value: srcVal, formula: undefined };
        }
      }
    }
  } else if (fillRight) {
    for (let c = sc + 1; c <= ec; c++) {
      const offset = c - sc;
      const key = cellKey(sr, c);
      if (srcFormula) {
        const newFormula = adjustFormula(srcFormula, 0, offset);
        sheet.cells[key] = { ...(sheet.cells[key] || {}), formula: newFormula, value: evaluateFormula(newFormula, activeSheet) };
      } else {
        const seqInfo = detectSequenceIndex(srcVal);
        if (seqInfo) {
          const newIdx = (seqInfo.idx + offset) % seqInfo.seq.length;
          sheet.cells[key] = { ...(sheet.cells[key] || {}), value: matchCase(seqInfo.seq[newIdx], seqInfo.origCase), formula: undefined };
        } else if (typeof srcVal === 'number') {
          sheet.cells[key] = { ...(sheet.cells[key] || {}), value: srcVal + offset, formula: undefined, detectedType: 'number' };
        } else {
          sheet.cells[key] = { ...(sheet.cells[key] || {}), value: srcVal, formula: undefined };
        }
      }
    }
  }

  recalcDependents();
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
  triggerAutoSave();
}

// ============================================
// Feature: Column Filters
// ============================================

let filtersActive = false;

function toggleFilters() {
  filtersActive = !filtersActive;
  const btn = document.getElementById('btn-filter');
  if (filtersActive) {
    btn.classList.add('filter-active');
    // Initialize hiddenRows on each sheet if not present
    sheets.forEach(s => { if (!s.hiddenFilterRows) s.hiddenFilterRows = []; });
  } else {
    btn.classList.remove('filter-active');
    // Clear all filter hidden rows
    sheets.forEach(s => { s.hiddenFilterRows = []; });
  }
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
}

function showFilterDropdown(colIdx, thEl) {
  removeFilterDropdown();
  const sheet = sheets[activeSheet];
  if (!sheet.hiddenFilterRows) sheet.hiddenFilterRows = [];

  // Collect unique values in this column
  const values = new Set();
  for (let r = 0; r < ROWS; r++) {
    const cell = sheet.cells[cellKey(r, colIdx)];
    if (cell && cell.value !== undefined && cell.value !== '') {
      values.add(String(getDisplayValue(cell)));
    }
  }

  const hiddenSet = new Set(sheet.hiddenFilterRows);

  const dropdown = document.createElement('div');
  dropdown.className = 'filter-dropdown';
  dropdown.id = 'active-filter-dropdown';
  const rect = thEl.getBoundingClientRect();
  dropdown.style.left = rect.left + 'px';
  dropdown.style.top = rect.bottom + 'px';

  let html = '';
  const sortedVals = [...values].sort();
  for (const val of sortedVals) {
    // Find rows with this value to check if they're hidden
    const rowsWithVal = [];
    for (let r = 0; r < ROWS; r++) {
      const cell = sheet.cells[cellKey(r, colIdx)];
      if (cell && String(getDisplayValue(cell)) === val) rowsWithVal.push(r);
    }
    const allHidden = rowsWithVal.every(r => hiddenSet.has(r));
    html += `<label><input type="checkbox" data-filter-col="${colIdx}" data-filter-val="${escapeAttr(val)}" ${allHidden ? '' : 'checked'}> ${escapeHTML(val)}</label>`;
  }
  html += `<div class="filter-actions"><button onclick="clearColumnFilter(${colIdx})">Clear</button><button onclick="applyFilterDropdown(${colIdx})">Apply</button></div>`;

  dropdown.innerHTML = html;
  document.body.appendChild(dropdown);

  // Adjust position
  const dr = dropdown.getBoundingClientRect();
  if (dr.right > window.innerWidth) dropdown.style.left = (window.innerWidth - dr.width - 5) + 'px';
  if (dr.bottom > window.innerHeight) dropdown.style.top = (rect.top - dr.height) + 'px';

  setTimeout(() => {
    document.addEventListener('click', function closeFilter(e) {
      if (!dropdown.contains(e.target)) {
        removeFilterDropdown();
        document.removeEventListener('click', closeFilter);
      }
    });
  }, 10);
}

function removeFilterDropdown() {
  const dd = document.getElementById('active-filter-dropdown');
  if (dd) dd.remove();
}

function applyFilterDropdown(colIdx) {
  const sheet = sheets[activeSheet];
  const checkboxes = document.querySelectorAll(`input[data-filter-col="${colIdx}"]`);
  const uncheckedVals = new Set();
  checkboxes.forEach(cb => {
    if (!cb.checked) uncheckedVals.add(cb.dataset.filterVal);
  });

  // Remove old filter hidden rows for this column, then re-add
  const newHidden = new Set(sheet.hiddenFilterRows || []);

  // First, unhide all rows that were hidden by this column's filter
  for (let r = 0; r < ROWS; r++) {
    const cell = sheet.cells[cellKey(r, colIdx)];
    if (cell && uncheckedVals.size === 0) {
      newHidden.delete(r);
    }
  }

  // Now hide rows with unchecked values
  for (let r = 0; r < ROWS; r++) {
    const cell = sheet.cells[cellKey(r, colIdx)];
    if (cell) {
      const dv = String(getDisplayValue(cell));
      if (uncheckedVals.has(dv)) {
        newHidden.add(r);
      } else {
        newHidden.delete(r);
      }
    }
  }

  sheet.hiddenFilterRows = [...newHidden];
  removeFilterDropdown();
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
  triggerAutoSave();
}

function clearColumnFilter(colIdx) {
  const sheet = sheets[activeSheet];
  // Remove rows hidden by filtering on this column
  const toRemove = new Set();
  for (let r = 0; r < ROWS; r++) {
    const cell = sheet.cells[cellKey(r, colIdx)];
    if (cell) toRemove.add(r);
  }
  sheet.hiddenFilterRows = (sheet.hiddenFilterRows || []).filter(r => !toRemove.has(r));
  removeFilterDropdown();
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
  triggerAutoSave();
}

// ============================================
// Feature: Cell Comments
// ============================================

let activeCommentTooltip = null;

function addComment(row, col) {
  const key = cellKey(row, col);
  const sheet = sheets[activeSheet];
  const cell = sheet.cells[key] || {};
  const existing = cell.comment || '';
  const text = prompt('Enter comment:', existing);
  if (text === null) return;
  if (!sheet.cells[key]) sheet.cells[key] = {};
  if (text === '') {
    delete sheet.cells[key].comment;
  } else {
    sheet.cells[key].comment = text;
  }
  renderSheet();
  selectCell(row, col);
  triggerAutoSave();
}

function deleteComment(row, col) {
  const key = cellKey(row, col);
  const sheet = sheets[activeSheet];
  if (sheet.cells[key]) {
    delete sheet.cells[key].comment;
    renderSheet();
    selectCell(row, col);
    triggerAutoSave();
  }
}

function showCommentTooltip(td, comment) {
  hideCommentTooltip();
  const tip = document.createElement('div');
  tip.className = 'comment-tooltip';
  tip.textContent = comment;
  document.body.appendChild(tip);
  const rect = td.getBoundingClientRect();
  tip.style.left = (rect.right + 5) + 'px';
  tip.style.top = rect.top + 'px';
  // Adjust if off-screen
  const tipRect = tip.getBoundingClientRect();
  if (tipRect.right > window.innerWidth) tip.style.left = (rect.left - tipRect.width - 5) + 'px';
  if (tipRect.bottom > window.innerHeight) tip.style.top = (window.innerHeight - tipRect.height - 5) + 'px';
  activeCommentTooltip = tip;
}

function hideCommentTooltip() {
  if (activeCommentTooltip) { activeCommentTooltip.remove(); activeCommentTooltip = null; }
}

// ============================================
// Feature: Print / Export PDF
// ============================================

function printSheet() {
  const fileName = document.getElementById('file-name').value || 'Spreadsheet';
  document.body.setAttribute('data-print-title', fileName);

  // Free users: add print watermark
  let watermark = null;
  if (!isPro()) {
    watermark = document.createElement('div');
    watermark.className = 'print-watermark';
    watermark.textContent = 'Created with Anomaly Quantix - anomaly-surround.github.io/anomaly-quantix';
    document.body.appendChild(watermark);
  }

  window.print();

  if (watermark) watermark.remove();
}

// ============================================
// Feature: Row/Column Hide
// ============================================

function hideRow(row) {
  const sheet = sheets[activeSheet];
  if (!sheet.hiddenRows) sheet.hiddenRows = [];
  if (!sheet.hiddenRows.includes(row)) sheet.hiddenRows.push(row);
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
  triggerAutoSave();
}

function unhideRow(row) {
  const sheet = sheets[activeSheet];
  if (!sheet.hiddenRows) return;
  // Unhide row and adjacent hidden rows
  const toUnhide = [row];
  // Also check row-1 and row+1
  for (let r = row - 1; r >= 0; r--) {
    if (sheet.hiddenRows.includes(r)) toUnhide.push(r); else break;
  }
  for (let r = row + 1; r < ROWS; r++) {
    if (sheet.hiddenRows.includes(r)) toUnhide.push(r); else break;
  }
  sheet.hiddenRows = sheet.hiddenRows.filter(r => !toUnhide.includes(r));
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
  triggerAutoSave();
}

function hideColumn(col) {
  const sheet = sheets[activeSheet];
  if (!sheet.hiddenCols) sheet.hiddenCols = [];
  if (!sheet.hiddenCols.includes(col)) sheet.hiddenCols.push(col);
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
  triggerAutoSave();
}

function unhideColumn(col) {
  const sheet = sheets[activeSheet];
  if (!sheet.hiddenCols) return;
  const toUnhide = [col];
  for (let c = col - 1; c >= 0; c--) {
    if (sheet.hiddenCols.includes(c)) toUnhide.push(c); else break;
  }
  for (let c = col + 1; c < COLS; c++) {
    if (sheet.hiddenCols.includes(c)) toUnhide.push(c); else break;
  }
  sheet.hiddenCols = sheet.hiddenCols.filter(c => !toUnhide.includes(c));
  renderSheet();
  selectCell(selectedCell.row, selectedCell.col);
  triggerAutoSave();
}

// ============================================
// Feature: Drag & Drop Rows
// ============================================

let rowDragSource = null;
let rowDragIndicator = null;

function startRowDrag(row, e) {
  e.preventDefault();
  rowDragSource = row;

  const onMove = (e2) => {
    const td = e2.target.closest('td[data-row]');
    if (!td) return;
    const targetRow = +td.dataset.row;
    // Remove old indicator
    document.querySelectorAll('.row-drop-indicator').forEach(tr => tr.classList.remove('row-drop-indicator'));
    document.querySelectorAll('.row-dragging').forEach(tr => tr.classList.remove('row-dragging'));
    // Mark source
    const srcTr = td.closest('tbody')?.querySelectorAll('tr');
    if (srcTr) {
      srcTr.forEach(tr => {
        const firstTd = tr.querySelector('td[data-row]');
        if (firstTd && +firstTd.dataset.row === rowDragSource) tr.classList.add('row-dragging');
        if (firstTd && +firstTd.dataset.row === targetRow) tr.classList.add('row-drop-indicator');
      });
    }
    rowDragIndicator = targetRow;
  };

  const onUp = () => {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    document.querySelectorAll('.row-drop-indicator, .row-dragging').forEach(tr => {
      tr.classList.remove('row-drop-indicator');
      tr.classList.remove('row-dragging');
    });
    if (rowDragSource !== null && rowDragIndicator !== null && rowDragSource !== rowDragIndicator) {
      moveRow(rowDragSource, rowDragIndicator);
    }
    rowDragSource = null;
    rowDragIndicator = null;
  };

  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

function moveRow(fromRow, toRow) {
  const sheet = sheets[activeSheet];
  // Collect all cells in the source row
  const srcCells = {};
  for (let c = 0; c < COLS; c++) {
    const key = cellKey(fromRow, c);
    if (sheet.cells[key]) { srcCells[c] = { ...sheet.cells[key] }; delete sheet.cells[key]; }
  }

  // Shift rows
  if (fromRow < toRow) {
    for (let r = fromRow; r < toRow; r++) {
      for (let c = 0; c < COLS; c++) {
        const below = cellKey(r + 1, c);
        const cur = cellKey(r, c);
        if (sheet.cells[below]) { sheet.cells[cur] = sheet.cells[below]; delete sheet.cells[below]; }
        else delete sheet.cells[cur];
      }
    }
    // Place source at toRow
    for (let c = 0; c < COLS; c++) {
      const key = cellKey(toRow, c);
      if (srcCells[c]) sheet.cells[key] = srcCells[c];
    }
  } else {
    for (let r = fromRow; r > toRow; r--) {
      for (let c = 0; c < COLS; c++) {
        const above = cellKey(r - 1, c);
        const cur = cellKey(r, c);
        if (sheet.cells[above]) { sheet.cells[cur] = sheet.cells[above]; delete sheet.cells[above]; }
        else delete sheet.cells[cur];
      }
    }
    for (let c = 0; c < COLS; c++) {
      const key = cellKey(toRow, c);
      if (srcCells[c]) sheet.cells[key] = srcCells[c];
    }
  }

  renderSheet();
  selectCell(toRow, selectedCell.col);
  triggerAutoSave();
}

// ============================================
// Feature: Block Drag (move selected range)
// ============================================

function isInSelectionRange(row, col) {
  if (!selectionRange) return false;
  const r1 = Math.min(selectionRange.startRow, selectionRange.endRow);
  const r2 = Math.max(selectionRange.startRow, selectionRange.endRow);
  const c1 = Math.min(selectionRange.startCol, selectionRange.endCol);
  const c2 = Math.max(selectionRange.startCol, selectionRange.endCol);
  return row >= r1 && row <= r2 && col >= c1 && col <= c2;
}

let blockDragData = null;

function startBlockDrag(r1, c1, r2, c2, grabRow, grabCol, e) {
  const offsetR = grabRow - r1;
  const offsetC = grabCol - c1;
  const blockH = r2 - r1;
  const blockW = c2 - c1;

  blockDragData = { r1, c1, r2, c2, offsetR, offsetC, blockH, blockW };
  document.getElementById('sheet-container').classList.add('block-dragging');

  const onMove = (e2) => {
    const td = e2.target.closest('td[data-row]');
    if (!td) return;
    const targetRow = +td.dataset.row;
    const targetCol = +td.dataset.col;

    // Calculate where the block top-left would land
    const newR1 = targetRow - offsetR;
    const newC1 = targetCol - offsetC;
    const newR2 = newR1 + blockH;
    const newC2 = newC1 + blockW;

    // Clear old drop indicators
    document.querySelectorAll('td.block-drop-target').forEach(td => td.classList.remove('block-drop-target'));

    // Show drop target
    for (let r = newR1; r <= newR2; r++) {
      for (let c = newC1; c <= newC2; c++) {
        const t = getCellTd(r, c);
        if (t) t.classList.add('block-drop-target');
      }
    }

    blockDragData.targetR1 = newR1;
    blockDragData.targetC1 = newC1;
  };

  const onUp = () => {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
    document.querySelectorAll('td.block-drop-target').forEach(td => td.classList.remove('block-drop-target'));
    document.getElementById('sheet-container').classList.remove('block-dragging');

    if (blockDragData && blockDragData.targetR1 !== undefined) {
      const { r1, c1, r2, c2, targetR1, targetC1 } = blockDragData;
      // Don't move if dropped in the same spot
      if (targetR1 !== r1 || targetC1 !== c1) {
        moveBlock(r1, c1, r2, c2, targetR1, targetC1);
      }
    }
    blockDragData = null;
  };

  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

function moveBlock(r1, c1, r2, c2, targetR1, targetC1) {
  const sheet = sheets[activeSheet];
  const blockH = r2 - r1;
  const blockW = c2 - c1;
  const targetR2 = targetR1 + blockH;
  const targetC2 = targetC1 + blockW;

  // Bounds check
  if (targetR1 < 0 || targetR2 >= ROWS || targetC1 < 0 || targetC2 >= COLS) return;

  // Collect source cells
  const srcCells = {};
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const key = cellKey(r, c);
      if (sheet.cells[key]) {
        srcCells[(r - r1) + '_' + (c - c1)] = { ...sheet.cells[key] };
      }
      delete sheet.cells[key];
    }
  }

  // Place at target
  for (let r = 0; r <= blockH; r++) {
    for (let c = 0; c <= blockW; c++) {
      const src = srcCells[r + '_' + c];
      const key = cellKey(targetR1 + r, targetC1 + c);
      if (src) {
        sheet.cells[key] = src;
      } else {
        delete sheet.cells[key];
      }
    }
  }

  // Update selection to new position
  selectionRange = {
    startRow: targetR1, startCol: targetC1,
    endRow: targetR2, endCol: targetC2
  };

  renderSheet();
  selectedCell = { row: targetR1, col: targetC1 };
  document.getElementById('cell-ref').textContent = COL_LETTERS[targetC1] + (targetR1 + 1);
  highlightRange();
  triggerAutoSave();
  document.getElementById('status-info').textContent = 'Block moved';
}

// ============================================
// Feature: Dark/Light Theme Toggle
// ============================================

function toggleTheme() {
  const html = document.documentElement;
  const current = html.getAttribute('data-theme');
  const newTheme = current === 'light' ? 'dark' : 'light';
  html.setAttribute('data-theme', newTheme);
  localStorage.setItem('quantix-theme', newTheme);
  updateThemeButton(newTheme);
}

function updateThemeButton(theme) {
  const btn = document.getElementById('btn-theme');
  if (!btn) return;
  if (theme === 'light') {
    btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.79A9 9 0 1111.21 3a7 7 0 009.79 9.79z"/></svg>';
    btn.title = 'Switch to Dark Theme';
  } else {
    btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>';
    btn.title = 'Switch to Light Theme';
  }
}

function loadTheme() {
  const saved = localStorage.getItem('quantix-theme');
  if (saved) {
    document.documentElement.setAttribute('data-theme', saved);
    updateThemeButton(saved);
  }
}

// ============================================
// Feature: Auto-fill Suggestions (Autocomplete)
// ============================================

let autocompleteDropdown = null;
let autocompleteItems = [];
let autocompleteSelectedIdx = -1;

function showAutocompleteSuggestions(input, row, col) {
  hideAutocompleteSuggestions();
  const val = input.value;
  if (!val || val.startsWith('=') || val.length < 2) return;

  const sheet = sheets[activeSheet];
  const lower = val.toLowerCase();
  const matches = new Set();

  // Check same column for matching values
  for (let r = 0; r < ROWS; r++) {
    if (r === row) continue;
    const cell = sheet.cells[cellKey(r, col)];
    if (cell && cell.value !== undefined && cell.value !== '') {
      const sv = String(cell.value);
      if (sv.toLowerCase().startsWith(lower) && sv.toLowerCase() !== lower) {
        matches.add(sv);
        if (matches.size >= 5) break;
      }
    }
  }

  if (matches.size === 0) return;

  autocompleteItems = [...matches];
  autocompleteSelectedIdx = -1;

  const dropdown = document.createElement('div');
  dropdown.className = 'autocomplete-dropdown';
  autocompleteDropdown = dropdown;

  autocompleteItems.forEach((item, idx) => {
    const btn = document.createElement('button');
    btn.className = 'ac-item';
    btn.textContent = item;
    btn.addEventListener('mousedown', (e) => {
      e.preventDefault();
      input.value = item;
      hideAutocompleteSuggestions();
    });
    dropdown.appendChild(btn);
  });

  document.body.appendChild(dropdown);

  // Position below the cell
  const td = getCellTd(row, col);
  if (td) {
    const rect = td.getBoundingClientRect();
    dropdown.style.left = rect.left + 'px';
    dropdown.style.top = rect.bottom + 'px';
    dropdown.style.minWidth = rect.width + 'px';
  }
}

function hideAutocompleteSuggestions() {
  if (autocompleteDropdown) { autocompleteDropdown.remove(); autocompleteDropdown = null; }
  autocompleteItems = [];
  autocompleteSelectedIdx = -1;
}

function handleAutocompleteKey(e, input) {
  if (!autocompleteDropdown) return false;

  if (e.key === 'ArrowDown') {
    e.preventDefault();
    autocompleteSelectedIdx = Math.min(autocompleteSelectedIdx + 1, autocompleteItems.length - 1);
    updateAutocompleteHighlight();
    return true;
  }
  if (e.key === 'ArrowUp') {
    e.preventDefault();
    autocompleteSelectedIdx = Math.max(autocompleteSelectedIdx - 1, -1);
    updateAutocompleteHighlight();
    return true;
  }
  if ((e.key === 'Tab' || e.key === 'Enter') && autocompleteSelectedIdx >= 0) {
    e.preventDefault();
    input.value = autocompleteItems[autocompleteSelectedIdx];
    hideAutocompleteSuggestions();
    return true;
  }
  if (e.key === 'Escape') {
    hideAutocompleteSuggestions();
    return true;
  }
  return false;
}

function updateAutocompleteHighlight() {
  if (!autocompleteDropdown) return;
  autocompleteDropdown.querySelectorAll('.ac-item').forEach((item, idx) => {
    item.classList.toggle('ac-selected', idx === autocompleteSelectedIdx);
  });
}

// ============================================
// Init
// ============================================

loadTheme();
init();

// PWA & File Handling
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js');
}

if ('launchQueue' in window) {
  window.launchQueue.setConsumer(async (launchParams) => {
    if (!launchParams.files || !launchParams.files.length) return;
    const fileHandle = launchParams.files[0];
    const file = await fileHandle.getFile();
    const ext = file.name.split('.').pop().toLowerCase();

    if (ext === 'xlsx' || ext === 'xls') {
      const buf = await file.arrayBuffer();
      importExcel(buf);
    } else if (ext === 'csv') {
      const text = await file.text();
      importCSV(text);
    } else {
      const text = await file.text();
      try {
        const data = JSON.parse(text);
        if (!data.sheets || !Array.isArray(data.sheets)) throw new Error('Invalid file');
        sheets = data.sheets.map(s => sanitizeSheetData(s));
        activeSheet = Math.max(0, Math.min(toNum(data.activeSheet || 0), sheets.length - 1));
        document.getElementById('file-name').value = String(data.fileName || file.name.replace('.qx', ''));
        renderSheetTabs();
        renderSheet();
        selectCell(0, 0);
      } catch (err) {
        alert('Failed to load file: ' + err.message);
      }
    }
    document.getElementById('file-name').value = file.name.replace(/\.\w+$/, '');
    document.getElementById('status-info').textContent = 'Loaded: ' + file.name;
  });
}

// ============================================
// Pro / Freemium
// ============================================

// Pro / Freemium Config (constants defined at top of file)

function isPro() {
  const data = localStorage.getItem('quantix-pro');
  if (!data) return false;
  try {
    const pro = JSON.parse(data);
    return pro && pro.pro === true;
  } catch { return false; }
}

function getProData() {
  try { return JSON.parse(localStorage.getItem('quantix-pro')) || null; } catch { return null; }
}

function getAuthToken() {
  return localStorage.getItem('quantix-token');
}

function updateProUI() {
  const btn = document.getElementById('pro-btn');
  const label = document.getElementById('pro-btn-label');
  const data = getProData();
  if (isPro()) {
    btn.classList.add('is-pro');
    label.textContent = 'Pro';
  } else if (data && data.email) {
    btn.classList.remove('is-pro');
    label.textContent = 'Upgrade to Pro';
  } else {
    btn.classList.remove('is-pro');
    label.textContent = 'Upgrade to Pro';
  }
}

function showProModal() {
  const modal = document.getElementById('pro-modal');
  const statusSection = document.getElementById('pro-status-section');
  const upgradeSection = document.getElementById('pro-upgrade-section');
  const data = getProData();

  if (isPro() && data) {
    statusSection.innerHTML = `
      <div class="pro-status">
        <h3>Pro Activated</h3>
        <p>${escapeHTML(data.email)}</p>
        <p style="margin-top:8px">Thank you for supporting Anomaly Quantix!</p>
      </div>
      <button class="btn-primary" onclick="signOut()" style="background:var(--danger);margin-right:8px">Sign Out</button>
      <button class="btn-primary" onclick="deactivatePro()" style="background:var(--border)">Deactivate License</button>
    `;
    upgradeSection.style.display = 'none';
  } else if (data && data.email && !data.pro) {
    // Signed in but not Pro
    statusSection.innerHTML = `
      <div class="pro-status" style="border-color:var(--border)">
        <h3>Signed in as</h3>
        <p>${escapeHTML(data.email)}</p>
        <button class="btn-primary" onclick="signOut()" style="background:var(--danger);margin-top:8px;font-size:11px;padding:5px 12px">Sign Out</button>
      </div>
    `;
    upgradeSection.style.display = '';
    document.getElementById('google-signin-btn').style.display = 'none';
    document.getElementById('google-signin-divider').style.display = 'none';
    document.getElementById('license-error').textContent = '';
  } else {
    statusSection.innerHTML = '';
    upgradeSection.style.display = '';
    document.getElementById('google-signin-btn').style.display = '';
    document.getElementById('google-signin-divider').style.display = '';
    document.getElementById('license-error').textContent = '';
  }

  modal.style.display = 'flex';
}

function signInWithGoogle() {
  window.location.href = WORKER_URL + '/auth/google';
}

function signOut() {
  localStorage.removeItem('quantix-pro');
  localStorage.removeItem('quantix-token');
  updateProUI();
  document.getElementById('pro-modal').style.display = 'none';
  document.getElementById('status-info').textContent = 'Signed out.';
}

function handleCheckout() {
  window.open(PRO_CHECKOUT_URL, '_blank');
}

function activateFromConfirm() {
  const input = document.getElementById('confirm-license-input');
  const error = document.getElementById('confirm-license-error');
  const key = input.value.trim();

  if (!key) { error.textContent = 'Please enter a license key.'; return; }

  error.textContent = 'Activating...';
  error.style.color = 'var(--text-dim)';

  tryActivateKey(key, error, () => {
    document.getElementById('confirm-modal').style.display = 'none';
  });
}

function activateLicense() {
  const input = document.getElementById('license-key-input');
  const error = document.getElementById('license-error');
  const key = input.value.trim();

  if (!key) { error.textContent = 'Please enter a license key.'; return; }

  error.textContent = 'Activating...';
  error.style.color = 'var(--text-dim)';

  tryActivateKey(key, error, () => {
    document.getElementById('pro-modal').style.display = 'none';
  });
}

function tryActivateKey(key, errorEl, onSuccess) {
  if (key.length < 6 || !/^[A-Za-z0-9\-_]+$/.test(key)) {
    errorEl.style.color = 'var(--danger)';
    errorEl.textContent = 'Invalid license key.';
    return;
  }

  const token = getAuthToken();
  if (!token) {
    errorEl.style.color = 'var(--danger)';
    errorEl.textContent = 'Please sign in with Google first.';
    return;
  }

  fetch(WORKER_URL + '/api/activate', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
    body: JSON.stringify({ key })
  })
  .then(r => r.json())
  .then(data => {
    if (data.success) {
      const proData = getProData() || {};
      proData.pro = true;
      proData.key = key;
      localStorage.setItem('quantix-pro', JSON.stringify(proData));
      updateProUI();
      document.getElementById('status-info').textContent = 'Pro activated! Thank you!';
      if (onSuccess) onSuccess();
    } else {
      errorEl.style.color = 'var(--danger)';
      errorEl.textContent = data.error || 'Activation failed.';
    }
  })
  .catch(() => {
    errorEl.style.color = 'var(--danger)';
    errorEl.textContent = 'Could not connect to server. Try again later.';
  });
}

function deactivatePro() {
  if (!confirm('Are you sure you want to deactivate your Pro license?')) return;

  const token = getAuthToken();
  if (token) {
    fetch(WORKER_URL + '/api/deactivate', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token }
    }).catch(() => {});
  }

  localStorage.removeItem('quantix-pro');
  updateProUI();
  document.getElementById('pro-modal').style.display = 'none';
  document.getElementById('status-info').textContent = 'Pro deactivated.';
}

// Handle OAuth callback params on page load
function handleAuthCallback() {
  const params = new URLSearchParams(window.location.search);
  const token = params.get('token');
  const error = params.get('error');

  if (!token && !error) return;

  // Clean URL
  window.history.replaceState({}, '', window.location.pathname);

  if (error) {
    document.getElementById('status-info').textContent = 'Login failed: ' + error;
    return;
  }

  if (token) {
    localStorage.setItem('quantix-token', token);

    // Fetch user info and pro status
    fetch(WORKER_URL + '/api/status', {
      headers: { 'Authorization': 'Bearer ' + token }
    })
    .then(r => r.json())
    .then(data => {
      if (data.email) {
        localStorage.setItem('quantix-pro', JSON.stringify({
          email: data.email,
          name: data.name,
          picture: data.picture,
          pro: data.pro,
          key: data.licenseKey
        }));
        updateProUI();
        if (data.pro) {
          document.getElementById('status-info').textContent = 'Welcome back, ' + data.name + '! Pro is active.';
        } else {
          document.getElementById('status-info').textContent = 'Signed in as ' + data.email;
        }
      }
    })
    .catch(() => {
      document.getElementById('status-info').textContent = 'Signed in (offline mode).';
    });
  }

}

// Check pro status on load if token exists
function checkProStatus() {
  const token = getAuthToken();
  if (!token) return;

  fetch(WORKER_URL + '/api/status', {
    headers: { 'Authorization': 'Bearer ' + token }
  })
  .then(r => r.json())
  .then(data => {
    if (data.email) {
      localStorage.setItem('quantix-pro', JSON.stringify({
        email: data.email,
        name: data.name,
        picture: data.picture,
        pro: data.pro,
        key: data.licenseKey
      }));
      updateProUI();
    }
  })
  .catch(() => {}); // Silently fail if offline
}

// Feature gating
const FREE_MAX_SHEETS = 3;

// Override addSheet to enforce limit
const _originalAddSheet = typeof addSheet === 'function' ? addSheet : null;

