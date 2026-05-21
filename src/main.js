import { mountLoginPage } from './login-ui.js';
import { getSupabase } from './supabase-client.js';
import { getJsPDF } from './pdf-libs.js';
import { renderMillProfileSummaryPdf } from './mill-profile-pdf-summary.js';

// ─── GLOBAL NAVIGATION NOTE: switchPanel is defined later in the file. ────
  let supplierWorkbook = null;
  window._sddUserRole = window._sddUserRole || 'STAFF';
  window._sddApproverRecordLoaded = false;
  window._sddIsLoadedSaved = false;
  window._sddLastInsertedRow = null;   // legacy — kept for compat; prefer _sddSubmissionId
  window._sddSubmissionId = null;      // relational primary key (submission_id from SDD_MAIN)
  window._sddSupplierType = window._sddSupplierType || '';

  /** Grievance / PRI tables: max rows in UI and columns saved to SDD_MAIN (GRV1… / PRI1…). */
  const SCR_GRV_PRI_MAX_ROWS = 10;

  function supplierHandleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    const supplierType = getCurrentSddSupplierType();
    if (!supplierType) {
      event.target.value = '';
      if (typeof window.showSddToast === 'function') {
        window.showSddToast('Pilih supplier type dulu (MILL / KCP / TRADER) sebelum import Excel.', 'error');
      }
      return;
    }
    window._sddIsLoadedSaved = false;
    window._sddLastInsertedRow = null;
    window._scrLoadedRowNum = null;
    window._scrLoadedKey = '';
    window._sddSubmissionId = null;
    window._loadedPrimarySddRow = null;
    // Anti-leak: import baru tidak boleh membawa data screening dari supplier sebelumnya.
    window._tmlScreeningData = {};
    window._ffbScreeningData = {};

    const reader = new FileReader();
    reader.onload = function(e) {
      const data = new Uint8Array(e.target.result);
      supplierWorkbook = XLSX.read(data, { type: 'array', cellDates: true });

      const sheetSelect = document.getElementById('supplierSheetSelect');
      sheetSelect.innerHTML = '';
      supplierWorkbook.SheetNames.forEach(name => {
        const option = document.createElement('option');
        option.value = name;
        option.textContent = name;
        sheetSelect.appendChild(option);
      });

      const preferSdd = supplierWorkbook.SheetNames.findIndex(function(n) {
        return String(n).trim().toLowerCase() === 'sdd data';
      });
      if (preferSdd >= 0) sheetSelect.selectedIndex = preferSdd;

      document.getElementById('supplierSheetSelectContainer').style.display = 'block';
      supplierDisplayAllSheets();
    };
    reader.readAsArrayBuffer(file);
  }

  function supplierDisplaySelectedSheet() {
    if (!supplierWorkbook) return;
    const sheetName = document.getElementById('supplierSheetSelect')?.value;
    if (!sheetName) return;
    const worksheet = supplierWorkbook.Sheets[sheetName];
    if (!worksheet) return;
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'dd mmm yyyy' });
    const container = document.getElementById('supplierExcelData');
    container.innerHTML = '';
    supplierDisplayExcelDataTo(container, jsonData);
    cacheSddImportRowsFromWorkbook();
  }

  /** Inline HTML (onchange="…") hanya melihat `window` — ES module tidak mengekspor fungsi ke global. */
  window.supplierHandleFileUpload = supplierHandleFileUpload;
  window.supplierDisplaySelectedSheet = supplierDisplaySelectedSheet;

  function normalizeCellText(v) {
    return String(v === undefined || v === null ? '' : v).replace(/\s+/g, ' ').trim();
  }

  function normalizeLooseKey(v) {
    return normalizeCellText(v).toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  /**
   * Normalize a coordinate value to use comma as decimal separator (Indonesian format).
   *
   * Root-cause guard:
   *   XLSX.read with raw:false converts numeric cells to their *display string*.
   *   A cell formatted as "0,318318" (Indonesian locale) may be stored internally
   *   as a Number (0.318318) and serialised by SheetJS as "0.318318" (dot) — or,
   *   if the sheet uses a text cell with a comma, it arrives correctly as "0,318318".
   *   Either way this helper ensures we always emit the comma form before saving.
   *
   * Rules:
   *   - Empty / null / undefined → ''
   *   - Already has comma AND no dot  → return as-is (already correct)
   *   - Has dot as decimal separator   → replace dot with comma
   *   - Integers (no separator)        → return as-is (no corruption possible)
   *   - Minus sign is preserved
   *
   * @param {*} v - raw cell value (string or number)
   * @returns {string} coordinate in dot-decimal format, or ''
   *
   * PATCH (coordinate round-trip fix): always store/send with DOT as decimal
   * separator. Google Sheets locale id_ID interprets comma as thousands separator,
   * so sending "0,318318" causes Sheets to store integer 318318 (data loss).
   * Display-only conversion to comma happens at render time, not here.
   */
  function normalizeCoordinate(v, fieldName) {
    if (v === undefined || v === null) return '';
    const s = String(v).trim();
    if (!s) return '';
    // Comma-decimal (Indonesian "0,318318") → convert to dot-decimal "0.318318".
    if (s.indexOf(',') !== -1 && s.indexOf('.') === -1) return s.replace(',', '.');
    // String contains a dot — could be (a) genuine decimal "101.38" or
    // (b) Indonesian thousands separator "318.318" (dot-thousands, no comma-decimal).
    // Guard: if the parsed float value is out of valid WGS-84 range it CANNOT be a
    // valid coordinate, so treat the dot as a thousands separator, strip it, then
    // fall through to the integer-recovery heuristic below.
    if (s.indexOf('.') !== -1) {
      const parsed = parseFloat(s);
      if (!isNaN(parsed) && Math.abs(parsed) <= 180) return s; // Looks like valid decimal
      // Out-of-range: strip all dots (thousands separators) and re-parse as integer
      const stripped = s.replace(/\./g, '');
      if (/^-?\d+$/.test(stripped)) {
        // Fall through to integer heuristic with the stripped value
        return normalizeCoordinate(stripped, fieldName);
      }
      return s; // Cannot recover — return as-is and let backend reject
    }
    // Pure integer: if within valid WGS-84 range, return as-is (e.g. "0", "-7")
    if (/^-?\d+$/.test(s)) {
      const absN = Math.abs(parseInt(s, 10));
      if (absN <= 180) return s;
      // Corrupted integer (e.g. 318318 from "0.318318", 1013838183 from "101.3838183"):
      // enumerate all decimal-point insertion positions and pick most plausible.
      const isLng = /long|lng/i.test(fieldName || '');
      const maxInt = isLng ? 180 : 90;
      const neg = s.charAt(0) === '-';
      const pfx = neg ? '-' : '';
      const digits = String(absN);
      const candidates = [];
      for (let k = 1; k <= digits.length; k++) {
        const pow = Math.pow(10, k);
        const leftInt = Math.floor(absN / pow);
        if (leftInt > maxInt) continue;
        const rightStr = String(absN % pow).padStart(k, '0');
        candidates.push({ leftInt: leftInt, k: k, valStr: pfx + leftInt + '.' + rightStr });
      }
      if (candidates.length) {
        const best = isLng
          ? candidates.reduce(function(a, b) { return b.leftInt > a.leftInt ? b : a; })
          : candidates.reduce(function(a, b) { return b.k > a.k ? b : a; });
        return best.valStr;
      }
    }
    return s;
  }

  function normalizeSddSupplierType(v) {
    const s = String(v || '').trim().toUpperCase();
    if (s === 'MILL' || s === 'MIL') return 'MILL';
    if (s === 'KCP') return 'KCP';
    if (s === 'TRADER' || s === 'TRD') return 'TRADER';
    return '';
  }

  function getCurrentSddSupplierType() {
    return normalizeSddSupplierType(window._sddSupplierType);
  }

  function debounce(fn, wait) {
    let timer = null;
    function wrapped() {
      const ctx = this;
      const args = arguments;
      if (timer) clearTimeout(timer);
      timer = setTimeout(function() {
        timer = null;
        fn.apply(ctx, args);
      }, wait);
    }
    wrapped.cancel = function() {
      if (timer) {
        clearTimeout(timer);
        timer = null;
      }
    };
    return wrapped;
  }

  function makeRafScheduler(fn) {
    let rafId = 0;
    function scheduled() {
      if (rafId) cancelAnimationFrame(rafId);
      rafId = requestAnimationFrame(function() {
        rafId = 0;
        fn();
      });
    }
    scheduled.flush = function() {
      if (rafId) {
        cancelAnimationFrame(rafId);
        rafId = 0;
      }
      fn();
    };
    return scheduled;
  }

  function syncSddSupplierTypeSelectorUI() {
    const sel = document.getElementById('sdd-supplier-type-select');
    if (!sel) return;
    const cur = getCurrentSddSupplierType();
    sel.value = cur || '';
    const fileInput = document.getElementById('supplierExcelFile');
    if (fileInput) fileInput.disabled = !cur;
  }

  function hasMeaningfulImportData(row) {
    if (!row || typeof row !== 'object') return false;
    const entries = Object.entries(row).filter(function(pair) {
      const key = pair[0];
      const val = pair[1];
      if (/^SCR - |^GRV\d+|^PRI\d+/.test(key)) return false;
      return val !== undefined && val !== null && String(val).trim() !== '';
    });
    if (!entries.length) return false;
    const priorityKeys = [
      'Mill ID', 'Mil ID', 'Group Name', 'Grup Name', 'Company Name',
      'Mill Name', 'Current Owner', 'Office Address', 'Latitude', 'Longitude'
    ];
    if (priorityKeys.some(function(k) { return row[k] && String(row[k]).trim() !== ''; })) return true;
    return entries.length >= 4;
  }

  function importRowSignature(row) {
    if (!row || typeof row !== 'object') return '';
    return Object.keys(row)
      .filter(function(k) { return !/^SCR - |^GRV\d+|^PRI\d+/.test(k); })
      .sort()
      .map(function(k) { return k + ':' + String(row[k] == null ? '' : row[k]).trim(); })
      .join('|');
  }

  /**
   * Mill List / FFB-only payloads. On first save we INSERT these rows; the Main Form row is saved once via merged payload (avoids duplicate primary rows).
   */
  function rowIsTraceabilityOnlyPayload(row) {
    if (!row || typeof row !== 'object') return false;
    const hasTml = String(row['TML - Mill Name'] || row['TML - Company Name'] || '').trim() !== '';
    const hasFfb = String(row['FFB - Supplier Name'] || '').trim() !== '' ||
      (String(row['FFB - Mill Name'] || '').trim() !== '' &&
        String(row['FFB - Supplier Category'] || row['FFB - Village'] || row['FFB - Sub District'] || '').trim() !== '');
    return hasTml || hasFfb;
  }

  /** Copy Mill ID / import identity from the Main Form row onto traceability rows so they share one submission group key in Sheets. */
  function stampSddIdentityOntoRow(row, stamp) {
    if (!row || !stamp || typeof row !== 'object' || typeof stamp !== 'object') return;
    const keys = ['Supplier Type', 'SUPPLIER_TYPE', 'Mill ID', 'Mil ID', 'Group Name', 'Grup Name', 'Company Name', 'Mill Name', 'Date Imported', 'Imported By'];
    keys.forEach(function(k) {
      if (!String(row[k] || '').trim() && stamp[k] != null && String(stamp[k]).trim() !== '') {
        row[k] = stamp[k];
      }
    });
  }

  /** Saved-screening list: hide "ghost" rows where only Mill ID / import meta survived a partial delete. */
  function sddSavedRowHasListableSubstance(r) {
    if (!r || typeof r !== 'object') return false;
    const skip = {
      _row: 1,
      'Mill ID': 1,
      'Mil ID': 1,
      'MILL ID': 1,
      'Date Imported': 1,
      'Imported By': 1,
    };
    for (const k in r) {
      if (!Object.prototype.hasOwnProperty.call(r, k)) continue;
      if (skip[k]) continue;
      const v = r[k];
      if (v === undefined || v === null || String(v).trim() === '') continue;
      return true;
    }
    return false;
  }

  /** All sheet row numbers for the currently loaded submission (main + traceability rows), descending for safe server-side deletes. */
  function collectSddSheetRowNumsForCurrentLoadedSubmission() {
    const nums = [];
    function add(n) {
      const x = parseInt(n, 10);
      if (!isNaN(x) && x > 0) nums.push(x);
    }
    const key = window._scrLoadedKey;
    const primary = window._loadedPrimarySddRow || (key && window._scrSavedRowsByKey && window._scrSavedRowsByKey[key]);
    const cache = window._sddAllRowsCache;
    if (key && primary && Array.isArray(cache)) {
      rowsMatchingSavedSubmissionKey(cache, key, primary).forEach(function(row) {
        if (row && row._row != null) add(row._row);
      });
    }
    if (!nums.length && key && window._scrSavedGroupsByKey && window._scrSavedGroupsByKey[key]) {
      const gr = window._scrSavedGroupsByKey[key].rows;
      if (Array.isArray(gr)) {
        gr.forEach(function(row) {
          if (row && row._row != null) add(row._row);
        });
      }
    }
    if (!nums.length && window._scrLoadedRowNum != null) add(window._scrLoadedRowNum);
    const seen = {};
    return nums.filter(function(n) {
      if (seen[n]) return false;
      seen[n] = true;
      return true;
    }).sort(function(a, b) {
      return b - a;
    });
  }

  function parseSheetAsKeyValueObject(rows) {
    const out = {};
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i] || [];
      const k = normalizeCellText(r[0]);
      const v = normalizeCellText(r[1]);
      if (!k || !v) continue;
      // only treat as key-value row when remaining columns are mostly empty
      const tailHasData = (r || []).slice(2).some(function(c) { return normalizeCellText(c) !== ''; });
      if (tailHasData) continue;
      const key = k.replace(/[:：]\s*$/, '');
      out[key] = v;
    }
    return out;
  }

  /** Parse imported worksheet into row objects keyed by SDD headers (header row can appear below row 1). */
  function parseSheetToSddRecords(worksheet) {
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'dd mmm yyyy' });
    if (!rows.length) return [];

    let headerRowIdx = -1;
    let headers = [];
    const scanLimit = Math.min(rows.length, 40);
    for (let i = 0; i < scanLimit; i++) {
      const candidate = (rows[i] || []).map(h => String(h || '').trim());
      if (!candidate.length) continue;
      const hasMillId = candidate.includes('Mill ID') || candidate.includes('Mil ID');
      const hasCompany = candidate.includes('Company Name');
      const hasGroup = candidate.includes('Group Name') || candidate.includes('Grup Name');
      if (hasMillId || (hasCompany && hasGroup)) {
        headerRowIdx = i;
        headers = candidate;
        break;
      }
    }
    if (headerRowIdx < 0) {
      headers = (rows[0] || []).map(h => String(h || '').trim());
      headerRowIdx = 0;
    }

    const tableRecords = rows.slice(headerRowIdx + 1).map(r => {
      const obj = {};
      headers.forEach((h, index) => {
        if (h) obj[h] = r[index] !== undefined ? r[index] : '';
      });
      return obj;
    }).filter(hasMeaningfulImportData);

    if (tableRecords.length) return tableRecords;

    // Fallback for non-tabular forms (key-value style sheet)
    const kvObj = parseSheetAsKeyValueObject(rows);
    if (hasMeaningfulImportData(kvObj)) return [kvObj];
    return [];
  }

  function cellAt(rows, r, c) {
    if (!rows[r]) return '';
    return normalizeCellText(rows[r][c]);
  }

  function findRowByFirstCell(rows, label, startIdx) {
    const from = startIdx || 0;
    const target = normalizeCellText(label).toLowerCase();
    for (let i = from; i < rows.length; i++) {
      if (cellAt(rows, i, 0).toLowerCase() === target) return i;
    }
    return -1;
  }

  function findRowByFirstCellAny(rows, labels, startIdx) {
    const from = startIdx || 0;
    const targets = (Array.isArray(labels) ? labels : [labels]).map(normalizeLooseKey).filter(Boolean);
    if (!targets.length) return -1;
    for (let i = from; i < rows.length; i++) {
      const key = normalizeLooseKey(cellAt(rows, i, 0));
      if (!key) continue;
      if (targets.some(function(t) { return key === t || key.indexOf(t) !== -1 || t.indexOf(key) !== -1; })) return i;
    }
    return -1;
  }

  function getValueNearLabel(rows, label) {
    const target = normalizeCellText(label).toLowerCase();
    function isPlaceholder(v) {
      const s = normalizeCellText(v).toLowerCase();
      return !s || s === 'please select' || s === '-' || s === '—';
    }
    const scanRows = Math.min(rows.length, 200);
    for (let r = 0; r < scanRows; r++) {
      const cols = rows[r] || [];
      for (let c = 0; c < Math.min(cols.length, 12); c++) {
        if (normalizeCellText(cols[c]).toLowerCase() !== target) continue;
        const near = [
          normalizeCellText(cols[c + 2]),
          normalizeCellText(cols[c + 1]),
          normalizeCellText(cols[c + 3]),
          normalizeCellText(cols[c + 4]),
          normalizeCellText(cols[c + 5]),
          normalizeCellText(cols[c + 6]),
        ];
        // Prefer first non-placeholder value (e.g. "150 KM" over "Please Select")
        const preferred = near.find(function(v) { return !isPlaceholder(v); });
        if (preferred) return preferred;
        // Fallback to first available value
        const any = near.find(function(v) { return !!v; });
        return any || '';
      }
    }
    return '';
  }

  /**
   * Like getValueNearLabel but reads the worksheet with raw:true so that
   * numeric cells return their actual Number value instead of a display string.
   * Use ONLY for coordinate fields where the cell may be formatted as an
   * integer in Excel (losing decimal places when raw:false is used).
   *
   * @param {Object} worksheet - raw SheetJS worksheet object
   * @param {string} label     - label to search for
   * @returns {string} raw numeric value as dot-decimal string, or ''
   */
  function getValueNearLabelRaw_(worksheet, label) {
    if (!worksheet) return '';
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
    const target  = label.trim().toLowerCase();
    function isPlaceholder(v) {
      const s = String(v === undefined || v === null ? '' : v).trim().toLowerCase();
      return !s || s === 'please select' || s === '-' || s === '—';
    }
    const scanRows = Math.min(rawRows.length, 200);
    for (let r = 0; r < scanRows; r++) {
      const cols = rawRows[r] || [];
      for (let c = 0; c < Math.min(cols.length, 12); c++) {
        const cell = String(cols[c] === undefined || cols[c] === null ? '' : cols[c]).trim();
        if (cell.toLowerCase() !== target) continue;
        const near = [cols[c + 2], cols[c + 1], cols[c + 3], cols[c + 4], cols[c + 5], cols[c + 6]];
        const preferred = near.find(v => v !== undefined && v !== null && !isPlaceholder(v));
        if (preferred !== undefined && preferred !== null) {
          // If it's a Number (raw:true), String() always gives dot-decimal notation
          return String(preferred).trim();
        }
        const any = near.find(v => v !== undefined && v !== null && String(v).trim() !== '');
        return any !== undefined ? String(any).trim() : '';
      }
    }
    return '';
  }

  function getDistanceFields(rows, label) {
    const target = normalizeCellText(label).toLowerCase();
    const out = { refinery: '', distanceKm: '' };
    function toKmValue(v) {
      const s = normalizeCellText(v);
      if (!s) return '';
      if (/^\s*\d+(\.\d+)?\s*km\s*$/i.test(s)) return s.replace(/\s+/g, ' ').trim().toUpperCase().replace('KM', ' KM');
      // Accept plain numeric cells (e.g. "150") as distance in KM.
      if (/^\d+([.,]\d+)?$/.test(s)) return s.replace(',', '.') + ' KM';
      return '';
    }
    const scanRows = Math.min(rows.length, 220);
    for (let r = 0; r < scanRows; r++) {
      const cols = rows[r] || [];
      for (let c = 0; c < Math.min(cols.length, 12); c++) {
        if (normalizeCellText(cols[c]).toLowerCase() !== target) continue;
        // Keep the actual selector value (can be "Please Select")
        const primary = normalizeCellText(cols[c + 2]) || normalizeCellText(cols[c + 1]) || normalizeCellText(cols[c + 3]) || '';
        // KM may sit immediately after the label row (c+1..c+3) or further right (legacy templates used c+4+)
        const right = [];
        for (let j = 1; j <= 12; j++) right.push(normalizeCellText(cols[c + j]));
        let kmVal = right.map(toKmValue).find(Boolean) || '';
        if (!kmVal) {
          const fb = getValueNearLabel(rows, 'Distance');
          kmVal = toKmValue(fb) || '';
        }
        out.refinery = primary;
        out.distanceKm = kmVal;
        return out;
      }
    }
    return out;
  }

  function extractMainFormToSddRecord(worksheet) {
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'dd mmm yyyy' });
    if (!rows.length) return null;
    const title = cellAt(rows, 0, 0).toLowerCase();

    // Accept Mill, KCP, and Trader supplier registration forms
    const isMill = title.indexOf('mill supplier registration form') !== -1;
    const isKcp  = title.indexOf('kcp supplier registration form') !== -1 ||
                   title.indexOf('kcp') !== -1 && title.indexOf('supplier') !== -1;
    const isTrader = title.indexOf('trader supplier registration form') !== -1 ||
                     title.indexOf('trader') !== -1 && title.indexOf('supplier') !== -1;
    if (!isMill && !isKcp && !isTrader) return null;

    // Company profile fields — same in both form types
    const record = {
      'Group Name': getValueNearLabel(rows, 'Group Name'),
      'Company Name': getValueNearLabel(rows, 'Company Name'),
      'Current Owner': getValueNearLabel(rows, 'Current Owner'),
      'Previous Owner': getValueNearLabel(rows, 'Previous Owner'),
      'Take Over From': getValueNearLabel(rows, 'Take Over From'),
      'Office Address': getValueNearLabel(rows, 'Office Address'),
      'Office Phone': getValueNearLabel(rows, 'Office Phone'),
      'Office Fax': getValueNearLabel(rows, 'Office Fax'),
      'Office Email': getValueNearLabel(rows, 'Office Email'),
      'Company Website': getValueNearLabel(rows, 'Company Website'),
      'Contact Person': getValueNearLabel(rows, 'Contact Person'),
      'Contact Position': getValueNearLabel(rows, 'Position'),
      'Contact Mobile': getValueNearLabel(rows, 'Mobile Phone'),
      'Contact Email': getValueNearLabel(rows, 'Contact Email'),
      'Sustainability PIC': getValueNearLabel(rows, 'Sustainability PIC'),
      'NDPE Commitment': getValueNearLabel(rows, 'NDPE Commitment'),
      'HRDD Commitment': getValueNearLabel(rows, 'HRDD Commitment'),
    };

    if (isKcp && !isMill) {
      // ── KCP-specific field mapping ──
      // KCP Name → Mill Name (same concept, different label)
      record['Mill Name'] = getValueNearLabel(rows, 'KCP Name') || getValueNearLabel(rows, 'Mill Name');
      record['Mill Address'] = getValueNearLabel(rows, 'KCP Address') || getValueNearLabel(rows, 'Mill Address');
      const kcpDistance = getDistanceFields(rows, 'Distance to Refinery');
      record['Distance to Refinery (km)'] = kcpDistance.refinery;
      record['Distance'] = kcpDistance.distanceKm;
      record['Latitude'] = normalizeCoordinate(
                           getValueNearLabelRaw_(worksheet, 'KCP Coordinate Latitude (Decimal)') ||
                           getValueNearLabelRaw_(worksheet, 'Mill Coordinate Latitude (Decimal)') ||
                           getValueNearLabelRaw_(worksheet, 'KCP Coordinate Latitude') ||
                           getValueNearLabel(rows, 'KCP Coordinate Latitude (Decimal)') ||
                           getValueNearLabel(rows, 'Mill Coordinate Latitude (Decimal)') ||
                           getValueNearLabel(rows, 'KCP Coordinate Latitude'), 'Latitude');
      record['Longitude'] = normalizeCoordinate(
                            getValueNearLabelRaw_(worksheet, 'Longitude (Decimal)') ||
                            getValueNearLabelRaw_(worksheet, 'Longitude') ||
                            getValueNearLabel(rows, 'Longitude (Decimal)') ||
                            getValueNearLabel(rows, 'Longitude'), 'Longitude');
      record['Mill Category'] = getValueNearLabel(rows, 'KCP Category') || getValueNearLabel(rows, 'Mill Category');
      record['Mill Capacity (Ton/Hour)'] = getValueNearLabel(rows, 'KCP Capacity (Ton/Hour)') ||
                                           getValueNearLabel(rows, 'KCP Capacity') ||
                                           getValueNearLabel(rows, 'Mill Capacity (Ton/Hour)');
      record['Storage Tank Capacity'] = getValueNearLabel(rows, 'Storage Tank Capacity');
      record['Silo Capacity'] = getValueNearLabel(rows, 'Silo Capacity');
      record['Commissioning Date'] = getValueNearLabel(rows, 'Commisioning Date') ||
                                     getValueNearLabel(rows, 'Commissioning Date');
    } else {
      // ── Mill-specific field mapping ──
      record['Mill Name'] = getValueNearLabel(rows, 'Mill Name');
      record['Mill Address'] = getValueNearLabel(rows, 'Mill Address');
      const millDistance = getDistanceFields(rows, 'Distance to Refinery');
      record['Distance to Refinery (km)'] = millDistance.refinery;
      record['Distance'] = millDistance.distanceKm;
      record['Latitude'] = normalizeCoordinate(
                           getValueNearLabelRaw_(worksheet, 'Mill Coordinate Latitude (Decimal)') ||
                           getValueNearLabel(rows, 'Mill Coordinate Latitude (Decimal)'), 'Latitude');
      record['Longitude'] = normalizeCoordinate(
                            getValueNearLabelRaw_(worksheet, 'Longitude (Decimal)') ||
                            getValueNearLabel(rows, 'Longitude (Decimal)'), 'Longitude');
      record['Mill Category'] = getValueNearLabel(rows, 'Mill Category');
      record['Mill Capacity (Ton/Hour)'] = getValueNearLabel(rows, 'Mill Capacity (Ton/Hour)');
      record['Sterilizer Type'] = getValueNearLabel(rows, 'Sterilizer Type');
      record['Storage Tank Capacity'] = getValueNearLabel(rows, 'Storage Tank Capacity');
      record['Silo Capacity'] = getValueNearLabel(rows, 'Silo Capacity');
      record['Commissioning Date'] = getValueNearLabel(rows, 'Commisioning Date');
    }

    // Capture second Mobile Phone as Sustainability PIC Mobile (if present)
    let mobileHits = [];
    for (let r = 0; r < Math.min(rows.length, 80); r++) {
      const row = rows[r] || [];
      for (let c = 0; c < Math.min(row.length, 12); c++) {
        if (normalizeCellText(row[c]).toLowerCase() !== 'mobile phone') continue;
        mobileHits.push(normalizeCellText(row[c + 2]) || normalizeCellText(row[c + 1]) || '');
      }
    }
    if (mobileHits.length > 1) record['Sustainability PIC Mobile'] = mobileHits[1];
    else record['Sustainability PIC Mobile'] = '';

    // ── Product table ────────────────────────────────────────────────────────
    // Supports both label variants:
    //   Mill form : "Product to be Supply"  (col2=Main Product, col3=Avg, col4=YIELD,
    //                                        col5=Other Product, col6=Avg)
    //   KCP  form : "Product to be Produce" (col2=Product, col3=Avg, col4=YIELD,
    //                                        no other-product side columns)
    //
    // Column positions are read dynamically from the header row itself so the parser
    // stays correct even if the sheet layout shifts by one column.
    //
    // Hard stop: the data loop exits as soon as it reaches a row whose col0 looks like
    // a certification / section marker (e.g. "C. CERTIFICATION", "RSPO", "Cert Number",
    // "Certification Members") so those labels can never be mistaken for product names.
    (function parseProductTable() {
      // Find the header row — accept either label variant
      const pHeader = findRowByFirstCellAny(rows,
        ['Product to be Supply', 'Product to be Produce'], 0);
      if (pHeader < 0) return;

      // ── Detect column positions from the header row itself ──────────────
      const hRow = rows[pHeader] || [];
      let mainCol = -1, avgCol = -1, yieldCol = -1, otherCol = -1, otherAvgCol = -1;

      for (let c = 0; c < hRow.length; c++) {
        const h = normalizeCellText(hRow[c]).toLowerCase();
        if (!h) continue;
        // "Main Product" or "Product" (KCP uses just "Product")
        if ((h === 'main product' || h === 'product') && mainCol < 0) { mainCol = c; continue; }
        // "Average …" — first occurrence = main avg, second = other avg
        if (h.indexOf('average') !== -1 || h.indexOf('avg') !== -1) {
          if (avgCol < 0) { avgCol = c; continue; }
          if (otherAvgCol < 0) { otherAvgCol = c; continue; }
        }
        // "YIELD"
        if (h === 'yield' && yieldCol < 0) { yieldCol = c; continue; }
        // "Other Product"
        if (h.indexOf('other') !== -1 && h.indexOf('product') !== -1 && otherCol < 0) {
          otherCol = c; continue;
        }
      }

      // Fallback to known fixed positions when header cells are blank / merged
      if (mainCol  < 0) mainCol  = 2;
      if (avgCol   < 0) avgCol   = 3;
      if (yieldCol < 0) yieldCol = 4;
      // otherCol / otherAvgCol stay -1 for KCP forms (no other-product columns)
      if (otherCol < 0 && !isKcp) { otherCol = 5; otherAvgCol = 6; }

      // KCP spreadsheet: merged headers often put "Average Production / Month" under col 4
      // while real data is still contiguous at mainCol+1 (tons) and mainCol+2 (yield %).
      if (isKcp && !isMill && mainCol >= 0) {
        let scanHdr = pHeader + 1;
        while (scanHdr < Math.min(rows.length, pHeader + 8)) {
          if (scanHdr >= rows.length) break;
          const qc0Scan = normalizeCellText(cellAt(rows, scanHdr, 0));
          if (qc0Scan && /\bquality\b/i.test(qc0Scan)) { scanHdr++; continue; }
          const candMain = normalizeCellText(cellAt(rows, scanHdr, mainCol));
          if (!candMain || candMain.toUpperCase() === 'PLEASE SELECT') { scanHdr++; continue; }
          const candA = normalizeCellText(cellAt(rows, scanHdr, mainCol + 1));
          const candB = normalizeCellText(cellAt(rows, scanHdr, mainCol + 2));
          const yLooksPct = !!(candB && (/%/.test(candB) || (!isNaN(parseFloat(candB)) && Math.abs(parseFloat(candB)) <= 1 && !/ton/i.test(candB))));
          const avgLooksQty = !!(candA && (/ton/i.test(candA) || /\d/.test(candA)));
          if (candA && candB && yLooksPct && avgLooksQty && (yieldCol !== mainCol + 2 || avgCol !== mainCol + 1)) {
            avgCol   = mainCol + 1;
            yieldCol = mainCol + 2;
          }
          break;
        }
      }

      // ── Hard-stop sentinel: labels that mark the end of the product block ──
      // Any row whose first cell matches one of these must not be parsed as a product.
      function isCertOrSectionRow(r) {
        const v = normalizeLooseKey(cellAt(rows, r, 0));
        if (!v) return false;
        // Explicit certification-section markers
        if (v === 'ccertification' || v.indexOf('certificationmembers') !== -1) return true;
        // Section letters: "C.", "D.", etc.
        if (/^[c-z]/.test(v) && (cellAt(rows, r, 0) || '').trim().match(/^[C-Z]\.\s/)) return true;
        // Known cert names that appear as first-column labels in the cert table
        const certNames = ['rspo','iscc','ispomill','ispohilir','ispo','ins','gmpplus','gmp','ggl','halal'];
        if (certNames.some(function(n) { return v === n || v.indexOf(n) === 0; })) return true;
        // Catch-all: if col1 is "Yes" or "No" and col2 is "Cert Number" it's a cert row
        const col1 = normalizeLooseKey(cellAt(rows, r, 1));
        const col2 = normalizeLooseKey(cellAt(rows, r, 2));
        if ((col1 === 'yes' || col1 === 'no') && col2 === 'certnumber') return true;
        return false;
      }

      // ── Iterate data rows — one JSON line per spreadsheet row (no lost mains) ──
      /** Row-aligned canonical product rows (survives save/load relational MAIN). */
      const productLines = [];

      for (let r = pHeader + 1; r < Math.min(rows.length, pHeader + 16); r++) {
        if (isCertOrSectionRow(r)) break;

        const qc0 = normalizeCellText(cellAt(rows, r, 0));
        // CPO Quality / PK Quality / … must never be mistaken for product data
        if (qc0 && /\bquality\b/i.test(qc0)) continue;

        const rawMain   = cellAt(rows, r, mainCol);
        const rawAvg    = cellAt(rows, r, avgCol);
        const rawYield  = cellAt(rows, r, yieldCol);
        const rawOther  = otherCol >= 0 ? cellAt(rows, r, otherCol) : '';
        const rawOthAvg = otherAvgCol >= 0 ? cellAt(rows, r, otherAvgCol) : '';

        const stripSel = function(v) {
          const s = normalizeCellText(v);
          if (!s || s.toUpperCase() === 'PLEASE SELECT') return '';
          return s;
        };
        const mainN  = stripSel(rawMain);
        const avgN   = stripSel(rawAvg);
        const yldN   = stripSel(rawYield);
        const otherN = stripSel(rawOther);
        const oAvgN  = stripSel(rawOthAvg);

        if (!mainN && !otherN) continue;

        productLines.push({
          main     : mainN,
          mainAvg  : avgN,
          yield    : yldN,
          other    : otherN,
          otherAvg : oAvgN,
        });
      }

      try {
        if (productLines.length) {
          record['SDD - Product Lines JSON'] = JSON.stringify(productLines);
        }
      } catch (jsonErr) { /* ignore */ }

      // Legacy flat columns: best-effort for old readers (JSON is authoritative)
      if (productLines.length) {
        const mainOnly = productLines.filter(function(pl) { return !!pl.main; }).slice(0, 3);
        if (mainOnly[0]) {
          record['Main Product']                              = mainOnly[0].main || '';
          record['Main Product Avg Production/Month (Ton)'] = mainOnly[0].mainAvg || '';
          record['Main Product Yield']                      = mainOnly[0].yield || '';
        }
        if (mainOnly[1]) {
          record['Main Product 2']                              = mainOnly[1].main || '';
          record['Main Product 2 Avg Production/Month (Ton)'] = mainOnly[1].mainAvg || '';
          record['Main Product 2 Yield']                      = mainOnly[1].yield || '';
        }
        if (mainOnly[2]) {
          record['Main Product 3']                              = mainOnly[2].main || '';
          record['Main Product 3 Avg Production/Month (Ton)'] = mainOnly[2].mainAvg || '';
          record['Main Product 3 Yield']                      = mainOnly[2].yield || '';
        }
        var oi = 1;
        productLines.forEach(function(pl) {
          if (oi > 3 || !pl.other) return;
          record['Other Product ' + oi]                        = pl.other || '';
          record['Other Product ' + oi + ' Avg/Month (Ton)'] = pl.otherAvg || '';
          oi++;
        });
      }
    })();

    // Quality rows — KCP uses CPKO/PKE Quality, Mill uses CPO/PK Quality
    // Try all variants so both forms are covered
    function fillQuality(labelVariants, ffa, mi, dobi) {
      for (let v = 0; v < labelVariants.length; v++) {
        const qRow = findRowByFirstCell(rows, labelVariants[v], 0);
        if (qRow < 0) continue;
        record[ffa]  = cellAt(rows, qRow, 3);
        record[mi]   = cellAt(rows, qRow, 5);
        record[dobi] = cellAt(rows, qRow, 7);
        return;
      }
    }
    fillQuality(['CPO Quality', 'CPKO Quality'], 'CPO Quality - FFA', 'CPO Quality - M&I', 'CPO Quality - DOBI');
    fillQuality(['PK Quality', 'PKE Quality'],   'PK Quality - FFA',  'PK Quality - MOIST', 'PK Quality - DIRT');

    // Certification rows — handle both Mill and KCP cert names
    function fillCert(certNameVariants, statusKey) {
      for (let v = 0; v < certNameVariants.length; v++) {
        const rr = findRowByFirstCell(rows, certNameVariants[v], 0);
        if (rr < 0) continue;
        record[statusKey + ' Status']     = cellAt(rows, rr, 1);
        record[statusKey + ' Cert No']    = cellAt(rows, rr, 3);
        record[statusKey + ' Valid Start'] = cellAt(rows, rr, 6);
        record[statusKey + ' Valid End']   = cellAt(rows, rr, 7);
        record[statusKey + ' Cert Body']   = cellAt(rows, rr, 8);
        return;
      }
    }
    fillCert(['RSPO'],                      'RSPO');
    fillCert(['ISCC'],                      'ISCC');
    fillCert(['ISPO Mill', 'ISPO Hilir', 'ISPO'], 'ISPO');
    fillCert(['INS', 'GMP PLUS', 'GMP+'],   'INS');
    fillCert(['GGL', 'HALAL'],              'GGL');

    // Legality rows — same labels in both form types
    function fillLegal(labels, keyPrefix) {
      const rr = findRowByFirstCellAny(rows, labels, 0);
      if (rr < 0) return;
      record[keyPrefix + ' Type']       = cellAt(rows, rr, 1);
      record[keyPrefix + ' No']         = cellAt(rows, rr, 3);
      record[keyPrefix + ' Issue Date'] = cellAt(rows, rr, 6);
    }
    fillLegal(['HGU/HGB', 'HGU / HGB', 'HGU HGB'], 'HGU/HGB');
    fillLegal(['IUP', 'IUI', 'IUP / IUP - P', 'IUP-IUP-P', 'IUP / IUP-P'], 'IUP');
    fillLegal(['Izin Lokasi', 'IZIN LOKASI'], 'Izin Lokasi');
    fillLegal(['Izin Lingkungan', 'IZIN LINGKUNGAN'], 'Izin Lingkungan');
    fillLegal(['IMB/PBG', 'IMB / PBG', 'IMB PBG'], 'IMB/PBG');
    fillLegal(['NIB'], 'NIB');

    return hasMeaningfulImportData(record) ? record : null;
  }

  function extractTraceabilityToSddRecords(worksheet) {
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'dd mmm yyyy' });
    if (!rows.length) return [];
    if (cellAt(rows, 0, 0).toLowerCase().indexOf('traceability data') === -1) return [];

    // Raw rows (raw:true) used ONLY for coordinate columns to bypass display-format
    // distortion (e.g. accounting format "_(* #,##0.00_)" truncating 101.734934 to "101.73").
    // For a Number cell, String(rawRows[r][c]) always produces dot-decimal notation.
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
    // Helper: prefer raw numeric value for coordinate cells; fall back to formatted string.
    function rawCellAt(r, c) {
      const raw = rawRows[r] && rawRows[r][c];
      if (raw === undefined || raw === null) return cellAt(rows, r, c);
      const sv = String(raw).trim();
      return sv !== '' ? sv : cellAt(rows, r, c);
    }

    const out = [];

    // A. Mill List -> TML columns
    const millStart = findRowByFirstCell(rows, 'A. Mill List', 0);
    if (millStart >= 0) {
      let header = -1;
      for (let r = millStart + 1; r < Math.min(rows.length, millStart + 20); r++) {
        if (cellAt(rows, r, 0).toUpperCase() === 'COMPANY NAME') { header = r; break; }
      }
      if (header >= 0) {
        for (let r = header + 1; r < rows.length; r++) {
          const first = cellAt(rows, r, 0);
          const ft = String(first || '').trim();
          if (ft.toLowerCase() === 'b. ffb supplier list') break;
          if (/^[C-Z]\.\s/i.test(ft) && !normalizeCellText(cellAt(rows, r, 1)) && !normalizeCellText(cellAt(rows, r, 2))) break;
          const company = cellAt(rows, r, 0);
          const mill = cellAt(rows, r, 1);
          if (!company && !mill) continue;
          out.push({
            'TML - Company Name': company,
            'TML - Mill Name': mill,
            'TML - UML ID': cellAt(rows, r, 2),
            'TML - Village': cellAt(rows, r, 3),
            'TML - Sub District': cellAt(rows, r, 4),
            'TML - District': cellAt(rows, r, 5),
            'TML - Capacity (Ton/Hour)': cellAt(rows, r, 6),
            'TML - Latitude': normalizeCoordinate(rawCellAt(r, 7), 'TML - Latitude'),
            'TML - Longitude': normalizeCoordinate(rawCellAt(r, 8), 'TML - Longitude'),
            'TML - Legality': cellAt(rows, r, 9),
            'TML - ISPO (Y/N)': cellAt(rows, r, 10),
            'TML - RSPO (Y/N)': cellAt(rows, r, 11),
            'TML - ISCC (Y/N)': cellAt(rows, r, 12),
            'TML - Total Supply CPO/PK (Ton)': cellAt(rows, r, 13)
          });
        }
      }
    }

    // B. FFB Supplier List -> FFB columns
    const ffbStart = findRowByFirstCell(rows, 'B. FFB Supplier List', 0);
    if (ffbStart >= 0) {
      let header = -1;
      for (let r = ffbStart + 1; r < Math.min(rows.length, ffbStart + 28); r++) {
        const c0 = normalizeCellText(cellAt(rows, r, 0)).toUpperCase();
        const c1 = normalizeCellText(cellAt(rows, r, 1)).toUpperCase();
        const c2 = normalizeCellText(cellAt(rows, r, 2)).toUpperCase();
        if (c0 === 'MILL NAME' && (c1.indexOf('SUPPLIER') >= 0 || c2.indexOf('SUPPLIER') >= 0)) {
          header = r;
          break;
        }
      }
      if (header >= 0) {
        const h1 = normalizeCellText(cellAt(rows, header, 1)).toUpperCase();
        const h2 = normalizeCellText(cellAt(rows, header, 2)).toUpperCase();
        const hasSupplierGroupName =
          h1 === 'SUPPLIER GROUP NAME' ||
          (h2.indexOf('SUPPLIER') >= 0 && h1.indexOf('GROUP') >= 0);
        const supplierCol = hasSupplierGroupName ? 2 : 1;
        const villageCol = hasSupplierGroupName ? 3 : 2;
        const subDistrictCol = hasSupplierGroupName ? 4 : 3;
        const districtCol = hasSupplierGroupName ? 5 : 4;
        const categoryCol = hasSupplierGroupName ? 6 : 5;
        const concessionCol = hasSupplierGroupName ? 7 : 6;
        const plantedAreaCol = hasSupplierGroupName ? 8 : 7;
        const smallholdersCol = hasSupplierGroupName ? 9 : 8;
        const plantedYearCol = hasSupplierGroupName ? 10 : 9;
        const legalityCol = hasSupplierGroupName ? 11 : 10;
        const latCol = hasSupplierGroupName ? 12 : 11;
        const longCol = hasSupplierGroupName ? 13 : 12;
        const ispoCol = hasSupplierGroupName ? 14 : 13;
        const rspoCol = hasSupplierGroupName ? 15 : 14;
        const isccCol = hasSupplierGroupName ? 16 : 15;
        const totalSupplyCol = hasSupplierGroupName ? 17 : 16;
        const ffbSeen = {};
        for (let r = header + 1; r < rows.length; r++) {
          const raw0 = cellAt(rows, r, 0);
          const t0 = String(raw0 || '').trim();
          if (t0.toLowerCase() === 'a. mill list') break;
          if (/^[D-Z]\.\s/i.test(t0)) break;
          const restEmpty = !normalizeCellText(cellAt(rows, r, 1)) && !normalizeCellText(cellAt(rows, r, 2));
          if (/^[C-Z]\.\s/i.test(t0) && restEmpty) break;
          const supplier = cellAt(rows, r, supplierCol) || (hasSupplierGroupName ? '' : cellAt(rows, r, 2));
          const category = cellAt(rows, r, categoryCol);
          if (!supplier || supplier.toUpperCase() === 'PLEASE SELECT') continue;
          if (!cellAt(rows, r, 0) && !supplier && !category) continue;
          const dedupeK = [normalizeCellText(raw0), normalizeCellText(supplier), normalizeCellText(category), normalizeCellText(cellAt(rows, r, villageCol))].join('|');
          if (ffbSeen[dedupeK]) continue;
          ffbSeen[dedupeK] = true;
          out.push({
            'FFB - Mill Name': cellAt(rows, r, 0),
            'FFB - Supplier Group Name': hasSupplierGroupName ? cellAt(rows, r, 1) : '',
            'FFB - Supplier Name': supplier,
            'FFB - Village': cellAt(rows, r, villageCol),
            'FFB - Sub District': cellAt(rows, r, subDistrictCol),
            'FFB - District': cellAt(rows, r, districtCol),
            'FFB - Supplier Category': category,
            'FFB - Concession Area (Ha)': cellAt(rows, r, concessionCol),
            'FFB - Planted Area (Ha)': cellAt(rows, r, plantedAreaCol),
            'FFB - Number of Smallholders': cellAt(rows, r, smallholdersCol),
            'FFB - Planted Year': cellAt(rows, r, plantedYearCol),
            'FFB - Legality': cellAt(rows, r, legalityCol),
            'FFB - Latitude': normalizeCoordinate(rawCellAt(r, latCol), 'FFB - Latitude'),
            'FFB - Longitude': normalizeCoordinate(rawCellAt(r, longCol), 'FFB - Longitude'),
            'FFB - ISPO (Y/N)': cellAt(rows, r, ispoCol),
            'FFB - RSPO (Y/N)': cellAt(rows, r, rspoCol),
            'FFB - ISCC (Y/N)': cellAt(rows, r, isccCol),
            'FFB - Total Supply FFB (Ton)': cellAt(rows, r, totalSupplyCol)
          });
        }
      }
    }

    return out.filter(hasMeaningfulImportData);
  }

  function buildImportedSddPayloadsFromWorkbook(workbook) {
    if (!workbook || !workbook.SheetNames || !workbook.Sheets) return [];
    let records = [];
    workbook.SheetNames.forEach(function(name) {
      const ws = workbook.Sheets[name];
      if (!ws) return;
      const n = String(name || '').trim().toLowerCase();
      if (n === 'main form') {
        const rec = extractMainFormToSddRecord(ws);
        if (rec) records.push(rec);
        return;
      }
      if (n === 'traceability') {
        records = records.concat(extractTraceabilityToSddRecords(ws));
        return;
      }
      records = records.concat(parseSheetToSddRecords(ws));
    });
    return records.filter(hasMeaningfulImportData);
  }

  /** Cache imported rows locally for merge on Save as Draft / Submit. */
  function cacheSddImportRowsFromWorkbook() {
    if (!supplierWorkbook) {
      /* Rows may come from "Load saved screening" — do not clear them here */
      return;
    }
    window._sddImportedRows = [];
    window._sddImportFirstRow = null;
    window._sddImportedRows = buildImportedSddPayloadsFromWorkbook(supplierWorkbook);
    window._sddImportFirstRow = window._sddImportedRows[0] || null;
    window._loadedPrimarySddRow = window._sddImportFirstRow
      ? Object.assign({}, window._sddImportFirstRow)
      : null;
    window._nblCheckResult = null;
  }

  /**
   * Excel row first, then manual screening fields.
   * Manual keys (SCR-/GRV/PRI) always win; other keys only overwrite import when non-empty
   * so empty form fields do not wipe Excel columns.
   */
  function mergeSddImportIntoPayload(payload) {
    const imp = window._sddImportFirstRow;
    if (!imp || typeof imp !== 'object') return Object.assign({}, payload);
    const out = Object.assign({}, imp);
    // Keputusan (statusSDD) hanya dari UI / server — jangan ambil dari baris Excel
    // (template sering punya kolom "Approved" dsb. yang menimpa Hold saat Submit).
    ['statusSDD', 'statusSdd', 'Status SDD', 'statusBossDecision', 'Status Boss Decision', 'StatusSDD', 'STATUSSDD'].forEach(function(k) {
      delete out[k];
    });
    const manualKeys = Object.keys(payload || {});
    manualKeys.forEach(function(k) {
      const v = payload[k];
      const isScreeningKey = /^SCR - |^GRV\d+|^PRI\d+/.test(k);
      if (isScreeningKey) {
        out[k] = v;
      } else if (v !== undefined && v !== null && String(v).trim() !== '') {
        out[k] = v;
      }
    });
    // PATCH: enforce dot-decimal on all coordinate fields before sending to GAS.
    // normalizeCoordinate converts "0,318318" -> "0.318318" and leaves
    // already-valid dot-decimal strings unchanged. This is the last fence before POST.
    ['Latitude', 'Longitude',
     'TML - Latitude', 'TML - Longitude',
     'FFB - Latitude', 'FFB - Longitude'].forEach(function(ck) {
      if (out[ck] !== undefined && out[ck] !== null) {
        out[ck] = normalizeCoordinate(out[ck], ck);
      }
    });
    return out;
  }

  function supplierDisplayAllSheets() {
    if (!supplierWorkbook) return;
    const container = document.getElementById('supplierExcelData');
    container.innerHTML = '';
    window._scrData = {};
    // Fresh workbook view must not reuse relational keys from a previous load/session;
    // otherwise Approve calls setSubmissionStatus on the wrong row and UI can desync.
    window._sddSubmissionId = null;
    window._scrLoadedKey = '';
    window._scrLoadedRowNum = null;
    window._sddIsLoadedSaved = false;
    const fileName = document.getElementById('supplierExcelFile')?.files[0]?.name || '';
    const tp = getCurrentSddSupplierType();
    window._scrKey = ((tp ? tp + '_' : '') + fileName.replace(/[^a-zA-Z0-9]/g,'_').toLowerCase()) || 'default';
    window._loadedPrimarySddRow = null;
    // Determine which sheet gets the SDD screening form (only one, to avoid duplicate IDs)
    const sddSheetIdx = supplierWorkbook.SheetNames.findIndex(function(n) {
      return String(n).trim().toLowerCase() === 'sdd data';
    });
    const scrFormSheetName = supplierWorkbook.SheetNames[sddSheetIdx >= 0 ? sddSheetIdx : 0];

    supplierWorkbook.SheetNames.forEach(function(name) {
      const worksheet = supplierWorkbook.Sheets[name];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'dd mmm yyyy' });
      // Create a wrapper div per sheet
      const wrapper = document.createElement('div');
      const sheetId = 'sheet-body-' + name.replace(/\s+/g,'-');
      wrapper.innerHTML = '<div onclick="(function(h){var b=document.getElementById(\'' + sheetId + '\');var open=b.style.display!==\'none\';b.style.display=open?\'none\':\'block\';h.querySelector(\'.sec-chev\').style.transform=open?\'rotate(-90deg)\':\'rotate(0deg)\';})(this)" style="display:flex;align-items:center;justify-content:space-between;font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#2C2828;margin:20px 0 0;padding:12px 16px;background:rgba(44,40,40,0.04);border:1.5px solid rgba(44,40,40,0.14);border-radius:10px 10px 0 0;cursor:pointer;user-select:none;">'
        + '<span>Sheet: ' + name + '</span>'
        + '<span class="sec-chev" style="display:inline-flex;width:22px;height:22px;align-items:center;justify-content:center;border-radius:50%;background:rgba(44,40,40,0.08);transition:transform 0.2s;"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#3D3535" stroke-width="2.5" stroke-linecap="round"><polyline points="6 9 12 15 18 9"/></svg></span>'
        + '</div><div id="' + sheetId + '" style="border:1.5px solid rgba(44,40,40,0.14);border-top:none;border-radius:0 0 10px 10px;padding:10px 6px 10px;margin-bottom:8px;"></div>';
      container.appendChild(wrapper);
      // Only the designated SDD sheet gets the screening form appended.
      // Other sheets use appendScrForm:false to prevent duplicate IDs
      // (scr-form-wrap, sdd-approver-approve, etc.) which break Save/Submit buttons.
      const sheetBody = wrapper.querySelector('[id^="sheet-body-"]');
      const sheetOpts = (name === scrFormSheetName) ? {} : { appendScrForm: false };
      supplierDisplayExcelDataTo(sheetBody, jsonData, sheetOpts);
    });
    cacheSddImportRowsFromWorkbook();
  }

  function supplierDisplayExcelDataTo(container, data, opts) {
    supplierDisplayExcelData(data, container, opts);
  }

  function supplierDisplayExcelData(data, container, opts) {
    opts = opts || {};
    if (!container) container = document.getElementById('supplierExcelData');
    const uiSupplierType = getCurrentSddSupplierType();
    const useKcpLabels = uiSupplierType === 'KCP';
    _sectionIdx = 0;
    if (!data || !data.length) {
      container.innerHTML = '<p style="color:#9C8080;font-size:13px;padding:16px;">No data found in this sheet.</p>';
      return;
    }

    function cellVal(row, col) {
      const v = row && row[col];
      if (v === undefined || v === null) return '';
      const s = String(v).trim();
      // ISO date string from XLSX with cellDates:true
      if (s.match(/^\d{4}-\d{2}-\d{2}T/)) {
        const d = new Date(s);
        return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
      }
      // Formatted date string from dateNF e.g. "02 Mar 2024" — pass through
      // Numbers: remove trailing .0
      if (s.match(/^\d+\.0+$/)) return s.replace(/\.0+$/, '');
      return s;
    }

    function isSectionHeader(row) {
      const v = cellVal(row, 0);
      if (!v) return false;
      const allOtherEmpty = (row || []).slice(1).every(c => c === undefined || c === null || String(c).trim() === '');
      return allOtherEmpty && (v.match(/^[A-Z]\.\s/) || v === 'MILL SUPPLIER REGISTRATION FORM' || v === 'Traceability Data');
    }

    function relabelForSupplier(k) {
      if (!useKcpLabels) return k;
      const src = String(k || '');
      if (!src) return src;
      return src
        .replace(/\bMILL\b/g, 'KCP')
        .replace(/\bMill\b/g, 'KCP')
        .replace(/\bmill\b/g, 'kcp');
    }

    function kv(k, v) {
      const shownKey = relabelForSupplier(k);
      return '<div style="display:flex;gap:0;padding:4px 0;border-bottom:1px solid rgba(74,28,28,0.05);font-size:13.5px;line-height:1.5;">'
        + '<span style="color:#5F4A48;font-weight:600;width:260px;flex-shrink:0;">' + shownKey + '</span>'
        + '<span style="color:#9C8080;font-weight:500;width:20px;flex-shrink:0;">:</span>'
        + '<span style="color:#1A0A0A;">' + (v || '<span style="color:#B09A9A;">—</span>') + '</span>'
        + '</div>';
    }

    var _sectionIdx = 0;
    function sectionCard(title, innerHtml) {
      var idx = _sectionIdx++;
      var sId = 'sec-body-' + idx;
      var hId = 'sec-hdr-' + idx;
      const shownTitle = relabelForSupplier(title);
      var isApprover = false;

      // Legacy: approver-only section colours (disabled — single staff UX).
      var secLetter = shownTitle ? shownTitle.trim().charAt(0).toUpperCase() : '';
      var headerBg, headerBorder, chevronBg, chevronStroke, labelColor;
      if (isApprover && secLetter) {
        if (secLetter === 'A') {
          headerBg = 'linear-gradient(135deg,#1e3a8a 0%,#2563eb 100%)';
          headerBorder = '#1e3a8a'; chevronBg = 'rgba(255,255,255,0.2)'; chevronStroke = '#fff'; labelColor = '#fff';
        } else if (secLetter === 'B') {
          headerBg = 'linear-gradient(135deg,#065f46 0%,#059669 100%)';
          headerBorder = '#065f46'; chevronBg = 'rgba(255,255,255,0.2)'; chevronStroke = '#fff'; labelColor = '#fff';
        } else if (secLetter === 'C') {
          headerBg = 'linear-gradient(135deg,#4c1d95 0%,#7c3aed 100%)';
          headerBorder = '#4c1d95'; chevronBg = 'rgba(255,255,255,0.2)'; chevronStroke = '#fff'; labelColor = '#fff';
        } else if (secLetter === 'D') {
          headerBg = 'linear-gradient(135deg,#92400e 0%,#d97706 100%)';
          headerBorder = '#92400e'; chevronBg = 'rgba(255,255,255,0.2)'; chevronStroke = '#fff'; labelColor = '#fff';
        } else {
          headerBg = 'rgba(44,40,40,0.03)'; headerBorder = 'rgba(44,40,40,0.1)';
          chevronBg = 'rgba(44,40,40,0.07)'; chevronStroke = '#3D3535'; labelColor = '#2C2828';
        }
      } else {
        // Staff: clean, subtle accent per section — dark left border, light header
        var SEC_ACCENTS = ['#1D4ED8','#0F766E','#6D28D9','#B45309','#0369A1','#065F46'];
        var accent = SEC_ACCENTS[idx % SEC_ACCENTS.length];
        headerBg = '#F8F7F6';
        headerBorder = 'rgba(0,0,0,0.07)';
        chevronBg = 'rgba(0,0,0,0.06)';
        chevronStroke = '#555';
        labelColor = accent;
      }

      // Approver: section B, C, D dimulai collapsed
      var startCollapsed = isApprover && (secLetter === 'B' || secLetter === 'C' || secLetter === 'D');
      var bodyDisplay   = startCollapsed ? 'none' : 'block';
      var chevronRotate = startCollapsed ? 'rotate(-90deg)' : 'rotate(0deg)';

      var cardBorderLeft = !isApprover ? '3px solid ' + labelColor : '1px solid rgba(44,40,40,0.09)';
      return '<div style="background:white;border:1px solid rgba(44,40,40,0.08);border-left:' + cardBorderLeft + ';border-radius:6px;margin-bottom:5px;overflow:hidden;">'
        + (shownTitle ? '<div id="' + hId + '" onclick="(function(h){var b=document.getElementById(\'' + sId + '\');var open=b.style.display!==\'none\';b.style.display=open?\'none\':\'block\';h.querySelector(\'.sec-chev\').style.transform=open?\'rotate(-90deg)\':\'rotate(0deg)\';})(this)" style="display:flex;align-items:center;justify-content:space-between;padding:9px 14px;cursor:pointer;user-select:none;border-bottom:1px solid ' + headerBorder + ';background:' + headerBg + ';">'
          + '<span style="font-size:10.5px;font-weight:700;letter-spacing:1.1px;text-transform:uppercase;color:' + labelColor + ';">' + shownTitle + '</span>'
          + '<span class="sec-chev" style="display:inline-flex;width:20px;height:20px;align-items:center;justify-content:center;border-radius:50%;background:' + chevronBg + ';transition:transform 0.2s;transform:' + chevronRotate + ';"><svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="' + chevronStroke + '" stroke-width="2.5" stroke-linecap="round"><polyline points="6 9 12 15 18 9"/></svg></span>'
          + '</div>' : '')
        + '<div id="' + sId + '" style="padding:12px 16px;display:' + bodyDisplay + ';">' + innerHtml + '</div>'
        + '</div>';
    }

    function tableCard(title, headers, rows) {
      var idx = _sectionIdx++;
      var sId = 'sec-body-' + idx;
      const colWidths = {
        'Company Name': 200, 'Mill Name': 200, 'Supplier Name': 200,
        'UML ID': 140, 'Village': 160, 'Sub District': 160, 'District': 160,
        'Capacity': 100, 'Lat': 130, 'Long': 130, 'Legality': 120,
        'ISPO': 70, 'RSPO': 70, 'ISCC': 70,
        'Total Supply CPO/PK (Ton)': 180, 'Total Supply FFB (Ton)': 180,
        'Category': 140, 'Concession Area': 150, 'Planted Area': 130,
        '# Smallholders': 130, 'Planted Year': 130, 'Action Request': 160,
      };
      var tableHtml = '<table style="border-collapse:collapse;font-size:13px;table-layout:auto;">';
      tableHtml += '<thead><tr>' + headers.map(h => {
        const w = colWidths[h] || 140;
        return '<th style="background:#f7f4f4;color:#5F4A48;font-weight:700;padding:9px 14px;border:1px solid #E8E3E3;text-align:left;white-space:nowrap;min-width:' + w + 'px;">' + h + '</th>';
      }).join('') + '</tr></thead>';
      tableHtml += '<tbody>' + rows.map((r, i) => '<tr style="background:' + (i%2===0?'#fff':'#fdfafa') + ';">' + r.map((c,ci) => {
        const w = colWidths[headers[ci]] || 140;
        return '<td style="padding:8px 14px;border:1px solid #E8E3E3;color:#1A0A0A;vertical-align:top;white-space:nowrap;min-width:' + w + 'px;">' + (c||'—') + '</td>';
      }).join('') + '</tr>').join('') + '</tbody>';
      tableHtml += '</table>';
      let t = '<div style="background:white;border:1px solid rgba(44,40,40,0.09);border-radius:12px;margin-bottom:10px;">';
      if (title) t += '<div onclick="(function(h){var b=document.getElementById(\'' + sId + '\');var open=b.style.display!==\'none\';b.style.display=open?\'none\':\'block\';h.querySelector(\'.sec-chev\').style.transform=open?\'rotate(-90deg)\':\'rotate(0deg)\';})(this)" style="display:flex;align-items:center;justify-content:space-between;padding:10px 16px;cursor:pointer;user-select:none;border-bottom:1.5px solid rgba(44,40,40,0.1);background:rgba(44,40,40,0.03);border-radius:12px 12px 0 0;">'
        + '<span style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#2C2828;">' + title + '</span>'
        + '<span class="sec-chev" style="display:inline-flex;width:22px;height:22px;align-items:center;justify-content:center;border-radius:50%;background:rgba(44,40,40,0.07);transition:transform 0.2s;"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#3D3535" stroke-width="2.5" stroke-linecap="round"><polyline points="6 9 12 15 18 9"/></svg></span>'
        + '</div>';
      t += '<div id="' + sId + '" style="padding:10px 16px;overflow-x:auto;-webkit-overflow-scrolling:touch;">' + tableHtml + '</div></div>';
      return t;
    }

    // Detect which sheet this is
    const firstNonEmpty = data.find(r => r && r.some(c => c !== null && c !== undefined && String(c).trim() !== ''));
    const firstVal = firstNonEmpty ? String(firstNonEmpty[0] || '').trim() : '';
    const isTraceability = firstVal === 'Traceability Data';

    if (isTraceability) {
      // ── TRACEABILITY SHEET ──
      let html = '';
      let i = 0;

      while (i < data.length) {
        const row = data[i];
        const v0 = cellVal(row, 0);

        if (v0 === 'A. Mill List') {
          // Find header row (next row with content in col 0 or 1)
          i++;
          // skip sub-header rows until we hit the real column headers
          while (i < data.length && !cellVal(data[i], 0) && !cellVal(data[i], 1)) i++;
          // col headers might be split across 2 rows (row 5 has group headers, row 6 has col headers)
          // find the row with COMPANY NAME
          let headerIdx = i;
          while (headerIdx < data.length && cellVal(data[headerIdx], 0) !== 'COMPANY NAME') headerIdx++;
          if (headerIdx < data.length) {
            const headers = ['Company Name','Mill Name','UML ID','Village','Sub District','District','Capacity','Lat','Long','Legality','ISPO','RSPO','ISCC','Total Supply CPO/PK (Ton)'];
            const tableRows = [];
            let j = headerIdx + 1;
            while (j < data.length) {
              const r = data[j];
              if (!r || !r.some(c => c !== null && c !== undefined && String(c).trim() !== '')) { j++; continue; }
              const v = cellVal(r, 0);
              if (v === 'B. FFB Supplier List' || isSectionHeader(r)) break;
              tableRows.push([cellVal(r,0),cellVal(r,1),cellVal(r,2),cellVal(r,3),cellVal(r,4),cellVal(r,5),cellVal(r,6),cellVal(r,7),cellVal(r,8),cellVal(r,9),cellVal(r,10),cellVal(r,11),cellVal(r,12),cellVal(r,13)]);
              j++;
            }
            html += tableCard('A. Mill List', headers, tableRows.filter(r => r[0] || r[1]));
            i = j;
          } else { i++; }
        }
        else if (v0 === 'B. FFB Supplier List') {
          i++;
          let headerIdx = i;
          while (headerIdx < data.length && cellVal(data[headerIdx], 0) !== 'MILL NAME') headerIdx++;
          if (headerIdx < data.length) {
            const hasSupplierGroupName =
              cellVal(data[headerIdx], 1).toUpperCase() === 'SUPPLIER GROUP NAME' ||
              (cellVal(data[headerIdx], 2).toUpperCase() === 'SUPPLIER NAME' && cellVal(data[headerIdx], 1).toUpperCase().indexOf('GROUP') !== -1);
            const headers = hasSupplierGroupName
              ? ['Mill Name','Supplier Group Name','Supplier Name','Village','Sub District','District','Category','Concession Area','Planted Area','# Smallholders','Planted Year','Legality','Lat','Long','ISPO','RSPO','ISCC','Total Supply FFB (Ton)']
              : ['Mill Name','Supplier Name','Village','Sub District','District','Category','Concession Area','Planted Area','# Smallholders','Planted Year','Legality','Lat','Long','ISPO','RSPO','ISCC','Total Supply FFB (Ton)'];
            const tableRows = [];
            let j = headerIdx + 1;
            const seenFfbCombos = new Set();
            while (j < data.length) {
              const r = data[j];
              if (!r || !r.some(c => c !== null && c !== undefined && String(c).trim() !== '')) { j++; continue; }
              const v = cellVal(r, 0);
              if (isSectionHeader(r)) break;
              // skip rows with only PLEASE SELECT
              const allPlease = r.every((c,idx) => !c || String(c).trim() === '' || String(c).trim() === 'PLEASE SELECT');
              if (allPlease) { j++; continue; }
              const supplierGroup = hasSupplierGroupName ? cellVal(r, 1) : '';
              const supplierName = cellVal(r, hasSupplierGroupName ? 2 : 1) || (hasSupplierGroupName ? '' : cellVal(r, 2));
              if (!supplierName) { j++; continue; }
              const millNameCell = cellVal(r, 0);
              // Deduplicate by logical supplier identity (not all columns),
              // so legacy duplicate rows in Sheets do not reappear in UI.
              const comboKey = hasSupplierGroupName
                ? [
                  normalizeCellText(millNameCell).toLowerCase(),
                  normalizeCellText(supplierGroup).toLowerCase(),
                  normalizeCellText(supplierName).toLowerCase(),
                  normalizeCellText(cellVal(r,3)).toLowerCase(),
                  normalizeCellText(cellVal(r,4)).toLowerCase(),
                  normalizeCellText(cellVal(r,5)).toLowerCase(),
                  normalizeCellText(cellVal(r,6)).toLowerCase()
                ].join('||')
                : [
                  normalizeCellText(millNameCell).toLowerCase(),
                  normalizeCellText(supplierName).toLowerCase(),
                  normalizeCellText(cellVal(r,2)).toLowerCase(),
                  normalizeCellText(cellVal(r,3)).toLowerCase(),
                  normalizeCellText(cellVal(r,4)).toLowerCase(),
                  normalizeCellText(cellVal(r,5)).toLowerCase()
                ].join('||');
              if (seenFfbCombos.has(comboKey)) { j++; continue; }
              seenFfbCombos.add(comboKey);
              if (hasSupplierGroupName) {
                tableRows.push([millNameCell, supplierGroup, supplierName, cellVal(r,3), cellVal(r,4), cellVal(r,5), cellVal(r,6), cellVal(r,7), cellVal(r,8), cellVal(r,9), cellVal(r,10), cellVal(r,11), cellVal(r,12), cellVal(r,13), cellVal(r,14), cellVal(r,15), cellVal(r,16), cellVal(r,17)]);
              } else {
                tableRows.push([millNameCell, supplierName, cellVal(r,2), cellVal(r,3), cellVal(r,4), cellVal(r,5), cellVal(r,6), cellVal(r,7), cellVal(r,8), cellVal(r,9), cellVal(r,10), cellVal(r,11), cellVal(r,12), cellVal(r,13), cellVal(r,14), cellVal(r,15), cellVal(r,16)]);
              }
              j++;
            }
            html += tableCard('B. FFB Supplier List', headers, tableRows.filter(r => hasSupplierGroupName ? r[2] : r[1]));
            i = j;
          } else { i++; }
        }
        else { i++; }
      }

      // collect mill names for traceability screening
      window._tmlMillNames = [];
      let _scanI = 0;
      while (_scanI < data.length) {
        const _sr = data[_scanI];
        const _sv = cellVal(_sr, 0);
        if (_sv === 'A. Mill List') {
          let _shi = _scanI + 1;
          while (_shi < data.length && cellVal(data[_shi], 0) !== 'COMPANY NAME') _shi++;
          let _sj = _shi + 1;
          while (_sj < data.length) {
            const _srow = data[_sj];
            if (!_srow || !_srow.some(c => c !== null && c !== undefined && String(c).trim() !== '')) { _sj++; continue; }
            if (cellVal(_srow, 0) === 'B. FFB Supplier List') break;
            const mn = cellVal(_srow, 1);
            if (mn) window._tmlMillNames.push(mn);
            _sj++;
          }
        }
        _scanI++;
      }

      // collect supplier names for FFB Supplier List screening
      window._ffbSupplierNames = [];
      let _scanF = 0;
      while (_scanF < data.length) {
        const _sfr = data[_scanF];
        const _sfv = cellVal(_sfr, 0);
        if (_sfv === 'B. FFB Supplier List') {
          let _fhi = _scanF + 1;
          while (_fhi < data.length && cellVal(data[_fhi], 0) !== 'MILL NAME') _fhi++;
          let _fj = _fhi + 1;
          while (_fj < data.length) {
            const _frow = data[_fj];
            if (!_frow || !_frow.some(c => c !== null && c !== undefined && String(c).trim() !== '')) { _fj++; continue; }
            if (isSectionHeader(_frow)) break;
            const hasSupplierGroupName =
              cellVal(data[_fhi], 1).toUpperCase() === 'SUPPLIER GROUP NAME' ||
              (cellVal(data[_fhi], 2).toUpperCase() === 'SUPPLIER NAME' && cellVal(data[_fhi], 1).toUpperCase().indexOf('GROUP') !== -1);
            const sn = cellVal(_frow, hasSupplierGroupName ? 2 : 1) || (hasSupplierGroupName ? '' : cellVal(_frow, 2));
            if (sn && sn !== 'PLEASE SELECT' && !window._ffbSupplierNames.includes(sn)) window._ffbSupplierNames.push(sn);
            _fj++;
          }
        }
        _scanF++;
      }

      // Screening action button: Start Screening (draft) vs View Screening (submitted)
      var _isSubmittedForBtn = (function() {
        var st = String((window._scrData && window._scrData.status) || (window._loadedPrimarySddRow && window._loadedPrimarySddRow['SCR - Screening Status']) || '').trim().toLowerCase();
        return st === 'submitted';
      })();
      html += '<div id="sdd-trace-action-btn-wrap" style="display:flex;justify-content:flex-end;flex-wrap:wrap;gap:10px;margin-top:16px;padding-right:4px;">'
        + '<button type="button" id="sdd-check-nbl-btn-trace" onclick="window.runSddNblCheck && window.runSddNblCheck()" style="padding:9px 18px;border-radius:8px;border:1.5px solid rgba(139,26,26,0.35);background:#fff;color:#8B1A1A;font-size:13px;font-weight:600;font-family:Inter,sans-serif;cursor:pointer;">Check NBL</button>'
        + (_isSubmittedForBtn
          ? '<button onclick="window.openViewScreeningPopup()" style="background:#1e40af;color:white;border:none;border-radius:8px;padding:9px 20px;font-size:13px;font-weight:600;font-family:Inter,sans-serif;cursor:pointer;display:flex;align-items:center;gap:8px;box-shadow:0 2px 8px rgba(30,64,175,0.25);" onmouseover="this.style.background=\'#1d4ed8\'" onmouseout="this.style.background=\'#1e40af\'">'
            + '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>'
            + ' View Screening</button>'
          : '<button onclick="window.openTmlScreeningPicker()" style="background:#8B1A1A;color:white;border:none;border-radius:8px;padding:9px 20px;font-size:13px;font-weight:600;font-family:Inter,sans-serif;cursor:pointer;display:flex;align-items:center;gap:8px;box-shadow:0 2px 8px rgba(139,26,26,0.25);" onmouseover="this.style.background=\'#A52020\'" onmouseout="this.style.background=\'#8B1A1A\'">'
            + '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><polyline points="15 11 11 11 11 7"/><line x1="16.5" y1="16.5" x2="22" y2="22"/></svg>'
            + ' Start Screening</button>')
        + '</div>'
        + '<div id="sdd-trace-bottom-card" style="margin-top:18px;padding:18px 16px 16px;border:1px solid rgba(139,26,26,0.12);border-radius:12px;background:#fff;">'
        + '<div style="display:flex;flex-direction:column;gap:14px;">'
        + '<div style="display:flex;gap:16px;align-items:stretch;">'
        + '<div style="flex:1;display:flex;flex-direction:column;gap:8px;">'
        + '<label for="traceRecInput" style="font-size:13px;font-weight:600;color:#1A0A0A;">Note</label>'
        + '<textarea id="traceRecInput" rows="4" placeholder="Note" style="flex:1;width:100%;min-height:120px;border:1px solid rgba(74,28,28,0.15);box-shadow:0 2px 10px rgba(0,0,0,0.04);border-radius:14px;padding:14px 16px;font-size:13px;color:#1A0A0A;line-height:1.5;resize:vertical;outline:none;box-sizing:border-box;"></textarea>'
        + '</div>'
        + '<div style="flex:1;display:flex;flex-direction:column;gap:8px;">'
        + '<label for="requestedDataInput" style="font-size:13px;font-weight:600;color:#1A0A0A;">Requested Data</label>'
        + '<textarea id="requestedDataInput" rows="4" placeholder="Requested Data" style="flex:1;width:100%;min-height:120px;border:1px solid rgba(74,28,28,0.15);box-shadow:0 2px 10px rgba(0,0,0,0.04);border-radius:14px;padding:14px 16px;font-size:13px;color:#1A0A0A;line-height:1.5;resize:vertical;outline:none;box-sizing:border-box;"></textarea>'
        + '</div>'
        + '</div>'
        + '<div style="display:flex;flex-direction:column;gap:10px;">'
        + '<label for="traceAttachmentInput" style="font-size:13px;font-weight:600;color:#1A0A0A;">Attach photo / PDF link</label>'
        + '<div style="display:flex;gap:10px;flex-wrap:wrap;">'
        + '<input id="traceAttachmentInput" type="url" placeholder="Paste Drive / image / PDF link here" style="flex:1;min-width:0;border:1px solid rgba(74,28,28,0.15);border-radius:12px;padding:12px 14px;background:#f9fafb;color:#111;font-size:13px;outline:none;">'
        + '<button type="button" onclick="window._addTraceAttachmentLink && window._addTraceAttachmentLink()" style="padding:9px 20px;border-radius:8px;border:none;background:#1F2937;color:white;cursor:pointer;font-size:13px;font-weight:600;font-family:Inter,sans-serif;box-shadow:0 2px 6px rgba(0,0,0,0.15);">Add link</button>'
        + '</div>'
        + '<div style="display:flex;justify-content:flex-end;gap:10px;flex-wrap:wrap;">'
        + '<button type="button" data-sdd-save="draft" onclick="window._saveScrScreening(\'draft\')" style="padding:9px 20px;border-radius:8px;border:none;background:#F59E0B;color:white;font-family:Inter,sans-serif;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 2px 8px rgba(245,158,11,0.3);">Save as Draft</button>'
        + '<button type="button" data-sdd-save="submit" onclick="window._saveScrScreening(\'submit\')" style="padding:9px 20px;border-radius:8px;border:none;background:#10B981;color:white;font-family:Inter,sans-serif;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 2px 8px rgba(16,185,129,0.3);">Submit</button>'
        + '</div>'
        + '<div style="font-size:12px;color:#6b7280;">Use Google Drive share link or direct image/PDF URL.</div>'
        + '<div id="traceAttachmentPreview" style="display:grid;gap:12px;color:#6b7280;font-size:13px;">No attachment yet.</div>'
        + '</div>'
        + '<div class="sdd-hide-for-approver" style="margin-top:20px;padding-top:16px;border-top:1px solid rgba(74,28,28,0.08);">'
        + '<div style="display:flex;align-items:center;gap:6px;margin-bottom:8px;">'
        + '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#8B1A1A" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>'
        + '<label for="noteBossDecision" style="font-size:11px;font-weight:700;color:#8B1A1A;letter-spacing:0.08em;text-transform:uppercase;cursor:pointer;">Approver Note <span style="font-weight:400;color:#9C8080;text-transform:none;letter-spacing:0;">(optional — saved with Draft / Submit)</span></label>'
        + '</div>'
        + '<textarea id="noteBossDecision" rows="3" placeholder="Tulis catatan untuk approver di sini..." style="width:100%;min-height:80px;border:1px solid rgba(139,26,26,0.18);border-radius:10px;padding:11px 14px;font-size:13px;color:#1A0A0A;line-height:1.5;resize:vertical;outline:none;box-sizing:border-box;font-family:Inter,sans-serif;background:#fff;box-shadow:0 1px 4px rgba(139,26,26,0.06);"></textarea>'
        + '<div id="sdd-staff-decision-wrap" style="display:none;margin-top:14px;padding-top:14px;border-top:1px dashed rgba(139,26,26,0.22);">'
        + '<div style="font-size:11px;font-weight:700;color:#8B1A1A;letter-spacing:0.07em;text-transform:uppercase;margin-bottom:10px;">Keputusan screening</div>'
        + '<div style="display:flex;flex-wrap:wrap;gap:10px;align-items:center;">'
        + '<span id="sdd-staff-decision-badge" style="font-size:12px;font-weight:600;color:#5F4A48;min-height:1.2em;flex:1;min-width:140px;">Decision: —</span>'
        + '<button type="button" id="sdd-approver-approve" disabled style="padding:8px 16px;border-radius:8px;border:none;background:#047857;color:#fff;font-size:12px;font-weight:600;cursor:pointer;font-family:Inter,sans-serif;opacity:0.5;">Approve</button>'
        + '<button type="button" id="sdd-approver-hold" disabled style="padding:8px 16px;border-radius:8px;border:none;background:#D97706;color:#fff;font-size:12px;font-weight:600;cursor:pointer;font-family:Inter,sans-serif;opacity:0.5;">On Hold</button>'
        + '<button type="button" id="sdd-approver-reject" disabled style="padding:8px 16px;border-radius:8px;border:1.5px solid #B91C1C;background:#fff;color:#B91C1C;font-size:12px;font-weight:600;cursor:pointer;font-family:Inter,sans-serif;opacity:0.5;">Reject</button>'
        + '</div></div>'
        + '<div id="sdd-bottom-actions-row" style="display:flex;justify-content:flex-end;flex-wrap:wrap;gap:10px;margin-top:14px;align-items:center;">'
        + '<button type="button" data-sdd-save="delete" onclick="window._saveScrScreening(\'delete\')" style="padding:9px 20px;border-radius:8px;border:1.5px solid rgba(239,68,68,0.4);background:#fff;color:#EF4444;font-family:Inter,sans-serif;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 1px 4px rgba(239,68,68,0.08);margin-right:auto;">Delete</button>'
        + '<button type="button" id="sdd-cancel-to-draft-btn" data-sdd-save="cancel" style="display:none;padding:9px 20px;border-radius:8px;border:1.5px solid rgba(217,119,6,0.4);background:#fff7ed;color:#9a3412;font-family:Inter,sans-serif;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 1px 4px rgba(217,119,6,0.12);">Cancel</button>'
        + '</div>'
        + '</div>'
        + '<p id="scr-save-ok" style="display:none;margin-top:12px;text-align:right;font-size:13px;color:#059669;font-weight:600;"></p>'
        + '</div>'
        + '</div>';

      container.innerHTML = html || '<p style="color:#9C8080;font-size:13px;padding:16px;">No traceability data found.</p>';
      if (typeof window.refreshSddApproverStaffTraceChrome === 'function') window.refreshSddApproverStaffTraceChrome();
      if (window._scrData) {
        const noteEl = document.getElementById('traceRecInput');
        if (noteEl) noteEl.value = window._scrData.traceNote || '';
        const reqDataEl = document.getElementById('requestedDataInput');
        if (reqDataEl) reqDataEl.value = window._scrData.requestedData || '';
        window._traceAttachments = window._scrData.attachments || [];
      }
      window._renderTraceAttachments && window._renderTraceAttachments();
      window.updateTraceRecButtons = function(el) {
        document.querySelectorAll('[data-trace-rec-button]').forEach(function(btn) {
          btn.style.background = '#fff';
          btn.style.borderColor = '#E8E3E3';
          btn.style.color = '#1A0A0A';
        });
        if (el) {
          el.style.background = '#8B1A1A';
          el.style.borderColor = '#8B1A1A';
          el.style.color = '#fff';
        }
      };
      if (!window._traceRecChoice) window._traceRecChoice = 'Recommend';
      window.updateTraceRecButtons(document.querySelector('[data-trace-rec-button][data-action="Recommend"]'));
      return;
    }

    // ── MAIN FORM SHEET ──
    let html = '';
    let currentSection = '';
    let sectionInner = '';

    // Track mode: normal KV, certification, legality
    let mode = 'kv';
    let certRows = [];    // accumulate cert rows for table rendering

    function flushSection() {
      if (mode === 'cert' && certRows.length) {
        // Render cert as KV profile style, one block per certification
        let certHtml = '';
        certRows.forEach(function(r) {
          // r = [name, status, certNum, start, end, body]
          const statusColor = r[1] === 'Yes' ? '#1a6b3a' : r[1] === 'No' ? '#b03030' : '#5F4A48';
          const statusBg   = r[1] === 'Yes' ? 'rgba(30,107,58,0.08)' : r[1] === 'No' ? 'rgba(176,48,48,0.07)' : 'rgba(95,74,72,0.06)';
          certHtml += '<div style="border:1px solid rgba(74,28,28,0.09);border-radius:9px;padding:14px 18px;margin-bottom:10px;background:#fdfafa;">';
          // cert name as sub-header
          certHtml += '<div style="font-size:12px;font-weight:700;color:#2C2828;letter-spacing:0.5px;margin-bottom:10px;padding-bottom:8px;border-bottom:1px solid rgba(74,28,28,0.07);">' + r[0] + '</div>';
          certHtml += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px 20px;">';
          // Status
          certHtml += '<div><div style="font-size:10px;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;color:#9C8080;margin-bottom:3px;">Status</div>'
            + '<span style="display:inline-block;padding:2px 10px;border-radius:20px;font-size:12px;font-weight:600;background:' + statusBg + ';color:' + statusColor + ';">' + r[1] + '</span></div>';
          // Cert No
          certHtml += '<div><div style="font-size:10px;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;color:#9C8080;margin-bottom:3px;">No. Sertifikat</div>'
            + '<div style="font-size:13px;color:#1A0A0A;">' + r[2] + '</div></div>';
          // Validity
          certHtml += '<div><div style="font-size:10px;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;color:#9C8080;margin-bottom:3px;">Validity Start</div>'
            + '<div style="font-size:13px;color:#1A0A0A;">' + r[3] + '</div></div>';
          certHtml += '<div><div style="font-size:10px;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;color:#9C8080;margin-bottom:3px;">Validity End</div>'
            + '<div style="font-size:13px;color:#1A0A0A;">' + r[4] + '</div></div>';
          // Cert body — full width
          certHtml += '<div style="grid-column:1/-1;"><div style="font-size:10px;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;color:#9C8080;margin-bottom:3px;">Certification Body</div>'
            + '<div style="font-size:13px;color:#1A0A0A;">' + r[5] + '</div></div>';
          certHtml += '</div></div>';
        });
        sectionInner = certHtml;
        certRows = [];
      }
      if (!sectionInner) return;
      html += sectionCard(currentSection, sectionInner);
      sectionInner = '';
    }

    // Some templates split a field across 2 rows, e.g.:
    //   Distance to Refinery : Please Select
    //   150 KM               :
    // We treat the second row as the value of the first.
    let pendingDistanceRow = null; // { label, value }

    for (let _ri = 0; _ri < data.length; _ri++) {
      const row = data[_ri];
      if (!row) continue;
      const col0 = cellVal(row, 0);

      if (col0 === 'MILL SUPPLIER REGISTRATION FORM') continue;

      // Section header
      if (isSectionHeader(row)) {
        // flush any pending distance row before closing section
        if (pendingDistanceRow) {
          sectionInner += kv(pendingDistanceRow.label, pendingDistanceRow.value);
          pendingDistanceRow = null;
        }
        flushSection();
        currentSection = col0;
        mode = col0.startsWith('C.') ? 'cert' : col0.startsWith('D.') ? 'legal' : 'kv';
        continue;
      }

      // Allow empty col0 rows through in kv mode (product sub-rows), but skip in cert/legal
      if (!col0 && mode !== 'kv') continue;
      // In kv mode, skip truly empty rows (no content anywhere)
      // Note: some templates place extra fields on the RIGHT side (label at col4, value at col6)
      // with empty col0/col2. Keep those rows so they can render as secondary KV pairs.
      if (!col0 && mode === 'kv' && !cellVal(row, 2) && !cellVal(row, 4) && !cellVal(row, 5) && !cellVal(row, 6)) continue;

      if (mode === 'cert') {
        // Skip the header row "Certification Members | Start | End | Certification Body"
        if (col0 === 'Certification Members') continue;
        // col0=name, col1=yes/no, col3=cert number, col6=start, col7=end, col8=body
        const status = cellVal(row, 1);
        const certNum = cellVal(row, 3);
        const start = cellVal(row, 6);
        const end = cellVal(row, 7);
        const body = cellVal(row, 8);
        certRows.push([col0.trim(), status||'—', certNum||'—', start||'—', end||'—', body||'—']);
      }
      else if (mode === 'legal') {
        // col0=name, col1=type(HGU/IUP), col3=doc num (col2="Document Num" label), col6=issue date (col5="Issue Date" label)
        const type = cellVal(row, 1);
        const docNum = cellVal(row, 3);
        const issueDate = cellVal(row, 6);
        // Build multi-line value: type on first line, then No:, then Issue Date:
        let lines = [];
        if (type) lines.push('<span style="font-weight:600;color:#1A0A0A;">' + type + '</span>');
        if (docNum) lines.push('<span style="color:#5F4A48;">No: <span style="color:#1A0A0A;">' + docNum + '</span></span>');
        if (issueDate) lines.push('<span style="color:#5F4A48;">Issue Date: <span style="color:#1A0A0A;">' + issueDate + '</span></span>');
        const val = lines.length ? lines.join('<br>') : '—';
        sectionInner += kv(col0, val || '—');
      }
      else {
        // Normal KV: col0=label, col2=value
        // Special case: "Product to be Produce" header row — skip, handled by sub-table accumulation
        const k2 = cellVal(row, 4);
        const v0 = cellVal(row, 2);
        const v6 = cellVal(row, 6);

        // If we had a pending distance row and this row is NOT its continuation, render it now.
        if (pendingDistanceRow) {
          const isKmRow = !!(col0 && /^\s*\d+(\.\d+)?\s*km\s*$/i.test(col0) && !v0 && !k2 && !cellVal(row, 5) && !cellVal(row, 6));
          if (!isKmRow) {
            sectionInner += kv(pendingDistanceRow.label, pendingDistanceRow.value);
            pendingDistanceRow = null;
          }
        }

        // Detect product-table header row — supports both label variants:
        //   "Product to be Supply"  (Mill form — col2 = "Main Product")
        //   "Product to be Produce" (KCP  form — col2 = "Product")
        const isProductHeader =
          (col0 === 'Product to be Supply' || col0 === 'Product to be Produce') &&
          (v0 === 'Main Product' || v0 === 'Product');
        if (isProductHeader) {
          // Dual-column (Mill) layout has 'Other Product' at col5 of the header row
          const hasOtherCol = (cellVal(row, 5) || '').toLowerCase().indexOf('other') !== -1;
          sectionInner += '<div style="padding:8px 0 6px;border-bottom:1px solid rgba(74,28,28,0.05);">'
            + '<span style="color:#5F4A48;font-weight:600;font-size:13.5px;">' + col0 + '</span></div>';
          if (hasOtherCol) {
            // Mill: dual-column header (Main Product | Avg | Yield | Other Product | Avg)
            sectionInner += '<div style="display:grid;grid-template-columns:130px 110px 60px 130px 110px;gap:0;padding:6px 0 4px;border-bottom:1px solid rgba(74,28,28,0.08);">'
              + '<span style="font-size:10px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;color:#9C8080;padding:0 8px 0 0;">Main Product</span>'
              + '<span style="font-size:10px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;color:#9C8080;padding:0 8px;">Avg/Month</span>'
              + '<span style="font-size:10px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;color:#9C8080;padding:0 8px;">Yield</span>'
              + '<span style="font-size:10px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;color:#9C8080;padding:0 8px;">Other Product</span>'
              + '<span style="font-size:10px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;color:#9C8080;padding:0 8px;">Avg/Month</span>'
              + '</div>';
          } else {
            // KCP: single-column header (Product | Avg | Yield)
            sectionInner += '<div style="display:grid;grid-template-columns:160px 130px 80px;gap:0;padding:6px 0 4px;border-bottom:1px solid rgba(74,28,28,0.08);">'
              + '<span style="font-size:10px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;color:#9C8080;padding:0 8px 0 0;">Product</span>'
              + '<span style="font-size:10px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;color:#9C8080;padding:0 8px;">Avg/Month</span>'
              + '<span style="font-size:10px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;color:#9C8080;padding:0 8px;">Yield</span>'
              + '</div>';
          }
          continue;
        }
        // Detect CPO/PK/CPKO/PKE Quality row (Mill CPO/PK; KCP CPKO/PKE — same column layout)
        if ((col0 === 'CPO Quality' || col0 === 'CPKO Quality' || col0 === 'PK Quality' || col0 === 'PKE Quality') && cellVal(row, 2) === 'FFA') {
          const qa = [
            ['FFA', cellVal(row, 3)],
            [cellVal(row, 4), cellVal(row, 5)],
            [cellVal(row, 6), cellVal(row, 7)],
          ].filter(p => p[0]);
          const qHtml = '<div style="display:flex;gap:20px;flex-wrap:wrap;">'
            + qa.map(p => '<span style="font-size:13px;"><span style="color:#5F4A48;font-weight:600;">' + p[0] + ':</span> <span style="color:#1A0A0A;">' + (p[1]||'—') + '</span></span>').join('')
            + '</div>';
          sectionInner += kv(col0, qHtml);
          continue;
        }

        // Detect product data rows (col0 empty, col2=product name OR col5=other product)
        if (!col0 && (v0 || cellVal(row, 5))) {
          const mainProd = v0;
          const mainAvg = cellVal(row, 3);
          const yld = cellVal(row, 4);
          const otherProd = cellVal(row, 5);
          const otherAvg = cellVal(row, 6);
          // skip rows that are entirely PLEASE SELECT
          if (!mainProd && !mainAvg && !otherProd && !otherAvg) return;
          // parse yield
          let yldStr = '—';
          if (yld && yld !== '' && yld !== 'PLEASE SELECT') {
            const ynum = parseFloat(yld);
            yldStr = !isNaN(ynum) ? (ynum < 1 ? (ynum * 100).toFixed(0) + '%' : yld) : yld;
          }
          let rowHtml = '<div style="display:grid;grid-template-columns:130px 110px 60px 130px 110px;gap:0;padding:5px 0;border-bottom:1px solid rgba(74,28,28,0.04);align-items:center;">';
          rowHtml += '<span style="font-size:13px;color:#1A0A0A;padding:0 8px 0 0;">' + (mainProd || '') + '</span>';
          rowHtml += '<span style="font-size:13px;color:#1A0A0A;padding:0 8px;">' + (mainProd ? (mainAvg || '—') : '') + '</span>';
          rowHtml += '<span style="font-size:12px;color:#7C6565;padding:0 8px;">' + (mainProd ? yldStr : '') + '</span>';
          rowHtml += '<span style="font-size:13px;color:#1A0A0A;padding:0 8px;">' + (otherProd || '') + '</span>';
          rowHtml += '<span style="font-size:13px;color:#1A0A0A;padding:0 8px;">' + (otherProd ? (otherAvg || '—') : '') + '</span>';
          rowHtml += '</div>';
          sectionInner += rowHtml;
          continue;
        }

        // Distance to Refinery can carry 2 different values:
        // - selector value near the label (e.g. "Please Select"/chosen refinery)
        // - physical distance in KM that may appear in various right-side columns
        if (col0 === 'Distance to Refinery') {
          const rightCandidates = [];
          for (let dc = 1; dc <= 12; dc++) rightCandidates.push(cellVal(row, dc));
          const distanceValue = (function() {
            for (let i = 0; i < rightCandidates.length; i++) {
              const s = String(rightCandidates[i] || '').trim();
              if (!s) continue;
              if (/^\d+([.,]\d+)?$/.test(s)) return s.replace(',', '.') + ' KM';
              if (/^\s*\d+(\.\d+)?\s*km\s*$/i.test(s)) return s.replace(/\s+/g, ' ').trim().toUpperCase().replace('KM', ' KM');
            }
            return '';
          })();
          if (distanceValue) {
            sectionInner += kv(col0, v0 || '—');
            sectionInner += kv('Distance', distanceValue);
            pendingDistanceRow = null;
            continue;
          }
        }
        // Special 2-row Distance to Refinery case
        if (col0 === 'Distance to Refinery' && (!v0 || String(v0).trim().toLowerCase() === 'please select')) {
          pendingDistanceRow = { label: col0, value: v0 || '' };
          continue;
        }
        // Continuation row for pending distance: "150 KM" placed as a standalone first-cell row
        if (pendingDistanceRow && col0 && /^\s*\d+(\.\d+)?\s*km\s*$/i.test(col0) && !v0 && !k2 && !cellVal(row, 5) && !cellVal(row, 6)) {
          sectionInner += kv(pendingDistanceRow.label, col0);
          pendingDistanceRow = null;
          continue;
        }

        // Rows with empty col0 but populated right-side label/value (k2/v6)
        if (!col0) {
          if (k2 && !/^(YIELD|Average|Other Product|Product)$/i.test(k2)) {
            sectionInner += kv(k2, cellVal(row, 6));
          }
          continue;
        }

        sectionInner += kv(col0, v0);
        // Suppress the second header column in product-table rows (col4 = YIELD /
        // Average Potential / Month / Other Product) from appearing as extra KV pairs.
        if (k2 && !/^(YIELD|Average|Other Product|Product)$/i.test(k2)) {
          sectionInner += kv(k2, cellVal(row, 6));
        }
      }
    }

    // Flush any pending distance row at end of sheet
    if (pendingDistanceRow) {
      sectionInner += kv(pendingDistanceRow.label, pendingDistanceRow.value);
      pendingDistanceRow = null;
    }

    flushSection();

    // ── screening form: after section D (Legality), before any following content (e.g. traceability in another block) ──
    if (opts.appendScrForm !== false) {
      html += buildScrForm();
    }

    container.innerHTML = html || '<p style="color:#9C8080;font-size:13px;padding:16px;">No data found.</p>';
    if (opts.appendScrForm !== false) {
      initScrForm();
      // Fresh Excel import is Draft: decision buttons must stay hidden
      // until SCR is Submitted.
      // Use setTimeout(0) so all synchronous initScrForm internal calls finish first,
      // and the DOM is fully flushed before we try to show sdd-staff-decision-wrap.
      setTimeout(function() {
        var wrap = document.getElementById('sdd-staff-decision-wrap');
        if (wrap) wrap.style.display = 'none';
      }, 0);
    }
  }

  // ─── SUPPLIER SCREENING BUILDER ─────────────────────────────────────────────
  function buildScrForm() {
    const s = window._scrData || {};

    function sel(id, opts, val) {
      let cls = 'scr-sel';
      if (val==='Yes'||val==='Complete') cls += ' scr-sel-yes';
      else if (val==='No'||val==='Non Complete') cls += ' scr-sel-no';
      return `<select id="${id}" class="${cls}">
        <option value="">— Pilih —</option>
        ${opts.map(v=>`<option${val===v?' selected':''}>${v}</option>`).join('')}
      </select>`;
    }
    function grvRowHtml(r) {
      r=r||{};
      return `<tr>
        <td><input type="text" class="grv-source" value="${(r.source||'').replace(/"/g,'&quot;')}" placeholder="Source"></td>
        <td><input type="text" class="grv-desc"   value="${(r.desc||'').replace(/"/g,'&quot;')}" placeholder="Description"></td>
        <td><input type="text" class="grv-pub"    value="${(r.pub||'').replace(/"/g,'&quot;')}" placeholder="Publisher"></td>
        <td><input type="text" class="grv-date"   value="${(r.date||'').replace(/"/g,'&quot;')}" placeholder="YYYY-MM-DD"></td>
        <td><input type="text" class="grv-status" value="${(r.status||'').replace(/"/g,'&quot;')}" placeholder="Status"></td>
        <td><input type="text" class="grv-attach" value="${(r.attach||'').replace(/"/g,'&quot;')}" placeholder="Link/filename"></td>
        <td><button type="button" class="scr-del-row">✕</button></td>
      </tr>`;
    }
    function priRowHtml(r) {
      r=r||{};
      return `<tr>
        <td><input type="text" class="pri-company" value="${(r.company||'').replace(/"/g,'&quot;')}" placeholder="Company"></td>
        <td><input type="text" class="pri-desc"    value="${(r.desc||'').replace(/"/g,'&quot;')}" placeholder="Description"></td>
        <td><input type="text" class="pri-pub"     value="${(r.pub||'').replace(/"/g,'&quot;')}" placeholder="Publisher"></td>
        <td><input type="text" class="pri-date"    value="${(r.date||'').replace(/"/g,'&quot;')}" placeholder="YYYY-MM-DD"></td>
        <td><input type="text" class="pri-attach"  value="${(r.attach||'').replace(/"/g,'&quot;')}" placeholder="Link/filename"></td>
        <td><input type="text" class="pri-action"  value="${(r.action||'').replace(/"/g,'&quot;')}" placeholder="Action Request"></td>
        <td><button type="button" class="scr-del-row">✕</button></td>
      </tr>`;
    }

    const grvYN  = s.grvYN||'';
    const priYN  = s.priYN||'';
    const grvRows = (s.grvRows&&s.grvRows.length) ? s.grvRows : [{}];
    const priRows = (s.priRows&&s.priRows.length) ? s.priRows : [{}];

    return `
    <div style="background:linear-gradient(135deg,rgba(139,26,26,0.03),rgba(139,26,26,0.06));border:1.5px solid rgba(139,26,26,0.18);border-radius:14px;padding:22px 26px;margin-top:14px;" id="scr-form-wrap">
      <div style="font-size:11px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;color:#8B1A1A;margin-bottom:18px;display:flex;align-items:center;gap:8px;">
        <span style="display:inline-block;width:3px;height:16px;background:#8B1A1A;border-radius:2px;"></span>
        Supplier Screening
      </div>

      <div id="sdd-nbl-check-section" style="margin-bottom:18px;padding:14px 16px;border:1px solid rgba(139,26,26,0.12);border-radius:10px;background:#fff;">
        <div style="display:flex;flex-wrap:wrap;align-items:flex-start;gap:12px;justify-content:space-between;">
          <div style="flex:1;min-width:200px;">
            <div style="font-size:12px;font-weight:700;color:#1A0A0A;">No Buy List check</div>
            <div style="font-size:12px;color:#6b7280;margin-top:4px;line-height:1.45;">Flags <strong>Yes</strong> if any imported name matches NBL (Group or Company) or Unilever NBL (Company or Mill). One similar name is enough.</div>
          </div>
          <button type="button" id="sdd-check-nbl-btn" onclick="window.runSddNblCheck && window.runSddNblCheck()" style="flex-shrink:0;padding:9px 18px;border-radius:8px;border:none;background:#8B1A1A;color:#fff;font-size:13px;font-weight:600;font-family:Inter,sans-serif;cursor:pointer;box-shadow:0 2px 8px rgba(139,26,26,0.2);">Check NBL</button>
        </div>
        <div id="sdd-nbl-check-result" class="sdd-nbl-check-result" style="display:none;" role="status" aria-live="polite"></div>
      </div>

      <div style="font-size:10px;font-weight:700;letter-spacing:1.1px;text-transform:uppercase;color:#8B1A1A;margin-bottom:12px;padding:6px 12px;background:rgba(139,26,26,0.07);border-radius:6px;">A. Company Checking</div>
      <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px 20px;margin-bottom:18px;">
        <div class="scr-field"><div class="scr-lbl">List Group by Owners</div>
          <textarea id="scr-owners" class="scr-ta" rows="2" placeholder="Isi nama group/owners...">${s.owners||''}</textarea></div>
        <div class="scr-field"><div class="scr-lbl">Previous News</div>
          <textarea id="scr-news" class="scr-ta" rows="2" placeholder="Isi berita terdahulu...">${s.news||''}</textarea></div>
        <div class="scr-field"><div class="scr-lbl">Supply To</div>
          <textarea id="scr-supplyto" class="scr-ta" rows="2" placeholder="Isi supply to...">${s.supplyto||''}</textarea></div>
        <div class="scr-field"><div class="scr-lbl">Legality</div>
          ${sel('scr-legality',['Complete','Non Complete'],s.legality||'')}</div>
        <div class="scr-field"><div class="scr-lbl">Certification</div>
          <textarea id="scr-cert" class="scr-ta" rows="2" placeholder="Isi sertifikasi...">${s.cert||''}</textarea></div>
      </div>

      <hr style="border:none;border-top:1.5px solid rgba(139,26,26,0.09);margin:18px 0;">

      <div style="font-size:10px;font-weight:700;letter-spacing:1.1px;text-transform:uppercase;color:#8B1A1A;margin-bottom:12px;padding:6px 12px;background:rgba(139,26,26,0.07);border-radius:6px;">B. Non Spatial Analysis</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px 20px;margin-bottom:14px;">
        <div class="scr-field"><div class="scr-lbl">NDPE Policy</div>
          ${sel('scr-ndpe',['Yes','No'],s.ndpe||'')}</div>
        <div class="scr-field"><div class="scr-lbl">No Buy List</div>
          ${sel('scr-nbl',['Yes','No'],s.nbl||'')}</div>
        <div class="scr-field"><div class="scr-lbl">Grievance Published by Other Buyer</div>
          ${sel('scr-grv-yn',['Yes','No'],grvYN)}</div>
        <div class="scr-field"><div class="scr-lbl">Public Reporting Information (PRI)</div>
          ${sel('scr-pri-yn',['Yes','No'],priYN)}</div>
      </div>

      <div id="scr-grv-wrap" style="display:${grvYN==='Yes'?'block':'none'};margin-bottom:14px;">
        <div style="font-size:10px;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;color:#8B1A1A;margin-bottom:8px;padding:5px 10px;background:rgba(139,26,26,0.05);border-radius:5px;">Grievance Details</div>
        <div style="overflow-x:auto;-webkit-overflow-scrolling:touch;">
        <table style="border-collapse:collapse;font-size:12px;white-space:nowrap;width:100%;" id="scr-grv-tbl">
          <thead><tr style="background:rgba(139,26,26,0.07);">
            <th class="scr-th" style="min-width:130px">Source</th>
            <th class="scr-th" style="min-width:200px">Description</th>
            <th class="scr-th" style="min-width:130px">Publisher</th>
            <th class="scr-th" style="min-width:120px">Date Publish</th>
            <th class="scr-th" style="min-width:110px">Status</th>
            <th class="scr-th" style="min-width:160px">Attachment</th>
            <th class="scr-th" style="width:36px"></th>
          </tr></thead>
          <tbody id="scr-grv-tbody">${grvRows.map(r=>grvRowHtml(r)).join('')}</tbody>
        </table></div>
        <button type="button" class="scr-add-btn" id="scr-grv-add" title="Maksimal ${SCR_GRV_PRI_MAX_ROWS} baris">+ Add Row</button>
      </div>

      <div id="scr-pri-wrap" style="display:${priYN==='Yes'?'block':'none'};margin-bottom:14px;">
        <div style="font-size:10px;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;color:#8B1A1A;margin-bottom:8px;padding:5px 10px;background:rgba(139,26,26,0.05);border-radius:5px;">PRI Details</div>
        <div style="overflow-x:auto;-webkit-overflow-scrolling:touch;">
        <table style="border-collapse:collapse;font-size:12px;white-space:nowrap;width:100%;" id="scr-pri-tbl">
          <thead><tr style="background:rgba(139,26,26,0.07);">
            <th class="scr-th" style="min-width:150px">Company</th>
            <th class="scr-th" style="min-width:200px">Description</th>
            <th class="scr-th" style="min-width:130px">Publisher</th>
            <th class="scr-th" style="min-width:120px">Date Publish</th>
            <th class="scr-th" style="min-width:160px">Attachment</th>
            <th class="scr-th" style="min-width:180px">Action Request</th>
            <th class="scr-th" style="width:36px"></th>
          </tr></thead>
          <tbody id="scr-pri-tbody">${priRows.map(r=>priRowHtml(r)).join('')}</tbody>
        </table></div>
        <button type="button" class="scr-add-btn" id="scr-pri-add" title="Maksimal ${SCR_GRV_PRI_MAX_ROWS} baris">+ Add Row</button>
      </div>
    </div>

    <style>
      .scr-field{display:flex;flex-direction:column;gap:4px;}
      .scr-lbl{font-size:10px;font-weight:700;letter-spacing:0.8px;text-transform:uppercase;color:#B09A9A;}
      .scr-ta{width:100%;padding:9px 12px;border:1.5px solid rgba(74,28,28,0.12);border-radius:8px;font-family:Inter,sans-serif;font-size:13px;color:#1A0A0A;background:white;resize:vertical;outline:none;transition:border-color 0.2s;}
      .scr-ta:focus{border-color:#8B1A1A;box-shadow:0 0 0 3px rgba(139,26,26,0.07);}
      .scr-sel{width:100%;padding:9px 30px 9px 12px;border:1.5px solid rgba(74,28,28,0.12);border-radius:8px;font-family:Inter,sans-serif;font-size:13px;color:#1A0A0A;background:white url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%238B3A3A' stroke-width='2.5'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E") no-repeat right 10px center;appearance:none;outline:none;cursor:pointer;transition:border-color 0.2s;}
      .scr-sel:focus{border-color:#8B1A1A;box-shadow:0 0 0 3px rgba(139,26,26,0.07);}
      .scr-sel-yes{border-color:rgba(30,107,58,0.4)!important;background-color:rgba(30,107,58,0.04)!important;color:#1e6b3a!important;}
      .scr-sel-no{border-color:rgba(192,57,43,0.4)!important;background-color:rgba(192,57,43,0.04)!important;color:#c0392b!important;}
      .scr-th{padding:8px 12px;text-align:left;color:#8B1A1A;font-weight:700;font-size:10px;letter-spacing:0.6px;text-transform:uppercase;border-bottom:1.5px solid rgba(139,26,26,0.12);}
      #scr-grv-tbl td,#scr-pri-tbl td{padding:7px 8px;border-bottom:1px solid rgba(74,28,28,0.07);vertical-align:top;}
      #scr-grv-tbl td input,#scr-pri-tbl td input{width:100%;padding:5px 8px;border:1px solid rgba(74,28,28,0.12);border-radius:5px;font-family:Inter,sans-serif;font-size:12px;color:#1A0A0A;background:white;outline:none;}
      #scr-grv-tbl td input:focus,#scr-pri-tbl td input:focus{border-color:#8B1A1A;}
      .scr-del-row{background:none;border:none;cursor:pointer;color:#c0392b;font-size:14px;padding:2px 5px;border-radius:4px;}
      .scr-del-row:hover{background:rgba(192,57,43,0.1);}
      .scr-add-btn{margin-top:8px;padding:5px 14px;background:rgba(139,26,26,0.07);border:1.5px solid rgba(139,26,26,0.2);border-radius:6px;color:#8B1A1A;font-size:12px;font-weight:600;font-family:Inter,sans-serif;cursor:pointer;transition:background 0.15s;}
      .scr-add-btn:hover{background:rgba(139,26,26,0.14);}
      #scr-saved-select option{color:#111;}
    </style>`;
  }

  window._scrSavedRowsByKey = window._scrSavedRowsByKey || {};
  window._scrSavedGroupsByKey = window._scrSavedGroupsByKey || {};
  window._pendingScrLoadKey = '';

  /** Trace chrome: single staff role (no separate approver account). */
  window.refreshSddApproverStaffTraceChrome = function() {
    document.querySelectorAll('.sdd-hide-for-approver').forEach(function(el) {
      if (el) el.style.display = '';
    });
    ['traceRecInput', 'requestedDataInput'].forEach(function(id) {
      var el = document.getElementById(id);
      if (!el) return;
      el.readOnly = false;
      el.style.background = '';
      el.style.color = '#1A0A0A';
      el.style.cursor = '';
    });
    var startBtn = document.querySelector('#sdd-trace-action-btn-wrap button[onclick*="openTmlScreeningPicker"]');
    if (startBtn) startBtn.style.display = '';

    // Boss decision panel — disabled, always hidden
    var bossPanel = document.getElementById('sdd-boss-decision-panel');
    if (bossPanel) bossPanel.style.display = 'none';
    var bossViewer = document.getElementById('sdd-boss-decision-viewer');
    if (bossViewer) bossViewer.style.display = 'none';
  };

  // Boss decision feature disabled — no-op
  window.refreshSddBossDecisionViewer = function() {
    var bossPanel = document.getElementById('sdd-boss-decision-panel');
    if (bossPanel) bossPanel.style.display = 'none';
    var bossViewer = document.getElementById('sdd-boss-decision-viewer');
    if (bossViewer) bossViewer.style.display = 'none';
  };

  /** Submitted mode: keep form locked, but allow decision action for tasklist gating. */
  window.refreshSddPostSubmitDecisionChrome = function() {
    var nb = document.getElementById('noteBossDecision');
    if (nb) {
      nb.readOnly = false;
      nb.disabled = false;
      nb.style.background = '#fff';
      nb.style.color = '#1A0A0A';
      nb.style.cursor = 'text';
    }
    var wrap = document.getElementById('sdd-staff-decision-wrap');
    if (wrap) wrap.style.display = 'block';
    ['sdd-approver-approve', 'sdd-approver-hold', 'sdd-approver-reject'].forEach(function(id) {
      var b = document.getElementById(id);
      if (!b) return;
      b.disabled = false;
      b.style.opacity = '1';
      b.style.cursor = 'pointer';
      b.style.filter = '';
      b.title = 'Keputusan ini menentukan masuk/tidaknya ke Task List Mill Onboarding.';
    });
    _syncDecisionBadge();
  };

  function normStr(v) {
    return String(v || '').trim().toLowerCase();
  }

  function escHtml(v) {
    return String(v === undefined || v === null ? '' : v)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function kvLine(label, value) {
    return '<div style="display:flex;gap:0;padding:4px 0;border-bottom:1px solid rgba(74,28,28,0.05);font-size:13.5px;line-height:1.5;">'
      + '<span style="color:#5F4A48;font-weight:600;width:260px;flex-shrink:0;">' + escHtml(label) + '</span>'
      + '<span style="color:#9C8080;font-weight:500;width:20px;flex-shrink:0;">:</span>'
      + '<span style="color:#1A0A0A;">' + (escHtml(value) || '<span style="color:#B09A9A;">—</span>') + '</span>'
      + '</div>';
  }

  function sectionCardSimple(title, innerHtml) {
    return '<div style="background:white;border:1px solid rgba(44,40,40,0.09);border-radius:12px;margin-bottom:6px;overflow:hidden;">'
      + '<div style="display:flex;align-items:center;justify-content:space-between;padding:12px 18px;border-bottom:1.5px solid rgba(44,40,40,0.1);background:rgba(44,40,40,0.03);">'
      + '<span style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#2C2828;">' + escHtml(title) + '</span>'
      + '</div>'
      + '<div style="padding:10px 16px;">' + innerHtml + '</div>'
      + '</div>';
  }

  /** Clone SDD row for merge/upsert (Apps Script ignores unknown keys). */
  function stripSddRowForApi(row) {
    if (!row || typeof row !== 'object') return {};
    const o = Object.assign({}, row);
    delete o._row;
    return o;
  }

  /** First non-empty value among possible sheet header spellings (Legality columns vary). */
  function pickSavedCol(p, keys) {
    p = p || {};
    const rowKeys = Object.keys(p);
    const keyByNorm = {};
    rowKeys.forEach(function(k) {
      const nk = normalizeLooseKey(k);
      if (nk && !keyByNorm[nk]) keyByNorm[nk] = k;
    });
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      let actualKey = k;
      if (!(actualKey in p)) {
        const nk = normalizeLooseKey(k);
        if (nk && keyByNorm[nk]) {
          actualKey = keyByNorm[nk];
        } else if (nk) {
          const fuzzy = rowKeys.find(function(rk) {
            const rnk = normalizeLooseKey(rk);
            return rnk === nk || rnk.indexOf(nk) !== -1 || nk.indexOf(rnk) !== -1;
          });
          if (fuzzy) actualKey = fuzzy;
        }
      }
      const v = p[actualKey];
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(+v)) {
          return v.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
        }
        return String(v).trim();
      }
    }
    return '';
  }

  /** Exact-key variant (no fuzzy contains match) to avoid collisions like
   * "Distance" vs "Distance to Refinery (km)". */
  function pickSavedColExact(p, keys) {
    p = p || {};
    const rowKeys = Object.keys(p);
    const keyByNorm = {};
    rowKeys.forEach(function(k) {
      const nk = normalizeLooseKey(k);
      if (nk && !keyByNorm[nk]) keyByNorm[nk] = k;
    });
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      let actualKey = k;
      if (!(actualKey in p)) {
        const nk = normalizeLooseKey(k);
        if (nk && keyByNorm[nk]) actualKey = keyByNorm[nk];
      }
      const v = p[actualKey];
      if (v !== undefined && v !== null && String(v).trim() !== '') {
        if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(+v)) {
          return v.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
        }
        return String(v).trim();
      }
    }
    return '';
  }

  function pickSavedDistanceValue(p) {
    const direct = pickSavedColExact(p, ['Distance', 'Distance (km)', 'Distance KM']);
    function normalizeDistance(v) {
      const s = String(v || '').trim();
      if (!s || /^please\s*select$/i.test(s) || s === '-' || s === '—') return '';
      if (/^\d+([.,]\d+)?$/.test(s)) return s.replace(',', '.') + ' KM';
      if (/^\s*\d+(\.\d+)?\s*km\s*$/i.test(s)) return s.replace(/\s+/g, ' ').trim().toUpperCase().replace('KM', ' KM');
      return '';
    }
    const directNorm = normalizeDistance(direct);
    if (directNorm) return directNorm;

    // Fallback for legacy/surprising headers in saved rows.
    const keys = Object.keys(p || {});
    for (let i = 0; i < keys.length; i++) {
      const k = String(keys[i] || '');
      const nk = normalizeLooseKey(k);
      if (!nk || nk.indexOf('distance') === -1) continue;
      const val = normalizeDistance(p[k]);
      if (val) return val;
    }
    return direct || '';
  }

  function buildPseudoMainFormDataFromSavedRow(p) {
    p = p || {};
    const d = [];
    d.push(['MILL SUPPLIER REGISTRATION FORM']);
    d.push([]); d.push([]); d.push(['A. COMPANY PROFILE']);
    d.push(['Group Name', '', pickSavedCol(p, ['Group Name', 'Grup Name']), '', 'Office Fax', '', pickSavedCol(p, ['Office Fax'])]);
    d.push(['Company Name', '', pickSavedCol(p, ['Company Name']), '', 'Office Email', '', pickSavedCol(p, ['Office Email'])]);
    d.push(['Current Owner', '', pickSavedCol(p, ['Current Owner']), '', 'Contact Person', '', pickSavedCol(p, ['Contact Person'])]);
    d.push(['Take Over From', '', pickSavedCol(p, ['Take Over From']), '', 'Position', '', pickSavedCol(p, ['Contact Position', 'Position'])]);
    d.push(['Previous Owner', '', pickSavedCol(p, ['Previous Owner']), '', 'Mobile Phone', '', pickSavedCol(p, ['Contact Mobile'])]);
    d.push(['Office Address', '', pickSavedCol(p, ['Office Address']), '', 'Contact Email', '', pickSavedCol(p, ['Contact Email'])]);
    d.push(['Office Phone', '', pickSavedCol(p, ['Office Phone']), '', 'Sustainability PIC', '', pickSavedCol(p, ['Sustainability PIC'])]);
    d.push(['Company Website', '', pickSavedCol(p, ['Company Website']), '', 'Mobile Phone', '', pickSavedCol(p, ['Sustainability PIC Mobile'])]);
    d.push(['NDPE Commitment', '', pickSavedCol(p, ['NDPE Commitment']), '', 'HRDD Commitment', '', pickSavedCol(p, ['HRDD Commitment'])]);

    d.push(['B. MILL INFORMATION']);
    d.push(['Mill Name', '', pickSavedCol(p, ['Mill Name'])]);
    d.push(['Mill Address', '', pickSavedCol(p, ['Mill Address'])]);
    d.push(['Distance to Refinery', '', pickSavedCol(p, ['Distance to Refinery (km)', 'Distance to Refinery'])]);
    d.push(['Distance', '', pickSavedDistanceValue(p)]);
    d.push(['Mill Coordinate Latitude (Decimal)', '', pickSavedCol(p, ['Latitude']), '', 'Longitude (Decimal)', '', pickSavedCol(p, ['Longitude'])]);
    d.push(['Mill Category', '', pickSavedCol(p, ['Mill Category'])]);
    d.push(['Mill Capacity (Ton/Hour)', '', pickSavedCol(p, ['Mill Capacity (Ton/Hour)'])]);
    d.push(['Sterilizer Type', '', pickSavedCol(p, ['Sterilizer Type'])]);
    d.push(['Storage Tank Capacity', '', pickSavedCol(p, ['Storage Tank Capacity'])]);
    d.push(['Silo Capacity', '', pickSavedCol(p, ['Silo Capacity'])]);
    d.push(['Commisioning Date', '', pickSavedCol(p, ['Commissioning Date', 'Commisioning Date'])]);

    // Prefer JSON product lines first — needed for KCP heuristic when type column is missing/wrong.
    var savedProductLines = null;
    try {
      var pj = pickSavedCol(p, ['SDD - Product Lines JSON']);
      if (pj && String(pj).trim()) {
        var arrPl = JSON.parse(String(pj));
        if (Array.isArray(arrPl) && arrPl.length) savedProductLines = arrPl;
      }
    } catch (ePl) {}

    // Relational MAIN often has only `supplier_type`; legacy rows use `Supplier Type`.
    var _stypeRaw = pickSavedCol(p, ['supplier_type', 'Supplier Type', 'SUPPLIER_TYPE', 'SupplierType']);
    var _isKcpSaved = normalizeSddSupplierType(_stypeRaw) === 'KCP';
    // If type was not stored on old rows, infer KCP from canonical product lines (CPKO/PKE, no "other" column).
    if (!_isKcpSaved && savedProductLines && savedProductLines.length) {
      var _allNoOther = savedProductLines.every(function(pl) {
        return !String(pl.other || '').trim();
      });
      var _hasKcpProducts = savedProductLines.some(function(pl) {
        var m = String(pl.main || '').trim().toUpperCase();
        return m === 'CPKO' || m === 'PKE';
      });
      if (_allNoOther && _hasKcpProducts) _isKcpSaved = true;
    }
    if (_isKcpSaved) {
      // KCP: single-column layout (Product | Avg | Yield — no Other Product columns)
      d.push(['Product to be Produce', '', 'Product', 'Average Production / Month', 'YIELD', '', '']);
    } else {
      // Mill / Trader: dual-column layout
      d.push(['Product to be Supply', '', 'Main Product', 'Average Potential / Month', 'YIELD', 'Other Product', 'Average Potential / Month']);
    }
    if (savedProductLines && savedProductLines.length) {
      savedProductLines.forEach(function(pl) {
        d.push([
          '', '',
          String(pl.main != null ? pl.main : ''),
          String(pl.mainAvg != null ? pl.mainAvg : ''),
          String(pl.yield != null ? pl.yield : ''),
          String(pl.other != null ? pl.other : ''),
          String(pl.otherAvg != null ? pl.otherAvg : ''),
        ]);
      });
    } else {
      d.push([
        '', '',
        p['Main Product'] || '',
        p['Main Product Avg Production/Month (Ton)'] || '',
        p['Main Product Yield'] || '',
        p['Other Product 1'] || '',
        p['Other Product 1 Avg/Month (Ton)'] || ''
      ]);
      d.push([
        '', '',
        p['Main Product 2'] || '',
        p['Main Product 2 Avg Production/Month (Ton)'] || '',
        p['Main Product 2 Yield'] || '',
        p['Other Product 2'] || '',
        p['Other Product 2 Avg/Month (Ton)'] || ''
      ]);
      d.push([
        '', '',
        p['Main Product 3'] || '',
        p['Main Product 3 Avg Production/Month (Ton)'] || '',
        p['Main Product 3 Yield'] || '',
        p['Other Product 3'] || '',
        p['Other Product 3 Avg/Month (Ton)'] || ''
      ]);
    }
    if (_isKcpSaved) {
      d.push(['CPKO Quality', '', 'FFA', p['CPO Quality - FFA'] || '', 'M&I', p['CPO Quality - M&I'] || '', 'DOBI', p['CPO Quality - DOBI'] || '']);
      d.push(['PKE Quality', '', 'FFA', p['PK Quality - FFA'] || '', 'MOIST', p['PK Quality - MOIST'] || '', 'DIRT', p['PK Quality - DIRT'] || '']);
    } else {
      d.push(['CPO Quality', '', 'FFA', p['CPO Quality - FFA'] || '', 'M&I', p['CPO Quality - M&I'] || '', 'DOBI', p['CPO Quality - DOBI'] || '']);
      d.push(['PK Quality', '', 'FFA', p['PK Quality - FFA'] || '', 'MOIST', p['PK Quality - MOIST'] || '', 'DIRT', p['PK Quality - DIRT'] || '']);
    }

    d.push([]); d.push(['C. CERTIFICATION']);
    d.push(['Certification Members', '', '', '', '', '', 'Start', 'End', 'Certification Body Name']);
    d.push(['RSPO', p['RSPO Status'] || '', 'Cert Number', p['RSPO Cert No'] || '', '', 'Validity Period', p['RSPO Valid Start'] || '', p['RSPO Valid End'] || '', p['RSPO Cert Body'] || '']);
    d.push(['ISCC', p['ISCC Status'] || '', 'Cert Number', p['ISCC Cert No'] || '', '', 'Validity Period', p['ISCC Valid Start'] || '', p['ISCC Valid End'] || '', p['ISCC Cert Body'] || '']);
    d.push(['ISPO Mill', p['ISPO Status'] || '', 'Cert Number', p['ISPO Cert No'] || '', '', 'Validity Period', p['ISPO Valid Start'] || '', p['ISPO Valid End'] || '', p['ISPO Cert Body'] || '']);
    d.push(['INS', p['INS Status'] || '', 'Cert Number', p['INS Cert No'] || '', '', 'Validity Period', p['INS Valid Start'] || '', p['INS Valid End'] || '', p['INS Cert Body'] || '']);
    d.push(['GGL', p['GGL Status'] || '', 'Cert Number', p['GGL Cert No'] || '', '', 'Validity Period', p['GGL Valid Start'] || '', p['GGL Valid End'] || '', p['GGL Cert Body'] || '']);

    d.push([]); d.push(['D. LEGALITY']);
    d.push(['HGU/HGB', pickSavedCol(p, ['HGU/HGB Type', 'HGU HGB Type', 'HGU/HGB', 'HGU HGB']), 'Document Num', pickSavedCol(p, ['HGU/HGB No', 'HGU HGB No', 'HGU/HGB Number', 'HGU HGB Number']), '', 'Issue Date', pickSavedCol(p, ['HGU/HGB Issue Date', 'HGU HGB Issue Date', 'HGU/HGB Issue', 'HGU HGB Issue'])]);
    d.push(['IUP', pickSavedCol(p, ['IUP Type']), 'Document Num', pickSavedCol(p, ['IUP No', 'IUP Number']), '', 'Issue Date', pickSavedCol(p, ['IUP Issue Date', 'IUP Issue'])]);
    d.push(['Izin Lokasi', pickSavedCol(p, ['Izin Lokasi Type', 'Izin Lokasi']), 'Document Num', pickSavedCol(p, ['Izin Lokasi No', 'Izin Lokasi Number', 'IZIN LOKASI']), '', 'Issue Date', pickSavedCol(p, ['Izin Lokasi Issue Date', 'Izin Lokasi Issue'])]);
    d.push(['Izin Lingkungan', pickSavedCol(p, ['Izin Lingkungan Type', 'Izin Lingkungan']), 'Document Num', pickSavedCol(p, ['Izin Lingkungan No', 'Izin Lingkungan Number', 'IZIN LINGKUNGAN']), '', 'Issue Date', pickSavedCol(p, ['Izin Lingkungan Issue Date', 'Izin Lingkungan Issue'])]);
    d.push(['IMB/PBG', pickSavedCol(p, ['IMB/PBG Type', 'IMB/PBG']), 'Document Num', pickSavedCol(p, ['IMB/PBG No', 'IMB/PBG Number', 'IMB', 'PBG']), '', 'Issue Date', pickSavedCol(p, ['IMB/PBG Issue Date', 'IMB/PBG Issue'])]);
    d.push(['NIB', pickSavedCol(p, ['NIB Type', 'NIB']), 'Document Num', pickSavedCol(p, ['NIB No', 'NIB Number', 'NIB']), '', 'Issue Date', pickSavedCol(p, ['NIB Issue Date', 'NIB Issue'])]);
    return d;
  }

  function scoreSavedRowAsMainForm(r) {
    if (!r || typeof r !== 'object') return 0;
    const keyFields = [
      'Group Name', 'Grup Name', 'Company Name', 'Mill Name',
      'Office Address', 'Office Phone', 'Office Fax', 'Office Email', 'Company Website',
      'Contact Person', 'Contact Position', 'Contact Mobile', 'Contact Email',
      'Current Owner', 'Take Over From', 'Previous Owner', 'Sustainability PIC', 'Sustainability PIC Mobile',
      'Mill Address', 'Latitude', 'Longitude',
      'HGU/HGB No', 'IUP No', 'Izin Lokasi No', 'Izin Lingkungan No', 'IMB/PBG No', 'NIB No',
      'HGU/HGB', 'IUP', 'IZIN LOKASI', 'IZIN LINGKUNGAN'
    ];
    return keyFields.reduce(function(acc, k) {
      return acc + (String(r[k] || '').trim() ? 1 : 0);
    }, 0);
  }

  function pickBestPrimaryRowForSubmission(rowList, fallback) {
    const gr = rowList && rowList.length ? rowList : (fallback ? [fallback] : []);
    if (!gr.length) return fallback || null;
    const nonTrace = gr.filter(function(r) { return r && !rowIsTraceabilityOnlyPayload(r); });
    const pool = nonTrace.length ? nonTrace : gr;
    return pool.slice().sort(function(a, b) {
      return scoreSavedRowAsMainForm(b) - scoreSavedRowAsMainForm(a);
    })[0] || gr[0] || fallback;
  }

  function mergeSavedMainFormFromRows(preferredRow, relatedRows) {
    const out = Object.assign({}, preferredRow || {});
    (relatedRows || []).forEach(function(r) {
      if (!r || typeof r !== 'object') return;
      Object.keys(r).forEach(function(k) {
        const cur = out[k];
        const next = r[k];
        const curEmpty = cur === undefined || cur === null || String(cur).trim() === '';
        if (curEmpty && next !== undefined && next !== null && String(next).trim() !== '') {
          out[k] = next;
        }
      });
    });
    return out;
  }

  function buildPseudoTraceabilityDataFromSavedRows(relatedRows) {
    const rows = relatedRows || [];
    const d = [];
    const seenTml = {};
    const seenFfb = {};
    d.push(['Traceability Data']);
    d.push(['*If there is no supply from the mill, continue filling in section B to complete the FFB supply data']);
    d.push([]); d.push(['A. Mill List']);
    d.push([]);
    d.push(['COMPANY NAME','MILL NAME','UML ID','VILLAGE','SUB DISTRICT','DISTRICT','CAPACITY','LAT','LONG','LEGALITY','ISPO (Y/N)','RSPO (Y/N)','ISCC (Y/N)','TOTAL SUPPLY CPO / PK (TON)']);
    rows.filter(function(r){
      return String(r['TML - Mill Name'] || '').trim() !== '' || String(r['TML - Company Name'] || '').trim() !== '';
    }).forEach(function(r){
      const tmlSig = [
        normalizeCellText(r['TML - Company Name']),
        normalizeCellText(r['TML - Mill Name']),
        normalizeCellText(r['TML - UML ID']),
        normalizeCellText(r['TML - Village']),
        normalizeCellText(r['TML - Sub District']),
        normalizeCellText(r['TML - District']),
        normalizeCellText(r['TML - Capacity (Ton/Hour)']),
        normalizeCellText(r['TML - Latitude']),
        normalizeCellText(r['TML - Longitude']),
        normalizeCellText(r['TML - Legality']),
        normalizeCellText(r['TML - ISPO (Y/N)']),
        normalizeCellText(r['TML - RSPO (Y/N)']),
        normalizeCellText(r['TML - ISCC (Y/N)']),
        normalizeCellText(r['TML - Total Supply CPO/PK (Ton)'])
      ].join('|');
      if (seenTml[tmlSig]) return;
      seenTml[tmlSig] = true;
      d.push([
        r['TML - Company Name'] || '',
        r['TML - Mill Name'] || '',
        r['TML - UML ID'] || '',
        r['TML - Village'] || '',
        r['TML - Sub District'] || '',
        r['TML - District'] || '',
        r['TML - Capacity (Ton/Hour)'] || '',
        r['TML - Latitude'] || '',
        r['TML - Longitude'] || '',
        r['TML - Legality'] || '',
        r['TML - ISPO (Y/N)'] || '',
        r['TML - RSPO (Y/N)'] || '',
        r['TML - ISCC (Y/N)'] || '',
        r['TML - Total Supply CPO/PK (Ton)'] || ''
      ]);
    });

    d.push([]); d.push(['B. FFB Supplier List']); d.push([]);
    d.push(['MILL NAME','SUPPLIER GROUP NAME','SUPPLIER NAME','VILLAGE','SUB DISTRICT','DISTRICT','SUPPLIER CATEGORY','CONSESION AREA','PLANTED AREA','NUMBERS OF SMALLHOLDERS','PLANTED YEAR','LEGALITY','Lat','Long','ISPO (Y/N)','RSPO (Y/N)','ISCC (Y/N)','TOTAL SUPPLY FFB (TON)']);
    rows.filter(function(r){
      // Only include rows that genuinely represent an FFB supplier entry.
      // A row with only FFB - Mill Name (no Supplier Name) is a main-form row stamped with
      // identity metadata — do NOT render it as an FFB supplier row.
      return String(r['FFB - Supplier Name'] || '').trim() !== '';
    }).forEach(function(r){
      const ffbSig = [
        normalizeCellText(r['FFB - Mill Name']),
        normalizeCellText(r['FFB - Supplier Group Name']),
        normalizeCellText(r['FFB - Supplier Name']),
        normalizeCellText(r['FFB - Village']),
        normalizeCellText(r['FFB - Sub District']),
        normalizeCellText(r['FFB - District']),
        normalizeCellText(r['FFB - Supplier Category']),
        normalizeCellText(r['FFB - Concession Area (Ha)']),
        normalizeCellText(r['FFB - Planted Area (Ha)']),
        normalizeCellText(r['FFB - Number of Smallholders']),
        normalizeCellText(r['FFB - Planted Year']),
        normalizeCellText(r['FFB - Legality']),
        normalizeCellText(r['FFB - Latitude']),
        normalizeCellText(r['FFB - Longitude']),
        normalizeCellText(r['FFB - ISPO (Y/N)']),
        normalizeCellText(r['FFB - RSPO (Y/N)']),
        normalizeCellText(r['FFB - ISCC (Y/N)']),
        normalizeCellText(r['FFB - Total Supply FFB (Ton)'])
      ].join('|');
      if (seenFfb[ffbSig]) return;
      seenFfb[ffbSig] = true;
      d.push([
        r['FFB - Mill Name'] || '',
        r['FFB - Supplier Group Name'] || '',
        r['FFB - Supplier Name'] || '',
        r['FFB - Village'] || '',
        r['FFB - Sub District'] || '',
        r['FFB - District'] || '',
        r['FFB - Supplier Category'] || '',
        r['FFB - Concession Area (Ha)'] || '',
        r['FFB - Planted Area (Ha)'] || '',
        r['FFB - Number of Smallholders'] || '',
        r['FFB - Planted Year'] || '',
        r['FFB - Legality'] || '',
        r['FFB - Latitude'] || '',
        r['FFB - Longitude'] || '',
        r['FFB - ISPO (Y/N)'] || '',
        r['FFB - RSPO (Y/N)'] || '',
        r['FFB - ISCC (Y/N)'] || '',
        r['FFB - Total Supply FFB (Ton)'] || ''
      ]);
    });
    return d;
  }

  function savedSubmissionGroupKey(r) {
    const supplierType = normalizeSddSupplierType(r['Supplier Type'] || r['SUPPLIER_TYPE'] || r['SupplierType']);
    const millId = normStr(r['Mill ID'] || r['Mil ID'] || r['MILL ID']);
    const dateImported = normStr(r['Date Imported']);
    const importedBy = normStr(r['Imported By']);
    const groupName = normStr(r['Group Name'] || r['Grup Name']);
    const company = normStr(r['Company Name']);
    const mill = normStr(r['Mill Name'] || r['TML - Mill Name'] || r['FFB - Mill Name']);
    const tp = supplierType || 'UNKNOWN';
    if (millId && dateImported && importedBy) return 'mdi|' + tp + '|' + millId + '|' + dateImported + '|' + importedBy;
    if (millId && dateImported) return 'md|' + tp + '|' + millId + '|' + dateImported;
    if (millId) return 'm|' + tp + '|' + millId;
    if (dateImported && importedBy && (groupName || company || mill)) return 'dib|' + tp + '|' + dateImported + '|' + importedBy + '|' + groupName + '|' + company + '|' + mill;
    return 'row|' + String(r._row || Math.random());
  }

  /**
   * Match traceability rows to a saved list key when they were stored before identity stamping, or when only Date+Imported By+Mill align.
   * Intentionally strict: requires BOTH Date+ImportedBy match AND (Mill ID match OR Mill Name match).
   * This prevents rows from other submissions being pulled in when Date/ImportedBy happen to collide.
   */
  function legacyRowBelongsToSavedSubmission(r, key, primaryRow) {
    if (!r || !primaryRow) return false;
    const pType = normalizeSddSupplierType(primaryRow['Supplier Type'] || primaryRow['SUPPLIER_TYPE'] || primaryRow['SupplierType']);
    const rType = normalizeSddSupplierType(r['Supplier Type'] || r['SUPPLIER_TYPE'] || r['SupplierType']);
    if (pType && rType && pType !== rType) return false;
    const di = normStr(primaryRow['Date Imported']);
    const ib = normStr(primaryRow['Imported By']);
    if (!di || !ib) return false;
    if (normStr(r['Date Imported']) !== di || normStr(r['Imported By']) !== ib) return false;
    // Must have traceability content to be a traceability row.
    const hasTrace = String(r['TML - Mill Name'] || r['FFB - Supplier Name'] || '').trim() !== '';
    if (!hasTrace) return false;
    // Mill ID match is the strongest signal.
    const pMid = normStr(primaryRow['Mill ID'] || primaryRow['Mil ID']);
    const rMid = normStr(r['Mill ID'] || r['Mil ID']);
    if (pMid && rMid && pMid === rMid) return true;
    // Fallback for legacy rows without Mill ID:
    // same Date Imported + Imported By, and same Group/Company when present.
    // Do NOT require Mill Name match because one submission can contain multiple mills/ffb rows.
    const pGroup = normStr(primaryRow['Group Name'] || primaryRow['Grup Name']);
    const rGroup = normStr(r['Group Name'] || r['Grup Name']);
    if (pGroup && rGroup && pGroup !== rGroup) return false;
    const pCompany = normStr(primaryRow['Company Name']);
    const rCompany = normStr(r['Company Name']);
    if (pCompany && rCompany && pCompany !== rCompany) return false;
    return true;
  }

  function rowsMatchingSavedSubmissionKey(allRows, key, primaryRow) {
    if (!Array.isArray(allRows) || !key || !primaryRow) return [];
    const out = [];
    const seen = {};
    for (let i = 0; i < allRows.length; i++) {
      const r = allRows[i];
      if (!r) continue;
      const match = savedSubmissionGroupKey(r) === key || legacyRowBelongsToSavedSubmission(r, key, primaryRow);
      if (!match) continue;
      const sid = String(r._row != null ? r._row : importRowSignature(r));
      if (seen[sid]) continue;
      seen[sid] = true;
      out.push(r);
    }
    return out;
  }

  /**
   * Fallback for legacy/TRADER rows where submission key can drift:
   * pull likely MAIN rows by stable import identity + type + company/group.
   */
  function supplementalMainRowsForReference(allRows, referenceRow) {
    if (!Array.isArray(allRows) || !referenceRow) return [];
    const refType = normalizeSddSupplierType(referenceRow['Supplier Type'] || referenceRow['SUPPLIER_TYPE'] || referenceRow['SupplierType']);
    const refDate = normStr(referenceRow['Date Imported']);
    const refBy = normStr(referenceRow['Imported By']);
    const refGroup = normStr(referenceRow['Group Name'] || referenceRow['Grup Name']);
    const refCompany = normStr(referenceRow['Company Name']);
    return allRows.filter(function(r) {
      if (!r || rowIsTraceabilityOnlyPayload(r)) return false;
      const rType = normalizeSddSupplierType(r['Supplier Type'] || r['SUPPLIER_TYPE'] || r['SupplierType']);
      if (refType && rType && refType !== rType) return false;
      if (refDate && normStr(r['Date Imported']) !== refDate) return false;
      if (refBy && normStr(r['Imported By']) !== refBy) return false;
      const rGroup = normStr(r['Group Name'] || r['Grup Name']);
      const rCompany = normStr(r['Company Name']);
      if (refGroup && rGroup && refGroup !== rGroup) return false;
      if (refCompany && rCompany && refCompany !== rCompany) return false;
      // Must have some substantive main-form fields
      return scoreSavedRowAsMainForm(r) > 0;
    });
  }

  function getSavedScreeningLabel(r) {
    const status = String(r['SCR - Screening Status'] || '');
    const tp = normalizeSddSupplierType(r['Supplier Type'] || r['SUPPLIER_TYPE'] || r['SupplierType']);
    const groupName = r['Group Name'] || r['Grup Name'] || '-';
    const mill = r['Mill Name'] || r['TML - Mill Name'] || r['FFB - Mill Name'] || '-';
    const millId = r['Mill ID'] || r['Mil ID'] || '';
    return (status ? '[' + status + '] ' : '') + (tp ? '[' + tp + '] ' : '') + groupName + ' · ' + mill + (millId ? ' · ' + millId : '');
  }

  /** Logical supplier key for display dedupe (hides old draft history entries). */
  function savedSubmissionLogicalDisplayKey(r) {
    if (!r || typeof r !== 'object') return '';
    const supplierType = normalizeSddSupplierType(r['Supplier Type'] || r['SUPPLIER_TYPE'] || r['SupplierType']) || 'UNKNOWN';
    const importedBy = normStr(r['Imported By']);
    const groupName = normStr(r['Group Name'] || r['Grup Name']);
    const company = normStr(r['Company Name']);
    const mill = normStr(r['Mill Name'] || r['TML - Mill Name'] || r['FFB - Mill Name']);
    const millId = normStr(r['Mill ID'] || r['Mil ID']);
    if (importedBy && (groupName || company || mill)) {
      return 'ib|' + supplierType + '|' + importedBy + '|' + groupName + '|' + company + '|' + mill;
    }
    if (groupName || company || mill) {
      return 'gcm|' + supplierType + '|' + groupName + '|' + company + '|' + mill;
    }
    if (millId) return 'mid|' + supplierType + '|' + millId;
    return 'row|' + String(r._row || Math.random());
  }

  function renderSavedScreeningListView(entries, selectedKey) {
    const holder = document.getElementById('scr-saved-list-wrap');
    if (!holder) return;
    if (!entries.length) {
      holder.innerHTML = '<div style="padding:20px 0;color:#A09090;font-size:12.5px;text-align:center;letter-spacing:0.01em;">Tidak ada screening tersimpan.</div>';
      return;
    }

    function fmtDate(raw) {
      const s = String(raw || '').trim();
      if (!s) return '—';
      const d = new Date(s);
      if (isNaN(d.getTime())) return s.slice(0, 10) || '—';
      return d.toLocaleDateString('id-ID', { day:'2-digit', month:'short', year:'numeric' });
    }

    function fmtStatus(r) {
      const scrStatus = String(r['SCR - Screening Status'] || '').trim().toLowerCase();
      if (scrStatus !== 'submitted') return 'Draft';
      const decRaw = String(
        r['statusSDD'] || r['statusSdd'] || r['Status SDD'] ||
        r['statusBossDecision'] || r['Status Boss Decision'] || ''
      ).trim();
      if (!decRaw) return 'Submitted';
      return 'Submitted · ' + _normalizeDecisionLabel(decRaw);
    }

    function parseDateTs(raw) {
      const s = String(raw || '').trim();
      if (!s) return 0;
      const t = new Date(s).getTime();
      return isNaN(t) ? 0 : t;
    }

    if (!window._scrSavedListSort || typeof window._scrSavedListSort !== 'object') {
      window._scrSavedListSort = { col: 'lastUpdate', dir: 'desc' };
    }
    if (!window._scrSavedColumnFilters || typeof window._scrSavedColumnFilters !== 'object') {
      window._scrSavedColumnFilters = {};
    }
    var sortState = window._scrSavedListSort;
    var sortCol = sortState.col || 'lastUpdate';
    var sortDir = (sortState.dir === 'asc' || sortState.dir === 'desc') ? sortState.dir : 'desc';
    var sortMul = sortDir === 'asc' ? 1 : -1;
    var colFilters = window._scrSavedColumnFilters;

    function compareText(a, b) {
      return String(a || '').localeCompare(String(b || ''), 'id', { sensitivity: 'base' });
    }

    var cols = [
      { key: 'dateImport',  label: 'Date Import' },
      { key: 'companyName', label: 'Company Name' },
      { key: 'category',    label: 'Category' },
      { key: 'status',      label: 'Status' },
      { key: 'lastUpdate',  label: 'Last Update' },
    ];

    function toRowItem(entry) {
      const r = entry.primary || {};
      return {
        entry: entry,
        key: entry.key,
        row: r,
        dateImport: fmtDate(r['Date Imported']),
        companyName: String(r['Company Name'] || r['Group Name'] || r['Grup Name'] || r['Mill Name'] || '—').trim() || '—',
        category: normalizeSddSupplierType(r['Supplier Type'] || r['SUPPLIER_TYPE'] || r['SupplierType'] || r['supplier_type']) || '—',
        status: fmtStatus(r),
        lastUpdate: fmtDate(r['updated_at'] || r['SCR - Last Updated']),
        dateImportTs: parseDateTs(r['Date Imported']),
        lastUpdateTs: parseDateTs(r['updated_at'] || r['SCR - Last Updated']),
      };
    }

    var rowItems = entries.map(toRowItem);
    var uniqueByCol = {};
    cols.forEach(function(c) {
      uniqueByCol[c.key] = Array.from(new Set(rowItems.map(function(it) { return String(it[c.key] || '—'); })))
        .sort(function(a, b) { return compareText(a, b); });
      if (!Array.isArray(colFilters[c.key])) colFilters[c.key] = [];
    });

    const filtered = rowItems.filter(function(it) {
      return cols.every(function(c) {
        var active = colFilters[c.key];
        if (!Array.isArray(active) || !active.length) return true;
        return active.indexOf(String(it[c.key] || '—')) !== -1;
      });
    });

    const sorted = filtered.slice().sort(function(a, b) {
      var cmp = 0;
      if (sortCol === 'dateImport') {
        cmp = a.dateImportTs - b.dateImportTs;
      } else if (sortCol === 'companyName') {
        cmp = compareText(a.companyName, b.companyName);
      } else if (sortCol === 'category') {
        cmp = compareText(a.category, b.category);
      } else if (sortCol === 'status') {
        cmp = compareText(a.status, b.status);
      } else {
        cmp = a.lastUpdateTs - b.lastUpdateTs;
      }
      if (cmp === 0) {
        cmp = a.lastUpdateTs - b.lastUpdateTs;
      }
      return cmp * sortMul;
    });

    function headLabel(colKey, text) {
      if (sortCol !== colKey) return text;
      return text + (sortDir === 'asc' ? ' ▲' : ' ▼');
    }

    function hasActiveFilter(colKey) {
      var vals = colFilters[colKey];
      return Array.isArray(vals) && vals.length > 0;
    }

    holder.innerHTML =
      '<div class="scr-saved-search-wrap"><span class="scr-saved-search-count">' + sorted.length + ' result(s)</span></div>'
      + '<div class="scr-saved-table-wrap"><table class="scr-saved-table">'
      + '<thead><tr>'
      + cols.map(function(c) {
        return '<th>'
          + '<div class="scr-th-wrap">'
          + '<span class="scr-th-sort" data-sort-col="' + escHtml(c.key) + '">' + escHtml(headLabel(c.key, c.label)) + '</span>'
          + '<button type="button" class="scr-th-filter-btn' + (hasActiveFilter(c.key) ? ' is-active' : '') + '" data-filter-col="' + escHtml(c.key) + '" title="Filter ' + escHtml(c.label) + '">▾</button>'
          + '</div>'
          + '</th>';
      }).join('')
      + '</tr></thead><tbody>'
      + sorted.map(function(entry) {
        const active = selectedKey && selectedKey === entry.key;
        return '<tr data-scr-key="' + escHtml(entry.key) + '" class="' + (active ? 'is-active' : '') + '">'
          + '<td>' + escHtml(entry.dateImport) + '</td>'
          + '<td>' + escHtml(entry.companyName) + '</td>'
          + '<td>' + escHtml(entry.category) + '</td>'
          + '<td>' + escHtml(entry.status) + '</td>'
          + '<td>' + escHtml(entry.lastUpdate) + '</td>'
          + '</tr>';
      }).join('')
      + '</tbody></table></div>';

    holder.querySelectorAll('tr[data-scr-key]').forEach(function(row) {
      row.addEventListener('click', function() {
        const key = row.getAttribute('data-scr-key') || '';
        const sel = document.getElementById('scr-saved-select');
        if (sel) sel.value = key;
        if (typeof window.loadSavedScrByKeyGlobal === 'function') window.loadSavedScrByKeyGlobal(key);
      });
    });
    holder.querySelectorAll('.scr-th-sort[data-sort-col]').forEach(function(th) {
      th.addEventListener('click', function() {
        var col = th.getAttribute('data-sort-col') || '';
        if (!col) return;
        if (window._scrSavedListSort && window._scrSavedListSort.col === col) {
          window._scrSavedListSort.dir = window._scrSavedListSort.dir === 'asc' ? 'desc' : 'asc';
        } else {
          window._scrSavedListSort = { col: col, dir: 'asc' };
        }
        renderSavedScreeningListView(entries, selectedKey);
      });
    });
    holder.querySelectorAll('.scr-th-filter-btn[data-filter-col]').forEach(function(btn) {
      btn.addEventListener('click', function(ev) {
        ev.stopPropagation();
        var col = btn.getAttribute('data-filter-col') || '';
        if (!col) return;
        holder.querySelectorAll('.scr-filter-menu').forEach(function(m) { m.remove(); });
        var options = uniqueByCol[col] || [];
        var current = Array.isArray(colFilters[col]) ? colFilters[col].slice() : [];
        var selectedVals = current.length ? current : options.slice();
        var menu = document.createElement('div');
        menu.className = 'scr-filter-menu';
        menu.innerHTML =
          '<div class="scr-filter-menu-head">' + escHtml((cols.find(function(c) { return c.key === col; }) || {}).label || col) + '</div>'
          + '<input type="text" class="scr-filter-menu-search" placeholder="Search value...">'
          + '<div class="scr-filter-menu-actions">'
          + '<button type="button" data-filter-action="all">Select all</button>'
          + '<button type="button" data-filter-action="none">Clear all</button>'
          + '</div>'
          + '<div class="scr-filter-menu-list"></div>'
          + '<div class="scr-filter-menu-foot">'
          + '<button type="button" data-filter-action="cancel">Cancel</button>'
          + '<button type="button" data-filter-action="ok" class="is-primary">OK</button>'
          + '</div>';
        holder.appendChild(menu);
        var rectBtn = btn.getBoundingClientRect();
        var rectHolder = holder.getBoundingClientRect();
        menu.style.top = Math.max(0, rectBtn.bottom - rectHolder.top + 4) + 'px';
        menu.style.left = Math.max(0, rectBtn.right - rectHolder.left - 250) + 'px';
        var listEl = menu.querySelector('.scr-filter-menu-list');
        var searchEl = menu.querySelector('.scr-filter-menu-search');

        function renderChecks() {
          if (!listEl) return;
          var kw = String((searchEl && searchEl.value) || '').trim().toLowerCase();
          listEl.innerHTML = options.filter(function(v) {
            return !kw || String(v || '').toLowerCase().indexOf(kw) !== -1;
          }).map(function(v) {
            var checked = selectedVals.indexOf(v) !== -1;
            return '<label class="scr-filter-menu-item"><input type="checkbox" value="' + escHtml(v) + '"' + (checked ? ' checked' : '') + '> <span>' + escHtml(v) + '</span></label>';
          }).join('') || '<div class="scr-filter-menu-empty">No values</div>';
          listEl.querySelectorAll('input[type="checkbox"]').forEach(function(cb) {
            cb.addEventListener('change', function() {
              var val = cb.value;
              if (cb.checked) {
                if (selectedVals.indexOf(val) === -1) selectedVals.push(val);
              } else {
                selectedVals = selectedVals.filter(function(x) { return x !== val; });
              }
            });
          });
        }
        renderChecks();
        if (searchEl) searchEl.addEventListener('input', renderChecks);
        menu.querySelectorAll('[data-filter-action]').forEach(function(actBtn) {
          actBtn.addEventListener('click', function() {
            var action = actBtn.getAttribute('data-filter-action');
            if (action === 'all') { selectedVals = options.slice(); renderChecks(); return; }
            if (action === 'none') { selectedVals = []; renderChecks(); return; }
            if (action === 'cancel') { menu.remove(); return; }
            if (action === 'ok') {
              if (selectedVals.length === 0 || selectedVals.length === options.length) {
                colFilters[col] = [];
              } else {
                colFilters[col] = selectedVals.slice();
              }
              menu.remove();
              renderSavedScreeningListView(entries, selectedKey);
            }
          });
        });
        setTimeout(function() {
          function closeMenuOutside(e) {
            if (!menu.contains(e.target) && e.target !== btn) {
              menu.remove();
              document.removeEventListener('click', closeMenuOutside, true);
            }
          }
          document.addEventListener('click', closeMenuOutside, true);
        }, 0);
      });
    });
  }

  function updateSavedScreeningPickerUI() {
    if (typeof window.refreshSddApproverStaffTraceChrome === 'function') {
      window.refreshSddApproverStaffTraceChrome();
    }
  }

  function renderSavedRecordImportedLikeInitial(primaryRow, relatedRows) {
    const holder = document.getElementById('supplierExcelData');
    if (!holder) return;
    holder.innerHTML = '';
    const mainData = buildPseudoMainFormDataFromSavedRow(primaryRow || {});
    const traceData = buildPseudoTraceabilityDataFromSavedRows(relatedRows || []);

    function appendSheet(name, data) {
      const wrapper = document.createElement('div');
      const sheetId = 'sheet-body-saved-' + name.replace(/\s+/g,'-').toLowerCase();
      wrapper.innerHTML = '<div style="display:flex;align-items:center;justify-content:space-between;font-size:12px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#2C2828;margin:20px 0 0;padding:12px 16px;background:rgba(44,40,40,0.04);border:1.5px solid rgba(44,40,40,0.14);border-radius:10px 10px 0 0;">'
        + '<span>Sheet: ' + escHtml(name) + '</span></div>'
        + '<div id="' + sheetId + '" style="border:1.5px solid rgba(44,40,40,0.14);border-top:none;border-radius:0 0 10px 10px;padding:10px 6px 10px;margin-bottom:8px;"></div>';
      holder.appendChild(wrapper);
      const body = wrapper.querySelector('#' + sheetId);
      supplierDisplayExcelDataTo(body, data, { appendScrForm: false });
    }

    appendSheet('Main Form', mainData);
    /* Single screening block: after Legality (end of Main Form), before Traceability — matches import layout */
    holder.insertAdjacentHTML('beforeend', buildScrForm());
    appendSheet('TRACEABILITY', traceData);
    if (typeof window.refreshSddApproverStaffTraceChrome === 'function') window.refreshSddApproverStaffTraceChrome();
  }

  /**
   * refreshSavedScreeningListGlobal — rewritten to use listSubmissions (MAIN only).
   *
   * Source of truth: SDD_MAIN via GET listSubmissions.
   * No more mixed-row grouping heuristics; each row IS one submission.
   * Dropdown label: [SCR status] [supplier_type] Group · Mill (updated_at)
   * selectKey is now a submission_id string.
   */
  window.refreshSavedScreeningListGlobal = async function(selectKey) {
    const sel = document.getElementById('scr-saved-select');
    if (!sel) return;
    try {
      // Fetch from MAIN sheet only — no mixed TML/FFB rows
      const listResult = await apiListSubmissions({ page_size: 200 });
      const mainRows = (listResult && Array.isArray(listResult.data)) ? listResult.data : [];

      const visibleRows = mainRows.filter(function(r) {
        const st = String(r['SCR - Screening Status'] || '').trim().toLowerCase();
        return st === 'draft' || st === 'submitted';
      });

      /**
       * Build label: [SCR Status] [SupplierType] GroupName · MillName (updated_at date)
       */
      function buildLabel(r) {
        const st  = String(r['SCR - Screening Status'] || '').trim();
        const tp  = String(r['supplier_type'] || r['Supplier Type'] || '').trim().toUpperCase();
        const grp = String(r['Group Name'] || r['Grup Name'] || '—').trim();
        const mill = String(r['Mill Name'] || '—').trim();
        const upd = String(r['updated_at'] || r['SCR - Last Updated'] || '').trim();
        let updStr = '';
        if (upd) {
          const d = new Date(upd);
          if (!isNaN(d.getTime())) {
            updStr = ' (' + d.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' }) + ')';
          } else {
            updStr = ' (' + upd.slice(0, 10) + ')';
          }
        }
        return (st ? '[' + st + '] ' : '') +
               (tp ? '[' + tp + '] ' : '') +
               grp + ' · ' + mill + updStr;
      }

      // Sort by updated_at desc (latest first)
      visibleRows.sort(function(a, b) {
        const da = new Date(a['updated_at'] || 0).getTime();
        const db = new Date(b['updated_at'] || 0).getTime();
        return db - da;
      });

      // Build entries keyed by submission_id (the stable, unambiguous key)
      const entries = visibleRows.map(function(r) {
        const sid = String(r['submission_id'] || '');
        return { key: sid, primary: r, label: buildLabel(r) };
      }).filter(function(e) { return !!e.key; });

      window._scrSavedRowsByKey   = {};
      window._scrSavedGroupsByKey = {};
      sel.innerHTML = '<option value="">— Select saved screening —</option>';
      entries.forEach(function(entry) {
        const sid = entry.key;
        window._scrSavedRowsByKey[sid]   = entry.primary;
        // _scrSavedGroupsByKey.rows is lazy-loaded via getSubmissionById on demand
        window._scrSavedGroupsByKey[sid] = { primary: entry.primary, rows: null };
        const op = document.createElement('option');
        op.value = sid;
        op.textContent = entry.label;
        sel.appendChild(op);
      });

      if (selectKey && window._scrSavedRowsByKey[selectKey]) {
        sel.value = selectKey;
      } else {
        sel.value = '';
      }
      renderSavedScreeningListView(entries, sel.value || '');
      updateSavedScreeningPickerUI();
    } catch (e) {
      if (typeof window.showSddToast === 'function') {
        window.showSddToast('Failed loading saved screening list: ' + (e.message || e), 'error');
      }
    }
  };

  /**
   * loadSavedScrByKeyGlobal — rewritten to use getSubmissionById.
   *
   * key is now a submission_id string.
   * Hydrates from the single API response { main, mills, ffb_rows } —
   * UI never reconstructs from mixed _row heuristics.
   * Empty cells from the API remain '' and are never inherited.
   */
  window.loadSavedScrByKeyGlobal = async function(key) {
    const sel = document.getElementById('scr-saved-select');
    if (!key || !sel) return;
    sel.value = key;

    // ── RESET ALL GLOBAL STATE before loading a different supplier record ──
    window._sddIsLoadedSaved     = false;
    window._sddLastInsertedRow   = null;
    window._sddSubmissionId      = null;
    window._scrLoadedRowNum      = null;
    window._scrLoadedKey         = '';
    window._loadedPrimarySddRow  = null;
    window._sddImportFirstRow    = null;
    window._sddImportedRows      = [];
    window._sddApproverRecordLoaded = false;

    // Use the cached MAIN row for quick supplier-type sync, then fetch full detail
    const cachedMain = window._scrSavedRowsByKey && window._scrSavedRowsByKey[key];
    if (cachedMain) {
      window._sddSupplierType = normalizeSddSupplierType(
        cachedMain['supplier_type'] || cachedMain['Supplier Type'] || cachedMain['SUPPLIER_TYPE']
      ) || window._sddSupplierType;
      syncSddSupplierTypeSelectorUI();
    }

    if (typeof window.showSddToast === 'function') {
      window.showSddToast('Loading submission…', 'info');
    }

    try {
      // ── Fetch full submission from relational API ──
      const result = await apiGetSubmissionById(key);
      // result: { success, main, mills, ffb_rows }
      const mainRow  = result.main     || {};
      const millRows = result.mills    || [];
      const ffbRows  = result.ffb_rows || [];

      // Null-safe: empty cells stay '' — never inherit from previous rows
      function safeStr(v) { return (v === undefined || v === null) ? '' : String(v); }

      // ── PATCH: coordinate keys that must stay dot-decimal ──────────────────
      const _COORD_KEYS = new Set([
        'Latitude', 'Longitude',
        'TML - Latitude', 'TML - Longitude',
        'FFB - Latitude', 'FFB - Longitude',
      ]);
      /**
       * recoverCoord — heuristic heal for values already corrupted by Google Sheets
       * locale id_ID (e.g. "0,318318" stored as integer 318318).
       *
       * Corruption mechanism: Sheets locale id_ID treats comma as thousands separator,
       * so "0,318318" → integer 318318 (leading-zero stripped + decimal point removed),
       * and "101,3838183" → integer 1013838183.
       *
       * Recovery strategy: enumerate ALL decimal-point insertion positions and pick the
       * most plausible one based on field type:
       *   • Latitude  fields: prefer split with MOST decimal digits (smallest int part,
       *                       handles 0.x equatorial coords like "0.318318").
       *   • Longitude fields: prefer split with LARGEST integer part within (-180,180)
       *                       (handles "101.3838183" where int part = 101).
       *
       * Field type is inferred from the key name (contains "long"/"lng" → longitude).
       *
       * Validated against all coordinate values in the sample Excel (Kalimantan/Sumatra):
       *   318318      (Lat) → "0.318318"     ✓
       *   1013838183  (Lng) → "101.3838183"  ✓
       *   321312321   (Lat) → "0.321312321"  ✓   etc.
       *
       * Limitation: best-effort only. Always run fixExistingCoordColumns() in GAS
       * to permanently heal rows; this guard is for rows not yet re-saved.
       *
       * @param {*}      raw        - raw value from API (may be number or string)
       * @param {string} fieldName  - column key name, used to detect lat vs lng
       * @returns {string} dot-decimal coordinate string, or original string if not recoverable
       */
      function recoverCoord(raw, fieldName) {
        const s = String(raw === null || raw === undefined ? '' : raw).trim();
        if (!s) return '';
        // Comma-decimal legacy → dot
        if (s.indexOf(',') !== -1 && s.indexOf('.') === -1) return s.replace(',', '.');
        // String contains a dot. Two cases:
        //   (a) Genuine decimal "0.318318" — parseFloat within WGS-84 range → return as-is.
        //   (b) id_ID thousands-formatted integer: Sheets stored 318318 and rendered it as
        //       "318.318" (dot = thousands sep) or "1.013.838.183" (multi-group).
        //       These are out of WGS-84 range → strip all dots and recover via integer heuristic.
        if (s.indexOf('.') !== -1) {
          const parsed = parseFloat(s);
          if (!isNaN(parsed) && Math.abs(parsed) <= 180) return s; // valid coordinate decimal
          // Out-of-range with dot: strip all dots (id_ID thousands separators) then re-run
          const stripped = s.replace(/\./g, '');
          if (/^-?\d+$/.test(stripped)) {
            // Recurse — will hit the integer-heuristic path below
            return recoverCoord(stripped, fieldName);
          }
          return s; // non-numeric, cannot recover
        }
        // Must be all-digits (with optional leading minus) to be a corrupted integer
        if (!/^-?\d+$/.test(s)) return s;
        const abs_n = Math.abs(parseInt(s, 10));
        const neg   = s.charAt(0) === '-';
        const pfx   = neg ? '-' : '';
        // Within valid WGS-84 range — not corrupted
        if (abs_n <= 180) return s;

        const digits  = String(abs_n);
        const isLng   = /long|lng/i.test(fieldName || '');
        const maxInt  = isLng ? 180 : 90;
        const candidates = [];

        // k = number of decimal digits in the recovered value
        // k ranges from 1 to len(digits): k == len(digits) handles the "0.xxx" case
        // where the leading "0" integer part was stripped.
        for (let k = 1; k <= digits.length; k++) {
          const pow     = Math.pow(10, k);
          const leftInt = Math.floor(abs_n / pow);
          const rightInt= abs_n % pow;
          if (leftInt > maxInt) continue;
          const rightStr = String(rightInt).padStart(k, '0');
          const valStr   = pfx + String(leftInt) + '.' + rightStr;
          candidates.push({ leftInt: leftInt, k: k, valStr: valStr });
        }

        if (!candidates.length) return s;

        let best;
        if (isLng) {
          // Longitude: largest integer part wins (101.38… preferred over 1.013… or 0.101…)
          best = candidates.reduce(function(a, b) { return b.leftInt > a.leftInt ? b : a; });
        } else {
          // Latitude: most decimal digits wins → smallest integer part (0.318318 preferred over 31.8318)
          best = candidates.reduce(function(a, b) { return b.k > a.k ? b : a; });
        }
        return best.valStr;
      }

      function cleanRow(r) {
        if (!r || typeof r !== 'object') return {};
        const out = {};
        Object.keys(r).forEach(function(k) {
          // PATCH: coordinate fields get heuristic recovery (pass key name for lat/lng disambiguation)
          out[k] = _COORD_KEYS.has(k) ? recoverCoord(r[k], k) : safeStr(r[k]);
        });
        return out;
      }
      const cleanMain = cleanRow(mainRow);
      const cleanMills = millRows.map(cleanRow);
      const cleanFfb   = ffbRows.map(cleanRow);

      // ── Hydrate in-memory screening state (REPLACE, bukan merge) ──────────
      // window._tml/_ffbScreeningData adalah authoritative untuk pre-fill modal.
      // REPLACE total — supplier B tidak akan mewarisi data screening dari supplier A.
      (function _hydrateScreeningMemory() {
        function parseItems(str) {
          if (!str) return [];
          return String(str).split(';').map(function(s) {
            s = s.trim();
            if (!s) return null;
            // "APL (123)" → { label: 'APL', ha: '123' }
            var m = s.match(/^(.+?)\s*\(([^)]*)\)$/);
            return m ? { label: m[1].trim(), ha: m[2].trim() } : { label: s, ha: '' };
          }).filter(Boolean);
        }

        var freshTml = {};
        cleanMills.forEach(function(m) {
          var millName = m['TML - Mill Name'] || '';
          if (!millName) return;
          freshTml[millName] = {
            coord       : m['SCR - TML Valid Coordinate']               || '',
            forestItems : parseItems(m['SCR - TML Forest Area']),
            peatItems   : parseItems(m['SCR - TML Peatland']),
            mora        : m['SCR - TML Moratorium']                     || '',
            moraHa      : m['SCR - TML Moratorium (Ha)']                || '',
            defbuf      : m['SCR - TML Deforestation Buffer 50KM (Ha)'] || '',
            status      : m['SCR - TML Screening Status']               || 'Draft',
            date        : m['SCR - TML Screening Date']                 || '',
          };
        });
        window._tmlScreeningData = freshTml;

        var freshFfb = {};
        cleanFfb.forEach(function(f) {
          var supplierName = f['FFB - Supplier Name'] || '';
          if (!supplierName) return;
          freshFfb[supplierName] = {
            coord       : f['FFB - Valid Coordinate']      || '',
            forestItems : parseItems(f['FFB - Forest Area']),
            peatItems   : parseItems(f['FFB - Peatland']),
            mora        : f['FFB - Moratorium']            || '',
            moraHa      : f['FFB - Moratorium (Ha)']       || '',
            distKm      : f['FFB - Distance to Mill (Km)'] || '',
            defor       : f['FFB - Deforestation (Ha)']    || '',
            burn        : f['FFB - Burn Area (Ha)']        || '',
            villageRisk : f['FFB - Village Risk']          || '',
            status      : f['FFB - Screening Status']      || 'Draft',
            date        : f['FFB - Screening Date']        || '',
          };
        });
        window._ffbScreeningData = freshFfb;
      })();

      window._sddSupplierType = normalizeSddSupplierType(
        cleanMain['supplier_type'] || cleanMain['Supplier Type']
      ) || window._sddSupplierType;
      syncSddSupplierTypeSelectorUI();

      // ── Build imported-like display from relational response ──
      // Convert mill/ffb objects back to the flat key format the renderer expects
      function millToFlatRow(m) {
        return Object.assign({}, cleanMain, m);
      }
      function ffbToFlatRow(f) {
        return Object.assign({}, { 'Mill Name': cleanMain['Mill Name'] || '' }, f);
      }

      const allRelatedRows = [].concat(
        cleanMills.map(millToFlatRow),
        cleanFfb.map(ffbToFlatRow)
      );

      // ── Render display ──
      if (typeof buildScrForm === 'function' && typeof initScrForm === 'function') {
        // Set state BEFORE rendering so buildScrForm() can check _loadedPrimarySddRow for submitted status
        const basePrimary = stripSddRowForApi(cleanMain);
        applySddHeaderAliases(basePrimary);
        window._sddImportFirstRow   = basePrimary;
        window._sddImportedRows     = allRelatedRows.map(stripSddRowForApi).filter(hasMeaningfulImportData);
        window._loadedPrimarySddRow = Object.assign({}, basePrimary);
        window._sddIsLoadedSaved    = true;
        window._sddSubmissionId     = key;   // relational PK — drives updateSubmission
        window._scrLoadedKey        = key;
        // _scrLoadedRowNum kept null — we use submission_id now, not sheet row number
        window._scrLoadedRowNum     = null;
        window._sddLastInsertedRow  = null;

        // ── Update cache so picker labels stay fresh ──
        if (!window._scrSavedGroupsByKey) window._scrSavedGroupsByKey = {};
        window._scrSavedGroupsByKey[key] = {
          primary : cleanMain,
          rows    : allRelatedRows,
          mills   : cleanMills,
          ffb_rows: cleanFfb,
        };
        window._scrSavedRowsByKey[key] = cleanMain;

        // NOW render form — buildScrForm() will use _loadedPrimarySddRow to determine button type
        renderSavedRecordImportedLikeInitial(cleanMain, allRelatedRows);

        initScrForm();
        if (typeof window._rowToScrData === 'function' && typeof window._fillScrFormFromData === 'function') {
          const loadedScr = window._rowToScrData(cleanMain);
          setTimeout(function() {
            window._fillScrFormFromData(loadedScr, 'from SDD Data', null, key);
            renderSavedScreeningListView(
              Object.keys(window._scrSavedRowsByKey || {}).map(function(k) {
                return {
                  key: k,
                  primary: window._scrSavedRowsByKey[k],
                  label: getSavedScreeningLabel(window._scrSavedRowsByKey[k]),
                };
              }),
              key
            );
            if (typeof window.refreshSddApproverStaffTraceChrome === 'function') {
              window.refreshSddApproverStaffTraceChrome();
            }
            if (typeof window.refreshSddBossDecisionViewer === 'function') {
              window.refreshSddBossDecisionViewer();
            }
            if (typeof window.showSddToast === 'function') {
              window.showSddToast('Loaded saved screening. Continue editing and save again.', 'success');
            }
          }, 0);
        }
      } else {
        window._pendingScrLoadKey = key;
        if (typeof window.showSddToast === 'function') {
          window.showSddToast('Saved screening selected. Upload/open supplier form to load editor.', 'info');
        }
      }
    } catch (e) {
      if (typeof window.showSddToast === 'function') {
        window.showSddToast('Failed to load submission: ' + (e.message || e), 'error');
      }
    }
  };

  window.loadSavedScrFromDropdownGlobal = function() {
    const sel = document.getElementById('scr-saved-select');
    if (!sel || !sel.value) return;
    window.loadSavedScrByKeyGlobal(sel.value);
  };

  // ─── INIT INTERAKSI FORM ─────────────────────────────────────────────────────
  function initScrForm() {
    function newGrvRow() {
      return `<tr>
        <td><input type="text" class="grv-source" placeholder="Source"></td>
        <td><input type="text" class="grv-desc"   placeholder="Description"></td>
        <td><input type="text" class="grv-pub"    placeholder="Publisher"></td>
        <td><input type="text" class="grv-date"   placeholder="YYYY-MM-DD"></td>
        <td><input type="text" class="grv-status" placeholder="Status"></td>
        <td><input type="text" class="grv-attach" placeholder="Link/filename"></td>
        <td><button type="button" class="scr-del-row">✕</button></td>
      </tr>`;
    }
    function newPriRow() {
      return `<tr>
        <td><input type="text" class="pri-company" placeholder="Company"></td>
        <td><input type="text" class="pri-desc"    placeholder="Description"></td>
        <td><input type="text" class="pri-pub"     placeholder="Publisher"></td>
        <td><input type="text" class="pri-date"    placeholder="YYYY-MM-DD"></td>
        <td><input type="text" class="pri-attach"  placeholder="Link/filename"></td>
        <td><input type="text" class="pri-action"  placeholder="Action Request"></td>
        <td><button type="button" class="scr-del-row">✕</button></td>
      </tr>`;
    }

    // Delete row
    ['scr-grv-tbody','scr-pri-tbody'].forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener('click', e => {
        if (e.target.classList.contains('scr-del-row')) {
          const tr = e.target.closest('tr');
          if (tr && tr.parentElement.querySelectorAll('tr').length > 1) tr.remove();
        }
      });
    });

    function tryAddGrvRow() {
      const tb = document.getElementById('scr-grv-tbody');
      if (!tb) return;
      if (tb.querySelectorAll('tr').length >= SCR_GRV_PRI_MAX_ROWS) {
        if (typeof window.showSddToast === 'function') {
          window.showSddToast('Maksimal ' + SCR_GRV_PRI_MAX_ROWS + ' baris grievance.', 'warning');
        }
        return;
      }
      tb.insertAdjacentHTML('beforeend', newGrvRow());
    }
    function tryAddPriRow() {
      const tb = document.getElementById('scr-pri-tbody');
      if (!tb) return;
      if (tb.querySelectorAll('tr').length >= SCR_GRV_PRI_MAX_ROWS) {
        if (typeof window.showSddToast === 'function') {
          window.showSddToast('Maksimal ' + SCR_GRV_PRI_MAX_ROWS + ' baris PRI.', 'warning');
        }
        return;
      }
      tb.insertAdjacentHTML('beforeend', newPriRow());
    }

    const grvAdd = document.getElementById('scr-grv-add');
    if (grvAdd) grvAdd.addEventListener('click', tryAddGrvRow);

    const priAdd = document.getElementById('scr-pri-add');
    if (priAdd) priAdd.addEventListener('click', tryAddPriRow);

    function updateSelColor(el) {
      el.className = 'scr-sel' +
        (el.value==='Yes'||el.value==='Complete'  ? ' scr-sel-yes' :
         el.value==='No'||el.value==='Non Complete'? ' scr-sel-no'  : '');
    }

    function applySddApproverStaffReadOnly(locked) {
      var wrap = document.getElementById('scr-form-wrap');
      if (wrap) wrap.classList.toggle('sdd-approver-readonly', !!locked);
      document.querySelectorAll('#scr-form-wrap input, #scr-form-wrap textarea, #scr-form-wrap select, #scr-form-wrap button').forEach(function(el) {
        if (!el) return;
        if (el.id === 'noteBossDecision') return;
        if (el.id === 'sdd-approver-approve' || el.id === 'sdd-approver-hold' || el.id === 'sdd-approver-reject') return;
        if (el.id === 'scr-grv-add' || el.id === 'scr-pri-add') {
          el.style.display = locked ? 'none' : '';
          el.disabled = !!locked;
          return;
        }
        if (el.classList && el.classList.contains('scr-del-row')) {
          el.style.display = locked ? 'none' : '';
          el.disabled = !!locked;
          return;
        }
        el.disabled = !!locked;
        if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') el.readOnly = !!locked;
      });
      var traceInput = document.getElementById('traceRecInput');
      if (traceInput) { traceInput.readOnly = !!locked; traceInput.disabled = !!locked; }
      var traceUrl = document.getElementById('traceAttachmentInput');
      if (traceUrl) traceUrl.disabled = !!locked;
      document.querySelectorAll('[data-trace-rec-button]').forEach(function(btn) { btn.disabled = !!locked; });
      var traceAddBtn = document.querySelector('button[onclick*="addTraceAttachmentLink"], button[onclick*="_addTraceAttachmentLink"]');
      if (traceAddBtn) traceAddBtn.disabled = !!locked;
    }

    function hideSddApproverPanel() {
      window._sddApproverRecordLoaded = false;
      applySddApproverStaffReadOnly(false);
      ['draft', 'submit', 'delete'].forEach(function(k) {
        var b = document.querySelector('[data-sdd-save="' + k + '"]');
        if (b) b.style.display = '';
      });
      var cancelBtn = document.querySelector('[data-sdd-save="cancel"]');
      if (cancelBtn) cancelBtn.style.display = 'none';
      if (typeof window.refreshSddApproverStaffTraceChrome === 'function') window.refreshSddApproverStaffTraceChrome();
    }

    function syncSddApproverDecisionUI(s, sourceLabel, rowNum, key) {
      var hasKey = String(key || '').trim() !== '';
      var fromSaved = (hasKey || (rowNum != null && rowNum !== '' && rowNum !== 0)) &&
        String(sourceLabel || '').indexOf('SDD') !== -1;
      if (!fromSaved) {
        hideSddApproverPanel();
        return;
      }
      var status = String((s && s.status) || '').toLowerCase();
      var isSubmitted = status === 'submitted';

      window._sddApproverRecordLoaded = true;

      if (isSubmitted) {
        ['draft', 'submit', 'delete'].forEach(function(k) {
          var btn = document.querySelector('[data-sdd-save="' + k + '"]');
          if (btn) btn.style.display = 'none';
        });
        var cancelBtnAD = document.querySelector('[data-sdd-save="cancel"]');
        if (cancelBtnAD) cancelBtnAD.style.display = 'none';
        var scrSaveOk = document.getElementById('scr-save-ok');
        if (scrSaveOk) scrSaveOk.style.display = 'none';
      } else {
        _refreshDecisionChromeForDraft();
      }

      if (typeof window.refreshSddApproverStaffTraceChrome === 'function') {
        window.refreshSddApproverStaffTraceChrome();
      }
    }

    async function cancelSubmittedToDraft() {
      // Use submission_id (relational) rather than _row number
      var sid = window._sddSubmissionId || window._scrLoadedKey || null;
      if (!sid) {
        if (typeof window.showSddToast === 'function') {
          window.showSddToast('Submission ID tidak ditemukan untuk cancel.', 'error');
        }
        return;
      }
      await apiSetSubmissionStatus({
        submission_id : sid,
        scr_status    : 'Draft',
      });
      if (window._loadedPrimarySddRow) window._loadedPrimarySddRow['SCR - Screening Status'] = 'Draft';
      if (window._scrData) window._scrData.status = 'Draft';
      if (typeof window.showSddToast === 'function') {
        window.showSddToast('Status dikembalikan ke Draft. Form bisa diedit lagi.', 'success');
      }
      await refreshSavedScreeningListGlobal(sid);
      // Full reload so traceability sheet re-renders "Start Screening" (not "View Screening").
      if (typeof window.loadSavedScrByKeyGlobal === 'function') {
        try {
          await window.loadSavedScrByKeyGlobal(sid);
        } catch (_e) {
          syncSubmittedStaffLockUI({ status: 'Draft' }, sid);
          _refreshDecisionChromeForDraft();
        }
      } else {
        syncSubmittedStaffLockUI({ status: 'Draft' }, sid);
        _refreshDecisionChromeForDraft();
      }
    }

    function syncSubmittedStaffLockUI(s, sidOrRowNum) {
      // fromSaved = true when we have either a submission_id or a legacy row num
      var fromSaved = !!(sidOrRowNum != null && sidOrRowNum !== '' && sidOrRowNum !== 0);
      var isSubmitted = String((s && s.status) || '').toLowerCase() === 'submitted';
      var cancelBtn = document.querySelector('[data-sdd-save="cancel"]');
      if (isSubmitted && fromSaved) {
        applySddApproverStaffReadOnly(true);
        if (typeof window.refreshSddPostSubmitDecisionChrome === 'function') {
          window.refreshSddPostSubmitDecisionChrome();
          setTimeout(function() {
            if (typeof window.refreshSddPostSubmitDecisionChrome === 'function') {
              window.refreshSddPostSubmitDecisionChrome();
            }
          }, 0);
        }
        ['draft', 'submit', 'delete'].forEach(function(k) {
          var btn = document.querySelector('[data-sdd-save="' + k + '"]');
          if (btn) btn.style.display = 'none';
        });
        if (cancelBtn) {
          cancelBtn.style.display = '';
          cancelBtn.onclick = async function() {
            if (!confirm('Ubah status Submitted menjadi Draft?')) return;
            cancelBtn.disabled = true;
            var oldText = cancelBtn.textContent;
            cancelBtn.textContent = 'Saving...';
            try {
              await cancelSubmittedToDraft();
            } catch (e) {
              if (typeof window.showSddToast === 'function') window.showSddToast('Cancel to Draft gagal: ' + (e.message || e), 'error');
            } finally {
              cancelBtn.disabled = false;
              cancelBtn.textContent = oldText;
            }
          };
        }
        // Show PDF export button
        var existingPdfBtn = document.getElementById('sdd-export-pdf-btn');
        if (!existingPdfBtn) {
          var pdfBtn = document.createElement('button');
          pdfBtn.id = 'sdd-export-pdf-btn';
          pdfBtn.type = 'button';
          pdfBtn.textContent = '⬇ Export PDF';
          pdfBtn.style.cssText = 'padding:9px 20px;border-radius:8px;border:none;background:#8B1A1A;color:#fff;font-size:13px;font-weight:600;font-family:Inter,sans-serif;cursor:pointer;box-shadow:0 2px 8px rgba(139,26,26,0.2);letter-spacing:0.2px;';
          pdfBtn.onmouseenter = function() { this.style.background = '#6e1414'; };
          pdfBtn.onmouseleave = function() { this.style.background = '#8B1A1A'; };
          pdfBtn.onclick = function() { sddExportPdf(); };
          // Insert after cancel button or append to action row
          var actionRow = cancelBtn ? cancelBtn.parentNode : null;
          if (actionRow) {
            actionRow.appendChild(pdfBtn);
          } else {
            var panelBox = document.querySelector('#panel-supplier-dd .panel-box');
            if (panelBox) panelBox.appendChild(pdfBtn);
          }
        } else {
          existingPdfBtn.style.display = '';
        }
      } else {
        applySddApproverStaffReadOnly(false);
        ['draft', 'submit', 'delete'].forEach(function(k) {
          var btn = document.querySelector('[data-sdd-save="' + k + '"]');
          if (btn) btn.style.display = '';
        });
        if (cancelBtn) cancelBtn.style.display = 'none';

        if (fromSaved) {
          _refreshDecisionChromeForDraft();
          var existingPdfDraft = document.getElementById('sdd-export-pdf-btn');
          if (!existingPdfDraft) {
            var pdfBtnD = document.createElement('button');
            pdfBtnD.id = 'sdd-export-pdf-btn';
            pdfBtnD.type = 'button';
            pdfBtnD.textContent = '⬇ Export PDF';
            pdfBtnD.style.cssText = 'padding:9px 20px;border-radius:8px;border:none;background:#8B1A1A;color:#fff;font-size:13px;font-weight:600;font-family:Inter,sans-serif;cursor:pointer;box-shadow:0 2px 8px rgba(139,26,26,0.2);letter-spacing:0.2px;';
            pdfBtnD.onmouseenter = function() { this.style.background = '#6e1414'; };
            pdfBtnD.onmouseleave = function() { this.style.background = '#8B1A1A'; };
            pdfBtnD.onclick = function() { sddExportPdf(); };
            var actionRowD = document.getElementById('sdd-bottom-actions-row');
            if (actionRowD) actionRowD.appendChild(pdfBtnD);
            else {
              var panelBoxD = document.querySelector('#panel-supplier-dd .panel-box');
              if (panelBoxD) panelBoxD.appendChild(pdfBtnD);
            }
          } else {
            existingPdfDraft.style.display = '';
          }
        } else {
          var pdfBtnDraft = document.getElementById('sdd-export-pdf-btn');
          if (pdfBtnDraft) pdfBtnDraft.style.display = 'none';
          // fromSaved=false means fresh Excel import — do NOT hide the decision wrap.
          // The setTimeout in supplierDisplayExcelData will show it. Only the PDF btn is hidden.
        }
      }
    }

    function rowToScrData(row) {
      const s = {
        owners: row['SCR - List Group By Owners'] || '',
        news: row['SCR - Previous News'] || '',
        supplyto: row['SCR - Supply To'] || '',
        legality: row['SCR - Legality Status'] || '',
        cert: row['SCR - Certification'] || '',
        ndpe: row['SCR - NDPE Policy'] || '',
        nbl: row['SCR - No Buy List'] || '',
        nblCheckResult: row['SCR - NBL Check Result'] || '',
        nblCheckDetail: row['SCR - NBL Match Detail'] || '',
        nblCheckedAt: row['SCR - NBL Checked At'] || '',
        grvYN: row['SCR - Grievance (Y/N)'] || '',
        priYN: row['SCR - PRI (Y/N)'] || '',
        traceNote: row['SCR - Notes'] || '',
        requestedData: row['SCR - Requested Data'] || row['Requested Data'] || '',
        attachments: [],
        status: row['SCR - Screening Status'] || '',
        date: row['SCR - Last Updated'] || '',
        noteSdd: row['noteSDD'] || row['noteSdd'] || row['Note SDD'] || '',
        noteBossDecision: row['noteBossDecision'] || row['noteSDD'] || row['noteSdd'] || '',
        statusSdd: row['statusSDD'] || row['statusSdd'] || row['Status SDD'] ||
          row['statusBossDecision'] || row['Status Boss Decision'] ||
          row['StatusSDD'] || row['STATUSSDD'] || '',
        grvRows: [],
        priRows: []
      };
      for (let i = 1; i <= SCR_GRV_PRI_MAX_ROWS; i++) {
        const g = {
          source: row['GRV' + i + ' - Source'] || '',
          desc: row['GRV' + i + ' - Description'] || '',
          pub: row['GRV' + i + ' - Publisher'] || '',
          date: row['GRV' + i + ' - Date Publish'] || '',
          status: row['GRV' + i + ' - Status'] || '',
          attach: row['GRV' + i + ' - Attachment'] || ''
        };
        if (Object.values(g).some(v => String(v || '').trim() !== '')) s.grvRows.push(g);
      }
      for (let i = 1; i <= SCR_GRV_PRI_MAX_ROWS; i++) {
        const p = {
          company: row['PRI' + i + ' - Company'] || '',
          desc: row['PRI' + i + ' - Description'] || '',
          pub: row['PRI' + i + ' - Publisher'] || '',
          date: row['PRI' + i + ' - Date Publish'] || '',
          attach: row['PRI' + i + ' - Attachment'] || '',
          action: row['PRI' + i + ' - Action Request'] || ''
        };
        if (Object.values(p).some(v => String(v || '').trim() !== '')) s.priRows.push(p);
      }
      if (!s.grvRows.length) s.grvRows = [{}];
      if (!s.priRows.length) s.priRows = [{}];
      return s;
    }

    function fillFormFromScrData(s, sourceLabel, rowNum, key) {
      const set = (id, v) => { const el = document.getElementById(id); if (el) el.value = v || ''; };
      set('scr-owners', s.owners); set('scr-news', s.news); set('scr-supplyto', s.supplyto);
      set('scr-legality', s.legality); set('scr-cert', s.cert);
      set('scr-ndpe', s.ndpe); set('scr-nbl', s.nbl);
      set('scr-grv-yn', s.grvYN); set('scr-pri-yn', s.priYN);
      set('traceRecInput', s.traceNote || '');
      set('requestedDataInput', s.requestedData || '');
      set('noteBossDecision', s.noteBossDecision || s.noteSdd || '');
      window._traceAttachments = s.attachments || [];
      window._renderTraceAttachments && window._renderTraceAttachments();

      const grvTb = document.getElementById('scr-grv-tbody');
      if (grvTb) grvTb.innerHTML = (s.grvRows && s.grvRows.length ? s.grvRows : [{}]).map(function(r) {
        return `<tr>
          <td><input type="text" class="grv-source" value="${(r.source||'').replace(/"/g,'&quot;')}" placeholder="Source"></td>
          <td><input type="text" class="grv-desc" value="${(r.desc||'').replace(/"/g,'&quot;')}" placeholder="Description"></td>
          <td><input type="text" class="grv-pub" value="${(r.pub||'').replace(/"/g,'&quot;')}" placeholder="Publisher"></td>
          <td><input type="text" class="grv-date" value="${(r.date||'').replace(/"/g,'&quot;')}" placeholder="YYYY-MM-DD"></td>
          <td><input type="text" class="grv-status" value="${(r.status||'').replace(/"/g,'&quot;')}" placeholder="Status"></td>
          <td><input type="text" class="grv-attach" value="${(r.attach||'').replace(/"/g,'&quot;')}" placeholder="Link/filename"></td>
          <td><button type="button" class="scr-del-row">✕</button></td>
        </tr>`;
      }).join('');
      const priTb = document.getElementById('scr-pri-tbody');
      if (priTb) priTb.innerHTML = (s.priRows && s.priRows.length ? s.priRows : [{}]).map(function(r) {
        return `<tr>
          <td><input type="text" class="pri-company" value="${(r.company||'').replace(/"/g,'&quot;')}" placeholder="Company"></td>
          <td><input type="text" class="pri-desc" value="${(r.desc||'').replace(/"/g,'&quot;')}" placeholder="Description"></td>
          <td><input type="text" class="pri-pub" value="${(r.pub||'').replace(/"/g,'&quot;')}" placeholder="Publisher"></td>
          <td><input type="text" class="pri-date" value="${(r.date||'').replace(/"/g,'&quot;')}" placeholder="YYYY-MM-DD"></td>
          <td><input type="text" class="pri-attach" value="${(r.attach||'').replace(/"/g,'&quot;')}" placeholder="Link/filename"></td>
          <td><input type="text" class="pri-action" value="${(r.action||'').replace(/"/g,'&quot;')}" placeholder="Action Request"></td>
          <td><button type="button" class="scr-del-row">✕</button></td>
        </tr>`;
      }).join('');

      ['scr-legality','scr-ndpe','scr-nbl','scr-grv-yn','scr-pri-yn'].forEach(id => {
        const el = document.getElementById(id); if (el) updateSelColor(el);
      });
      const grvWrap = document.getElementById('scr-grv-wrap');
      if (grvWrap) grvWrap.style.display = s.grvYN === 'Yes' ? 'block' : 'none';
      const priWrap = document.getElementById('scr-pri-wrap');
      if (priWrap) priWrap.style.display = s.priYN === 'Yes' ? 'block' : 'none';

      window._scrLoadedRowNum = rowNum || null;
      window._scrLoadedKey = key || '';
      window._scrData = s;
      if (window._loadedPrimarySddRow && typeof window.restoreNblCheckResultFromRow_ === 'function') {
        window.restoreNblCheckResultFromRow_(window._loadedPrimarySddRow);
      }
      syncSddApproverDecisionUI(s, sourceLabel, rowNum, key);
      // Use submission_id (key) as the "fromSaved" signal; fall back to legacy rowNum
      syncSubmittedStaffLockUI(s, key || rowNum || null);
      const saveOk = document.getElementById('scr-save-ok');
      if (saveOk && s.status) {
        saveOk.style.color = '#059669';
        saveOk.textContent = '✓ Loaded: ' + s.status + (s.date ? ' (' + s.date + ')' : '') + (sourceLabel ? ' · ' + sourceLabel : '');
        saveOk.style.display = 'inline';
      }
    }
    window._rowToScrData = rowToScrData;
    window._fillScrFormFromData = fillFormFromScrData;

    function refreshSavedScreeningList(selectKey) {
      if (typeof window.refreshSavedScreeningListGlobal === 'function') {
        return window.refreshSavedScreeningListGlobal(selectKey);
      }
      return Promise.resolve();
    }

    window._loadSavedScrFromDropdown = function() {
      if (typeof window.loadSavedScrFromDropdownGlobal === 'function') {
        return window.loadSavedScrFromDropdownGlobal();
      }
    };
    window.exitSavedScrModeGlobal = function(silent) {
      window._scrLoadedRowNum = null;
      window._scrLoadedKey = '';
      window._loadedPrimarySddRow = null;
      window._sddIsLoadedSaved = false;
      window._sddLastInsertedRow = null;
      window._sddSubmissionId = null;   // relational PK cleared on exit
      window._scrData = {};
      window._traceAttachments = [];
      fillFormFromScrData({ grvRows:[{}], priRows:[{}], attachments:[] }, 'reset', null, '');
      const sel = document.getElementById('scr-saved-select');
      if (sel) sel.value = '';
      renderSavedScreeningListView(
        Object.keys(window._scrSavedRowsByKey || {}).map(function(k) {
          return { key: k, primary: window._scrSavedRowsByKey[k], label: getSavedScreeningLabel(window._scrSavedRowsByKey[k]) };
        }),
        ''
      );
      const saveOk = document.getElementById('scr-save-ok');
      if (saveOk) saveOk.style.display = 'none';
      if (supplierWorkbook) {
        supplierDisplayAllSheets();
      } else {
        const holder = document.getElementById('supplierExcelData');
        if (holder) holder.innerHTML = '<p style="color:#9C8080;font-size:13px;padding:16px;">Select saved screening or import an Excel file to view data.</p>';
      }
      if (!silent && typeof window.showSddToast === 'function') {
        window.showSddToast('Exited load mode. You are back to import/select view.', 'info');
      }
      if (typeof window.refreshSddApproverStaffTraceChrome === 'function') window.refreshSddApproverStaffTraceChrome();
    };

    ['scr-legality','scr-ndpe','scr-nbl','scr-grv-yn','scr-pri-yn'].forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener('change', () => updateSelColor(el));
    });

    const grvYnEl = document.getElementById('scr-grv-yn');
    if (grvYnEl) grvYnEl.addEventListener('change', function() {
      document.getElementById('scr-grv-wrap').style.display = this.value==='Yes' ? 'block' : 'none';
    });
    const priYnEl = document.getElementById('scr-pri-yn');
    if (priYnEl) priYnEl.addEventListener('change', function() {
      document.getElementById('scr-pri-wrap').style.display = this.value==='Yes' ? 'block' : 'none';
    });

    const refBtn = document.getElementById('scr-refresh-list');
    if (refBtn) refBtn.addEventListener('click', function() { refreshSavedScreeningList(); });
    const exitBtn = document.getElementById('scr-exit-load');
    if (exitBtn) exitBtn.addEventListener('click', function() {
      if (typeof window.exitSavedScrModeGlobal === 'function') window.exitSavedScrModeGlobal(false);
    });

    window._traceAttachments = window._traceAttachments || [];
    window._renderTraceAttachments = function() {
      const preview = document.getElementById('traceAttachmentPreview');
      if (!preview) return;
      if (!window._traceAttachments.length) {
        preview.innerHTML = 'No attachment yet.';
        return;
      }
      preview.innerHTML = window._traceAttachments.map(function(file, idx) {
        const isImage = file.url && /\.(jpe?g|png|gif|webp|svg)$/i.test(file.url);
        const thumb = isImage ? '<img src="' + file.url + '" style="max-width:140px;max-height:120px;border-radius:12px;object-fit:cover;border:1px solid rgba(74,28,28,0.12);">' : '<div style="padding:18px 14px;border-radius:12px;border:1px solid rgba(74,28,28,0.12);background:#f8fafc;color:#111;font-size:13px;font-weight:600;min-width:120px;text-align:center;">Link</div>';
        return '<div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;border:1px solid rgba(226,232,240,0.9);border-radius:14px;padding:12px;">'
          + thumb
          + '<div style="flex:1;min-width:160px;">'
          + '<div style="font-size:13px;font-weight:600;color:#111;margin-bottom:4px;">' + file.name + '</div>'
          + '<div style="font-size:12px;color:#6b7280;margin-bottom:10px;word-break:break-all;">' + file.url + '</div>'
          + '<div style="display:flex;gap:8px;flex-wrap:wrap;">'
          + '<a href="' + file.url + '" target="_blank" style="padding:8px 14px;border-radius:8px;background:#111;color:#fff;text-decoration:none;font-size:12px;">Open</a>'
          + '<button type="button" onclick="window._removeTraceAttachment(' + idx + ')" style="padding:9px 20px;border-radius:8px;border:1.5px solid #E5E7EB;background:#fff;color:#374151;font-size:13px;font-weight:600;font-family:Inter,sans-serif;cursor:pointer;">Remove</button>'
          + '</div>'
          + '</div>'
          + '</div>';
      }).join('');
    };

    window._removeTraceAttachment = function(index) {
      if (!window._traceAttachments) return;
      window._traceAttachments.splice(index, 1);
      window._renderTraceAttachments();
    };

    window._addTraceAttachmentLink = function() {
      const input = document.getElementById('traceAttachmentInput');
      if (!input) return;
      const url = (input.value || '').trim();
      if (!url) return;
      const name = url.split('/').pop().split('?')[0] || 'Attachment';
      window._traceAttachments = window._traceAttachments || [];
      window._traceAttachments.push({ name: name, url: url });
      input.value = '';
      window._renderTraceAttachments();
    };

    window._renderTraceAttachments();

    if (typeof window.refreshSddApproverStaffTraceChrome === 'function') {
      window.refreshSddApproverStaffTraceChrome();
    }

    // Save flow: source of truth is Google Sheets (no localStorage persistence for screening data)
    window._saveScrScreening = async function(status) {
      if (window._sddSaveInFlight) return;
      if (typeof location !== 'undefined' && location.protocol === 'file:') {
        var fileMsg = 'Cannot save while opened as file://. Please run this HTML from http://localhost (for example with Live Server) and try again.';
        if (typeof window.showSddToast === 'function') window.showSddToast(fileMsg, 'error');
        if (typeof window.showSddNotification === 'function') window.showSddNotification('Save Blocked', fileMsg, 'error');
        return;
      }
      var _carryDecision = '';
      if (window._scrData && String(window._scrData.statusSdd || '').trim()) {
        _carryDecision = String(window._scrData.statusSdd).trim();
      } else if (window._loadedPrimarySddRow) {
        _carryDecision = String(
          window._loadedPrimarySddRow['statusSDD'] || window._loadedPrimarySddRow['statusSdd'] ||
          window._loadedPrimarySddRow['statusBossDecision'] || ''
        ).trim();
      }
      const scrData = {
        owners:   document.getElementById('scr-owners')?.value||'',
        news:     document.getElementById('scr-news')?.value||'',
        supplyto: document.getElementById('scr-supplyto')?.value||'',
        legality: document.getElementById('scr-legality')?.value||'',
        cert:     document.getElementById('scr-cert')?.value||'',
        ndpe:     document.getElementById('scr-ndpe')?.value||'',
        nbl:      document.getElementById('scr-nbl')?.value||'',
        nblCheckResult: (window._scrData && window._scrData.nblCheckResult) || '',
        nblCheckDetail: (window._scrData && window._scrData.nblCheckDetail) || '',
        nblCheckedAt: (window._scrData && window._scrData.nblCheckedAt) || '',
        grvYN:    document.getElementById('scr-grv-yn')?.value||'',
        priYN:    document.getElementById('scr-pri-yn')?.value||'',
        traceNote: document.getElementById('traceRecInput')?.value||'',
        requestedData: document.getElementById('requestedDataInput')?.value||'',
        noteSdd:  document.getElementById('noteBossDecision')?.value||document.getElementById('noteSDD')?.value||'',
        noteBossDecision: document.getElementById('noteBossDecision')?.value||document.getElementById('noteSDD')?.value||'',
        statusSdd: _carryDecision,
        statusBossDecision: _carryDecision,
        attachments: window._traceAttachments || [],
        grvRows: [], priRows: [],
      };
      if (scrData.nbl && !scrData.nblCheckResult) {
        scrData.nblCheckResult = scrData.nbl === 'Yes'
          ? 'YES — Supplier IS ON the No Buy List (NBL)'
          : 'NO — Supplier is NOT on the No Buy List';
      }
      if (scrData.nbl && !scrData.nblCheckDetail && window._nblCheckResult) {
        scrData.nblCheckDetail = window._scrData.nblCheckDetail
          || (window._nblCheckResult.matches && window._nblCheckResult.matches.length
            ? window._nblCheckResult.matches.map(function(m) { return m.source + ': ' + m.detail; }).join(' | ')
            : (scrData.nbl === 'No'
              ? 'No matching Group Name, Company Name, or Mill Name in NBL or Unilever NBL sheets.'
              : 'Supplier matched a name on the NBL or Unilever NBL registry.'));
        scrData.nblCheckedAt = scrData.nblCheckedAt || window._scrData.nblCheckedAt || window._nblCheckResult.checkedAt || '';
      }
      document.querySelectorAll('#scr-grv-tbody tr').forEach(tr => {
        scrData.grvRows.push({
          source: tr.querySelector('.grv-source')?.value||'',
          desc:   tr.querySelector('.grv-desc')?.value||'',
          pub:    tr.querySelector('.grv-pub')?.value||'',
          date:   tr.querySelector('.grv-date')?.value||'',
          status: tr.querySelector('.grv-status')?.value||'',
          attach: tr.querySelector('.grv-attach')?.value||'',
        });
      });
      document.querySelectorAll('#scr-pri-tbody tr').forEach(tr => {
        scrData.priRows.push({
          company: tr.querySelector('.pri-company')?.value||'',
          desc:    tr.querySelector('.pri-desc')?.value||'',
          pub:     tr.querySelector('.pri-pub')?.value||'',
          date:    tr.querySelector('.pri-date')?.value||'',
          attach:  tr.querySelector('.pri-attach')?.value||'',
          action:  tr.querySelector('.pri-action')?.value||'',
        });
      });

      if (status === 'delete') {
        if (!confirm('Yakin ingin delete data screening ini?')) return;
        // ── DELETE via relational deleteSubmission API ──
        const sidToDelete = window._sddSubmissionId || window._scrLoadedKey || null;
        if (!sidToDelete) {
          if (typeof window.showSddToast === 'function') {
            window.showSddToast('Tidak ada submission_id. Muat screening tersimpan dulu, lalu hapus.', 'error');
          }
          return;
        }
        try {
          await apiDeleteSubmission({ submission_id: sidToDelete });
        } catch (e) {
          if (typeof window.showSddToast === 'function') window.showSddToast('Delete gagal: ' + (e.message || e), 'error');
          return;
        }
        await refreshSavedScreeningListGlobal('');
        supplierWorkbook = null;
        window._sddImportedRows = [];
        window._sddImportFirstRow = null;
        window._sddSubmissionId = null;
        const fileInput = document.getElementById('supplierExcelFile');
        if (fileInput) fileInput.value = '';
        const sheetWrap = document.getElementById('supplierSheetSelectContainer');
        if (sheetWrap) sheetWrap.style.display = 'none';
        if (typeof window.exitSavedScrModeGlobal === 'function') {
          window.exitSavedScrModeGlobal(true);
        } else {
          fillFormFromScrData({ grvRows: [{}], priRows: [{}], attachments: [] }, 'reset', null, '');
        }
        if (typeof window.showSddToast === 'function') window.showSddToast('Data screening berhasil dihapus.', 'success');
        if (typeof window.showSddNotification === 'function') {
          window.showSddNotification('Delete Success', 'Data screening berhasil dihapus.', 'success');
        }
        return;
      }

      scrData.status = status === 'submit' ? 'Submitted' : 'Draft';
      scrData.date = new Date().toLocaleDateString('id-ID');
      window._scrData = scrData;
      var saveOk = document.getElementById('scr-save-ok');
      try {
        window._sddSaveInFlight = true;
        var apiResult = await handleFinalSave(scrData);
        var importedCount = (apiResult && typeof apiResult.importedSyncedRows === 'number')
          ? apiResult.importedSyncedRows
          : (Array.isArray(window._sddImportedRows) ? window._sddImportedRows.length : 0);
        // Capture sid BEFORE state-clear: createSubmission returns submission_id; updateSubmission often does not.
        var savedSid = (apiResult && apiResult.submission_id)
          ? String(apiResult.submission_id)
          : (window._sddSubmissionId ? String(window._sddSubmissionId) : '');
        var okMsg = 'Data berhasil disimpan ke Google Sheets sebagai ' + scrData.status + '.';
        if (savedSid) okMsg += ' Submission ID: ' + savedSid + '.';
        if (importedCount > 0) okMsg += ' ' + importedCount + ' baris tersinkron.';
        if (typeof window.showSddToast === 'function') window.showSddToast(okMsg, 'success');
        if (typeof window.showSddNotification === 'function') window.showSddNotification('Tersimpan', okMsg, 'success');

        // ── POST-SAVE: Refresh list dan load ulang jika submit ──
        // Jika status=submit dan ada savedSid, load ulang form TANPA reset state dulu
        // sehingga tombol berubah dari "Start Screening" ke "View Screening" otomatis.
        // Jika status=draft, reset dan exit seperti biasanya.
        
        // Bersihkan sessionStorage backup
        try { sessionStorage.removeItem('sdd_form_html_backup'); sessionStorage.removeItem('sdd_form_html_backup_key'); } catch(e) {}

        // Update list (untuk refresh status di dropdown)
        await refreshSavedScreeningListGlobal(savedSid || '');

        if (status === 'submit' && savedSid && typeof window.loadSavedScrByKeyGlobal === 'function') {
          // For SUBMIT: Load ulang form dari server dengan status "Submitted".
          // Form akan ter-render dengan tombol "View Screening" biru dan read-only UI.
          try {
            await window.loadSavedScrByKeyGlobal(savedSid);
            // loadSavedScrByKeyGlobal akan handle state set & UI refresh internally
          } catch (_reloadErr) {
            console.error('Failed to reload submitted screening:', _reloadErr);
            // Fallback: reset state dan exit
            if (typeof window.exitSavedScrModeGlobal === 'function') window.exitSavedScrModeGlobal(true);
          }
        } else {
          // For DRAFT atau non-submit: Reset ke state awal dan show import view
          window._sddIsLoadedSaved    = false;
          window._sddLastInsertedRow  = null;
          window._sddSubmissionId     = null;   // relational PK cleared after save
          window._scrLoadedRowNum     = null;
          window._scrLoadedKey        = '';
          window._loadedPrimarySddRow = null;
          window._sddImportedRows     = [];
          window._sddImportFirstRow   = null;
          window._scrData             = {};
          window._traceAttachments    = [];
          // Anti-leak: clear in-memory screening data supaya import/submission berikutnya bersih.
          window._tmlScreeningData    = {};
          window._ffbScreeningData    = {};

          supplierWorkbook = null;
          const fileInputAfterSave = document.getElementById('supplierExcelFile');
          if (fileInputAfterSave) fileInputAfterSave.value = '';
          const sheetWrapAfterSave = document.getElementById('supplierSheetSelectContainer');
          if (sheetWrapAfterSave) sheetWrapAfterSave.style.display = 'none';

          if (typeof window.exitSavedScrModeGlobal === 'function') {
            window.exitSavedScrModeGlobal(true);
          } else {
            const holder = document.getElementById('supplierExcelData');
            if (holder) holder.innerHTML = '<p style="color:#9C8080;font-size:13px;padding:16px;">Data berhasil disimpan. Pilih dari list di atas untuk melanjutkan edit, atau import Excel baru.</p>';
          }
        }
      } catch (e) {
        var errMsg = (e && e.message) ? e.message : String(e);
        console.error('SDD save failed:', e);
        if (typeof window.showSddToast === 'function') {
          window.showSddToast('Could not save to Google Sheets: ' + errMsg, 'error');
        }
        if (saveOk) {
          saveOk.textContent = '✗ Save failed — see message above';
          saveOk.style.color = '#dc2626';
          saveOk.style.display = 'block';
        }
        if (typeof window.showSddNotification === 'function') window.showSddNotification('Save Failed', 'Could not save to Google Sheets: ' + errMsg, 'error');
      } finally {
        window._sddSaveInFlight = false;
      }
    };

    // Do not hydrate screening form from localStorage; always rely on Google Sheets / loaded records.

    updateSavedScreeningPickerUI();
    refreshSavedScreeningList(window._pendingScrLoadKey || '');
    if (window._pendingScrLoadKey) {
      window.loadSavedScrFromDropdownGlobal && window.loadSavedScrFromDropdownGlobal();
      window._pendingScrLoadKey = '';
    }
  }

/** Fallback web app URL — override with window.SDD_WEBAPP_URL or localStorage SDD_WEBAPP_URL (full …/exec URL). */
var SDD_DEFAULT_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbwXJXN-vATCurs7C_OLEa-z1-i-qTgb6j37HBsL3MbnW7XFCuLL2X5xG26A1H6v5ilxyg/exec';

function getSddApiUrl() {
  var custom = (typeof window !== 'undefined' && window.SDD_WEBAPP_URL) || '';
  if (typeof custom === 'string' && custom.indexOf('http') === 0) {
    var u = custom.trim();
    if (u.indexOf('/exec') === -1) u = u.replace(/\/?$/, '') + '/exec';
    return u;
  }
  try {
    var ls = localStorage.getItem('SDD_WEBAPP_URL');
    if (ls && ls.indexOf('http') === 0) {
      var u2 = ls.trim();
      if (u2.indexOf('/exec') === -1) u2 = u2.replace(/\/?$/, '') + '/exec';
      return u2;
    }
  } catch (e) {}
  return SDD_DEFAULT_WEBAPP_URL;
}

/** Kept for optional checks; do not block saves — use http(s) + text/plain POST for GAS compatibility. */
function assertNotFileProtocol() {
  if (typeof location !== 'undefined' && location.protocol === 'file:') {
    console.warn('SDD: file:// may block reading Google responses. Prefer opening this page via http://localhost or https://');
  }
}

window.showSddToast = function(message, type) {
  type = type || 'info';
  var el = document.getElementById('sdd-global-toast');
  if (!el) {
    el = document.createElement('div');
    el.id = 'sdd-global-toast';
    el.setAttribute('role', 'status');
    el.style.cssText = 'position:fixed;bottom:24px;right:24px;max-width:min(420px,calc(100vw - 48px));padding:14px 18px;border-radius:12px;font-size:13px;font-weight:500;z-index:999999;box-shadow:0 8px 32px rgba(0,0,0,0.18);font-family:Inter,sans-serif;line-height:1.45;transition:opacity .25s;';
    document.body.appendChild(el);
  }
  el.style.background = type === 'error' ? '#fef2f2' : type === 'success' ? '#ecfdf5' : '#f1f5f9';
  el.style.color = type === 'error' ? '#991b1b' : type === 'success' ? '#065f46' : '#0f172a';
  el.style.border = type === 'error' ? '1px solid #fecaca' : type === 'success' ? '1px solid #a7f3d0' : '1px solid #e2e8f0';
  el.textContent = message;
  el.style.opacity = '1';
  clearTimeout(window.showSddToast._t);
  window.showSddToast._t = setTimeout(function() {
    el.style.opacity = '0';
  }, 7000);
};

window.showSddNotification = function(title, message, type) {
  type = type || 'info';
  let overlay = document.getElementById('sdd-notification-overlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'sdd-notification-overlay';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,0.4);z-index:1000000;display:none;align-items:center;justify-content:center;padding:20px;';
    overlay.innerHTML = '<div id="sdd-notification-modal" style="width:min(520px,92vw);background:#fff;border-radius:16px;box-shadow:0 20px 48px rgba(0,0,0,0.28);border:1px solid #e5e7eb;overflow:hidden;">'
      + '<div id="sdd-notification-head" style="padding:14px 16px;font-size:14px;font-weight:700;"></div>'
      + '<div style="padding:14px 16px 18px;"><div id="sdd-notification-msg" style="font-size:13px;line-height:1.55;color:#111827;"></div>'
      + '<div style="display:flex;justify-content:flex-end;margin-top:16px;"><button type="button" id="sdd-notification-ok" style="padding:8px 14px;border-radius:8px;border:1px solid #d1d5db;background:#fff;color:#111827;font-size:12px;font-weight:600;cursor:pointer;">OK</button></div></div>'
      + '</div>';
    document.body.appendChild(overlay);
    overlay.addEventListener('click', function(e) {
      if (e.target === overlay) overlay.style.display = 'none';
    });
    const okBtn = overlay.querySelector('#sdd-notification-ok');
    if (okBtn) okBtn.addEventListener('click', function() { overlay.style.display = 'none'; });
  }
  const head = overlay.querySelector('#sdd-notification-head');
  const msg = overlay.querySelector('#sdd-notification-msg');
  if (head) {
    if (type === 'error') {
      head.style.background = '#fef2f2';
      head.style.color = '#991b1b';
      head.style.borderBottom = '1px solid #fecaca';
    } else if (type === 'success') {
      head.style.background = '#ecfdf5';
      head.style.color = '#065f46';
      head.style.borderBottom = '1px solid #a7f3d0';
    } else {
      head.style.background = '#f8fafc';
      head.style.color = '#0f172a';
      head.style.borderBottom = '1px solid #e2e8f0';
    }
    head.textContent = title || 'Notification';
  }
  if (msg) msg.textContent = String(message || '');
  overlay.style.display = 'flex';
};

/** GAS may return a bare JSON array or an object like { data: [...] } / { rows: [...] }. */
function normalizeGetArray(data) {
  if (Array.isArray(data)) return data;
  if (!data || typeof data !== 'object') return [];
  if (Array.isArray(data.data)) return data.data;
  if (Array.isArray(data.rows)) return data.rows;
  if (Array.isArray(data.records)) return data.records;
  if (Array.isArray(data.items)) return data.items;
  if (Array.isArray(data.result)) return data.result;
  return [];
}

async function apiGet(sheet) {
  var url = getSddApiUrl();
  var params = new URLSearchParams({ action: 'getAll', sheet: sheet });
  if (sheet === 'sdd') params.set('_ts', String(Date.now()));
  var fullUrl = url + '?' + params.toString();
  var controller = new AbortController();
  var tid = setTimeout(function () { controller.abort(); }, 45000);
  var res;
  try {
    res = await fetch(fullUrl, {
      method: 'GET',
      mode: 'cors',
      credentials: 'omit',
      redirect: 'follow',
      signal: controller.signal
    });
  } catch (e) {
    clearTimeout(tid);
    var ename = e && e.name;
    var emsg = (e && e.message) || String(e);
    if (ename === 'AbortError') {
      throw new Error('GET timeout (45s). Periksa jaringan atau URL Apps Script (SDD_WEBAPP_URL).');
    }
    if (emsg.indexOf('Failed to fetch') !== -1 || ename === 'TypeError') {
      throw new Error('GET gagal (jaringan/CORS/adblock). Pastikan halaman di http://localhost atau https, dan URL Web App benar.');
    }
    throw new Error('GET failed: ' + emsg);
  }
  clearTimeout(tid);
  var text = await res.text();
  var data;
  try {
    data = text ? JSON.parse(text) : {};
  } catch (parseErr) {
    throw new Error('GET: server did not return JSON (HTTP ' + res.status + '). Check SDD_WEBAPP_URL. Body: ' + text.slice(0, 180));
  }
  if (!res.ok) throw new Error((data && data.error) || ('API GET failed: ' + res.status));
  if (data && typeof data === 'object' && data.error && !Array.isArray(data)) throw new Error(data.error);
  var rows = normalizeGetArray(data);
  if (rows.length === 0 && data && typeof data === 'object' && !Array.isArray(data)) {
    var keys = Object.keys(data);
    if (keys.length && keys.indexOf('success') === -1) {
      console.warn('[apiGet] Response was not an array; keys:', keys.slice(0, 12));
    }
  }
  return rows;
}

async function apiPost(body) {
  var url = getSddApiUrl();
  var payload = JSON.stringify(body);
  console.log('📤 API POST →', url, body);
  /* Google Apps Script: use text/plain so the browser sends a "simple" request (no CORS preflight). doPost still receives JSON in e.postData.contents. */
  var res = await fetch(url, {
    method: 'POST',
    mode: 'cors',
    credentials: 'omit',
    redirect: 'follow',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: payload
  });
  var text = await res.text();
  var data;
  try {
    data = text ? JSON.parse(text) : {};
  } catch (parseErr) {
    throw new Error('POST: server did not return JSON (HTTP ' + res.status + '). Deploy Web app as "Anyone", verify /exec URL. Response: ' + text.slice(0, 220));
  }
  if (!res.ok) throw new Error((data && data.error) || ('API POST failed: HTTP ' + res.status));
  if (data && data.error) throw new Error(data.error);
  return data;
}

async function apiUpsertSDD(data) {
  return apiPost({ action: 'upsertSDD', sheet: 'sdd', data });
}

async function apiInsertSDD(data) {
  return apiPost({ action: 'insertSDD', sheet: 'sdd', data });
}

async function apiUpdateSDD(rowNum, data) {
  var r = Number(rowNum);
  if (!r || r < 2) throw new Error('Invalid SDD row for update: ' + rowNum);
  return apiPost({ action: 'updateSDD', sheet: 'sdd', row: r, data });
}

// ─── RELATIONAL API HELPERS (submission_id-centric) ──────────────────────────

/**
 * POST createSubmission — first save of a new submission.
 * payload: { main: { supplier_type, ...fields }, mills?: [...], ffb_rows?: [...] }
 * Returns { success, submission_id, mills_inserted, ffb_inserted }
 */
async function apiCreateSubmission(payload) {
  return apiPost({ action: 'createSubmission', payload });
}

/**
 * POST updateSubmission — partial update of existing submission.
 * payload: { submission_id, main?: { ...fields }, mills?: [...], ffb_rows?: [...] }
 * Only keys present in main are overwritten; missing keys keep existing values.
 */
async function apiUpdateSubmission(payload) {
  return apiPost({ action: 'updateSubmission', payload });
}

/**
 * POST setSubmissionStatus — lightweight status-only patch.
 * payload: { submission_id, scr_status?, statusSDD?, scr_notes?,
 *            scr_recommendation?, scr_risk_level? }
 */
async function apiSetSubmissionStatus(payload) {
  return apiPost({ action: 'setSubmissionStatus', payload });
}

/**
 * POST deleteSubmission — soft-deletes MAIN + all child rows.
 * payload: { submission_id }
 */
async function apiDeleteSubmission(payload) {
  return apiPost({ action: 'deleteSubmission', payload });
}

/**
 * GET getSubmissionById — canonical hydration source.
 * Returns { success, main, mills, ffb_rows }
 * UI MUST use this response shape; never reconstruct from mixed row heuristics.
 */
async function apiGetSubmissionById(submissionId) {
  var url = getSddApiUrl();
  var params = new URLSearchParams({
    action: 'getSubmissionById',
    submission_id: String(submissionId)
  });
  var res = await fetch(url + '?' + params.toString(), {
    method: 'GET', mode: 'cors', credentials: 'omit', redirect: 'follow'
  });
  var text = await res.text();
  var data;
  try { data = text ? JSON.parse(text) : {}; } catch (e) {
    throw new Error('getSubmissionById: non-JSON response. ' + text.slice(0, 180));
  }
  if (data && data.error) throw new Error(data.error);
  return data; // { success, main, mills, ffb_rows }
}
window.apiGetSubmissionById = apiGetSubmissionById;

/**
 * GET listSubmissions — MAIN rows only (no mixed traceability rows).
 * params: { scr_status?, supplier_type?, page?, page_size? }
 * Returns { success, total, page, page_size, data: [mainRow, ...] }
 */
async function apiListSubmissions(params) {
  var url = getSddApiUrl();
  var qp = Object.assign({ action: 'listSubmissions', _ts: String(Date.now()) }, params || {});
  var res = await fetch(url + '?' + new URLSearchParams(qp).toString(), {
    method: 'GET', mode: 'cors', credentials: 'omit', redirect: 'follow'
  });
  var text = await res.text();
  var data;
  try { data = text ? JSON.parse(text) : {}; } catch (e) {
    throw new Error('listSubmissions: non-JSON response. ' + text.slice(0, 180));
  }
  if (data && data.error) throw new Error(data.error);
  return data; // { success, total, page, page_size, data }
}

// ─────────────────────────────────────────────────────────────────────────────

// ─── SDD DECISION HELPERS (used by initDashboardApp + PDF export) ───────────

function _normalizeDecisionLabel(raw) {
  if (!raw) return '';
  var s = String(raw).trim().toLowerCase();
  if (s === 'approve' || s === 'approved') return 'APPROVED';
  if (s === 'hold' || s === 'on hold') return 'ON HOLD';
  if (s === 'reject' || s === 'rejected') return 'REJECTED';
  return String(raw).trim().toUpperCase();
}

function _readCurrentDecisionRaw() {
  if (window._scrData && String(window._scrData.statusSdd || '').trim()) {
    return String(window._scrData.statusSdd).trim();
  }
  if (window._loadedPrimarySddRow) {
    return String(
      window._loadedPrimarySddRow['statusSDD'] ||
      window._loadedPrimarySddRow['statusSdd'] ||
      window._loadedPrimarySddRow['statusBossDecision'] || ''
    ).trim();
  }
  return '';
}

function _syncDecisionBadge() {
  var badge = document.getElementById('sdd-staff-decision-badge');
  if (!badge) return;
  var raw = _readCurrentDecisionRaw();
  badge.textContent = raw
    ? 'Decision: ' + _normalizeDecisionLabel(raw)
    : 'Decision: — (pilih tombol di bawah)';
}

function _refreshDecisionChromeForDraft() {
  var wrap = document.getElementById('sdd-staff-decision-wrap');
  if (wrap) wrap.style.display = 'none';
  ['sdd-approver-approve', 'sdd-approver-hold', 'sdd-approver-reject'].forEach(function(id) {
    var b = document.getElementById(id);
    if (!b) return;
    b.disabled = true;
    b.style.opacity = '0.5';
    b.style.cursor = 'not-allowed';
    b.style.filter = '';
    b.title = 'Tombol keputusan aktif setelah status Submitted.';
  });
  var nb = document.getElementById('noteBossDecision');
  if (nb) {
    nb.readOnly = false;
    nb.disabled = false;
    nb.style.background = '#fff';
    nb.style.color = '#1A0A0A';
    nb.style.cursor = 'text';
  }
  _syncDecisionBadge();
}
// Expose to window so nested scopes (e.g. supplierDisplayExcelData inside initDashboardApp) can call it
window._refreshDecisionChromeForDraft = _refreshDecisionChromeForDraft;
window._syncDecisionBadge = _syncDecisionBadge;

// ─────────────────────────────────────────────────────────────────────────────

/**
 * Screening decision (Approve / Hold / Reject) — only while SCR status is Draft.
 * Same mental model as screening fields: only updates in-memory state here; the server
 * gets statusSDD / noteBossDecision when the user clicks Save Draft or Submit (create/update).
 */
window._submitSddApproverDecision = function(statusSdd) {
  var sid = window._sddSubmissionId || window._scrLoadedKey || null;
  var scrSt = String(
    (window._scrData && window._scrData.status) ||
    (window._loadedPrimarySddRow && window._loadedPrimarySddRow['SCR - Screening Status']) || ''
  ).trim().toLowerCase();
  if (scrSt !== 'submitted') {
    if (typeof window.showSddToast === 'function') {
      window.showSddToast('Keputusan Approve / On Hold / Reject baru bisa dipilih setelah Submit.', 'info');
    }
    return;
  }

  var ta   = document.getElementById('noteBossDecision') || document.getElementById('noteSDD');
  var note = ta ? String(ta.value || '').trim() : '';

  if (!window._scrData) window._scrData = {};
  window._scrData.statusSdd = statusSdd;
  window._scrData.statusBossDecision = statusSdd;
  window._scrData.noteSdd = note;
  window._scrData.noteBossDecision = note;

  if (window._loadedPrimarySddRow) {
    window._loadedPrimarySddRow['noteSDD'] = note;
    window._loadedPrimarySddRow['noteBossDecision'] = note;
    window._loadedPrimarySddRow['statusSDD'] = statusSdd;
    window._loadedPrimarySddRow['statusBossDecision'] = statusSdd;
  }

  _syncDecisionBadge();
  _refreshDecisionChromeForDraft();
  if (scrSt === 'submitted' && sid) {
    apiSetSubmissionStatus({ submission_id: sid, statusSDD: statusSdd })
      .then(function(statusResp) {
        window._lastTtpSyncResult = statusResp && statusResp.ttp_sync ? statusResp.ttp_sync : null;
        return apiUpdateSubmission({
          submission_id: sid,
          main: {
            noteSDD: note,
            noteBossDecision: note,
            statusBossDecision: statusSdd,
          }
        });
      })
      .then(function() {
        if (typeof window.refreshSavedScreeningListGlobal === 'function') {
          return window.refreshSavedScreeningListGlobal(sid);
        }
      })
      .then(function() {
        if (typeof renderMillTaskList === 'function') renderMillTaskList();
        var saveOk = document.getElementById('scr-save-ok');
        if (saveOk) {
          saveOk.style.display = 'block';
          saveOk.style.color = '#059669';
          saveOk.textContent = '✓ Decision saved: ' + statusSdd + ' · Submission ID: ' + sid;
        }
        if (typeof window.__contactListInvalidate === 'function') window.__contactListInvalidate();
        if (typeof window.__ttpInvalidate === 'function') window.__ttpInvalidate();
        var normDecision = _normalizeDecisionLabel(statusSdd);
        if (normDecision === 'APPROVED') {
          var clsPanel = document.getElementById('panel-contact-list-supplier');
          if (clsPanel && clsPanel.classList.contains('active') && typeof window.loadContactListData === 'function') {
            return window.loadContactListData(true);
          }
          var ttpPanel = document.getElementById('panel-ttm-ttp');
          if (ttpPanel && ttpPanel.classList.contains('active') && typeof loadTTPData === 'function') {
            return loadTTPData();
          }
        }
      })
      .then(function() {
        if (typeof window.showSddToast === 'function') {
          var normToast = _normalizeDecisionLabel(statusSdd);
          var extraCls = normToast === 'APPROVED'
            ? ' Kontak Sustainability PIC disimpan ke Contact List Supplier.'
            : '';
          var ttpSync = window._lastTtpSyncResult;
          if (normToast === 'APPROVED' && ttpSync) {
            if (ttpSync.synced) {
              extraCls += ' Monitoring TTM/TTP: ' + String(ttpSync.inserted || 0) + ' baru, '
                + String(ttpSync.updated || 0) + ' diperbarui.';
            } else if (ttpSync.skipped && ttpSync.reason === 'no_ffb_rows') {
              extraCls += ' Monitoring TTM/TTP: tidak ada baris FFB di traceability.';
            }
          }
          window.showSddToast('Keputusan Submitted disimpan: ' + statusSdd + ' (SID: ' + sid + ').' + extraCls, 'success');
        }
      })
      .catch(function(e) {
        var msg = (e && e.message) ? e.message : String(e);
        var saveErr = document.getElementById('scr-save-ok');
        if (saveErr) {
          saveErr.style.display = 'block';
          saveErr.style.color = '#dc2626';
          saveErr.textContent = '✗ Decision save failed untuk Submission ID: ' + sid;
        }
        if (typeof window.showSddToast === 'function') {
          window.showSddToast('Gagal menyimpan keputusan Submitted: ' + msg, 'error');
        }
      });
    return;
  }
  if (typeof window.showSddToast === 'function') {
    window.showSddToast(
      'Keputusan: ' + statusSdd + '. Disimpan ke server saat Save Draft atau Submit.',
      'info'
    );
  }
};

if (!window.__sddApproverDecisionClickBound) {
  window.__sddApproverDecisionClickBound = true;
  document.addEventListener('click', function(ev) {
    var btn = ev.target && ev.target.closest ? ev.target.closest('button') : null;
    if (!btn || !btn.id) return;
    if (btn.id === 'sdd-approver-hold' || btn.id === 'boss-btn-hold') { ev.preventDefault(); window._submitSddApproverDecision('Hold'); }
    else if (btn.id === 'sdd-approver-approve' || btn.id === 'boss-btn-approve') { ev.preventDefault(); window._submitSddApproverDecision('Approve'); }
    else if (btn.id === 'sdd-approver-reject' || btn.id === 'boss-btn-reject') { ev.preventDefault(); window._submitSddApproverDecision('Reject'); }
  });
}

function formatSddDateImported() {
  const d = new Date();
  const p = function(n) { return String(n).padStart(2, '0'); };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

function setAliasGroup(obj, keys) {
  if (!obj || typeof obj !== 'object' || !Array.isArray(keys) || !keys.length) return;
  const firstNonEmpty = keys.map(function(k) { return obj[k]; }).find(function(v) {
    return v !== undefined && v !== null && String(v).trim() !== '';
  });
  if (firstNonEmpty === undefined) return;
  keys.forEach(function(k) {
    if (obj[k] === undefined || obj[k] === null || String(obj[k]).trim() === '') obj[k] = firstNonEmpty;
  });
}

/** Mirror important SDD fields across common header spelling variants. */
function applySddHeaderAliases(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  setAliasGroup(obj, ['Mill ID', 'Mil ID']);
  setAliasGroup(obj, ['Group Name', 'Grup Name']);

  // Legality aliases (slash vs spaced slash vs uppercase).
  setAliasGroup(obj, ['HGU/HGB Type', 'HGU / HGB Type', 'HGU HGB Type', 'HGU/HGB']);
  setAliasGroup(obj, ['HGU/HGB No', 'HGU / HGB No', 'HGU HGB No', 'HGU/HGB Number', 'HGU / HGB Number']);
  setAliasGroup(obj, ['HGU/HGB Issue Date', 'HGU / HGB Issue Date', 'HGU HGB Issue Date']);

  setAliasGroup(obj, ['IUP Type', 'IUP / IUP - P Type']);
  setAliasGroup(obj, ['IUP No', 'IUP / IUP - P No', 'IUP Number']);
  setAliasGroup(obj, ['IUP Issue Date', 'IUP / IUP - P Issue Date']);

  setAliasGroup(obj, ['Izin Lokasi Type', 'IZIN LOKASI Type', 'Izin Lokasi']);
  setAliasGroup(obj, ['Izin Lokasi No', 'IZIN LOKASI No', 'Izin Lokasi Number', 'IZIN LOKASI']);
  setAliasGroup(obj, ['Izin Lokasi Issue Date', 'IZIN LOKASI Issue Date']);

  setAliasGroup(obj, ['Izin Lingkungan Type', 'IZIN LINGKUNGAN Type', 'Izin Lingkungan']);
  setAliasGroup(obj, ['Izin Lingkungan No', 'IZIN LINGKUNGAN No', 'Izin Lingkungan Number', 'IZIN LINGKUNGAN']);
  setAliasGroup(obj, ['Izin Lingkungan Issue Date', 'IZIN LINGKUNGAN Issue Date']);

  setAliasGroup(obj, ['IMB/PBG Type', 'IMB / PBG Type', 'IMB PBG Type', 'IMB/PBG']);
  setAliasGroup(obj, ['IMB/PBG No', 'IMB / PBG No', 'IMB PBG No', 'IMB/PBG Number']);
  setAliasGroup(obj, ['IMB/PBG Issue Date', 'IMB / PBG Issue Date', 'IMB PBG Issue Date']);

  setAliasGroup(obj, ['NIB Type', 'NIB']);
  setAliasGroup(obj, ['NIB No', 'NIB Number']);
  setAliasGroup(obj, ['NIB Issue Date']);

  // Numeric / label km line from Excel → sheet must expose one of these headers or backend adds "Distance".
  setAliasGroup(obj, ['Distance', 'Distance (km)', 'Distance KM']);
  return obj;
}

/** Manual screening fields only (keys match SDD Data sheet headers). */
function buildScrDataPayload(scrData) {
  const payload = {
    'SCR - List Group By Owners': scrData.owners || '',
    'SCR - Previous News': scrData.news || '',
    'SCR - Supply To': scrData.supplyto || '',
    'SCR - Legality Status': scrData.legality || '',
    'SCR - Certification': scrData.cert || '',
    'SCR - NDPE Policy': scrData.ndpe || '',
    'SCR - No Buy List': scrData.nbl || '',
    'SCR - NBL Check Result': scrData.nblCheckResult || '',
    'SCR - NBL Match Detail': scrData.nblCheckDetail || '',
    'SCR - NBL Checked At': scrData.nblCheckedAt || '',
    'SCR - Grievance (Y/N)': scrData.grvYN || '',
    'SCR - PRI (Y/N)': scrData.priYN || '',
    'SCR - Notes': scrData.traceNote || '',
    'SCR - Requested Data': scrData.requestedData || '',
    'noteSDD': scrData.noteSdd || '',
    'noteBossDecision': scrData.noteBossDecision || scrData.noteSdd || '',
    'statusSDD': scrData.statusSdd || scrData.statusSDD || '',
    'statusBossDecision': scrData.statusBossDecision || scrData.statusSdd || scrData.statusSDD || '',
    'SCR - Screening Status': scrData.status || '',
    'SCR - Last Updated': scrData.date || ''
  };
  const grv = scrData.grvRows || [];
  const pri = scrData.priRows || [];
  for (let idx = 1; idx <= SCR_GRV_PRI_MAX_ROWS; idx++) {
    const row = grv[idx - 1];
    payload['GRV' + idx + ' - Source'] = row && row.source ? String(row.source) : '';
    payload['GRV' + idx + ' - Description'] = row && row.desc ? String(row.desc) : '';
    payload['GRV' + idx + ' - Publisher'] = row && row.pub ? String(row.pub) : '';
    payload['GRV' + idx + ' - Date Publish'] = row && row.date ? String(row.date) : '';
    payload['GRV' + idx + ' - Status'] = row && row.status ? String(row.status) : '';
    payload['GRV' + idx + ' - Attachment'] = row && row.attach ? String(row.attach) : '';
  }
  for (let idx = 1; idx <= SCR_GRV_PRI_MAX_ROWS; idx++) {
    const row = pri[idx - 1];
    payload['PRI' + idx + ' - Company'] = row && row.company ? String(row.company) : '';
    payload['PRI' + idx + ' - Description'] = row && row.desc ? String(row.desc) : '';
    payload['PRI' + idx + ' - Publisher'] = row && row.pub ? String(row.pub) : '';
    payload['PRI' + idx + ' - Date Publish'] = row && row.date ? String(row.date) : '';
    payload['PRI' + idx + ' - Attachment'] = row && row.attach ? String(row.attach) : '';
    payload['PRI' + idx + ' - Action Request'] = row && row.action ? String(row.action) : '';
  }
  return payload;
}

/**
 * handleFinalSave — rewritten to use the relational submission_id API.
 *
 * Flow:
 *   First save  (_sddSubmissionId == null) → createSubmission({ main, mills, ffb_rows })
 *   Subsequent  (_sddSubmissionId != null)  → updateSubmission({ submission_id, main, mills, ffb_rows })
 *
 * Null-safe: empty Excel cells stay '' and are never inherited from previous rows.
 * Partial updates only overwrite keys that are present in the payload.
 */
async function handleFinalSave(scrData) {
  // Recompute import cache so stale/empty cache edge cases are handled.
  cacheSddImportRowsFromWorkbook();

  // If a loaded draft is open alongside a new Excel workbook, prefer loaded identity fields.
  // PATCH: coordinate fields are excluded from this merge — the freshly-imported Excel
  // values (already dot-decimal via getValueNearLabelRaw_) are always more reliable than
  // whatever came back from Sheets (which may be a corrupted integer like 318318).
  if (window._loadedPrimarySddRow && supplierWorkbook && window._sddImportFirstRow) {
    const _COORD_FIELD_SET = new Set([
      'Latitude', 'Longitude',
      'TML - Latitude', 'TML - Longitude',
      'FFB - Latitude', 'FFB - Longitude',
    ]);
    // Save the good coordinate values from the import before the merge overwrites them
    const _coordBackup = {};
    _COORD_FIELD_SET.forEach(function(k) {
      if (window._sddImportFirstRow[k] !== undefined && window._sddImportFirstRow[k] !== null
          && String(window._sddImportFirstRow[k]).trim() !== '') {
        _coordBackup[k] = window._sddImportFirstRow[k];
      }
    });
    window._sddImportFirstRow = Object.assign({}, window._sddImportFirstRow, window._loadedPrimarySddRow);
    // Re-apply the good Excel coordinate values so they win over any corrupt sheet value
    Object.assign(window._sddImportFirstRow, _coordBackup);
  }
  // PATCH: always normalise coordinate fields in _sddImportFirstRow to dot-decimal
  // right before mergeSddImportIntoPayload reads it. Guards against any path that
  // set a comma-decimal value (e.g. browser locale rendering of a raw float).
  if (window._sddImportFirstRow) {
    ['Latitude', 'Longitude',
     'TML - Latitude', 'TML - Longitude',
     'FFB - Latitude', 'FFB - Longitude'].forEach(function(ck) {
      if (window._sddImportFirstRow[ck] !== undefined && window._sddImportFirstRow[ck] !== null) {
        window._sddImportFirstRow[ck] = normalizeCoordinate(window._sddImportFirstRow[ck], ck);
      }
    });
  }

  // ── Build main-form payload ──────────────────────────────────────────────
  const manual = buildScrDataPayload(scrData);
  const merged = mergeSddImportIntoPayload(manual);
  applySddHeaderAliases(merged);
  if (!merged['Supplier Type']) merged['Supplier Type'] = getCurrentSddSupplierType() || 'MILL';
  merged['supplier_type']             = getCurrentSddSupplierType() || 'MILL';
  merged['SCR - Screening Status']    = scrData.status || '';
  merged['SCR - Last Updated']        = scrData.date   || '';
  if (!merged['Date Imported']) merged['Date Imported'] = formatSddDateImported();
  if (!merged['Imported By'])   merged['Imported By']   = window._scrKey || 'UI';
  // Strip internal bookkeeping keys the backend must not receive
  delete merged['_row'];
  delete merged['submission_id'];

  // ── Build mill (TML) rows ────────────────────────────────────────────────
  const importedRows = Array.isArray(window._sddImportedRows) ? window._sddImportedRows : [];
  const dedupeSeen   = {};

  function buildMillRow(r) {
    const obj = {};
    Object.keys(r).forEach(function(k) {
      if (/^TML - |^SCR - TML/.test(k)) obj[k] = (r[k] === undefined || r[k] === null) ? '' : String(r[k]);
    });
    return obj;
  }
  function buildFfbRow(r) {
    const obj = {};
    Object.keys(r).forEach(function(k) {
      if (/^FFB - /.test(k)) obj[k] = (r[k] === undefined || r[k] === null) ? '' : String(r[k]);
    });
    return obj;
  }

  const millRows = [];
  const ffbRows  = [];

  importedRows.filter(hasMeaningfulImportData).forEach(function(r) {
    var sig = importRowSignature(r);
    if (!sig || dedupeSeen[sig]) return;
    dedupeSeen[sig] = true;

    if (!rowIsTraceabilityOnlyPayload(r)) return; // main-form row handled above

    const hasTml = String(r['TML - Mill Name'] || r['TML - Company Name'] || '').trim() !== '';
    const hasFfb = String(r['FFB - Supplier Name'] || '').trim() !== '';

    if (hasTml) {
      const mr = buildMillRow(r);
      // Carry line_id if we have one (from a previously loaded submission)
      if (r['line_id']) mr['line_id'] = String(r['line_id']);
      millRows.push(mr);
    }
    if (hasFfb) {
      const fr = buildFfbRow(r);
      if (r['line_id']) fr['line_id'] = String(r['line_id']);
      ffbRows.push(fr);
    }
  });

  // ── Also include TML/FFB rows loaded from getSubmissionById (already clean objects) ──
  const cachedGroup = window._scrSavedGroupsByKey && window._scrSavedGroupsByKey[window._sddSubmissionId];
  if (cachedGroup && !millRows.length && Array.isArray(cachedGroup.mills)) {
    cachedGroup.mills.forEach(function(m) {
      const mr = Object.assign({}, m);
      delete mr['submission_id']; delete mr['created_at']; delete mr['created_by']; delete mr['is_deleted'];
      millRows.push(mr);
    });
  }
  if (cachedGroup && !ffbRows.length && Array.isArray(cachedGroup.ffb_rows)) {
    cachedGroup.ffb_rows.forEach(function(f) {
      const fr = Object.assign({}, f);
      delete fr['submission_id']; delete fr['created_at']; delete fr['created_by']; delete fr['is_deleted'];
      ffbRows.push(fr);
    });
  }

  // ── Merge in-memory TML screening data → millRows ─────────────────────────
  // Untuk setiap mill yang di-screening via modal (tersimpan di window._tmlScreeningData),
  // cari row yang match di millRows berdasarkan TML - Mill Name (case-insensitive trim),
  // lalu merge field SCR - TML *. Kalau tidak ada match, push row baru.
  // Field name IDENTIK dengan saveTmlScreeningToSdd supaya skema backend tetap konsisten.
  if (window._tmlScreeningData && typeof window._tmlScreeningData === 'object') {
    Object.keys(window._tmlScreeningData).forEach(function(millName) {
      var scrD = window._tmlScreeningData[millName];
      if (!scrD) return;
      var tmlFields = {
        'SCR - TML Valid Coordinate'               : scrD.coord   || '',
        'SCR - TML Forest Area'                    : (scrD.forestItems || []).map(function(i) {
                                                       return i.label + ' (' + i.ha + ')';
                                                     }).join('; '),
        'SCR - TML Peatland'                       : (scrD.peatItems  || []).map(function(i) {
                                                       return i.label + ' (' + i.ha + ')';
                                                     }).join('; '),
        'SCR - TML Moratorium'                     : scrD.mora    || '',
        'SCR - TML Moratorium (Ha)'                : scrD.moraHa  || '',
        'SCR - TML Deforestation Buffer 50KM (Ha)' : scrD.defbuf  || '',
        'SCR - TML Screening Date'                 : scrD.date    || '',
      };
      var millNorm = normalizeCellText(millName).toLowerCase();
      var matchIdx = millRows.findIndex(function(mr) {
        return normalizeCellText(mr['TML - Mill Name'] || '').toLowerCase() === millNorm;
      });
      if (matchIdx >= 0) {
        Object.assign(millRows[matchIdx], tmlFields);
      } else {
        millRows.push(Object.assign({ 'TML - Mill Name': millName }, tmlFields));
      }
    });
  }

  // ── Merge in-memory FFB screening data → ffbRows ──────────────────────────
  // Field name IDENTIK dengan saveFfbScreeningToSdd.
  if (window._ffbScreeningData && typeof window._ffbScreeningData === 'object') {
    Object.keys(window._ffbScreeningData).forEach(function(supplierName) {
      var scrD = window._ffbScreeningData[supplierName];
      if (!scrD) return;
      var ffbFields = {
        'FFB - Valid Coordinate'      : scrD.coord       || '',
        'FFB - Forest Area'           : (scrD.forestItems || []).map(function(i) {
                                          return i.label + ' (' + i.ha + ')';
                                        }).join('; '),
        'FFB - Peatland'              : (scrD.peatItems  || []).map(function(i) {
                                          return i.label + ' (' + i.ha + ')';
                                        }).join('; '),
        'FFB - Moratorium'            : scrD.mora        || '',
        'FFB - Moratorium (Ha)'       : scrD.moraHa      || '',
        'FFB - Distance to Mill (Km)' : scrD.distKm      || '',
        'FFB - Deforestation (Ha)'    : scrD.defor       || '',
        'FFB - Burn Area (Ha)'        : scrD.burn        || '',
        'FFB - Village Risk'          : scrD.villageRisk || '',
      };
      var supplierNorm = normalizeCellText(supplierName).toLowerCase();
      var matchIdx = ffbRows.findIndex(function(fr) {
        return normalizeCellText(fr['FFB - Supplier Name'] || '').toLowerCase() === supplierNorm;
      });
      if (matchIdx >= 0) {
        Object.assign(ffbRows[matchIdx], ffbFields);
      } else {
        ffbRows.push(Object.assign({ 'FFB - Supplier Name': supplierName }, ffbFields));
      }
    });
  }

  // ── CREATE or UPDATE ─────────────────────────────────────────────────────
  const existingSid    = window._sddSubmissionId || null;
  const isLoadedSaved  = !!(window._sddIsLoadedSaved && existingSid);

  var finalResult;
  var syncedRows = 0;

  if (!isLoadedSaved) {
    // ── FIRST SAVE: createSubmission ──
    if (typeof window.showSddToast === 'function') window.showSddToast('Menyimpan submission baru…', 'info');
    finalResult = await apiCreateSubmission({
      main     : merged,
      mills    : millRows.length  ? millRows  : undefined,
      ffb_rows : ffbRows.length   ? ffbRows   : undefined,
    });
    // finalResult: { success, submission_id, mills_inserted, ffb_inserted }
    if (finalResult && finalResult.submission_id) {
      window._sddSubmissionId  = finalResult.submission_id;
      window._scrLoadedKey     = finalResult.submission_id;
      window._sddIsLoadedSaved = true;
      window._sddLastInsertedRow = null; // no longer needed
      window._scrLoadedRowNum    = null;
    }
    syncedRows = (finalResult && (Number(finalResult.mills_inserted || 0) + Number(finalResult.ffb_inserted || 0))) || 0;
  } else {
    // ── SUBSEQUENT SAVE: updateSubmission (partial — only present keys overwritten) ──
    if (typeof window.showSddToast === 'function') window.showSddToast('Memperbarui submission…', 'info');
    finalResult = await apiUpdateSubmission({
      submission_id : existingSid,
      main          : merged,
      mills         : millRows.length  ? millRows  : undefined,
      ffb_rows      : ffbRows.length   ? ffbRows   : undefined,
    });
    syncedRows = finalResult
      ? ((finalResult.mills && (Number(finalResult.mills.upserted || 0) + Number(finalResult.mills.inserted || 0))) || 0) +
        ((finalResult.ffb   && (Number(finalResult.ffb.upserted   || 0) + Number(finalResult.ffb.inserted   || 0))) || 0)
      : 0;
  }

  if (finalResult && typeof finalResult === 'object') {
    finalResult.importedSyncedRows = syncedRows;
    finalResult.importedSyncedMillIds = [];  // submission_id replaces millId tracking
  }
  return finalResult;
}

/**
 * saveTmlScreeningToSdd — write TML screening fields for one mill.
 *
 * SAFE FULL-SET UPDATE: The backend treats `mills` as a FULL REPLACEMENT SET
 * (upsertChildSheet_ soft-deletes any existing row not present in the payload).
 * This function ALWAYS sends the complete mills array — never a single-row slice —
 * so sibling rows are never accidentally deleted.
 *
 * Identity matching (scored — highest scorer wins, 0 = no match → append):
 *   score +100 : line_id exact match  ← screeningData.line_id (highest priority)
 *   score  +10 : TML - UML ID match   ← screeningData.tmlUmlId
 *   score   +5 : TML - Company Name   ← screeningData.tmlCompanyName
 *   score   +1 : TML - Mill Name      ← mill param (name-only fallback)
 * Ties broken by lowest array index (backend order = line_id order).
 * A row scoring 0 on ALL dimensions is never selected; a new row is appended.
 *
 * Callers can upgrade accuracy by adding optional identity fields to screeningData:
 *   { ..., line_id: '42', tmlUmlId: 'UML-001', tmlCompanyName: 'PT Maju' }
 * Old callers that pass only screening fields continue to work via the name fallback.
 *
 * Relational flow (submission_id present):
 *   1. Ensure _scrSavedGroupsByKey[sid].mills is populated; fetch via
 *      apiGetSubmissionById if the cache slot is missing or has no mills array.
 *   2. Score every cached row; pick highest scorer as target.
 *   3. Clone all rows; merge tmlFields ONLY into target; siblings untouched.
 *   4. Send apiUpdateSubmission({ submission_id: sid, mills: nextMills }).
 *   5. Re-fetch getSubmissionById to sync cache with backend-assigned line_ids.
 *
 * Legacy fallback (no submission_id): unchanged apiUpdateSDD / apiInsertSDD path.
 */
async function saveTmlScreeningToSdd(mill, screeningData) {
  const tmlFields = {
    'TML - Mill Name'                          : mill || '',
    'SCR - TML Valid Coordinate'               : screeningData.coord || '',
    'SCR - TML Forest Area'                    : (screeningData.forestItems || []).map(function(i) { return i.label + ' (' + i.ha + ')'; }).join('; '),
    'SCR - TML Peatland'                       : (screeningData.peatItems  || []).map(function(i) { return i.label + ' (' + i.ha + ')'; }).join('; '),
    'SCR - TML Moratorium'                     : screeningData.mora   || '',
    'SCR - TML Moratorium (Ha)'                : screeningData.moraHa || '',
    'SCR - TML Deforestation Buffer 50KM (Ha)' : screeningData.defbuf || '',
    'SCR - TML Screening Date'                 : screeningData.date   || '',
  };

  var sid = window._sddSubmissionId || window._scrLoadedKey || null;
  if (sid) {
    // ── Step 1: Ensure cache is hydrated with full mills array ────────────
    if (!window._scrSavedGroupsByKey) window._scrSavedGroupsByKey = {};
    var cachedGroup = window._scrSavedGroupsByKey[sid];
    if (!cachedGroup || !Array.isArray(cachedGroup.mills)) {
      var fetched = await apiGetSubmissionById(sid);
      function _tmlSafeStr(v) { return (v === undefined || v === null) ? '' : String(v); }
      function _tmlCleanRow(r) {
        if (!r || typeof r !== 'object') return {};
        var out = {}; Object.keys(r).forEach(function(k) { out[k] = _tmlSafeStr(r[k]); }); return out;
      }
      cachedGroup = window._scrSavedGroupsByKey[sid] = Object.assign(cachedGroup || {}, {
        mills    : (fetched.mills    || []).map(_tmlCleanRow),
        ffb_rows : (fetched.ffb_rows || []).map(_tmlCleanRow),
      });
    }

    // ── Step 2: Score every cached row; highest scorer = target ──────────
    //
    // Optional identity fields the caller may add to screeningData:
    //   screeningData.line_id        → TML - line_id exact match  (+100)
    //   screeningData.tmlUmlId       → TML - UML ID match         (+10)
    //   screeningData.tmlCompanyName → TML - Company Name match   (+5)
    //   mill (always present)        → TML - Mill Name match      (+1)
    //
    var existingMills  = Array.isArray(cachedGroup.mills) ? cachedGroup.mills : [];
    var millNameNorm   = normalizeCellText(mill || '').toLowerCase();
    // Normalise optional caller-supplied identity hints once
    var hintLineId     = screeningData.line_id       ? String(screeningData.line_id).trim()                                    : '';
    var hintUmlId      = screeningData.tmlUmlId      ? normalizeCellText(screeningData.tmlUmlId).toLowerCase()      : '';
    var hintCompany    = screeningData.tmlCompanyName ? normalizeCellText(screeningData.tmlCompanyName).toLowerCase() : '';

    var bestScore = 0;
    var targetIdx = -1;

    existingMills.forEach(function(m, i) {
      var score = 0;
      // line_id: authoritative — if caller supplies it and it matches, row is unambiguous
      if (hintLineId && String(m['line_id'] || '').trim() === hintLineId)                                                   score += 100;
      // TML - UML ID: strong secondary key (unique per mill)
      if (hintUmlId  && normalizeCellText(m['TML - UML ID']      || '').toLowerCase() === hintUmlId)                       score += 10;
      // TML - Company Name: composite refinement
      if (hintCompany && normalizeCellText(m['TML - Company Name'] || '').toLowerCase() === hintCompany)                    score += 5;
      // TML - Mill Name: baseline name match (required for fallback path)
      if (normalizeCellText(m['TML - Mill Name'] || '').toLowerCase() === millNameNorm)                                     score += 1;

      if (score > bestScore) { bestScore = score; targetIdx = i; }
      // Ties: first occurrence wins (lower line_id from backend ordering)
    });

    if (bestScore === 0) {
      // No cached row matches at all — this is a new mill being screened.
      console.warn('[saveTmlScreeningToSdd] No matching row for mill:', mill,
        '— appending new row. Provide screeningData.line_id for exact targeting.');
    }

    // ── Step 3: Build nextMills — full set, only target patched ───────────
    var nextMills;
    if (targetIdx >= 0) {
      nextMills = existingMills.map(function(m, i) {
        if (i !== targetIdx) return Object.assign({}, m);   // sibling — untouched
        return Object.assign({}, m, tmlFields);             // target  — SCR fields updated, line_id & others preserved
      });
    } else {
      // Append new row; backend assigns line_id
      nextMills = existingMills.map(function(m) { return Object.assign({}, m); });
      nextMills.push(Object.assign({}, tmlFields));
    }

    // ── Step 4: Send FULL replacement set ─────────────────────────────────
    var result = await apiUpdateSubmission({ submission_id: sid, mills: nextMills });

    // ── Step 5: Sync cache — re-fetch for backend-assigned line_ids ────────
    if (result && result.success !== false) {
      try {
        var refreshed = await apiGetSubmissionById(sid);
        function _tmlSafeStr2(v) { return (v === undefined || v === null) ? '' : String(v); }
        function _tmlCleanRow2(r) {
          if (!r || typeof r !== 'object') return {};
          var out = {}; Object.keys(r).forEach(function(k) { out[k] = _tmlSafeStr2(r[k]); }); return out;
        }
        window._scrSavedGroupsByKey[sid].mills    = (refreshed.mills    || []).map(_tmlCleanRow2);
        window._scrSavedGroupsByKey[sid].ffb_rows = (refreshed.ffb_rows || []).map(_tmlCleanRow2);
      } catch (cacheErr) {
        // Network blip: fall back to optimistic cache with what we sent
        window._scrSavedGroupsByKey[sid].mills = nextMills;
      }
    }

    return result;
  }

  // ── Legacy fallback (no submission_id) — unchanged ────────────────────────
  const payload = Object.assign({
    'Supplier Type'          : getCurrentSddSupplierType() || 'MILL',
    'SCR - Screening Status' : screeningData.status || '',
    'SCR - Last Updated'     : screeningData.date   || '',
  }, tmlFields);
  const merged  = mergeSddImportIntoPayload(payload);
  const rowNum  = Number(window._scrLoadedRowNum || window._sddLastInsertedRow || 0);
  if (window._sddIsLoadedSaved && rowNum >= 2) return apiUpdateSDD(rowNum, merged);
  const inserted = await apiInsertSDD(merged);
  if (inserted && inserted.row) {
    window._sddLastInsertedRow = Number(inserted.row) || null;
    window._scrLoadedRowNum    = window._sddLastInsertedRow;
    window._sddIsLoadedSaved   = true;
  }
  return inserted;
}

/**
 * saveFfbScreeningToSdd — write FFB screening fields for one supplier.
 *
 * SAFE FULL-SET UPDATE: Mirrors saveTmlScreeningToSdd — always sends the
 * complete ffb_rows array so sibling rows are never soft-deleted by the backend.
 *
 * Identity matching (scored — highest scorer wins, 0 = no match → append):
 *   score +100 : line_id exact match           ← screeningData.line_id
 *   score  +10 : FFB - Mill Name composite     ← screeningData.ffbMillName
 *   score   +5 : FFB - Village composite       ← screeningData.ffbVillage
 *   score   +3 : FFB - Sub District composite  ← screeningData.ffbSubDistrict
 *   score   +1 : FFB - Supplier Name (name-only fallback, always scored)
 * Ties broken by lowest array index (backend line_id ordering).
 *
 * This mirrors the backend's own FFB identity hierarchy (FINAL.gs ~565):
 *   composite = Supplier Name + Mill Name, fallback = Supplier Name alone.
 *
 * Callers upgrade accuracy by adding optional identity fields to screeningData:
 *   { ..., line_id: '7', ffbMillName: 'PKS Sinar', ffbVillage: 'Desa Maju' }
 * Old callers that pass only screening fields continue to work via name fallback.
 *
 * Relational flow (submission_id present):
 *   1. Ensure _scrSavedGroupsByKey[sid].ffb_rows is populated; fetch via
 *      apiGetSubmissionById if the cache slot is missing.
 *   2. Score every cached row; pick highest scorer as target.
 *   3. Clone all rows; merge ffbFields ONLY into target; siblings untouched.
 *   4. Send apiUpdateSubmission({ submission_id: sid, ffb_rows: nextFfbRows }).
 *   5. Re-fetch getSubmissionById to sync cache with backend-assigned line_ids.
 *
 * Legacy fallback (no submission_id): unchanged row-number UPDATE/INSERT path.
 */
async function saveFfbScreeningToSdd(supplier, screeningData) {
  const ffbFields = {
    'FFB - Supplier Name'        : supplier || '',
    'FFB - Valid Coordinate'     : screeningData.coord       || '',
    'FFB - Forest Area'          : (screeningData.forestItems || []).map(function(i) { return i.label + ' (' + i.ha + ')'; }).join('; '),
    'FFB - Peatland'             : (screeningData.peatItems  || []).map(function(i) { return i.label + ' (' + i.ha + ')'; }).join('; '),
    'FFB - Moratorium'           : screeningData.mora        || '',
    'FFB - Moratorium (Ha)'      : screeningData.moraHa      || '',
    'FFB - Distance to Mill (Km)': screeningData.distKm      || '',
    'FFB - Deforestation (Ha)'   : screeningData.defor       || '',
    'FFB - Burn Area (Ha)'       : screeningData.burn        || '',
    'FFB - Village Risk'         : screeningData.villageRisk || '',
  };

  var sid = window._sddSubmissionId || window._scrLoadedKey || null;
  if (sid) {
    // ── Step 1: Ensure cache is hydrated with full ffb_rows array ─────────
    if (!window._scrSavedGroupsByKey) window._scrSavedGroupsByKey = {};
    var cachedGroup = window._scrSavedGroupsByKey[sid];
    if (!cachedGroup || !Array.isArray(cachedGroup.ffb_rows)) {
      var fetched = await apiGetSubmissionById(sid);
      function _ffbSafeStr(v) { return (v === undefined || v === null) ? '' : String(v); }
      function _ffbCleanRow(r) {
        if (!r || typeof r !== 'object') return {};
        var out = {}; Object.keys(r).forEach(function(k) { out[k] = _ffbSafeStr(r[k]); }); return out;
      }
      cachedGroup = window._scrSavedGroupsByKey[sid] = Object.assign(cachedGroup || {}, {
        mills    : (fetched.mills    || []).map(_ffbCleanRow),
        ffb_rows : (fetched.ffb_rows || []).map(_ffbCleanRow),
      });
    }

    // ── Step 2: Score every cached row; highest scorer = target ──────────
    //
    // Optional identity fields the caller may add to screeningData:
    //   screeningData.line_id        → FFB line_id exact match         (+100)
    //   screeningData.ffbMillName    → FFB - Mill Name composite       (+10)
    //   screeningData.ffbVillage     → FFB - Village composite         (+5)
    //   screeningData.ffbSubDistrict → FFB - Sub District composite    (+3)
    //   supplier (always present)    → FFB - Supplier Name fallback    (+1)
    //
    // Mirrors the backend composite key: Supplier Name + Mill Name (FINAL.gs ~565).
    //
    var existingFfb    = Array.isArray(cachedGroup.ffb_rows) ? cachedGroup.ffb_rows : [];
    var supplierNorm   = normalizeCellText(supplier || '').toLowerCase();
    var hintLineId     = screeningData.line_id        ? String(screeningData.line_id).trim()                                       : '';
    var hintMillName   = screeningData.ffbMillName    ? normalizeCellText(screeningData.ffbMillName).toLowerCase()    : '';
    var hintVillage    = screeningData.ffbVillage     ? normalizeCellText(screeningData.ffbVillage).toLowerCase()     : '';
    var hintSubDist    = screeningData.ffbSubDistrict ? normalizeCellText(screeningData.ffbSubDistrict).toLowerCase() : '';

    var bestScore = 0;
    var targetIdx = -1;

    existingFfb.forEach(function(f, i) {
      var score = 0;
      if (hintLineId   && String(f['line_id'] || '').trim() === hintLineId)                                                        score += 100;
      if (hintMillName && normalizeCellText(f['FFB - Mill Name']    || '').toLowerCase() === hintMillName)                         score += 10;
      if (hintVillage  && normalizeCellText(f['FFB - Village']      || '').toLowerCase() === hintVillage)                          score += 5;
      if (hintSubDist  && normalizeCellText(f['FFB - Sub District'] || '').toLowerCase() === hintSubDist)                          score += 3;
      if (normalizeCellText(f['FFB - Supplier Name'] || '').toLowerCase() === supplierNorm)                                        score += 1;
      if (score > bestScore) { bestScore = score; targetIdx = i; }
    });

    if (bestScore === 0) {
      console.warn('[saveFfbScreeningToSdd] No matching row for supplier:', supplier,
        '— appending new row. Provide screeningData.line_id for exact targeting.');
    }

    // ── Step 3: Build nextFfbRows — full set, only target patched ─────────
    var nextFfbRows;
    if (targetIdx >= 0) {
      nextFfbRows = existingFfb.map(function(f, i) {
        if (i !== targetIdx) return Object.assign({}, f);   // sibling — untouched
        return Object.assign({}, f, ffbFields);             // target  — SCR fields updated, line_id & others preserved
      });
    } else {
      nextFfbRows = existingFfb.map(function(f) { return Object.assign({}, f); });
      nextFfbRows.push(Object.assign({}, ffbFields));
    }

    // ── Step 4: Send FULL replacement set ─────────────────────────────────
    var result = await apiUpdateSubmission({ submission_id: sid, ffb_rows: nextFfbRows });

    // ── Step 5: Sync cache — re-fetch for backend-assigned line_ids ────────
    if (result && result.success !== false) {
      try {
        var refreshed = await apiGetSubmissionById(sid);
        function _ffbSafeStr2(v) { return (v === undefined || v === null) ? '' : String(v); }
        function _ffbCleanRow2(r) {
          if (!r || typeof r !== 'object') return {};
          var out = {}; Object.keys(r).forEach(function(k) { out[k] = _ffbSafeStr2(r[k]); }); return out;
        }
        window._scrSavedGroupsByKey[sid].mills    = (refreshed.mills    || []).map(_ffbCleanRow2);
        window._scrSavedGroupsByKey[sid].ffb_rows = (refreshed.ffb_rows || []).map(_ffbCleanRow2);
      } catch (cacheErr) {
        window._scrSavedGroupsByKey[sid].ffb_rows = nextFfbRows;
      }
    }

    return result;
  }

  // ── Legacy fallback (no submission_id) — unchanged ────────────────────────
  const payload = Object.assign({
    'Supplier Type'          : getCurrentSddSupplierType() || 'MILL',
    'SCR - Screening Status' : screeningData.status || '',
    'SCR - Last Updated'     : screeningData.date   || '',
  }, ffbFields);
  const merged = mergeSddImportIntoPayload(payload);

  // Look up legacy row number for this specific FFB supplier from the flat cache
  let ffbRowNum = null;
  const cacheRows  = Array.isArray(window._sddAllRowsCache) ? window._sddAllRowsCache : [];
  const loadedKey  = window._scrLoadedKey || '';
  const primaryRow = window._loadedPrimarySddRow || {};
  if (loadedKey && supplier) {
    const supplierNorm = normalizeCellText(supplier).toLowerCase();
    const allRelated   = rowsMatchingSavedSubmissionKey(cacheRows, loadedKey, primaryRow);
    const matchedFfbRow = allRelated.find(function(r) {
      return normalizeCellText(r['FFB - Supplier Name'] || '').toLowerCase() === supplierNorm &&
             r._row != null && Number(r._row) >= 2;
    });
    if (matchedFfbRow) ffbRowNum = Number(matchedFfbRow._row);
  }

  if (ffbRowNum && ffbRowNum >= 2) return apiUpdateSDD(ffbRowNum, merged);
  const inserted = await apiInsertSDD(merged);
  if (inserted && inserted.row) {
    window._sddLastInsertedRow = Number(inserted.row) || null;
    if (!window._scrLoadedRowNum) window._scrLoadedRowNum = window._sddLastInsertedRow;
    window._sddIsLoadedSaved = true;
  }
  return inserted;
}

function initDashboardApp() {
  try {

  (function ensureLoginUiMounted() {
    let root = document.getElementById('login');
    if (!root) {
      root = document.createElement('div');
      root.id = 'login';
      root.className = 'page active';
      document.body.prepend(root);
    }
    if (!document.getElementById('btn-login-submit')) mountLoginPage(root);
  })();

  console.log('🚀 Sustainability Dashboard loaded and connected to Google Sheets backend');
  console.log('ℹ️ SDD Apps Script URL:', typeof getSddApiUrl === 'function' ? getSddApiUrl() : '(n/a)');
  console.log('ℹ️ To use your own deployment: localStorage.setItem("SDD_WEBAPP_URL", "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec")');
  if (location.protocol === 'file:') {
    console.warn('⚠️ Page is file:// — saving to Google Sheets will not work until you open this HTML via http(s) (e.g. Live Server).');
  }

  const supplierExcelFileEl = document.getElementById('supplierExcelFile');
  const supplierCard = supplierExcelFileEl && supplierExcelFileEl.closest('.panel-card');
  if (supplierCard && !document.getElementById('sdd-supplier-type-wrap')) {
    const typeWrap = document.createElement('div');
    typeWrap.id = 'sdd-supplier-type-wrap';
    typeWrap.className = 'field';
    typeWrap.style.marginBottom = '12px';
    typeWrap.innerHTML =
      '<label for="sdd-supplier-type-select">Choose Supplier Type</label>'
      + '<select id="sdd-supplier-type-select">'
      + '<option value="">-- Select type first --</option>'
      + '<option value="MILL">MILL</option>'
      + '<option value="KCP">KCP</option>'
      + '<option value="TRADER">TRADER</option>'
      + '</select>';
    supplierCard.insertBefore(typeWrap, supplierCard.firstChild);
  }
  const supplierTypeSel = document.getElementById('sdd-supplier-type-select');
  if (supplierTypeSel) {
    supplierTypeSel.addEventListener('change', function() {
      window._sddSupplierType = normalizeSddSupplierType(this.value);
      syncSddSupplierTypeSelectorUI();
      if (typeof window.refreshSavedScreeningListGlobal === 'function') {
        window.refreshSavedScreeningListGlobal(window._scrLoadedKey || '');
      }
    });
  }
  if (supplierExcelFileEl) {
    supplierExcelFileEl.addEventListener('click', function() {
      if (!getCurrentSddSupplierType()) {
        if (typeof window.showSddToast === 'function') {
          window.showSddToast('Pilih supplier type dulu sebelum pilih file Excel.', 'error');
        }
        return;
      }
      this.value = '';
    });
  }
  syncSddSupplierTypeSelectorUI();

  // Saved screening: list + refresh (load is click-on-list).
  const topRefreshBtn = document.getElementById('scr-refresh-list');
  if (topRefreshBtn) topRefreshBtn.addEventListener('click', function() {
    if (typeof window.refreshSavedScreeningListGlobal === 'function') window.refreshSavedScreeningListGlobal();
  });
  const topExitBtn = document.getElementById('scr-exit-load');
  if (topExitBtn) topExitBtn.addEventListener('click', function() {
    if (typeof window.exitSavedScrModeGlobal === 'function') window.exitSavedScrModeGlobal(false);
  });
  if (typeof window.refreshSavedScreeningListGlobal === 'function') {
    window.refreshSavedScreeningListGlobal();
  }

  const MILL_FIELDS = ['QUARTER','YEAR','COMPANY CODE','TRADER NAME','GROUP NAME','COMPANY NAME','MILL NAME','UML ID','ADDRESS','PROVINCE','COORDINATES','MILL CATEGORY','MILL CAPACITY (TON/HOUR)','HGU/HGB','IZIN LOKASI','IUP','IZIN LINGKUNGAN','SCORE','MILL LOC','COMPLIMENT/NOT COMPLIMENT','DEFORESTATION SPATIAL','BURN AREA SPATIAL','PEAT','LEGALITY','DEFORESTATION GRIEVANCES','BURN AREA GRIEVANCES','HUMAN RIGHT','SAFETY','SOCIAL','ENVIRONMENT','TOTAL GRIEVANCES','NDPE','HRDD','TOTAL POLICY','CERTIFICATION','TOTAL CERTIFICATION','TOTAL SCORE','SUPPLIER LEVEL','BUYER NO BUY LIST','VOLUME SUPPLY STATUS','RECOMMENDATION LEVEL','SIGN','SUPPLIER STATUS','RISK LEVEL','RESULT RISK LEVEL','FACILITY NAME CPO','FACILITY NAME PK','PRODUCT SUPPLY'];
  let modalSheet = '', modalMode = '', modalRow = null, modalFields = [];
  let modalTaskKey = ''; // submission_id dari SDD yang di-add via Task List
  let allData = [];
  let currentFilter = 'All';
  let currentSearch = '';
  let ttpData = [], ttpFields = [], ttpLoaded = false, ttpPctCol = '', ttpPkPctCol = '', ttpCategoryCol = '', ttpSearch = '';
  let ttpPkTraceVolCol = '', ttpCpoTraceVolCol = '', ttpPkTraceDenomCol = '', ttpCpoTraceDenomCol = '';
  let ttpPeriodMode = 'overall'; // 'overall' (full year) | 'quarter'
  let ttpPeriodYear = '';
  let ttpPeriodQuarter = 'Q1';
  let millLoadPromise = null;
  let ttpLoadPromise = null;
  let grvLoadPromise = null;

  function prepareTtpRowPerfCache(row, fields) {
    if (!row || typeof row !== 'object') return row;
    row._sddSearchBlob = (fields || []).map(function(h) {
      return String(row[h] || '').toLowerCase();
    }).join('|');
    return row;
  }

  const scheduleRenderTTPTable = makeRafScheduler(function() {
    renderTTPTable();
  });

  // Field definitions for mill form
  const YESNO_FIELDS = ['HGU/HGB','IZIN LOKASI','IUP','IZIN LINGKUNGAN','LEGALITY','DEFORESTATION SPATIAL','BURN AREA SPATIAL','PEAT','DEFORESTATION GRIEVANCES','BURN AREA GRIEVANCES','HUMAN RIGHT','SAFETY','SOCIAL','ENVIRONMENT','NDPE','HRDD'];
  const SCORE_SOURCES = ['HGU/HGB','IZIN LOKASI','IUP','IZIN LINGKUNGAN'];           // → SCORE
  const GRIEVANCE_SOURCES = ['DEFORESTATION GRIEVANCES','BURN AREA GRIEVANCES','HUMAN RIGHT','SAFETY','SOCIAL','ENVIRONMENT']; // → TOTAL GRIEVANCES
  const POLICY_SOURCES = ['NDPE','HRDD'];                                             // → TOTAL POLICY
  const AUTO_FIELDS = ['SCORE','TOTAL GRIEVANCES','TOTAL POLICY','TOTAL SCORE'];      // auto-calculated

  const DROPDOWN_FIELDS = {
    'SUPPLIER STATUS': ['Active','Inactive','Conditional'],
    'SUPPLIER LEVEL': ['Tier 1','Tier 2','Tier 3'],
    'RISK LEVEL': ['High','Medium','Low'],
    'RESULT RISK LEVEL': ['High','Medium','Low'],
    'RECOMMENDATION LEVEL': ['Approved','Conditional','Not Approved'],
    'VOLUME SUPPLY STATUS': ['Active','Inactive'],
    'COMPLIMENT/NOT COMPLIMENT': ['Compliment','Not Compliment'],
    'MILL LOC': ['Inside Concession','Outside Concession'],
    'BUYER NO BUY LIST': ['Yes','No'],
  };

  const FIELD_SECTIONS = [
    { title: 'Identitas Mill', fields: ['QUARTER','YEAR','COMPANY CODE','TRADER NAME','GROUP NAME','COMPANY NAME','MILL NAME','UML ID','PROVINCE','ADDRESS','COORDINATES','MILL CATEGORY','MILL CAPACITY (TON/HOUR)'] },
    { title: 'Legalitas', fields: ['HGU/HGB','IZIN LOKASI','IUP','IZIN LINGKUNGAN','LEGALITY'], totalField: 'SCORE', totalLabel: 'Score (auto)' },
    { title: 'Spatial & Peat', fields: ['DEFORESTATION SPATIAL','BURN AREA SPATIAL','PEAT','MILL LOC','COMPLIMENT/NOT COMPLIMENT'] },
    { title: 'Grievances', fields: ['DEFORESTATION GRIEVANCES','BURN AREA GRIEVANCES','HUMAN RIGHT','SAFETY','SOCIAL','ENVIRONMENT'], totalField: 'TOTAL GRIEVANCES', totalLabel: 'Total Grievances (auto)' },
    { title: 'Policy', fields: ['NDPE','HRDD'], totalField: 'TOTAL POLICY', totalLabel: 'Total Policy (auto)' },
    { title: 'Sertifikasi', fields: ['CERTIFICATION','TOTAL CERTIFICATION'] },
    { title: 'Supply & Status', fields: ['SUPPLIER LEVEL','SUPPLIER STATUS','BUYER NO BUY LIST','VOLUME SUPPLY STATUS','RECOMMENDATION LEVEL','SIGN','RISK LEVEL','RESULT RISK LEVEL','FACILITY NAME CPO','FACILITY NAME PK','PRODUCT SUPPLY'] },
  ];

  function calcTotals() {
    const getVal = f => {
      const el = document.querySelector(`[data-field="${f}"]`);
      return el ? el.value : '';
    };
    const countYes = arr => arr.filter(f => getVal(f) === 'Yes').length;

    const score = countYes(SCORE_SOURCES);
    const grv = countYes(GRIEVANCE_SOURCES);
    const pol = countYes(POLICY_SOURCES);
    const total = score + grv + pol;

    const setAuto = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
    setAuto('auto-SCORE', score);
    setAuto('auto-TOTAL GRIEVANCES', grv);
    setAuto('auto-TOTAL POLICY', pol);
    setAuto('auto-TOTAL SCORE', total);

    // store values as hidden inputs
    ['SCORE','TOTAL GRIEVANCES','TOTAL POLICY','TOTAL SCORE'].forEach((f, i) => {
      const vals = [score, grv, pol, total];
      let hidden = document.querySelector(`input[data-field="${f}"]`);
      if (!hidden) {
        hidden = document.createElement('input');
        hidden.type = 'hidden';
        hidden.dataset.field = f;
        document.getElementById('modalFormGrid').appendChild(hidden);
      }
      hidden.value = vals[i];
    });
  }

  // ─── CUSTOM DROPDOWN BUILDER ────────────────────────────
  function buildCustomSelect(field, options, currentVal, isYesNo, isFull) {
    const wrapClass = ['custom-select-wrap', isYesNo ? 'yesno-wrap' : '', isFull ? '' : ''].join(' ').trim();
    const arrowSvg = `<svg class="cs-arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><polyline points="6 9 12 15 18 9"/></svg>`;
    const checkSvg = `<svg class="cs-opt-check" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><polyline points="20 6 9 17 4 12"/></svg>`;

    let triggerInner, optionsHtml;
    if (isYesNo) {
      const dotHtml = currentVal ? `<span class="cs-yn-dot"></span>` : '';
      triggerInner = currentVal
        ? `${dotHtml}<span class="cs-value">${currentVal}</span>`
        : `<span class="cs-placeholder">— Pilih —</span>`;
      optionsHtml = `
        <div class="cs-option cs-option-placeholder" data-val="">— Pilih —</div>
        <div class="cs-option cs-yes${currentVal==='Yes'?' selected':''}" data-val="Yes"><span class="cs-yn-dot"></span>${checkSvg}Yes</div>
        <div class="cs-option cs-no${currentVal==='No'?' selected':''}" data-val="No"><span class="cs-yn-dot"></span>${checkSvg}No</div>`;
    } else {
      triggerInner = currentVal
        ? `<span class="cs-value">${currentVal}</span>`
        : `<span class="cs-placeholder">— Pilih —</span>`;
      optionsHtml = `<div class="cs-option cs-option-placeholder" data-val="">— Pilih —</div>`
        + options.map(o => `<div class="cs-option${currentVal===o?' selected':''}" data-val="${o}">${checkSvg}${o}</div>`).join('');
    }

    return `<div class="form-field${isFull ? ' full' : ''}">
      <label>${field}</label>
      <input type="hidden" data-field="${field}" value="${currentVal}">
      <div class="${wrapClass}${currentVal==='Yes'?' val-yes':currentVal==='No'?' val-no':''}">
        <div class="custom-select-trigger">${triggerInner}${arrowSvg}</div>
        <div class="custom-select-dropdown">${optionsHtml}</div>
      </div>
    </div>`;
  }

  function initCustomSelects(container) {
    container.querySelectorAll('.custom-select-wrap').forEach(wrap => {
      const trigger = wrap.querySelector('.custom-select-trigger');
      const dropdown = wrap.querySelector('.custom-select-dropdown');
      const hidden = wrap.parentElement.querySelector('input[data-field]');
      const isYesNo = wrap.classList.contains('yesno-wrap');

      trigger.addEventListener('click', function(e) {
        e.stopPropagation();
        const isOpen = wrap.classList.contains('open');
        // close all others
        container.querySelectorAll('.custom-select-wrap.open').forEach(w => w.classList.remove('open'));
        if (!isOpen) wrap.classList.add('open');
      });

      dropdown.querySelectorAll('.cs-option').forEach(opt => {
        opt.addEventListener('click', function(e) {
          e.stopPropagation();
          const val = this.dataset.val;
          hidden.value = val;

          // update trigger display
          const arrowSvg = trigger.querySelector('.cs-arrow').outerHTML;
          if (!val) {
            trigger.innerHTML = `<span class="cs-placeholder">— Pilih —</span>${arrowSvg}`;
          } else if (isYesNo) {
            trigger.innerHTML = `<span class="cs-yn-dot"></span><span class="cs-value">${val}</span>${arrowSvg}`;
          } else {
            trigger.innerHTML = `<span class="cs-value">${val}</span>${arrowSvg}`;
          }

          // update selected class
          dropdown.querySelectorAll('.cs-option').forEach(o => o.classList.remove('selected'));
          if (val) this.classList.add('selected');

          // yesno coloring on wrap
          if (isYesNo) {
            wrap.classList.remove('val-yes','val-no');
            if (val === 'Yes') wrap.classList.add('val-yes');
            else if (val === 'No') wrap.classList.add('val-no');
          }

          wrap.classList.remove('open');
          calcTotals();
        });
      });
    });

    // Bind once globally to avoid stacking document listeners each modal open.
    if (!window._sddCustomSelectOutsideBound) {
      window._sddCustomSelectOutsideBound = true;
      document.addEventListener('click', function() {
        document.querySelectorAll('.custom-select-wrap.open').forEach(function(wrap) {
          wrap.classList.remove('open');
        });
      });
    }
  }

  function buildMillForm(data) {
    const grid = document.getElementById('modalFormGrid');
    grid.className = 'modal-form-grid cols-1';
    let html = '';
    FIELD_SECTIONS.forEach(sec => {
      html += `<div class="mill-form-section"><div class="mill-form-section-title">${sec.title}</div><div class="mill-form-grid">`;
      sec.fields.forEach(f => {
        const val = data ? (data[f] || '') : '';
        const isFull = f === 'ADDRESS' || f === 'COORDINATES';
        if (YESNO_FIELDS.includes(f)) {
          html += buildCustomSelect(f, ['Yes','No'], val, true, isFull);
        } else if (DROPDOWN_FIELDS[f]) {
          html += buildCustomSelect(f, DROPDOWN_FIELDS[f], val, false, isFull);
        } else {
          html += `<div class="form-field${isFull ? ' full' : ''}">
            <label>${f}</label>
            <input type="text" data-field="${f}" value="${val}" placeholder="${f}"></div>`;
        }
      });
      if (sec.totalField) {
        html += `<div class="form-field auto-total-field full">
          <label>${sec.totalLabel}</label>
          <div class="auto-total-val"><span id="auto-${sec.totalField}">0</span></div>
        </div>`;
      }
      html += `</div></div>`;
    });
    html += `<div class="mill-form-section"><div class="mill-form-grid">
      <div class="form-field auto-total-field full">
        <label>Total Score (auto)</label>
        <div class="auto-total-val grand"><span id="auto-TOTAL SCORE">0</span></div>
      </div></div></div>`;
    grid.innerHTML = html;

    initCustomSelects(grid);
    calcTotals();
  }

  function openModal(sheet, fields, mode, data) {
    modalSheet = sheet;
    modalMode = mode;
    modalRow = data ? data._row : null;
    modalFields = fields;

    const modalTitleEl = document.getElementById('modalTitle');
    if (modalTitleEl) modalTitleEl.textContent = mode === 'add' ? 'Add New Record' : 'Edit Record';
    const grid = document.getElementById('modalFormGrid');
    if (!grid) return;

    if (sheet === 'mill') {
      buildMillForm(data);
    } else {
      grid.className = 'modal-form-grid';
      grid.innerHTML = fields.map(f => `
        <div class="form-field ${f === 'ADDRESS' ? 'full' : ''}">
          <label>${f}</label>
          <input type="text" data-field="${f}" value="${data ? (data[f] || '') : ''}" placeholder="${f}">
        </div>`).join('');
    }

    document.getElementById('modalOverlay')?.classList.add('active');
    const firstInput = grid.querySelector('input, select');
    if (firstInput) firstInput.focus();
  }

  function closeModal() {
    document.getElementById('modalOverlay')?.classList.remove('active');
  }

  (function bindCoreModals() {
    const modalClose = document.getElementById('modalClose');
    const modalCancel = document.getElementById('modalCancel');
    const modalOverlay = document.getElementById('modalOverlay');
    const modalSave = document.getElementById('modalSave');
    if (!modalClose || !modalCancel || !modalOverlay || !modalSave) {
      console.error('[dashboard] Modal nodes missing (modalClose / modalCancel / modalOverlay / modalSave). Pastikan `npm run dev` dari folder sustain-dashboard (bukan buka file HTML mentah).');
      return;
    }
    modalClose.addEventListener('click', closeModal);
    modalCancel.addEventListener('click', closeModal);
    modalOverlay.addEventListener('click', function(e) {
      if (e.target === this) closeModal();
    });

    modalSave.addEventListener('click', async function() {
    const btn = this;
    btn.disabled = true;
    btn.textContent = 'Saving...';
    const data = {};
    document.querySelectorAll('#modalFormGrid [data-field]').forEach(el => {
      data[el.dataset.field] = el.value;
    });
    try {
      if (modalMode === 'add') {
        await apiPost({ action: 'add', sheet: modalSheet, data });
      } else {
        await apiPost({ action: 'update', sheet: modalSheet, row: modalRow, data });
      }
      closeModal();
      if (modalSheet === 'mill') {
        if (modalTaskKey) {
          const taskKey = modalTaskKey;
          modalTaskKey = '';
          try {
            await apiUpdateSubmission({ submission_id: taskKey, main: { mill_added: 'true' } });
            // Update local cache so re-render immediately reflects change
            if (window._scrSavedRowsByKey && window._scrSavedRowsByKey[taskKey]) {
              window._scrSavedRowsByKey[taskKey]['mill_added'] = 'true';
            }
          } catch(e) {
            console.warn('[TaskList] Failed to mark mill_added on submission:', e);
          }
          if (currentFilter === 'Task List') renderMillTaskList();
        }
        // Mark supply draft row as submitted if opened from supply import context
        if (window._supplyModalContext) {
          const ctx = window._supplyModalContext;
          window._supplyModalContext = null;
          try {
            const ctxBatch = (window._supplyDraftBatches || []).find(function(b) { return b.batch_id === ctx.batchId; });
            if (ctxBatch && ctxBatch.rows[ctx.rowIdx]) {
              ctxBatch.rows[ctx.rowIdx]._submitted = true;
              const allDone = ctxBatch.rows.every(function(r) { return r._submitted; });
              if (allDone) ctxBatch.status = 'submitted';
              // Mark on server
              apiPost({ action: 'submitSupplyDraft', batch_id: ctx.batchId, rows: [ctxBatch.rows[ctx.rowIdx]] })
                .catch(function(err) { console.warn('[supplyDraft] Mark submitted failed:', err.message); });
              if (typeof renderSupplyDraftList_ === 'function') renderSupplyDraftList_();
            }
          } catch(e) {
            console.warn('[supplyDraft] Failed to mark row submitted after modal save:', e);
          }
        }
        await loadMillData();
      } else if (modalSheet === 'ttp') {
        ttpLoaded = false; await loadTTPData();
      } else if (modalSheet === 'contactSupplier') {
        contactListLoaded = false; await loadContactListData(true);
      } else if (modalSheet === 'grievance') {
        grvLoaded = false; await loadGrvData();
      }
    } catch(err) {
      alert('Error saving: ' + err.message);
    }
    btn.disabled = false;
    btn.textContent = 'Save';
  });
  })();

  // ─── CONFIRM DELETE ─────────────────────────────────────
  let pendingDelete = null;

  function openConfirm(sheet, row) {
    pendingDelete = { sheet, row };
    const co = document.getElementById('confirmOverlay');
    if (co) co.classList.add('active');
  }

  (function bindConfirmModals() {
    const confirmCancel = document.getElementById('confirmCancel');
    const confirmDelete = document.getElementById('confirmDelete');
    const confirmOverlay = document.getElementById('confirmOverlay');
    if (!confirmCancel || !confirmDelete || !confirmOverlay) {
      console.warn('[dashboard] Confirm dialog nodes missing; delete confirmations disabled.');
      return;
    }
    confirmCancel.addEventListener('click', () => {
      confirmOverlay.classList.remove('active');
      pendingDelete = null;
    });

    confirmDelete.addEventListener('click', async function() {
    if (!pendingDelete) return;
    this.textContent = 'Deleting...';
    try {
      await apiPost({ action: 'delete', sheet: pendingDelete.sheet, row: pendingDelete.row });
      document.getElementById('confirmOverlay').classList.remove('active');
      if (pendingDelete.sheet === 'mill') {
        await loadMillData();
      } else if (pendingDelete.sheet === 'ttp') {
        ttpLoaded = false; await loadTTPData();
      } else if (pendingDelete.sheet === 'grievance') {
        grvLoaded = false; await loadGrvData();
      }
      pendingDelete = null;
    } catch(err) {
      alert('Error deleting: ' + err.message);
    }
    this.textContent = 'Yes, Delete';
  });
  })();

  // ─── MILL DATA ──────────────────────────────────────────
  let millFilteredRows = [];
  let millProfileVariantRows_ = [];
  let millSortKey = null;
  let millSortAsc = true;
  let millTableDelegationBound = false;
  let millColumnFilters = {};
  let millFilterOptions = {};

  /** Normalize sheet header key (trim, NBSP → space, lower). */
  function millHeaderNorm_(k) {
    return String(k || '').replace(/\u00a0/g, ' ').trim().toLowerCase();
  }

  /**
   * Resolve Quarter / Year from row regardless of exact header spelling.
   * Direct keys first, then scan all properties (handles "Quarter ", "QUARTER", "Kuartal", etc.).
   */
  function pickMillColumnValue_(row, wanted) {
    if (!row || typeof row !== 'object') return '';
    wanted = String(wanted || '').toLowerCase();
    const tryVals = wanted === 'quarter'
      ? [row['QUARTER'], row['Quarter'], row['quarter']]
      : [row['YEAR'], row['Year'], row['year']];
    for (let i = 0; i < tryVals.length; i++) {
      const v = tryVals[i];
      if (v != null && String(v).trim() !== '') return String(v).trim();
    }
    const keys = Object.keys(row);
    for (let j = 0; j < keys.length; j++) {
      const k = keys[j];
      if (k === '_row' || (String(k).length && String(k)[0] === '_')) continue;
      const nk = millHeaderNorm_(k);
      if (wanted === 'quarter') {
        if (nk === 'quarter' || nk === 'qtr' || nk === 'kuartal') {
          const v = row[k];
          if (v != null && String(v).trim() !== '') return String(v).trim();
        }
      } else {
        if (nk === 'year' || nk === 'yr' || nk === 'tahun') {
          if (/planted|tanam|ffb|tml|issue|expiry|commissioning|average|capacity|mill name/i.test(k)) continue;
          const v = row[k];
          if (v != null && String(v).trim() !== '') return String(v).trim();
        }
      }
    }
    return '';
  }

  function millQuarterVal(row) {
    return pickMillColumnValue_(row, 'quarter');
  }
  function millYearVal(row) {
    return pickMillColumnValue_(row, 'year');
  }
  function parseMillYearSort(v) {
    const s = String(v || '').trim();
    const m = s.match(/(19|20)\d{2}/);
    if (m) return parseInt(m[0], 10);
    const n = parseInt(s, 10);
    return isNaN(n) ? 0 : n;
  }
  function parseMillQuarterSort(v) {
    const s = String(v || '').trim().toUpperCase().replace(/\s+/g, '');
    let m = s.match(/^Q([1-4])$/);
    if (m) return parseInt(m[1], 10);
    m = s.match(/^([1-4])$/);
    if (m) return parseInt(m[1], 10);
    m = s.match(/TR[IY]M?(?:ESTER)?[:\s-]?([1-4])/);
    if (m) return parseInt(m[1], 10);
    m = s.match(/QUARTER[:\s-]?([1-4])/);
    if (m) return parseInt(m[1], 10);
    return 99;
  }
  function normalizeMillApiRow(row) {
    if (!row || typeof row !== 'object') return row;
    const o = Object.assign({}, row);
    const q = pickMillColumnValue_(o, 'quarter');
    if (q) o['QUARTER'] = q;
    const y = pickMillColumnValue_(o, 'year');
    if (y) o['YEAR'] = y;
    return o;
  }

  function bindMillTableDelegationOnce() {
    if (millTableDelegationBound) return;
    const body = document.getElementById('millTableBody');
    if (!body) return;
    millTableDelegationBound = true;
    body.addEventListener('click', function(e) {
      const row = e.target.closest('.mill-row-clickable');
      if (!row || !body.contains(row)) return;
      const idx = parseInt(row.dataset.idx, 10);
      if (isNaN(idx) || !millFilteredRows[idx]) return;
      openMillProfile(millFilteredRows[idx]);
    });
  }

  function prepareMillRowPerfCache(row) {
    if (!row || typeof row !== 'object') return row;
    const nblLower = String(row['BUYER NO BUY LIST'] || '').toLowerCase();
    const searchBlob = [
      row['QUARTER'],
      row['YEAR'],
      row['Quarter'],
      row['Year'],
      row['MILL NAME'],
      row['GROUP NAME'],
      row['COMPANY NAME'],
      row['PROVINCE'],
      row['TRADER NAME'],
      row['RISK LEVEL'],
      row['RESULT RISK LEVEL'],
    ].map(function(v) {
      return String(v || '').toLowerCase();
    }).join('|');
    row._sddNblLower = nblLower;
    row._sddSearchBlob = searchBlob;
    return row;
  }

  /** Prefer computed result risk when present; else input risk (for summary cards). */
  function millResolvedRiskLevelForStats_(d) {
    const rr = String(d && d['RESULT RISK LEVEL'] != null ? d['RESULT RISK LEVEL'] : '').trim();
    if (rr) return rr;
    return String(d && d['RISK LEVEL'] != null ? d['RISK LEVEL'] : '').trim();
  }

  // ─── MILL PDF EXPORT (toolbar pattern aligned with Monitoring TTM/TTP) ──
  const MILL_PDF_EXPORT_COLS = [
    { key: 'QUARTER', label: 'Quarter' },
    { key: 'YEAR', label: 'Year' },
    { key: 'GROUP NAME', label: 'Group' },
    { key: 'COMPANY NAME', label: 'Company' },
    { key: 'MILL NAME', label: 'Mill' },
    { key: 'UML ID', label: 'UML ID' },
    { key: 'PROVINCE', label: 'Province' },
    { key: 'SUPPLIER STATUS', label: 'Supplier Status' },
    { key: 'RISK LEVEL', label: 'Risk Level' },
    { key: 'RESULT RISK LEVEL', label: 'Result Risk Level' },
    { key: 'BUYER NO BUY LIST', label: 'No Buy List' },
    { key: '__NBL_BY__', label: 'NBL by (riser)' },
    { key: 'CERTIFICATION', label: 'Certification' },
    { key: 'FACILITY NAME CPO', label: 'Facility CPO' },
    { key: 'FACILITY NAME PK', label: 'Facility PK' },
    { key: 'PRODUCT SUPPLY', label: 'Product Supply' },
  ];
  const MILL_PDF_COL_DEFAULT_KEYS = ['QUARTER', 'YEAR', 'GROUP NAME', 'COMPANY NAME', 'MILL NAME', 'PROVINCE', 'SUPPLIER STATUS', 'RISK LEVEL', 'BUYER NO BUY LIST', '__NBL_BY__', 'CERTIFICATION'];
  let millPdfColSelected = new Set(MILL_PDF_COL_DEFAULT_KEYS);
  let millPdfDimFilters = {
    quarter: new Set(),
    year: new Set(),
    group: new Set(),
    province: new Set(),
  };
  let millPdfDimOptions = { quarter: [], year: [], group: [], province: [] };

  function millPdfEscHtml(s) {
    return String(s === undefined || s === null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function millPdfTokenForCell(val) {
    const t = String(val || '').trim();
    return t ? t : '__EMPTY__';
  }

  function millPdfLabelForToken(tok) {
    return tok === '__EMPTY__' ? '(kosong)' : tok;
  }

  function millRegistryChipFilter() {
    return currentFilter === 'Task List' ? 'All' : currentFilter;
  }

  const MILL_TABLE_FILTER_COLS = [
    'RESULT RISK LEVEL','GROUP NAME','COMPANY NAME','MILL NAME','PROVINCE',
    'SUPPLIER STATUS','BUYER NO BUY LIST','CERTIFICATION',
    'FACILITY NAME CPO','FACILITY NAME PK','PRODUCT SUPPLY'
  ];

  function getMillFilterCellValue(row, colKey) {
    if (colKey === 'QUARTER') return String(millQuarterVal(row) || '—');
    if (colKey === 'YEAR') return String(millYearVal(row) || '—');
    return String((row && row[colKey]) || '—').trim() || '—';
  }

  function millRefreshFilterOptions(baseRows) {
    millFilterOptions = {};
    MILL_TABLE_FILTER_COLS.forEach(function(col) {
      millFilterOptions[col] = Array.from(new Set(baseRows.map(function(r) {
        return getMillFilterCellValue(r, col);
      }))).sort(function(a, b) {
        return String(a).localeCompare(String(b), 'id', { sensitivity: 'base' });
      });
      if (!Array.isArray(millColumnFilters[col])) millColumnFilters[col] = [];
    });
  }

  function millRowMatchesColumnFilters(row) {
    return MILL_TABLE_FILTER_COLS.every(function(col) {
      var active = millColumnFilters[col];
      if (!Array.isArray(active) || !active.length) return true;
      return active.indexOf(getMillFilterCellValue(row, col)) !== -1;
    });
  }

  function millRowMatchesChipAndSearch(d) {
    const nbl = d._sddNblLower || (d['BUYER NO BUY LIST'] || '').toString().toLowerCase();
    const chip = millRegistryChipFilter();
    const matchFilter = chip === 'All' ||
      (chip === 'NBL' && (nbl === 'yes' || nbl.includes('nbl'))) ||
      (chip === 'Non-NBL' && nbl !== 'yes' && !nbl.includes('nbl'));
    const q = currentSearch;
    const matchSearch = !q || (d._sddSearchBlob || '').includes(q);
    return matchFilter && matchSearch;
  }

  function millRowMatchesPdfDimFilters(d) {
    const qTok = millPdfTokenForCell(millQuarterVal(d));
    const yTok = millPdfTokenForCell(millYearVal(d));
    const gTok = millPdfTokenForCell(d['GROUP NAME']);
    const pTok = millPdfTokenForCell(d['PROVINCE']);
    if (millPdfDimFilters.quarter.size && !millPdfDimFilters.quarter.has(qTok)) return false;
    if (millPdfDimFilters.year.size && !millPdfDimFilters.year.has(yTok)) return false;
    if (millPdfDimFilters.group.size && !millPdfDimFilters.group.has(gTok)) return false;
    if (millPdfDimFilters.province.size && !millPdfDimFilters.province.has(pTok)) return false;
    return true;
  }

  function sortMillRowsForDisplay(filtered) {
    let sorted = filtered;
    if (millSortKey === 'YEAR') {
      sorted = filtered.slice().sort(function(a, b) {
        const ya = parseMillYearSort(millYearVal(a));
        const yb = parseMillYearSort(millYearVal(b));
        if (ya !== yb) return millSortAsc ? (ya - yb) : (yb - ya);
        const qa = parseMillQuarterSort(millQuarterVal(a));
        const qb = parseMillQuarterSort(millQuarterVal(b));
        return millSortAsc ? (qa - qb) : (qb - qa);
      });
    } else if (millSortKey === 'QUARTER') {
      sorted = filtered.slice().sort(function(a, b) {
        const qa = parseMillQuarterSort(millQuarterVal(a));
        const qb = parseMillQuarterSort(millQuarterVal(b));
        if (qa !== qb) return millSortAsc ? (qa - qb) : (qb - qa);
        const ya = parseMillYearSort(millYearVal(a));
        const yb = parseMillYearSort(millYearVal(b));
        return millSortAsc ? (ya - yb) : (yb - ya);
      });
    }
    return sorted;
  }

  function bindMillHeaderFiltersOnce() {
    const table = document.getElementById('millTable');
    if (!table || table.dataset.millFilterBound === '1') return;
    table.dataset.millFilterBound = '1';
    table.addEventListener('click', function(e) {
      const btn = e.target.closest('[data-mill-filter-btn]');
      if (!btn || !table.contains(btn)) return;
      e.preventDefault();
      e.stopPropagation();
      document.querySelectorAll('.mill-col-filter-menu').forEach(function(m) { m.remove(); });
      const col = btn.getAttribute('data-mill-filter-btn') || '';
      if (!col) return;
      const options = millFilterOptions[col] || [];
      let selectedVals = (millColumnFilters[col] && millColumnFilters[col].length) ? millColumnFilters[col].slice() : options.slice();
      const menu = document.createElement('div');
      menu.className = 'mill-col-filter-menu';
      menu.innerHTML =
        '<div class="mill-col-filter-head">' + col + '</div>'
        + '<input type="text" class="mill-col-filter-search" placeholder="Search value...">'
        + '<div class="mill-col-filter-actions">'
        + '<button type="button" data-act="all">Select all</button>'
        + '<button type="button" data-act="none">Clear all</button>'
        + '</div>'
        + '<div class="mill-col-filter-list"></div>'
        + '<div class="mill-col-filter-foot">'
        + '<button type="button" data-act="cancel">Cancel</button>'
        + '<button type="button" data-act="ok" class="is-primary">OK</button>'
        + '</div>';
      document.body.appendChild(menu);
      const br = btn.getBoundingClientRect();
      menu.style.top = (br.bottom + 6 + window.scrollY) + 'px';
      menu.style.left = Math.max(8, br.right - 260 + window.scrollX) + 'px';

      const listEl = menu.querySelector('.mill-col-filter-list');
      const searchEl = menu.querySelector('.mill-col-filter-search');
      function renderItems() {
        if (!listEl) return;
        const q = String((searchEl && searchEl.value) || '').trim().toLowerCase();
        const items = options.filter(function(v) { return !q || String(v).toLowerCase().indexOf(q) !== -1; });
        listEl.innerHTML = items.map(function(v) {
          const checked = selectedVals.indexOf(v) !== -1;
          return '<label class="mill-col-filter-item"><input type="checkbox" value="' + String(v).replace(/"/g,'&quot;') + '"' + (checked ? ' checked' : '') + '> <span>' + v + '</span></label>';
        }).join('') || '<div class="mill-col-filter-empty">No values</div>';
        listEl.querySelectorAll('input[type="checkbox"]').forEach(function(cb) {
          cb.addEventListener('change', function() {
            const val = cb.value;
            if (cb.checked) {
              if (selectedVals.indexOf(val) === -1) selectedVals.push(val);
            } else {
              selectedVals = selectedVals.filter(function(x) { return x !== val; });
            }
          });
        });
      }
      renderItems();
      if (searchEl) searchEl.addEventListener('input', renderItems);
      menu.querySelectorAll('[data-act]').forEach(function(ab) {
        ab.addEventListener('click', function() {
          const act = ab.getAttribute('data-act');
          if (act === 'all') { selectedVals = options.slice(); renderItems(); return; }
          if (act === 'none') { selectedVals = []; renderItems(); return; }
          if (act === 'cancel') { menu.remove(); return; }
          if (act === 'ok') {
            millColumnFilters[col] = (selectedVals.length === 0 || selectedVals.length === options.length) ? [] : selectedVals.slice();
            menu.remove();
            scheduleRenderMillTable();
          }
        });
      });
      setTimeout(function() {
        function closeOut(ev) {
          if (!menu.contains(ev.target) && ev.target !== btn) {
            menu.remove();
            document.removeEventListener('click', closeOut, true);
          }
        }
        document.addEventListener('click', closeOut, true);
      }, 0);
    });
  }

  function millRowsAfterRegistryDimFilters() {
    return allData.filter(function(d) {
      return millRowMatchesChipAndSearch(d) && millRowMatchesPdfDimFilters(d);
    });
  }

  function getMillRowsForPdfExport() {
    return sortMillRowsForDisplay(millRowsAfterRegistryDimFilters());
  }

  function updateMillPdfExportScope() {
    const el = document.getElementById('millPdfExportScopeText');
    if (!el) return;
    if (!allData.length) {
      el.textContent = 'Muat data mill terlebih dahulu';
      return;
    }
    const pdfN = getMillRowsForPdfExport().length;
    const tableN = sortMillRowsForDisplay(millRowsAfterRegistryDimFilters().filter(millRowMatchesColumnFilters)).length;
    if (tableN === pdfN) {
      el.textContent = pdfN + ' rows · table view matches export';
    } else {
      el.textContent = pdfN + ' rows exported · ' + tableN + ' rows shown in table (column filters active)';
    }
  }

  function collectMillPdfDimTokens_() {
    const sq = new Set(); const sy = new Set(); const sg = new Set(); const sp = new Set();
    allData.forEach(function(row) {
      sq.add(millPdfTokenForCell(millQuarterVal(row)));
      sy.add(millPdfTokenForCell(millYearVal(row)));
      sg.add(millPdfTokenForCell(row['GROUP NAME']));
      sp.add(millPdfTokenForCell(row['PROVINCE']));
    });
    function sortQuarterTokens(arr) {
      return arr.slice().sort(function(a, b) {
        if (a === '__EMPTY__') return 1;
        if (b === '__EMPTY__') return -1;
        const c = parseMillQuarterSort(a) - parseMillQuarterSort(b);
        return c !== 0 ? c : String(a).localeCompare(String(b), 'id');
      });
    }
    function sortYearTokens(arr) {
      return arr.slice().sort(function(a, b) {
        if (a === '__EMPTY__') return 1;
        if (b === '__EMPTY__') return -1;
        const c = parseMillYearSort(b) - parseMillYearSort(a);
        return c !== 0 ? c : String(a).localeCompare(String(b), 'id');
      });
    }
    function sortStrTokens(arr) {
      return arr.slice().sort(function(a, b) {
        if (a === '__EMPTY__') return 1;
        if (b === '__EMPTY__') return -1;
        return String(a).localeCompare(String(b), 'id', { sensitivity: 'base' });
      });
    }
    millPdfDimOptions.quarter = sortQuarterTokens(Array.from(sq));
    millPdfDimOptions.year = sortYearTokens(Array.from(sy));
    millPdfDimOptions.group = sortStrTokens(Array.from(sg));
    millPdfDimOptions.province = sortStrTokens(Array.from(sp));
  }

  function millPdfDimItemHtml(dim, tok, idx) {
    const id = 'mill-pdf-chk-' + dim + '-' + idx;
    const checked = millPdfDimFilters[dim] && millPdfDimFilters[dim].has(tok) ? ' checked' : '';
    const lab = millPdfEscHtml(millPdfLabelForToken(tok));
    const escTok = millPdfEscHtml(tok);
    const searchHay = millPdfEscHtml(millPdfLabelForToken(tok).toLowerCase().replace(/"/g, ''));
    return '<div class="ttp-dropdown-item mill-pdf-dim-item" data-mill-pdf-dim="' + dim + '" data-mill-pdf-search="' + searchHay + '">'
      + '<input type="checkbox" id="' + id + '" data-mill-pdf-dim="' + dim + '" data-mill-pdf-val="' + escTok + '"' + checked + ' />'
      + '<label for="' + id + '">' + lab + '</label></div>';
  }

  function renderMillPdfDimListEl(dim, listId) {
    const el = document.getElementById(listId);
    if (!el) return;
    const arr = millPdfDimOptions[dim] || [];
    el.innerHTML = arr.map(function(tok, idx) { return millPdfDimItemHtml(dim, tok, idx); }).join('');
  }

  function millPdfApplyGroupSearchFilter() {
    const inp = document.getElementById('millPdfGroupSearch');
    const q = (inp && inp.value || '').trim().toLowerCase();
    const items = document.querySelectorAll('#millPdfDimListGroup .mill-pdf-dim-item');
    items.forEach(function(node) {
      const hay = (node.getAttribute('data-mill-pdf-search') || '').toLowerCase();
      node.style.display = !q || hay.indexOf(q) !== -1 ? '' : 'none';
    });
  }

  function millPdfRebuildDimPanels() {
    collectMillPdfDimTokens_();
    renderMillPdfDimListEl('quarter', 'millPdfDimListQuarter');
    renderMillPdfDimListEl('year', 'millPdfDimListYear');
    renderMillPdfDimListEl('group', 'millPdfDimListGroup');
    renderMillPdfDimListEl('province', 'millPdfDimListProvince');
    millPdfApplyGroupSearchFilter();
  }

  function millPdfRenderColPanel() {
    const el = document.getElementById('millPdfColList');
    if (!el) return;
    el.innerHTML = MILL_PDF_EXPORT_COLS.map(function(col, idx) {
      const id = 'mill-pdf-col-' + idx + '-' + col.key.replace(/\s+/g, '-');
      const checked = millPdfColSelected.has(col.key) ? ' checked' : '';
      return '<div class="ttp-dropdown-item"><input type="checkbox" id="' + id + '" data-mill-pdf-col="' + millPdfEscHtml(col.key) + '"' + checked + ' />'
        + '<label for="' + id + '">' + millPdfEscHtml(col.label) + '</label></div>';
    }).join('');
  }

  function millPdfResetDim(dim) {
    if (millPdfDimFilters[dim]) millPdfDimFilters[dim].clear();
    document.querySelectorAll('#millPdfFilterDimsPanel input[data-mill-pdf-dim="' + dim + '"][data-mill-pdf-val]').forEach(function(cb) { cb.checked = false; });
    updateMillPdfExportScope();
    scheduleRenderMillTable();
  }

  function millPdfResetAllDims() {
    millPdfDimFilters = { quarter: new Set(), year: new Set(), group: new Set(), province: new Set() };
    document.querySelectorAll('#millPdfFilterDimsPanel input[data-mill-pdf-val]').forEach(function(cb) { cb.checked = false; });
    updateMillPdfExportScope();
    scheduleRenderMillTable();
  }

  async function millExportToPdf() {
    const toastErr = function(msg) {
      if (typeof window.showSddToast === 'function') window.showSddToast(msg, 'error');
    };
    const rows = getMillRowsForPdfExport();
    if (!rows.length) {
      toastErr('Tidak ada baris untuk diekspor. Sesuaikan chip, pencarian, filter registry, atau filter kolom.');
      return;
    }
    const cols = MILL_PDF_EXPORT_COLS.filter(function(c) { return millPdfColSelected.has(c.key); });
    if (!cols.length) {
      toastErr('Pilih minimal satu kolom di menu Kolom PDF.');
      return;
    }

    const btn = document.getElementById('btn-mill-export-pdf');
    const prevHtml = btn ? btn.innerHTML : '';
    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="ttp-btn-icon">…</span> Menghasilkan…'; }

    try {
      const nblLists = await ensureNblListsForCheck_();
      const JsPDFLib = getJsPDF();
      const doc = new JsPDFLib({ unit: 'mm', format: 'a4', orientation: 'landscape' });
      const RED = [139, 26, 26];
      const WHITE = [255, 255, 255];
      const GRY = [110, 96, 96];
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(13);
      doc.setTextColor.apply(doc, RED);
      doc.text('Mill Onboarding — Registry Export', 14, 14);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.setTextColor.apply(doc, GRY);
      const stamp = new Date().toLocaleString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
      doc.text('Diekspor: ' + stamp + ' · ' + rows.length + ' baris', 14, 20);

      const body = rows.map(function(row) {
        return cols.map(function(c) {
          if (c.key === '__NBL_BY__') {
            if (!millIsNblYes_(row['BUYER NO BUY LIST'])) return '';
            var info = millNblByInfoFromMatches_(millNblSourceMatchesForRow_(row, nblLists));
            return String(info && info.label ? info.label : '').replace(/\r?\n/g, ' ');
          }
          const v = row[c.key];
          return String(v !== undefined && v !== null ? v : '').replace(/\r?\n/g, ' ');
        });
      });
      const nblByColIdx = cols.findIndex(function(c) { return c.key === '__NBL_BY__'; });
      const columnStyles = {};
      if (nblByColIdx !== -1) {
        columnStyles[nblByColIdx] = { cellWidth: 42, overflow: 'linebreak' };
      }
      doc.autoTable({
        head: [cols.map(function(c) { return c.label; })],
        body: body,
        startY: 24,
        margin: { left: 10, right: 10 },
        styles: { fontSize: 6.2, cellPadding: 1.4, textColor: [26, 10, 10] },
        headStyles: { fillColor: RED, textColor: WHITE, fontStyle: 'bold', fontSize: 7 },
        alternateRowStyles: { fillColor: [253, 250, 250] },
        theme: 'striped',
        columnStyles: columnStyles,
      });

      const fname = 'Mill-Registry-' + new Date().toISOString().slice(0, 10) + '.pdf';
      doc.save(fname);
      if (typeof window.showSddToast === 'function') window.showSddToast('PDF berhasil diunduh.', 'success');
    } catch (err) {
      toastErr('Export PDF gagal: ' + (err.message || err));
    } finally {
      if (btn) { btn.disabled = false; btn.innerHTML = prevHtml; }
    }
  }

  (function bindMillPdfToolbarOnce() {
    const btnPdf = document.getElementById('btn-mill-export-pdf');
    const bFilter = document.getElementById('millPdfBtnFilterDims');
    const bCols = document.getElementById('millPdfBtnCols');
    const pFilter = document.getElementById('millPdfFilterDimsPanel');
    const pCols = document.getElementById('millPdfColsPanel');
    const gSearch = document.getElementById('millPdfGroupSearch');
    if (!btnPdf || !bFilter || !bCols || !pFilter || !pCols) return;

    bFilter.addEventListener('click', function(e) {
      e.stopPropagation();
      openTTPDropdown('millPdfBtnFilterDims', 'millPdfFilterDimsPanel', 'millPdfColsPanel', 'millPdfBtnCols');
    });
    bCols.addEventListener('click', function(e) {
      e.stopPropagation();
      openTTPDropdown('millPdfBtnCols', 'millPdfColsPanel', 'millPdfFilterDimsPanel', 'millPdfBtnFilterDims');
    });

    if (!window.__sddMillPdfDropdownOutsideClickBound) {
      window.__sddMillPdfDropdownOutsideClickBound = true;
      document.addEventListener('click', function(e) {
        const w1 = document.getElementById('millPdfFilterWrap');
        const w2 = document.getElementById('millPdfColWrap');
        if (w1 && pFilter && !w1.contains(e.target)) {
          pFilter.classList.remove('open');
          bFilter.classList.remove('active');
        }
        if (w2 && pCols && !w2.contains(e.target)) {
          pCols.classList.remove('open');
          bCols.classList.remove('active');
        }
      });
    }

    pFilter.addEventListener('change', function(e) {
      const t = e.target;
      if (!t || t.type !== 'checkbox' || !t.dataset.millPdfDim || !t.dataset.millPdfVal) return;
      const dim = t.dataset.millPdfDim;
      const tok = t.dataset.millPdfVal;
      if (!millPdfDimFilters[dim]) return;
      if (t.checked) millPdfDimFilters[dim].add(tok);
      else millPdfDimFilters[dim].delete(tok);
      updateMillPdfExportScope();
      scheduleRenderMillTable();
    });

    const resetAll = document.getElementById('millPdfDimResetAll');
    if (resetAll) resetAll.addEventListener('click', function(e) { e.stopPropagation(); millPdfResetAllDims(); });

    pFilter.querySelectorAll('[data-mill-pdf-dim-reset]').forEach(function(b) {
      b.addEventListener('click', function(e) {
        e.stopPropagation();
        const dim = this.getAttribute('data-mill-pdf-dim-reset');
        if (dim) millPdfResetDim(dim);
      });
    });

    if (gSearch) {
      gSearch.addEventListener('input', function() { millPdfApplyGroupSearchFilter(); });
      gSearch.addEventListener('click', function(e) { e.stopPropagation(); });
    }

    pCols.addEventListener('change', function(e) {
      const t = e.target;
      if (!t || t.type !== 'checkbox' || !t.dataset.millPdfCol) return;
      const k = t.dataset.millPdfCol;
      if (t.checked) millPdfColSelected.add(k);
      else millPdfColSelected.delete(k);
      if (millPdfColSelected.size === 0) {
        millPdfColSelected.add(k);
        t.checked = true;
        if (typeof window.showSddToast === 'function') window.showSddToast('Minimal satu kolom harus dipilih.', 'warning');
      }
    });

    const bAll = document.getElementById('millPdfColAll');
    const bDef = document.getElementById('millPdfColDefault');
    if (bAll) bAll.addEventListener('click', function(e) {
      e.stopPropagation();
      millPdfColSelected = new Set(MILL_PDF_EXPORT_COLS.map(function(c) { return c.key; }));
      millPdfRenderColPanel();
    });
    if (bDef) bDef.addEventListener('click', function(e) {
      e.stopPropagation();
      millPdfColSelected = new Set(MILL_PDF_COL_DEFAULT_KEYS);
      millPdfRenderColPanel();
    });

    btnPdf.addEventListener('click', function() { millExportToPdf(); });
    millPdfRenderColPanel();
    updateMillPdfExportScope();
  })();

  const scheduleRenderMillTable = makeRafScheduler(function() {
    renderMillTable();
  });

  (function bindMillTableSortOnce() {
    const tr = document.querySelector('#millTable thead tr');
    if (!tr || tr.dataset.millSortBound) return;
    tr.dataset.millSortBound = '1';
    tr.addEventListener('click', function(e) {
      if (e.target && e.target.closest && e.target.closest('[data-mill-filter-btn]')) return;
      const th = e.target.closest('[data-mill-sort]');
      if (!th || !tr.contains(th)) return;
      const key = th.getAttribute('data-mill-sort');
      if (millSortKey === key) millSortAsc = !millSortAsc;
      else { millSortKey = key; millSortAsc = true; }
      scheduleRenderMillTable();
    });
  })();

  async function loadMillDataImpl() {
    const loading = document.getElementById('mill-loading');
    const errorEl = document.getElementById('mill-error');
    const table = document.getElementById('millTable');
    if (!loading || !errorEl || !table) {
      console.warn('[dashboard] Mill panel DOM missing; skip loadMillData.');
      return;
    }
    try {
      loading.style.display = 'block';
      errorEl.style.display = 'none';
      table.style.display = 'none';
      const rawRows = await apiGet('mill');
      const rawData = (Array.isArray(rawRows) ? rawRows : []).map(normalizeMillApiRow);

      // ── DEDUP: per Mill ID / UML ID, keep the row with most non-empty fields ──
      const dedupMap = new Map();
      const rowsWithoutId = [];
      function countFilled(row) {
        return Object.values(row).filter(v => v !== undefined && v !== null && String(v).trim() !== '').length;
      }
      rawData.forEach(function(row) {
        // Try multiple possible ID key names from the sheet
        const id = (row['Mill ID'] || row['MILL ID'] || row['UML ID'] || row['UML_ID'] || '').toString().trim();
        if (!id) {
          rowsWithoutId.push(row);
          return;
        }
        const existing = dedupMap.get(id);
        if (!existing || countFilled(row) > countFilled(existing)) {
          dedupMap.set(id, Object.assign({}, row));
        }
      });
      const deduped = Array.from(dedupMap.values());
      allData = rowsWithoutId.concat(deduped).map(prepareMillRowPerfCache);
      // If no IDs were found (different column name), fall back to raw
      if (allData.length === 0 && rawData.length > 0) allData = rawData.map(prepareMillRowPerfCache);

      // ── SORT by Date Imported DESC for Top 5 ──
      function parseDateImported(row) {
        const raw = row['Date Imported'] || row['DATE IMPORTED'] || row['Last Updated'] || row['LAST UPDATED'] || row['Timestamp'] || row['TIMESTAMP'] || '';
        if (!raw) return 0;
        const d = new Date(String(raw).trim());
        return isNaN(d.getTime()) ? 0 : d.getTime();
      }
      const sortedByDate = allData.slice().sort((a, b) => parseDateImported(b) - parseDateImported(a));
      const top5 = sortedByDate.filter(d => parseDateImported(d) > 0).slice(0, 5);

      const top5Card = document.getElementById('top5-card');
      const top5List = document.getElementById('top5-list');
      if (top5.length > 0) {
        top5Card.style.display = 'block';
        const rankColors = ['#8B1A1A','#A52020','#C03030','#9C8080','#7C6565'];
        top5List.innerHTML = top5.map((d, i) => {
          const company = d['Company Name'] || d['COMPANY NAME'] || d['Group Name'] || d['GROUP NAME'] || d['Mill Name'] || d['MILL NAME'] || '(Unknown)';
          const mill    = d['Mill Name']    || d['MILL NAME']    || '';
          const millId  = d['Mill ID']  || d['UML ID'] || d['MILL ID'] || '';
          const rawDate = d['Date Imported'] || d['DATE IMPORTED'] || d['Last Updated'] || d['LAST UPDATED'] || d['Timestamp'] || d['TIMESTAMP'] || '';
          const dObj = new Date(String(rawDate).trim());
          const dateStr = isNaN(dObj.getTime()) ? rawDate : dObj.toLocaleString('id-ID', { day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit', hour12:false });
          return `<div style="display:flex;align-items:center;gap:12px;padding:10px 14px;border-radius:10px;background:${i===0?'rgba(139,26,26,0.05)':'rgba(74,28,28,0.02)'};border:1px solid ${i===0?'rgba(139,26,26,0.15)':'rgba(74,28,28,0.06)'};">
            <div style="width:24px;height:24px;border-radius:50%;background:${rankColors[i]};display:flex;align-items:center;justify-content:center;flex-shrink:0;">
              <span style="font-size:11px;font-weight:700;color:white;">${i+1}</span>
            </div>
            <div style="flex:1;min-width:0;">
              <div style="font-size:13px;font-weight:600;color:#1A0A0A;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${company}${mill && mill !== company ? ' <span style="color:#9C8080;font-weight:400;font-size:12px;">· '+mill+'</span>' : ''}</div>
              <div style="font-size:11px;color:#9C8080;margin-top:2px;">${millId ? millId+' · ' : ''}${dateStr}</div>
            </div>
            <div style="flex-shrink:0;">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#C03030" stroke-width="2.2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
            </div>
          </div>`;
        }).join('');
      } else {
        top5Card.style.display = 'none';
      }

      document.getElementById('stat-total').textContent = allData.length;
      document.getElementById('stat-groups').textContent = new Set(allData.map(d => d['GROUP NAME']).filter(Boolean)).size;
      document.getElementById('stat-high-risk').textContent = allData.filter(d => (millResolvedRiskLevelForStats_(d) || '').toLowerCase().includes('high')).length;
      document.getElementById('stat-nbl').textContent = allData.filter(d => (d['BUYER NO BUY LIST']||'').toLowerCase() === 'yes' || (d['BUYER NO BUY LIST']||'').toLowerCase().includes('nbl')).length;
      loading.style.display = 'none';
      table.style.display = 'table';
      millPdfDimFilters = { quarter: new Set(), year: new Set(), group: new Set(), province: new Set() };
      millPdfRebuildDimPanels();
      scheduleRenderMillTable();
    } catch(err) {
      loading.style.display = 'none';
      errorEl.style.display = 'block';
      errorEl.textContent = 'Gagal memuat data: ' + err.message;
    }
  }

  async function loadMillData() {
    if (millLoadPromise) return millLoadPromise;
    millLoadPromise = loadMillDataImpl();
    try {
      await millLoadPromise;
    } finally {
      millLoadPromise = null;
    }
  }

  function nblBadge(val) {
    if (!val || val === '-') return '<span style="color:#ccc">—</span>';
    const v = val.toString().toLowerCase();
    const isNBL = v === 'yes' || v.includes('nbl') || v.includes('no buy');
    return isNBL
      ? `<span class="status-badge" style="background:rgba(192,57,43,0.1);color:#c0392b">${val}</span>`
      : `<span class="status-badge s-active">${val}</span>`;
  }

  function supplierBadge(val) {
    if (!val) return '—';
    const lower = val.toLowerCase();
    let cls = 's-pending';
    if (['active','compliant'].some(k => lower.includes(k))) cls = 's-active';
    else if (['review','pending','conditional'].some(k => lower.includes(k))) cls = 's-review';
    return `<span class="status-badge ${cls}">${val}</span>`;
  }

  function riskBadgeLevel(val) {
    if (!val || val === '-') return '<span style="color:#ccc">—</span>';
    const lower = val.toString().toLowerCase();
    let bg = 'rgba(124,101,101,0.1)', color = '#7C6565';
    if (lower.includes('high') || lower.includes('tinggi')) { bg = 'rgba(192,57,43,0.1)'; color = '#c0392b'; }
    else if (lower.includes('med') || lower.includes('sedang')) { bg = 'rgba(200,168,75,0.15)'; color = '#8a6e1a'; }
    else if (lower.includes('low') || lower.includes('rendah')) { bg = 'rgba(39,174,96,0.1)'; color = '#1e8449'; }
    return `<span class="status-badge" style="background:${bg};color:${color}">${val}</span>`;
  }

  /** Glossy pill for Mill Registry "Result Risk Level" only (HIGH / MEDIUM / LOW). */
  function resultRiskLevelPill(val) {
    if (val == null || val === '-') {
      return '<span class="mill-rrl mill-rrl--empty" aria-hidden="true">—</span>';
    }
    const raw = String(val).trim();
    if (!raw) {
      return '<span class="mill-rrl mill-rrl--empty" aria-hidden="true">—</span>';
    }
    const lower = raw.toLowerCase();
    let tier = 'other';
    if (lower.includes('high') || lower.includes('tinggi')) tier = 'high';
    else if (lower.includes('med') || lower.includes('sedang')) tier = 'medium';
    else if (lower.includes('low') || lower.includes('rendah')) tier = 'low';
    const lbl = millPdfEscHtml(raw);
    const aria = millPdfEscHtml('Result risk level: ' + raw);
    return '<span class="mill-rrl mill-rrl--' + tier + '" role="img" aria-label="' + aria + '">'
      + '<span class="mill-rrl-pill"><span class="mill-rrl-pill__sheen" aria-hidden="true"></span>'
      + '<span class="mill-rrl-pill__lbl">' + lbl + '</span></span></span>';
  }

  function renderMillTable() {
    const body = document.getElementById('millTableBody');
    if (!body) return;
    bindMillTableDelegationOnce();
    bindMillHeaderFiltersOnce();
    const baseFiltered = millRowsAfterRegistryDimFilters();
    millRefreshFilterOptions(baseFiltered);
    const filtered = baseFiltered.filter(millRowMatchesColumnFilters);
    const sorted = sortMillRowsForDisplay(filtered);
    millFilteredRows = sorted;
    updateMillPdfExportScope();

    const theadRow = document.querySelector('#millTable thead tr');
    if (millSortKey && theadRow && !theadRow.querySelector('[data-mill-sort="' + millSortKey + '"]')) {
      millSortKey = null;
    }
    if (theadRow) {
      theadRow.querySelectorAll('[data-mill-sort]').forEach(function(th) {
        th.classList.remove('is-sorted', 'is-sorted-asc', 'is-sorted-desc');
        if (millSortKey && th.getAttribute('data-mill-sort') === millSortKey) {
          th.classList.add('is-sorted', millSortAsc ? 'is-sorted-asc' : 'is-sorted-desc');
        }
      });
      theadRow.querySelectorAll('[data-mill-filter-btn]').forEach(function(btn) {
        const key = btn.getAttribute('data-mill-filter-btn') || '';
        const hasActive = Array.isArray(millColumnFilters[key]) && millColumnFilters[key].length > 0;
        btn.classList.toggle('is-active', hasActive);
      });
    }

    body.innerHTML = sorted.length === 0
      ? `<tr><td colspan="11" style="text-align:center;padding:32px;color:#9C8A8A;">No data found</td></tr>`
      : sorted.map((d, i) => `
        <tr class="mill-row-clickable" data-idx="${i}" title="Klik untuk lihat detail lengkap">
          <td>${resultRiskLevelPill(d['RESULT RISK LEVEL'])}</td>
          <td>${d['GROUP NAME'] || '—'}</td>
          <td>${d['COMPANY NAME'] || '—'}</td>
          <td><span class="mill-name">${d['MILL NAME'] || '—'}</span><div class="mill-id">${d['UML ID'] || ''}</div></td>
          <td>${d['PROVINCE'] || '—'}</td>
          <td>${supplierBadge(d['SUPPLIER STATUS'])}</td>
          <td>${nblBadge(d['BUYER NO BUY LIST'])}</td>
          <td>${d['CERTIFICATION'] || '—'}</td>
          <td class="mill-cell-long">${d['FACILITY NAME CPO'] || '—'}</td>
          <td class="mill-cell-long">${d['FACILITY NAME PK'] || '—'}</td>
          <td class="mill-cell-long">${d['PRODUCT SUPPLY'] || '—'}</td>
        </tr>`).join('');
  }

  const btnAddMill = document.getElementById('btn-add-mill');
  if (btnAddMill) btnAddMill.addEventListener('click', () => openModal('mill', MILL_FIELDS, 'add', null));

  // ─── MILL PROFILE POPUP ─────────────────────────────────

  function millProfileSameEntityRows_(anchorRow) {
    if (!anchorRow || typeof anchorRow !== 'object') return [];
    if (!Array.isArray(allData) || !allData.length) return [anchorRow];
    function nk(r, k) {
      return String(r && r[k] != null ? r[k] : '').trim().toLowerCase();
    }
    const c = nk(anchorRow, 'COMPANY NAME');
    const m = nk(anchorRow, 'MILL NAME');
    const g = nk(anchorRow, 'GROUP NAME');
    const rows = allData.filter(function(r) {
      return nk(r, 'COMPANY NAME') === c && nk(r, 'MILL NAME') === m && nk(r, 'GROUP NAME') === g;
    });
    return rows.length ? rows : [anchorRow];
  }

  function millProfileComparePeriodDesc_(a, b) {
    const ya = parseMillYearSort(millYearVal(a));
    const yb = parseMillYearSort(millYearVal(b));
    if (ya !== yb) return yb - ya;
    const qa = parseMillQuarterSort(millQuarterVal(a));
    const qb = parseMillQuarterSort(millQuarterVal(b));
    return qb - qa;
  }

  function millProfileSortYearTokDesc_(toks) {
    return toks.slice().sort(function(a, b) {
      if (a === '__EMPTY__') return 1;
      if (b === '__EMPTY__') return -1;
      return parseMillYearSort(b) - parseMillYearSort(a);
    });
  }

  function millProfileSortQuarterTokDesc_(toks) {
    return toks.slice().sort(function(a, b) {
      if (a === '__EMPTY__') return 1;
      if (b === '__EMPTY__') return -1;
      return parseMillQuarterSort(b) - parseMillQuarterSort(a);
    });
  }

  function millProfileCollectYearToks_(siblings) {
    const s = new Set();
    siblings.forEach(function(r) { s.add(millPdfTokenForCell(millYearVal(r))); });
    return millProfileSortYearTokDesc_(Array.from(s));
  }

  function millProfileCollectQuarterToksForYear_(siblings, yearTok) {
    const s = new Set();
    siblings.forEach(function(r) {
      if (millPdfTokenForCell(millYearVal(r)) === yearTok) {
        s.add(millPdfTokenForCell(millQuarterVal(r)));
      }
    });
    return millProfileSortQuarterTokDesc_(Array.from(s));
  }

  function millProfileFillSelectToks_(sel, toks) {
    if (!sel) return;
    const list = (toks && toks.length) ? toks : ['__EMPTY__'];
    sel.innerHTML = list.map(function(t) {
      return '<option value="' + millPdfEscHtml(t) + '">' + millPdfEscHtml(millPdfLabelForToken(t)) + '</option>';
    }).join('');
  }

  function millProfileFindRowByPeriodTok_(siblings, yTok, qTok) {
    for (let i = 0; i < siblings.length; i++) {
      const r = siblings[i];
      if (millPdfTokenForCell(millYearVal(r)) === yTok && millPdfTokenForCell(millQuarterVal(r)) === qTok) {
        return r;
      }
    }
    return null;
  }

  function millProfileTitleCaseWords_(s) {
    return String(s == null ? '' : s).trim().split(/\s+/).map(function(w) {
      if (!w) return '';
      return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    }).filter(Boolean).join(' ');
  }

  function millIsNblYes_(raw) {
    var v = String(raw == null ? '' : raw).trim().toLowerCase();
    return v === 'yes' || v.includes('nbl') || v.includes('no buy');
  }

  function millNameSimilarLoose_(a, b) {
    var na = normalizeLooseKey(a);
    var nb = normalizeLooseKey(b);
    if (!na || !nb) return false;
    if (na === nb) return true;
    // "mirip" match: one contained in another, but avoid very short false positives
    if (na.length >= 4 && nb.includes(na)) return true;
    if (nb.length >= 4 && na.includes(nb)) return true;
    // Token-level fuzzy match for small spelling variants (e.g. sumatra/sumatera).
    var ta = String(a || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(function(t) { return t.length >= 3; });
    var tb = String(b || '').toLowerCase().replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(function(t) { return t.length >= 3; });
    if (!ta.length || !tb.length) return false;
    function levDist_(x, y) {
      if (x === y) return 0;
      var m = x.length, n = y.length;
      if (!m) return n;
      if (!n) return m;
      var prev = [];
      for (var j = 0; j <= n; j++) prev[j] = j;
      for (var i = 1; i <= m; i++) {
        var cur = [i];
        for (var j2 = 1; j2 <= n; j2++) {
          var cost = x.charAt(i - 1) === y.charAt(j2 - 1) ? 0 : 1;
          cur[j2] = Math.min(
            cur[j2 - 1] + 1,
            prev[j2] + 1,
            prev[j2 - 1] + cost
          );
        }
        prev = cur;
      }
      return prev[n];
    }
    function tokenNear_(x, y) {
      if (x === y) return true;
      if (x.length >= 4 && y.includes(x)) return true;
      if (y.length >= 4 && x.includes(y)) return true;
      // allow one-char typo on longer tokens
      if (x.length >= 6 && y.length >= 6 && Math.abs(x.length - y.length) <= 1) {
        return levDist_(x, y) <= 1;
      }
      return false;
    }
    var hits = 0;
    for (var i = 0; i < ta.length; i++) {
      var t1 = ta[i];
      var ok = false;
      for (var j = 0; j < tb.length; j++) {
        if (tokenNear_(t1, tb[j])) { ok = true; break; }
      }
      if (ok) hits++;
    }
    // Require at least 2 token hits to reduce false positives.
    if (hits >= 2) return true;
    return false;
  }

  function millNblSourceMatchesForRow_(row, lists) {
    var group = String(row && row['GROUP NAME'] != null ? row['GROUP NAME'] : '').trim();
    var company = String(row && row['COMPANY NAME'] != null ? row['COMPANY NAME'] : '').trim();
    var out = [];
    var seen = {};

    function pushOnce_(key, item) {
      if (seen[key]) return;
      seen[key] = true;
      out.push(item);
    }

    // Source: NBL registry (Group/Company)
    (lists.registry || []).forEach(function(r, i) {
      var groupHit = !!(group && (millNameSimilarLoose_(group, r._nblGroup) || millNameSimilarLoose_(group, r._nblCompany)));
      var companyHit = !!(company && (millNameSimilarLoose_(company, r._nblGroup) || millNameSimilarLoose_(company, r._nblCompany)));
      if (!groupHit && !companyHit) return;
      var hitBy = [];
      if (groupHit) hitBy.push('GROUP NAME');
      if (companyHit) hitBy.push('COMPANY NAME');
      pushOnce_('nbl-' + i, {
        source: 'NBL',
        by: hitBy.join(' / '),
        target: (r._nblGroup || r._nblCompany || '—'),
        riser: (r._nblRiser || '').trim(),
      });
    });

    // Source: Unilever NBL (match by GROUP/COMPANY vs Unilever COMPANY)
    (lists.unilever || []).forEach(function(r, i) {
      var groupHit = !!(group && millNameSimilarLoose_(group, r._nblCompany));
      var companyHit = !!(company && millNameSimilarLoose_(company, r._nblCompany));
      if (!groupHit && !companyHit) return;
      var hitBy = [];
      if (groupHit) hitBy.push('GROUP NAME');
      if (companyHit) hitBy.push('COMPANY NAME');
      pushOnce_('uni-' + i, {
        source: 'Unilever NBL',
        by: hitBy.join(' / '),
        target: (r._nblCompany || '—'),
        riser: (String(r._nblRiser || '').trim() || 'Unilever'),
      });
    });

    return out;
  }

  function millNblByInfoFromMatches_(matches) {
    var risers = [];
    var seen = {};
    var hasUnilever = false;
    (matches || []).forEach(function(m) {
      var src = String(m && m.source ? m.source : '').trim();
      var riser = String(m && m.riser ? m.riser : '').trim();
      if (src === 'Unilever NBL') hasUnilever = true;
      if (!riser) return;
      var key = riser.toLowerCase();
      if (seen[key]) return;
      seen[key] = true;
      risers.push(riser);
    });

    function formatRisersLabel_(vals) {
      var clean = (vals || []).map(function(v) {
        return millProfileTitleCaseWords_(String(v || '').trim());
      }).filter(Boolean);
      if (!clean.length) return '';
      clean.sort(function(a, b) { return a.localeCompare(b, 'id', { sensitivity: 'base' }); });
      return clean.join(', ');
    }

    if (risers.length) {
      return {
        label: 'NBL by ' + formatRisersLabel_(risers),
        risers: risers,
        hasUnilever: hasUnilever
      };
    }
    if (hasUnilever) {
      return { label: 'NBL by Unilever NBL', risers: [], hasUnilever: true };
    }
    return { label: 'Yes (source unresolved)', risers: [], hasUnilever: false };
  }

  async function millResolveNblByInfo_(row) {
    if (!millIsNblYes_(row && row['BUYER NO BUY LIST'])) {
      return { label: '', risers: [], hasUnilever: false, matches: [] };
    }
    var lists = await ensureNblListsForCheck_();
    var matches = millNblSourceMatchesForRow_(row || {}, lists);
    var info = millNblByInfoFromMatches_(matches);
    info.matches = matches;
    return info;
  }

  let millProfileNblInfoReqSeq_ = 0;
  async function millProfileUpdateNblSourceInfo_(row) {
    var infoEl = document.getElementById('mp-mill-nbl-source');
    if (!infoEl) return;
    var reqId = ++millProfileNblInfoReqSeq_;

    if (!millIsNblYes_(row && row['BUYER NO BUY LIST'])) {
      infoEl.innerHTML = '';
      return;
    }

    infoEl.innerHTML = 'Checking NBL source…';
    try {
      if (reqId !== millProfileNblInfoReqSeq_) return; // stale request
      var info = await millResolveNblByInfo_(row || {});
      if (reqId !== millProfileNblInfoReqSeq_) return; // stale request
      if (!info.label) {
        infoEl.innerHTML = '';
        return;
      }
      infoEl.innerHTML = '<strong>' + escHtml(info.label) + '</strong>';
    } catch (err) {
      if (reqId !== millProfileNblInfoReqSeq_) return;
      infoEl.innerHTML = '<strong>Yes (source unresolved)</strong>';
    }
  }

  function millProfileFormatHeaderLoc_(row) {
    const g = String(row && row['GROUP NAME'] != null ? row['GROUP NAME'] : '').trim();
    const p = String(row && row['PROVINCE'] != null ? row['PROVINCE'] : '').trim();
    if (!g && !p) return '—';
    if (!g) return millProfileTitleCaseWords_(p);
    if (!p) return millProfileTitleCaseWords_(g);
    return millProfileTitleCaseWords_(g) + ' — ' + millProfileTitleCaseWords_(p);
  }

  function millProfileUpdateHeaderFromRow_(row) {
    if (!row || typeof row !== 'object') return;
    const nameEl = document.getElementById('mp-mill-name');
    const locEl = document.getElementById('mp-mill-loc');
    const supEl = document.getElementById('mp-mill-supplier');
    const nblEl = document.getElementById('mp-mill-nbl');
    const nblSourceEl = document.getElementById('mp-mill-nbl-source');
    const rrEl = document.getElementById('mp-mill-result-risk');
    if (!nameEl || !locEl || !supEl || !nblEl || !rrEl) return;
    const companyName = String(row['COMPANY NAME'] != null ? row['COMPANY NAME'] : '').trim();
    const groupName = String(row['GROUP NAME'] != null ? row['GROUP NAME'] : '').trim();
    const millName = String(row['MILL NAME'] != null ? row['MILL NAME'] : '').trim();
    nameEl.textContent = companyName
      ? millProfileTitleCaseWords_(companyName)
      : (millName ? millProfileTitleCaseWords_(millName) : '—');
    locEl.textContent = groupName
      ? millProfileTitleCaseWords_(groupName)
      : millProfileFormatHeaderLoc_(row);
    const sup = String(row['SUPPLIER STATUS'] != null ? row['SUPPLIER STATUS'] : '').trim();
    supEl.textContent = sup ? millProfileTitleCaseWords_(sup) : '—';
    nblEl.innerHTML = nblBadge(row['BUYER NO BUY LIST']);
    if (nblSourceEl) nblSourceEl.innerHTML = '';
    rrEl.innerHTML = resultRiskLevelPill(row['RESULT RISK LEVEL']);
    millProfileUpdateNblSourceInfo_(row);
  }

  const MILL_PROFILE_FIELD_ALIASES_ = {
    'MILL LOC': ['MILL LOC', 'MILL LOCATION', 'Mill Location', 'LOC'],
    'DEFORESTATION SPATIAL': ['DEFORESTATION SPATIAL', 'Deforestation Spatial'],
    'BURN AREA SPATIAL': ['BURN AREA SPATIAL', 'Burn Area Spatial'],
    'PEAT': ['PEAT', 'Peat'],
    'NDPE': ['NDPE'],
    'HRDD': ['HRDD'],
  };

  const MILL_PROFILE_YESNO_KEYS_ = new Set([
    'DEFORESTATION SPATIAL', 'BURN AREA SPATIAL', 'PEAT', 'NDPE', 'HRDD',
  ]);

  /** Sheet column LEGALITY SCORE (or SCORE): 1 → Complete, 0 → Not Complete. */
  function millProfileLegalityFromScore_(d) {
    const raw = pickSavedCol(d, ['LEGALITY SCORE', 'SCORE', 'Legality Score']);
    if (raw === '' || raw === null || raw === undefined) return '';
    if (typeof raw === 'number' && !isNaN(raw)) {
      if (raw === 1) return 'Complete';
      if (raw === 0) return 'Not Complete';
    }
    const s = String(raw).trim();
    const t = s.toLowerCase();
    if (t === '1' || t === 'complete') return 'Complete';
    if (t === '0' || t === 'not complete' || t === 'non complete') return 'Not Complete';
    return s;
  }

  function millProfileFormatYesNo_(raw) {
    const s = String(raw == null ? '' : raw).trim();
    if (!s) return '';
    const t = s.toLowerCase();
    if (t === 'yes' || t === 'y' || t === '1' || t === 'true' || t === 'ada') return 'Yes';
    if (t === 'no' || t === 'n' || t === '0' || t === 'false' || t === 'tidak') return 'No';
    return s;
  }

  function millProfileResolveField_(d, key, opts) {
    opts = opts || {};
    const aliases = MILL_PROFILE_FIELD_ALIASES_[key] || [key];
    const v = pickSavedCol(d, aliases);
    if (!v) return '';
    if (opts.raw) return v;
    if (opts.yesNo || MILL_PROFILE_YESNO_KEYS_.has(key)) return millProfileFormatYesNo_(v) || v;
    return v;
  }

  function millProfileBuildSectionsHtml_(d) {
    function millProfileFormatSupplyValue_(raw) {
      if (raw == null) return '';
      if (typeof raw === 'string') return raw.trim();
      if (typeof raw === 'number') {
        if (!isFinite(raw)) return '';
        if (Number.isInteger(raw) && Math.abs(raw) >= 1000) {
          return String(raw).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
        }
        return String(raw);
      }
      const s = String(raw).trim();
      if (!s) return '';
      if (/^[+-]?\d+$/.test(s) && s.length > 3) {
        const neg = s[0] === '-';
        const digits = neg ? s.slice(1) : s;
        const grouped = digits.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
        return neg ? '-' + grouped : grouped;
      }
      return s;
    }

    const sections = [
      {
        title: 'Mill Identity',
        fields: [
          ['COMPANY CODE','Company Code'], ['MILL NAME','Mill Name'], ['TRADER NAME','Trader Name'],
          ['ADDRESS','Address'], ['PROVINCE','Province'], ['COORDINATES','Coordinates'],
          ['MILL CATEGORY','Mill Category'], ['UML ID','UML ID'], ['MILL CAPACITY (TON/HOUR)','Capacity (Ton/Hour)'],
        ]
      },
      {
        title: 'Legality',
        fields: [
          ['LEGALITY', 'Legality'],
          ['MILL LOC', 'Mill Location'],
        ],
      },
      {
        title: 'Certification',
        fields: [
          ['CERTIFICATION','Certification'],
        ]
      },
      {
        title: 'Spatial',
        fields: [
          ['DEFORESTATION SPATIAL','Deforestation Spatial'], ['BURN AREA SPATIAL','Burn Area Spatial'],
          ['PEAT','Peat'],
        ]
      },
      {
        title: 'Grievances',
        fields: [
          ['TOTAL GRIEVANCES','Grievance'],
        ]
      },
      {
        title: 'Policy',
        fields: [
          ['NDPE','NDPE'], ['HRDD','HRDD'],
        ]
      },
      {
        title: 'Supplied Data',
        fields: [
          ['PRODUCT SUPPLY','Product Supply'],
          ['SUPPLY CPO','Supply CPO'],
          ['SUPPLY PK','Supply PK'],
          ['FACILITY NAME','Facility Name'],
        ]
      },
    ];
    return sections.map(function(sec) {
      return `
      <div class="mp-section">
        <div class="mp-section-title">${sec.title}</div>
        <div class="mp-grid${sec.fields.length <= 4 ? ' cols2' : ''}${sec.title === 'Legality' ? ' legalitas-stack' : ''}${sec.title === 'Spatial' ? ' spatial-stack' : ''}${sec.title === 'Grievances' ? ' grievances-stack' : ''}${sec.title === 'Policy' ? ' policy-stack' : ''}${sec.title === 'Supplied Data' ? ' supplied-stack' : ''}">
          ${sec.fields.map(function(fl) {
            const key = fl[0];
            const label = fl[1];
            let val = '';
            if (key === 'LEGALITY') {
              val = millProfileLegalityFromScore_(d);
            } else {
              val = millProfileResolveField_(d, key, {
                yesNo: MILL_PROFILE_YESNO_KEYS_.has(key),
              });
            }
            if (key === 'FACILITY NAME') {
              const facCpo = String(d['FACILITY NAME CPO'] || '').trim();
              const facPk = String(d['FACILITY NAME PK'] || '').trim();
              val = [facCpo, facPk].filter(Boolean).join(' / ');
            }
            if (key === 'SUPPLY CPO') {
              val =
                (d['SUPPLY CPO'] != null ? d['SUPPLY CPO'] : '') ||
                (d['CPO SUPPLY to REFINERY'] != null ? d['CPO SUPPLY to REFINERY'] : '') ||
                (d['SCR - CPO Supply'] != null ? d['SCR - CPO Supply'] : '');
              val = millProfileFormatSupplyValue_(val);
            }
            if (key === 'SUPPLY PK') {
              val =
                (d['SUPPLY PK'] != null ? d['SUPPLY PK'] : '') ||
                (d['PK SUPPLY to KCP'] != null ? d['PK SUPPLY to KCP'] : '') ||
                (d['SCR - PK Supply'] != null ? d['SCR - PK Supply'] : '');
              val = millProfileFormatSupplyValue_(val);
            }
            if (key === 'TOTAL POLICY') {
              const polTok = String(val == null ? '' : val).trim().toLowerCase();
              if (polTok === '1' || polTok === 'yes') val = 'Yes';
              else if (polTok === '0' || polTok === 'no') val = 'No';
            }
            if (key === 'TOTAL GRIEVANCES') {
              const grvTok = String(val == null ? '' : val).trim().toLowerCase();
              if (grvTok === '1' || grvTok === 'yes') val = 'Yes';
              else if (grvTok === '0' || grvTok === 'no') val = 'No';
            }
            const isWide = key === 'ADDRESS';
            const isLong = key === 'COORDINATES' || (sec.title === 'Certification' && key === 'CERTIFICATION');
            return `<div class="mp-field${isWide ? ' wide' : ''}${isLong ? ' full' : ''}">
              <div class="mp-label">${label}</div>
              <div class="mp-val">${val || '—'}</div>
            </div>`;
          }).join('')}
        </div>
      </div>`;
    }).join('');
  }

  function millProfileRenderBody_(d) {
    const mpBody = document.getElementById('millProfileBody');
    if (!mpBody) return;
    mpBody.innerHTML = millProfileBuildSectionsHtml_(d);
  }

  function millProfileCurrentRow_() {
    const siblings = millProfileVariantRows_ || [];
    if (!siblings.length) return null;
    const ySel = document.getElementById('millProfileYearSel');
    const qSel = document.getElementById('millProfileQuarterSel');
    if (!ySel || !qSel) return siblings[0];
    const row = millProfileFindRowByPeriodTok_(siblings, ySel.value, qSel.value);
    return row || siblings[0];
  }

  async function millProfileExportPdf() {
    const row = millProfileCurrentRow_();
    const btn = document.getElementById('millProfileExportPdfBtn');
    const toastErr = function(msg) {
      if (typeof window.showSddToast === 'function') window.showSddToast(msg, 'error');
      else alert(msg);
    };
    if (!row) {
      toastErr('Data mill tidak ditemukan untuk diekspor.');
      return;
    }
    const prevTxt = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = 'Exporting...'; }
    try {
      const nblByInfo = await millResolveNblByInfo_(row);
      const JsPDFLib = getJsPDF();
      const doc = new JsPDFLib({ unit: 'mm', format: 'a4' });
      const title = String(row['COMPANY NAME'] || row['MILL NAME'] || 'Mill Profile').trim() || 'Mill Profile';
      const group = String(row['GROUP NAME'] || '').trim();
      const quarter = String(millQuarterVal(row) || '').trim();
      const year = String(millYearVal(row) || '').trim();
      const periodText = (year || quarter) ? ('Period: ' + [year, quarter ? ('Q' + quarter) : ''].filter(Boolean).join(' ')) : '';
      const sections = [
        { title: 'Mill Identity', fields: [['COMPANY CODE','Company Code'], ['MILL NAME','Mill Name'], ['TRADER NAME','Trader Name'], ['ADDRESS','Address'], ['PROVINCE','Province'], ['COORDINATES','Coordinates'], ['MILL CATEGORY','Mill Category'], ['UML ID','UML ID'], ['MILL CAPACITY (TON/HOUR)','Capacity (Ton/Hour)']] },
        { title: 'Legality', fields: [['LEGALITY', 'Legality'], ['MILL LOC', 'Mill Location']] },
        { title: 'Certification', fields: [['CERTIFICATION','Certification']] },
        { title: 'Spatial', fields: [['DEFORESTATION SPATIAL','Deforestation Spatial'], ['BURN AREA SPATIAL','Burn Area Spatial'], ['PEAT','Peat']] },
        { title: 'Grievances', fields: [['TOTAL GRIEVANCES','Grievance']] },
        { title: 'Policy', fields: [['NDPE','NDPE'], ['HRDD','HRDD']] },
        { title: 'Supplied Data', fields: [['PRODUCT SUPPLY','Product Supply'], ['SUPPLY CPO','Supply CPO'], ['SUPPLY PK','Supply PK'], ['FACILITY NAME','Facility Name']] },
      ];
      const valOrDash = function(v) {
        const s = String(v == null ? '' : v).trim();
        return s ? s : '—';
      };
      const mapBool01 = function(v) {
        const s = String(v == null ? '' : v).trim().toLowerCase();
        if (s === '1' || s === 'yes') return 'Yes';
        if (s === '0' || s === 'no') return 'No';
        return valOrDash(v);
      };
      const formatSupplyLikeUi = function(raw) {
        if (raw == null) return '';
        if (typeof raw === 'string') return raw.trim();
        if (typeof raw === 'number') {
          if (!isFinite(raw)) return '';
          if (Number.isInteger(raw) && Math.abs(raw) >= 1000) {
            return String(raw).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
          }
          return String(raw);
        }
        return String(raw).trim();
      };
      const resolveExportVal = function(key) {
        if (key === 'SUPPLY CPO') {
          const v = (row['SUPPLY CPO'] != null ? row['SUPPLY CPO'] : '') ||
                    (row['CPO SUPPLY to REFINERY'] != null ? row['CPO SUPPLY to REFINERY'] : '') ||
                    (row['SCR - CPO Supply'] != null ? row['SCR - CPO Supply'] : '');
          return formatSupplyLikeUi(v);
        }
        if (key === 'SUPPLY PK') {
          const v = (row['SUPPLY PK'] != null ? row['SUPPLY PK'] : '') ||
                    (row['PK SUPPLY to KCP'] != null ? row['PK SUPPLY to KCP'] : '') ||
                    (row['SCR - PK Supply'] != null ? row['SCR - PK Supply'] : '');
          return formatSupplyLikeUi(v);
        }
        if (key === 'FACILITY NAME') {
          const facCpo = String(row['FACILITY NAME CPO'] || '').trim();
          const facPk = String(row['FACILITY NAME PK'] || '').trim();
          return [facCpo, facPk].filter(Boolean).join(' / ');
        }
        if (key === 'SUPPLIER STATUS') {
          return millProfileTitleCaseWords_(row['SUPPLIER STATUS']);
        }
        if (key === 'LEGALITY') {
          return valOrDash(millProfileLegalityFromScore_(row));
        }
        if (MILL_PROFILE_YESNO_KEYS_.has(key)) {
          return valOrDash(millProfileResolveField_(row, key, { yesNo: true }));
        }
        const resolved = millProfileResolveField_(row, key);
        if (resolved) return resolved;
        return row[key];
      };

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(14);
      doc.setTextColor(26, 26, 26);
      doc.text(title, 14, 14);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(10);
      let y = 20;
      if (group) { doc.text(group, 14, y); y += 5; }
      if (periodText) { doc.text(periodText, 14, y); y += 5; }
      if (nblByInfo && nblByInfo.label) { doc.text(nblByInfo.label, 14, y); y += 5; }
      doc.text('Exported: ' + new Date().toLocaleString('id-ID'), 14, y);
      y += 4;

      y = renderMillProfileSummaryPdf(doc, row, y + 4, resolveExportVal);

      sections.forEach(function(sec) {
        const body = sec.fields.map(function(pair) {
          const key = pair[0];
          const label = pair[1];
          let value = resolveExportVal(key);
          if (key === 'TOTAL GRIEVANCES' || key === 'TOTAL POLICY') value = mapBool01(value);
          return [label, valOrDash(value)];
        });
        doc.autoTable({
          head: [[sec.title, 'Value']],
          body: body,
          startY: y + 4,
          margin: { left: 14, right: 14 },
          theme: 'grid',
          styles: { fontSize: 9, cellPadding: 2 },
          headStyles: { fillColor: [139, 26, 26], textColor: [255, 255, 255], fontStyle: 'bold' },
          columnStyles: { 0: { cellWidth: 58 }, 1: { cellWidth: 'auto' } },
        });
        y = (doc.lastAutoTable && doc.lastAutoTable.finalY) ? doc.lastAutoTable.finalY : y + 8;
      });

      const safeName = title.replace(/[^\w\s-]+/g, '').trim().replace(/\s+/g, '-').slice(0, 60) || 'Mill-Profile';
      const fileName = safeName + '-' + new Date().toISOString().slice(0, 10) + '.pdf';
      doc.save(fileName);
      if (typeof window.showSddToast === 'function') window.showSddToast('PDF berhasil diunduh.', 'success');
    } catch (e) {
      toastErr('Export PDF gagal: ' + (e && e.message ? e.message : e));
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = prevTxt || 'Export PDF'; }
    }
  }

  function millProfileSyncBodyToQySelections_(changeSourceId) {
    const siblings = millProfileVariantRows_;
    const ySel = document.getElementById('millProfileYearSel');
    const qSel = document.getElementById('millProfileQuarterSel');
    if (!siblings || !siblings.length || !ySel || !qSel) return;
    if (changeSourceId === 'millProfileYearSel') {
      const yTok = ySel.value;
      const qToks = millProfileCollectQuarterToksForYear_(siblings, yTok);
      millProfileFillSelectToks_(qSel, qToks);
      if (qSel.options.length) qSel.selectedIndex = 0;
    }
    const yTok = ySel.value;
    const qTok = qSel.value;
    const row = millProfileFindRowByPeriodTok_(siblings, yTok, qTok) || siblings[0];
    if (!row) return;
    millProfileUpdateHeaderFromRow_(row);
    millProfileRenderBody_(row);
    millProfileResetScroll_();
  }

  function millProfileResetScroll_() {
    const mpOverlay = document.getElementById('millProfileOverlay');
    if (!mpOverlay) return;
    const mpBox = mpOverlay.querySelector('.mill-profile-box');
    const mpBody = document.getElementById('millProfileBody');
    [mpOverlay, mpBox, mpBody].forEach(function(el) {
      if (el) el.scrollTop = 0;
    });
  }

  function openMillProfile(d) {
    millProfileVariantRows_ = millProfileSameEntityRows_(d).slice().sort(millProfileComparePeriodDesc_);
    const chosen = millProfileVariantRows_[0] || d;
    const mpOverlay = document.getElementById('millProfileOverlay');
    const yBar = document.getElementById('millProfileQyBar');
    const ySel = document.getElementById('millProfileYearSel');
    const qSel = document.getElementById('millProfileQuarterSel');

    if (yBar && ySel && qSel) {
      yBar.hidden = false;
      const yearToks = millProfileCollectYearToks_(millProfileVariantRows_);
      millProfileFillSelectToks_(ySel, yearToks);
      const yTokChosen = millPdfTokenForCell(millYearVal(chosen));
      if (yearToks.indexOf(yTokChosen) !== -1) ySel.value = yTokChosen;
      else if (ySel.options.length) ySel.selectedIndex = 0;
      const yTok = ySel.value;
      const qToks = millProfileCollectQuarterToksForYear_(millProfileVariantRows_, yTok);
      millProfileFillSelectToks_(qSel, qToks);
      const qTokChosen = millPdfTokenForCell(millQuarterVal(chosen));
      if (qToks.indexOf(qTokChosen) !== -1) qSel.value = qTokChosen;
      else if (qSel.options.length) qSel.selectedIndex = 0;
    }

    let yTokFinal = millPdfTokenForCell(millYearVal(chosen));
    let qTokFinal = millPdfTokenForCell(millQuarterVal(chosen));
    if (ySel && qSel) {
      yTokFinal = ySel.value;
      qTokFinal = qSel.value;
    }
    const displayRow = millProfileFindRowByPeriodTok_(millProfileVariantRows_, yTokFinal, qTokFinal) || chosen;
    millProfileUpdateHeaderFromRow_(displayRow);
    millProfileRenderBody_(displayRow);

    mpOverlay?.classList.add('active');
    millProfileResetScroll_();
    requestAnimationFrame(function() { millProfileResetScroll_(); });
  }

  (function bindMillProfileOverlay() {
    const mpc = document.getElementById('millProfileClose');
    const mpo = document.getElementById('millProfileOverlay');
    if (!mpc || !mpo) {
      console.warn('[dashboard] Mill profile overlay nodes missing.');
      return;
    }
    function closeMillProfileOverlay_() {
      mpo.classList.remove('active');
      millProfileResetScroll_();
    }
    mpc.addEventListener('click', closeMillProfileOverlay_);
    mpo.addEventListener('click', function(e) {
      if (e.target === this) closeMillProfileOverlay_();
    });
    mpo.addEventListener('change', function(e) {
      const t = e.target;
      if (!t || (t.id !== 'millProfileYearSel' && t.id !== 'millProfileQuarterSel')) return;
      e.stopPropagation();
      millProfileSyncBodyToQySelections_(t.id);
    });
    mpo.addEventListener('click', function(e) {
      const btn = e.target && e.target.closest && e.target.closest('#millProfileExportPdfBtn');
      if (!btn) return;
      e.preventDefault();
      e.stopPropagation();
      millProfileExportPdf();
    });
  })();

  // ─── TTP DATA ───────────────────────────────────────────
  let ttpViewMode = 'flat'; // default: flat table (not grouped)
  let ttpFlatFilteredRows = [];
  let ttpDetailCurrentRow = null;
  function normalizeTtpHeaderKey_(h) {
    return String(h || '').trim().replace(/\s+/g, ' ').toUpperCase();
  }

  function pickTtpCategoryCol_(fields) {
    if (!fields || !fields.length) return 'CATEGORY';
    const exact = fields.find(function(h) { return normalizeTtpHeaderKey_(h) === 'CATEGORY'; });
    if (exact) return exact;
    const fuzzy = fields.find(function(h) {
      const u = normalizeTtpHeaderKey_(h);
      return u.includes('CATEGORY') && u !== 'SUPPLIER CATEGORY' && !u.includes('MILL CATEGORY');
    });
    return fuzzy || 'CATEGORY';
  }

  function ttpRowCategoryValue_(row) {
    if (!row || typeof row !== 'object') return '—';
    const key = ttpCategoryCol || 'CATEGORY';
    if (row[key] != null && String(row[key]).trim()) return ttpCategoryDisplay_(row[key]);
    const alt = Object.keys(row).find(function(k) {
      if (k === '_row' || k === '_sddSearchBlob') return false;
      return normalizeTtpHeaderKey_(k) === 'CATEGORY';
    });
    return alt ? ttpCategoryDisplay_(row[alt]) : '—';
  }

  function ttpCategoryDisplay_(raw) {
    const s = String(raw == null ? '' : raw).trim();
    return s || '—';
  }

  function ttpCategoryGroupSummary_(rows) {
    if (!rows || !rows.length) return '—';
    const vals = rows.map(function(r) { return ttpRowCategoryValue_(r); }).filter(function(v) { return v && v !== '—'; });
    if (!vals.length) return '—';
    const uniq = [];
    vals.forEach(function(v) { if (uniq.indexOf(v) === -1) uniq.push(v); });
    if (uniq.length === 1) return uniq[0];
    if (uniq.length <= 3) return uniq.join(', ');
    return uniq.slice(0, 2).join(', ') + ' +' + (uniq.length - 2);
  }

  const TTP_CATEGORY_MIX_BUCKETS = [
    { id: 'own_estate',      label: 'Own Estate',      color: '#2d5a3d' },
    { id: 'external_estate', label: 'External Estate', color: '#1a6b5c' },
    { id: 'dealer',          label: 'Dealer',          color: '#9a6b1a' },
    { id: 'plasma',          label: 'Plasma',          color: '#2a5f8f' },
    { id: 'cooperative',     label: 'Cooperative',     color: '#6b3f8f' },
  ];

  function ttpNormalizeCategoryBucket_(raw) {
    const s = String(raw == null ? '' : raw).trim().toLowerCase().replace(/\s+/g, ' ');
    if (!s || s === '—' || s === '-') return '';
    if (/^own\s*estate/.test(s) || s === 'own estate') return 'own_estate';
    if (/^external\s*estate/.test(s) || s === 'external estate') return 'external_estate';
    if (s === 'dealer' || /^dealer\b/.test(s)) return 'dealer';
    if (/plasma/.test(s)) return 'plasma';
    if (/cooperat|co-op|kooperasi|\bcoop\b/.test(s)) return 'cooperative';
    return 'other';
  }

  /**
   * Largest-remainder on tenths (×10) so displayed % labels always sum to 100.0%.
   * Bar/track widths use the same adjusted values (sum = 100).
   */
  function ttpApplyMixPctLabels_(items, basis) {
    if (!basis) {
      items.forEach(function(it) {
        it.pct = 0;
        it.pctLabel = '0%';
      });
      return items;
    }

    const work = items.map(function(it) {
      const exactTenths = (it.count / basis) * 1000;
      const floorTenths = Math.floor(exactTenths);
      return {
        it: it,
        floorTenths: floorTenths,
        remainder: exactTenths - floorTenths,
      };
    });

    let sumTenths = work.reduce(function(s, w) { return s + w.floorTenths; }, 0);
    let need = 1000 - sumTenths;
    work.sort(function(a, b) { return b.remainder - a.remainder; });
    for (let i = 0; i < need && i < work.length; i++) {
      work[i].floorTenths++;
    }

    work.forEach(function(w) {
      const tenths = w.floorTenths;
      const it = w.it;
      it.pct = tenths / 10;
      if (it.count > 0 && tenths === 0) {
        it.pctLabel = '<0.1%';
        it.pct = 0.05;
      } else if (tenths % 10 === 0) {
        it.pctLabel = String(tenths / 10) + '%';
      } else {
        it.pctLabel = (tenths / 10).toFixed(1) + '%';
      }
    });

    return items;
  }

  function ttpComputeCategoryMix_(rows) {
    const counts = Object.create(null);
    TTP_CATEGORY_MIX_BUCKETS.forEach(function(b) { counts[b.id] = 0; });

    (rows || []).forEach(function(row) {
      const raw = ttpRowCategoryValue_(row);
      const bucket = ttpNormalizeCategoryBucket_(raw);
      if (!bucket || bucket === 'other') return;
      if (counts[bucket] !== undefined) counts[bucket]++;
    });

    const basis = TTP_CATEGORY_MIX_BUCKETS.reduce(function(sum, b) {
      return sum + (counts[b.id] || 0);
    }, 0);

    const items = TTP_CATEGORY_MIX_BUCKETS.map(function(b) {
      return {
        id: b.id,
        label: b.label,
        color: b.color,
        count: counts[b.id] || 0,
        pct: 0,
        pctLabel: '0%',
      };
    });

    ttpApplyMixPctLabels_(items, basis);
    return { items: items, basis: basis };
  }

  function ttpQuarterToken_(row) {
    const q = millQuarterVal(row);
    if (!q) return '';
    const n = parseMillQuarterSort(q);
    if (n >= 1 && n <= 4) return 'Q' + n;
    return String(q).trim().toUpperCase();
  }

  function ttpYearToken_(row) {
    const y = millYearVal(row);
    if (!y) return '';
    const n = parseMillYearSort(y);
    return n ? String(n) : String(y).trim();
  }

  function ttpFilterByPeriod_(rows) {
    const list = rows || [];
    if (ttpPeriodMode === 'quarter') {
      const wantY = String(ttpPeriodYear || '');
      const wantQ = String(ttpPeriodQuarter || '');
      return list.filter(function(r) {
        return ttpYearToken_(r) === wantY && ttpQuarterToken_(r) === wantQ;
      });
    }
    if (ttpPeriodMode === 'overall') {
      const wantY = String(ttpPeriodYear || '');
      if (!wantY) return list.slice();
      return list.filter(function(r) { return ttpYearToken_(r) === wantY; });
    }
    return list.slice();
  }

  function ttpGetPeriodRows_() {
    return ttpFilterByPeriod_(ttpData);
  }

  function ttpApplyTableFilters_(rows) {
    let filtered = ttpFilterByPeriod_(rows || []);
    filtered = filtered.filter(function(d) {
      return !ttpSearch || (d._sddSearchBlob || '').includes(ttpSearch);
    });
    if (ttpSelectedCompanies !== null) {
      const filterCol = ttpActiveFilterCol || ttpCompanyCol;
      if (filterCol) {
        filtered = filtered.filter(function(d) {
          return ttpSelectedCompanies.has((d[filterCol] || '').toString());
        });
      }
    }
    return filtered;
  }

  function ttpParsePctValue_(raw) {
    if (raw === null || raw === undefined) return NaN;
    let s = String(raw).trim();
    if (!s || s === '—' || s === '-') return NaN;
    s = s.replace(/%/g, '').trim().replace(/\s/g, '');
    if (/,\d/.test(s) && !/\.\d{1,}/.test(s.replace(/,\d+$/, ''))) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
    const n = parseFloat(s);
    return isNaN(n) ? NaN : n;
  }

  /** Sheets getValues() returns % cells as 0–1 fractions (0.7081 = 70.81%). */
  function ttpNormalizePctNumber_(n) {
    if (isNaN(n)) return NaN;
    if (n >= 0 && n <= 1) return n * 100;
    return n;
  }

  function ttpParseNumber_(raw) {
    if (raw === null || raw === undefined) return NaN;
    if (typeof raw === 'number' && !isNaN(raw)) return raw;
    let s = String(raw).trim();
    if (!s || s === '—' || s === '-') return NaN;
    s = s.replace(/%/g, '').replace(/\s/g, '');
    const dotCount = (s.match(/\./g) || []).length;
    const commaCount = (s.match(/,/g) || []).length;
    if (commaCount && /,\d{1,3}$/.test(s)) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else if (dotCount > 1) {
      s = s.replace(/\./g, '');
    } else if (dotCount === 1 && /^\d{1,3}(\.\d{3})+$/.test(s)) {
      s = s.replace(/\./g, '');
    } else {
      s = s.replace(/,/g, '');
    }
    const n = parseFloat(s);
    return isNaN(n) ? NaN : n;
  }

  function ttpIsDataRow_(row) {
    if (!row || typeof row !== 'object') return false;
    if (row._sddSearchBlob && /total traceable cpo|total traceable pk/.test(row._sddSearchBlob)) {
      return false;
    }
    let hasMill = false;
    let hasCompany = false;
    let hasSupplier = false;
    const keys = Object.keys(row);
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      if (k === '_row' || (String(k).length && String(k)[0] === '_')) continue;
      const v = String(row[k] || '').trim();
      if (/^total traceable/i.test(v)) return false;
      const u = normalizeTtpHeaderKey_(k);
      if (u === 'MILL NAME' && v && v !== '—' && v !== '-') hasMill = true;
      if ((u === 'COMPANY NAME' || u === 'COMPANY CODE') && v && v !== '—' && v !== '-') hasCompany = true;
      if ((u === 'FFB SUPPLIER NAME' || u === 'FFB SUPPLIER GROUP NAME') && v && v !== '—' && v !== '-') {
        hasSupplier = true;
      }
    }
    return hasMill || hasCompany || hasSupplier;
  }

  /** Sheet footer: PK = SUM(PK Traceable) / SUM(PK SUPPLY to KCP); CPO = SUM(CPO Traceable) / SUM(CPO SUPPLY to REFINERY). */
  var TTP_TRACEABLE_COL_CANDIDATES_ = {
    pk: {
      numerator: ['PK TRACEABLE VOLUME', 'PK TRACEABLE'],
      denominator: ['PK SUPPLY TO KCP', 'PK SUPPLY TO KCP (TON)', 'PK SUPPLY']
    },
    cpo: {
      numerator: ['CPO TRACEABLE', 'CPO TRACEABLE VOLUME'],
      denominator: ['CPO SUPPLY TO REFINERY']
    }
  };

  function ttpPickHeaderCol_(fields, candidates, opts) {
    const list = fields || [];
    const norms = (candidates || []).map(function(c) { return normalizeTtpHeaderKey_(c); });
    let i;
    for (i = 0; i < norms.length; i++) {
      const exact = list.find(function(h) { return normalizeTtpHeaderKey_(h) === norms[i]; });
      if (exact) return exact;
    }
    const excludePct = opts && opts.excludePct;
    const excludeTotal = opts && opts.excludeTotal;
    const tag = opts && opts.tag;
    for (i = 0; i < norms.length; i++) {
      const want = norms[i];
      const hit = list.find(function(h) {
        const u = normalizeTtpHeaderKey_(h);
        if (excludePct && u.includes('%')) return false;
        if (excludeTotal && u.includes('TOTAL')) return false;
        if (tag && !u.includes(tag)) return false;
        if (want.includes('SUPPLY') && u.includes('CONVERSION')) return false;
        if (want.includes('CPO SUPPLY') && !u.includes('REFINERY')) return false;
        if (want.includes('PK SUPPLY') && !u.includes('KCP')) return false;
        if (u.includes('TRACEABLE') && u.includes('%')) return false;
        return u.includes(want) || u === want;
      });
      if (hit) return hit;
    }
    return '';
  }

  function ttpFindTraceableVolCol_(fields, product) {
    const cfg = TTP_TRACEABLE_COL_CANDIDATES_[product];
    if (!cfg) return '';
    return ttpPickHeaderCol_(fields, cfg.numerator, {
      excludePct: true,
      excludeTotal: true,
      tag: product === 'pk' ? 'PK' : 'CPO'
    });
  }

  function ttpFindTraceableDenomCol_(fields, product) {
    const cfg = TTP_TRACEABLE_COL_CANDIDATES_[product];
    if (!cfg) return '';
    return ttpPickHeaderCol_(fields, cfg.denominator, {
      excludePct: true,
      tag: product === 'pk' ? 'PK' : 'CPO'
    });
  }

  function ttpFormatCellPct_(raw) {
    const n = ttpNormalizePctNumber_(ttpParsePctValue_(raw));
    if (isNaN(n)) {
      if (raw === null || raw === undefined || String(raw).trim() === '') return '—';
      return String(raw);
    }
    if (n > 0 && n < 0.05) return '<0.1%';
    if (n >= 10) return (Math.round(n * 10) / 10).toFixed(1).replace(/\.0$/, '') + '%';
    return (Math.round(n * 10) / 10).toFixed(1) + '%';
  }

  function ttpFindSupplyCol_() {
    if (!ttpFields || !ttpFields.length) return '';
    return ttpFields.find(function(h) {
      const u = normalizeTtpHeaderKey_(h);
      return u.includes('FFB SUPPLY') && u.includes('MILL');
    }) || ttpFields.find(function(h) {
      return normalizeTtpHeaderKey_(h).includes('FFB SUPPLY');
    }) || '';
  }

  /** Same as sheet footer: SUM(traceable volume) ÷ SUM(supply volume). */
  function ttpAggregateTotalTraceablePct_(rows, product) {
    const numCol = product === 'pk' ? ttpPkTraceVolCol : ttpCpoTraceVolCol;
    const denCol = product === 'pk' ? ttpPkTraceDenomCol : ttpCpoTraceDenomCol;
    const pctCol = product === 'pk' ? ttpPkPctCol : ttpPctCol;
    const dataRows = (rows || []).filter(ttpIsDataRow_);

    if (!numCol || !denCol) {
      const legacy = ttpAggregateTraceablePctFromCol_(dataRows, pctCol);
      legacy.method = 'average';
      return legacy;
    }

    let sumNum = 0;
    let sumDen = 0;
    let rowsWithNum = 0;
    let rowsWithDen = 0;
    dataRows.forEach(function(row) {
      const n = ttpParseNumber_(row[numCol]);
      const d = ttpParseNumber_(row[denCol]);
      if (!isNaN(n)) {
        sumNum += n;
        rowsWithNum++;
      }
      if (!isNaN(d)) {
        sumDen += d;
        if (d > 0) rowsWithDen++;
      }
    });

    if (!rowsWithNum || sumDen <= 0) {
      return {
        value: NaN,
        rowsUsed: 0,
        totalRows: dataRows.length,
        method: 'sum',
        numCol: numCol,
        denCol: denCol,
        sumNum: sumNum,
        sumDen: sumDen
      };
    }

    const ratio = sumNum / sumDen;
    const value = ratio <= 1.5 ? ratio * 100 : ratio;

    return {
      value: value,
      rowsUsed: dataRows.length,
      rowsWithNum: rowsWithNum,
      rowsWithDen: rowsWithDen,
      totalRows: dataRows.length,
      method: 'sum',
      numCol: numCol,
      denCol: denCol,
      sumNum: sumNum,
      sumDen: sumDen
    };
  }

  function ttpAggregateTraceablePctFromCol_(rows, pctCol) {
    if (!pctCol) return { value: NaN, rowsUsed: 0, totalRows: (rows || []).length };
    let sumSimple = 0;
    let rowsUsed = 0;
    (rows || []).forEach(function(row) {
      const pct = ttpNormalizePctNumber_(ttpParsePctValue_(row[pctCol]));
      if (isNaN(pct)) return;
      rowsUsed++;
      sumSimple += pct;
    });
    if (!rowsUsed) return { value: NaN, rowsUsed: 0, totalRows: (rows || []).length };
    return { value: sumSimple / rowsUsed, rowsUsed: rowsUsed, totalRows: (rows || []).length };
  }

  function ttpFormatTraceablePct_(n) {
    if (isNaN(n)) return '—';
    if (n > 0 && n < 0.05) return '<0.1%';
    if (n >= 10) return (Math.round(n * 10) / 10).toFixed(1).replace(/\.0$/, '') + '%';
    return (Math.round(n * 10) / 10).toFixed(1) + '%';
  }

  function ttpPeriodScopeLabel_() {
    const y = ttpPeriodYear || '—';
    if (ttpPeriodMode === 'overall') return 'Full year ' + y;
    return y + ' · ' + (ttpPeriodQuarter || '—');
  }

  function ttpFormatTtpTon_(n) {
    if (n === null || n === undefined || isNaN(n)) return '—';
    return Math.round(n).toLocaleString('en-US');
  }

  function ttpPeriodMetaText_(periodRows, cpoAgg, pkAgg) {
    const dataRows = (periodRows || []).filter(ttpIsDataRow_);
    const n = dataRows.length;
    const total = ttpData.length;
    const scope = ttpPeriodScopeLabel_();
    if (!total) return 'No data loaded';
    let line = scope + ' · ' + n.toLocaleString() + ' supplier rows';
    if (ttpPeriodMode === 'quarter') {
      line += ' (filtered)';
    } else if (n < total) {
      line += ' · ' + total.toLocaleString() + ' total in database';
    }
    if (cpoAgg && cpoAgg.method === 'sum' && cpoAgg.sumDen > 0) {
      line += ' · CPO: ' + ttpFormatTtpTon_(cpoAgg.sumNum) + ' / ' + ttpFormatTtpTon_(cpoAgg.sumDen) + ' ton';
    }
    return line;
  }

  function buildTtpPeriodDropdowns_() {
    const yearSel = document.getElementById('ttpPeriodYear');
    if (!yearSel) return;
    const years = [];
    const seenY = Object.create(null);
    ttpData.forEach(function(row) {
      const y = ttpYearToken_(row);
      if (y && !seenY[y]) { seenY[y] = true; years.push(y); }
    });
    years.sort(function(a, b) { return parseInt(b, 10) - parseInt(a, 10); });
    yearSel.innerHTML = years.length
      ? years.map(function(y) { return '<option value="' + escHtml(y) + '">' + escHtml(y) + '</option>'; }).join('')
      : '<option value="">—</option>';
    if (years.length && (!ttpPeriodYear || years.indexOf(ttpPeriodYear) === -1)) {
      ttpPeriodYear = years[0];
    }
    yearSel.value = ttpPeriodYear || '';
  }

  function syncTtpPeriodPickersUi_() {
    const pickers = document.getElementById('ttpPeriodPickers');
    const yearWrap = document.getElementById('ttpPeriodYearWrap');
    const quarterWrap = document.getElementById('ttpPeriodQuarterWrap');
    const modeBtns = document.querySelectorAll('[data-ttp-period-mode]');
    modeBtns.forEach(function(btn) {
      btn.classList.toggle('active', btn.getAttribute('data-ttp-period-mode') === ttpPeriodMode);
    });
    const isQuarter = ttpPeriodMode === 'quarter';
    if (pickers) pickers.hidden = false;
    if (yearWrap) yearWrap.style.display = '';
    if (quarterWrap) quarterWrap.style.display = isQuarter ? '' : 'none';
    const qSel = document.getElementById('ttpPeriodQuarter');
    if (qSel && ttpPeriodQuarter) qSel.value = ttpPeriodQuarter;
    const ySel = document.getElementById('ttpPeriodYear');
    if (ySel && ttpPeriodYear) ySel.value = ttpPeriodYear;
  }

  function renderTtpTraceableStats_() {
    const cpoEl = document.getElementById('ttp-stat-cpo-traceable');
    const pkEl = document.getElementById('ttp-stat-pk-traceable');
    const metaEl = document.getElementById('ttpPeriodMeta');
    const periodRows = ttpGetPeriodRows_();
    const cpoAgg = ttpAggregateTotalTraceablePct_(periodRows, 'cpo');
    const pkAgg = ttpAggregateTotalTraceablePct_(periodRows, 'pk');
    if (metaEl) metaEl.textContent = ttpPeriodMetaText_(periodRows, cpoAgg, pkAgg);

    if (cpoEl) {
      cpoEl.textContent = ttpFormatTraceablePct_(cpoAgg.value);
      cpoEl.title = cpoAgg.method === 'sum' && cpoAgg.sumDen > 0
        ? 'SUM(' + cpoAgg.numCol + ') ÷ SUM(' + cpoAgg.denCol + ')'
          + '\n' + ttpFormatTtpTon_(cpoAgg.sumNum) + ' ÷ ' + ttpFormatTtpTon_(cpoAgg.sumDen)
          + ' ton = ' + ttpFormatTraceablePct_(cpoAgg.value)
          + '\nBandingkan dengan total kolom CPO Traceable & CPO SUPPLY to REFINERY di sheet (periode sama).'
        : 'Tidak ada data CPO traceable pada periode ini';
    }
    if (pkEl) {
      pkEl.textContent = ttpFormatTraceablePct_(pkAgg.value);
      pkEl.title = pkAgg.method === 'sum' && pkAgg.sumDen > 0
        ? 'SUM(' + pkAgg.numCol + ') ÷ SUM(' + pkAgg.denCol + ')'
          + '\n' + ttpFormatTtpTon_(pkAgg.sumNum) + ' ÷ ' + ttpFormatTtpTon_(pkAgg.sumDen)
          + ' ton = ' + ttpFormatTraceablePct_(pkAgg.value)
        : 'Tidak ada data PK traceable pada periode ini';
    }
  }

  function refreshTtpPeriodDashboard_() {
    syncTtpPeriodPickersUi_();
    renderTtpTraceableStats_();
    renderTtpCategoryMix_();
    scheduleRenderTTPTable();
  }

  let ttpPeriodBarBound = false;
  function bindTtpPeriodBarOnce_() {
    if (ttpPeriodBarBound) return;
    const panel = document.getElementById('panel-ttm-ttp');
    if (!panel) return;
    ttpPeriodBarBound = true;
    panel.addEventListener('click', function(e) {
      const btn = e.target.closest('[data-ttp-period-mode]');
      if (!btn || !panel.contains(btn)) return;
      const mode = btn.getAttribute('data-ttp-period-mode');
      if (!mode || mode === ttpPeriodMode) return;
      ttpPeriodMode = mode;
      refreshTtpPeriodDashboard_();
    });
    const yearSel = document.getElementById('ttpPeriodYear');
    const qSel = document.getElementById('ttpPeriodQuarter');
    if (yearSel) {
      yearSel.addEventListener('change', function() {
        ttpPeriodYear = this.value;
        refreshTtpPeriodDashboard_();
      });
    }
    if (qSel) {
      qSel.addEventListener('change', function() {
        ttpPeriodQuarter = this.value;
        refreshTtpPeriodDashboard_();
      });
    }
    syncTtpPeriodPickersUi_();
  }

  function renderTtpCategoryMix_() {
    const section = document.getElementById('ttpCategoryMixSection');
    const barEl = document.getElementById('ttpCategoryMixBar');
    const gridEl = document.getElementById('ttpCategoryMixGrid');
    const descEl = document.querySelector('.ttp-category-mix-desc');
    if (!section || !barEl || !gridEl) return;

    if (!ttpData.length) {
      section.hidden = true;
      return;
    }

    const periodRows = ttpGetPeriodRows_();
    if (descEl) {
      descEl.textContent = 'Share of five supplier categories for ' + ttpPeriodScopeLabel_().toLowerCase()
        + ' (' + periodRows.length.toLocaleString() + ' records)';
    }

    const mix = ttpComputeCategoryMix_(periodRows);
    section.hidden = false;

    const barParts = mix.items.filter(function(it) { return it.pct > 0; });
    barEl.innerHTML = barParts.map(function(it) {
      return '<span class="ttp-category-mix-seg" style="flex:' + it.pct + ' 1 0;background:' + it.color + ';" title="'
        + escHtml(it.label) + ': ' + escHtml(it.pctLabel) + '"></span>';
    }).join('');

    gridEl.innerHTML = mix.items.map(function(it) {
      return ''
        + '<div class="ttp-category-mix-item" data-cat="' + escHtml(it.id) + '">'
        + '<div class="ttp-category-mix-pct" style="color:' + escHtml(it.color) + '">' + escHtml(it.pctLabel) + '</div>'
        + '<div class="ttp-category-mix-label">' + escHtml(it.label) + '</div>'
        + '<div class="ttp-category-mix-track" aria-hidden="true"><span class="ttp-category-mix-fill" style="width:'
        + Math.min(100, Math.max(0, it.pct)) + '%;background:' + escHtml(it.color) + ';"></span></div>'
        + '</div>';
    }).join('');

    section.setAttribute('aria-label', 'Supplier category distribution across five categories');
  }

  function ttpMainTableColspan_(grouped) {
    return grouped ? 7 : 6;
  }

  function ttpGroupAvgPctHtml_(rows, product, subCount) {
    const col = product === 'pk' ? ttpPkPctCol : ttpPctCol;
    if (!col) return '—';
    const agg = ttpAggregateTotalTraceablePct_(rows, product);
    const val = isNaN(agg.value) ? '—' : ttpFormatTraceablePct_(agg.value);
    const avgSuffix = subCount > 1 ? '<span class="ttp-group-meta"> avg</span>' : '';
    return '<span style="font-weight:600;color:var(--forest)">' + val + '</span>' + avgSuffix;
  }

  /** Nested supplier grid when a mill group row is expanded (grouped view only). */
  function ttpGroupedNestedTableHtml_(rows) {
    const cols = [
      { label: 'FFB Supplier Group Name', get: function(d) { return ttpRowField_(d, ['FFB SUPPLIER GROUP NAME']); } },
      { label: 'FFB Supplier Name', get: function(d) { return ttpRowField_(d, ['FFB SUPPLIER NAME']); } },
      { label: 'Category', get: function(d) { return ttpRowCategoryValue_(d); } },
      { label: 'Lat', get: function(d) { return ttpRowField_(d, ['LAT', 'LATITUDE']); } },
      { label: 'Long', get: function(d) { return ttpRowField_(d, ['LONG', 'LONGITUDE']); } },
      { label: 'Village', get: function(d) { return ttpRowField_(d, ['VILLAGE']); } },
      { label: 'Sub District', get: function(d) { return ttpRowField_(d, ['SUBDISTRICT', 'SUB DISTRICT']); } },
    ];
    const gridCols = 'minmax(120px, 1.4fr) minmax(120px, 1.4fr) minmax(88px, 0.9fr) minmax(72px, 0.7fr) minmax(72px, 0.7fr) minmax(88px, 0.9fr) minmax(88px, 0.9fr)';
    let html = '<div class="ttp-group-nested-wrap" role="region" aria-label="FFB suppliers for this mill">'
      + '<div class="ttp-group-nested-grid">';
    html += '<div class="ttp-group-nested-head" style="grid-template-columns:' + gridCols + '">';
    cols.forEach(function(c) {
      html += '<div class="ttp-group-nested-cell ttp-group-nested-th">' + escHtml(c.label) + '</div>';
    });
    html += '</div>';
    rows.forEach(function(d) {
      const rowJson = JSON.stringify(d).replace(/'/g, '&#39;');
      html += '<div class="ttp-group-nested-row" style="grid-template-columns:' + gridCols + '" tabindex="0" role="button" data-row=\'' + rowJson
        + '\' title="Click for full supplier detail">';
      cols.forEach(function(c) {
        html += '<div class="ttp-group-nested-cell">' + escHtml(c.get(d)) + '</div>';
      });
      html += '</div>';
    });
    html += '<p class="ttp-group-nested-hint">Click a row to open supplier detail.</p></div></div>';
    return html;
  }

  function ttpMainTableHeadHtml_(grouped) {
    const cpoLabel = ttpPctCol || '% CPO TRACEABLE';
    const pkLabel = ttpPkPctCol || '% PK TRACEABLE';
    const isGrouped = !!grouped;
    let row = '<tr>'
      + '<th class="ttp-th ttp-th-group">Group Name</th>'
      + '<th class="ttp-th ttp-th-company">Company Name</th>'
      + '<th class="ttp-th ttp-th-mill">Mill Name</th>'
      + '<th class="ttp-th ttp-th-pct ttp-th-pct-cpo">' + cpoLabel + '</th>'
      + '<th class="ttp-th ttp-th-pct ttp-th-pct-pk">' + pkLabel + '</th>'
      + '<th class="ttp-th ttp-th-category">Category</th>';
    if (isGrouped) row += '<th class="ttp-th ttp-th-actions"></th>';
    return row + '</tr>';
  }

  function ttpSyncTableLayoutClass_() {
    const table = document.getElementById('ttpTable');
    if (!table) return;
    table.classList.toggle('ttp-table-grouped', ttpViewMode === 'grouped');
  }

  function ttpPickField_(patterns) {
    if (!ttpFields || !ttpFields.length) return '';
    for (let pi = 0; pi < patterns.length; pi++) {
      const p = patterns[pi];
      const want = typeof p === 'string' ? normalizeTtpHeaderKey_(p) : null;
      for (let i = 0; i < ttpFields.length; i++) {
        const h = ttpFields[i];
        const u = normalizeTtpHeaderKey_(h);
        if (want && u === want) return h;
        if (p instanceof RegExp && p.test(h)) return h;
      }
    }
    return '';
  }

  function ttpRowField_(row, patterns) {
    if (!row) return '—';
    const col = ttpPickField_(patterns);
    if (col) {
      const v = row[col];
      if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
    }
    for (const k of Object.keys(row)) {
      if (k === '_row' || k === '_sddSearchBlob') continue;
      const u = normalizeTtpHeaderKey_(k);
      for (let pi = 0; pi < patterns.length; pi++) {
        const p = patterns[pi];
        if (typeof p === 'string' && u === normalizeTtpHeaderKey_(p)) {
          const s = String(row[k]).trim();
          if (s) return s;
        }
      }
    }
    return '—';
  }

  function ttpRowCertificationHtml_(row) {
    const items = [
      { label: 'ISPO', patterns: ['ISPO (Y/N)', 'ISPO'] },
      { label: 'RSPO', patterns: ['RSPO (Y/N)', 'RSPO'] },
      { label: 'ISCC', patterns: ['ISCC (Y/N)', 'ISCC'] },
    ];
    return items.map(function(it) {
      const v = ttpRowField_(row, it.patterns);
      const yn = v === '—' ? '—' : v;
      let cls = 'na';
      if (/^y(es)?$/i.test(yn) || yn === '1') cls = 'yes';
      else if (/^n(o)?$/i.test(yn) || yn === '0') cls = 'no';
      return '<span class="ttp-cert-badge ttp-cert-' + cls + '">' + escHtml(it.label) + ': ' + escHtml(yn) + '</span>';
    }).join('');
  }

  function ttpDetailItemHtml_(label, value) {
    const v = value === undefined || value === null || String(value).trim() === '' ? '—' : String(value).trim();
    return '<div class="ttp-detail-item">'
      + '<div class="ttp-detail-label">' + escHtml(label) + '</div>'
      + '<div class="ttp-detail-val">' + escHtml(v) + '</div>'
      + '</div>';
  }

  function buildTtpDetailBodyHtml_(row) {
    const certHtml = '<div class="ttp-detail-item ttp-detail-item--wide">'
      + '<div class="ttp-detail-label">Certification</div>'
      + '<div class="ttp-detail-val"><div class="ttp-cert-badges">' + ttpRowCertificationHtml_(row) + '</div></div>'
      + '</div>';
    return ''
      + '<div class="ttp-detail-section"><div class="ttp-detail-section-title">Supplier</div><div class="ttp-detail-grid">'
      + ttpDetailItemHtml_('FFB Supplier Group Name', ttpRowField_(row, ['FFB SUPPLIER GROUP NAME']))
      + ttpDetailItemHtml_('FFB Supplier Name', ttpRowField_(row, ['FFB SUPPLIER NAME']))
      + ttpDetailItemHtml_('Category', ttpRowCategoryValue_(row))
      + '</div></div>'
      + '<div class="ttp-detail-section"><div class="ttp-detail-section-title">Location</div><div class="ttp-detail-grid">'
      + ttpDetailItemHtml_('Latitude', ttpRowField_(row, ['LAT', 'LATITUDE']))
      + ttpDetailItemHtml_('Longitude', ttpRowField_(row, ['LONG', 'LONGITUDE']))
      + ttpDetailItemHtml_('Village', ttpRowField_(row, ['VILLAGE']))
      + ttpDetailItemHtml_('Sub district', ttpRowField_(row, ['SUBDISTRICT', 'SUB DISTRICT']))
      + ttpDetailItemHtml_('District', ttpRowField_(row, ['DISTRICT']))
      + ttpDetailItemHtml_('Province', ttpRowField_(row, ['PROVINCE']))
      + '</div></div>'
      + '<div class="ttp-detail-section"><div class="ttp-detail-section-title">Supply</div><div class="ttp-detail-grid">'
      + ttpDetailItemHtml_('Planted Area', ttpRowField_(row, ['PLANTED AREA', 'PLANTED AREA (HA)']))
      + ttpDetailItemHtml_('FFB Supply', ttpRowField_(row, ['FFB SUPPLY TO MILL (TON)', 'FFB SUPPLY']))
      + certHtml
      + '</div></div>';
  }

  function mountTtpDetailOverlayOnce_() {
    const overlay = document.getElementById('ttpDetailOverlay');
    if (!overlay) {
      console.warn('[TTP] #ttpDetailOverlay tidak ditemukan — cek partials/modals-shared.html di index.html.');
      return null;
    }
    if (overlay.parentElement !== document.body) {
      document.body.appendChild(overlay);
    }
    return overlay;
  }

  function syncTtpViewUi_() {
    const flatBtn = document.getElementById('ttpViewFlat');
    const groupedBtn = document.getElementById('ttpViewGrouped');
    if (!flatBtn || !groupedBtn) return;
    flatBtn.classList.toggle('active', ttpViewMode === 'flat');
    groupedBtn.classList.toggle('active', ttpViewMode === 'grouped');
  }

  function closeTtpDetailModal_() {
    const overlay = mountTtpDetailOverlayOnce_();
    if (overlay) overlay.classList.remove('active');
    document.body.classList.remove('ttp-detail-open');
    ttpDetailCurrentRow = null;
    const body = document.getElementById('ttpDetailBody');
    if (body) body.innerHTML = '';
  }

  function openTtpDetailModal_(row) {
    if (!row) return;
    const overlay = mountTtpDetailOverlayOnce_();
    const titleEl = document.getElementById('ttpDetailTitle');
    const subEl = document.getElementById('ttpDetailSubtitle');
    const bodyEl = document.getElementById('ttpDetailBody');
    if (!overlay || !titleEl || !subEl || !bodyEl) return;

    ttpDetailCurrentRow = row;

    const supplier = ttpRowField_(row, ['FFB SUPPLIER NAME']);
    const mill = ttpRowField_(row, ['MILL NAME']);
    const company = ttpRowField_(row, ['COMPANY NAME']);
    const group = ttpRowField_(row, ['GROUP NAME']);

    titleEl.textContent = supplier !== '—' ? supplier : 'Supplier detail';
    const subParts = [group, company, mill].filter(function(p) { return p && p !== '—'; });
    subEl.textContent = subParts.length ? subParts.join(' · ') : 'Monitoring TTM/TTP';

    try {
      bodyEl.innerHTML = buildTtpDetailBodyHtml_(row);
    } catch (err) {
      console.error('[TTP] Gagal render detail supplier', err, row);
      bodyEl.innerHTML = '<p class="ttp-detail-error">Gagal memuat detail. Muat ulang halaman (Ctrl+Shift+R).</p>';
    }

    document.body.classList.add('ttp-detail-open');
    overlay.classList.add('active');
    bodyEl.scrollTop = 0;
    overlay.scrollTop = 0;
  }

  (function bindTtpDetailOverlay() {
    const overlay = mountTtpDetailOverlayOnce_();
    if (!overlay) return;
    const closeBtn = document.getElementById('ttpDetailClose');
    const closeBtn2 = document.getElementById('ttpDetailCloseBtn');
    const editBtn = document.getElementById('ttpDetailEditBtn');
    const delBtn = document.getElementById('ttpDetailDeleteBtn');

    function onOverlayClick_(e) {
      if (e.target === overlay) closeTtpDetailModal_();
    }
    if (closeBtn) closeBtn.addEventListener('click', closeTtpDetailModal_);
    if (closeBtn2) closeBtn2.addEventListener('click', closeTtpDetailModal_);
    overlay.addEventListener('click', onOverlayClick_);
    if (editBtn) {
      editBtn.addEventListener('click', function() {
        const row = ttpDetailCurrentRow;
        if (!row) return;
        closeTtpDetailModal_();
        openModal('ttp', ttpFields, 'edit', row);
      });
    }
    if (delBtn) {
      delBtn.addEventListener('click', function() {
        const row = ttpDetailCurrentRow;
        if (!row || !row._row) return;
        closeTtpDetailModal_();
        openConfirm('ttp', parseInt(row._row, 10));
      });
    }
    if (!window.__sddTtpDetailEscBound) {
      window.__sddTtpDetailEscBound = true;
      document.addEventListener('keydown', function(e) {
        if (e.key !== 'Escape') return;
        if (overlay.classList.contains('active')) closeTtpDetailModal_();
      });
    }
  })();

  async function loadTTPDataImpl() {
    const loading = document.getElementById('ttp-loading');
    const errorEl = document.getElementById('ttp-error');
    const table = document.getElementById('ttpTable');
    if (!loading || !errorEl || !table) {
      console.warn('[dashboard] TTP panel DOM missing; skip loadTTPData.');
      return;
    }
    try {
      loading.style.display = 'block';
      errorEl.style.display = 'none';
      table.style.display = 'none';
      ttpData = await apiGet('ttp');
      ttpLoaded = true;
      ttpFields = ttpData.length > 0 ? Object.keys(ttpData[0]).filter(k => k !== '_row') : [];
      ttpData = ttpData.map(function(row) {
        return prepareTtpRowPerfCache(row, ttpFields);
      });
      ttpPkTraceVolCol = ttpFindTraceableVolCol_(ttpFields, 'pk');
      ttpCpoTraceVolCol = ttpFindTraceableVolCol_(ttpFields, 'cpo');
      ttpPkTraceDenomCol = ttpFindTraceableDenomCol_(ttpFields, 'pk');
      ttpCpoTraceDenomCol = ttpFindTraceableDenomCol_(ttpFields, 'cpo');
      ttpData = ttpData.filter(ttpIsDataRow_);
      ttpUniqueValuesCache = Object.create(null);
      // Traceability % columns
      ttpPctCol = (function pickTtpPctCol(fields) {
        if (!fields || !fields.length) return '% CPO TRACEABLE';
        const U = function (h) { return String(h || '').toUpperCase(); };
        const cpo = fields.find(function (h) { return U(h).includes('% CPO TRACEABLE'); });
        if (cpo) return cpo;
        const legacy = fields.find(function (h) {
          return U(h).includes('PERCENTAGE TRACEABILITY') ||
            U(h) === 'PERCENTAGE TRACEABILITY' ||
            h.toLowerCase().includes('percentage');
        });
        if (legacy) return legacy;
        return '% CPO TRACEABLE';
      })(ttpFields);
      ttpPkPctCol = (function pickTtpPkPctCol(fields) {
        if (!fields || !fields.length) return '% PK TRACEABLE';
        const U = function (h) { return String(h || '').toUpperCase(); };
        const pk = fields.find(function (h) { return U(h).includes('% PK TRACEABLE'); });
        if (pk) return pk;
        return fields.find(function (h) {
          return normalizeTtpHeaderKey_(h).includes('% PK TRACEABLE');
        }) || '% PK TRACEABLE';
      })(ttpFields);
      ttpCategoryCol = pickTtpCategoryCol_(ttpFields);
      document.getElementById('ttpTableHead').innerHTML = ttpMainTableHeadHtml_(false);
      ttpSyncTableLayoutClass_();
      loading.style.display = 'none';
      table.style.display = 'table';
      // reset selection state on fresh load
      ttpSelectedCompanies = null;
      ttpColFilterMode = 'ttm';
      ttpFieldSectionsCache = null;
      buildCompanyDropdown();
      buildColumnModePanel();
      document.getElementById('btn-export-ttp-xlsx').disabled = false;
      ttpViewMode = 'flat';
      syncTtpViewUi_();
      buildTtpPeriodDropdowns_();
      bindTtpPeriodBarOnce_();
      refreshTtpPeriodDashboard_();
    } catch(err) {
      loading.style.display = 'none';
      errorEl.style.display = 'block';
      errorEl.textContent = 'Gagal memuat data: ' + err.message;
      const mixSection = document.getElementById('ttpCategoryMixSection');
      if (mixSection) mixSection.hidden = true;
    }
  }

  window.__ttpInvalidate = function() {
    ttpLoaded = false;
  };

  async function loadTTPData() {
    if (ttpLoadPromise) return ttpLoadPromise;
    ttpLoadPromise = loadTTPDataImpl();
    try {
      await ttpLoadPromise;
    } finally {
      ttpLoadPromise = null;
    }
  }

  // ─── TTP VIEW MODE (grouped / flat) ────────────────────
  let ttpTableDelegationBound = false;

  function bindTtpTableDelegationOnce() {
    if (ttpTableDelegationBound) return;
    const body = document.getElementById('ttpTableBody');
    if (!body) return;
    ttpTableDelegationBound = true;
    body.addEventListener('click', function(e) {
      const editBtn = e.target.closest('.btn-edit');
      if (editBtn && body.contains(editBtn)) {
        e.stopPropagation();
        openModal('ttp', ttpFields, 'edit', JSON.parse(editBtn.dataset.row.replace(/&#39;/g, "'")));
        return;
      }

      const delBtn = e.target.closest('.btn-delete');
      if (delBtn && body.contains(delBtn)) {
        e.stopPropagation();
        openConfirm(delBtn.dataset.sheet, parseInt(delBtn.dataset.rownum, 10));
        return;
      }

      const flatRow = e.target.closest('.ttp-flat-row');
      if (flatRow && body.contains(flatRow) && ttpViewMode === 'flat') {
        const idx = parseInt(flatRow.dataset.flatIdx, 10);
        const row = ttpFlatFilteredRows[idx];
        if (row) openTtpDetailModal_(row);
        return;
      }

      const nestedRow = e.target.closest('.ttp-group-nested-row');
      if (nestedRow && body.contains(nestedRow) && ttpViewMode === 'grouped') {
        if (nestedRow.dataset.row) {
          try {
            const parsed = JSON.parse(nestedRow.dataset.row.replace(/&#39;/g, "'"));
            openTtpDetailModal_(parsed);
          } catch (err) {
            console.warn('[TTP] Could not parse row for detail popup', err);
          }
        }
        return;
      }

      const groupRow = e.target.closest('.ttp-group-row');
      if (groupRow && body.contains(groupRow) && !e.target.closest('.ttp-group-detail-row')) {
        const groupId = groupRow.dataset.group;
        const expanded = groupRow.dataset.expanded === '1';
        const children = body.querySelectorAll(`[data-parent="${groupId}"]`);
        if (expanded) {
          children.forEach(c => c.classList.add('hidden'));
          groupRow.dataset.expanded = '0';
          groupRow.classList.remove('expanded');
        } else {
          children.forEach(c => c.classList.remove('hidden'));
          groupRow.dataset.expanded = '1';
          groupRow.classList.add('expanded');
        }
      }
    });
    body.addEventListener('keydown', function(e) {
      if (ttpViewMode !== 'grouped') return;
      if (e.key !== 'Enter' && e.key !== ' ') return;
      const nestedRow = e.target.closest('.ttp-group-nested-row');
      if (!nestedRow || !body.contains(nestedRow) || !nestedRow.dataset.row) return;
      e.preventDefault();
      try {
        const parsed = JSON.parse(nestedRow.dataset.row.replace(/&#39;/g, "'"));
        openTtpDetailModal_(parsed);
      } catch (err) {
        console.warn('[TTP] Could not parse row for detail popup', err);
      }
    });
  }

  function renderTTPTable() {
    syncTtpViewUi_();
    ttpSyncTableLayoutClass_();
    if (ttpViewMode === 'grouped') {
      renderTTPGrouped();
    } else {
      renderTTPFlat();
    }
    updateTTPSelectionInfo();
  }

  // ── GROUPED VIEW ─────────────────────────────────────────
  function renderTTPGrouped() {
    const body = document.getElementById('ttpTableBody');
    if (!body) return;
    bindTtpTableDelegationOnce();

    const groupCol   = ttpFields.find(h => normalizeTtpHeaderKey_(h) === 'GROUP NAME')
      || ttpFields.find(h => { const u = normalizeTtpHeaderKey_(h); return u.includes('GROUP') && u.includes('NAME') && !u.includes('FFB'); })
      || '';
    const companyCol = ttpFields.find(h => normalizeTtpHeaderKey_(h) === 'COMPANY NAME') || '';
    const millCol    = ttpFields.find(h => normalizeTtpHeaderKey_(h) === 'MILL NAME') || '';
    // Fixed header — no "Detail" column, fixed widths that don't depend on content
    document.getElementById('ttpTableHead').innerHTML = ttpMainTableHeadHtml_(true);

    let filtered = ttpApplyTableFilters_(ttpData);

    if (filtered.length === 0) {
      body.innerHTML = '<tr><td colspan="' + ttpMainTableColspan_(true) + '" style="text-align:center;padding:32px;color:#9C8A8A;">No data found</td></tr>';
      document.getElementById('ttpGroupInfo').textContent = '';
      return;
    }

    // Group by MILL NAME
    const groups = new Map();
    filtered.forEach(d => {
      const mill = (d[millCol] || '').toString().trim() || '(No Mill Name)';
      if (!groups.has(mill)) groups.set(mill, []);
      groups.get(mill).push(d);
    });

    document.getElementById('ttpGroupInfo').textContent =
      groups.size + ' mills · ' + filtered.length + ' records';

    let html = '';
    let gIdx = 0;
    groups.forEach((rows, millName) => {
      const first     = rows[0];
      const groupName = first[groupCol]   || '—';
      const compName  = first[companyCol] || '—';
      const subCount  = rows.length;
      const groupId   = 'ttpg-' + gIdx;

      const groupCategory = ttpCategoryGroupSummary_(rows);
      const avgCpoHtml = ttpGroupAvgPctHtml_(rows, 'cpo', subCount);
      const avgPkHtml = ttpGroupAvgPctHtml_(rows, 'pk', subCount);

      // ── Parent row ──
      html += `<tr class="ttp-group-row" data-group="${groupId}" data-expanded="0">
        <td>${groupName}</td>
        <td>${compName}</td>
        <td>
          <div class="ttp-mill-name-cell">
            <span class="ttp-group-chevron"><svg viewBox="0 0 24 24"><polyline points="9 18 15 12 9 6"/></svg></span>
            <span class="mill-name">${millName}</span>
            ${subCount > 1 ? `<span class="ttp-sub-badge">${subCount}</span>` : ''}
          </div>
        </td>
        <td>${avgCpoHtml}</td>
        <td>${avgPkHtml}</td>
        <td><span class="ttp-category-val">${groupCategory}</span></td>
        <td></td>
      </tr>`;

      html += '<tr class="ttp-group-detail-row hidden" data-parent="' + groupId + '">'
        + '<td colspan="' + ttpMainTableColspan_(true) + '" class="ttp-group-detail-cell">'
        + ttpGroupedNestedTableHtml_(rows)
        + '</td></tr>';

      gIdx++;
    });

    body.innerHTML = html;

  }

  // ── FLAT VIEW (default) — one row per record; click row for detail popup ──
  function renderTTPFlat() {
    const body = document.getElementById('ttpTableBody');
    if (!body) return;
    bindTtpTableDelegationOnce();

    let filtered = ttpApplyTableFilters_(ttpData);
    ttpFlatFilteredRows = filtered;

    const groupCol   = ttpFields.find(h => normalizeTtpHeaderKey_(h) === 'GROUP NAME')
      || ttpFields.find(h => { const u = normalizeTtpHeaderKey_(h); return u.includes('GROUP') && u.includes('NAME') && !u.includes('FFB'); })
      || '';
    const companyCol = ttpFields.find(h => normalizeTtpHeaderKey_(h) === 'COMPANY NAME') || '';
    const millCol    = ttpFields.find(h => normalizeTtpHeaderKey_(h) === 'MILL NAME') || '';

    document.getElementById('ttpTableHead').innerHTML = ttpMainTableHeadHtml_(false);
    document.getElementById('ttpGroupInfo').textContent =
      filtered.length + ' records · click a row for supplier detail';

    if (filtered.length === 0) {
      body.innerHTML = '<tr><td colspan="' + ttpMainTableColspan_(false) + '" style="text-align:center;padding:32px;color:#9C8A8A;">No data found</td></tr>';
      return;
    }

    body.innerHTML = filtered.map((d, i) => {
      const cpoPct = ttpPctCol ? ttpFormatCellPct_(d[ttpPctCol]) : '—';
      const pkPct = ttpPkPctCol ? ttpFormatCellPct_(d[ttpPkPctCol]) : '—';
      const category = ttpRowCategoryValue_(d);
      const supplierHint = ttpRowField_(d, ['FFB SUPPLIER NAME']);
      const titleAttr = supplierHint !== '—'
        ? ' title="View: ' + supplierHint.replace(/"/g, '&quot;') + '"'
        : ' title="View supplier detail"';
      return ''
        + '<tr class="mill-row-clickable ttp-flat-row" data-flat-idx="' + i + '"' + titleAttr + '>'
        + '<td>' + escHtml(d[groupCol] || '—') + '</td>'
        + '<td>' + escHtml(d[companyCol] || '—') + '</td>'
        + '<td><span class="mill-name">' + escHtml(d[millCol] || '—') + '</span></td>'
        + '<td><span class="ttp-pct-val">' + escHtml(String(cpoPct)) + '</span></td>'
        + '<td><span class="ttp-pct-val">' + escHtml(String(pkPct)) + '</span></td>'
        + '<td><span class="ttp-category-val">' + escHtml(category) + '</span></td>'
        + '</tr>';
    }).join('');
  }


  (function bindTtpToolbarIfPresent() {
    const btnAddTtp = document.getElementById('btn-add-ttp');
    const ttpSearchInp = document.getElementById('ttpSearch');
    const ttpSearchClearBtn = document.getElementById('ttpSearchClear');
    const ttpViewGrouped = document.getElementById('ttpViewGrouped');
    const ttpViewFlat = document.getElementById('ttpViewFlat');
    const ttpBtnSelect = document.getElementById('ttpBtnSelect');
    const ttpBtnFilter = document.getElementById('ttpBtnFilter');
    const ttpExportBtn = document.getElementById('btn-export-ttp-xlsx');
    if (!btnAddTtp || !ttpSearchInp || !ttpSearchClearBtn || !ttpViewGrouped || !ttpViewFlat || !ttpBtnSelect || !ttpBtnFilter || !ttpExportBtn) {
      console.warn('[dashboard] TTP panel toolbar incomplete; TTP controls skipped.');
      return;
    }
    const debouncedRenderTTPTable = debounce(function() {
      scheduleRenderTTPTable();
    }, 120);
    btnAddTtp.addEventListener('click', () => openModal('ttp', ttpFields.length ? ttpFields : [''], 'add', null));
    ttpSearchInp.addEventListener('input', function() {
      ttpSearch = this.value.toLowerCase().trim();
      if (this.value) {
        ttpSearchClearBtn.classList.add('show');
      } else {
        ttpSearchClearBtn.classList.remove('show');
      }
      debouncedRenderTTPTable();
    });

    ttpSearchClearBtn.addEventListener('click', function() {
      const searchInput = ttpSearchInp;
      if (!searchInput) return;
      searchInput.value = '';
      ttpSearch = '';
      this.classList.remove('show');
      debouncedRenderTTPTable.cancel();
      scheduleRenderTTPTable.flush();
      searchInput.focus();
    });

    ttpViewGrouped.addEventListener('click', function() {
      ttpViewMode = 'grouped';
      syncTtpViewUi_();
      scheduleRenderTTPTable();
    });
    ttpViewFlat.addEventListener('click', function() {
      ttpViewMode = 'flat';
      syncTtpViewUi_();
      scheduleRenderTTPTable();
    });
    syncTtpViewUi_();

    ttpBtnSelect.addEventListener('click', function(e) {
      e.stopPropagation();
      openTTPDropdown('ttpBtnSelect', 'ttpSelectPanel', 'ttpFilterPanel', 'ttpBtnFilter');
    });
    ttpBtnFilter.addEventListener('click', function(e) {
      e.stopPropagation();
      openTTPDropdown('ttpBtnFilter', 'ttpFilterPanel', 'ttpSelectPanel', 'ttpBtnSelect');
    });
    const ttpBtnReset = document.getElementById('ttpBtnResetFilters');
    if (ttpBtnReset) {
      ttpBtnReset.addEventListener('click', function() {
        resetTtpTableFilters_();
      });
    }
    if (!window.__sddTtpDropdownOutsideClickBound) {
      window.__sddTtpDropdownOutsideClickBound = true;
      document.addEventListener('click', function(e) {
        const ttpSelectWrap = document.getElementById('ttpSelectWrap');
        const ttpFilterWrap = document.getElementById('ttpFilterWrap');
        const ttpSelectPanel = document.getElementById('ttpSelectPanel');
        const ttpFilterPanel = document.getElementById('ttpFilterPanel');
        const ttpBtnSelectEl = document.getElementById('ttpBtnSelect');
        if (ttpSelectWrap && ttpSelectPanel && ttpBtnSelectEl && !ttpSelectWrap.contains(e.target)) {
          ttpSelectPanel.classList.remove('open');
          ttpBtnSelectEl.classList.remove('active');
        }
        if (ttpFilterWrap && ttpFilterPanel && !ttpFilterWrap.contains(e.target)) {
          ttpFilterPanel.classList.remove('open');
        }
      });
    }

    ttpExportBtn.addEventListener('click', function() {
      if (!ttpData.length) return;

      const colsToExport = ttpFieldsForColMode_(ttpColFilterMode);

      const rowsToExport = ttpApplyTableFilters_(ttpData);

      if (!rowsToExport.length) { alert('Tidak ada data untuk di-export.'); return; }

      if (typeof XLSX === 'undefined') {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
        script.onload = () => doExportXLSX(rowsToExport, colsToExport, ttpColFilterMode);
        document.head.appendChild(script);
      } else {
        doExportXLSX(rowsToExport, colsToExport, ttpColFilterMode);
      }
    });
  })();

  // ─── TTP SELECT (dynamic column-based filter) ─────────────
  let ttpSelectedCompanies = null; // null = all selected (values of the active filter col)
  let ttpColFilterMode = 'ttm';    // 'ttm' | 'ttp'
  let ttpFieldSectionsCache = null;
  let ttpCompanyCol = '';          // currently active filter column (legacy name kept for compat)
  let ttpActiveFilterCol = '';     // which column is currently being used for row-filter
  let ttpCompanyDropdownListenersBound = false;
  let ttpColumnDropdownListenersBound = false;
  let ttpUniqueValuesCache = Object.create(null);

  function buildCompanyDropdown() {
    if (!ttpData.length) return;
    // Default filter column: COMPANY NAME
    ttpCompanyCol = ttpFields.find(h => h.toUpperCase() === 'COMPANY NAME') ||
                    ttpFields.find(h => h.toLowerCase().includes('company name')) ||
                    ttpFields[1] || ttpFields[0] || '';
    ttpActiveFilterCol = ttpCompanyCol;

    buildColPickerChips();
    const values = getUniqueValuesForCol(ttpActiveFilterCol);
    if (ttpSelectedCompanies === null) ttpSelectedCompanies = new Set(values);

    renderCompanyList(values, '');

    if (!ttpCompanyDropdownListenersBound) {
      const ttpCoSearchEl = document.getElementById('ttpCompanySearch');
      const ttpSelAll = document.getElementById('ttpSelectAll');
      const ttpSelNone = document.getElementById('ttpSelectNone');
      if (!ttpCoSearchEl || !ttpSelAll || !ttpSelNone) return;
      ttpCompanyDropdownListenersBound = true;
      const debouncedRenderCompanyList = debounce(function(keyword) {
        renderCompanyList(getUniqueValuesForCol(ttpActiveFilterCol), keyword);
      }, 120);
      ttpCoSearchEl.addEventListener('input', function() {
        debouncedRenderCompanyList(this.value.toLowerCase());
      });

      ttpSelAll.addEventListener('click', () => {
        debouncedRenderCompanyList.cancel();
        ttpSelectedCompanies = new Set(getUniqueValuesForCol(ttpActiveFilterCol));
        renderCompanyList(getUniqueValuesForCol(ttpActiveFilterCol), (ttpCoSearchEl.value || '').toLowerCase());
        updateTTPSelection();
      });
      ttpSelNone.addEventListener('click', () => {
        debouncedRenderCompanyList.cancel();
        ttpSelectedCompanies = new Set();
        renderCompanyList(getUniqueValuesForCol(ttpActiveFilterCol), (ttpCoSearchEl.value || '').toLowerCase());
        updateTTPSelection();
      });
    }
  }

  function getUniqueValuesForCol(col) {
    if (!col) return [];
    if (ttpUniqueValuesCache[col]) return ttpUniqueValuesCache[col];
    const values = [...new Set(ttpData.map(d => (d[col] || '').toString()).filter(Boolean))].sort();
    ttpUniqueValuesCache[col] = values;
    return values;
  }

  function buildColPickerChips() {
    const container = document.getElementById('ttpColPickerList');
    if (!container) return;
    container.innerHTML = ttpFields.map(f => {
      const active = f === ttpActiveFilterCol ? 'style="background:var(--primary);color:white;border-color:var(--primary);"' : '';
      return `<button onclick="ttpSwitchFilterCol('${f.replace(/'/g,"\\'")}',this)"
        style="padding:3px 9px;border-radius:20px;border:1.5px solid rgba(74,28,28,0.15);background:white;color:#4A3535;font-size:11px;font-weight:500;cursor:pointer;white-space:nowrap;transition:all 0.15s;" ${active}>${f}</button>`;
    }).join('');
  }

  window.ttpSwitchFilterCol = function(col, btn) {
    ttpActiveFilterCol = col;
    ttpCompanyCol = col;
    ttpSelectedCompanies = null; // reset selection
    const values = getUniqueValuesForCol(col);
    ttpSelectedCompanies = new Set(values);
    // update button label
    const btnLabel = document.getElementById('ttpBtnSelectLabel');
    if (btnLabel) btnLabel.textContent = col.length > 18 ? col.substring(0,17)+'…' : col;
    // re-render chips to show active state
    buildColPickerChips();
    const ttpCoSearch = document.getElementById('ttpCompanySearch');
    if (ttpCoSearch) ttpCoSearch.value = '';
    renderCompanyList(values, '');
    updateTTPSelection();
  };

  function renderCompanyList(companies, query) {
    const list = document.getElementById('ttpCompanyList');
    if (!list) return;
    const filtered = query ? companies.filter(c => c.toLowerCase().includes(query)) : companies;
    if (!filtered.length) { list.innerHTML = '<div class="ttp-dropdown-empty">Tidak ditemukan</div>'; return; }
    list.innerHTML = filtered.map(c => {
      const checked = ttpSelectedCompanies && ttpSelectedCompanies.has(c) ? 'checked' : '';
      const id = 'ttp-co-' + btoa(encodeURIComponent(c)).replace(/[^a-zA-Z0-9]/g,'').substring(0,16);
      return `<div class="ttp-dropdown-item">
        <input type="checkbox" id="${id}" value="${c.replace(/"/g,'&quot;')}" ${checked} onchange="ttpToggleCompany(this)">
        <label for="${id}">${c}</label>
      </div>`;
    }).join('');
  }

  window.ttpToggleCompany = function(cb) {
    const val = cb.value;
    if (!ttpSelectedCompanies) ttpSelectedCompanies = new Set();
    if (cb.checked) ttpSelectedCompanies.add(val);
    else ttpSelectedCompanies.delete(val);
    updateTTPSelection();
  };

  function updateTTPSelection() {
    scheduleRenderTTPTable();
  }

  function getTtpDefaultFilterCol_() {
    return ttpFields.find(function(h) { return h.toUpperCase() === 'COMPANY NAME'; }) ||
      ttpFields.find(function(h) { return h.toLowerCase().includes('company name'); }) ||
      ttpFields[1] || ttpFields[0] || '';
  }

  function closeTtpDropdownPanels_() {
    const selectPanel = document.getElementById('ttpSelectPanel');
    const filterPanel = document.getElementById('ttpFilterPanel');
    const selBtn = document.getElementById('ttpBtnSelect');
    const filBtn = document.getElementById('ttpBtnFilter');
    if (selectPanel) selectPanel.classList.remove('open');
    if (filterPanel) filterPanel.classList.remove('open');
    if (selBtn) selBtn.classList.remove('active');
    if (filBtn) filBtn.classList.remove('active');
  }

  function ttpHasActiveTableFilters_() {
    if (ttpSearch) return true;
    if (!ttpData.length) return false;
    const allValues = getUniqueValuesForCol(ttpActiveFilterCol);
    if (ttpSelectedCompanies && ttpSelectedCompanies.size < allValues.length) return true;
    if (ttpColFilterMode !== 'ttm') return true;
    const defCol = getTtpDefaultFilterCol_();
    if (defCol && ttpActiveFilterCol && ttpActiveFilterCol !== defCol) return true;
    return false;
  }

  function resetTtpTableFilters_() {
    if (!ttpLoaded || !ttpData.length) return;
    ttpSearch = '';
    const searchInp = document.getElementById('ttpSearch');
    const searchClear = document.getElementById('ttpSearchClear');
    if (searchInp) searchInp.value = '';
    if (searchClear) searchClear.classList.remove('show');
    ttpSelectedCompanies = null;
    ttpColFilterMode = 'ttm';
    ttpFieldSectionsCache = null;
    const defCol = getTtpDefaultFilterCol_();
    ttpActiveFilterCol = defCol;
    ttpCompanyCol = defCol;
    const btnLabel = document.getElementById('ttpBtnSelectLabel');
    if (btnLabel) {
      btnLabel.textContent = defCol && defCol.length > 18 ? defCol.substring(0, 17) + '…' : (defCol || 'Select');
    }
    const coSearch = document.getElementById('ttpCompanySearch');
    if (coSearch) coSearch.value = '';
    closeTtpDropdownPanels_();
    buildColPickerChips();
    buildCompanyDropdown();
    buildColumnModePanel();
    scheduleRenderTTPTable();
  }

  function updateTTPSelectionInfo() {
    const info = document.getElementById('ttpSelectionInfo');
    const txt = document.getElementById('ttpSelectionText');
    const exportBtn = document.getElementById('btn-export-ttp-xlsx');
    if (!info || !txt || !exportBtn) return;
    const allValues = ttpData.length ? getUniqueValuesForCol(ttpActiveFilterCol) : [];
    const selCount = ttpSelectedCompanies ? ttpSelectedCompanies.size : allValues.length;
    const colModeLabel = ttpColFilterMode === 'ttp' ? 'TTP Only' : 'TTM Only';
    const isDefault = selCount === allValues.length && ttpColFilterMode === 'ttm';
    const colLabel = ttpActiveFilterCol || 'nilai';
    if (!isDefault || selCount < allValues.length || ttpColFilterMode !== 'ttm') {
      info.style.display = 'flex';
      const parts = [];
      if (selCount < allValues.length) parts.push('<span class="ttp-badge">' + selCount + ' ' + colLabel.toLowerCase() + '</span>');
      if (ttpColFilterMode !== 'ttm') parts.push('<span class="ttp-badge">' + colModeLabel + '</span>');
      txt.innerHTML = parts.join(' · ');
    } else {
      info.style.display = 'none';
    }
    exportBtn.disabled = false;
    const selBtn = document.getElementById('ttpBtnSelect');
    const filBtn = document.getElementById('ttpBtnFilter');
    if (selBtn) selBtn.classList.toggle('active', ttpSelectedCompanies !== null && ttpSelectedCompanies.size < allValues.length);
    if (filBtn) filBtn.classList.toggle('active', ttpColFilterMode !== 'ttm');
    const resetBtn = document.getElementById('ttpBtnResetFilters');
    if (resetBtn) resetBtn.disabled = !ttpHasActiveTableFilters_();
  }

  // ─── TTP FILTER COLUMNS (TTM Only / TTP Only) ─────────────
  function resolveTtpFieldSections_() {
    if (ttpFieldSectionsCache) return ttpFieldSectionsCache;
    const ttm = [];
    const ttp = [];
    let inTtp = false;
    let closedTtp = false;
    ttpFields.forEach(function(f) {
      if (closedTtp) return;
      const key = normalizeTtpHeaderKey_(f);
      if (!inTtp) {
        if (key.includes('FFB SUPPLIER')) {
          inTtp = true;
          ttp.push(f);
        } else {
          ttm.push(f);
        }
      } else {
        ttp.push(f);
        if (key.includes('ISCC')) closedTtp = true;
      }
    });
    ttpFieldSectionsCache = { ttm: ttm, ttp: ttp };
    return ttpFieldSectionsCache;
  }

  function ttpFieldsForColMode_(mode) {
    const sections = resolveTtpFieldSections_();
    if (mode === 'ttp') return sections.ttp.slice();
    return sections.ttm.slice();
  }

  function buildColumnModePanel() {
    if (!ttpFields.length) return;
    const list = document.getElementById('ttpColumnModeList');
    if (!list) return;
    const sections = resolveTtpFieldSections_();
    const modes = [
      { id: 'ttm', label: 'TTM Only', desc: 'No – UML ID', cols: sections.ttm },
      { id: 'ttp', label: 'TTP Only', desc: 'FFB Supplier – ISCC (Y/N)', cols: sections.ttp },
    ];
    list.innerHTML = modes.map(function(m) {
      const checked = ttpColFilterMode === m.id ? ' checked' : '';
      return ''
        + '<label class="ttp-col-mode-opt">'
        + '<input type="radio" name="ttpColMode" value="' + m.id + '"' + checked + '>'
        + '<span class="ttp-col-mode-opt-body">'
        + '<span class="ttp-col-mode-opt-title">' + escHtml(m.label) + '</span>'
        + '<span class="ttp-col-mode-opt-desc">' + escHtml(m.desc) + ' · ' + m.cols.length + ' kolom</span>'
        + '</span>'
        + '</label>';
    }).join('');

    if (!ttpColumnDropdownListenersBound) {
      ttpColumnDropdownListenersBound = true;
      list.addEventListener('change', function(e) {
        const inp = e.target;
        if (!inp || inp.name !== 'ttpColMode') return;
        ttpColFilterMode = inp.value === 'ttp' ? 'ttp' : 'ttm';
        updateTTPSelectionInfo();
      });
    }
  }

  // ─── TTP DROPDOWN TOGGLE (absolute under .ttp-dropdown-wrap — avoid position:fixed + ancestor transform mismatch) ──
  function openTTPDropdown(btnId, panelId, otherPanelId, otherBtnId) {
    const btn    = document.getElementById(btnId);
    const panel  = document.getElementById(panelId);
    const other  = document.getElementById(otherPanelId);
    if (!btn || !panel || !other) return;

    // Close the other panel
    other.classList.remove('open');
    if (otherBtnId) document.getElementById(otherBtnId).classList.remove('active');

    const isOpen = panel.classList.contains('open');
    if (isOpen) {
      panel.classList.remove('open');
      btn.classList.remove('active');
      panel.style.position = '';
      panel.style.top = '';
      panel.style.left = '';
      panel.style.right = '';
      panel.style.width = '';
      panel.style.minWidth = '';
      panel.style.maxWidth = '';
      return;
    }

    panel.style.position = '';
    panel.style.top = '';
    panel.style.left = '';
    panel.style.right = '';
    panel.style.width = '';
    panel.style.minWidth = '';
    panel.style.maxWidth = '';

    panel.classList.add('open');
    btn.classList.add('active');
  }

  function doExportXLSX(rows, cols, colMode) {
    const mode = colMode === 'ttp' ? 'ttp' : 'ttm';
    const sheetName = mode === 'ttp' ? 'TTP' : 'TTM';
    const filePrefix = mode === 'ttp' ? 'monitoring_ttp' : 'monitoring_ttm';
    const wsData = [cols, ...rows.map(d => cols.map(f => d[f] !== undefined ? d[f] : ''))];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // column widths
    ws['!cols'] = cols.map(f => ({ wch: Math.min(Math.max(f.length + 4, 14), 36) }));

    // Define border style
    const borderStyle = {
      top: { style: 'thin', color: { rgb: '000000' } },
      bottom: { style: 'thin', color: { rgb: '000000' } },
      left: { style: 'thin', color: { rgb: '000000' } },
      right: { style: 'thin', color: { rgb: '000000' } }
    };

    // Style header (row 0) - Blue background, white text, bold, centered
    cols.forEach((_, ci) => {
      const cellAddr = XLSX.utils.encode_cell({ r: 0, c: ci });
      if (!ws[cellAddr]) ws[cellAddr] = { t: 's', v: '' };
      ws[cellAddr].s = {
        font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
        fill: { fgColor: { rgb: '4472C4' } },  // Blue color
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: borderStyle
      };
    });

    // Style data cells (rows 1 onwards) - Add borders and alignment
    rows.forEach((row, ri) => {
      cols.forEach((_, ci) => {
        const cellAddr = XLSX.utils.encode_cell({ r: ri + 1, c: ci });
        if (!ws[cellAddr]) ws[cellAddr] = { t: 's', v: '' };
        if (!ws[cellAddr].s) ws[cellAddr].s = {};
        ws[cellAddr].s = {
          ...ws[cellAddr].s,
          border: borderStyle,
          alignment: { vertical: 'top', wrapText: false }
        };
      });
    });

    // Set row height for header
    if (!ws['!rows']) ws['!rows'] = [];
    ws['!rows'][0] = { hpt: 25, hpx: 25 };

    // Add autofilter to header row
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ 
      s: { r: 0, c: 0 }, 
      e: { r: rows.length, c: cols.length - 1 } 
    })};

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
    XLSX.writeFile(wb, `${filePrefix}_${stamp}.xlsx`, { cellStyles: true });
  }

  // ─── GRIEVANCE DATA ─────────────────────────────────────
  const GRV_FIELDS = ['Grievance ID','Date Received','Complainant','Grievance Source','Grievance Publisher','Grievance Category','Subject','Relationship','Grievance Subject','Grievance Subject Group','Subject ID','Grievance Description','Risk Classification','Verification Findings','Corrective Action','Preventive Action','Responsible Div./Dep.','Grievance Status','Date Closed','Action Taken','Published'];
  const GRV_LONG = ['Grievance Description','Verification Findings','Corrective Action','Preventive Action','Action Taken'];
  const GRV_SHORT = ['Grievance ID','Date Received','Grievance Category','Subject','Grievance Subject Group','Risk Classification','Grievance Status','Date Closed'];

  let grvData = [], grvLoaded = false, grvSearch = '';
  let grvTableDelegationBound = false;

  function prepareGrvRowPerfCache(row) {
    if (!row || typeof row !== 'object') return row;
    row._sddSearchBlob = GRV_FIELDS.map(function(f) {
      return String(row[f] || '').toLowerCase();
    }).join('|');
    return row;
  }

  const scheduleRenderGrvTable = makeRafScheduler(function() {
    renderGrvTable();
  });

  function bindGrvTableDelegationOnce() {
    if (grvTableDelegationBound) return;
    const body = document.getElementById('grvTableBody');
    if (!body) return;
    grvTableDelegationBound = true;
    body.addEventListener('click', function(e) {
      const expandBtn = e.target.closest('.btn-expand');
      if (expandBtn && body.contains(expandBtn)) {
        const idx = expandBtn.dataset.idx;
        const detail = document.getElementById('grv-detail-' + idx);
        if (!detail) return;
        const isOpen = detail.classList.contains('open');
        document.querySelectorAll('.grv-detail.open').forEach(d => d.classList.remove('open'));
        document.querySelectorAll('.btn-expand.open').forEach(b => b.classList.remove('open'));
        if (!isOpen) {
          detail.classList.add('open');
          expandBtn.classList.add('open');
        }
        return;
      }

      const editBtn = e.target.closest('.btn-edit');
      if (editBtn && body.contains(editBtn)) {
        const row = JSON.parse(editBtn.dataset.row.replace(/&#39;/g, "'"));
        openModal('grievance', GRV_FIELDS, 'edit', row);
        return;
      }

      const delBtn = e.target.closest('.btn-delete');
      if (delBtn && body.contains(delBtn)) {
        openConfirm(delBtn.dataset.sheet, parseInt(delBtn.dataset.rownum, 10));
      }
    });
  }

  async function loadGrvDataImpl() {
    const loading = document.getElementById('grv-loading');
    const errorEl = document.getElementById('grv-error');
    const table = document.getElementById('grvTable');
    try {
      loading.style.display = 'block';
      errorEl.style.display = 'none';
      table.style.display = 'none';
      grvData = await apiGet('grievance');
      grvData = grvData.map(prepareGrvRowPerfCache);
      grvLoaded = true;

      document.getElementById('grv-stat-total').textContent = grvData.length;
      document.getElementById('grv-stat-open').textContent = grvData.filter(d => (d['Grievance Status']||'').toLowerCase().includes('open')).length;
      document.getElementById('grv-stat-closed').textContent = grvData.filter(d => (d['Grievance Status']||'').toLowerCase().includes('closed')).length;
      document.getElementById('grv-stat-high').textContent = grvData.filter(d => (d['Risk Classification']||'').toLowerCase().includes('high')).length;

      loading.style.display = 'none';
      table.style.display = 'table';
      scheduleRenderGrvTable();
    } catch(err) {
      loading.style.display = 'none';
      errorEl.style.display = 'block';
      errorEl.textContent = 'Gagal memuat data: ' + err.message;
    }
  }

  async function loadGrvData() {
    if (grvLoadPromise) return grvLoadPromise;
    grvLoadPromise = loadGrvDataImpl();
    try {
      await grvLoadPromise;
    } finally {
      grvLoadPromise = null;
    }
  }

  function riskBadge(val) {
    if (!val) return '—';
    const v = val.toLowerCase();
    const cls = v.includes('high') ? 'risk-high' : v.includes('med') ? 'risk-med' : 'risk-low';
    return `<span class="status-badge ${cls}"><span class="s-dot"></span>${val}</span>`;
  }

  function statusBadgeGrv(val) {
    if (!val) return '—';
    const cls = val.toLowerCase().includes('open') ? 'grv-open' : 'grv-closed';
    return `<span class="status-badge ${cls}"><span class="s-dot"></span>${val}</span>`;
  }

  function renderGrvTable() {
    const body = document.getElementById('grvTableBody');
    if (!body) return;
    bindGrvTableDelegationOnce();
    const q = grvSearch;
    const filtered = grvData.filter(d =>
      !q || (d._sddSearchBlob || '').includes(q)
    );
    if (filtered.length === 0) {
      body.innerHTML = `<tr><td colspan="9" style="text-align:center;padding:32px;color:#9C8A8A;">No data found</td></tr>`;
      return;
    }
    body.innerHTML = filtered.map((d, i) => {
      const detailHTML = GRV_FIELDS.map(f => `
        <div class="grv-detail-item ${GRV_LONG.includes(f) ? 'full' : ''}">
          <div class="grv-detail-label">${f}</div>
          <div class="grv-detail-val">${d[f] || '—'}</div>
        </div>`).join('');
      return `
        <tr class="grv-main-row" data-idx="${i}">
          <td><span class="mill-id">${d['Grievance ID'] || '—'}</span></td>
          <td>${d['Date Received'] || '—'}</td>
          <td>${d['Grievance Category'] || '—'}</td>
          <td><span class="mill-name">${d['Subject'] || '—'}</span></td>
          <td>${d['Grievance Subject Group'] || '—'}</td>
          <td>${riskBadge(d['Risk Classification'])}</td>
          <td>${statusBadgeGrv(d['Grievance Status'])}</td>
          <td>${d['Date Closed'] || '—'}</td>
          <td>
            <div class="row-actions">
              <button class="btn-expand" data-idx="${i}" title="View detail">▾</button>
              <button class="btn-row btn-edit" data-row='${JSON.stringify(d).replace(/'/g,"&#39;")}' data-sheet="grievance">Edit</button>
              <button class="btn-row btn-delete" data-rownum="${d._row}" data-sheet="grievance">Del</button>
            </div>
          </td>
        </tr>
        <tr class="grv-expand-row">
          <td colspan="9"><div class="grv-detail" id="grv-detail-${i}">${detailHTML}</div></td>
        </tr>`;
    }).join('');

  }

  (function bindGrievanceToolbarIfPresent() {
    const grvSearchEl = document.getElementById('grvSearch');
    const grvSearchClearEl = document.getElementById('grvSearchClear');
    const btnAddGrv = document.getElementById('btn-add-grv');
    const btnExportGrv = document.getElementById('btn-export-grv');
    if (!grvSearchEl || !grvSearchClearEl || !btnAddGrv || !btnExportGrv) {
      console.warn('[dashboard] Grievance panel toolbar incomplete; grievance controls skipped.');
      return;
    }
    const debouncedRenderGrvTable = debounce(function() {
      scheduleRenderGrvTable();
    }, 120);
    grvSearchEl.addEventListener('input', function() {
      grvSearch = this.value.toLowerCase().trim();
      if (this.value) {
        grvSearchClearEl.classList.add('show');
      } else {
        grvSearchClearEl.classList.remove('show');
      }
      debouncedRenderGrvTable();
    });

    grvSearchClearEl.addEventListener('click', function() {
      const searchInput = grvSearchEl;
      if (!searchInput) return;
      searchInput.value = '';
      grvSearch = '';
      this.classList.remove('show');
      debouncedRenderGrvTable.cancel();
      scheduleRenderGrvTable.flush();
      searchInput.focus();
    });
    btnAddGrv.addEventListener('click', () => openModal('grievance', GRV_FIELDS, 'add', null));
    btnExportGrv.addEventListener('click', function() {
      if (!grvData.length) return;
      const csv = '\uFEFF' + [GRV_FIELDS, ...grvData.map(d => GRV_FIELDS.map(f => `"${(d[f]||'').toString().replace(/"/g,'""')}"`))]
        .map(r => r.join(',')).join('\n');
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
      a.download = 'grievance_monitoring.csv';
      a.click();
    });
  })();

  // ─── CONTACT LIST SUPPLIER ───────────────────────────────
  const CLS_EXPORT_FIELDS = [
    'Group Name', 'Company Name', 'Supplier Type',
    'Sustainability PIC', 'Phone Number', 'Email', 'approved_at', 'submission_id',
  ];
  const CLS_MODAL_FIELDS = [
    'Group Name', 'Company Name', 'Supplier Type',
    'Sustainability PIC', 'Phone Number', 'Email',
  ];
  let contactListData = [];
  let contactListLoaded = false;
  let contactListSearch = '';
  let contactLoadPromise = null;
  let clsRenderScheduled = false;

  function formatClsApprovedAt_(raw) {
    const s = String(raw || '').trim();
    if (!s) return '—';
    if (s.length >= 10 && s.indexOf('T') !== -1) return s.slice(0, 10);
    return s.length > 16 ? s.slice(0, 16) : s;
  }

  function prepareContactListRow_(d) {
    const parts = [
      d['Group Name'], d['Company Name'], d['Supplier Type'],
      d['Sustainability PIC'], d['Phone Number'], d['Email'], d['submission_id'],
    ].map(function(x) { return String(x || '').toLowerCase(); });
    d._clsSearchBlob = parts.join(' ');
    return d;
  }

  function updateContactListStats_() {
    const totalEl = document.getElementById('cls-stat-total');
    if (!totalEl) return;
    const rows = contactListData;
    totalEl.textContent = String(rows.length);
    const countType = function(t) {
      return rows.filter(function(d) {
        return String(d['Supplier Type'] || '').toUpperCase() === t;
      }).length;
    };
    const millEl = document.getElementById('cls-stat-mill');
    const kcpEl = document.getElementById('cls-stat-kcp');
    const traderEl = document.getElementById('cls-stat-trader');
    if (millEl) millEl.textContent = String(countType('MILL'));
    if (kcpEl) kcpEl.textContent = String(countType('KCP'));
    if (traderEl) traderEl.textContent = String(countType('TRADER'));
  }

  function renderContactListTable_() {
    const body = document.getElementById('clsTableBody');
    if (!body) return;
    const q = contactListSearch;
    const filtered = contactListData.filter(function(d) {
      return !q || (d._clsSearchBlob || '').includes(q);
    });
    filtered.sort(function(a, b) {
      const ta = String(a['approved_at'] || a['updated_at'] || '');
      const tb = String(b['approved_at'] || b['updated_at'] || '');
      return tb.localeCompare(ta);
    });
    if (!filtered.length) {
      body.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:#9C8A8A;">'
        + (q ? 'No results match your search.' : 'No contacts yet. PIC details appear here after an SDD screening is approved.')
        + '</td></tr>';
      return;
    }
    body.innerHTML = filtered.map(function(d) {
      const pic = d['Sustainability PIC'] || '—';
      const phone = d['Phone Number'] || '—';
      const email = String(d['Email'] || '').trim();
      const phoneCell = phone !== '—'
        ? '<a class="cls-phone-link" href="tel:' + escHtml(String(phone).replace(/\s/g, '')) + '">' + escHtml(phone) + '</a>'
        : '—';
      const emailCell = email
        ? '<a class="cls-phone-link" href="mailto:' + escHtml(email) + '">' + escHtml(email) + '</a>'
        : '—';
      return '<tr>'
        + '<td><span class="mill-name">' + escHtml(d['Group Name'] || '—') + '</span></td>'
        + '<td>' + escHtml(d['Company Name'] || '—') + '</td>'
        + '<td><span class="status-badge risk-low"><span class="s-dot"></span>' + escHtml(d['Supplier Type'] || '—') + '</span></td>'
        + '<td><strong>' + escHtml(pic) + '</strong></td>'
        + '<td>' + phoneCell + '</td>'
        + '<td>' + emailCell + '</td>'
        + '<td>' + escHtml(formatClsApprovedAt_(d['approved_at'])) + '</td>'
        + '<td style="text-align:right;"><button type="button" class="btn-sm btn-outline cls-edit-btn" data-row="' + Number(d._row || 0) + '" title="Edit contact">Edit</button></td>'
        + '</tr>';
    }).join('');

    body.querySelectorAll('.cls-edit-btn').forEach(function(btn) {
      btn.addEventListener('click', function() {
        var rowNum = Number(btn.getAttribute('data-row') || 0);
        if (!rowNum) return;
        var hit = contactListData.find(function(it) { return Number(it._row || 0) === rowNum; });
        if (!hit) return;
        openModal('contactSupplier', CLS_MODAL_FIELDS, 'edit', hit);
      });
    });
  }

  function scheduleRenderContactListTable_() {
    if (clsRenderScheduled) return;
    clsRenderScheduled = true;
    requestAnimationFrame(function() {
      clsRenderScheduled = false;
      renderContactListTable_();
    });
  }

  async function loadContactListDataImpl_(force) {
    const loading = document.getElementById('cls-loading');
    const errorEl = document.getElementById('cls-error');
    const table = document.getElementById('clsTable');
    if (!loading || !errorEl || !table) return;
    try {
      loading.style.display = 'block';
      errorEl.style.display = 'none';
      table.style.display = 'none';
      const raw = await apiGet('contactSupplier');
      contactListData = (Array.isArray(raw) ? raw : []).map(prepareContactListRow_);
      contactListLoaded = true;
      updateContactListStats_();
      loading.style.display = 'none';
      table.style.display = 'table';
      scheduleRenderContactListTable_();
    } catch (err) {
      loading.style.display = 'none';
      errorEl.style.display = 'block';
      errorEl.textContent = 'Failed to load Contact List Supplier: ' + ((err && err.message) ? err.message : String(err));
      if (force) throw err;
    }
  }

  async function loadContactListData(force) {
    if (contactListLoaded && !force) {
      scheduleRenderContactListTable_();
      return;
    }
    if (contactLoadPromise) return contactLoadPromise;
    contactLoadPromise = loadContactListDataImpl_(!!force);
    try {
      await contactLoadPromise;
    } finally {
      contactLoadPromise = null;
    }
  }

  window.__contactListInvalidate = function() {
    contactListLoaded = false;
  };
  window.loadContactListData = loadContactListData;

  (function bindContactListToolbarIfPresent_() {
    const searchEl = document.getElementById('clsSearch');
    const clearEl = document.getElementById('clsSearchClear');
    const btnRefresh = document.getElementById('btn-refresh-cls');
    const btnExport = document.getElementById('btn-export-cls');
    const btnAdd = document.getElementById('btn-add-cls');
    if (!searchEl || !clearEl || !btnRefresh || !btnExport || !btnAdd) return;

    const debouncedRender = debounce(function() {
      scheduleRenderContactListTable_();
    }, 120);

    searchEl.addEventListener('input', function() {
      contactListSearch = this.value.toLowerCase().trim();
      if (this.value) clearEl.classList.add('show');
      else clearEl.classList.remove('show');
      debouncedRender();
    });

    clearEl.addEventListener('click', function() {
      searchEl.value = '';
      contactListSearch = '';
      this.classList.remove('show');
      if (debouncedRender.cancel) debouncedRender.cancel();
      if (debouncedRender.flush) debouncedRender.flush();
      else scheduleRenderContactListTable_();
      searchEl.focus();
    });

    btnRefresh.addEventListener('click', function() {
      contactListLoaded = false;
      loadContactListData(true);
    });

    btnAdd.addEventListener('click', function() {
      openModal('contactSupplier', CLS_MODAL_FIELDS, 'add', null);
    });

    btnExport.addEventListener('click', function() {
      if (!contactListData.length) return;
      const csv = '\uFEFF' + [CLS_EXPORT_FIELDS].concat(
        contactListData.map(function(d) {
          return CLS_EXPORT_FIELDS.map(function(f) {
            return '"' + String(d[f] || '').replace(/"/g, '""') + '"';
          });
        })
      ).map(function(r) { return r.join(','); }).join('\n');
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
      a.download = 'contact_list_supplier.csv';
      a.click();
    });
  })();

  // ─── NO BUY LIST (NBL + Unilever NBL) ────────────────────
  const NBL_REGISTRY_FIELDS = ['Riser', 'Group Name NBL', 'Company Name NBL', 'SOURCE'];
  const UNILEVER_NBL_FIELDS = [
    'Riser', 'UML ID', 'COMPANY NAME', 'MILL NAME', 'COUNTRY', 'PROVINCE',
    'DISTRICT / REGENCY', 'LAT.', 'LONG.',
  ];
  let nblRegistryData = [];
  let nblUnileverData = [];
  let nblRegistryLoaded = false;
  let nblUnileverLoaded = false;
  let nblLoadPromise = null;
  let nblActiveSource = 'nbl';
  let nblSearchRegistry = '';
  let nblSearchUnilever = '';
  let nblRenderScheduled = false;

  function nblPickField_(row, keys) {
    for (var i = 0; i < keys.length; i++) {
      var v = row[keys[i]];
      if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
    }
    var norm = {};
    Object.keys(row || {}).forEach(function(k) {
      if (k === '_row') return;
      norm[String(k).trim().toLowerCase()] = row[k];
    });
    for (var j = 0; j < keys.length; j++) {
      var lk = String(keys[j]).trim().toLowerCase();
      if (norm[lk] !== undefined && norm[lk] !== null && String(norm[lk]).trim() !== '') {
        return String(norm[lk]).trim();
      }
    }
    return '';
  }

  function prepareNblRegistryRow_(d) {
    d._nblRiser = nblPickField_(d, ['Riser']);
    d._nblGroup = nblPickField_(d, ['Group Name NBL', 'Group Name']);
    d._nblCompany = nblPickField_(d, ['Company Name NBL', 'Company Name']);
    d._nblSource = nblPickField_(d, ['SOURCE', 'Source']);
    d._nblSearchBlob = [d._nblRiser, d._nblGroup, d._nblCompany, d._nblSource]
      .join(' ').toLowerCase();
    return d;
  }

  function prepareNblUnileverRow_(d) {
    d._nblRiser = nblPickField_(d, ['Riser', 'RISER', 'NO.', 'NO', 'No.']);
    d._nblNo = d._nblRiser;
    d._nblUml = nblPickField_(d, ['UML ID']);
    d._nblCompany = nblPickField_(d, ['COMPANY NAME', 'Company Name']);
    d._nblMill = nblPickField_(d, ['MILL NAME', 'Mill Name']);
    d._nblCountry = nblPickField_(d, ['COUNTRY', 'Country']);
    d._nblProvince = nblPickField_(d, ['PROVINCE', 'Province']);
    d._nblDistrict = nblPickField_(d, ['DISTRICT / REGENCY', 'DISTRICT/REGENCY', 'District / Regency']);
    d._nblLat = nblPickField_(d, ['LAT.', 'LAT', 'Latitude', 'Lat']);
    d._nblLong = nblPickField_(d, ['LONG.', 'LONG', 'Longitude', 'Long']);
    d._nblSearchBlob = [
      d._nblRiser, d._nblNo, d._nblUml, d._nblCompany, d._nblMill, d._nblCountry,
      d._nblProvince, d._nblDistrict, d._nblLat, d._nblLong,
    ].join(' ').toLowerCase();
    return d;
  }

  function updateNblStats_() {
    var regEl = document.getElementById('nbl-stat-registry');
    var uniEl = document.getElementById('nbl-stat-unilever');
    var grpEl = document.getElementById('nbl-stat-groups');
    var ctyEl = document.getElementById('nbl-stat-countries');
    if (regEl) regEl.textContent = String(nblRegistryData.length);
    if (uniEl) uniEl.textContent = String(nblUnileverData.length);
    if (grpEl) {
      var groups = {};
      nblRegistryData.forEach(function(d) {
        var g = d._nblGroup;
        if (g) groups[g.toLowerCase()] = true;
      });
      grpEl.textContent = String(Object.keys(groups).length);
    }
    if (ctyEl) {
      var countries = {};
      nblUnileverData.forEach(function(d) {
        var c = d._nblCountry;
        if (c) countries[c.toLowerCase()] = true;
      });
      ctyEl.textContent = String(Object.keys(countries).length);
    }
  }

  function setNblActiveSource_(source) {
    nblActiveSource = source === 'unilever' ? 'unilever' : 'nbl';
    document.querySelectorAll('#panel-no-buy-list .nbl-source-tab').forEach(function(btn) {
      var on = btn.getAttribute('data-nbl-source') === nblActiveSource;
      btn.classList.toggle('active', on);
      btn.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    var paneReg = document.getElementById('nbl-pane-registry');
    var paneUni = document.getElementById('nbl-pane-unilever');
    if (paneReg) {
      paneReg.classList.toggle('active', nblActiveSource === 'nbl');
      paneReg.hidden = nblActiveSource !== 'nbl';
    }
    if (paneUni) {
      paneUni.classList.toggle('active', nblActiveSource === 'unilever');
      paneUni.hidden = nblActiveSource !== 'unilever';
    }
    var titleEl = document.getElementById('nbl-table-title');
    if (titleEl) {
      titleEl.textContent = nblActiveSource === 'unilever' ? 'Unilever NBL' : 'NBL registry';
    }
    var searchEl = document.getElementById('nblSearch');
    if (searchEl) {
      searchEl.value = nblActiveSource === 'unilever' ? nblSearchUnilever : nblSearchRegistry;
      searchEl.placeholder = nblActiveSource === 'unilever'
        ? 'Search riser, UML ID, company, mill, location...'
        : 'Search riser, group, company, source...';
    }
    scheduleRenderNblTable_();
  }

  function renderNblRegistryTable_() {
    var body = document.getElementById('nblTableRegistryBody');
    var table = document.getElementById('nblTableRegistry');
    if (!body || !table) return;
    var q = nblSearchRegistry;
    var filtered = nblRegistryData.filter(function(d) {
      return !q || (d._nblSearchBlob || '').includes(q);
    });
    if (!filtered.length) {
      body.innerHTML = '<tr><td colspan="4" style="text-align:center;padding:32px;color:#9C8A8A;">'
        + (q ? 'No results match your search.' : 'No NBL registry rows in the sheet yet.')
        + '</td></tr>';
      table.style.display = 'table';
      return;
    }
    body.innerHTML = filtered.map(function(d) {
      return '<tr>'
        + '<td>' + escHtml(d._nblRiser || '—') + '</td>'
        + '<td><span class="mill-name">' + escHtml(d._nblGroup || '—') + '</span></td>'
        + '<td>' + escHtml(d._nblCompany || '—') + '</td>'
        + '<td>' + escHtml(d._nblSource || '—') + '</td>'
        + '</tr>';
    }).join('');
    table.style.display = 'table';
  }

  function renderNblUnileverTable_() {
    var body = document.getElementById('nblTableUnileverBody');
    var table = document.getElementById('nblTableUnilever');
    if (!body || !table) return;
    var q = nblSearchUnilever;
    var filtered = nblUnileverData.filter(function(d) {
      return !q || (d._nblSearchBlob || '').includes(q);
    });
    if (!filtered.length) {
      body.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:32px;color:#9C8A8A;">'
        + (q ? 'No results match your search.' : 'No Unilever NBL rows in the sheet yet.')
        + '</td></tr>';
      table.style.display = 'table';
      return;
    }
    body.innerHTML = filtered.map(function(d) {
      return '<tr>'
        + '<td>' + escHtml(d._nblRiser || d._nblNo || '—') + '</td>'
        + '<td><span class="mill-id">' + escHtml(d._nblUml || '—') + '</span></td>'
        + '<td>' + escHtml(d._nblCompany || '—') + '</td>'
        + '<td><span class="mill-name">' + escHtml(d._nblMill || '—') + '</span></td>'
        + '<td>' + escHtml(d._nblCountry || '—') + '</td>'
        + '<td>' + escHtml(d._nblProvince || '—') + '</td>'
        + '<td>' + escHtml(d._nblDistrict || '—') + '</td>'
        + '<td class="nbl-coord">' + escHtml(d._nblLat || '—') + '</td>'
        + '<td class="nbl-coord">' + escHtml(d._nblLong || '—') + '</td>'
        + '</tr>';
    }).join('');
    table.style.display = 'table';
  }

  function scheduleRenderNblTable_() {
    if (nblRenderScheduled) return;
    nblRenderScheduled = true;
    requestAnimationFrame(function() {
      nblRenderScheduled = false;
      if (nblActiveSource === 'unilever') renderNblUnileverTable_();
      else renderNblRegistryTable_();
    });
  }

  function showNblTablesAfterLoad_() {
    var loading = document.getElementById('nbl-loading');
    if (loading) loading.style.display = 'none';
    var tblReg = document.getElementById('nblTableRegistry');
    var tblUni = document.getElementById('nblTableUnilever');
    if (nblActiveSource === 'unilever') {
      if (tblReg) tblReg.style.display = 'none';
      if (tblUni) tblUni.style.display = 'table';
    } else {
      if (tblUni) tblUni.style.display = 'none';
      if (tblReg) tblReg.style.display = 'table';
    }
    scheduleRenderNblTable_();
  }

  async function loadNoBuyListDataImpl_(force) {
    var loading = document.getElementById('nbl-loading');
    var errorEl = document.getElementById('nbl-error');
    if (!loading) return;
    try {
      loading.style.display = 'block';
      if (errorEl) errorEl.style.display = 'none';
      var tblReg = document.getElementById('nblTableRegistry');
      var tblUni = document.getElementById('nblTableUnilever');
      if (tblReg) tblReg.style.display = 'none';
      if (tblUni) tblUni.style.display = 'none';

      var results = await Promise.all([
        apiGet('nbl'),
        apiGet('unileverNbl'),
      ]);
      nblRegistryData = (Array.isArray(results[0]) ? results[0] : []).map(prepareNblRegistryRow_);
      nblUnileverData = (Array.isArray(results[1]) ? results[1] : []).map(prepareNblUnileverRow_);
      nblRegistryLoaded = true;
      nblUnileverLoaded = true;
      updateNblStats_();
      showNblTablesAfterLoad_();
    } catch (err) {
      loading.style.display = 'none';
      if (errorEl) {
        errorEl.style.display = 'block';
        errorEl.textContent = 'Failed to load No Buy List: ' + ((err && err.message) ? err.message : String(err));
      }
      if (force) throw err;
    }
  }

  async function loadNoBuyListData(force) {
    if (nblRegistryLoaded && nblUnileverLoaded && !force) {
      showNblTablesAfterLoad_();
      return;
    }
    if (nblLoadPromise) return nblLoadPromise;
    nblLoadPromise = loadNoBuyListDataImpl_(!!force);
    try {
      await nblLoadPromise;
    } finally {
      nblLoadPromise = null;
    }
  }

  (function bindNoBuyListToolbarIfPresent_() {
    var panel = document.getElementById('panel-no-buy-list');
    if (!panel) return;

    panel.querySelectorAll('.nbl-source-tab').forEach(function(btn) {
      btn.addEventListener('click', function() {
        setNblActiveSource_(btn.getAttribute('data-nbl-source') || 'nbl');
      });
    });

    var searchEl = document.getElementById('nblSearch');
    var clearEl = document.getElementById('nblSearchClear');
    var btnRefresh = document.getElementById('btn-refresh-nbl');
    var btnExport = document.getElementById('btn-export-nbl');
    if (!searchEl || !clearEl || !btnRefresh || !btnExport) return;

    var debouncedRender = debounce(function() {
      scheduleRenderNblTable_();
    }, 120);

    searchEl.addEventListener('input', function() {
      var q = this.value.toLowerCase().trim();
      if (nblActiveSource === 'unilever') nblSearchUnilever = q;
      else nblSearchRegistry = q;
      if (this.value) clearEl.classList.add('show');
      else clearEl.classList.remove('show');
      debouncedRender();
    });

    clearEl.addEventListener('click', function() {
      searchEl.value = '';
      if (nblActiveSource === 'unilever') nblSearchUnilever = '';
      else nblSearchRegistry = '';
      this.classList.remove('show');
      if (debouncedRender.cancel) debouncedRender.cancel();
      if (debouncedRender.flush) debouncedRender.flush();
      else scheduleRenderNblTable_();
      searchEl.focus();
    });

    btnRefresh.addEventListener('click', function() {
      nblRegistryLoaded = false;
      nblUnileverLoaded = false;
      loadNoBuyListData(true);
    });

    btnExport.addEventListener('click', function() {
      var fields = nblActiveSource === 'unilever' ? UNILEVER_NBL_FIELDS : NBL_REGISTRY_FIELDS;
      var rows = nblActiveSource === 'unilever' ? nblUnileverData : nblRegistryData;
      if (!rows.length) return;
      var csv = '\uFEFF' + [fields].concat(
        rows.map(function(d) {
          return fields.map(function(f) {
            if (nblActiveSource === 'unilever') {
              if (f === 'Riser') return '"' + String(d._nblRiser || d._nblNo || '').replace(/"/g, '""') + '"';
              if (f === 'UML ID') return '"' + String(d._nblUml || '').replace(/"/g, '""') + '"';
              if (f === 'COMPANY NAME') return '"' + String(d._nblCompany || '').replace(/"/g, '""') + '"';
              if (f === 'MILL NAME') return '"' + String(d._nblMill || '').replace(/"/g, '""') + '"';
              if (f === 'COUNTRY') return '"' + String(d._nblCountry || '').replace(/"/g, '""') + '"';
              if (f === 'PROVINCE') return '"' + String(d._nblProvince || '').replace(/"/g, '""') + '"';
              if (f === 'DISTRICT / REGENCY') return '"' + String(d._nblDistrict || '').replace(/"/g, '""') + '"';
              if (f === 'LAT.') return '"' + String(d._nblLat || '').replace(/"/g, '""') + '"';
              if (f === 'LONG.') return '"' + String(d._nblLong || '').replace(/"/g, '""') + '"';
            }
            if (f === 'Riser') return '"' + String(d._nblRiser || '').replace(/"/g, '""') + '"';
            if (f === 'Group Name NBL') return '"' + String(d._nblGroup || '').replace(/"/g, '""') + '"';
            if (f === 'Company Name NBL') return '"' + String(d._nblCompany || '').replace(/"/g, '""') + '"';
            if (f === 'SOURCE') return '"' + String(d._nblSource || '').replace(/"/g, '""') + '"';
            return '"' + String(d[f] || '').replace(/"/g, '""') + '"';
          });
        })
      ).map(function(r) { return r.join(','); }).join('\n');
      var a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
      a.download = nblActiveSource === 'unilever' ? 'unilever_nbl.csv' : 'nbl_registry.csv';
      a.click();
    });
  })();

  // ─── SDD: CHECK NBL (import vs NBL + Unilever NBL sheets) ───
  let _nblListsCache = null;
  let _nblListsCacheAt = 0;
  const NBL_LISTS_CACHE_MS = 120000;

  function nblNamesEqual_(a, b) {
    var na = normalizeLooseKey(a);
    var nb = normalizeLooseKey(b);
    return !!(na && nb && na === nb);
  }

  function getSddPrimaryForNblCheck_() {
    var p = window._loadedPrimarySddRow || window._sddImportFirstRow || {};
    return {
      group: normalizeCellText(p['Group Name'] || p['Grup Name'] || ''),
      company: normalizeCellText(p['Company Name'] || ''),
      mill: normalizeCellText(p['Mill Name'] || ''),
    };
  }

  async function ensureNblListsForCheck_() {
    var now = Date.now();
    if (_nblListsCache && (now - _nblListsCacheAt) < NBL_LISTS_CACHE_MS) {
      return _nblListsCache;
    }
    var results = await Promise.all([apiGet('nbl'), apiGet('unileverNbl')]);
    _nblListsCache = {
      registry: (Array.isArray(results[0]) ? results[0] : []).map(prepareNblRegistryRow_),
      unilever: (Array.isArray(results[1]) ? results[1] : []).map(prepareNblUnileverRow_),
    };
    _nblListsCacheAt = now;
    return _nblListsCache;
  }

  function runNblMatchCheck_(primary, lists) {
    var matches = [];
    var seen = {};
    var group = primary.group;
    var company = primary.company;
    var mill = primary.mill;

    function pushMatch_(key, entry) {
      if (seen[key]) return;
      seen[key] = true;
      matches.push(entry);
    }

    // NBL: any single field match counts (Group OR Company)
    if (group || company) {
      lists.registry.forEach(function(r, i) {
        var groupMatch = !!(group && nblNamesEqual_(r._nblGroup, group));
        var companyMatch = !!(company && nblNamesEqual_(r._nblCompany, company));
        if (!groupMatch && !companyMatch) return;
        var hit = [];
        if (groupMatch) hit.push('Group Name');
        if (companyMatch) hit.push('Company Name');
        pushMatch_('nbl-' + i, {
          source: 'NBL',
          detail: 'Matched (' + hit.join(' or ') + '): Group ' + (r._nblGroup || '—')
            + ' · Company ' + (r._nblCompany || '—')
            + (r._nblSource ? ' · Source: ' + r._nblSource : ''),
        });
      });
    }

    // Unilever NBL: any single field match counts (Company OR Mill)
    if (company || mill) {
      lists.unilever.forEach(function(r, i) {
        var companyMatch = !!(company && nblNamesEqual_(r._nblCompany, company));
        var millMatch = !!(mill && nblNamesEqual_(r._nblMill, mill));
        if (!companyMatch && !millMatch) return;
        var hit = [];
        if (companyMatch) hit.push('Company Name');
        if (millMatch) hit.push('Mill Name');
        pushMatch_('uni-' + i, {
          source: 'Unilever NBL',
          detail: 'Matched (' + hit.join(' or ') + '): Company ' + (r._nblCompany || '—')
            + ' · Mill ' + (r._nblMill || '—')
            + (r._nblUml ? ' · UML: ' + r._nblUml : ''),
        });
      });
    }

    var status = matches.length ? 'Yes' : 'No';
    return { status: status, matches: matches, checkedAt: new Date().toISOString(), primary: primary };
  }

  function renderSddNblCheckBanner_(result) {
    var boxes = document.querySelectorAll('#sdd-nbl-check-result');
    if (!boxes.length) return;
    var isYes = result.status === 'Yes';
    var bg = isYes ? 'rgba(185,28,28,0.08)' : 'rgba(30,107,58,0.08)';
    var border = isYes ? 'rgba(185,28,28,0.35)' : 'rgba(30,107,58,0.35)';
    var title = isYes ? 'Yes — on No Buy List' : 'No — not on No Buy List';
    var html = '<strong style="color:#1A0A0A;">' + title + '</strong>';
    html += '<div style="margin-top:6px;color:#5F4A48;">Compared: Group <em>' + escHtml(result.primary.group || '—')
      + '</em> · Company <em>' + escHtml(result.primary.company || '—')
      + '</em> · Mill <em>' + escHtml(result.primary.mill || '—') + '</em></div>';
    if (result.matches.length) {
      html += '<ul style="margin:8px 0 0;padding-left:18px;color:#4a1c1c;">';
      result.matches.forEach(function(m) {
        html += '<li style="margin-bottom:4px;"><span style="font-weight:600;">' + escHtml(m.source) + ':</span> '
          + escHtml(m.detail) + '</li>';
      });
      html += '</ul>';
    } else {
      html += '<div style="margin-top:6px;">No similar Group Name, Company Name, or Mill Name found in NBL or Unilever NBL sheets.</div>';
    }
    html += '<div style="margin-top:8px;font-size:11px;color:#9C8080;">Screening field <strong>No Buy List</strong> set to '
      + escHtml(result.status) + '. Included in PDF export.</div>';
    boxes.forEach(function(box) {
      box.style.display = 'block';
      box.style.background = bg;
      box.style.border = '1px solid ' + border;
      box.style.borderRadius = '8px';
      box.style.padding = '10px 12px';
      box.innerHTML = html;
    });
  }

  function persistNblCheckToSdd_(result) {
    var isYes = result.status === 'Yes';
    var statusLabel = isYes
      ? 'YES — Supplier IS ON the No Buy List (NBL)'
      : 'NO — Supplier is NOT on the No Buy List';
    var detail = result.matches.length
      ? result.matches.map(function(m) { return m.source + ': ' + m.detail; }).join(' | ')
      : 'No matching Group Name, Company Name, or Mill Name in NBL or Unilever NBL sheets.';
    window._scrData = window._scrData || {};
    window._scrData.nblCheckResult = statusLabel;
    window._scrData.nblCheckDetail = detail;
    window._scrData.nblCheckedAt = result.checkedAt || new Date().toISOString();
    if (window._loadedPrimarySddRow) {
      window._loadedPrimarySddRow['SCR - NBL Check Result'] = statusLabel;
      window._loadedPrimarySddRow['SCR - NBL Match Detail'] = detail;
      window._loadedPrimarySddRow['SCR - NBL Checked At'] = window._scrData.nblCheckedAt;
    }
  }

  function applyNblCheckToScrForm_(status) {
    var val = status === 'Yes' ? 'Yes' : 'No';
    var el = document.getElementById('scr-nbl');
    if (el) {
      el.value = val;
      el.classList.remove('scr-sel-yes', 'scr-sel-no');
      if (val === 'Yes') el.classList.add('scr-sel-yes');
      else if (val === 'No') el.classList.add('scr-sel-no');
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }
    window._scrData = window._scrData || {};
    window._scrData.nbl = val;
    if (window._loadedPrimarySddRow) {
      window._loadedPrimarySddRow['SCR - No Buy List'] = val;
    }
  }

  function restoreNblCheckResultFromRow_(row) {
    if (!row || typeof row !== 'object') return;
    var stored = String(row['SCR - NBL Check Result'] || '').trim();
    var detail = String(row['SCR - NBL Match Detail'] || '').trim();
    var checkedAt = String(row['SCR - NBL Checked At'] || '').trim();
    var nbl = String(row['SCR - No Buy List'] || '').trim();
    if (!stored && !detail && !nbl) return;
    var status = /^yes/i.test(nbl) || /on the no buy list/i.test(stored) ? 'Yes' : 'No';
    if (stored && /not on the no buy list/i.test(stored)) status = 'No';
    var matches = [];
    if (detail && !/^no matching/i.test(detail)) {
      detail.split(/\s*\|\s*/).forEach(function(part) {
        var idx = part.indexOf(':');
        if (idx > -1) {
          matches.push({ source: part.slice(0, idx).trim(), detail: part.slice(idx + 1).trim() });
        }
      });
    }
    window._nblCheckResult = {
      status: status,
      matches: matches,
      checkedAt: checkedAt || '',
      primary: getSddPrimaryForNblCheck_(),
    };
    window._scrData = window._scrData || {};
    window._scrData.nblCheckResult = stored;
    window._scrData.nblCheckDetail = detail;
    window._scrData.nblCheckedAt = checkedAt;
  }
  window.restoreNblCheckResultFromRow_ = restoreNblCheckResultFromRow_;

  window.runSddNblCheck = async function() {
    var primary = getSddPrimaryForNblCheck_();
    if (!primary.company && !primary.group && !primary.mill) {
      if (typeof window.showSddToast === 'function') {
        window.showSddToast('Import Excel dulu — Group Name, Company Name, dan Mill Name diperlukan untuk Check NBL.', 'error');
      }
      return;
    }

    var btns = document.querySelectorAll('#sdd-check-nbl-btn, #sdd-check-nbl-btn-trace');
    btns.forEach(function(b) { b.disabled = true; });
    var boxes = document.querySelectorAll('#sdd-nbl-check-result');
    boxes.forEach(function(box) {
      box.style.display = 'block';
      box.style.background = 'rgba(249,250,251,0.9)';
      box.style.border = '1px solid rgba(74,28,28,0.12)';
      box.style.borderRadius = '8px';
      box.style.padding = '10px 12px';
      box.textContent = 'Checking against NBL sheets…';
    });

    try {
      var lists = await ensureNblListsForCheck_();
      var result = runNblMatchCheck_(primary, lists);
      window._nblCheckResult = result;
      applyNblCheckToScrForm_(result.status);
      persistNblCheckToSdd_(result);
      renderSddNblCheckBanner_(result);
      if (typeof window.showSddToast === 'function') {
        window.showSddToast('NBL check: ' + result.status + (result.matches.length
          ? ' (' + result.matches.length + ' match' + (result.matches.length > 1 ? 'es' : '') + ')'
          : ''), result.status === 'Yes' ? 'warning' : 'success');
      }
    } catch (err) {
      var msg = (err && err.message) ? err.message : String(err);
      boxes.forEach(function(box) {
        box.style.display = 'block';
        box.style.background = 'rgba(185,28,28,0.06)';
        box.style.border = '1px solid rgba(185,28,28,0.25)';
        box.textContent = 'Check NBL failed: ' + msg;
      });
      if (typeof window.showSddToast === 'function') {
        window.showSddToast('Check NBL failed: ' + msg, 'error');
      }
    } finally {
      btns.forEach(function(b) { b.disabled = false; });
    }
  };

  // ─── NAVIGATION ─────────────────────────────────────────
  const pageEls = Array.from(document.querySelectorAll('.page'));
  pageEls.forEach(function(p) {
    p.setAttribute('aria-hidden', p.classList.contains('active') ? 'false' : 'true');
  });
  const panelEls = Array.from(document.querySelectorAll('.panel'));
  const navPanelEls = Array.from(document.querySelectorAll('.nav-item[data-panel]'));
  const filterChipEls = Array.from(document.querySelectorAll('.filter-chip'));
  const runWhenIdle = function(cb, timeoutMs) {
    if (typeof window.requestIdleCallback === 'function') {
      window.requestIdleCallback(cb, { timeout: timeoutMs || 1200 });
    } else {
      setTimeout(cb, timeoutMs || 1200);
    }
  };

  function resetScrollToTopEverywhere() {
    const run = function() {
      const mainContent = document.querySelector('.main-content');
      if (mainContent) {
        mainContent.scrollTop = 0;
        mainContent.scrollLeft = 0;
      }
      const activePage = document.querySelector('.page.active');
      if (activePage) activePage.scrollTop = 0;
      const activePanel = document.querySelector('.panel.active');
      if (activePanel) activePanel.scrollTop = 0;
      document.documentElement.scrollTop = 0;
      document.body.scrollTop = 0;
      window.scrollTo(0, 0);
    };
    run();
    requestAnimationFrame(run);
  }

  function showPage(id) {
    pageEls.forEach(function(p) {
      p.classList.remove('active');
      p.style.removeProperty('display');
      p.setAttribute('aria-hidden', 'true');
    });
    const target = document.getElementById(id);
    if (target) {
      target.classList.add('active');
      target.setAttribute('aria-hidden', 'false');
    }
    const dash = document.getElementById('dashboard');
    const loginEl = document.getElementById('login');
    if (dash) dash.inert = id !== 'dashboard';
    if (loginEl) loginEl.inert = id !== 'login';
    resetScrollToTopEverywhere();
  }

  (function syncInitialPageInert_() {
    var active = document.querySelector('.page.active');
    var aid = active && active.id;
    var dash = document.getElementById('dashboard');
    var lo = document.getElementById('login');
    if (dash && aid) dash.inert = aid !== 'dashboard';
    if (lo && aid) lo.inert = aid !== 'login';
  })();

  function switchPanel(name) {
    panelEls.forEach(function(p) { p.classList.remove('active'); });
    navPanelEls.forEach(function(n) { n.classList.remove('active'); });
    const panel = document.getElementById('panel-' + name);
    const navItem = document.querySelector('[data-panel="' + name + '"]');
    if (panel) panel.classList.add('active');
    if (navItem) navItem.classList.add('active');
    const grp = document.getElementById('navGroupTrace');
    if (grp && !grp.classList.contains('open')) grp.classList.add('open');
    const grpPrograms = document.getElementById('navGroupPrograms');
    if (grpPrograms && ['no-buy-list', 'performa-facility', 'eudr-potential', 'contact-list-supplier', 'priority-supplier-engagement'].indexOf(name) !== -1) {
      if (!grpPrograms.classList.contains('open')) grpPrograms.classList.add('open');
    }
    if (name === 'mill-onboarding' && allData.length) {
      if (currentFilter === 'Task List') {
        currentFilter = 'All';
        filterChipEls.forEach(function(c) { c.classList.toggle('active', c.dataset.filter === 'All'); });
        const taskPanel = document.getElementById('mill-task-list-panel');
        const tableCard = document.querySelector('#panel-mill-onboarding .table-card');
        if (taskPanel) taskPanel.style.display = 'none';
        if (tableCard) tableCard.style.display = '';
      }
      scheduleRenderMillTable();
    }
    if (name === 'ttm-ttp' && !ttpLoaded) loadTTPData();
    if (name === 'grievance' && !grvLoaded) loadGrvData();
    if (name === 'contact-list-supplier') loadContactListData();
    if (name === 'no-buy-list') loadNoBuyListData();
    if (name === 'performa-facility') initPerformaFacility_();
    resetScrollToTopEverywhere();
  }
  // expose globally for onclick handlers
  window.switchPanel = switchPanel;

  // ─── PERFORMA FACILITY ────────────────────────────────────────────────────────
  (function setupPerformaFacility_() {
    function pfPopulateYears_() {
      const sel = document.getElementById('pfYearSel');
      if (!sel) return;
      const existing = Array.from(sel.options).map(function(o) { return o.value; });
      const years = new Set();
      (allData || []).forEach(function(r) {
        const y = String(millYearVal(r) || '').trim();
        if (y) years.add(y);
      });
      Array.from(years).sort(function(a, b) { return Number(b) - Number(a); }).forEach(function(y) {
        if (existing.indexOf(y) === -1) {
          const opt = document.createElement('option');
          opt.value = y;
          opt.textContent = y;
          sel.appendChild(opt);
        }
      });
    }

    function pfSplitFacility_(raw) {
      const s = String(raw || '').trim();
      if (!s) return ['—'];
      return s.split(/[,;/]+/).map(function(f) { return f.trim(); }).filter(Boolean);
    }

    function pfBuildRows_() {
      const qSel = document.getElementById('pfQuarterSel');
      const ySel = document.getElementById('pfYearSel');
      const q = qSel ? qSel.value : '';
      const y = ySel ? ySel.value : '';

      // Filter mill data by Quarter/Year
      const millRows = (allData || []).filter(function(r) {
        if (q && String(millQuarterVal(r) || '').trim() !== q) return false;
        if (y && String(millYearVal(r) || '').trim() !== y) return false;
        return true;
      });

      // Build TTP lookup: company name (lowercase) → avg % CPO TRACEABLE
      const ttpLookup = {};
      const ttpColKey = ttpPctCol || '% CPO TRACEABLE';
      (ttpData || []).forEach(function(r) {
        const co = String(r['COMPANY NAME'] || r['Company Name'] || '').trim().toLowerCase();
        if (!co) return;
        const raw = r[ttpColKey] || r['% CPO TRACEABLE'] || r['% Cpo Traceable'] || '';
        const num = parseFloat(String(raw).replace(',', '.').replace('%', ''));
        if (!isNaN(num)) {
          if (!ttpLookup[co]) ttpLookup[co] = { sum: 0, count: 0 };
          ttpLookup[co].sum += num;
          ttpLookup[co].count += 1;
        }
      });

      // Expand mill rows by facility
      const pfRows = [];
      millRows.forEach(function(r) {
        const facilities = pfSplitFacility_(r['FACILITY NAME CPO']);
        const coLower = String(r['COMPANY NAME'] || '').trim().toLowerCase();
        const ttpEntry = ttpLookup[coLower];
        const cpoPct = ttpEntry
          ? (ttpEntry.sum / ttpEntry.count).toFixed(1) + '%'
          : '—';

        facilities.forEach(function(fac) {
          pfRows.push({
            company:   String(r['COMPANY NAME']   || '').trim() || '—',
            group:     String(r['GROUP NAME']     || '').trim() || '—',
            facility:  fac,
            province:  String(r['PROVINCE']       || '').trim() || '—',
            nbl:       String(r['BUYER NO BUY LIST'] || '').trim() || '—',
            riskLevel: String(r['RESULT RISK LEVEL'] || '').trim() || '—',
            grievance: r['TOTAL GRIEVANCES'] != null ? r['TOTAL GRIEVANCES'] : '—',
            cpoPct:    cpoPct,
          });
        });
      });

      return pfRows;
    }

    function pfRiskBadge_(v) {
      const s = String(v || '').toLowerCase();
      if (s.includes('high'))   return '<span class="status-badge risk-high"><span class="s-dot"></span>' + escHtml(v) + '</span>';
      if (s.includes('med'))    return '<span class="status-badge risk-med"><span class="s-dot"></span>'  + escHtml(v) + '</span>';
      if (s.includes('low'))    return '<span class="status-badge risk-low"><span class="s-dot"></span>'  + escHtml(v) + '</span>';
      return '<span>' + escHtml(v) + '</span>';
    }

    function pfNblBadge_(v) {
      const s = String(v || '').toLowerCase();
      const isYes = s === 'yes' || s.includes('nbl') || s.includes('no buy');
      if (isYes) return '<span class="status-badge risk-high"><span class="s-dot"></span>Yes</span>';
      if (s === 'no') return '<span class="status-badge risk-low"><span class="s-dot"></span>No</span>';
      return '<span>' + escHtml(v) + '</span>';
    }

    function pfRenderTable_(rows) {
      const tbl   = document.getElementById('pfTable');
      const tbody = document.getElementById('pfTableBody');
      const loading = document.getElementById('pf-loading');
      const scope = document.getElementById('pfScopeText');
      if (!tbl || !tbody) return;

      const q = String((document.getElementById('pfSearch') || {}).value || '').toLowerCase().trim();
      const filtered = q
        ? rows.filter(function(r) {
            return [r.company, r.group, r.facility, r.province].some(function(v) {
              return v.toLowerCase().includes(q);
            });
          })
        : rows;

      // Stats
      const totalEl   = document.getElementById('pf-stat-total');
      const highEl    = document.getElementById('pf-stat-high');
      const nblEl     = document.getElementById('pf-stat-nbl');
      const grvEl     = document.getElementById('pf-stat-grievance');
      if (totalEl) totalEl.textContent = String(filtered.length);
      if (highEl)  highEl.textContent  = String(filtered.filter(function(r) { return r.riskLevel.toLowerCase().includes('high'); }).length);
      if (nblEl)   nblEl.textContent   = String(filtered.filter(function(r) { const s = r.nbl.toLowerCase(); return s === 'yes' || s.includes('nbl'); }).length);
      if (grvEl)   grvEl.textContent   = String(filtered.filter(function(r) { const g = parseFloat(String(r.grievance)); return !isNaN(g) && g > 0; }).length);
      if (scope)   scope.textContent   = filtered.length + ' rows';

      if (!filtered.length) {
        tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:#9C8A8A;">'
          + (q ? 'Tidak ada hasil pencarian.' : 'Tidak ada data. Coba ubah filter Quarter / Year.')
          + '</td></tr>';
        if (loading) loading.style.display = 'none';
        tbl.style.display = 'table';
        return;
      }

      tbody.innerHTML = filtered.map(function(r) {
        return '<tr>'
          + '<td><span class="mill-name">' + escHtml(r.company) + '</span></td>'
          + '<td>' + escHtml(r.group) + '</td>'
          + '<td><strong>' + escHtml(r.facility) + '</strong></td>'
          + '<td>' + escHtml(r.province) + '</td>'
          + '<td>' + pfNblBadge_(r.nbl) + '</td>'
          + '<td>' + pfRiskBadge_(r.riskLevel) + '</td>'
          + '<td style="text-align:right;">' + escHtml(String(r.grievance)) + '</td>'
          + '<td style="text-align:right;">' + escHtml(r.cpoPct) + '</td>'
          + '</tr>';
      }).join('');

      if (loading) loading.style.display = 'none';
      tbl.style.display = 'table';
    }

    let _pfCurrentRows = [];

    window.initPerformaFacility_ = function initPerformaFacility_() {
      pfPopulateYears_();

      const applyBtn = document.getElementById('pfApplyBtn');
      const searchEl = document.getElementById('pfSearch');

      if (applyBtn && !applyBtn._pfBound) {
        applyBtn._pfBound = true;
        applyBtn.addEventListener('click', function() {
          const loading = document.getElementById('pf-loading');
          const tbl     = document.getElementById('pfTable');
          if (loading) { loading.style.display = 'block'; loading.textContent = 'Memuat…'; }
          if (tbl)     tbl.style.display = 'none';

          // Ensure both data sources are loaded first
          Promise.all([
            allData && allData.length ? Promise.resolve() : loadMillData(),
            ttpLoaded ? Promise.resolve() : loadTTPData(),
          ]).then(function() {
            pfPopulateYears_();
            _pfCurrentRows = pfBuildRows_();
            pfRenderTable_(_pfCurrentRows);
          }).catch(function(err) {
            const errEl = document.getElementById('pf-error');
            if (loading) loading.style.display = 'none';
            if (errEl) { errEl.style.display = 'block'; errEl.textContent = 'Gagal memuat data: ' + (err && err.message ? err.message : err); }
          });
        });
      }

      if (searchEl && !searchEl._pfBound) {
        searchEl._pfBound = true;
        searchEl.addEventListener('input', function() {
          pfRenderTable_(_pfCurrentRows);
        });
      }

      // If data already loaded, auto-apply
      if (allData && allData.length) {
        pfPopulateYears_();
        _pfCurrentRows = pfBuildRows_();
        pfRenderTable_(_pfCurrentRows);
      }
    };
  })();
  // ─── END PERFORMA FACILITY ────────────────────────────────────────────────────

  // ─── LOGIN (simple page, no modal needed) ─────────────────────────────────────────
  // focus email on load
  setTimeout(() => { const e = document.getElementById('loginEmail'); if (e) e.focus(); }, 100);

  // ─── LOGIN / LOGOUT ─────────────────────────────────────
  function resolveSddRoleFromSupabaseUser_(user) {
    if (!user) return 'STAFF';
    // Approver role retired — all accounts use the same SDD workflow (decision on staff UI).
    return 'STAFF';
  }

  async function finalizeSuccessfulLogin_(displayEmail, role) {
    window._sddUserRole = 'STAFF';
    document.body.classList.remove('sdd-role-approver');
    if (typeof window.refreshSddApproverStaffTraceChrome === 'function') window.refreshSddApproverStaffTraceChrome();
    if (typeof updateSavedScreeningPickerUI === 'function') updateSavedScreeningPickerUI();
    var err = document.getElementById('loginErr');
    if (err) err.style.display = 'none';
    var dashUser = document.getElementById('dashUserName');
    var avatar = document.getElementById('userAvatar');
    if (dashUser) dashUser.textContent = displayEmail;
    if (avatar) avatar.textContent = (displayEmail || '?').charAt(0).toUpperCase();
    if (typeof window.updateOverviewWelcome === 'function') window.updateOverviewWelcome(displayEmail);
    showPage('dashboard');
    if (typeof switchPanel === 'function') switchPanel('overview');
    var sbNav = document.getElementById('mainSidebar');
    if (sbNav) sbNav.classList.remove('expanded');
    try {
      await loadMillData();
      console.log('✅ Dashboard loaded successfully');
    } catch (error) {
      console.error('❌ Error loading dashboard:', error);
    }
    runWhenIdle(function() {
      if (!ttpLoaded) loadTTPData();
      if (!grvLoaded) loadGrvData();
    }, 1400);
  }

  async function doLogin() {
    const emailEl = document.getElementById('loginEmail');
    const passEl = document.getElementById('loginPass');
    if (!emailEl || !passEl) {
      console.warn('[dashboard] Login fields missing.');
      return;
    }
    const emailRaw = emailEl.value.trim();
    const emailNorm = emailRaw.toLowerCase().replace(/\s+/g, '');
    const pass = String(passEl.value || '');
    const passNorm = pass.trim();
    const passCompact = passNorm.replace(/\s+/g, '');
    const err = document.getElementById('loginErr');
    if (!emailNorm || !passCompact) {
      if (err) {
        err.textContent = 'Invalid credentials. Please try again.';
        err.style.display = 'block';
      }
      return;
    }

    const sb = getSupabase();
    if (sb) {
      const { data, error } = await sb.auth.signInWithPassword({
        email: emailRaw,
        password: passNorm,
      });
      if (error) {
        if (err) {
          err.textContent = error.message || 'Sign in failed.';
          err.style.display = 'block';
        }
        return;
      }
      const u = data.user;
      const role = resolveSddRoleFromSupabaseUser_(u);
      await finalizeSuccessfulLogin_(u.email || emailRaw, role);
      return;
    }

    // Demo account fallback (no Supabase env).
    const isStaffDemo =
      emailNorm === 'trace123@gmail.com' &&
      (passCompact === 'trace123!' || passCompact === 'trace123');
    if (isStaffDemo) {
      await finalizeSuccessfulLogin_(emailRaw, 'STAFF');
    } else {
      if (err) {
        err.textContent = 'Invalid credentials. Please try again.';
        err.style.display = 'block';
      }
    }
  }

  const btnLoginSubmit = document.getElementById('btn-login-submit');
  const loginPassEl = document.getElementById('loginPass');
  const btnLogout = document.getElementById('btn-logout');
  if (btnLoginSubmit && loginPassEl) {
    btnLoginSubmit.addEventListener('click', async () => {
      try {
        await doLogin();
      } catch (error) {
        console.error('Login error:', error);
      }
    });
    loginPassEl.addEventListener('keydown', async (e) => {
      if (e.key === 'Enter') {
        try {
          await doLogin();
        } catch (error) {
          console.error('Login error:', error);
        }
      }
    });
  } else {
    console.warn('[dashboard] Login form nodes missing (#btn-login-submit / #loginPass). Check entry mounts before main.js.');
  }
  if (btnLogout) btnLogout.addEventListener('click', async function() {
    var sbOut = getSupabase();
    if (sbOut) {
      try {
        await sbOut.auth.signOut();
      } catch (e) {
        console.warn('[dashboard] Supabase signOut:', e);
      }
    }
    clearClientSessionForLogout_();
  });

  function clearClientSessionForLogout_() {
    var le = document.getElementById('loginEmail');
    var lp = document.getElementById('loginPass');
    if (le) le.value = '';
    if (lp) lp.value = '';
    ttpLoaded = false;
    ttpData = [];
    allData = [];
    millPdfDimFilters = { quarter: new Set(), year: new Set(), group: new Set(), province: new Set() };
    millPdfColSelected = new Set(MILL_PDF_COL_DEFAULT_KEYS);
    millPdfRebuildDimPanels();
    millPdfRenderColPanel();
    updateMillPdfExportScope();
    currentFilter = 'All';
    currentSearch = '';
    grvLoaded = false;
    grvData = [];
    window._sddUserRole = 'STAFF';
    document.body.classList.remove('sdd-role-approver');
    if (typeof window.refreshSddApproverStaffTraceChrome === 'function') window.refreshSddApproverStaffTraceChrome();
    if (typeof updateSavedScreeningPickerUI === 'function') updateSavedScreeningPickerUI();
    window._sddApproverRecordLoaded = false;
    window._sddSubmissionId      = null;
    window._sddIsLoadedSaved     = false;
    window._sddLastInsertedRow   = null;
    window._scrLoadedRowNum      = null;
    window._scrLoadedKey         = '';
    window._loadedPrimarySddRow  = null;
    window._sddImportFirstRow    = null;
    window._sddImportedRows      = [];
    window._scrData              = {};
    window._tmlScreeningData     = {};
    window._ffbScreeningData     = {};
    window._scrSavedRowsByKey    = {};
    window._scrSavedGroupsByKey  = {};
    window._sddAllRowsCache      = null;
    supplierWorkbook             = null;
    if (typeof switchPanel === 'function') switchPanel('overview');
    showPage('login');
  }

  (async function restoreSupabaseSession_() {
    var sb = getSupabase();
    if (!sb) return;
    var res = await sb.auth.getSession();
    if (res.data.session && res.data.session.user) {
      var u = res.data.session.user;
      await finalizeSuccessfulLogin_(u.email || '', resolveSddRoleFromSupabaseUser_(u));
    }
    sb.auth.onAuthStateChange(function(event) {
      if (event === 'SIGNED_OUT') clearClientSessionForLogout_();
    });
  })();

  // ─── SIDEBAR ────────────────────────────────────────────
  navPanelEls.forEach(function(item) {
    item.addEventListener('click', function() {
      const panel = this.dataset.panel;
      switchPanel(panel);
      if (panel === 'grievance' && !grvLoaded) loadGrvData();
    });
  });

  const btnBlMonitoringBack = document.getElementById('btn-bl-monitoring-back');
  if (btnBlMonitoringBack) {
    btnBlMonitoringBack.addEventListener('click', function() {
      switchPanel('overview');
    });
  }
  document.querySelectorAll('.nav-soon-back').forEach(function(btn) {
    btn.addEventListener('click', function() {
      switchPanel('overview');
    });
  });

  // ─── SIDEBAR TOGGLE ─────────────────────────────────────
  const sidebar = document.getElementById('mainSidebar');
  const sidebarToggle = document.getElementById('sidebarToggle');
  if (sidebar && sidebarToggle) {
    sidebarToggle.addEventListener('click', function(e) {
      e.stopPropagation();
      sidebar.classList.toggle('expanded');
    });
  }

  // ─── TRACEABILITY GROUP TOGGLE ──────────────────────────
  const navGroupTraceHeader = document.getElementById('navGroupTraceHeader');
  const navGroupTrace = document.getElementById('navGroupTrace');
  if (navGroupTraceHeader && navGroupTrace) {
    navGroupTraceHeader.addEventListener('click', function() {
      navGroupTrace.classList.toggle('open');
    });
  }

  const navGroupProgramsHeader = document.getElementById('navGroupProgramsHeader');
  const navGroupPrograms = document.getElementById('navGroupPrograms');
  if (navGroupProgramsHeader && navGroupPrograms) {
    navGroupProgramsHeader.addEventListener('click', function() {
      navGroupPrograms.classList.toggle('open');
    });
  }

  // ─── FILTERS ────────────────────────────────────────────
  filterChipEls.forEach(function(chip) {
    chip.addEventListener('click', function() {
      currentFilter = this.dataset.filter;
      filterChipEls.forEach(function(c) { c.classList.remove('active'); });
      this.classList.add('active');

      const taskPanel = document.getElementById('mill-task-list-panel');
      const tableCard = document.querySelector('#panel-mill-onboarding .table-card');

      if (currentFilter === 'Task List') {
        if (taskPanel) taskPanel.style.display = 'block';
        if (tableCard) tableCard.style.display = 'none';
        renderMillTaskList();
        updateMillPdfExportScope();
      } else {
        if (taskPanel) taskPanel.style.display = 'none';
        if (tableCard) tableCard.style.display = '';
        scheduleRenderMillTable();
      }
    });
  });

  // ─── TASK LIST ───────────────────────────────────────────
  function escapeHtmlMin(s) {
    return String(s === undefined || s === null ? '' : s)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function renderMillTaskList() {
    const body    = document.getElementById('mill-task-list-body');
    const empty   = document.getElementById('mill-task-empty');
    const countEl = document.getElementById('mill-task-count');
    if (!body) return;

    const rows = window._scrSavedRowsByKey || {};
    const submitted = Object.entries(rows)
      .filter(function([, r]) {
        const status    = String(r['SCR - Screening Status'] || '').toLowerCase();
        const decision  = String(
          r['statusSDD'] || r['statusSdd'] || r['Status SDD'] ||
          r['statusBossDecision'] || r['Status Boss Decision'] || ''
        ).trim().toLowerCase();
        const millAdded = String(r['mill_added'] || '').toLowerCase();
        return status === 'submitted' && (decision === 'approve' || decision === 'approved') && millAdded !== 'true';
      })
      .sort(function([, a], [, b]) {
        return new Date(b['updated_at'] || 0) - new Date(a['updated_at'] || 0);
      });

    if (countEl) countEl.textContent = submitted.length + ' entr' + (submitted.length === 1 ? 'y' : 'ies');

    if (!submitted.length) {
      body.innerHTML = '';
      if (empty) empty.style.display = 'block';
      return;
    }
    if (empty) empty.style.display = 'none';

    body.innerHTML = submitted.map(function([key, r]) {
      const grp      = String(r['Group Name'] || r['Grup Name'] || '—').trim();
      const mill     = String(r['Mill Name'] || '').trim();
      const tp       = String(r['Supplier Type'] || r['supplier_type'] || '').trim().toUpperCase();
      const updRaw   = String(r['updated_at'] || r['SCR - Last Updated'] || '').trim();
      const updDate  = updRaw ? new Date(updRaw).toLocaleDateString('id-ID', { day:'2-digit', month:'short', year:'numeric' }) : '';
      const metaParts = [tp, mill, updDate ? 'Submitted ' + updDate : ''].filter(Boolean);

      return '<div class="mill-task-card">'
        + '<div class="mill-task-card-icon">'
        + '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2.2" stroke-linecap="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>'
        + '</div>'
        + '<div class="mill-task-card-info">'
        + '<div class="mill-task-card-group">' + escapeHtmlMin(grp) + '</div>'
        + (metaParts.length ? '<div class="mill-task-card-meta">' + metaParts.map(escapeHtmlMin).join(' &nbsp;·&nbsp; ') + '</div>' : '')
        + '</div>'
        + '<button type="button" class="mill-task-card-btn"'
        + ' data-task-key="' + escapeHtmlMin(key) + '"'
        + ' data-task-group="' + escapeHtmlMin(grp) + '">+ Add to Mill</button>'
        + '</div>';
    }).join('');
  }

  /**
   * Pure mapping: SDD_MAIN flat row → partial payload for Mill modal fields.
   *
   * COORDINATES format: "lat; long"  (semicolon + space separator).
   * Rationale: comma is already the decimal separator for Indonesian coordinate
   * notation (e.g. "0,318318"), so using comma as a field separator would be
   * ambiguous. Semicolon is unambiguous. Preserve original decimal notation of
   * each component (do not re-normalise here — normalizeCoordinate already ran
   * at import time). If only one component is present, no separator is emitted.
   * Both empty → COORDINATES stays ''.
   *
   * CAPACITY field: mapped to 'MILL CAPACITY (TON/HOUR)' in the modal.
   * Sources tried in order: 'Mill Capacity (Ton/Hour)', 'KCP Capacity (Ton/Hour)'.
   * This requires the column 'MILL CAPACITY (TON/HOUR)' to exist in the
   * Mill Onboarding Profile sheet — add it once manually next to MILL CATEGORY.
   *
   * @param {Object} r - flat SDD_MAIN row object from window._scrSavedRowsByKey
   * @returns {Object} partial payload keyed by Mill modal data-field attribute values
   */
  function mapSddRowToMillPayload(r) {
    if (!r || typeof r !== 'object') return {};

    // ── Coordinates ───────────────────────────────────────────────────────────
    // Values are stored as dot-decimal strings (e.g. "0.318318").
    // toDot: replace ALL commas with dots in case a legacy comma-decimal value
    // slips through (replaceAll covers "101,383,8183" edge-cases too).
    // Separator changed to ", " (koma-spasi) per product requirement.
    const toDot  = v => v.split(',').join('.');
    const lat    = toDot(String(r['Latitude']  || '').trim());
    const lng    = toDot(String(r['Longitude'] || '').trim());
    const coords = [lat, lng].filter(Boolean).join(', ');

    // ── Capacity: Mill first, KCP fallback ────────────────────────────────────
    const capacity = String(
      r['Mill Capacity (Ton/Hour)'] || r['KCP Capacity (Ton/Hour)'] || ''
    ).trim();

    // ── Mill/KCP name ─────────────────────────────────────────────────────────
    const millName = String(r['Mill Name'] || r['KCP Name'] || '').trim();

    // ── Address: Mill first, KCP fallback ─────────────────────────────────────
    const address = String(r['Mill Address'] || r['KCP Address'] || '').trim();

    // ── Mill Category: Mill first, KCP fallback ───────────────────────────────
    const millCat = String(r['Mill Category'] || r['KCP Category'] || '').trim();

    return {
      'GROUP NAME':               String(r['Group Name'] || r['Grup Name'] || '').trim(),
      'COMPANY NAME':             String(r['Company Name'] || '').trim(),
      'MILL NAME':                millName,
      'ADDRESS':                  address,
      'COORDINATES':              coords,
      'MILL CATEGORY':            millCat,
      'MILL CAPACITY (TON/HOUR)': capacity,
    };
  }

  (function bindMillTaskListDelegationIfPresent() {
    const body = document.getElementById('mill-task-list-body');
    if (!body) return;
    body.addEventListener('click', function(e) {
      const btn = e.target.closest('.mill-task-card-btn');
      if (!btn || !body.contains(btn)) return;

      const taskKey = btn.dataset.taskKey || '';
      modalTaskKey  = taskKey;

      // Resolve cached SDD row for pre-fill (guard: key may be absent from cache)
      const sddRow  = (taskKey && window._scrSavedRowsByKey)
                      ? (window._scrSavedRowsByKey[taskKey] || null)
                      : null;
      const prefill = sddRow ? mapSddRowToMillPayload(sddRow) : {};

      // Derive modal title: prefer mill name, fall back to group, then taskGroup attr
      const titleName = prefill['MILL NAME'] || prefill['GROUP NAME']
                        || btn.dataset.taskGroup || '';

      // Open modal with pre-filled data; buildMillForm reads data[f] for each field
      openModal('mill', MILL_FIELDS, 'add', prefill);

      // After DOM settles: patch title + fire input events on pre-filled text inputs
      requestAnimationFrame(function() {
        const titleEl = document.getElementById('modalTitle');
        if (titleEl) titleEl.textContent = 'Add Mill \u2014 ' + titleName;

        // Fire 'input' on every pre-filled text input so any live listeners pick up values
        document.querySelectorAll('#modalFormGrid [data-field]').forEach(function(el) {
          const field = el.dataset.field;
          if (prefill[field] !== undefined && el.tagName === 'INPUT') {
            el.dispatchEvent(new Event('input', { bubbles: true }));
          }
        });

        // Re-run calcTotals; YESNO fields were not pre-filled so score stays 0
        calcTotals();
      });
    });
  })();

  // ─── SEARCH ─────────────────────────────────────────────
  const millSearch = document.getElementById('millSearch');
  const millSearchClear = document.getElementById('millSearchClear');
  if (millSearch && millSearchClear) {
    const debouncedRenderMillTable = debounce(function() {
      scheduleRenderMillTable();
    }, 120);
    millSearch.addEventListener('input', function() {
      currentSearch = this.value.toLowerCase().trim();
      if (this.value) {
        millSearchClear.classList.add('show');
      } else {
        millSearchClear.classList.remove('show');
      }
      debouncedRenderMillTable();
    });

    millSearchClear.addEventListener('click', function() {
      const searchInput = millSearch;
      searchInput.value = '';
      currentSearch = '';
      this.classList.remove('show');
      debouncedRenderMillTable.cancel();
      scheduleRenderMillTable.flush();
      searchInput.focus();
    });
  }

  // ─── SDD PDF EXPORT ─────────────────────────────────────────────────────────
  async function sddExportPdf() {
    var status = String(
      (window._scrData && window._scrData.status) ||
      (window._loadedPrimarySddRow && window._loadedPrimarySddRow['SCR - Screening Status']) || ''
    ).trim().toLowerCase();
    if (!status || (status !== 'submitted' && status !== 'draft')) {
      if (typeof window.showSddToast === 'function') {
        window.showSddToast('Load screening record terlebih dahulu sebelum export PDF.', 'error');
      }
      return;
    }

    var pdfBtn = document.getElementById('sdd-export-pdf-btn');
    if (pdfBtn) { pdfBtn.disabled = true; pdfBtn.innerHTML = 'Generating…'; }

    try {
      if (window._loadedPrimarySddRow && typeof window.restoreNblCheckResultFromRow_ === 'function') {
        window.restoreNblCheckResultFromRow_(window._loadedPrimarySddRow);
      }

      var jsPDFLib = getJsPDF();
      var doc = new jsPDFLib({ orientation: 'portrait', unit: 'mm', format: 'a4' });

      // ─── Data sources ────────────────────────────────────────────────────
      var p  = window._loadedPrimarySddRow || {};
      var sc = window._scrData || {};

      // Pick first non-empty value from a list of field aliases in p
      function pick() {
        var keys = Array.prototype.slice.call(arguments);
        for (var i = 0; i < keys.length; i++) {
          var v = String(p[keys[i]] || '').trim();
          if (v) return v;
        }
        return '';
      }

      // DOM → _scrData → API field fallback
      function val(domId, scrKey, apiKeys) {
        var el = domId ? document.getElementById(domId) : null;
        if (el && String(el.value || '').trim()) return el.value.trim();
        if (scrKey && String(sc[scrKey] || '').trim()) return String(sc[scrKey]).trim();
        if (apiKeys) {
          var arr = Array.isArray(apiKeys) ? apiKeys : [apiKeys];
          for (var i = 0; i < arr.length; i++) {
            if (String(p[arr[i]] || '').trim()) return String(p[arr[i]]).trim();
          }
        }
        return '';
      }

      // All screening fields
      var f = {
        owners      : val('scr-owners',   'owners',   'SCR - List Group By Owners'),
        news        : val('scr-news',      'news',     'SCR - Previous News'),
        supplyto    : val('scr-supplyto',  'supplyto', 'SCR - Supply To'),
        legality    : val('scr-legality',  'legality', 'SCR - Legality Status'),
        cert        : val('scr-cert',      'cert',     'SCR - Certification'),
        ndpe        : val('scr-ndpe',      'ndpe',     'SCR - NDPE Policy'),
        nbl         : val('scr-nbl',       'nbl',      'SCR - No Buy List'),
        nblCheckResult: val(null, 'nblCheckResult', 'SCR - NBL Check Result'),
        nblCheckDetail: val(null, 'nblCheckDetail', 'SCR - NBL Match Detail'),
        nblCheckedAt  : val(null, 'nblCheckedAt',   'SCR - NBL Checked At'),
        grvYN       : val('scr-grv-yn',    'grvYN',    'SCR - Grievance (Y/N)'),
        priYN       : val('scr-pri-yn',    'priYN',    'SCR - PRI (Y/N)'),
        note        : val('traceRecInput', 'traceNote','SCR - Notes'),
        reqData     : val('requestedDataInput','requestedData',['SCR - Requested Data','Requested Data']),
        approverNote: val('noteBossDecision','noteBossDecision',['noteBossDecision','noteSDD','noteSdd']),
        approverStat: val(null,'statusSdd',['statusSDD','statusSdd','Status SDD','statusBossDecision','Status Boss Decision','StatusSDD','STATUSSDD']),
        status      : (sc.status || p['SCR - Screening Status'] || '').toUpperCase(),
        date        : (sc.date || p['SCR - Last Updated'] || p['updated_at'] || '')
      };

      var sid          = String(window._sddSubmissionId || window._scrLoadedKey || '').trim();
      var supplierType = String(window._sddSupplierType || p['Supplier Type'] || '').toUpperCase();
      var supplierName = String(p['Company Name'] || p['Mill Name'] || '').trim() || '—';
      var exportedAt   = new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });

      // Header "Status" = decision (Approve/Hold/Reject); else PENDING.
      var approverRaw = String(f.approverStat || '').trim();
      var metaStatusLine;
      if (approverRaw) {
        metaStatusLine = _normalizeDecisionLabel(approverRaw) || approverRaw.toUpperCase();
      } else {
        metaStatusLine = 'PENDING';
      }

      // Mill & FFB rows
      var cachedGroup = (window._scrSavedGroupsByKey && sid) ? window._scrSavedGroupsByKey[sid] : null;
      var millRows = (cachedGroup && Array.isArray(cachedGroup.mills))    ? cachedGroup.mills    : [];
      var ffbRows  = (cachedGroup && Array.isArray(cachedGroup.ffb_rows)) ? cachedGroup.ffb_rows : [];
      if (!millRows.length && !ffbRows.length && Array.isArray(window._sddImportedRows)) {
        window._sddImportedRows.forEach(function(r) {
          if (String(r['TML - Mill Name'] || r['TML - Company Name'] || '').trim()) millRows.push(r);
          else if (String(r['FFB - Supplier Name'] || '').trim()) ffbRows.push(r);
        });
      }

      // GRV / PRI rows with DOM fallback
      function readTblRows(tbodyId, cols) {
        var rows = [];
        var tb = document.getElementById(tbodyId);
        if (tb) {
          tb.querySelectorAll('tr').forEach(function(tr) {
            var inputs = tr.querySelectorAll('input');
            var obj = {};
            cols.forEach(function(c, i) { obj[c] = inputs[i] ? inputs[i].value : ''; });
            if (Object.values(obj).some(function(v) { return String(v).trim(); })) rows.push(obj);
          });
        }
        return rows;
      }

      var grvRows = ((sc.grvRows || []).filter(function(r) {
        return Object.values(r).some(function(v) { return String(v || '').trim(); });
      }));
      if (!grvRows.length) grvRows = readTblRows('scr-grv-tbody', ['source','desc','pub','date','status','attach']);

      var priRows = ((sc.priRows || []).filter(function(r) {
        return Object.values(r).some(function(v) { return String(v || '').trim(); });
      }));
      if (!priRows.length) priRows = readTblRows('scr-pri-tbody', ['company','desc','pub','date','attach','action']);

      // ─── Palette ─────────────────────────────────────────────────────────
      var RED     = [139, 26, 26];
      var RED_LT  = [245, 235, 235];
      var RED_MD  = [160, 40, 40];
      var GRY_LBL = [120, 100, 100];
      var GRY_DRK = [80, 80, 80];
      var BLACK   = [28, 28, 28];
      var WHITE   = [255, 255, 255];
      var ROW_ALT = [252, 248, 248];

      var pageW = doc.internal.pageSize.getWidth();
      var pageH = doc.internal.pageSize.getHeight();
      var mL = 14, mR = 14, cW = pageW - mL - mR;
      var y = 0;

      // ─── HEADER ──────────────────────────────────────────────────────────
      doc.setFillColor.apply(doc, RED);
      doc.rect(0, 0, pageW, 30, 'F');
      doc.setTextColor.apply(doc, WHITE);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(15);
      doc.text('Supplier Due Diligence — Screening Report', mL, 13);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.text('CONFIDENTIAL  ·  Status: ' + metaStatusLine + '  ·  Exported: ' + exportedAt, mL, 21);
      if (sid) {
        doc.setFont('helvetica', 'bold');
        doc.text('ID: ' + sid, pageW - mR, 21, { align: 'right' });
      }
      y = 38;

      // ─── LAYOUT HELPERS ──────────────────────────────────────────────────
      function checkPage(need) {
        if (y + (need || 20) > pageH - 14) { doc.addPage(); y = 16; }
      }

      function sectionTitle(title) {
        checkPage(14);
        doc.setFillColor.apply(doc, RED_LT);
        doc.rect(mL, y - 4.5, cW, 8, 'F');
        doc.setTextColor.apply(doc, RED);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(8.5);
        doc.text(title.toUpperCase(), mL + 2.5, y + 1);
        doc.setTextColor.apply(doc, BLACK);
        y += 8;
      }

      var KV_LW = 52, KV_GAP = 3, KV_LH = 5;

      function kv(label, value) {
        checkPage(KV_LH + 2);
        var v = String(value === undefined || value === null || value === '' ? '—' : value) || '—';
        doc.setFont('helvetica', 'bold'); doc.setFontSize(7.8); doc.setTextColor.apply(doc, GRY_LBL);
        doc.text(label + ' :', mL, y);
        doc.setFont('helvetica', 'normal'); doc.setTextColor.apply(doc, BLACK);
        var lines = doc.splitTextToSize(v, cW - KV_LW - KV_GAP);
        doc.text(lines, mL + KV_LW + KV_GAP, y);
        y += Math.max(KV_LH, lines.length * KV_LH);
      }

      function kv2(l1, v1, l2, v2) {
        checkPage(KV_LH + 2);
        var half = cW / 2 - 3, xR = mL + cW / 2 + 3;
        var s1 = String(v1 === undefined || v1 === null || v1 === '' ? '—' : v1) || '—';
        var s2 = String(v2 === undefined || v2 === null || v2 === '' ? '—' : v2) || '—';
        doc.setFont('helvetica', 'bold'); doc.setFontSize(7.8); doc.setTextColor.apply(doc, GRY_LBL);
        doc.text(l1 + ' :', mL, y);
        doc.setFont('helvetica', 'normal'); doc.setTextColor.apply(doc, BLACK);
        var ll1 = doc.splitTextToSize(s1, half - KV_LW - KV_GAP);
        doc.text(ll1, mL + KV_LW + KV_GAP, y);
        if (l2) {
          doc.setFont('helvetica', 'bold'); doc.setTextColor.apply(doc, GRY_LBL);
          doc.text(l2 + ' :', xR, y);
          doc.setFont('helvetica', 'normal'); doc.setTextColor.apply(doc, BLACK);
          var ll2 = doc.splitTextToSize(s2, half - KV_LW - KV_GAP);
          doc.text(ll2, xR + KV_LW + KV_GAP, y);
        }
        y += Math.max(KV_LH, ll1.length * KV_LH);
      }

      function gap(n) { y += n || 3; }

      function tbl(head, body, opts) {
        checkPage(24);
        doc.autoTable(Object.assign({
          head, body,
          startY : y,
          margin : { left: mL, right: mR },
          styles : { fontSize: 7, cellPadding: 2, textColor: BLACK, lineColor: [215, 200, 200], lineWidth: 0.18 },
          headStyles: { fillColor: RED, textColor: WHITE, fontStyle: 'bold', fontSize: 7 },
          alternateRowStyles: { fillColor: ROW_ALT },
          theme  : 'grid',
          tableLineColor: [200, 185, 185], tableLineWidth: 0.18,
        }, opts || {}));
        y = doc.lastAutoTable.finalY + 4;
      }

      function subHead(title) {
        checkPage(18);
        doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5); doc.setTextColor.apply(doc, GRY_DRK);
        doc.text(title, mL, y);
        y += 5;
      }

      var sec = 1;
      function S(title) { sectionTitle(sec + '. ' + title); sec++; }

      // ═══════════════════════════════════════════════════════════════════════
      // 1. SUPPLIER IDENTITY
      // ═══════════════════════════════════════════════════════════════════════
      S('Supplier Identity');
      kv2('Supplier Type',    supplierType,                                  'Submission ID',   sid);
      kv2('Company Name',     p['Company Name']  || '—',                     'Mill Name',       p['Mill Name']  || '—');
      kv2('Group Name',       p['Group Name']    || p['Grup Name'] || '—',   'Current Owner',   p['Current Owner'] || '—');
      kv2('UML ID / Mill ID', p['Mill ID']       || p['Mil ID']    || '—',   'UML ID (Alt)',    p['UML ID']     || '—');
      kv ('Office Address',   p['Office Address'] || '—');
      kv2('Latitude',         p['Latitude']       || '—',                     'Longitude',       p['Longitude']  || '—');
      gap(4);

      // ═══════════════════════════════════════════════════════════════════════
      // 2. MILL / KCP INFORMATION
      // ═══════════════════════════════════════════════════════════════════════
      var isKcp = (supplierType === 'KCP');

      // Determine if any mill/kcp info field has a value
      var millInfoLabel = isKcp ? 'KCP Information' : (supplierType === 'TRADER' ? 'Trader Mill Information' : 'Mill Information');
      var millName_       = String(p['Mill Name'] || '').trim();
      var millAddr_       = String(p['Mill Address'] || '').trim();
      var distRef_        = String(p['Distance to Refinery (km)'] || p['Distance to Refinery'] || '').trim();
      var dist_           = String(p['Distance'] || p['Distance (km)'] || p['Distance KM'] || '').trim();
      var millLat_        = String(p['Latitude'] || '').trim();
      var millLng_        = String(p['Longitude'] || '').trim();
      var millCat_        = String(p['Mill Category'] || '').trim();
      var millCap_        = String(p['Mill Capacity (Ton/Hour)'] || p['KCP Capacity (Ton/Hour)'] || p['Capacity (Ton/Hour)'] || '').trim();
      var sterType_       = String(p['Sterilizer Type'] || '').trim();
      var storageTank_    = String(p['Storage Tank Capacity'] || '').trim();
      var siloCap_        = String(p['Silo Capacity'] || '').trim();
      var commDate_       = String(p['Commissioning Date'] || p['Commisioning Date'] || '').trim();
      var hasMillInfo     = millName_ || millAddr_ || distRef_ || dist_ || millCat_ || millCap_ || sterType_ || storageTank_ || siloCap_ || commDate_;

      if (hasMillInfo) {
        S(millInfoLabel);
        var nameLabel = isKcp ? 'KCP Name' : 'Mill Name';
        var addrLabel = isKcp ? 'KCP Address' : 'Mill Address';
        var catLabel  = isKcp ? 'KCP Category' : 'Mill Category';
        var capLabel  = isKcp ? 'KCP Capacity (Ton/Hour)' : 'Mill Capacity (Ton/Hour)';
        kv (nameLabel,           millName_  || '—');
        kv (addrLabel,           millAddr_  || '—');
        kv2('Distance to Refinery', distRef_ || '—', 'Distance', dist_ || '—');
        kv2('Latitude',          millLat_   || '—',  'Longitude', millLng_ || '—');
        kv2(catLabel,            millCat_   || '—',  capLabel,    millCap_ || '—');
        if (!isKcp) kv('Sterilizer Type', sterType_ || '—');
        kv2('Storage Tank Capacity', storageTank_ || '—', 'Silo Capacity', siloCap_ || '—');
        kv ('Commissioning Date', commDate_ || '—');
        gap(4);
      }

      // ═══════════════════════════════════════════════════════════════════════
      // 3. PRODUCT TABLE + QUALITY
      // ═══════════════════════════════════════════════════════════════════════

      // --- Determine supplier type for product layout (same heuristic as pseudo-sheet builder) ---
      var _stypeRawPdf = String(p['supplier_type'] || p['Supplier Type'] || supplierType || '').trim().toUpperCase();
      var _isKcpPdf    = (_stypeRawPdf === 'KCP');

      // --- Parse product lines (JSON first, then individual columns) ---
      var _productLinesPdf = null;
      try {
        var _plJson = String(p['SDD - Product Lines JSON'] || '').trim();
        if (_plJson) {
          var _plArr = JSON.parse(_plJson);
          if (Array.isArray(_plArr) && _plArr.length) _productLinesPdf = _plArr;
        }
      } catch (e_pl) {}

      // If type was not set, infer KCP from product line content (CPKO/PKE, no other col)
      if (!_isKcpPdf && _productLinesPdf && _productLinesPdf.length) {
        var _allNoOtherPdf = _productLinesPdf.every(function(pl) { return !String(pl.other || '').trim(); });
        var _hasKcpProdPdf = _productLinesPdf.some(function(pl) {
          var m = String(pl.main || '').trim().toUpperCase();
          return m === 'CPKO' || m === 'PKE';
        });
        if (_allNoOtherPdf && _hasKcpProdPdf) _isKcpPdf = true;
      }

      // Fallback: build product lines from individual columns
      if (!_productLinesPdf) {
        _productLinesPdf = [];
        // Main products 1-3
        for (var _mi = 1; _mi <= 3; _mi++) {
          var _mk = _mi === 1 ? 'Main Product' : ('Main Product ' + _mi);
          var _mav = _mi === 1 ? 'Main Product Avg Production/Month (Ton)' : ('Main Product ' + _mi + ' Avg Production/Month (Ton)');
          var _myl = _mi === 1 ? 'Main Product Yield' : ('Main Product ' + _mi + ' Yield');
          var _ov  = _mi === 1 ? 'Other Product 1' : ('Other Product ' + _mi);
          var _oav = _mi === 1 ? 'Other Product 1 Avg/Month (Ton)' : ('Other Product ' + _mi + ' Avg/Month (Ton)');
          var _mv  = String(p[_mk]  || '').trim();
          var _mav2 = String(p[_mav] || '').trim();
          var _myl2 = String(p[_myl] || '').trim();
          var _ov2  = String(p[_ov]  || '').trim();
          var _oav2 = String(p[_oav] || '').trim();
          if (_mv || _ov2) {
            _productLinesPdf.push({ main: _mv, mainAvg: _mav2, yield: _myl2, other: _ov2, otherAvg: _oav2 });
          }
        }
        // Other products 4-12 (other-only rows)
        for (var _oi = 4; _oi <= 12; _oi++) {
          var _ok  = 'Other Product ' + _oi;
          var _oav3 = 'Other Product ' + _oi + ' Avg/Month (Ton)';
          var _ov3  = String(p[_ok]   || '').trim();
          var _oav3v = String(p[_oav3] || '').trim();
          if (_ov3) _productLinesPdf.push({ main: '', mainAvg: '', yield: '', other: _ov3, otherAvg: _oav3v });
        }
      }

      // Quality fields (shared between KCP and Mill, same canonical keys)
      // fmtPct: normalise percentage-ish values stored as decimals.
      // "0.05" -> "5%",  "5%" -> "5%",  "20%" -> "20%",  "0.2" -> "20%"
      // Values >= 1 assumed already correct (plain number/string, not a fraction)
      function fmtPct(v) {
        var s = String(v === undefined || v === null ? '' : v).trim();
        if (!s || s === '\u2014') return s;
        if (s.charAt(s.length - 1) === '%') return s;
        var n = parseFloat(s.replace(',', '.'));
        if (isNaN(n)) return s;
        if (n > 0 && n < 1) return (n * 100).toFixed(2).replace(/\.?0+$/, '') + '%';
        return s;
      }

      var _cpoCpoFfa  = fmtPct(String(p['CPO Quality - FFA']   || '').trim());
      var _cpoCpoMi   = fmtPct(String(p['CPO Quality - M&I']   || '').trim());
      var _cpoCpoDobi = fmtPct(String(p['CPO Quality - DOBI']  || '').trim());
      var _pkFfa      = fmtPct(String(p['PK Quality - FFA']    || '').trim());
      var _pkMoist    = fmtPct(String(p['PK Quality - MOIST']  || '').trim());
      var _pkDirt     = fmtPct(String(p['PK Quality - DIRT']   || '').trim());

      var _hasProducts = _productLinesPdf && _productLinesPdf.length;
      var _hasQuality  = _cpoCpoFfa || _cpoCpoMi || _cpoCpoDobi || _pkFfa || _pkMoist || _pkDirt;

      if (_hasProducts || _hasQuality) {
        var _prodSectionTitle = _isKcpPdf ? 'Product to be Produce' : 'Product to be Supply';
        S(_prodSectionTitle);

        if (_hasProducts) {
          if (_isKcpPdf) {
            // KCP: Product | Avg/Month | Yield (no "Other Product" column)
            tbl(
              [['Product', 'Avg Production/Month', 'Yield']],
              _productLinesPdf
                .filter(function(pl) { return String(pl.main || '').trim(); })
                .map(function(pl) { return [
                  String(pl.main    || ''),
                  String(pl.mainAvg || ''),
                  fmtPct(String(pl.yield || ''))
                ]; }),
              {
                styles:    { fontSize: 7.5, cellPadding: 2.5 },
                headStyles: { fillColor: RED, textColor: WHITE, fontStyle: 'bold', fontSize: 7.5 },
                columnStyles: { 0: { cellWidth: 50 }, 1: { cellWidth: 60 } }
              }
            );
          } else {
            // Mill / Trader: Main Product | Avg | Yield | Other Product | Avg
            var _prodBody = _productLinesPdf
              .filter(function(pl) { return String(pl.main || pl.other || '').trim(); })
              .map(function(pl) { return [
                String(pl.main     || ''),
                String(pl.mainAvg  || ''),
                fmtPct(String(pl.yield || '')),
                String(pl.other    || ''),
                String(pl.otherAvg || '')
              ]; });
            if (_prodBody.length) {
              tbl(
                [['Main Product', 'Avg/Month (Ton)', 'Yield', 'Other Product', 'Avg/Month (Ton)']],
                _prodBody,
                {
                  styles:    { fontSize: 7.5, cellPadding: 2.5 },
                  headStyles: { fillColor: RED, textColor: WHITE, fontStyle: 'bold', fontSize: 7.5 },
                  columnStyles: { 0: { cellWidth: 35 }, 1: { cellWidth: 30 }, 2: { cellWidth: 18 }, 3: { cellWidth: 35 } }
                }
              );
            }
          }
        }

        if (_hasQuality) {
          gap(2);
          var _q1Label = _isKcpPdf ? 'CPKO Quality' : 'CPO Quality';
          var _q2Label = _isKcpPdf ? 'PKE Quality'  : 'PK Quality';
          var _q1Mi    = _isKcpPdf ? 'M&I' : 'M&I';
          var _q1DoLbl = _isKcpPdf ? 'DOBI' : 'DOBI';
          tbl(
            [['Quality', 'FFA', (_isKcpPdf ? 'M&I' : 'M&I'), (_isKcpPdf ? 'DOBI' : 'DOBI'), 'MOIST', 'DIRT']],
            [
              [ _q1Label, _cpoCpoFfa || '—', _cpoCpoMi || '—', _cpoCpoDobi || '—', '—', '—' ],
              [ _q2Label, _pkFfa  || '—',  '—', '—', _pkMoist || '—', _pkDirt || '—' ]
            ],
            {
              styles:    { fontSize: 7.5, cellPadding: 2.5 },
              headStyles: { fillColor: RED_MD, textColor: WHITE, fontStyle: 'bold', fontSize: 7.5 },
              columnStyles: { 0: { cellWidth: 28 } }
            }
          );
        }
        gap(3);
      }

      // ═══════════════════════════════════════════════════════════════════════
      // (old section 2 → now renumbered by the auto-counter `sec`)
      // 4. SCREENING SUMMARY
      // ═══════════════════════════════════════════════════════════════════════
      S('Screening Summary');

      // ── No Buy List (NBL) — prominent summary for PDF readers ──
      (function renderNblPdfBlock_() {
        var nblRaw = String(f.nbl || '').trim();
        var nblYes = nblRaw.toLowerCase() === 'yes';
        var nblNo = nblRaw.toLowerCase() === 'no';
        var resultLine = String(f.nblCheckResult || pick('SCR - NBL Check Result') || '').trim();
        if (!resultLine) {
          if (nblYes) resultLine = 'YES — Supplier IS ON the No Buy List (NBL)';
          else if (nblNo) resultLine = 'NO — Supplier is NOT on the No Buy List';
          else resultLine = 'Not checked — run Check NBL before export for a definitive result';
        }
        var detailLine = String(f.nblCheckDetail || pick('SCR - NBL Match Detail') || '').trim();
        if (!detailLine && window._nblCheckResult) {
          if (window._nblCheckResult.matches && window._nblCheckResult.matches.length) {
            detailLine = window._nblCheckResult.matches.map(function(m) {
              return m.source + ': ' + m.detail;
            }).join(' | ');
          } else if (window._nblCheckResult.status === 'No') {
            detailLine = 'No matching Group Name, Company Name, or Mill Name in NBL or Unilever NBL sheets.';
          }
        }
        if (!detailLine && nblNo) {
          detailLine = 'No matching Group Name, Company Name, or Mill Name in NBL or Unilever NBL sheets.';
        }
        if (!detailLine && nblYes) {
          detailLine = 'Supplier matched a name on the NBL or Unilever NBL registry (screening: Yes).';
        }
        var checkedLine = String(f.nblCheckedAt || pick('SCR - NBL Checked At') || '').trim();
        if (!checkedLine && window._nblCheckResult && window._nblCheckResult.checkedAt) {
          checkedLine = window._nblCheckResult.checkedAt;
        }
        if (checkedLine) {
          try {
            var dChk = new Date(checkedLine);
            if (!isNaN(dChk.getTime())) {
              checkedLine = dChk.toLocaleString('id-ID', {
                day: '2-digit', month: 'short', year: 'numeric',
                hour: '2-digit', minute: '2-digit'
              });
            }
          } catch (eChk) {}
        }

        var detailPdfLines = detailLine ? doc.splitTextToSize(detailLine, cW - 8) : [];
        var resultPdfLines = doc.splitTextToSize(resultLine, cW - 8);
        var boxH = 7 + resultPdfLines.length * KV_LH
          + (detailPdfLines.length ? detailPdfLines.length * KV_LH + 2 : 0)
          + (checkedLine ? KV_LH + 2 : 0);
        checkPage(boxH + 4);
        doc.setFillColor.apply(doc, nblYes ? [255, 235, 235] : (nblNo ? [235, 245, 238] : [249, 250, 251]));
        doc.setDrawColor.apply(doc, nblYes ? RED : (nblNo ? [30, 107, 58] : [200, 185, 185]));
        doc.setLineWidth(0.35);
        doc.rect(mL, y, cW, boxH, 'FD');
        var ty = y + 5;
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(8.5);
        doc.setTextColor.apply(doc, nblYes ? RED : (nblNo ? [30, 107, 58] : GRY_DRK));
        doc.text(resultPdfLines, mL + 4, ty);
        ty += resultPdfLines.length * KV_LH + 1;
        if (detailPdfLines.length) {
          doc.setFont('helvetica', 'normal');
          doc.setFontSize(7.5);
          doc.setTextColor.apply(doc, BLACK);
          doc.text(detailPdfLines, mL + 4, ty);
          ty += detailPdfLines.length * KV_LH + 1;
        }
        if (checkedLine) {
          doc.setFont('helvetica', 'italic');
          doc.setFontSize(7);
          doc.setTextColor.apply(doc, GRY_LBL);
          doc.text('NBL checked: ' + checkedLine, mL + 4, ty);
        }
        doc.setTextColor.apply(doc, BLACK);
        doc.setFont('helvetica', 'normal');
        y += boxH + 4;
      })();

      kv ('Group / Owners',  f.owners   || '—');
      kv ('Previous News',   f.news     || '—');
      kv ('Supply To',       f.supplyto || '—');
      kv2('Legality Status', f.legality || '—',  'Certification', f.cert  || '—');
      kv2('NDPE Policy',     f.ndpe     || '—',  'No Buy List (Y/N)', f.nbl || '—');
      kv2('Grievance (Y/N)', f.grvYN    || '—',  'PRI (Y/N)',     f.priYN || '—');
      kv ('Screening Notes', f.note     || '—');
      kv ('Requested Data',  f.reqData  || '—');
      kv2('Status',          f.status   || '—',  'Last Updated',  f.date  || '—');
      gap(4);

      // ═══════════════════════════════════════════════════════════════════════
      // 3. CERTIFICATION DETAILS (C. Certification from primary row)
      // ═══════════════════════════════════════════════════════════════════════
      var certDefs = [
        { label: 'RSPO',      statusKey: 'RSPO Status',      noKey: 'RSPO Cert No',  startKey: 'RSPO Valid Start', endKey: 'RSPO Valid End',  bodyKey: 'RSPO Cert Body'  },
        { label: 'ISCC',      statusKey: 'ISCC Status',      noKey: 'ISCC Cert No',  startKey: 'ISCC Valid Start', endKey: 'ISCC Valid End',  bodyKey: 'ISCC Cert Body'  },
        { label: 'ISPO Mill', statusKey: 'ISPO Status',      noKey: 'ISPO Cert No',  startKey: 'ISPO Valid Start', endKey: 'ISPO Valid End',  bodyKey: 'ISPO Cert Body'  },
        { label: 'INS',       statusKey: 'INS Status',       noKey: 'INS Cert No',   startKey: 'INS Valid Start',  endKey: 'INS Valid End',   bodyKey: 'INS Cert Body'   },
        { label: 'GGL',       statusKey: 'GGL Status',       noKey: 'GGL Cert No',   startKey: 'GGL Valid Start',  endKey: 'GGL Valid End',   bodyKey: 'GGL Cert Body'   }
      ];
      // Only show section if there is any cert data
      var hasCert = certDefs.some(function(c) {
        return String(p[c.statusKey] || p[c.noKey] || p[c.startKey] || '').trim();
      });
      if (hasCert) {
        S('Certification Details');
        tbl(
          [['Certification', 'Status', 'No. Sertifikat', 'Validity Start', 'Validity End', 'Certification Body']],
          certDefs.map(function(c) {
            return [
              c.label,
              p[c.statusKey] || '—',
              p[c.noKey]     || '—',
              p[c.startKey]  || '—',
              p[c.endKey]    || '—',
              p[c.bodyKey]   || '—'
            ];
          }),
          {
            styles: { fontSize: 7.5, cellPadding: 2.5 },
            headStyles: { fillColor: RED, textColor: WHITE, fontStyle: 'bold', fontSize: 7.5 },
            columnStyles: { 0: { cellWidth: 22 }, 1: { cellWidth: 16 }, 2: { cellWidth: 40 }, 3: { cellWidth: 24 }, 4: { cellWidth: 24 } }
          }
        );
        gap(2);
      }

      // ═══════════════════════════════════════════════════════════════════════
      // 4. LEGALITY DETAILS (D. Legality from primary row)
      // ═══════════════════════════════════════════════════════════════════════
      var legalDefs = [
        { label: 'HGU/HGB',         typeKeys: ['HGU/HGB Type','HGU HGB Type','HGU/HGB','HGU HGB'],
          noKeys:   ['HGU/HGB No','HGU HGB No','HGU/HGB Number','HGU HGB Number'],
          dateKeys: ['HGU/HGB Issue Date','HGU HGB Issue Date','HGU/HGB Issue','HGU HGB Issue'] },
        { label: 'IUP',             typeKeys: ['IUP Type','IUP / IUP - P Type'],
          noKeys:   ['IUP No','IUP / IUP - P No','IUP Number'],
          dateKeys: ['IUP Issue Date','IUP / IUP - P Issue Date','IUP Issue'] },
        { label: 'Izin Lokasi',     typeKeys: ['Izin Lokasi Type','IZIN LOKASI Type','Izin Lokasi'],
          noKeys:   ['Izin Lokasi No','IZIN LOKASI No','Izin Lokasi Number','IZIN LOKASI'],
          dateKeys: ['Izin Lokasi Issue Date','IZIN LOKASI Issue Date','Izin Lokasi Issue'] },
        { label: 'Izin Lingkungan', typeKeys: ['Izin Lingkungan Type','IZIN LINGKUNGAN Type','Izin Lingkungan'],
          noKeys:   ['Izin Lingkungan No','IZIN LINGKUNGAN No','Izin Lingkungan Number','IZIN LINGKUNGAN'],
          dateKeys: ['Izin Lingkungan Issue Date','IZIN LINGKUNGAN Issue Date','Izin Lingkungan Issue'] },
        { label: 'IMB/PBG',         typeKeys: ['IMB/PBG Type','IMB / PBG Type','IMB PBG Type','IMB/PBG'],
          noKeys:   ['IMB/PBG No','IMB / PBG No','IMB PBG No','IMB/PBG Number','IMB','PBG'],
          dateKeys: ['IMB/PBG Issue Date','IMB / PBG Issue Date','IMB PBG Issue Date','IMB/PBG Issue'] },
        { label: 'NIB',             typeKeys: ['NIB Type','NIB'],
          noKeys:   ['NIB No','NIB Number'],
          dateKeys: ['NIB Issue Date'] }
      ];
      var hasLegal = legalDefs.some(function(d) {
        return pick.apply(null, d.typeKeys.concat(d.noKeys).concat(d.dateKeys));
      });
      if (hasLegal) {
        S('Legality Details');
        tbl(
          [['Document', 'Type', 'No. Dokumen', 'Issue Date']],
          legalDefs.map(function(d) {
            return [
              d.label,
              pick.apply(null, d.typeKeys)  || '—',
              pick.apply(null, d.noKeys)    || '—',
              pick.apply(null, d.dateKeys)  || '—'
            ];
          }),
          {
            styles: { fontSize: 7.5, cellPadding: 2.5 },
            headStyles: { fillColor: RED, textColor: WHITE, fontStyle: 'bold', fontSize: 7.5 },
            columnStyles: { 0: { cellWidth: 30 }, 1: { cellWidth: 22 }, 2: { cellWidth: 60 } }
          }
        );
        gap(2);
      }

      // ═══════════════════════════════════════════════════════════════════════
      // 5. MILL LIST (TML)
      // ═══════════════════════════════════════════════════════════════════════
      if (millRows.length) {
        S('Mill List (TML)');
        tbl(
          [['Company', 'Mill Name', 'UML ID', 'Village', 'Sub District', 'District', 'Cap.(T/H)', 'Lat', 'Long', 'Legality', 'ISPO', 'RSPO', 'ISCC', 'Supply CPO/PK (T)']],
          millRows.map(function(m) { return [
            m['TML - Company Name']              || '',
            m['TML - Mill Name']                 || '',
            m['TML - UML ID']                    || '',
            m['TML - Village']                   || '',
            m['TML - Sub District']              || '',
            m['TML - District']                  || '',
            m['TML - Capacity (Ton/Hour)']       || '',
            m['TML - Latitude']                  || '',
            m['TML - Longitude']                 || '',
            m['TML - Legality']                  || '',
            m['TML - ISPO (Y/N)']                || '',
            m['TML - RSPO (Y/N)']                || '',
            m['TML - ISCC (Y/N)']                || '',
            m['TML - Total Supply CPO/PK (Ton)'] || ''
          ]; }),
          { styles: { fontSize: 6, cellPadding: 1.8 }, headStyles: { fillColor: RED, textColor: WHITE, fontStyle: 'bold', fontSize: 6 },
            columnStyles: { 0: { cellWidth: 20 }, 1: { cellWidth: 20 }, 13: { cellWidth: 14 } } }
        );

        var fmtItems = function(items) {
          return (items || []).map(function(i) { return i.label + ' (' + (i.ha || '') + ')'; }).join('; ');
        };
        var tmlScr = millRows.filter(function(m) {
          var k = m['TML - Mill Name'] || '';
          var sd = (window._tmlScreeningData && k) ? window._tmlScreeningData[k] : null;
          return (sd && Object.values(sd).some(function(v) { return String(v || '').trim(); })) ||
            String(m['SCR - TML Screening Status'] || m['SCR - TML Valid Coordinate'] || m['SCR - TML Moratorium'] || '').trim();
        });
        if (tmlScr.length) {
          gap(1); subHead('Mill Screening Detail');
          tbl(
            [['Mill Name', 'Valid Coord', 'Forest Area', 'Peatland', 'Moratorium', 'Mora (Ha)', 'Defbuf (Ha)', 'Status', 'Date']],
            tmlScr.map(function(m) {
              var k = m['TML - Mill Name'] || '';
              var sd = (window._tmlScreeningData && k) ? (window._tmlScreeningData[k] || {}) : {};
              return [
                k,
                m['SCR - TML Valid Coordinate']               || sd.coord  || '',
                m['SCR - TML Forest Area']                    || fmtItems(sd.forestItems) || '',
                m['SCR - TML Peatland']                       || fmtItems(sd.peatItems)  || '',
                m['SCR - TML Moratorium']                     || sd.mora   || '',
                m['SCR - TML Moratorium (Ha)']                || sd.moraHa || '',
                m['SCR - TML Deforestation Buffer 50KM (Ha)'] || sd.defbuf || '',
                m['SCR - TML Screening Status']               || sd.status || '',
                m['SCR - TML Screening Date']                 || sd.date   || ''
              ];
            }),
            { styles: { fontSize: 6.2, cellPadding: 1.8 }, headStyles: { fillColor: RED_MD, textColor: WHITE, fontStyle: 'bold', fontSize: 6.2 },
              columnStyles: { 2: { cellWidth: 26 }, 3: { cellWidth: 20 } } }
          );
        }
        gap(2);
      }

      // ═══════════════════════════════════════════════════════════════════════
      // 6. FFB SUPPLIER LIST
      // ═══════════════════════════════════════════════════════════════════════
      if (ffbRows.length) {
        S('FFB Supplier List');
        tbl(
          [['Mill Name', 'Supplier Name', 'Supplier Group', 'Village', 'Sub District', 'District', 'Category', 'Concession (Ha)', 'Planted (Ha)', 'Plant Year', 'Legality', 'ISPO', 'RSPO', 'ISCC', 'Supply FFB (T)']],
          ffbRows.map(function(r) { return [
            r['FFB - Mill Name']              || '',
            r['FFB - Supplier Name']          || '',
            r['FFB - Supplier Group Name']    || '',
            r['FFB - Village']                || '',
            r['FFB - Sub District']           || '',
            r['FFB - District']               || '',
            r['FFB - Supplier Category']      || r['FFB - Category'] || '',
            r['FFB - Concession Area (Ha)']   || '',
            r['FFB - Planted Area (Ha)']      || '',
            r['FFB - Planted Year']           || '',
            r['FFB - Legality']               || '',
            r['FFB - ISPO (Y/N)']             || '',
            r['FFB - RSPO (Y/N)']             || '',
            r['FFB - ISCC (Y/N)']             || '',
            r['FFB - Total Supply FFB (Ton)'] || ''
          ]; }),
          { styles: { fontSize: 5.8, cellPadding: 1.6 }, headStyles: { fillColor: RED, textColor: WHITE, fontStyle: 'bold', fontSize: 5.8 },
            columnStyles: { 1: { cellWidth: 20 }, 6: { cellWidth: 18 }, 14: { cellWidth: 14 } } }
        );

        var ffbScr = ffbRows.filter(function(r) {
          var k = r['FFB - Supplier Name'] || '';
          var sd = (window._ffbScreeningData && k) ? window._ffbScreeningData[k] : null;
          return (sd && Object.values(sd).some(function(v) { return String(v || '').trim(); })) ||
            String(r['FFB - Screening Status'] || r['FFB - Valid Coordinate'] || r['FFB - Moratorium'] || '').trim();
        });
        if (ffbScr.length) {
          gap(1); subHead('FFB Screening Detail');
          tbl(
            [['Supplier', 'Valid Coord', 'Forest Area', 'Peatland', 'Moratorium', 'Mora (Ha)', 'Dist (Km)', 'Defor (Ha)', 'Burn (Ha)', 'Village Risk', 'Status', 'Date']],
            ffbScr.map(function(r) {
              var k = r['FFB - Supplier Name'] || '';
              var sd = (window._ffbScreeningData && k) ? (window._ffbScreeningData[k] || {}) : {};
              return [
                k,
                r['FFB - Valid Coordinate']      || sd.coord       || '',
                r['FFB - Forest Area']           || fmtItems(sd.forestItems) || '',
                r['FFB - Peatland']              || fmtItems(sd.peatItems)  || '',
                r['FFB - Moratorium']            || sd.mora        || '',
                r['FFB - Moratorium (Ha)']       || sd.moraHa      || '',
                r['FFB - Distance to Mill (Km)'] || sd.distKm      || '',
                r['FFB - Deforestation (Ha)']    || sd.defor       || '',
                r['FFB - Burn Area (Ha)']        || sd.burn        || '',
                r['FFB - Village Risk']          || sd.villageRisk || '',
                r['FFB - Screening Status']      || sd.status      || '',
                r['FFB - Screening Date']        || sd.date        || ''
              ];
            }),
            { styles: { fontSize: 6, cellPadding: 1.8 }, headStyles: { fillColor: RED_MD, textColor: WHITE, fontStyle: 'bold', fontSize: 6 },
              columnStyles: { 2: { cellWidth: 24 }, 3: { cellWidth: 18 } } }
          );
        }
        gap(2);
      }

      // ═══════════════════════════════════════════════════════════════════════
      // 7. GRIEVANCE DETAILS
      // ═══════════════════════════════════════════════════════════════════════
      if (grvRows.length) {
        S('Grievance Details');
        tbl(
          [['Source', 'Description', 'Publisher', 'Date', 'Status', 'Attachment']],
          grvRows.map(function(r) { return [r.source||'', r.desc||'', r.pub||'', r.date||'', r.status||'', r.attach||'']; }),
          { columnStyles: { 1: { cellWidth: 42 }, 5: { cellWidth: 28 } } }
        );
      }

      // ═══════════════════════════════════════════════════════════════════════
      // 8. PRI DETAILS
      // ═══════════════════════════════════════════════════════════════════════
      if (priRows.length) {
        S('PRI (Pending Resolution Issue) Details');
        tbl(
          [['Company', 'Description', 'Publisher', 'Date', 'Attachment', 'Action Request']],
          priRows.map(function(r) { return [r.company||'', r.desc||'', r.pub||'', r.date||'', r.attach||'', r.action||'']; }),
          { columnStyles: { 1: { cellWidth: 38 }, 5: { cellWidth: 34 } } }
        );
      }

      // ═══════════════════════════════════════════════════════════════════════
      // 9. APPROVER DECISION
      // ═══════════════════════════════════════════════════════════════════════
      if (f.approverNote || f.approverStat) {
        S('Approver Decision');
        kv2('Decision Status', f.approverStat || '—', '', '');
        kv ('Decision Note',   f.approverNote || '—');
        gap(3);
      }

      // ─── FOOTER ──────────────────────────────────────────────────────────
      var totalPages = doc.internal.getNumberOfPages();
      for (var pg = 1; pg <= totalPages; pg++) {
        doc.setPage(pg);
        doc.setDrawColor.apply(doc, RED);
        doc.setLineWidth(0.35);
        doc.line(mL, pageH - 10, pageW - mR, pageH - 10);
        doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor.apply(doc, [120, 100, 100]);
        doc.text('Supplier Due Diligence Report  ·  ' + supplierName + (supplierType ? '  ·  ' + supplierType : ''), mL, pageH - 6);
        doc.text('Page ' + pg + ' of ' + totalPages, pageW - mR, pageH - 6, { align: 'right' });
      }

      // ─── SAVE ────────────────────────────────────────────────────────────
      var safeName = supplierName.replace(/[^a-zA-Z0-9_\- ]/g, '').replace(/\s+/g, '_').slice(0, 40);
      doc.save('SDD_' + (safeName || 'Report') + '_' + exportedAt.replace(/ /g, '-') + '.pdf');
      if (typeof window.showSddToast === 'function') window.showSddToast('PDF berhasil di-generate ✓', 'success');

    } catch (pdfErr) {
      console.error('[sddExportPdf]', pdfErr);
      if (typeof window.showSddToast === 'function') window.showSddToast('Gagal generate PDF: ' + (pdfErr.message || pdfErr), 'error');
    } finally {
      if (pdfBtn) {
        pdfBtn.disabled = false;
        pdfBtn.innerHTML = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Export PDF';
      }
    }
  }
  window.sddExportPdf = sddExportPdf;
  // ─── END SDD PDF EXPORT ──────────────────────────────────────────────────

  // ═══════════════════════════════════════════════════════════════════════════
  //  SUPPLY IMPORT — Task List (Excel → Draft → Submit to Mill)
  // ═══════════════════════════════════════════════════════════════════════════
  (function initSupplyImport() {
    // ── Year options ─────────────────────────────────────────────────────────
    const yearSel = document.getElementById('supply-import-year');
    if (yearSel) {
      const curYear = new Date().getFullYear();
      for (let y = curYear + 1; y >= curYear - 4; y--) {
        const opt = document.createElement('option');
        opt.value = String(y); opt.textContent = String(y);
        yearSel.appendChild(opt);
      }
      yearSel.value = String(curYear);
    }

    // ── Modal open/close ──────────────────────────────────────────────────────
    const overlay   = document.getElementById('supply-import-modal-overlay');
    const openBtn   = document.getElementById('btn-supply-import-open');
    const closeBtn  = document.getElementById('btn-supply-import-close');
    const cancelBtn = document.getElementById('btn-supply-import-cancel');
    function closeModal() {
      if (overlay) overlay.style.display = 'none';
      resetImportModal_();
    }
    if (openBtn)   openBtn.addEventListener('click', function() { if (overlay) overlay.style.display = 'block'; });
    if (closeBtn)  closeBtn.addEventListener('click', closeModal);
    if (cancelBtn) cancelBtn.addEventListener('click', closeModal);
    if (overlay)   overlay.addEventListener('click', function(e) { if (e.target === overlay) closeModal(); });

    // ── Drag & drop ───────────────────────────────────────────────────────────
    const dropZone  = document.getElementById('supply-import-drop-zone');
    const fileInput = document.getElementById('supply-import-file-input');
    if (dropZone) {
      dropZone.addEventListener('dragover', function(e) { e.preventDefault(); dropZone.style.background = 'rgba(139,26,26,0.04)'; });
      dropZone.addEventListener('dragleave', function() { dropZone.style.background = ''; });
      dropZone.addEventListener('drop', function(e) { e.preventDefault(); dropZone.style.background = ''; if (e.dataTransfer.files[0]) handleSupplyFile_(e.dataTransfer.files[0]); });
    }
    if (fileInput) fileInput.addEventListener('change', function() { if (this.files[0]) handleSupplyFile_(this.files[0]); });

    // ── Quarter/Year validation → enable proceed btn ──────────────────────────
    const quarterSel = document.getElementById('supply-import-quarter');
    function checkProceedReady_() {
      const q = quarterSel ? quarterSel.value : '';
      const y = yearSel ? yearSel.value : '';
      const hasData = window._supplyImportParsedRows && window._supplyImportParsedRows.length > 0;
      const proceedBtn = document.getElementById('btn-supply-import-proceed');
      if (proceedBtn) {
        const ready = !!(q && y && hasData);
        proceedBtn.disabled = !ready;
        proceedBtn.style.opacity = ready ? '1' : '0.45';
        proceedBtn.style.cursor  = ready ? 'pointer' : 'not-allowed';
      }
    }
    if (quarterSel) quarterSel.addEventListener('change', checkProceedReady_);
    if (yearSel)    yearSel.addEventListener('change', checkProceedReady_);

    // ── Proceed → build task list ─────────────────────────────────────────────
    const proceedBtn = document.getElementById('btn-supply-import-proceed');
    if (proceedBtn) proceedBtn.addEventListener('click', function() {
      const q = quarterSel ? quarterSel.value : '';
      const y = yearSel ? yearSel.value : '';
      if (!q || !y || !window._supplyImportParsedRows) return;
      buildSupplyTaskList_(window._supplyImportParsedRows, q, y);
      closeModal();
    });

    // ── Load existing drafts on panel open ────────────────────────────────────
    loadSupplyDraftsFromServer_();
  })();

  function resetImportModal_() {
    window._supplyImportParsedRows = null;
    const fileInput  = document.getElementById('supply-import-file-input');
    const fileInfo   = document.getElementById('supply-import-file-info');
    const fileError  = document.getElementById('supply-import-file-error');
    const preview    = document.getElementById('supply-import-preview-wrap');
    const proceedBtn = document.getElementById('btn-supply-import-proceed');
    if (fileInput)   { fileInput.value = ''; }
    if (fileInfo)    { fileInfo.style.display = 'none'; }
    if (fileError)   { fileError.style.display = 'none'; fileError.textContent = ''; }
    if (preview)     { preview.style.display = 'none'; }
    if (proceedBtn)  { proceedBtn.disabled = true; proceedBtn.style.opacity = '0.45'; }
  }

  function handleSupplyFile_(file) {
    const fileInfo  = document.getElementById('supply-import-file-info');
    const fileError = document.getElementById('supply-import-file-error');
    const fileName  = document.getElementById('supply-import-file-name');
    const rowCount  = document.getElementById('supply-import-row-count');

    function showErr(msg) {
      if (fileError) { fileError.textContent = msg; fileError.style.display = 'block'; }
      if (fileInfo)  fileInfo.style.display = 'none';
    }
    if (fileError) fileError.style.display = 'none';

    if (typeof XLSX === 'undefined') {
      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
      script.onload = function() { parseSupplyFile_(file, showErr); };
      document.head.appendChild(script);
    } else {
      parseSupplyFile_(file, showErr);
    }

    function parseSupplyFile_(f, errCb) {
      const reader = new FileReader();
      reader.onload = function(evt) {
        try {
          const wb = XLSX.read(evt.target.result, { type: 'array' });
          const sheetName = wb.SheetNames.find(function(n) {
            return n.toLowerCase().trim().includes('contoh pengisian');
          });
          if (!sheetName) {
            errCb('Sheet "Contoh Pengisian" tidak ditemukan. Pastikan file sudah sesuai template.');
            return;
          }
          const ws   = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          if (rows.length < 2) { errCb('Sheet "Contoh Pengisian" kosong.'); return; }

          const headers = rows[0].map(function(h) { return String(h || '').trim().toUpperCase(); });
          const dataRows = rows.slice(1).filter(function(r) {
            return r.some(function(c) { return String(c || '').trim() !== ''; });
          });
          if (!dataRows.length) { errCb('Tidak ada data di sheet "Contoh Pengisian".'); return; }

          // Map columns
          const colPlant    = headers.indexOf('PLANT');
          const colCategory = headers.indexOf('CATEGORY');
          const colCoGroup  = headers.indexOf('COMPANY GROUP NAME');
          const colCoName   = headers.indexOf('COMPANY NAME');
          const colMillName = headers.indexOf('MILL NAME');
          const colSumQty   = headers.indexOf('SUM OF QTY KG');

          if (colCoName < 0 || colMillName < 0) {
            errCb('Kolom COMPANY NAME atau MILL NAME tidak ditemukan di sheet "Contoh Pengisian".');
            return;
          }

          window._supplyImportParsedRows = dataRows.map(function(r) {
            return {
              PLANT:              String(r[colPlant]    != null ? r[colPlant]    : '').trim(),
              CATEGORY:           String(r[colCategory] != null ? r[colCategory] : '').trim(),
              COMPANY_GROUP_NAME: String(r[colCoGroup]  != null ? r[colCoGroup]  : '').trim(),
              COMPANY_NAME:       String(r[colCoName]   != null ? r[colCoName]   : '').trim(),
              MILL_NAME:          String(r[colMillName] != null ? r[colMillName] : '').trim(),
              SUM_QTY_KG:         r[colSumQty] != null ? r[colSumQty] : '',
            };
          }).filter(function(r) { return r.COMPANY_NAME || r.MILL_NAME; });

          // Show file info
          if (fileName) fileName.textContent = f.name;
          if (rowCount) rowCount.textContent = window._supplyImportParsedRows.length + ' baris';
          if (fileInfo) fileInfo.style.display = 'flex';

          // Build preview table
          buildImportPreview_(window._supplyImportParsedRows);

          // Re-check proceed readiness
          const quarterSel = document.getElementById('supply-import-quarter');
          const yearSel    = document.getElementById('supply-import-year');
          const q = quarterSel ? quarterSel.value : '';
          const y = yearSel ? yearSel.value : '';
          const proceedBtn = document.getElementById('btn-supply-import-proceed');
          if (proceedBtn) {
            const ready = !!(q && y);
            proceedBtn.disabled = !ready;
            proceedBtn.style.opacity = ready ? '1' : '0.45';
            proceedBtn.style.cursor  = ready ? 'pointer' : 'not-allowed';
          }
        } catch (ex) {
          errCb('Gagal membaca file: ' + ex.message);
        }
      };
      reader.readAsArrayBuffer(f);
    }
  }

  function buildImportPreview_(parsedRows) {
    const wrap   = document.getElementById('supply-import-preview-wrap');
    const tHead  = document.getElementById('supply-import-preview-head');
    const tBody  = document.getElementById('supply-import-preview-body');
    if (!wrap || !tHead || !tBody) return;

    tHead.innerHTML = '<tr>'
      + ['Company Name','Mill Name','Plant','Category','Supply (Kg)'].map(function(h) {
          return '<th style="padding:6px 10px;text-align:left;font-size:10.5px;font-weight:700;color:#5A3030;background:#f9f3f3;border-bottom:1px solid rgba(139,26,26,0.1);white-space:nowrap;">' + escHtml(h) + '</th>';
        }).join('') + '</tr>';

    const preview = parsedRows.slice(0, 8);
    tBody.innerHTML = preview.map(function(r, i) {
      const bg = i % 2 === 0 ? '#fff' : '#fdf9f9';
      return '<tr style="background:' + bg + ';">'
        + [r.COMPANY_NAME, r.MILL_NAME, r.PLANT, r.CATEGORY, r.SUM_QTY_KG].map(function(v) {
            return '<td style="padding:6px 10px;font-size:11.5px;color:#2A1010;border-bottom:1px solid rgba(139,26,26,0.05);max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">'
              + escHtml(String(v || '—')) + '</td>';
          }).join('') + '</tr>';
    }).join('');
    if (parsedRows.length > 8) {
      tBody.innerHTML += '<tr><td colspan="5" style="padding:8px 10px;font-size:11px;color:#9C8080;text-align:center;">… dan ' + (parsedRows.length - 8) + ' baris lagi</td></tr>';
    }
    wrap.style.display = 'block';
  }

  // ── Build task list in panel ──────────────────────────────────────────────
  // State: current open/editing draft batches
  window._supplyDraftBatches = window._supplyDraftBatches || [];

  function buildSupplyTaskList_(parsedRows, quarter, year) {
    const batchId = 'batch-' + Date.now();
    const now     = new Date().toISOString();
    const rows    = parsedRows.map(function(r, idx) {
      // Try to find matching mill in allData
      const match = allData.find(function(d) {
        const cn = String(d['COMPANY NAME'] || '').trim().toLowerCase();
        const mn = String(d['MILL NAME']    || '').trim().toLowerCase();
        return cn === r.COMPANY_NAME.toLowerCase() && mn === r.MILL_NAME.toLowerCase();
      });

      const draft = {
        draft_id:    batchId + '_' + idx,
        batch_id:    batchId,
        status:      'draft',
        quarter:     quarter,
        year:        year,
        match_status: match ? 'matched' : 'new',
        QUARTER:     quarter,
        YEAR:        year,
      };

      // Mapping: Excel → Mill columns
      draft['FACILITY NAME CPO'] = r.PLANT;
      draft['TRADER NAME']       = r.CATEGORY;
      draft['GROUP NAME']        = r.COMPANY_GROUP_NAME || (match ? (match['GROUP NAME'] || '') : '');
      draft['COMPANY NAME']      = r.COMPANY_NAME;
      draft['MILL NAME']         = r.MILL_NAME;
      draft['SUPPLY CPO']        = r.SUM_QTY_KG;

      // Prefill all other mill fields from matched record (EXCEPT SUPPLY CPO)
      if (match) {
        const PREFILL_SKIP = new Set(['SUPPLY CPO', 'QUARTER', 'YEAR', 'COMPANY NAME', 'MILL NAME', 'GROUP NAME', 'FACILITY NAME CPO', 'TRADER NAME']);
        (window.MILL_FIELDS_LIST || [
          'COMPANY CODE','UML ID','ADDRESS','PROVINCE','COORDINATES',
          'MILL CATEGORY','MILL CAPACITY (TON/HOUR)','HGU/HGB','IZIN LOKASI',
          'IUP','IZIN LINGKUNGAN','SCORE','MILL LOC','COMPLIMENT/NOT COMPLIMENT',
          'DEFORESTATION SPATIAL','BURN AREA SPATIAL','PEAT','LEGALITY',
          'DEFORESTATION GRIEVANCES','BURN AREA GRIEVANCES','HUMAN RIGHT',
          'SAFETY','SOCIAL','ENVIRONMENT','TOTAL GRIEVANCES','NDPE','HRDD',
          'TOTAL POLICY','CERTIFICATION','TOTAL CERTIFICATION','TOTAL SCORE',
          'SUPPLIER LEVEL','BUYER NO BUY LIST','VOLUME SUPPLY STATUS',
          'RECOMMENDATION LEVEL','SIGN','SUPPLIER STATUS','RISK LEVEL',
          'RESULT RISK LEVEL','FACILITY NAME PK','PRODUCT SUPPLY',
        ]).forEach(function(f) {
          if (PREFILL_SKIP.has(f)) return;
          if (match[f] !== undefined && match[f] !== null && String(match[f]).trim() !== '') {
            draft[f] = match[f];
          }
        });
      }

      return draft;
    });

    const batch = { batch_id: batchId, quarter: quarter, year: year, rows: rows, status: 'draft', created_at: now };
    window._supplyDraftBatches.push(batch);
    renderSupplyDraftList_();

    // Auto-save draft to server
    apiPost({ action: 'saveSupplyDraft', batch_id: batchId, rows: rows, meta: { quarter, year } })
      .catch(function(err) { console.warn('[supplyDraft] Auto-save failed:', err.message); });
  }

  function renderSupplyDraftList_() {
    const container = document.getElementById('supply-draft-list');
    const empty     = document.getElementById('supply-draft-empty');
    if (!container) return;

    const batches = window._supplyDraftBatches || [];
    if (!batches.length) {
      container.innerHTML = '';
      if (empty) empty.style.display = 'block';
      return;
    }
    if (empty) empty.style.display = 'none';

    container.innerHTML = batches.map(function(b) {
      const rowCount    = b.rows ? b.rows.length : 0;
      const matched     = b.rows ? b.rows.filter(function(r) { return r.match_status === 'matched'; }).length : 0;
      const newCount    = b.rows ? b.rows.filter(function(r) { return r.match_status !== 'matched'; }).length : 0;
      const doneCount   = b.rows ? b.rows.filter(function(r) { return r._submitted; }).length : 0;
      const isSubmitted = b.status === 'submitted';
      const statusBadge = isSubmitted
        ? '<span class="supply-badge supply-badge--submitted">Submitted</span>'
        : '<span class="supply-badge supply-badge--draft">Draft</span>';
      const createdAt = b.created_at ? new Date(b.created_at).toLocaleString('id-ID', { day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' }) : '';

      return '<div class="supply-batch-card" data-batch-id="' + escHtml(b.batch_id) + '">'
        + '<div class="supply-batch-card-head">'
        + '<div class="supply-batch-card-meta">'
        + '<span class="supply-batch-period">' + escHtml(b.quarter) + ' · ' + escHtml(b.year) + '</span>'
        + statusBadge
        + '<span class="supply-batch-info">' + rowCount + ' rows · ' + matched + ' matched · ' + newCount + ' new'
        + (doneCount > 0 ? ' · ' + doneCount + ' done' : '') + '</span>'
        + (createdAt ? '<span class="supply-batch-date">' + createdAt + '</span>' : '')
        + '</div>'
        + '<div class="supply-batch-actions">'
        + (!isSubmitted ? '<button type="button" class="supply-btn supply-btn--ghost supply-btn--expand" data-batch="' + escHtml(b.batch_id) + '">Lihat / Edit</button>' : '')
        + '<button type="button" class="supply-btn supply-btn--danger" data-action="delete-batch" data-batch="' + escHtml(b.batch_id) + '">Hapus</button>'
        + '</div>'
        + '</div>'
        + '<div class="supply-batch-table-wrap" id="supply-batch-table-' + escHtml(b.batch_id) + '" style="display:none;">'
        + renderSupplyBatchTable_(b)
        + '</div>'
        + '</div>';
    }).join('');

    // Bind events
    container.querySelectorAll('[data-action]').forEach(function(btn) {
      btn.addEventListener('click', handleSupplyBatchAction_);
    });
    container.querySelectorAll('.supply-btn--expand').forEach(function(btn) {
      btn.addEventListener('click', function() {
        const batchId = btn.dataset.batch;
        const wrap    = document.getElementById('supply-batch-table-' + batchId);
        if (!wrap) return;
        const isOpen = wrap.style.display !== 'none';
        wrap.style.display = isOpen ? 'none' : 'block';
        btn.textContent    = isOpen ? 'Lihat / Edit' : 'Tutup';
      });
    });
  }

  function renderSupplyBatchTable_(batch) {
    if (!batch.rows || !batch.rows.length) return '<p style="padding:12px;font-size:12px;color:#9C8080;">Tidak ada baris.</p>';
    const isSubmitted = batch.status === 'submitted';
    const SHOW_COLS = [
      ['COMPANY NAME',    'Company Name'],
      ['MILL NAME',       'Mill Name'],
      ['GROUP NAME',      'Group Name'],
      ['TRADER NAME',     'Trader Name (Category)'],
      ['FACILITY NAME CPO','Facility CPO (Plant)'],
      ['SUPPLY CPO',      'Supply CPO (Kg)'],
      ['PROVINCE',        'Province'],
    ];

    const head = '<tr>'
      + (isSubmitted ? '' : '<th style="width:24px;"></th>')
      + SHOW_COLS.map(function(c) {
          return '<th style="padding:7px 10px;text-align:left;font-size:10.5px;font-weight:700;color:#5A3030;background:#f9f3f3;border-bottom:1px solid rgba(139,26,26,0.1);white-space:nowrap;">' + escHtml(c[1]) + '</th>';
        }).join('')
      + '<th style="width:80px;background:#f9f3f3;border-bottom:1px solid rgba(139,26,26,0.1);"></th>'
      + '</tr>';

    const body = batch.rows.map(function(row, i) {
      const isMatched  = row.match_status === 'matched';
      const isRowDone  = row._submitted === true;
      const matchBadge = isMatched
        ? '<span style="font-size:9.5px;font-weight:700;background:rgba(46,125,50,0.12);color:#2e7d32;border-radius:4px;padding:1px 5px;margin-left:4px;">✓ Match</span>'
        : '<span style="font-size:9.5px;font-weight:700;background:rgba(139,26,26,0.1);color:#8B1A1A;border-radius:4px;padding:1px 5px;margin-left:4px;">New</span>';
      const cells = SHOW_COLS.map(function(c) {
        const key = c[0];
        const val = row[key] != null ? row[key] : '';
        if (key === 'COMPANY NAME') {
          return '<td style="padding:7px 10px;font-size:12px;color:#1A0A0A;border-bottom:1px solid rgba(139,26,26,0.05);">'
            + escHtml(String(val || '—')) + matchBadge + '</td>';
        }
        if (!isSubmitted && !isRowDone && (key === 'SUPPLY CPO' || key === 'FACILITY NAME CPO' || key === 'TRADER NAME' || key === 'GROUP NAME')) {
          return '<td style="padding:4px 6px;border-bottom:1px solid rgba(139,26,26,0.05);">'
            + '<input class="supply-inline-input" data-batch="' + escHtml(batch.batch_id) + '" data-row="' + i + '" data-field="' + escHtml(key) + '"'
            + ' value="' + escHtml(String(val || '')) + '"'
            + ' style="width:100%;height:28px;border:1px solid rgba(139,26,26,0.18);border-radius:5px;padding:0 7px;font-size:11.5px;color:#1A0A0A;background:#fff;">'
            + '</td>';
        }
        return '<td style="padding:7px 10px;font-size:12px;color:#2A1010;border-bottom:1px solid rgba(139,26,26,0.05);">' + escHtml(String(val || '—')) + '</td>';
      }).join('');

      // Per-row action button — all rows open modal for review/edit before save
      let rowAction = '';
      if (!isSubmitted) {
        if (isRowDone) {
          rowAction = '<span style="font-size:10px;font-weight:700;color:#2e7d32;">✓ Submitted</span>';
        } else {
          const btnLabel = isMatched ? 'Edit / Submit →' : 'Lengkapi →';
          const btnStyle = isMatched ? 'supply-btn--primary' : 'supply-btn--ghost';
          rowAction = '<button type="button" class="supply-btn ' + btnStyle + '" style="padding:4px 10px;font-size:11px;white-space:nowrap;"'
            + ' data-action="open-modal-row" data-batch="' + escHtml(batch.batch_id) + '" data-row="' + i + '">' + btnLabel + '</button>';
        }
      }

      return '<tr' + (isRowDone ? ' style="opacity:0.5;"' : '') + '>'
        + (isSubmitted ? '' : '<td style="padding:7px 6px;border-bottom:1px solid rgba(139,26,26,0.05);vertical-align:middle;"><input type="checkbox" class="supply-row-check" data-batch="' + escHtml(batch.batch_id) + '" data-row="' + i + '"' + (isRowDone ? ' disabled' : ' checked') + '></td>')
        + cells
        + '<td style="padding:4px 8px;border-bottom:1px solid rgba(139,26,26,0.05);text-align:right;white-space:nowrap;">' + rowAction + '</td>'
        + '</tr>';
    }).join('');

    return '<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:12px;">'
      + '<thead>' + head + '</thead><tbody>' + body + '</tbody></table></div>'
      + (!isSubmitted ? supplyBatchFooterHtml_(batch.batch_id) : '');
  }

  function supplyBatchFooterHtml_(batchId) {
    const batch   = (window._supplyDraftBatches || []).find(function(b) { return b.batch_id === batchId; });
    const matched = batch ? (batch.rows || []).filter(function(r) { return r.match_status === 'matched' && !r._submitted; }).length : 0;
    return '<div style="display:flex;align-items:center;justify-content:space-between;gap:10px;padding:12px 10px 4px;border-top:1px solid rgba(139,26,26,0.07);flex-wrap:wrap;">'
      + '<span style="font-size:11px;color:#9C8080;">Klik <strong>Edit / Submit →</strong> atau <strong>Lengkapi →</strong> untuk review/isi data lalu Save. Atau bulk-submit semua Matched langsung.</span>'
      + '<div style="display:flex;gap:8px;">'
      + '<button type="button" class="supply-btn supply-btn--ghost" data-action="save-draft" data-batch="' + escHtml(batchId) + '">💾 Save as Draft</button>'
      + (matched > 0 ? '<button type="button" class="supply-btn supply-btn--primary" data-action="submit-matched" data-batch="' + escHtml(batchId) + '">✓ Submit Matched (' + matched + ')</button>' : '')
      + '</div>'
      + '</div>';
  }

  function handleSupplyBatchAction_(e) {
    const btn    = e.currentTarget;
    const action = btn.dataset.action;
    const bId    = btn.dataset.batch;
    const batch  = (window._supplyDraftBatches || []).find(function(b) { return b.batch_id === bId; });
    if (!batch) return;

    if (action === 'delete-batch') {
      if (!confirm('Hapus draft batch ' + batch.quarter + ' ' + batch.year + '? Tidak bisa dibatalkan.')) return;
      window._supplyDraftBatches = window._supplyDraftBatches.filter(function(b) { return b.batch_id !== bId; });
      renderSupplyDraftList_();
      apiPost({ action: 'deleteSupplyDraft', batch_id: bId })
        .catch(function(err) { console.warn('[supplyDraft] Delete failed:', err.message); });
      return;
    }

    if (action === 'save-draft') {
      // Collect inline edits
      collectInlineEdits_(bId);
      btn.textContent = 'Menyimpan…'; btn.disabled = true;
      apiPost({ action: 'saveSupplyDraft', batch_id: bId, rows: batch.rows, meta: {} })
        .then(function() {
          btn.textContent = '✓ Tersimpan';
          setTimeout(function() { btn.textContent = '💾 Save as Draft'; btn.disabled = false; }, 2000);
        })
        .catch(function(err) {
          btn.textContent = '💾 Save as Draft'; btn.disabled = false;
          alert('Gagal menyimpan: ' + err.message);
        });
      return;
    }

    // Submit all matched rows directly (data already complete from prefill)
    if (action === 'submit-matched') {
      collectInlineEdits_(bId);
      const matchedRows = (batch.rows || []).filter(function(r) { return r.match_status === 'matched' && !r._submitted; });
      if (!matchedRows.length) { alert('Tidak ada baris matched yang belum disubmit.'); return; }
      if (!confirm('Submit ' + matchedRows.length + ' baris Matched ke Mill Onboarding Profile?')) return;
      btn.textContent = 'Submitting…'; btn.disabled = true;
      apiPost({ action: 'submitSupplyDraft', batch_id: bId, rows: matchedRows })
        .then(function(res) {
          matchedRows.forEach(function(r) { r._submitted = true; });
          const allDone = (batch.rows || []).every(function(r) { return r._submitted; });
          if (allDone) batch.status = 'submitted';
          renderSupplyDraftList_();
          millLoadPromise = null;
          alert('✓ ' + (res.submitted || matchedRows.length) + ' baris Matched berhasil di-submit.');
        })
        .catch(function(err) {
          btn.textContent = '✓ Submit Matched'; btn.disabled = false;
          alert('Gagal submit: ' + err.message);
        });
      return;
    }

    // Submit a single matched row directly
    if (action === 'submit-row') {
      collectInlineEdits_(bId);
      const rowIdx = parseInt(btn.dataset.row, 10);
      const row    = batch.rows && batch.rows[rowIdx];
      if (!row) return;
      btn.textContent = '…'; btn.disabled = true;
      apiPost({ action: 'submitSupplyDraft', batch_id: bId, rows: [row] })
        .then(function() {
          row._submitted = true;
          const allDone = (batch.rows || []).every(function(r) { return r._submitted; });
          if (allDone) batch.status = 'submitted';
          millLoadPromise = null;
          renderSupplyDraftList_();
        })
        .catch(function(err) {
          btn.textContent = 'Submit'; btn.disabled = false;
          alert('Gagal submit: ' + err.message);
        });
      return;
    }

    // Open full Add New Record modal for a New (unmatched) row
    if (action === 'open-modal-row') {
      collectInlineEdits_(bId);
      const rowIdx  = parseInt(btn.dataset.row, 10);
      const draftRow = batch.rows && batch.rows[rowIdx];
      if (!draftRow) return;

      // Build prefill object for the mill modal (uppercase MILL_FIELDS keys)
      const prefill = {};
      Object.keys(draftRow).forEach(function(k) {
        if (k === 'draft_id' || k === 'batch_id' || k === 'status' || k === 'match_status' || k === '_submitted'
            || k === 'created_at' || k === 'updated_at' || k === 'created_by') return;
        if (draftRow[k] !== undefined && draftRow[k] !== null && String(draftRow[k]).trim() !== '') {
          prefill[k] = draftRow[k];
        }
      });
      // Ensure QUARTER and YEAR are set from batch meta if missing in row
      if (!prefill['QUARTER'] && batch.quarter) prefill['QUARTER'] = batch.quarter;
      if (!prefill['YEAR']    && batch.year)    prefill['YEAR']    = String(batch.year);

      // Store context so after save we can mark draft row as submitted
      window._supplyModalContext = { batchId: bId, rowIdx: rowIdx };
      openModal('mill', MILL_FIELDS, 'add', prefill);
      return;
    }
  }

  function collectInlineEdits_(batchId) {
    const container = document.getElementById('supply-draft-list');
    if (!container) return;
    const inputs = container.querySelectorAll('.supply-inline-input[data-batch="' + batchId + '"]');
    const batch  = (window._supplyDraftBatches || []).find(function(b) { return b.batch_id === batchId; });
    if (!batch) return;
    inputs.forEach(function(inp) {
      const rowIdx = parseInt(inp.dataset.row, 10);
      const field  = inp.dataset.field;
      if (batch.rows[rowIdx] !== undefined) batch.rows[rowIdx][field] = inp.value;
    });
  }

  function getSelectedRows_(batchId, allRows) {
    const container = document.getElementById('supply-draft-list');
    if (!container) return allRows;
    const checks = container.querySelectorAll('.supply-row-check[data-batch="' + batchId + '"]');
    if (!checks.length) return allRows;
    const selected = [];
    checks.forEach(function(cb) {
      if (cb.checked) {
        const idx = parseInt(cb.dataset.row, 10);
        if (allRows[idx]) selected.push(allRows[idx]);
      }
    });
    return selected.length ? selected : allRows;
  }

  function loadSupplyDraftsFromServer_() {
    apiGet('supplyDraft')
      .then(function(data) {
        if (!Array.isArray(data) || !data.length) {
          renderSupplyDraftList_();
          return;
        }
        // Group by batch_id
        const batches = {};
        data.forEach(function(row) {
          const bid = row.batch_id || row.draft_id || 'unknown';
          if (!batches[bid]) {
            batches[bid] = {
              batch_id:   bid,
              quarter:    row.quarter || '',
              year:       row.year    || '',
              status:     row.status  || 'draft',
              created_at: row.created_at || '',
              rows:       [],
            };
          }
          batches[bid].rows.push(row);
        });
        window._supplyDraftBatches = Object.values(batches);
        renderSupplyDraftList_();
      })
      .catch(function(err) {
        console.warn('[supplyDraft] Load drafts failed:', err.message);
        renderSupplyDraftList_();
      });
  }
  // ─── END SUPPLY IMPORT ────────────────────────────────────────────────────

  window.sddExportPdf = sddExportPdf;
  // ─── END SDD PDF EXPORT ──────────────────────────────────────────────────

  } catch (bootErr) {
    console.error('[initDashboardApp]', bootErr);
    var _fatal = document.createElement('div');
    _fatal.setAttribute('role', 'alert');
    _fatal.style.cssText = 'position:fixed;inset:0;display:flex;align-items:center;justify-content:center;padding:24px;background:#e8ddd8;z-index:2147483647;font:15px/1.5 system-ui,sans-serif;color:#2a0e0e;';
    var _box = document.createElement('div');
    _box.style.cssText = 'max-width:420px;background:#fff;border-radius:16px;padding:24px;box-shadow:0 16px 48px rgba(0,0,0,.12);border:1px solid rgba(0,0,0,.06)';
    var _t = document.createElement('strong');
    _t.textContent = 'Startup error';
    var _p = document.createElement('p');
    _p.style.cssText = 'margin:12px 0 0;color:#6a4a4a';
    _p.textContent = String(bootErr && bootErr.message ? bootErr.message : bootErr);
    var _hint = document.createElement('p');
    _hint.style.cssText = 'margin:12px 0 0;font-size:13px;color:#9a7070';
    _hint.textContent =
      'Jalankan npm run dev di folder sustain-dashboard, lalu buka URL di terminal (biasanya http://127.0.0.1:5340 — ikuti baris Local jika port lain). Jangan membuka file index.html langsung dari Finder.';
    _box.appendChild(_t);
    _box.appendChild(_p);
    _box.appendChild(_hint);
    _fatal.appendChild(_box);
    document.body.appendChild(_fatal);
  }

}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initDashboardApp);
} else {
  initDashboardApp();
}