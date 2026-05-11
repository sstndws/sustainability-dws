import { mountLoginPage } from './login-ui.js';
import { getSupabase } from './supabase-client.js';

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
    const fileName = document.getElementById('supplierExcelFile')?.files[0]?.name || '';
    const tp = getCurrentSddSupplierType();
    window._scrKey = ((tp ? tp + '_' : '') + fileName.replace(/[^a-zA-Z0-9]/g,'_').toLowerCase()) || 'default';
    window._loadedPrimarySddRow = null;
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
      // Render into the sheet body div
      const sheetBody = wrapper.querySelector('[id^="sheet-body-"]');
      supplierDisplayExcelDataTo(sheetBody, jsonData);
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
      var isApprover = window._sddUserRole === 'APPROVER';

      // Approver: tiap section punya warna header berbeda
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
        + '<div style="font-size:12px;color:#6b7280;">Use Google Drive share link or direct image/PDF URL.</div>'
        + '<div id="traceAttachmentPreview" style="display:grid;gap:12px;color:#6b7280;font-size:13px;">No attachment yet.</div>'
        + '</div>'
        + '<div class="sdd-hide-for-approver" style="margin-top:20px;padding-top:16px;border-top:1px solid rgba(74,28,28,0.08);">'
        + '<div style="display:flex;align-items:center;gap:6px;margin-bottom:8px;">'
        + '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#8B1A1A" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>'
        + '<label for="noteBossDecision" style="font-size:11px;font-weight:700;color:#8B1A1A;letter-spacing:0.08em;text-transform:uppercase;cursor:pointer;">Approver Note <span style="font-weight:400;color:#9C8080;text-transform:none;letter-spacing:0;">(optional — saved with Draft / Submit)</span></label>'
        + '</div>'
        + '<textarea id="noteBossDecision" rows="3" placeholder="Tulis catatan untuk approver di sini..." style="width:100%;min-height:80px;border:1px solid rgba(139,26,26,0.18);border-radius:10px;padding:11px 14px;font-size:13px;color:#1A0A0A;line-height:1.5;resize:vertical;outline:none;box-sizing:border-box;font-family:Inter,sans-serif;background:#fff;box-shadow:0 1px 4px rgba(139,26,26,0.06);"></textarea>'
        + '<div style="display:flex;justify-content:flex-end;flex-wrap:wrap;gap:10px;margin-top:14px;align-items:center;">'
        + '<button type="button" data-sdd-save="delete" onclick="window._saveScrScreening(\'delete\')" style="padding:9px 20px;border-radius:8px;border:1.5px solid rgba(239,68,68,0.4);background:#fff;color:#EF4444;font-family:Inter,sans-serif;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 1px 4px rgba(239,68,68,0.08);margin-right:auto;">Delete</button>'
        + '<button type="button" id="sdd-cancel-to-draft-btn" data-sdd-save="cancel" style="display:none;padding:9px 20px;border-radius:8px;border:1.5px solid rgba(217,119,6,0.4);background:#fff7ed;color:#9a3412;font-family:Inter,sans-serif;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 1px 4px rgba(217,119,6,0.12);">Cancel</button>'
        + '<button type="button" data-sdd-save="draft" onclick="window._saveScrScreening(\'draft\')" style="padding:9px 20px;border-radius:8px;border:none;background:#F59E0B;color:white;font-family:Inter,sans-serif;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 2px 8px rgba(245,158,11,0.3);">Save as Draft</button>'
        + '<button type="button" data-sdd-save="submit" onclick="window._saveScrScreening(\'submit\')" style="padding:9px 20px;border-radius:8px;border:none;background:#10B981;color:white;font-family:Inter,sans-serif;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 2px 8px rgba(16,185,129,0.3);">Submit</button>'
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

  /** Hide staff-only trace controls for approver accounts. Note/RequestedData visible tapi readonly. */
  window.refreshSddApproverStaffTraceChrome = function() {
    var ap = window._sddUserRole === 'APPROVER';
    document.querySelectorAll('.sdd-hide-for-approver').forEach(function(el) {
      if (el) el.style.display = ap ? 'none' : '';
    });
    // Note & Requested Data: tampil untuk approver tapi readonly
    ['traceRecInput', 'requestedDataInput'].forEach(function(id) {
      var el = document.getElementById(id);
      if (!el) return;
      if (ap) {
        el.readOnly = true;
        el.style.background = '#f8f6f5';
        el.style.color = '#5F4A48';
        el.style.cursor = 'default';
      } else {
        el.readOnly = false;
        el.style.background = '';
        el.style.color = '#1A0A0A';
        el.style.cursor = '';
      }
    });
    // Start Screening button: sembunyikan untuk approver, View Screening tetap tampil
    var startBtn = document.querySelector('#sdd-trace-action-btn-wrap button[onclick*="openTmlScreeningPicker"]');
    if (startBtn) startBtn.style.display = ap ? 'none' : '';

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

    function fmtDate(r) {
      const upd = String(r['updated_at'] || r['SCR - Last Updated'] || '').trim();
      if (!upd) return '';
      const d = new Date(upd);
      if (isNaN(d.getTime())) return upd.slice(0,10);
      return d.toLocaleDateString('id-ID', { day:'2-digit', month:'short', year:'numeric' });
    }

    // Sort by updated_at descending (newest first), then render all — container scrolls
    const sorted = entries.slice().sort(function(a, b) {
      const ta = new Date(String((a.primary||{})['updated_at'] || (a.primary||{})['SCR - Last Updated'] || 0)).getTime() || 0;
      const tb = new Date(String((b.primary||{})['updated_at'] || (b.primary||{})['SCR - Last Updated'] || 0)).getTime() || 0;
      return tb - ta;
    });

    holder.innerHTML = sorted.map(function(entry) {
      const active  = selectedKey && selectedKey === entry.key;
      const r       = entry.primary || {};
      const st      = String(r['SCR - Screening Status'] || '').trim();
      const tp      = normalizeSddSupplierType(r['Supplier Type'] || r['SUPPLIER_TYPE'] || r['SupplierType'] || r['supplier_type']);
      const grp     = String(r['Group Name'] || r['Grup Name'] || '').trim();
      const mill    = String(r['Mill Name'] || '').trim();
      const dateStr = fmtDate(r);

      const isDraft     = st.toLowerCase() === 'draft';
      const statusClass = isDraft ? 'scr-status-draft' : 'scr-status-submitted';
      const statusLabel = isDraft ? 'Draft' : 'Submitted';

      return '<button type="button" data-scr-key="' + escHtml(entry.key) + '" class="scr-saved-item' + (active ? ' selected' : '') + '">'

        // Row 1: status badge + type chip + date
        + '<div style="display:flex;align-items:center;gap:6px;margin-bottom:5px;">'
        + '<span class="scr-saved-status ' + statusClass + '"><span class="scr-status-dot"></span>' + statusLabel + '</span>'
        + (tp ? '<span style="font-size:10.5px;font-weight:700;color:#9C8A8A;letter-spacing:0.07em;text-transform:uppercase;">' + escHtml(tp) + '</span>' : '')
        + '<span style="margin-left:auto;font-size:11px;color:#B09090;white-space:nowrap;">' + (dateStr ? escHtml(dateStr) : '') + '</span>'
        + '</div>'

        // Row 2: group name (bold)
        + '<div style="font-size:13px;font-weight:' + (active ? '700' : '600') + ';color:#1A0A0A;line-height:1.35;' + (mill ? 'margin-bottom:2px;' : '') + '">'
        + (grp ? escHtml(grp) : '<span style="color:#C4BAB4;font-weight:400;">—</span>')
        + '</div>'

        // Row 3: mill name (muted)
        + (mill ? '<div style="font-size:11.5px;color:#9C8A8A;">' + escHtml(mill) + '</div>' : '')

        + '</button>';
    }).join('');

    // hover handled by CSS .scr-saved-item:hover — no inline override needed
    holder.querySelectorAll('button[data-scr-key]').forEach(function(btn) {
      btn.addEventListener('click', function() {
        const key = btn.getAttribute('data-scr-key') || '';
        const sel = document.getElementById('scr-saved-select');
        if (sel) sel.value = key;
        if (typeof window.loadSavedScrByKeyGlobal === 'function') window.loadSavedScrByKeyGlobal(key);
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
      const scrFilter = window._sddUserRole === 'APPROVER' ? 'Submitted' : undefined;
      const listResult = await apiListSubmissions(
        scrFilter ? { scr_status: scrFilter, page_size: 200 } : { page_size: 200 }
      );
      const mainRows = (listResult && Array.isArray(listResult.data)) ? listResult.data : [];

      // For STAFF also include Draft (listSubmissions returns all active by default without filter).
      // When no scr_status filter is passed, we get everything — filter client-side.
      const visibleRows = mainRows.filter(function(r) {
        const st = String(r['SCR - Screening Status'] || '').trim().toLowerCase();
        if (window._sddUserRole === 'APPROVER') return st === 'submitted';
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
      applySddApproverStaffReadOnly(window._sddUserRole === 'APPROVER');
      ['draft', 'submit', 'delete'].forEach(function(k) {
        var b = document.querySelector('[data-sdd-save="' + k + '"]');
        if (b) b.style.display = '';
      });
      var cancelBtn = document.querySelector('[data-sdd-save="cancel"]');
      if (cancelBtn) cancelBtn.style.display = 'none';
      if (typeof window.refreshSddApproverStaffTraceChrome === 'function') window.refreshSddApproverStaffTraceChrome();
    }

    function syncSddApproverDecisionUI(s, sourceLabel, rowNum) {
      var fromSaved = rowNum != null && rowNum !== '' && String(sourceLabel || '').indexOf('SDD') !== -1;
      if (!fromSaved) {
        hideSddApproverPanel();
        return;
      }
      window._sddApproverRecordLoaded = true;
      applySddApproverStaffReadOnly(false);
      ['draft', 'submit', 'delete'].forEach(function(k) {
        var btn = document.querySelector('[data-sdd-save="' + k + '"]');
        if (btn) btn.style.display = 'none';
      });
      var cancelBtn = document.querySelector('[data-sdd-save="cancel"]');
      if (cancelBtn) cancelBtn.style.display = 'none';
      var scrSaveOk = document.getElementById('scr-save-ok');
      if (scrSaveOk) scrSaveOk.style.display = 'none';
      if (typeof window.refreshSddApproverStaffTraceChrome === 'function') window.refreshSddApproverStaffTraceChrome();
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
        }
      } else {
        syncSubmittedStaffLockUI({ status: 'Draft' }, sid);
      }
    }

    function syncSubmittedStaffLockUI(s, sidOrRowNum) {
      if (window._sddUserRole === 'APPROVER') return;
      // fromSaved = true when we have either a submission_id or a legacy row num
      var fromSaved = !!(sidOrRowNum != null && sidOrRowNum !== '' && sidOrRowNum !== 0);
      var isSubmitted = String((s && s.status) || '').toLowerCase() === 'submitted';
      var cancelBtn = document.querySelector('[data-sdd-save="cancel"]');
      if (isSubmitted && fromSaved) {
        applySddApproverStaffReadOnly(true);
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
        // Hide PDF export button when back to Draft
        var pdfBtnDraft = document.getElementById('sdd-export-pdf-btn');
        if (pdfBtnDraft) pdfBtnDraft.style.display = 'none';
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
        grvYN: row['SCR - Grievance (Y/N)'] || '',
        priYN: row['SCR - PRI (Y/N)'] || '',
        traceNote: row['SCR - Notes'] || '',
        requestedData: row['SCR - Requested Data'] || row['Requested Data'] || '',
        attachments: [],
        status: row['SCR - Screening Status'] || '',
        date: row['SCR - Last Updated'] || '',
        noteSdd: row['noteSDD'] || row['noteSdd'] || row['Note SDD'] || '',
        noteBossDecision: row['noteBossDecision'] || row['noteSDD'] || row['noteSdd'] || '',
        statusSdd: row['statusSDD'] || row['statusSdd'] || row['Status SDD'] || '',
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
      syncSddApproverDecisionUI(s, sourceLabel, rowNum);
      // Use submission_id (key) as the "fromSaved" signal; fall back to legacy rowNum
      syncSubmittedStaffLockUI(s, key || rowNum || null);
      const saveOk = document.getElementById('scr-save-ok');
      if (window._sddUserRole !== 'APPROVER' && saveOk && s.status) {
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
      if (window._sddUserRole === 'APPROVER') {
        if (typeof window.showSddToast === 'function') {
          window.showSddToast('Akun approver tidak bisa menyimpan draft/submit di sini. Gunakan catatan dan tombol Approve / Hold / Reject di bawah.', 'info');
        }
        return;
      }
      if (typeof location !== 'undefined' && location.protocol === 'file:') {
        var fileMsg = 'Cannot save while opened as file://. Please run this HTML from http://localhost (for example with Live Server) and try again.';
        if (typeof window.showSddToast === 'function') window.showSddToast(fileMsg, 'error');
        if (typeof window.showSddNotification === 'function') window.showSddNotification('Save Blocked', fileMsg, 'error');
        return;
      }
      const scrData = {
        owners:   document.getElementById('scr-owners')?.value||'',
        news:     document.getElementById('scr-news')?.value||'',
        supplyto: document.getElementById('scr-supplyto')?.value||'',
        legality: document.getElementById('scr-legality')?.value||'',
        cert:     document.getElementById('scr-cert')?.value||'',
        ndpe:     document.getElementById('scr-ndpe')?.value||'',
        nbl:      document.getElementById('scr-nbl')?.value||'',
        grvYN:    document.getElementById('scr-grv-yn')?.value||'',
        priYN:    document.getElementById('scr-pri-yn')?.value||'',
        traceNote: document.getElementById('traceRecInput')?.value||'',
        requestedData: document.getElementById('requestedDataInput')?.value||'',
        noteSdd:  document.getElementById('noteBossDecision')?.value||document.getElementById('noteSDD')?.value||'',
        noteBossDecision: document.getElementById('noteBossDecision')?.value||document.getElementById('noteSDD')?.value||'',
        attachments: window._traceAttachments || [],
        grvRows: [], priRows: [],
      };
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
var SDD_DEFAULT_WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbzRVFziViQwwL7N_3yRSplKfCyS6N9gW4rYqZGSxLYPmiCcKJ84yIbPRpzf9fKhoeI3iA/exec';

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

/**
 * Approver workflow — rewritten to use relational APIs.
 * statusSDD is written via setSubmissionStatus; noteSDD (free-text) via updateSubmission.
 */
window._submitSddApproverDecision = async function(statusSdd) {
  if (window._sddUserRole !== 'APPROVER') {
    if (typeof window.showSddToast === 'function') window.showSddToast('Approver sign-in required.', 'error');
    return;
  }
  var sid = window._sddSubmissionId || window._scrLoadedKey || null;
  if (!sid) {
    if (typeof window.showSddToast === 'function') window.showSddToast('Load a screening record first.', 'error');
    return;
  }
  var ta   = document.getElementById('noteBossDecision') || document.getElementById('noteSDD');
  var note = ta ? String(ta.value || '').trim() : '';
  try {
    // Write statusSDD via lightweight status patch
    await apiSetSubmissionStatus({ submission_id: sid, statusSDD: statusSdd });
    // Write approver fields to relational MAIN row
    await apiUpdateSubmission({
      submission_id: sid,
      main: {
        noteSDD: note,
        noteBossDecision: note,
        statusBossDecision: statusSdd,
      }
    });
    if (window._loadedPrimarySddRow) {
      window._loadedPrimarySddRow['noteSDD']          = note;
      window._loadedPrimarySddRow['noteBossDecision']  = note;
      window._loadedPrimarySddRow['statusSDD']        = statusSdd;
      window._loadedPrimarySddRow['statusBossDecision'] = statusSdd;
    }
    if (window._scrData && typeof window._scrData === 'object') {
      window._scrData.noteSdd           = note;
      window._scrData.statusSdd         = statusSdd;
      window._scrData.noteBossDecision  = note;
      window._scrData.statusBossDecision = statusSdd;
    }
    if (typeof window.refreshSavedScreeningListGlobal === 'function') {
      await window.refreshSavedScreeningListGlobal(sid);
    }
    if (typeof window.refreshSddBossDecisionViewer === 'function') {
      window.refreshSddBossDecisionViewer();
    }
    if (typeof window.showSddToast === 'function') {
      window.showSddToast('Approver decision saved (' + statusSdd + ').', 'success');
    }
    if (typeof window.showSddNotification === 'function') {
      window.showSddNotification('Saved', 'Approver decision saved to Google Sheets.', 'success');
    }
  } catch (e) {
    var msg = (e && e.message) ? e.message : String(e);
    if (typeof window.showSddToast === 'function') window.showSddToast('Save failed: ' + msg, 'error');
    if (typeof window.showSddNotification === 'function') window.showSddNotification('Save Failed', msg, 'error');
  }
};

if (!window.__sddApproverDecisionClickBound) {
  window.__sddApproverDecisionClickBound = true;
  document.addEventListener('click', function(ev) {
    var t = ev.target;
    if (!t || !t.id) return;
    if (t.id === 'sdd-approver-hold' || t.id === 'boss-btn-hold') { ev.preventDefault(); window._submitSddApproverDecision('Hold'); }
    else if (t.id === 'sdd-approver-approve' || t.id === 'boss-btn-approve') { ev.preventDefault(); window._submitSddApproverDecision('Approve'); }
    else if (t.id === 'sdd-approver-reject' || t.id === 'boss-btn-reject') { ev.preventDefault(); window._submitSddApproverDecision('Reject'); }
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
    'SCR - Grievance (Y/N)': scrData.grvYN || '',
    'SCR - PRI (Y/N)': scrData.priYN || '',
    'SCR - Notes': scrData.traceNote || '',
    'SCR - Requested Data': scrData.requestedData || '',
    'noteSDD': scrData.noteSdd || '',
    'noteBossDecision': scrData.noteBossDecision || scrData.noteSdd || '',
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

  const MILL_FIELDS = ['QUARTER','YEAR','COMPANY CODE','TRADER NAME','GROUP NAME','COMPANY NAME','MILL NAME','UML ID','ADDRESS','PROVINCE','COORDINATES','MILL CATEGORY','MILL CAPACITY (TON/HOUR)','HGU/HGB','IZIN LOKASI','IUP','IZIN LINGKUNGAN','SCORE','MILL LOC','COMPLIMENT/NOT COMPLIMENT','DEFORESTATION SPATIAL','BURN AREA SPATIAL','PEAT','LEGALITY','DEFORESTATION GRIEVANCES','BURN AREA GRIEVANCES','HUMAN RIGHT','SAFETY','SOCIAL','ENVIRONMENT','TOTAL GRIEVANCES','NDPE','HRDD','TOTAL POLICY','CERTIFICATION','TOTAL CERTIFICATION','TOTAL SCORE','SUPPLIER LEVEL','BUYER NO BUY LIST','VOLUME SUPPLY STATUS','RECOMMENDATION LEVEL','SIGN','SUPPLIER STATUS','RISK LEVEL','FACILITY NAME CPO','FACILITY NAME PK','PRODUCT SUPPLY'];
  let modalSheet = '', modalMode = '', modalRow = null, modalFields = [];
  let modalTaskKey = ''; // submission_id dari SDD yang di-add via Task List
  let allData = [];
  let currentFilter = 'All';
  let currentSearch = '';
  let ttpData = [], ttpFields = [], ttpLoaded = false, ttpPctCol = '', ttpSearch = '';
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
    { title: 'Supply & Status', fields: ['SUPPLIER LEVEL','SUPPLIER STATUS','BUYER NO BUY LIST','VOLUME SUPPLY STATUS','RECOMMENDATION LEVEL','SIGN','RISK LEVEL','FACILITY NAME CPO','FACILITY NAME PK','PRODUCT SUPPLY'] },
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
        await loadMillData();
      } else if (modalSheet === 'ttp') {
        ttpLoaded = false; await loadTTPData();
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
  let millSortKey = null;
  let millSortAsc = true;
  let millTableDelegationBound = false;

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
    ].map(function(v) {
      return String(v || '').toLowerCase();
    }).join('|');
    row._sddNblLower = nblLower;
    row._sddSearchBlob = searchBlob;
    return row;
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
    { key: 'BUYER NO BUY LIST', label: 'No Buy List' },
    { key: 'CERTIFICATION', label: 'Certification' },
    { key: 'FACILITY NAME CPO', label: 'Facility CPO' },
    { key: 'FACILITY NAME PK', label: 'Facility PK' },
    { key: 'PRODUCT SUPPLY', label: 'Product Supply' },
  ];
  const MILL_PDF_COL_DEFAULT_KEYS = ['QUARTER', 'YEAR', 'GROUP NAME', 'COMPANY NAME', 'MILL NAME', 'PROVINCE', 'SUPPLIER STATUS', 'RISK LEVEL', 'BUYER NO BUY LIST', 'CERTIFICATION'];
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

  function getMillRowsForPdfExport() {
    const filtered = allData.filter(function(d) {
      return millRowMatchesChipAndSearch(d) && millRowMatchesPdfDimFilters(d);
    });
    return sortMillRowsForDisplay(filtered);
  }

  function updateMillPdfExportScope() {
    const el = document.getElementById('millPdfExportScopeText');
    if (!el) return;
    if (!allData.length) {
      el.textContent = 'Muat data mill terlebih dahulu';
      return;
    }
    const tableN = sortMillRowsForDisplay(allData.filter(millRowMatchesChipAndSearch)).length;
    const pdfN = getMillRowsForPdfExport().length;
    if (tableN === pdfN) {
      el.textContent = pdfN + ' baris · sama dengan tabel';
    } else {
      el.textContent = pdfN + ' baris untuk PDF · ' + tableN + ' di tabel';
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
  }

  function millPdfResetAllDims() {
    millPdfDimFilters = { quarter: new Set(), year: new Set(), group: new Set(), province: new Set() };
    document.querySelectorAll('#millPdfFilterDimsPanel input[data-mill-pdf-val]').forEach(function(cb) { cb.checked = false; });
    updateMillPdfExportScope();
  }

  function millExportToPdf() {
    const toastErr = function(msg) {
      if (typeof window.showSddToast === 'function') window.showSddToast(msg, 'error');
    };
    const rows = getMillRowsForPdfExport();
    if (!rows.length) {
      toastErr('Tidak ada baris untuk diekspor. Sesuaikan chip, pencarian, atau saringan export.');
      return;
    }
    const cols = MILL_PDF_EXPORT_COLS.filter(function(c) { return millPdfColSelected.has(c.key); });
    if (!cols.length) {
      toastErr('Pilih minimal satu kolom di menu Kolom PDF.');
      return;
    }
    const JsPDFLib = (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF : window.jsPDF;
    if (!JsPDFLib) {
      toastErr('Library jsPDF belum dimuat. Refresh halaman.');
      return;
    }
    let probe;
    try {
      probe = new JsPDFLib({ unit: 'mm', format: 'a4', orientation: 'landscape' });
    } catch (e) {
      toastErr('Gagal membuat PDF: ' + (e.message || e));
      return;
    }
    if (typeof probe.autoTable !== 'function') {
      toastErr('Plugin AutoTable belum dimuat. Refresh halaman.');
      return;
    }

    const btn = document.getElementById('btn-mill-export-pdf');
    const prevHtml = btn ? btn.innerHTML : '';
    if (btn) { btn.disabled = true; btn.innerHTML = '<span class="ttp-btn-icon">…</span> Menghasilkan…'; }

    try {
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
          const v = row[c.key];
          return String(v !== undefined && v !== null ? v : '').replace(/\r?\n/g, ' ');
        });
      });
      doc.autoTable({
        head: [cols.map(function(c) { return c.label; })],
        body: body,
        startY: 24,
        margin: { left: 10, right: 10 },
        styles: { fontSize: 6.2, cellPadding: 1.4, textColor: [26, 10, 10] },
        headStyles: { fillColor: RED, textColor: WHITE, fontStyle: 'bold', fontSize: 7 },
        alternateRowStyles: { fillColor: [253, 250, 250] },
        theme: 'striped',
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
      document.getElementById('stat-high-risk').textContent = allData.filter(d => (d['RISK LEVEL']||'').toLowerCase().includes('high')).length;
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
      ? `<span class="status-badge" style="background:rgba(192,57,43,0.1);color:#c0392b"><span class="s-dot"></span>${val}</span>`
      : `<span class="status-badge s-active"><span class="s-dot"></span>${val}</span>`;
  }

  function supplierBadge(val) {
    if (!val) return '—';
    const lower = val.toLowerCase();
    let cls = 's-pending';
    if (['active','compliant'].some(k => lower.includes(k))) cls = 's-active';
    else if (['review','pending','conditional'].some(k => lower.includes(k))) cls = 's-review';
    return `<span class="status-badge ${cls}"><span class="s-dot"></span>${val}</span>`;
  }

  function riskBadgeLevel(val) {
    if (!val || val === '-') return '<span style="color:#ccc">—</span>';
    const lower = val.toString().toLowerCase();
    let bg = 'rgba(124,101,101,0.1)', color = '#7C6565';
    if (lower.includes('high') || lower.includes('tinggi')) { bg = 'rgba(192,57,43,0.1)'; color = '#c0392b'; }
    else if (lower.includes('med') || lower.includes('sedang')) { bg = 'rgba(200,168,75,0.15)'; color = '#8a6e1a'; }
    else if (lower.includes('low') || lower.includes('rendah')) { bg = 'rgba(39,174,96,0.1)'; color = '#1e8449'; }
    return `<span class="status-badge" style="background:${bg};color:${color}"><span class="s-dot"></span>${val}</span>`;
  }

  function renderMillTable() {
    const body = document.getElementById('millTableBody');
    if (!body) return;
    bindMillTableDelegationOnce();
    const filtered = allData.filter(d => millRowMatchesChipAndSearch(d));
    const sorted = sortMillRowsForDisplay(filtered);
    millFilteredRows = sorted;
    updateMillPdfExportScope();

    const theadRow = document.querySelector('#millTable thead tr');
    if (theadRow) {
      theadRow.querySelectorAll('[data-mill-sort]').forEach(function(th) {
        th.classList.remove('is-sorted', 'is-sorted-asc', 'is-sorted-desc');
        if (millSortKey && th.getAttribute('data-mill-sort') === millSortKey) {
          th.classList.add('is-sorted', millSortAsc ? 'is-sorted-asc' : 'is-sorted-desc');
        }
      });
    }

    body.innerHTML = sorted.length === 0
      ? `<tr><td colspan="13" style="text-align:center;padding:32px;color:#9C8A8A;">No data found</td></tr>`
      : sorted.map((d, i) => `
        <tr class="mill-row-clickable" data-idx="${i}" title="Klik untuk lihat detail lengkap">
          <td>${millQuarterVal(d) || '—'}</td>
          <td>${millYearVal(d) || '—'}</td>
          <td>${d['GROUP NAME'] || '—'}</td>
          <td>${d['COMPANY NAME'] || '—'}</td>
          <td><span class="mill-name">${d['MILL NAME'] || '—'}</span><div class="mill-id">${d['UML ID'] || ''}</div></td>
          <td>${d['PROVINCE'] || '—'}</td>
          <td>${supplierBadge(d['SUPPLIER STATUS'])}</td>
          <td>${riskBadgeLevel(d['RISK LEVEL'])}</td>
          <td>${nblBadge(d['BUYER NO BUY LIST'])}</td>
          <td>${d['CERTIFICATION'] || '—'}</td>
          <td class="td-truncate" title="${d['FACILITY NAME CPO'] || ''}">${d['FACILITY NAME CPO'] ? (d['FACILITY NAME CPO'].length > 18 ? d['FACILITY NAME CPO'].substring(0,18)+'…' : d['FACILITY NAME CPO']) : '—'}</td>
          <td class="td-truncate" title="${d['FACILITY NAME PK'] || ''}">${d['FACILITY NAME PK'] ? (d['FACILITY NAME PK'].length > 18 ? d['FACILITY NAME PK'].substring(0,18)+'…' : d['FACILITY NAME PK']) : '—'}</td>
          <td class="td-truncate" title="${d['PRODUCT SUPPLY'] || ''}">${d['PRODUCT SUPPLY'] ? (d['PRODUCT SUPPLY'].length > 18 ? d['PRODUCT SUPPLY'].substring(0,18)+'…' : d['PRODUCT SUPPLY']) : '—'}</td>
        </tr>`).join('');
  }

  const btnAddMill = document.getElementById('btn-add-mill');
  if (btnAddMill) btnAddMill.addEventListener('click', () => openModal('mill', MILL_FIELDS, 'add', null));

  // ─── MILL PROFILE POPUP ─────────────────────────────────
  function openMillProfile(d) {
    document.getElementById('mp-mill-name').textContent = d['MILL NAME'] || '—';
    document.getElementById('mp-mill-sub').textContent =
      [d['UML ID'], d['GROUP NAME'], d['PROVINCE']].filter(Boolean).join(' • ');

    const sections = [
      {
        title: 'Identitas Mill',
        fields: [
          ['QUARTER','Quarter'], ['YEAR','Year'],
          ['COMPANY CODE','Company Code'], ['TRADER NAME','Trader Name'], ['GROUP NAME','Group Name'],
          ['COMPANY NAME','Company Name'], ['MILL NAME','Mill Name'], ['UML ID','UML ID'],
          ['PROVINCE','Province'],
          ['ADDRESS','Address'], ['COORDINATES','Coordinates'],
          ['MILL CATEGORY','Mill Category'], ['MILL CAPACITY (TON/HOUR)','Capacity (Ton/Hour)'],
        ]
      },
      {
        title: 'Legalitas',
        fields: [
          ['HGU/HGB','HGU/HGB'], ['IZIN LOKASI','Izin Lokasi'], ['IUP','IUP'],
          ['IZIN LINGKUNGAN','Izin Lingkungan'], ['LEGALITY','Legality'],
        ]
      },
      {
        title: 'Skor & Sertifikasi',
        fields: [
          ['SCORE','Score'], ['MILL LOC','Mill Loc'],
          ['COMPLIMENT/NOT COMPLIMENT','Compliment/Not Compliment'],
          ['CERTIFICATION','Certification'], ['TOTAL CERTIFICATION','Total Certification'],
          ['TOTAL SCORE','Total Score'],
        ]
      },
      {
        title: 'Spatial & Lingkungan',
        fields: [
          ['DEFORESTATION SPATIAL','Deforestation Spatial'], ['BURN AREA SPATIAL','Burn Area Spatial'],
          ['PEAT','Peat'],
        ]
      },
      {
        title: 'Grievances',
        fields: [
          ['DEFORESTATION GRIEVANCES','Deforestation Grievances'], ['BURN AREA GRIEVANCES','Burn Area Grievances'],
          ['HUMAN RIGHT','Human Right'], ['SAFETY','Safety'], ['SOCIAL','Social'],
          ['ENVIRONMENT','Environment'], ['TOTAL GRIEVANCES','Total Grievances'],
        ]
      },
      {
        title: 'Policy & Supply',
        fields: [
          ['NDPE','NDPE'], ['HRDD','HRDD'], ['TOTAL POLICY','Total Policy'],
          ['SUPPLIER LEVEL','Supplier Level'], ['BUYER NO BUY LIST','No Buy List'],
          ['VOLUME SUPPLY STATUS','Volume Supply Status'], ['RECOMMENDATION LEVEL','Recommendation Level'],
          ['SIGN','Sign'], ['SUPPLIER STATUS','Supplier Status'],
          ['FACILITY NAME CPO','Facility Name CPO'], ['FACILITY NAME PK','Facility Name PK'],
          ['PRODUCT SUPPLY','Product Supply'],
        ]
      },
    ];

    const mpBody = document.getElementById('millProfileBody');
    if (!mpBody) return;
    mpBody.innerHTML = sections.map(sec => `
      <div class="mp-section">
        <div class="mp-section-title">${sec.title}</div>
        <div class="mp-grid${sec.fields.length <= 4 ? ' cols2' : ''}">
          ${sec.fields.map(([key, label]) => {
            const val = d[key] || '';
            const isLong = key === 'ADDRESS' || key === 'COORDINATES';
            return `<div class="mp-field${isLong ? ' full' : ''}">
              <div class="mp-label">${label}</div>
              <div class="mp-val">${val || '—'}</div>
            </div>`;
          }).join('')}
        </div>
      </div>`).join('');

    document.getElementById('millProfileOverlay')?.classList.add('active');
  }

  (function bindMillProfileOverlay() {
    const mpc = document.getElementById('millProfileClose');
    const mpo = document.getElementById('millProfileOverlay');
    if (!mpc || !mpo) {
      console.warn('[dashboard] Mill profile overlay nodes missing.');
      return;
    }
    mpc.addEventListener('click', () => {
      mpo.classList.remove('active');
    });
    mpo.addEventListener('click', function(e) {
      if (e.target === this) this.classList.remove('active');
    });
  })();

  // ─── TTP DATA ───────────────────────────────────────────
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
      ttpUniqueValuesCache = Object.create(null);
      document.getElementById('ttp-stat-total').textContent = ttpData.length;
      document.getElementById('ttp-stat-loaded').textContent = ttpData.length;
      // Traceability % column: legacy "PERCENTAGE TRACEABILITY" or new "% CPO / % PK TRACEABLE" headers
      ttpPctCol = (function pickTtpPctCol(fields) {
        if (!fields || !fields.length) return '% CPO TRACEABLE';
        const U = function (h) { return String(h || '').toUpperCase(); };
        const cpo = fields.find(function (h) { return U(h).includes('% CPO TRACEABLE'); });
        if (cpo) return cpo;
        const pk = fields.find(function (h) { return U(h).includes('% PK TRACEABLE'); });
        if (pk) return pk;
        const legacy = fields.find(function (h) {
          return U(h).includes('PERCENTAGE TRACEABILITY') ||
            U(h) === 'PERCENTAGE TRACEABILITY' ||
            h.toLowerCase().includes('percentage');
        });
        if (legacy) return legacy;
        // Header label fallback (values resolve only if sheet uses this exact column name)
        return '% CPO TRACEABLE';
      })(ttpFields);
      document.getElementById('ttpTableHead').innerHTML =
        '<tr><th>Group Name</th><th>Company Name</th><th>Mill Name</th><th>' + (ttpPctCol || 'Traceability %') + '</th><th></th></tr>';
      loading.style.display = 'none';
      table.style.display = 'table';
      // reset selection state on fresh load
      ttpSelectedCompanies = null;
      ttpVisibleCols = null;
      buildCompanyDropdown();
      buildColumnDropdown();
      document.getElementById('btn-export-ttp-xlsx').disabled = false;
      scheduleRenderTTPTable();
    } catch(err) {
      loading.style.display = 'none';
      errorEl.style.display = 'block';
      errorEl.textContent = 'Gagal memuat data: ' + err.message;
    }
  }

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
  let ttpViewMode = 'grouped'; // 'grouped' | 'flat'
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

      const expandBtn = e.target.closest('.ttp-expand');
      if (expandBtn && body.contains(expandBtn)) {
        e.stopPropagation();
        const idx = expandBtn.dataset.idx;
        const detail = document.getElementById('ttp-detail-' + idx);
        if (!detail) return;
        const isOpen = detail.classList.contains('open');
        document.querySelectorAll('.grv-detail').forEach(d => d.classList.remove('open'));
        document.querySelectorAll('.btn-expand').forEach(b => b.classList.remove('open'));
        if (!isOpen) {
          detail.classList.add('open');
          expandBtn.classList.add('open');
        }
        return;
      }

      const groupRow = e.target.closest('.ttp-group-row');
      if (groupRow && body.contains(groupRow)) {
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
  }

  function renderTTPTable() {
    if (ttpViewMode === 'grouped') {
      renderTTPGrouped();
    } else {
      renderTTPFlat();
    }
  }

  // ── GROUPED VIEW ─────────────────────────────────────────
  function renderTTPGrouped() {
    const body = document.getElementById('ttpTableBody');
    if (!body) return;
    bindTtpTableDelegationOnce();

    const groupCol   = ttpFields.find(h => h.toLowerCase().includes('group')) || '';
    const companyCol = ttpFields.find(h => h.toUpperCase() === 'COMPANY NAME') || ttpFields.find(h => h.toLowerCase().includes('company name')) || '';
    const millCol    = ttpFields.find(h => h.toUpperCase().includes('MILL NAME') || h.toLowerCase() === 'mill name') || ttpFields.find(h => h.toLowerCase().includes('mill')) || '';
    const villageCol = ttpFields.find(h => /village|desa|kebun|estate|location|lokasi/i.test(h)) || '';

    // Columns to show inline in child row (exclude the 4 main cols already shown)
    const mainCols = new Set([groupCol, companyCol, millCol, ttpPctCol].filter(Boolean));
    const extraCols = (ttpVisibleCols && ttpVisibleCols.size > 0)
      ? ttpFields.filter(f => ttpVisibleCols.has(f) && !mainCols.has(f))
      : ttpFields.filter(f => !mainCols.has(f));

    // Fixed header — no "Detail" column, fixed widths that don't depend on content
    document.getElementById('ttpTableHead').innerHTML = `<tr>
      <th style="width:160px;min-width:130px">Group Name</th>
      <th style="width:200px;min-width:160px">Company Name</th>
      <th style="min-width:200px">Mill Name</th>
      <th style="width:160px;min-width:120px">${ttpPctCol || 'Traceability %'}</th>
      <th style="width:90px;min-width:80px"></th>
    </tr>`;

    // Apply filters
    let filtered = ttpData.filter(d =>
      !ttpSearch || (d._sddSearchBlob || '').includes(ttpSearch)
    );
    if (ttpSelectedCompanies !== null) {
      filtered = filtered.filter(d => ttpSelectedCompanies.has((d[ttpCompanyCol] || '').toString()));
    }

    if (filtered.length === 0) {
      body.innerHTML = `<tr><td colspan="5" style="text-align:center;padding:32px;color:#9C8A8A;">No data found</td></tr>`;
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

      const avgPct = ttpPctCol ? (() => {
        const nums = rows.map(r => parseFloat(r[ttpPctCol])).filter(n => !isNaN(n));
        if (!nums.length) return '—';
        const avg = nums.reduce((a,b) => a+b, 0) / nums.length;
        const hasPct = (rows[0][ttpPctCol]||'').toString().includes('%');
        return (Number.isInteger(avg) ? avg : avg.toFixed(1)) + (hasPct ? '%' : '');
      })() : '—';

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
        <td><span style="font-weight:600;color:var(--forest)">${avgPct}</span>${subCount > 1 ? '<span class="ttp-group-meta"> avg</span>' : ''}</td>
        <td></td>
      </tr>`;

      // ── Child rows — clean horizontal layout ──
      rows.forEach((d, ri) => {
        const pct        = ttpPctCol ? (d[ttpPctCol] !== undefined && d[ttpPctCol] !== '' ? d[ttpPctCol] : '—') : '—';
        const childLabel = villageCol && d[villageCol] ? d[villageCol] : ('Record #' + (ri + 1));

        // Pick up to 2 most informative extra fields to show as subtle text in each cell
        const extraInfo = extraCols.slice(0, 6).filter(f => {
          const v = (d[f] || '').toString().trim();
          return v && v !== '—' && v !== '-';
        }).map(f => `<span class="ttp-child-meta">${f}: <strong>${d[f]}</strong></span>`).join('');

        html += `<tr class="ttp-child-row hidden" data-parent="${groupId}">
          <td style="padding-left:12px;">
            <span style="color:#D0B8B8;font-size:13px;line-height:1;">└</span>
          </td>
          <td style="color:var(--muted);font-size:11px;">${ri + 1} / ${rows.length}</td>
          <td class="ttp-child-indent">
            <div class="ttp-child-village">${childLabel}</div>
            ${extraInfo ? `<div class="ttp-child-extra">${extraInfo}</div>` : ''}
          </td>
          <td style="font-weight:500;color:var(--text-light);font-size:12px;">${pct}</td>
          <td>
            <div class="row-actions">
              <button class="btn-row btn-edit" data-row='${JSON.stringify(d).replace(/'/g,"&#39;")}' data-sheet="ttp">Edit</button>
              <button class="btn-row btn-delete" data-rownum="${d._row}" data-sheet="ttp">Del</button>
            </div>
          </td>
        </tr>`;
      });

      gIdx++;
    });

    body.innerHTML = html;

  }

  // ── FLAT VIEW (original behaviour) ────────────────────
  function renderTTPFlat() {
    const body = document.getElementById('ttpTableBody');
    if (!body) return;
    bindTtpTableDelegationOnce();

    let filtered = ttpData.filter(d =>
      !ttpSearch || (d._sddSearchBlob || '').includes(ttpSearch)
    );
    if (ttpSelectedCompanies !== null) {
      filtered = filtered.filter(d => ttpSelectedCompanies.has((d[ttpCompanyCol] || '').toString()));
    }

    const colsForDetail = (ttpVisibleCols && ttpVisibleCols.size > 0) ? ttpFields.filter(f => ttpVisibleCols.has(f)) : ttpFields;
    const groupCol   = ttpFields.find(h => h.toLowerCase().includes('group')) || '';
    const companyCol = ttpFields.find(h => h.toUpperCase() === 'COMPANY NAME') || ttpFields.find(h => h.toLowerCase().includes('company name')) || '';
    const millCol    = ttpFields.find(h => h.toUpperCase().includes('MILL NAME') || h.toLowerCase() === 'mill name') || ttpFields.find(h => h.toLowerCase().includes('mill')) || '';

    document.getElementById('ttpTableHead').innerHTML =
      '<tr><th>Group Name</th><th>Company Name</th><th>Mill Name</th><th>' + (ttpPctCol || 'Traceability %') + '</th><th></th></tr>';
    document.getElementById('ttpGroupInfo').textContent = filtered.length + ' records';

    if (filtered.length === 0) {
      body.innerHTML = `<tr><td colspan="5" style="text-align:center;padding:32px;color:#9C8A8A;">No data found</td></tr>`;
      return;
    }

    body.innerHTML = filtered.map((d, i) => {
      const detailHTML = colsForDetail.map(f => `
        <div class="grv-detail-item">
          <div class="grv-detail-label">${f}</div>
          <div class="grv-detail-val">${d[f] !== undefined && d[f] !== '' ? d[f] : '—'}</div>
        </div>`).join('');
      const pct = ttpPctCol ? (d[ttpPctCol] !== undefined && d[ttpPctCol] !== '' ? d[ttpPctCol] : '—') : '—';
      return `
        <tr class="mill-row-clickable ttp-main-row" data-idx="${i}">
          <td>${d[groupCol] || '—'}</td>
          <td>${d[companyCol] || '—'}</td>
          <td><span class="mill-name">${d[millCol] || '—'}</span></td>
          <td>${pct}</td>
          <td>
            <div class="row-actions">
              <button class="btn-expand ttp-expand" data-idx="${i}" title="View detail">▾</button>
              <button class="btn-row btn-edit" data-row='${JSON.stringify(d).replace(/'/g,"&#39;")}' data-sheet="ttp">Edit</button>
              <button class="btn-row btn-delete" data-rownum="${d._row}" data-sheet="ttp">Del</button>
            </div>
          </td>
        </tr>
        <tr class="grv-expand-row">
          <td colspan="5">
            <div class="grv-detail" id="ttp-detail-${i}" style="grid-template-columns:1fr 1fr 1fr;">${detailHTML}</div>
          </td>
        </tr>`;
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
      this.classList.add('active');
      document.getElementById('ttpViewFlat')?.classList.remove('active');
      scheduleRenderTTPTable();
    });
    ttpViewFlat.addEventListener('click', function() {
      ttpViewMode = 'flat';
      this.classList.add('active');
      document.getElementById('ttpViewGrouped')?.classList.remove('active');
      scheduleRenderTTPTable();
    });

    ttpBtnSelect.addEventListener('click', function(e) {
      e.stopPropagation();
      openTTPDropdown('ttpBtnSelect', 'ttpSelectPanel', 'ttpFilterPanel', 'ttpBtnFilter');
    });
    ttpBtnFilter.addEventListener('click', function(e) {
      e.stopPropagation();
      openTTPDropdown('ttpBtnFilter', 'ttpFilterPanel', 'ttpSelectPanel', 'ttpBtnSelect');
    });
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

      const colsToExport = ttpVisibleCols && ttpVisibleCols.size > 0
        ? ttpFields.filter(f => ttpVisibleCols.has(f))
        : ttpFields;

      const rowsToExport = ttpData.filter(d => {
        if (ttpSelectedCompanies === null) return true;
        const co = (d[ttpCompanyCol] || '').toString();
        return ttpSelectedCompanies.has(co);
      }).filter(d =>
        !ttpSearch || (d._sddSearchBlob || '').includes(ttpSearch)
      );

      if (!rowsToExport.length) { alert('Tidak ada data untuk di-export.'); return; }

      if (typeof XLSX === 'undefined') {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js';
        script.onload = () => doExportXLSX(rowsToExport, colsToExport);
        document.head.appendChild(script);
      } else {
        doExportXLSX(rowsToExport, colsToExport);
      }
    });
  })();

  // ─── TTP SELECT (dynamic column-based filter) ─────────────
  let ttpSelectedCompanies = null; // null = all selected (values of the active filter col)
  let ttpVisibleCols = null;       // null = all columns
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
    updateTTPSelectionInfo();
  }

  function updateTTPSelectionInfo() {
    const info = document.getElementById('ttpSelectionInfo');
    const txt = document.getElementById('ttpSelectionText');
    const exportBtn = document.getElementById('btn-export-ttp-xlsx');
    if (!info || !txt || !exportBtn) return;
    const allValues = ttpData.length ? getUniqueValuesForCol(ttpActiveFilterCol) : [];
    const selCount = ttpSelectedCompanies ? ttpSelectedCompanies.size : allValues.length;
    const colCount = ttpVisibleCols ? ttpVisibleCols.size : ttpFields.length;
    const isDefault = selCount === allValues.length && colCount === ttpFields.length;
    const colLabel = ttpActiveFilterCol || 'nilai';
    if (!isDefault || selCount < allValues.length || colCount < ttpFields.length) {
      info.style.display = 'flex';
      txt.innerHTML = `<span class="ttp-badge">${selCount} ${colLabel.toLowerCase()}</span> · <span class="ttp-badge">${colCount} kolom</span>`;
    } else {
      info.style.display = 'none';
    }
    exportBtn.disabled = false;
    const selBtn = document.getElementById('ttpBtnSelect');
    const filBtn = document.getElementById('ttpBtnFilter');
    if (selBtn) selBtn.classList.toggle('active', ttpSelectedCompanies !== null && ttpSelectedCompanies.size < allValues.length);
    if (filBtn) filBtn.classList.toggle('active', ttpVisibleCols !== null && ttpVisibleCols.size < ttpFields.length);
  }

  // ─── TTP FILTER COLUMNS ───────────────────────────────────
  function buildColumnDropdown() {
    if (!ttpFields.length) return;
    if (ttpVisibleCols === null) ttpVisibleCols = new Set(ttpFields);

    const list = document.getElementById('ttpColumnList');
    if (!list) return;
    list.innerHTML = ttpFields.map(f => {
      const checked = ttpVisibleCols.has(f) ? 'checked' : '';
      const id = 'ttp-col-' + f.replace(/[^a-zA-Z0-9]/g,'_').substring(0,20);
      return `<div class="ttp-dropdown-item">
        <input type="checkbox" id="${id}" value="${f.replace(/"/g,'&quot;')}" ${checked} onchange="ttpToggleCol(this)">
        <label for="${id}">${f}</label>
      </div>`;
    }).join('');

    if (!ttpColumnDropdownListenersBound) {
      const ttpColAll = document.getElementById('ttpColAll');
      const ttpColNone = document.getElementById('ttpColNone');
      if (!ttpColAll || !ttpColNone) return;
      ttpColumnDropdownListenersBound = true;
      ttpColAll.addEventListener('click', () => {
        ttpVisibleCols = new Set(ttpFields);
        buildColumnDropdown();
        scheduleRenderTTPTable();
        updateTTPSelectionInfo();
      });
      ttpColNone.addEventListener('click', () => {
        ttpVisibleCols = new Set();
        buildColumnDropdown();
        scheduleRenderTTPTable();
        updateTTPSelectionInfo();
      });
    }
  }

  window.ttpToggleCol = function(cb) {
    if (!ttpVisibleCols) ttpVisibleCols = new Set(ttpFields);
    if (cb.checked) ttpVisibleCols.add(cb.value);
    else ttpVisibleCols.delete(cb.value);
    scheduleRenderTTPTable();
    updateTTPSelectionInfo();
  };

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

  function doExportXLSX(rows, cols) {
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
    XLSX.utils.book_append_sheet(wb, ws, 'TTP-TTM');

    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
    XLSX.writeFile(wb, `monitoring_ttp_ttm_${stamp}.xlsx`, { cellStyles: true });
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
    resetScrollToTopEverywhere();
  }

  function switchPanel(name) {
    panelEls.forEach(function(p) { p.classList.remove('active'); });
    navPanelEls.forEach(function(n) { n.classList.remove('active'); });
    const panel = document.getElementById('panel-' + name);
    const navItem = document.querySelector('[data-panel="' + name + '"]');
    if (panel) panel.classList.add('active');
    if (navItem) navItem.classList.add('active');
    const grp = document.getElementById('navGroupTrace');
    if (grp && !grp.classList.contains('open')) grp.classList.add('open');
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
    resetScrollToTopEverywhere();
  }
  // expose globally for onclick handlers
  window.switchPanel = switchPanel;

  // ─── LOGIN (simple page, no modal needed) ─────────────────────────────────────────
  // focus email on load
  setTimeout(() => { const e = document.getElementById('loginEmail'); if (e) e.focus(); }, 100);

  // ─── LOGIN / LOGOUT ─────────────────────────────────────
  function resolveSddRoleFromSupabaseUser_(user) {
    if (!user) return 'STAFF';
    var um = user.user_metadata || {};
    var am = user.app_metadata || {};
    var r = String(um.sdd_role || um.role || am.sdd_role || '').trim().toUpperCase();
    if (r === 'APPROVER') return 'APPROVER';
    return 'STAFF';
  }

  async function finalizeSuccessfulLogin_(displayEmail, role) {
    window._sddUserRole = role || 'STAFF';
    document.body.classList.toggle('sdd-role-approver', role === 'APPROVER');
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
        const millAdded = String(r['mill_added'] || '').toLowerCase();
        return status === 'submitted' && millAdded !== 'true';
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
  function sddExportPdf() {
    var status = String(
      (window._scrData && window._scrData.status) ||
      (window._loadedPrimarySddRow && window._loadedPrimarySddRow['SCR - Screening Status']) || ''
    ).trim().toLowerCase();
    if (status !== 'submitted') {
      if (typeof window.showSddToast === 'function') window.showSddToast('Export PDF hanya tersedia untuk screening dengan status Submitted.', 'error');
      return;
    }
    if (typeof window.jspdf === 'undefined' && typeof window.jsPDF === 'undefined') {
      if (typeof window.showSddToast === 'function') window.showSddToast('Library jsPDF belum dimuat.', 'error');
      return;
    }

    var pdfBtn = document.getElementById('sdd-export-pdf-btn');
    if (pdfBtn) { pdfBtn.disabled = true; pdfBtn.innerHTML = 'Generating…'; }

    try {
      var jsPDFLib = (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF : window.jsPDF;
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
        grvYN       : val('scr-grv-yn',    'grvYN',    'SCR - Grievance (Y/N)'),
        priYN       : val('scr-pri-yn',    'priYN',    'SCR - PRI (Y/N)'),
        note        : val('traceRecInput', 'traceNote','SCR - Notes'),
        reqData     : val('requestedDataInput','requestedData',['SCR - Requested Data','Requested Data']),
        approverNote: val('noteBossDecision','noteBossDecision',['noteBossDecision','noteSDD','noteSdd']),
        approverStat: val(null,'statusSdd',['statusSDD','statusSdd','Status SDD']),
        status      : (sc.status || p['SCR - Screening Status'] || '').toUpperCase(),
        date        : (sc.date || p['SCR - Last Updated'] || p['updated_at'] || '')
      };

      var sid          = String(window._sddSubmissionId || window._scrLoadedKey || '').trim();
      var supplierType = String(window._sddSupplierType || p['Supplier Type'] || '').toUpperCase();
      var supplierName = String(p['Company Name'] || p['Mill Name'] || '').trim() || '—';
      var exportedAt   = new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });

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
      doc.text('CONFIDENTIAL  ·  Status: SUBMITTED  ·  Exported: ' + exportedAt, mL, 21);
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
      // 2. SCREENING SUMMARY
      // ═══════════════════════════════════════════════════════════════════════
      S('Screening Summary');
      kv ('Group / Owners',  f.owners   || '—');
      kv ('Previous News',   f.news     || '—');
      kv ('Supply To',       f.supplyto || '—');
      kv2('Legality Status', f.legality || '—',  'Certification', f.cert  || '—');
      kv2('NDPE Policy',     f.ndpe     || '—',  'No Buy List',   f.nbl   || '—');
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