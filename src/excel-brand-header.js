/**
 * Shared Excel branding block for dashboard exports.
 * Matches the SPC palm-oil traceability submission template header.
 */

export const EXCEL_REPORT_TITLE = 'PALM OIL TRACEABILITY DATA SUBMISSION';

export const EXCEL_COMPANY_INFO = {
  name: 'SUMBER PANGAN CEMERLANG',
  country: 'INDONESIA',
  address: 'JL RAYA LUBUK GAUNG, SUNGAI SEMBILAN, DUMAI, RIAU',
  latitude: '1.742806',
  longitude: '101.371881',
};

/** Rows above the column-header line (title … blank before table). */
export function excelBrandPreambleRowCount_() {
  return 9;
}

function companyLines_() {
  const info = EXCEL_COMPANY_INFO;
  return [
    'NAME: ' + info.name,
    'COUNTRY: ' + info.country,
    'ADDRESS: ' + info.address,
    'LATITUDE: ' + info.latitude,
    'LONGITUDE: ' + info.longitude,
  ];
}

/**
 * Build a SheetJS worksheet with company header + maroon table header + data.
 * Layout (0-based rows):
 *   0 title (merged across all columns, centered, bordered)
 *   1 blank
 *   2 COMPANY INFORMATION
 *   3–7 NAME / COUNTRY / ADDRESS / LATITUDE / LONGITUDE (single cells in col A)
 *   8 blank
 *   9 table headers
 *   10+ data
 *
 * @param {typeof window.XLSX} XLSX
 * @param {string[]} headers
 * @param {Array<Array<string|number|null|undefined>>} bodyRows
 * @param {{ headerFill?: string, zebra?: boolean, freeze?: boolean }=} opts
 */
export function buildBrandedExcelSheet_(XLSX, headers, bodyRows, opts) {
  opts = opts || {};
  const cols = Array.isArray(headers) ? headers : [];
  const rows = Array.isArray(bodyRows) ? bodyRows : [];
  const colCount = Math.max(cols.length, 1);
  const headerRow = excelBrandPreambleRowCount_();

  // Build AOA without stuffing empty strings into every preamble column
  // (empty cells in a merge range make the title look broken / off-center).
  const wsData = [];
  wsData[0] = [EXCEL_REPORT_TITLE];
  wsData[1] = [];
  wsData[2] = ['COMPANY INFORMATION'];
  companyLines_().forEach(function(line, i) {
    wsData[3 + i] = [line];
  });
  wsData[8] = [];
  wsData[headerRow] = cols.slice();
  rows.forEach(function(row, i) {
    const out = [];
    for (let c = 0; c < colCount; c++) {
      const v = row && row[c] != null ? row[c] : '';
      out[c] = v;
    }
    wsData[headerRow + 1 + i] = out;
  });

  const ws = XLSX.utils.aoa_to_sheet(wsData, { skipHidden: false });

  // Drop phantom empty cells inside the title merge range (B1…Zn).
  for (let c = 1; c < colCount; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c: c });
    if (ws[addr] && (ws[addr].v === '' || ws[addr].v == null)) delete ws[addr];
  }

  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(colCount - 1, 0) } },
  ];

  // Column widths from data — keep A readable for company lines.
  const companyMax = companyLines_().reduce(function(m, line) {
    return Math.max(m, line.length);
  }, 'COMPANY INFORMATION'.length);
  ws['!cols'] = Array.from({ length: colCount }, function(_, ci) {
    let maxLen = cols[ci] != null ? String(cols[ci]).length : 10;
    rows.forEach(function(row) {
      const cell = row && row[ci] != null ? String(row[ci]) : '';
      if (cell.length > maxLen) maxLen = cell.length;
    });
    if (ci === 0) maxLen = Math.max(maxLen, Math.min(companyMax, 42));
    return { wch: Math.min(Math.max(maxLen + 2, 12), 48) };
  });

  const headerFill = opts.headerFill || '8B1A1A';
  const headerFont = 'FFFFFF';
  const borderThin = {
    top: { style: 'thin', color: { rgb: 'BFA8A8' } },
    bottom: { style: 'thin', color: { rgb: 'BFA8A8' } },
    left: { style: 'thin', color: { rgb: 'BFA8A8' } },
    right: { style: 'thin', color: { rgb: 'BFA8A8' } },
  };
  const borderTitle = {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } },
  };

  if (!ws['!rows']) ws['!rows'] = [];
  ws['!rows'][0] = { hpt: 26 };
  ws['!rows'][headerRow] = { hpt: 22 };

  // Title — only top-left of merge gets the value + style.
  const titleAddr = XLSX.utils.encode_cell({ r: 0, c: 0 });
  ws[titleAddr] = {
    t: 's',
    v: EXCEL_REPORT_TITLE,
    s: {
      font: { bold: true, sz: 14, name: 'Calibri', color: { rgb: '1A0A0A' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: borderTitle,
    },
  };

  // COMPANY INFORMATION
  const infoAddr = XLSX.utils.encode_cell({ r: 2, c: 0 });
  ws[infoAddr] = {
    t: 's',
    v: 'COMPANY INFORMATION',
    s: {
      font: { bold: true, sz: 11, name: 'Calibri', color: { rgb: '1A0A0A' } },
      alignment: { horizontal: 'left', vertical: 'center' },
    },
  };

  // Company detail lines (single cell each — matches template)
  for (let r = 3; r <= 7; r++) {
    const addr = XLSX.utils.encode_cell({ r: r, c: 0 });
    if (!ws[addr]) continue;
    ws[addr].s = {
      font: { sz: 11, name: 'Calibri', color: { rgb: '1A0A0A' } },
      alignment: { horizontal: 'left', vertical: 'center' },
    };
  }

  // Table header row
  cols.forEach(function(h, ci) {
    const cellAddr = XLSX.utils.encode_cell({ r: headerRow, c: ci });
    ws[cellAddr] = {
      t: 's',
      v: h,
      s: {
        font: { bold: true, color: { rgb: headerFont }, sz: 11, name: 'Calibri' },
        fill: { patternType: 'solid', fgColor: { rgb: headerFill } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: borderThin,
      },
    };
  });

  // Body rows
  const zebra = opts.zebra !== false;
  rows.forEach(function(rowVals, ri) {
    const sheetRow = headerRow + 1 + ri;
    for (let ci = 0; ci < colCount; ci++) {
      const cellAddr = XLSX.utils.encode_cell({ r: sheetRow, c: ci });
      const raw = rowVals && rowVals[ci] != null ? rowVals[ci] : '';
      const isEven = ri % 2 === 0;
      ws[cellAddr] = {
        t: typeof raw === 'number' ? 'n' : 's',
        v: typeof raw === 'number' ? raw : String(raw),
        s: {
          font: { sz: 10, name: 'Calibri', color: { rgb: '3D2020' } },
          fill: zebra
            ? { patternType: 'solid', fgColor: { rgb: isEven ? 'FFFFFF' : 'FBF7F7' } }
            : undefined,
          alignment: { vertical: 'center', wrapText: false },
          border: borderThin,
        },
      };
    }
  });

  // Ensure sheet range covers preamble + table.
  const lastRow = headerRow + Math.max(rows.length, 0);
  const lastCol = Math.max(colCount - 1, 0);
  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(lastRow, headerRow), c: lastCol },
  });

  if (cols.length) {
    ws['!autofilter'] = {
      ref: XLSX.utils.encode_range({
        s: { r: headerRow, c: 0 },
        e: { r: Math.max(headerRow, lastRow), c: lastCol },
      }),
    };
    if (opts.freeze !== false) {
      ws['!freeze'] = {
        xSplit: 0,
        ySplit: headerRow + 1,
        topLeftCell: XLSX.utils.encode_cell({ r: headerRow + 1, c: 0 }),
        activePane: 'bottomLeft',
        state: 'frozen',
      };
    }
  }

  return ws;
}
