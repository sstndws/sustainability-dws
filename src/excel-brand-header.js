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

function padRow_(row, colCount) {
  const out = Array.isArray(row) ? row.slice() : [];
  while (out.length < colCount) out.push('');
  return out;
}

export function excelBrandPreambleAoa_(colCount) {
  const n = Math.max(2, Number(colCount) || 2);
  const info = EXCEL_COMPANY_INFO;
  return [
    padRow_([EXCEL_REPORT_TITLE], n),
    padRow_([], n),
    padRow_(['COMPANY INFORMATION'], n),
    padRow_(['NAME:', info.name], n),
    padRow_(['COUNTRY:', info.country], n),
    padRow_(['ADDRESS:', info.address], n),
    padRow_(['LATITUDE:', info.latitude], n),
    padRow_(['LONGITUDE:', info.longitude], n),
    padRow_([], n),
  ];
}

/**
 * Build a SheetJS worksheet with company header + maroon table header + data.
 * @param {typeof window.XLSX} XLSX
 * @param {string[]} headers
 * @param {Array<Array<string|number|null|undefined>>} bodyRows
 * @param {{ headerFill?: string, zebra?: boolean, freeze?: boolean, sheetTitleStyle?: boolean }=} opts
 */
export function buildBrandedExcelSheet_(XLSX, headers, bodyRows, opts) {
  opts = opts || {};
  const cols = Array.isArray(headers) ? headers : [];
  const rows = Array.isArray(bodyRows) ? bodyRows : [];
  const colCount = Math.max(cols.length, 2);
  const preamble = excelBrandPreambleAoa_(colCount);
  const headerRow = preamble.length;
  const wsData = preamble.concat([cols]).concat(rows);
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  ws['!merges'] = (ws['!merges'] || []).concat([
    { s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } },
  ]);

  ws['!cols'] = Array.from({ length: colCount }, function(_, ci) {
    let maxLen = cols[ci] != null ? String(cols[ci]).length : 10;
    if (ci === 0) maxLen = Math.max(maxLen, 14);
    if (ci === 1) maxLen = Math.max(maxLen, 28);
    rows.forEach(function(row) {
      const cell = row && row[ci] != null ? String(row[ci]) : '';
      if (cell.length > maxLen) maxLen = cell.length;
    });
    // Keep address readable in the brand block.
    if (ci === 1) maxLen = Math.max(maxLen, String(EXCEL_COMPANY_INFO.address).length);
    return { wch: Math.min(Math.max(maxLen + 2, 12), 56) };
  });

  const headerFill = opts.headerFill || '8B1A1A';
  const headerFont = 'FFFFFF';
  const borderThin = {
    top: { style: 'thin', color: { rgb: 'D4C4C4' } },
    bottom: { style: 'thin', color: { rgb: 'D4C4C4' } },
    left: { style: 'thin', color: { rgb: 'D4C4C4' } },
    right: { style: 'thin', color: { rgb: 'D4C4C4' } },
  };
  const borderTitle = {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } },
  };

  if (!ws['!rows']) ws['!rows'] = [];
  ws['!rows'][0] = { hpt: 28 };
  ws['!rows'][headerRow] = { hpt: 24 };

  // Title
  const titleAddr = XLSX.utils.encode_cell({ r: 0, c: 0 });
  if (!ws[titleAddr]) ws[titleAddr] = { t: 's', v: EXCEL_REPORT_TITLE };
  ws[titleAddr].s = {
    font: { bold: true, sz: 14, name: 'Calibri', color: { rgb: '1A0A0A' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: borderTitle,
  };

  // COMPANY INFORMATION label
  const infoAddr = XLSX.utils.encode_cell({ r: 2, c: 0 });
  if (!ws[infoAddr]) ws[infoAddr] = { t: 's', v: 'COMPANY INFORMATION' };
  ws[infoAddr].s = {
    font: { bold: true, sz: 11, name: 'Calibri', color: { rgb: '1A0A0A' } },
    alignment: { horizontal: 'left', vertical: 'center' },
  };

  // Company field labels + values
  for (let r = 3; r <= 7; r++) {
    const labelAddr = XLSX.utils.encode_cell({ r: r, c: 0 });
    const valueAddr = XLSX.utils.encode_cell({ r: r, c: 1 });
    if (ws[labelAddr]) {
      ws[labelAddr].s = {
        font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: '3D2020' } },
        alignment: { horizontal: 'left', vertical: 'center' },
      };
    }
    if (ws[valueAddr]) {
      ws[valueAddr].s = {
        font: { sz: 10, name: 'Calibri', color: { rgb: '3D2020' } },
        alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
      };
    }
  }

  // Table header row
  cols.forEach(function(h, ci) {
    const cellAddr = XLSX.utils.encode_cell({ r: headerRow, c: ci });
    if (!ws[cellAddr]) ws[cellAddr] = { t: 's', v: h };
    ws[cellAddr].s = {
      font: { bold: true, color: { rgb: headerFont }, sz: 11, name: 'Calibri' },
      fill: { fgColor: { rgb: headerFill } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: borderThin,
    };
  });

  // Body rows
  const zebra = opts.zebra !== false;
  rows.forEach(function(rowVals, ri) {
    const sheetRow = headerRow + 1 + ri;
    for (let ci = 0; ci < colCount; ci++) {
      const cellAddr = XLSX.utils.encode_cell({ r: sheetRow, c: ci });
      if (!ws[cellAddr]) {
        const v = rowVals && rowVals[ci] != null ? rowVals[ci] : '';
        ws[cellAddr] = { t: 's', v: String(v) };
      }
      const isEven = ri % 2 === 0;
      ws[cellAddr].s = {
        font: { sz: 10, name: 'Calibri', color: { rgb: '3D2020' } },
        fill: zebra ? { fgColor: { rgb: isEven ? 'FFFFFF' : 'FBF7F7' } } : undefined,
        alignment: { vertical: 'center', wrapText: true },
        border: borderThin,
      };
    }
  });

  if (cols.length) {
    const lastDataRow = headerRow + rows.length;
    ws['!autofilter'] = {
      ref: XLSX.utils.encode_range({
        s: { r: headerRow, c: 0 },
        e: { r: Math.max(headerRow, lastDataRow), c: cols.length - 1 },
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
