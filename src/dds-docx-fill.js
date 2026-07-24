/**
 * Shared EUDR DDS template fill — used by DOCX download and PDF render pipeline.
 */
import PizZip from 'pizzip';
import { buildDdsExportModel_ } from './dds-export-model.js';

export const DDS_TEMPLATE_URL = '/templates/EUDR_DDS_Revised.docx';

function xmlEsc_(s) {
  return String(s != null ? s : '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function replaceOnce_(xml, search, replacement) {
  const idx = xml.indexOf(search);
  if (idx === -1) return xml;
  return xml.slice(0, idx) + replacement + xml.slice(idx + search.length);
}

function replaceAll_(xml, search, replacement) {
  return xml.split(search).join(replacement);
}

function findTableEnd_(xml, tblStart) {
  let depth = 0;
  let i = tblStart;
  while (i < xml.length) {
    if (xml.startsWith('<w:tbl>', i)) {
      depth += 1;
      i += 7;
      continue;
    }
    if (xml.startsWith('</w:tbl>', i)) {
      depth -= 1;
      i += 8;
      if (depth === 0) return i;
      continue;
    }
    i += 1;
  }
  return -1;
}

function removeTableColumnIndices_(tableXml, removeCols) {
  const sorted = removeCols.slice().sort(function(a, b) { return b - a; });
  const rows = tableRows_(tableXml).map(function(rowXml) {
    const cells = tableCells_(rowXml);
    sorted.forEach(function(ci) {
      if (ci >= 0 && ci < cells.length) cells.splice(ci, 1);
    });
    return rebuildRow_(rowXml, cells);
  });

  const gridMatch = tableXml.match(/<w:tblGrid[\s\S]*?<\/w:tblGrid>/);
  if (gridMatch) {
    let gridCols = gridMatch[0].match(/<w:gridCol[^/]*\/>/g) || [];
    sorted.forEach(function(ci) {
      if (ci >= 0 && ci < gridCols.length) gridCols.splice(ci, 1);
    });
    tableXml = tableXml.replace(/<w:tblGrid[\s\S]*?<\/w:tblGrid>/, '<w:tblGrid>' + gridCols.join('') + '</w:tblGrid>');
  }

  const tblPr = tableXml.match(/<w:tblPr[\s\S]*?<\/w:tblPr>/)?.[0] || '';
  const tblGrid = tableXml.match(/<w:tblGrid[\s\S]*?<\/w:tblGrid>/)?.[0] || '';
  return '<w:tbl>' + tblPr + tblGrid + rows.join('') + '</w:tbl>';
}

function patchTableByMarker_(xml, marker, patchFn) {
  const slice = sliceTable_(xml, marker);
  if (!slice) return xml;
  const patched = patchFn(slice.xml);
  return xml.slice(0, slice.start) + patched + xml.slice(slice.end);
}

/** Drop English subtitles in section headers; keep numbering (8., 9., 10., …). */
function stripDdsSectionEnglishSubtitles_(xml) {
  const pairs = [
    ['Ringkasan Penilaian Risiko &amp; Mitigasi (Risk Assessment Summary)', 'Ringkasan Penilaian Risiko &amp; Mitigasi'],
    ['Ringkasan Penilaian Risiko & Mitigasi (Risk Assessment Summary)', 'Ringkasan Penilaian Risiko & Mitigasi'],
    ['Pernyataan Operator (Operator Declaration)', 'Pernyataan Operator'],
    ['Catatan Penyimpanan Dokumen (Record Keeping)', 'Catatan Penyimpanan Dokumen'],
  ];
  pairs.forEach(function(pair) {
    xml = replaceAll_(xml, pair[0], pair[1]);
  });
  return xml;
}

/** User revisions: label text, supplier/geo block order, supplier columns, geo column title. */
function applyDdsTemplateRevisions_(xml) {
  const textPairs = [
    ['Nomor Izin Usaha (NIB / SIUP)', 'Nomor Izin Usaha (NIB)'],
    ['Nama Perusahaan Pembeli/Importir UE', 'Nama Perusahaan Pembeli'],
    ['Alamat Pembeli/Importir UE', 'Alamat Pembeli'],
    ['Nomor EORI Pembeli/Importir UE (opsional, bila diketahui)', 'Nomor Pembeli (opsional, bila diketahui)'],
    ['Kode HS / Kode CN', 'Kode HS'],
    ['Detail Pengiriman / Konsinyasi Ekspor', 'Detail Pengiriman'],
    ['Nomor Invoice Komersial / Kontrak', 'Nomor Invoice / Kontrak'],
    ['Informasi Pemasok dan Ketertelusuran Rantai Pasok (Annex II, Butir 5) — WAJIB', 'Informasi Pemasok dan Ketertelusuran Rantai Pasok'],
    ['Informasi Pemasok dan Ketertelusuran Rantai Pasok (Annex II, Butir 5)', 'Informasi Pemasok dan Ketertelusuran Rantai Pasok'],
  ];
  textPairs.forEach(function(pair) {
    xml = replaceAll_(xml, pair[0], pair[1]);
  });

  xml = patchTableByMarker_(xml, 'Nama Pemasok (PKS/Kebun)', function(tableXml) {
    return removeTableColumnIndices_(tableXml, [3, 4]);
  });

  const geoIdx = xml.indexOf('Tabel Geolokasi Plot');
  const supIdx = xml.indexOf('Informasi Pemasok');
  if (geoIdx !== -1 && supIdx !== -1 && geoIdx < supIdx) {
    const riskIdx = xml.indexOf('Kategori Risiko Negara');
    if (riskIdx !== -1) {
      const metaTblStart = xml.lastIndexOf('<w:tbl>', riskIdx);
      const metaTblEnd = findTableEnd_(xml, metaTblStart);
      const docIdx = xml.indexOf('Daftar Dokumen');
      const sec7Start = docIdx !== -1 ? xml.lastIndexOf('<w:p>', docIdx) : -1;
      if (metaTblEnd > 0 && sec7Start > supIdx) {
        const supBlockStart = xml.lastIndexOf('<w:p>', supIdx);
        const geoBlockStart = xml.lastIndexOf('<w:p>', geoIdx);
        const aStart = geoBlockStart > metaTblEnd ? geoBlockStart : metaTblEnd;
        const aEnd = supBlockStart;
        const bStart = supBlockStart;
        const bEnd = sec7Start;
        if (aStart > 0 && aEnd > aStart && bEnd > bStart) {
          const partGeo = xml.slice(aStart, aEnd);
          const partSup = xml.slice(bStart, bEnd);
          xml = xml.slice(0, aStart) + partSup + partGeo + xml.slice(bEnd);
        }
      }
    }
  }

  xml = stripDdsSectionEnglishSubtitles_(xml);

  return xml;
}

function setCellText_(cellXml, text) {
  const val = xmlEsc_(text);
  let cell = blackenRunsInCell_(cellXml);
  if (/<w:t[\s\S]*?<\/w:t>/.test(cell)) {
    return cell.replace(
      /<w:t(\s+xml:space="preserve")?>[\s\S]*?<\/w:t>/,
      '<w:t xml:space="preserve">' + val + '</w:t>'
    );
  }
  return cell.replace(
    /(<w:p>[\s\S]*?<w:pPr>[\s\S]*?<\/w:pPr>)/,
    '$1<w:r><w:rPr><w:color w:val="000000"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">' + val + '</w:t></w:r>'
  );
}

/** User-fill fields: placeholder & value cells use black text (was gray italic). */
function blackenRunProperties_(inner) {
  let pr = inner || '';
  pr = pr.replace(/<w:i\/>/g, '').replace(/<w:iCs\/>/g, '');
  if (/<w:color\s/.test(pr)) {
    pr = pr.replace(/<w:color[^/]*\/>/g, '<w:color w:val="000000"/>');
  } else {
    pr = '<w:color w:val="000000"/>' + pr;
  }
  return pr;
}

function blackenRunsInCell_(cellXml) {
  return cellXml.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/g, function(_m, inner) {
    if (inner.indexOf('999999') === -1) return _m;
    return '<w:rPr>' + blackenRunProperties_(inner) + '</w:rPr>';
  });
}

function applyDdsFillableFieldTextStyle_(xml) {
  return xml.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/g, function(m, inner) {
    if (inner.indexOf('999999') === -1) return m;
    return '<w:rPr>' + blackenRunProperties_(inner) + '</w:rPr>';
  });
}

function tableRows_(tableXml) {
  return tableXml.match(/<w:tr[\s\S]*?<\/w:tr>/g) || [];
}

function tableCells_(rowXml) {
  return rowXml.match(/<w:tc[\s\S]*?<\/w:tc>/g) || [];
}

function rebuildRow_(rowXml, cells) {
  let i = 0;
  return rowXml.replace(/<w:tc[\s\S]*?<\/w:tc>/g, function() {
    return cells[i++];
  });
}

function sliceTable_(xml, marker) {
  const markerIdx = xml.indexOf(marker);
  if (markerIdx === -1) return null;
  const start = xml.lastIndexOf('<w:tbl>', markerIdx);
  const end = xml.indexOf('</w:tbl>', markerIdx);
  if (start === -1 || end === -1) return null;
  return { start: start, end: end + 8, xml: xml.slice(start, end + 8) };
}

function fillDataTable_(xml, marker, dataRows, startRow, colMap) {
  const slice = sliceTable_(xml, marker);
  if (!slice) return xml;

  let tableXml = slice.xml;
  let rows = tableRows_(tableXml);
  const templateRow = rows[Math.min(startRow, rows.length - 1)] || rows[rows.length - 1];

  while (rows.length < startRow + dataRows.length) {
    rows.push(templateRow);
  }

  dataRows.forEach(function(rowData, ri) {
    const rowIdx = startRow + ri;
    const cells = tableCells_(rows[rowIdx]);
    colMap.forEach(function(map) {
      const colIdx = map.col;
      const value = typeof map.value === 'function' ? map.value(rowData, ri) : rowData[map.key];
      if (colIdx < cells.length) cells[colIdx] = setCellText_(cells[colIdx], value != null ? value : '');
    });
    rows[rowIdx] = rebuildRow_(rows[rowIdx], cells);
  });

  const tblPr = tableXml.match(/<w:tblPr[\s\S]*?<\/w:tblPr>/)?.[0] || '';
  const tblGrid = tableXml.match(/<w:tblGrid[\s\S]*?<\/w:tblGrid>/)?.[0] || '';
  tableXml = '<w:tbl>' + tblPr + tblGrid + rows.join('') + '</w:tbl>';

  return xml.slice(0, slice.start) + tableXml + xml.slice(slice.end);
}

function buildPlaceholderMap_(model) {
  const s = model.sections;
  return [
    ['[NAMA PERUSAHAAN / LOGO]', model.company],
    ['[... KG net mass/tahun — persyaratan amandemen Desember 2025]', s.s1[2][1]],
    ['[PT. ...]', s.s2[0][1]],
    ['[Alamat lengkap, Kota, Indonesia]', s.s2[1][1]],
    ['[Nomor NIB]', s.s2[2][1]],
    ['[Nama buyer UE penerima dokumen ini]', s.s2[4][1]],
    ['[Alamat lengkap, Kota, Negara UE]', s.s2[5][1]],
    ['[... / Tidak diketahui]', s.s2[6][1]],
    ['[1511.10 / 1511.90 / 1513.21 / 1513.29]', s.s3[0][1]],
    ['[CPO / RBDPO / RBD Palm Olein / RBD Palm Stearin / CPKO / RBDPKO]', s.s3[1][1]],
    ['[Nama produk persis sesuai sales contract/invoice — cth: &quot;RBDPO&quot;, &quot;RBD Palm Olein IV58&quot;, &quot;CPO SumSel Grade&quot;]', s.s3[2][1]],
    ['[... KG net mass — sesuai B/L atau kontrak]', s.s3[4][1]],
    ['[Nomor Invoice]', s.s4[0][1]],
    ['[Nomor B/L]', s.s4[1][1]],
    ['[Nama Kapal / Nomor Tangki]', s.s4[2][1]],
    ['[Nama Pelabuhan, Indonesia]', s.s4[3][1]],
    ['[Nama Pelabuhan, Negara UE]', s.s4[4][1]],
    ['[DD/MM/YYYY — tanggal kapal berangkat / tanggal B/L. Gunakan ETD, bukan ETA. ETA dikelola oleh buyer UE di sistem TRACES mereka.]', s.s4[5][1]],
    ['[Standard / Low / High — sesuai penetapan Komisi Eropa terbaru]', s.s5Meta[1][1]],
    ['[No. Dok. RA-EUDR-YYYY-XX]', s.s8[0][1]],
    ['[Sesuai Pasal 10(2) EUDR — lihat Risk Assessment Report]', s.s8[2][1]],
    ['[Negligible / Low / Standard / High]', s.s8[3][1]],
    ['[Risiko berhasil ditekan ke tingkat Negligible / Masih memerlukan tindak lanjut]', s.s8[4][1]],
    ['[Ringkasan singkat — detail di Risk Assessment Report]', s.s8[5][1]],
    ['[Nama Lengkap]', s.s9[0][1]],
    ['[Jabatan / Divisi]', s.s9[1][1]],
    ['[Kota, DD/MM/YYYY]', s.s9[2][1]],
  ];
}

function fillDocChecklist_(xml, docRows) {
  const slice = sliceTable_(xml, 'Jenis Dokumen');
  if (!slice) return xml;

  let tableXml = slice.xml;
  const rows = tableRows_(tableXml);
  const docMap = {};
  docRows.forEach(function(row) {
    if (row.type === 'doc') docMap[row.code] = row;
  });

  const filled = rows.map(function(rowXml) {
    const texts = (rowXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [])
      .map(function(t) { return t.replace(/<[^>]+>/g, ''); });
    const code = texts[0] || '';
    const doc = docMap[code];
    if (!doc) return rowXml;

    const cells = tableCells_(rowXml);
    if (cells.length >= 6) {
      cells[2] = setCellText_(cells[2], doc.number);
      cells[3] = setCellText_(cells[3], doc.date);
      cells[4] = setCellText_(cells[4], doc.available);
      cells[5] = setCellText_(cells[5], doc.notes);
    }
    return rebuildRow_(rowXml, cells);
  });

  tableXml = tableXml.replace(/<w:tr[\s\S]*?<\/w:tr>/g, function() {
    return filled.shift();
  });

  return xml.slice(0, slice.start) + tableXml + xml.slice(slice.end);
}

export function fillTemplateXml_(xml, model) {
  xml = applyDdsTemplateRevisions_(xml);

  buildPlaceholderMap_(model).forEach(function(pair) {
    xml = replaceAll_(xml, pair[0], xmlEsc_(pair[1]));
  });

  xml = replaceOnce_(xml, '[DD/MM/YYYY]', xmlEsc_(model.sections.s1[0][1]));
  xml = replaceOnce_(xml, '[DD/MM/YYYY]', xmlEsc_(model.sections.s8[1][1]));

  xml = fillDataTable_(xml, 'Identifikasi Plot / Kebun / PKS', model.geoRows, 1, [
    { col: 0, key: 0 },
    { col: 1, key: 1 },
    { col: 2, key: 2 },
    { col: 3, key: 3 },
    { col: 4, key: 4 },
    { col: 5, key: 5 },
  ]);

  xml = fillDataTable_(xml, 'Nama Pemasok (PKS/Kebun)', model.supplierRows, 1, [
    { col: 0, key: 1 },
    { col: 1, key: 2 },
    { col: 2, key: 3 },
  ]);

  xml = replaceAll_(xml, 'Identifikasi Plot / Kebun / PKS', 'Nama Pemasok');

  xml = applyDdsFillableFieldTextStyle_(xml);

  return fillDocChecklist_(xml, model.docRows);
}

/**
 * @param {{ master: object, suppliers: object[], geolocation: object[], documents: object[] }} bundle
 * @returns {Promise<{ blob: Blob, fileKey: string, model: ReturnType<typeof buildDdsExportModel_> }>}
 */
export async function buildFilledDdsDocxBlob_(bundle) {
  const model = buildDdsExportModel_(bundle);
  const res = await fetch(DDS_TEMPLATE_URL);
  if (!res.ok) throw new Error('Template DOCX tidak ditemukan (' + DDS_TEMPLATE_URL + ')');

  const buf = await res.arrayBuffer();
  const zip = new PizZip(buf);
  const docPath = 'word/document.xml';
  let xml = zip.file(docPath).asText();
  xml = fillTemplateXml_(xml, model);
  zip.file(docPath, xml);

  const blob = zip.generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  return { blob: blob, fileKey: model.fileKey, exportBaseName: model.exportBaseName, model: model };
}
