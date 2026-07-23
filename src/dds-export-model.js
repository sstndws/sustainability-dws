/** Shared EUDR DDS export data — PDF & DOCX read the same sheet bundle. */

export const DDS_EXPORT_DOC_ROWS = [
  { group: 'A. Dokumen Bebas Deforestasi' },
  { code: 'A1', label: 'Citra Satelit / Analisis NDVI per Plot Kebun' },
  { code: 'A2', label: 'Laporan Pemantauan Deforestasi (Forest Cover Analysis)' },
  { code: 'A3', label: 'Sertifikasi ISPO (Indonesia Sustainable Palm Oil)' },
  { code: 'A4', label: 'Sertifikasi RSPO (opsional)' },
  { code: 'A5', label: 'STDB (Surat Tanda Daftar Budidaya) — untuk kebun plasma/swadaya' },
  { group: 'B. Dokumen Kepatuhan Hukum' },
  { code: 'B1', label: 'HGU (Hak Guna Usaha) / SHM (Sertifikat Hak Milik)' },
  { code: 'B2', label: 'AMDAL / UKL-UPL (Izin Lingkungan)' },
  { code: 'B3', label: 'Izin Usaha Perkebunan (IUP)' },
  { code: 'B4', label: 'Izin Usaha Industri Pengolahan (IUIP) / NIB Refinery' },
  { code: 'B5', label: 'Dokumen FPIC (Free Prior Informed Consent) — jika ada lahan adat' },
  { code: 'B6', label: 'Dokumen Ketenagakerjaan (sesuai UU yang berlaku)' },
];

export const DDS_EXPORT_STATIC = {
  introInfo: 'Dokumen ini adalah paket data due diligence untuk perusahaan Indonesia yang mengekspor CPO/minyak kelapa sawit ke pembeli/importir di Uni Eropa, tanpa memiliki entitas atau Authorised Representative terdaftar di UE. Pengajuan DDS ke sistem resmi UE (TRACES NT) dilakukan oleh pembeli/importir UE Anda menggunakan data yang Anda sediakan melalui dokumen ini.',
  operatorType: 'Operator — Eksportir (Pemasok Data ke Importir/Operator UE)',
  tracesRef: 'Tidak berlaku — diisi oleh pembeli/importir UE setelah submit ke TRACES',
  ueStatus: 'Perusahaan TIDAK memiliki entitas atau Authorised Representative terdaftar di UE. Dokumen ini disediakan sebagai paket data due diligence agar pembeli/importir UE dapat men-submit DDS ke sistem TRACES atas nama rantai pasoknya.',
  hsInfo: 'Pilih HS Code sesuai produk yang dikirimkan: CPO = 1511.10 | RBDPO / Palm Olein / Palm Stearin = 1511.90 | CPKO = 1513.21 | RBDPKO = 1513.29',
  commodity: 'Kelapa sawit (oil palm)',
  shipmentInfo: 'Bagian ini menautkan DDS dengan dokumen pabean dan pengapalan aktual, sesuai praktik ekspor CPO curah via tanker.',
  geoInfo: 'Seluruh plot sumber wajib dicantumkan, termasuk kebun plasma dan swadaya. Untuk jumlah plot banyak, lampirkan file GeoJSON terpisah dan cantumkan referensi file di kolom keterangan.',
  geoTableTitle: 'Tabel Geolokasi Plot/Kebun Sumber TBS',
  supplierInfo: 'Bagian ini WAJIB diisi. Buyer UE membutuhkan informasi ini untuk membuktikan ketertelusuran hingga ke kebun. Cantumkan minimal nama & alamat PKS pemasok CPO/TBS ke refinery Anda.',
  docInfo: 'Bagian ini berisi daftar referensi dokumen. Dokumen asli diserahkan sebagai lampiran terpisah kepada buyer. Centang (✓) kolom "Tersedia" untuk setiap dokumen yang dapat disediakan.',
  riskInfo: 'Bagian ini hanya memuat KESIMPULAN dari Risk Assessment. Uraian lengkap metodologi, identifikasi risiko, dan tindakan mitigasi disediakan dalam dokumen terpisah: "EUDR Risk Assessment & Mitigation Report" yang menyertai paket data ini.',
  declaration: 'Dengan pernyataan uji tuntas (due diligence statement) ini, perusahaan yang bertanda tangan di bawah menegaskan bahwa uji tuntas (due diligence) telah dilaksanakan sesuai Pasal 8 Regulation (EU) 2023/1115, dan bahwa tidak ditemukan risiko atau hanya ditemukan risiko yang dapat diabaikan (negligible) bahwa produk relevan tidak memenuhi ketentuan Pasal 3 regulasi tersebut — yaitu produk bebas deforestasi, diproduksi sesuai hukum negara produksi, dan telah dicakup oleh pernyataan uji tuntas ini.',
  declarationInfo: 'Keakuratan seluruh data dalam dokumen ini menjadi tanggung jawab perusahaan sebagai penyedia data, meskipun pengajuan formal DDS ke sistem TRACES dilakukan oleh pembeli/importir UE.',
  recordKeeping: 'Seluruh dokumen pendukung due diligence wajib disimpan minimal 5 (lima) tahun sejak tanggal pengajuan DDS ke sistem UE, sesuai Pasal 12 EUDR. Nomor referensi DDS yang diterbitkan sistem (oleh pembeli/importir UE) wajib dicantumkan pada dokumen pabean, invoice, dan B/L terkait konsinyasi ini.',
  disclaimer: 'Disclaimer: Template ini disusun berdasarkan Annex II Regulation (EU) 2023/1115 (EUDR) sebagaimana diubah oleh Regulation (EU) 2025/2650 dan panduan resmi Komisi Eropa. Bukan pengganti nasihat hukum profesional. Persyaratan teknis dapat berubah — selalu rujuk panduan resmi EU Information System/TRACES NT terbaru.',
  regulationLine: 'Sesuai Pasal 4(2) dan Annex II, Regulation (EU) 2023/1115 (EUDR) sebagaimana diubah oleh Regulation (EU) 2025/2650',
};

export function ddsPick_(row, field) {
  if (!row) return '';
  if (row[field] != null && String(row[field]).trim() !== '') return String(row[field]).trim();
  const want = String(field || '').trim().toUpperCase();
  for (const k of Object.keys(row)) {
    if (k.charAt(0) === '_') continue;
    if (String(k).trim().toUpperCase() === want) return String(row[k]).trim();
  }
  return '';
}

export function ddsSheetOr_(row, field, placeholder) {
  const v = ddsPick_(row, field);
  return v || placeholder;
}

export function ddsSheetCell_(val) {
  return String(val != null ? val : '').trim();
}

function qtyKg_(master) {
  const v = ddsPick_(master, 'QUANTITY (KG)');
  return v ? v + ' KG net mass — sesuai B/L atau kontrak' : '[... KG net mass — sesuai B/L atau kontrak]';
}

function annualQty_(master) {
  const v = ddsPick_(master, 'ANNUAL QUANTITY ESTIMATE (KG)');
  return v ? v + ' KG net mass/tahun — persyaratan amandemen Desember 2025' : '[... KG net mass/tahun — persyaratan amandemen Desember 2025]';
}

function etd_(master) {
  const v = ddsPick_(master, 'ETD DATE');
  return v
    ? v + ' — tanggal kapal berangkat / tanggal B/L. Gunakan ETD, bukan ETA. ETA dikelola oleh buyer UE di sistem TRACES mereka.'
    : '[DD/MM/YYYY — tanggal kapal berangkat / tanggal B/L. Gunakan ETD, bukan ETA. ETA dikelola oleh buyer UE di sistem TRACES mereka.]';
}

function vessel_(master) {
  const parts = [
    ddsPick_(master, 'SHIPPING NAME'),
    ddsPick_(master, 'TANK NUMBER'),
    ddsPick_(master, 'BATCH NUMBER'),
  ].filter(Boolean);
  return parts.length ? parts.join(' / ') : '[Nama Kapal / Nomor Tangki]';
}

function signPlaceDate_(master) {
  const place = ddsPick_(master, 'SIGNATORY PLACE');
  const date = ddsPick_(master, 'SIGNATORY DATE');
  if (place && date) return place + ', ' + date;
  if (place || date) return place || date;
  return '[Kota, DD/MM/YYYY]';
}

function padRows_(rows, minCount, emptyFn) {
  const out = rows.slice();
  while (out.length < minCount) out.push(emptyFn(out.length + 1));
  return out;
}

function mergeDocMap_(docs) {
  const docMap = {};
  (docs || []).forEach(function(d) {
    docMap[String(d['DOC CODE'] || '').trim()] = d;
  });
  return docMap;
}

function buildDocRows_(docMap) {
  const rows = [];
  DDS_EXPORT_DOC_ROWS.forEach(function(def) {
    if (def.group) {
      rows.push({ type: 'group', label: def.group });
      return;
    }
    const row = docMap[def.code] || {};
    const avail = String(row.AVAILABLE || '').trim().toUpperCase();
    rows.push({
      type: 'doc',
      code: def.code,
      label: def.label,
      number: ddsSheetCell_(row['DOC NUMBER']),
      date: ddsSheetCell_(row['DOC DATE']),
      available: avail === 'Y' ? '\u2713' : (avail === 'N' ? '\u2717' : ''),
      notes: ddsSheetCell_(row.NOTES),
    });
  });
  return rows;
}

/** Normalised export payload for PDF/DOCX renderers. */
export function buildDdsExportModel_(bundle) {
  const master = bundle.master || {};
  const suppliers = (bundle.suppliers || []).slice().sort(function(a, b) {
    return Number(a['LINE NO'] || 0) - Number(b['LINE NO'] || 0);
  });
  const geo = (bundle.geolocation || []).slice().sort(function(a, b) {
    return Number(a['LINE NO'] || 0) - Number(b['LINE NO'] || 0);
  });
  const docMap = mergeDocMap_(bundle.documents);

  const supSource = suppliers.length
    ? suppliers
    : (ddsPick_(master, 'SUPPLIER NAME')
      ? [{
        'LINE NO': '1',
        'SUPPLIER NAME': ddsPick_(master, 'SUPPLIER NAME'),
        'SUPPLIER ADDRESS': '',
        'SUPPLIER TYPE': '',
        'SUPPLIER DDS REF': '',
        'PROOF DOC': '',
      }]
      : []);

  const supEmpty = !supSource.length;

  const plantLabel = ddsPick_(master, 'PLANT')
    || ddsPick_(master, 'EXPORTER COMPANY NAME')
    || ddsPick_(master, 'SD NUMBER')
    || 'DDS';
  const exportBaseName = 'EUDR DDS (' + String(plantLabel).replace(/[\\/:*?"<>|]/g, '').trim() + ')';

  return {
    fileKey: ddsPick_(master, 'SD NUMBER').replace(/[^\w\-]+/g, '_') || 'DDS',
    exportBaseName: exportBaseName,
    company: ddsSheetOr_(master, 'EXPORTER COMPANY NAME', '')
      || ddsSheetOr_(master, 'PLANT', '')
      || '[NAMA PERUSAHAAN / LOGO]',
    sections: {
      s1: [
        ['Tanggal Penyusunan Dokumen', ddsSheetOr_(master, 'DOCUMENT DATE', '[DD/MM/YYYY]')],
        ['Jenis Operator', DDS_EXPORT_STATIC.operatorType],
        ['Estimasi Kuantitas Tahunan Produk yang Dicakup', annualQty_(master)],
        ['Nomor Referensi DDS dari Sistem TRACES', DDS_EXPORT_STATIC.tracesRef],
      ],
      s2: [
        ['Nama Perusahaan Eksportir', ddsSheetOr_(master, 'EXPORTER COMPANY NAME', ddsSheetOr_(master, 'PLANT', '[PT. ...]'))],
        ['Alamat Perusahaan', ddsSheetOr_(master, 'EXPORTER ADDRESS', '[Alamat lengkap, Kota, Indonesia]')],
        ['Nomor Izin Usaha (NIB / SIUP)', ddsSheetOr_(master, 'EXPORTER NIB', '[Nomor NIB]')],
        ['Status terhadap sistem UE', DDS_EXPORT_STATIC.ueStatus],
        ['Nama Perusahaan Pembeli/Importir UE', ddsSheetOr_(master, 'BUYER NAME', '[Nama buyer UE penerima dokumen ini]')],
        ['Alamat Pembeli/Importir UE', ddsSheetOr_(master, 'BUYER ADDRESS', '[Alamat lengkap, Kota, Negara UE]')],
        ['Nomor EORI Pembeli/Importir UE (opsional, bila diketahui)', ddsSheetOr_(master, 'BUYER EORI', '[... / Tidak diketahui]')],
      ],
      s3: [
        ['Kode HS / Kode CN', ddsSheetOr_(master, 'HS CODE', '[1511.10 / 1511.90 / 1513.21 / 1513.29]')],
        ['Deskripsi Produk', ddsSheetOr_(master, 'PRODUCT', '[CPO / RBDPO / RBD Palm Olein / RBD Palm Stearin / CPKO / RBDPKO]')],
        ['Nama Dagang (Trade Name)', ddsSheetOr_(master, 'TRADE NAME', '[Nama produk persis sesuai sales contract/invoice — cth: "RBDPO", "RBD Palm Olein IV58", "CPO SumSel Grade"]')],
        ['Komoditas Relevan EUDR', DDS_EXPORT_STATIC.commodity],
        ['Kuantitas (net mass, KG)', qtyKg_(master)],
      ],
      s4: [
        ['Nomor Invoice Komersial / Kontrak', ddsSheetOr_(master, 'CONTRACT NUMBER', '[Nomor Invoice]')],
        ['Nomor Bill of Lading (B/L)', ddsSheetOr_(master, 'BL NUMBER', '[Nomor B/L]')],
        ['Nama Kapal / Nomor Tangki / Nomor Batch', vessel_(master)],
        ['Pelabuhan Muat (Port of Loading)', ddsSheetOr_(master, 'PORT OF LOADING', '[Nama Pelabuhan, Indonesia]')],
        ['Pelabuhan Bongkar di UE (Port of Discharge)', ddsSheetOr_(master, 'PORT OF DISCHARGE', '[Nama Pelabuhan, Negara UE]')],
        ['Tanggal Pengapalan — ETD (Estimated Time of Departure)', etd_(master)],
      ],
      s5Meta: (function() {
        const pairs = [
          ['Negara Produksi', 'Indonesia'],
          ['Kategori Risiko Negara (per Komisi Eropa, Pasal 29)', ddsSheetOr_(master, 'COUNTRY RISK CATEGORY', '[Standard / Low / High — sesuai penetapan Komisi Eropa terbaru]')],
        ];
        if (ddsPick_(master, 'GEOJSON FILE REF')) {
          pairs.push(['Referensi file GeoJSON (lampiran)', ddsPick_(master, 'GEOJSON FILE REF')]);
        }
        return pairs;
      })(),
      s8: [
        ['Nomor / Judul Dokumen Risk Assessment', ddsSheetOr_(master, 'RA DOCUMENT NO', '[No. Dok. RA-EUDR-YYYY-XX]')],
        ['Tanggal Penilaian Risiko Dilakukan', ddsSheetOr_(master, 'RA DATE', '[DD/MM/YYYY]')],
        ['Metodologi yang Digunakan', ddsSheetOr_(master, 'RA METHODOLOGY', '[Sesuai Pasal 10(2) EUDR — lihat Risk Assessment Report]')],
        ['Kesimpulan Tingkat Risiko Keseluruhan', ddsSheetOr_(master, 'RA OVERALL RISK', '[Negligible / Low / Standard / High]')],
        ['Status Setelah Mitigasi', ddsSheetOr_(master, 'RA POST MITIGATION STATUS', '[Risiko berhasil ditekan ke tingkat Negligible / Masih memerlukan tindak lanjut]')],
        ['Tindakan Mitigasi Utama yang Diterapkan', ddsSheetOr_(master, 'RA MITIGATION SUMMARY', '[Ringkasan singkat — detail di Risk Assessment Report]')],
      ],
      s9: [
        ['Nama Penanggung Jawab', ddsSheetOr_(master, 'SIGNATORY NAME', '[Nama Lengkap]')],
        ['Jabatan', ddsSheetOr_(master, 'SIGNATORY TITLE', '[Jabatan / Divisi]')],
        ['Tempat dan Tanggal', signPlaceDate_(master)],
        ['Tanda Tangan & Cap Perusahaan', ''],
      ],
    },
    geoRows: padRows_(
      geo.map(function(r, i) {
        return [
          String(r['LINE NO'] || i + 1),
          ddsSheetCell_(r['PLOT ID']),
          ddsSheetCell_(r['AREA HA']),
          ddsSheetCell_(r.COORDINATES),
          ddsSheetCell_(r['HARVEST DATE']),
          ddsSheetCell_(r.NOTES),
        ];
      }),
      3,
      function(n) { return [String(n), '', '', '', '', '']; }
    ),
    supplierRows: padRows_(
      supSource.map(function(r, i) {
        return [
          String(r['LINE NO'] || i + 1),
          ddsSheetCell_(r['SUPPLIER NAME']),
          ddsSheetCell_(r['SUPPLIER ADDRESS']),
          ddsSheetCell_(r['SUPPLIER TYPE']),
          ddsSheetCell_(r['SUPPLIER DDS REF']),
          ddsSheetCell_(r['PROOF DOC']),
        ];
      }),
      3,
      function(n) {
        return [
          String(n),
          '',
          '',
          (supEmpty && n === 1) ? 'Kebun Inti / PKS / Plasma / Swadaya / Trader' : '',
          '',
          '',
        ];
      }
    ),
    docRows: buildDocRows_(docMap),
  };
}
