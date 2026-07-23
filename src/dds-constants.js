/** Shared EUDR DDS field definitions. */

export const DDS_LIST_FIELDS = [
  'PLANT', 'SD NUMBER', 'BUYER NAME', 'CONTRACT NUMBER', 'HS CODE', 'PRODUCT',
  'QUANTITY (KG)', 'BL NUMBER', 'SHIPPING NAME', 'PORT OF DISCHARGE', 'SUPPLIER NAME',
];

export const DDS_FIELD_SECTIONS = {
  main: ['PLANT', 'SD NUMBER', 'BUYER NAME', 'CONTRACT NUMBER', 'HS CODE', 'PRODUCT', 'QUANTITY (KG)', 'BL NUMBER', 'SHIPPING NAME', 'PORT OF DISCHARGE', 'SUPPLIER NAME'],
  admin: ['DOCUMENT DATE', 'ANNUAL QUANTITY ESTIMATE (KG)'],
  identity: ['EXPORTER COMPANY NAME', 'EXPORTER ADDRESS', 'EXPORTER NIB', 'BUYER ADDRESS', 'BUYER EORI'],
  product: ['TRADE NAME'],
  shipment: ['PORT OF LOADING', 'TANK NUMBER', 'BATCH NUMBER', 'ETD DATE'],
  geoMeta: ['COUNTRY RISK CATEGORY', 'GEOJSON FILE REF'],
  risk: ['RA DOCUMENT NO', 'RA DATE', 'RA METHODOLOGY', 'RA OVERALL RISK', 'RA POST MITIGATION STATUS', 'RA MITIGATION SUMMARY'],
  sign: ['SIGNATORY NAME', 'SIGNATORY TITLE', 'SIGNATORY PLACE', 'SIGNATORY DATE'],
};

export const DDS_DATE_FIELDS = new Set([
  'DOCUMENT DATE', 'ETD DATE', 'RA DATE', 'SIGNATORY DATE',
]);

export const DDS_DOC_CODES = [
  { code: 'A1', label: 'Citra Satelit / Analisis NDVI per Plot Kebun' },
  { code: 'A2', label: 'Laporan Pemantauan Deforestasi (Forest Cover Analysis)' },
  { code: 'A3', label: 'Sertifikasi ISPO (Indonesia Sustainable Palm Oil)' },
  { code: 'A4', label: 'Sertifikasi RSPO (opsional)' },
  { code: 'A5', label: 'STDB (Surat Tanda Daftar Budidaya)' },
  { code: 'B1', label: 'HGU (Hak Guna Usaha) / SHM (Sertifikat Hak Milik)' },
  { code: 'B2', label: 'AMDAL / UKL-UPL (Izin Lingkungan)' },
  { code: 'B3', label: 'Izin Usaha Perkebunan (IUP)' },
  { code: 'B4', label: 'Izin Usaha Industri Pengolahan (IUIP) / NIB Refinery' },
  { code: 'B5', label: 'FPIC documentation (Free Prior Informed Consent)' },
  { code: 'B6', label: 'Labour / employment documentation' },
];

export const DDS_SUPPLIER_TYPES = ['Kebun Inti', 'PKS', 'Plasma', 'Swadaya', 'Trader'];

/** EUDR Annex I CN/HS headings — palm oil products (Regulation EU 2023/1115). */
export const DDS_HS_CODES = [
  { code: '1207', description: 'Palm nuts and kernels' },
  { code: '1511', description: 'Palm oil and its fractions, whether or not refined (not chemically modified) — CPO, RBD olein, stearin' },
  { code: '1513', description: 'Palm kernel / babassu oil and fractions, crude or refined, not chemically modified — incl. 1513 29 (refined, excl. crude)' },
  { code: '1516 20 / 1518 00', description: 'Hydrogenated, inter-esterified, re-esterified or elaidinised palm oil; chemically modified palm oil and fractions' },
  { code: '1520 00', description: 'Crude glycerol; glycerol waters and lyes (from palm processing)' },
  { code: 'ex 2905 45', description: 'Glycerol, purity ≥ 95%' },
  { code: '2915 70', description: 'Palmitic acid, stearic acid, their salts and esters' },
  { code: '2915 90', description: 'Saturated acyclic monocarboxylic acids/derivatives (with specified exclusions)' },
  { code: '3823 11 / 3823 12 / 3823 19', description: 'Industrial stearic acid; industrial oleic acid; other industrial monocarboxylic fatty acids and acid oils from refining' },
  { code: '2306', description: 'Oilcake and other solid residues from extraction of palm kernel/nut oil' },
];

export const DDS_HS_CUSTOM_STORAGE_KEY = 'sustain-dds-custom-hs-codes';

export const DDS_COUNTRY_RISK = ['Standard', 'Low', 'High'];
export const DDS_RA_RISK = ['Negligible', 'Low', 'Standard', 'High'];

export const DDS_FIELD_LABELS = {
  'PLANT': 'Plant',
  'SD NUMBER': 'SD Number',
  'BUYER NAME': 'Buyer Name',
  'CONTRACT NUMBER': 'Contract Number',
  'HS CODE': 'HS Code',
  'PRODUCT': 'Product',
  'QUANTITY (KG)': 'Quantity (KG)',
  'BL NUMBER': 'BL Number',
  'SHIPPING NAME': 'Shipping / Vessel Name',
  'PORT OF DISCHARGE': 'Port of Discharge',
  'SUPPLIER NAME': 'Primary Supplier Name',
  'DOCUMENT DATE': 'Document Date',
  'ANNUAL QUANTITY ESTIMATE (KG)': 'Annual Quantity Estimate (KG)',
  'EXPORTER COMPANY NAME': 'Exporter Company Name',
  'EXPORTER ADDRESS': 'Exporter Address',
  'EXPORTER NIB': 'Exporter NIB / SIUP',
  'BUYER ADDRESS': 'Buyer Address (EU)',
  'BUYER EORI': 'Buyer EORI (optional)',
  'TRADE NAME': 'Trade Name',
  'PORT OF LOADING': 'Port of Loading',
  'TANK NUMBER': 'Tank Number',
  'BATCH NUMBER': 'Batch Number',
  'ETD DATE': 'ETD Date',
  'COUNTRY RISK CATEGORY': 'Country Risk Category',
  'GEOJSON FILE REF': 'GeoJSON File Reference',
  'RA DOCUMENT NO': 'Risk Assessment Document No.',
  'RA DATE': 'Risk Assessment Date',
  'RA METHODOLOGY': 'Risk Assessment Methodology',
  'RA OVERALL RISK': 'Overall Risk Conclusion',
  'RA POST MITIGATION STATUS': 'Status After Mitigation',
  'RA MITIGATION SUMMARY': 'Main Mitigation Actions',
  'SIGNATORY NAME': 'Signatory Name',
  'SIGNATORY TITLE': 'Signatory Title',
  'SIGNATORY PLACE': 'Signatory Place',
  'SIGNATORY DATE': 'Signatory Date',
};
