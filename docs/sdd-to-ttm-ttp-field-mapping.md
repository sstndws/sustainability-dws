# Mapping: Form Registrasi SDD → Monitoring TTM/TTP

Dokumen ini memetakan **tiga template Excel** (KCP, Mill, Trader) ke kolom sheet **`Monitoring TTP/TTM`** (`TTP_HEADERS` di Apps Script).

**Prinsip baris TTP (sesuai UI detail modal):**  
Satu baris monitoring = **satu pemasok FFB** dari sheet **Traceability → B. FFB Supplier List**, dengan konteks **Group / Company / Mill (KCP)** dari **Main Form** submission yang sudah **Approved**.

---

## 1. Unit baris & sumber data

| Lapisan | Sheet Excel | Sheet SDD (setelah import/save) | Peran di TTP |
|--------|-------------|----------------------------------|--------------|
| Identitas registran | Main Form | `SDD_MAIN` | `GROUP NAME`, `COMPANY NAME`, `MILL NAME` (KCP/Mill/Trader) |
| Mill upstream (opsional) | Traceability → A. Mill List | `SDD_MILL_LIST` (`TML - …`) | Bisa dipakai untuk `UML ID` / konteks mill pemasok; **bukan** baris utama TTP |
| **Pemasok FFB (inti TTP)** | Traceability → B. FFB Supplier List | `SDD_FFB_LIST` (`FFB - …`) | **Isi utama** modal Supplier detail |

**Tidak disarankan:** hanya copy Main Form → TTP (akan kosong di FFB Supplier Group/Name seperti contoh screenshot).

---

## 2. Mapping kolom TTP ← SDD (FFB row + main context)

Legenda: **✅** = map langsung · **⚙️** = isi manual / monitoring · **—** = tidak ada di form registrasi

| # | Kolom Monitoring TTM/TTP | Kolom SDD / Excel (sumber) | KCP | Mill | Trader |
|---|--------------------------|----------------------------|-----|------|--------|
| 1 | `NO` | — (auto increment di TTP) | ⚙️ | ⚙️ | ⚙️ |
| 2 | `COMPANY CODE` | — | ⚙️ | ⚙️ | ⚙️ |
| 3 | `GROUP NAME` | Main: `Group Name` | ✅ | ✅ | ✅ |
| 4 | `COMPANY NAME` | Main: `Company Name` | ✅ | ✅ | ✅ |
| 5 | `MILL NAME` | Main: `KCP Name` → SDD `Mill Name` (KCP) · Main: `Mill Name` (Mill) · Trader: `Company Name` atau kosong* | ✅ | ✅ | ⚠️* |
| 6 | `UML ID` | Traceability A: `UML ID` → `TML - UML ID` (mill terkait) · atau — | ⚠️ | ⚠️ | — |
| 7 | `FFB SUPPLIER GROUP NAME` | Traceability B: `SUPPLIER GROUP NAME` → `FFB - Supplier Group Name` | ✅ | ✅ | ✅ |
| 8 | `FFB SUPPLIER NAME` | Traceability B: `SUPPLIER NAME` → `FFB - Supplier Name` | ✅ | ✅ | ✅ |
| 9 | `CATEGORY` | Traceability B: `SUPPLIER CATEGORY` → `FFB - Supplier Category` | ✅ | ✅ | ✅ |
| 10 | `LAT` | Traceability B: `Lat` → `FFB - Latitude` · fallback Main: `Latitude` | ✅ | ✅ | ⚠️ |
| 11 | `LONG` | Traceability B: `Long` → `FFB - Longitude` · fallback Main: `Longitude` | ✅ | ✅ | ⚠️ |
| 12 | `VILLAGE ID` | — | — | — | — |
| 13 | `VILLAGE` | Traceability B: `VILLAGE` → `FFB - Village` | ✅ | ✅ | ✅ |
| 14 | `SUBDISTRICT` | Traceability B: `SUB DISTRICT` → `FFB - Sub District` | ✅ | ✅ | ✅ |
| 15 | `DISTRICT` | Traceability B: `DISTRICT` → `FFB - District` | ✅ | ✅ | ✅ |
| 16 | `PROVINCE` | — (form tidak punya kolom province di FFB) | ⚙️ | ⚙️ | ⚙️ |
| 17 | `CONCESION AREA` | Traceability B: `CONSESION AREA` → `FFB - Concession Area (Ha)` | ✅ | ✅ | ✅ |
| 18 | `PLANTED AREA` | Traceability B: `PLANTED AREA` → `FFB - Planted Area (Ha)` | ✅ | ✅ | ✅ |
| 19 | `NUMBER OD SMALLHOLDERS` | Traceability B: `NUMBERS OF SMALLHOLDERS` → `FFB - Number of Smallholders` | ✅ | ✅ | ✅ |
| 20 | `TAHUN TANAM` | Traceability B: `PLANTED YEAR` → `FFB - Planted Year` | ✅ | ✅ | ✅ |
| 21 | `LEGALITAS` | Traceability B: `LEGALITY` → `FFB - Legality` | ✅ | ✅ | ✅ |
| 22 | `ISPO (Y/N)` | Traceability B: `ISPO (Y/N)` → `FFB - ISPO (Y/N)` | ✅ | ✅ | ✅ |
| 23 | `RSPO (Y/N)` | Traceability B: `RSPO (Y/N)` → `FFB - RSPO (Y/N)` | ✅ | ✅ | ✅ |
| 24 | `ISCC (Y/N)` | Traceability B: `ISCC (Y/N)` → `FFB - ISCC (Y/N)` | ✅ | ✅ | ✅ |
| 25 | `FFB SUPPLY to MILL (TON)` | Traceability B: kolom supply FFB (ton) → `FFB - Total Supply FFB (Ton)` | ✅ | ✅ | ✅ |
| 26 | `CONVERSION FFB to PK (5%)` | — | ⚙️ | ⚙️ | ⚙️ |
| 27 | `PK SUPPLY to KCP` | Main KCP: produk PKE/CPKO (bukan per FFB) | ⚙️ | — | — |
| 28 | `CONVERSION FFB to CPO (20%)` | — | ⚙️ | ⚙️ | ⚙️ |
| 29 | `CPO SUPPLY to REFINERY` | — | ⚙️ | ⚙️ | ⚙️ |
| 30 | `% PK TRACEABLE` | — | ⚙️ | ⚙️ | ⚙️ |
| 31 | `% CPO TRACEABLE` | — | ⚙️ | ⚙️ | ⚙️ |
| 32 | `Total PK % Traceable` | — | ⚙️ | ⚙️ | ⚙️ |
| 33 | `Total CPO % Traceable` | — | ⚙️ | ⚙️ | ⚙️ |
| 34 | `MSD` | — | ⚙️ | ⚙️ | ⚙️ |
| 35 | `PK Traceable Volume` | — | ⚙️ | ⚙️ | ⚙️ |
| 36 | `CPO Traceable Volume` | — | ⚙️ | ⚙️ | ⚙️ |

\* **Trader:** Main Form tidak punya `Mill Name` / `KCP Name`; section B sering kosong di template. Opsi: `MILL NAME` = `Company Name` trader, atau kosong sampai FFB diisi.

**Metadata sync (disarankan ditambah di TTP, belum di `TTP_HEADERS`):**  
`submission_id`, `ffb_line_id`, `supplier_type`, `synced_at` — untuk upsert saat re-approve.

---

## 3. Traceability B — header Excel → kolom SDD (`SDD_FFB_LIST`)

Header baris di ketiga template (sama untuk KCP & Mill; Trader struktur sama):

| Kolom Excel (Traceability B) | Kolom SDD setelah import |
|-------------------------------|---------------------------|
| MILL NAME | `FFB - Mill Name` |
| SUPPLIER GROUP NAME | `FFB - Supplier Group Name` |
| SUPPLIER NAME | `FFB - Supplier Name` |
| VILLAGE | `FFB - Village` |
| SUB DISTRICT | `FFB - Sub District` |
| DISTRICT | `FFB - District` |
| SUPPLIER CATEGORY | `FFB - Supplier Category` |
| CONSESION AREA | `FFB - Concession Area (Ha)` |
| PLANTED AREA | `FFB - Planted Area (Ha)` |
| NUMBERS OF SMALLHOLDERS | `FFB - Number of Smallholders` |
| PLANTED YEAR | `FFB - Planted Year` |
| LEGALITY | `FFB - Legality` |
| Lat | `FFB - Latitude` |
| Long | `FFB - Longitude` |
| ISPO (Y/N) | `FFB - ISPO (Y/N)` |
| RSPO (Y/N) | `FFB - RSPO (Y/N)` |
| ISCC (Y/N) | `FFB - ISCC (Y/N)` |
| (kolom tonase FFB di akhir baris) | `FFB - Total Supply FFB (Ton)` |

**Konteks baris (diisi dari `SDD_MAIN` yang sama `submission_id`):**

| TTP | SDD_MAIN |
|-----|----------|
| `GROUP NAME` | `Group Name` |
| `COMPANY NAME` | `Company Name` |
| `MILL NAME` | `Mill Name` (= KCP Name / Mill Name / fallback Trader) |

| TTP (opsional) | SDD FFB |
|----------------|---------|
| Mill pemasok terdekat | `FFB - Mill Name` (bisa sama dengan mill di kolom 0 traceability, **bukan** selalu = MILL NAME registran) |

---

## 4. Contoh konkret — KCP (`Form_Registrasi_Supplier_KCP.xlsx`)

### Main Form → konteks TTP

| Excel (Main) | SDD_MAIN | → TTP |
|--------------|----------|-------|
| Group Name: PT PASIFIK AGRO SENTOSA | `Group Name` | `GROUP NAME` |
| Company Name: PT CIPTA USAHA SEJATI | `Company Name` | `COMPANY NAME` |
| KCP Name: SEGATI KCP | `Mill Name` | `MILL NAME` |
| KCP Coordinate Lat/Long | `Latitude` / `Longitude` | fallback `LAT`/`LONG` jika FFB tanpa koordinat |

### Traceability B baris 1 (contoh data) → satu baris TTP

| Excel (baris data) | → TTP |
|--------------------|-------|
| MILL NAME: MITRASARI PRIMA | (konteks `FFB - Mill Name`; bukan `MILL NAME` registran) |
| SUPPLIER GROUP NAME: TOGAR | `FFB SUPPLIER GROUP NAME` |
| SUPPLIER NAME: BATU AMPAR | `FFB SUPPLIER NAME` |
| VILLAGE: LANGGAM | `VILLAGE` |
| SUB DISTRICT: PELALAWAN | `SUBDISTRICT` |
| SUPPLIER CATEGORY: Own Estate | `CATEGORY` |
| PLANTED AREA: 2450 | `PLANTED AREA` |
| … | … |
| Supply: 15000 | `FFB SUPPLY to MILL (TON)` |

**Subtitle modal TTP** (contoh: `SAMPLING GROUP · ABDI BORNEO · …`) = gabungan `GROUP NAME` · `COMPANY NAME` · `MILL NAME` **registran**, bukan nama FFB.

---

## 5. Per tipe supplier

### KCP

| Bagian Excel | Masuk TTP? |
|--------------|------------|
| Main A–D (profil, KCP info, produk, sertifikasi, legalitas) | Konteks + tidak duplikasi ke setiap baris FFB kecuali field di tabel §2 |
| Traceability A (Mill List) | Opsional: referensi `UML ID`; **bukan** baris TTP utama |
| Traceability B (FFB List) | **1 baris TTP per baris FFB valid** |

### Mill

Sama dengan KCP, beda label Main: `Mill Name` / `Mill Address` / `Mill Coordinate` (bukan KCP).

| Tambahan Mill Main | TTP |
|--------------------|-----|
| Product to be Supply (CPO/PK/Other) | Tidak map ke FFB row; relevan untuk monitoring produk mill, bukan kolom FFB modal |
| Sterilizer Type | — |

### Trader

| Bagian Excel | Masuk TTP? |
|--------------|------------|
| Main (profil + produk supply, sertifikasi, NIB saja) | `GROUP`/`COMPANY`; `MILL NAME` = perlu keputusan bisnis |
| Traceability A | Template sample: dropdown kosong |
| Traceability B | **Wajib diisi** jika ingin baris TTP; tanpa B → **tidak ada baris TTP** |

---

## 6. Traceability A (Mill List) — mapping opsional

Jika tim ingin baris terpisah per **mill upstream** (bukan FFB):

| TTP | Traceability A / `SDD_MILL_LIST` |
|-----|----------------------------------|
| `COMPANY NAME` | `TML - Company Name` |
| `MILL NAME` | `TML - Mill Name` |
| `UML ID` | `TML - UML ID` |
| `VILLAGE` | `TML - Village` |
| `SUBDISTRICT` | `TML - Sub District` |
| `DISTRICT` | `TML - District` |
| `LAT` / `LONG` | `TML - Latitude` / `TML - Longitude` |
| `LEGALITAS` | `TML - Legality` |
| `ISPO/RSPO/ISCC` | `TML - ISPO/RSPO/ISCC (Y/N)` |
| Supply CPO/PK | `TML - Total Supply CPO/PK (Ton)` → bisa ke `CPO SUPPLY` / konversi (⚙️ aturan bisnis) |

**Catatan:** UI detail TTP saat ini menekankan **FFB Supplier**, jadi mode ini perlu keputusan produk (dual registry vs FFB-only).

---

## 7. Kolom TTP = monitoring lanjutan (kosong saat sync Approved)

Isi belakangan oleh tim TTM/TTP:

- `% PK TRACEABLE`, `% CPO TRACEABLE`, `Total PK/CPO % Traceable`
- `PK Traceable Volume`, `CPO Traceable Volume`
- `MSD`, `COMPANY CODE`, `NO`
- `PROVINCE` (kecuali ditambah ke form / geocoding)
- Konversi & alokasi: `CONVERSION FFB to PK/CPO`, `PK SUPPLY to KCP`, `CPO SUPPLY to REFINERY`

---

## 8. Implementasi sync (backend)

**Status: implemented** in `GoogleAppsScript-backend-v3-full.gs` → `syncTtpFromApprovedSubmission_()`.

1. **Trigger:** `statusSDD = APPROVED` via `setSubmissionStatus` / `updateSubmission` (sama pola Contact List).  
2. **Loop:** setiap baris `SDD_FFB_LIST` untuk `submission_id`.  
3. **Map:** tabel §2 + konteks §3.  
4. **Skip:** baris FFB kosong / `PLEASE SELECT`.  
5. **Upsert:** by `submission_id` + `ffb_line_id` (fallback: nama FFB).  
6. **Preserve:** kolom monitoring TTP (`% TRACEABLE`, volume, MSD, dll.) tidak ditimpa jika sudah terisi.  
7. **Trader tanpa FFB:** sync dilewati (`reason: no_ffb_rows`).

Kolom metadata di sheet TTP: `submission_id`, `ffb_line_id`, `supplier_type`, `synced_at`, `synced_by`.

---

*Generated from: `Form_Registrasi_Supplier_KCP.xlsx`, `Form_Registrasi_Supplier_Mill (1).xlsx`, `Form_Registrasi_Supplier_Trader.xlsx`, `TTP_HEADERS` in `GoogleAppsScript-backend-v3-full.gs`, import logic in `src/main.js`.*
