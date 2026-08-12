# Barcode Label Generator — Tom & Jerry

Web app **100% client-side** (HTML + JavaScript) untuk generate barcode Code128 dari data produk Excel, export ke Word siap print label **Tom & Jerry No. 103**, dan update Excel dengan kolom **Kode Barcode**.

Tidak perlu Python/server — **cocok untuk GitHub Pages**.

## Fitur

- Upload Excel dengan kolom **Nama** produk
- Generate kode barcode otomatis (`BR001`, `BR002`, ...)
- Export **Word (.docx)** layout label Tom & Jerry No. 103 (64 × 32 mm, **12 label/halaman**, 3×4)
- Export **Excel (.xlsx)** yang sudah ditambah kolom Kode Barcode
- Download hasil dalam satu file ZIP
- Semua proses di browser — data tidak di-upload ke server

## Deploy ke GitHub Pages

1. Push repo ke GitHub
2. Buka **Settings → Pages**
3. Source: **Deploy from a branch**
4. Branch: `main` / folder: **`/` (root)**
5. Akses di: `https://<username>.github.io/<repo-name>/`

File utama ada di root:
- `index.html`
- `css/style.css`
- `js/*.js`

## Jalankan Lokal

**Opsi 1 — Buka langsung:** double-click `index.html` (perlu koneksi internet untuk library CDN).

**Opsi 2 — Local server (disarankan):**

```bash
python3 -m http.server 8080
```

Buka `http://localhost:8080`

## Format Excel Input

| Nama |
|------|
| Sabun Lifebuoy |
| Shampoo Clear 170ml |
| Minyak Goreng 1L |

## Format Excel Output

| Nama | Kode Barcode |
|------|--------------|
| Sabun Lifebuoy | BR001 |
| Shampoo Clear 170ml | BR002 |
| Minyak Goreng 1L | BR003 |

## Output ZIP

- `labels.docx` — dokumen Word siap print
- `products_updated.xlsx` — Excel dengan kolom Kode Barcode

## Tips Print Label

1. Test print di **kertas HVS** dulu sebelum pakai stiker asli
2. Setting printer: scale **100%**, ukuran kertas **custom 20,5 × 16,5 cm** (bukan A4)
3. Jika posisi meleset, sesuaikan margin di `LABEL_CONFIG` pada `js/app.js`

## Desain Label

Setiap label berisi:
- **Logo** (upload opsional) di kiri
- **Barcode Code128** + kode (contoh: `SHM - ME - 0261`) di kanan
- **Footer** dengan teks & warna dinamis

Layout: Tom & Jerry **No. 103** lembar kuning — **12 label/halaman** (3×4), label **64 × 32 mm**, kertas **20,5 × 16,5 cm**. Saat print orientasi **landscape** (sisi panjang horizontal).

## Library (CDN)

- [SheetJS](https://sheetjs.com/) — baca/tulis Excel
- [JsBarcode](https://github.com/lindell/JsBarcode) — generate Code128
- [JSZip](https://stuk.github.io/jszip/) — buat file ZIP & Word (.docx)
- [FileSaver.js](https://github.com/eligrey/FileSaver.js/) — download file

## Struktur Project

```
barcode/
├── index.html
├── css/style.css
└── js/app.js
```
