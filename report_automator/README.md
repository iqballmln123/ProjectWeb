# 📡 Telkomsel Coverage Report Automator

Aplikasi berbasis **Streamlit** untuk mengotomasi pembuatan laporan **Coverage Activity Report** Telkomsel dalam format PowerPoint (`.pptx`). Upload data Excel/CSV → preview per site → generate presentasi lengkap secara otomatis.

---

## 🗂️ Struktur Project

```
report_automator/
├── app.py                  # Entry point Streamlit (UI 3-step)
├── data_processor.py       # Modul load & validasi data Excel/CSV
├── chart_generator.py      # Modul pembuatan chart (Matplotlib)
├── ppt_generator.py        # Modul pembuatan file .pptx (python-pptx)
├── requirements.txt        # Dependensi Python
└── sample_data/
    ├── data_penjualan.csv              # Contoh data CSV sederhana
    └── test_2_multi_kolom.csv          # Contoh data multi-kolom
```

---

## 🚀 Cara Menjalankan

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Jalankan aplikasi
```bash
streamlit run app.py
```

### 3. Buka di browser
```
http://localhost:8501
```

---

## ✨ Fitur Utama

| Fitur | Detail |
|-------|--------|
| **Upload Excel / CSV** | Mendukung `.xlsx`, `.xls`, dan `.csv` |
| **Multi-sheet Excel** | Deteksi otomatis sheet yang direkomendasikan |
| **Upload Gambar per Site** | Payload Chart, Maps/COVMO, hingga 3 Support Images |
| **Preview Data** | Tabel preview 10 baris pertama sebelum generate |
| **Header Profesional** | Bar navy dengan 2 baris: PURPOSE\|SOW (putih) + \[City\] Plan Action (merah maroon) |
| **Logo Telkomsel** | Otomatis tampil di pojok kanan atas setiap slide |
| **Slide Cover & Penutup** | Slide pertama dan terakhir otomatis digenerate |
| **Download .pptx** | Tombol download langsung dari UI |

---

## 📋 Format Kolom Excel yang Dibutuhkan

Setiap baris = 1 site = 1 slide PPT.

| Kolom | Keterangan |
|-------|-----------|
| `NO` | Nomor urut |
| `SITE ID` | ID unik site (e.g. `BDG212`) |
| `SITE NAME` | Nama site (e.g. `BDG212_SETRAMURNI-DMT`) |
| `FINDING` | Temuan di lapangan (bisa multi-baris) |
| `CITY` | Kota / area (tampil di header baris merah) |
| `PURPOSE HEADER` | Judul tujuan (tampil di header baris putih) |
| `SOW` | Scope of Work (tampil di header baris putih, setelah PURPOSE) |
| `PLAN ACTION` | Rencana tindakan (tampil di header baris merah & section ②) |
| `SUPPORT NEEDED` | Kebutuhan dukungan (material, dll.) |
| `GOALS` | Target hasil |
| `INCREAMENT PAYLOAD AND REVENUE` | Estimasi peningkatan payload & revenue |
| `LONGITUDE` | Koordinat bujur site |
| `LATITUDE` | Koordinat lintang site |

> Untuk field dengan beberapa poin, pisahkan dengan baris baru (Enter di Excel).

---

## 🖼️ Struktur Slide yang Dihasilkan

```
Slide 1        : Cover (Judul, Sub-judul, Pembuat, Tanggal)
Slide 2..N     : Satu slide per baris data (site)
Slide terakhir : Penutup
```

### Layout Slide Data (per Site)

```
┌──────────────────────────────────────────────── HEADER NAVY ──────────────────────────────────── [Telkomsel] ─┐
│  PURPOSE HEADER  |  SOW                                                                     (putih bold)       │
│  [City]  Plan Action baris pertama                                                         (merah maroon)      │
├───────────────────────┬───────────────────────────────────┬──────────────────────────────────────────────────┤
│ ① FINDING             │  ⑥ PAYLOAD site Surrounding       │                                                  │
│   • bullet            │     [chart image]                 │  [Support Image 1]                               │
│ ② PLAN ACTION         ├───────────────────────────────────┤                                                  │
│   teks                │  ⑦ MAPS Preview & COVMO           ├──────────────────────────────────────────────────┤
│ ③ SUPPORT NEEDED      │     [maps/COVMO image]            │  [Support Image 2]  │  [Support Image 3]         │
│   • bullet            │                                   │                     │                            │
│ ④ GOALS               │                                   │                     │                            │
│   - bullet            │                                   │                     │                            │
│ ⑤ INCREMENT P&R       │                                   │                     │                            │
│   • bullet            │                                   │                     │                            │
└───────────────────────┴───────────────────────────────────┴──────────────────────────────────────────────────┘
│                                FOOTER (navy)                                                                  │
└───────────────────────────────────────────────────────────────────────────────────────────────────────────────┘
```

---

## 🏗️ Arsitektur Modul

```
app.py  (UI 3-step)
  ├── data_processor.get_excel_info()     → Info sheet & rekomendasi
  ├── data_processor.load_excel_sheet()   → Baca sheet tertentu → DataFrame
  ├── data_processor.load_data()          → Baca CSV → DataFrame
  └── ppt_generator.generate_pptx()      → .pptx bytes
        ├── _slide_cover()               → Slide 1 (Cover)
        ├── _slide_data_site()           → Slide per site (2..N)
        └── _slide_penutup()             → Slide terakhir (Penutup)
```

---

## ⚙️ Konfigurasi (Sidebar)

| Setting | Keterangan |
|---------|-----------|
| **Judul Laporan** | Tampil di Slide Cover (default: *Tracking Activity NOP 2026*) |
| **Sub-judul / Departemen** | Sub-teks di bawah judul (default: *Divisi Network Operation*) |
| **Nama Pembuat** | Nama tim / individu pembuat laporan |
| **Tanggal Laporan** | Otomatis hari ini, bisa diubah manual |

---

## 🎨 Desain Header Slide

| Elemen | Konten | Style |
|--------|--------|-------|
| Background | Seluruh header bar | Navy biru gelap `#1B2A4A` |
| Baris 1 | `PURPOSE HEADER  \|  SOW` | Putih, bold, 13pt |
| Baris 2 | `[City]  Plan Action` | Merah maroon, italic, 9.5pt |
| Logo | `Telkomsel` pojok kanan | Merah, bold italic, Times New Roman, 20pt |
