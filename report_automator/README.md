# 📊 Report Automator

Aplikasi automasi laporan berbasis **Streamlit** yang mengkonversi data CSV/JSON menjadi presentasi PowerPoint (.pptx) secara otomatis — lengkap dengan bar chart profesional dan executive summary.

---

## 🗂️ Struktur Project

```
report_automator/
├── app.py                  # Entry point Streamlit
├── data_processor.py       # Modul load, agregasi & ringkasan data
├── chart_generator.py      # Modul pembuatan bar chart (Matplotlib)
├── ppt_generator.py        # Modul pembuatan file .pptx (python-pptx)
├── requirements.txt        # Dependensi Python
└── sample_data/
    ├── data_penjualan.csv  # Contoh data CSV (penjualan bulanan)
    └── data_departemen.json # Contoh data JSON (anggaran per departemen)
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
| **File Uploader** | Mendukung CSV (berbagai delimiter & encoding) dan JSON |
| **Auto-detect Kolom** | Deteksi otomatis kolom kategori (X) dan numerik (Y) |
| **Agregasi Cerdas** | SUM otomatis jika ada duplikasi; passthrough jika data sudah unik |
| **Bar Chart Korporat** | Minimalis, gradient opacity, value labels, watermark |
| **PPT 3 Slide** | Title → Executive Summary → Visualisasi + Analisis |
| **Download Langsung** | Tombol download file .pptx dari dalam UI |

---

## 📋 Format Data yang Didukung

### CSV
```csv
bulan,penjualan,target
Januari,125000000,100000000
Februari,98000000,100000000
```

### JSON (Array of Objects)
```json
[
  {"departemen": "Engineering", "anggaran": 2500000000},
  {"departemen": "Marketing", "anggaran": 1800000000}
]
```

---

## 🏗️ Arsitektur Modul

```
app.py
  ├── data_processor.load_data()          → Baca CSV/JSON → DataFrame
  ├── data_processor.aggregate_data()     → Agregasi SUM per kategori
  ├── data_processor.generate_summary_bullets() → Bullet points statistik
  ├── chart_generator.create_bar_chart()  → PNG bytes bar chart
  └── ppt_generator.generate_pptx()      → .pptx bytes (3 slide)
```

---

## ⚙️ Konfigurasi (Sidebar)

- **Judul Laporan** — Tampil di Slide 1
- **Sub-judul / Departemen** — Tampil di bawah judul
- **Nama Pembuat** — Tercantum di slide
- **Tanggal Laporan** — Otomatis dari hari ini, bisa diubah
- **Warna Batang Chart** — Color picker bebas
- **Kolom X/Y** — Nama kolom custom atau kosongkan untuk auto-detect
