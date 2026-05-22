"""
data_processor.py - Modul pemrosesan dan agregasi data
========================================================
Bertanggung jawab untuk:
- Membaca file CSV atau JSON menjadi DataFrame Pandas
- Melakukan deteksi kolom otomatis (kategori vs numerik)
- Mengagregas data (sum, mean, atau count per kategori)
- Menghasilkan bullet points ringkasan untuk executive summary
"""

import pandas as pd
import io
import csv
from typing import Optional, Tuple, List


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI: load_data
# ─────────────────────────────────────────────────────────────────────────────
def load_data(uploaded_file) -> pd.DataFrame:
    """
    Membaca file yang diupload (CSV atau JSON) menjadi Pandas DataFrame.

    Parameter:
    ----------
    uploaded_file : UploadedFile (Streamlit)
        Objek file yang diterima dari st.file_uploader.

    Returns:
    --------
    pd.DataFrame
        DataFrame yang berisi data dari file yang diupload.

    Raises:
    -------
    ValueError
        Jika format file tidak didukung atau file kosong.
    """
    filename = uploaded_file.name.lower()
    content = uploaded_file.read()

    if filename.endswith(".csv"):
        # ── Langkah 1: Decode bytes → str dengan fallback encoding ───────────
        decoded_str = None
        for encoding in ("utf-8", "utf-8-sig", "latin1", "cp1252"):
            try:
                decoded_str = content.decode(encoding)
                break
            except (UnicodeDecodeError, LookupError):
                continue
        if decoded_str is None:
            decoded_str = content.decode("latin1", errors="replace")

        # ── Langkah 2: Deteksi delimiter via csv.Sniffer ─────────────────────
        detected_sep = ","
        try:
            dialect = csv.Sniffer().sniff(decoded_str[:4096], delimiters=",;\t|")
            detected_sep = dialect.delimiter
        except csv.Error:
            # Fallback: hitung kemunculan tiap delimiter di 1024 karakter pertama
            counts = {d: decoded_str[:1024].count(d) for d in [",", ";", "\t"]}
            detected_sep = max(counts, key=counts.get)

        # ── Langkah 3: Parse menggunakan csv stdlib → DataFrame ───────────────
        # Menghindari sepenuhnya pandas CSV parser yang bermasalah di pandas 3.x
        try:
            reader = csv.reader(io.StringIO(decoded_str), delimiter=detected_sep)
            rows = [row for row in reader if any(cell.strip() for cell in row)]
            if not rows:
                raise ValueError("File CSV tidak mengandung baris data.")
            headers = [h.strip() for h in rows[0]]
            data = rows[1:]
            df = pd.DataFrame(data, columns=headers)
        except Exception as csv_err:
            raise ValueError(f"Gagal mem-parse CSV: {csv_err}")

    else:
        raise ValueError(
            f"Format file '{filename}' tidak didukung. "
            "Hanya file CSV (.csv) yang diterima."
        )

    if df.empty:
        raise ValueError("File yang diupload tidak mengandung data (kosong).")

    # ── FIX 1: Bersihkan spasi tersembunyi di semua nama kolom ───────────────
    # Menangani kasus seperti ' Bulan' atau 'Bulan ' yang menyebabkan KeyError.
    df.columns = df.columns.str.strip()

    # ── FIX 2: Hapus kolom indeks otomatis 'Unnamed: X' ──────────────────────
    # Kolom ini muncul ketika CSV di-export dengan df.to_csv() tanpa index=False.
    # Pola regex menangkap semua variannya: 'Unnamed: 0', 'Unnamed: 1', dst.
    unnamed_cols = [c for c in df.columns if c.startswith("Unnamed:")]
    if unnamed_cols:
        df = df.drop(columns=unnamed_cols)

    # Coba konversi kolom yang mungkin berformat angka tapi bertipe string
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col], errors="raise")
        except Exception:
            pass

    return df


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI: _detect_columns
# ─────────────────────────────────────────────────────────────────────────────
def _detect_columns(df: pd.DataFrame) -> Tuple[str, str]:
    """
    Mendeteksi secara otomatis kolom kategori (X) dan kolom numerik (Y)
    dari DataFrame jika tidak ditentukan oleh pengguna.

    Strategi:
    - Kolom X (kategori): kolom non-numerik pertama, atau kolom numerik
      pertama jika semua numerik (dijadikan string).
    - Kolom Y (nilai): kolom numerik pertama yang bukan kolom X.

    Parameter:
    ----------
    df : pd.DataFrame
        DataFrame sumber.

    Returns:
    --------
    Tuple[str, str]
        (nama_kolom_x, nama_kolom_y)
    """
    numeric_cols = df.select_dtypes(include="number").columns.tolist()
    non_numeric_cols = df.select_dtypes(exclude="number").columns.tolist()

    if not numeric_cols:
        raise ValueError(
            "Tidak ditemukan kolom numerik dalam data. "
            "Pastikan data memiliki minimal satu kolom angka untuk divisualisasikan."
        )

    # Pilih kolom X: prioritaskan kolom non-numerik (kategori)
    if non_numeric_cols:
        x_col = non_numeric_cols[0]
        # Pilih kolom Y: kolom numerik pertama
        y_col = numeric_cols[0]
    else:
        # Semua kolom numerik: gunakan kolom pertama sebagai label X
        x_col = numeric_cols[0]
        y_col = numeric_cols[1] if len(numeric_cols) > 1 else numeric_cols[0]

    return x_col, y_col


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI: aggregate_data
# ─────────────────────────────────────────────────────────────────────────────
def aggregate_data(
    df: pd.DataFrame,
    x_col: Optional[str] = None,
    y_col: Optional[str] = None,
) -> Tuple[pd.DataFrame, str, str, str]:
    """
    Mengagregas DataFrame berdasarkan kolom X (kategori) terhadap kolom Y (nilai).
    Jika data sudah teragregasi (satu baris per kategori), nilai dikembalikan apa adanya.

    Parameter:
    ----------
    df : pd.DataFrame
        DataFrame mentah.
    x_col : str, optional
        Nama kolom kategori (sumbu X). Jika None, dideteksi otomatis.
    y_col : str, optional
        Nama kolom nilai numerik (sumbu Y). Jika None, dideteksi otomatis.

    Returns:
    --------
    Tuple[pd.DataFrame, str, str, str]
        (df_agregasi, nama_x_col, nama_y_col, label_agregasi)
        label_agregasi: string seperti "Total Penjualan" atau "Rata-rata Nilai"
    """
    # Deteksi kolom jika tidak disediakan user
    if x_col is None or y_col is None:
        x_col_auto, y_col_auto = _detect_columns(df)
        x_col = x_col or x_col_auto
        y_col = y_col or y_col_auto

    # Validasi bahwa kolom tersebut ada di DataFrame
    if x_col not in df.columns:
        raise ValueError(f"Kolom '{x_col}' tidak ditemukan dalam data. "
                         f"Kolom yang tersedia: {list(df.columns)}")
    if y_col not in df.columns:
        raise ValueError(f"Kolom '{y_col}' tidak ditemukan dalam data. "
                         f"Kolom yang tersedia: {list(df.columns)}")

    # Pastikan kolom Y bersifat numerik
    if not pd.api.types.is_numeric_dtype(df[y_col]):
        raise ValueError(f"Kolom '{y_col}' harus bertipe numerik untuk divisualisasikan.")

    # Cek apakah perlu agregasi (ada duplikasi nilai di kolom X)
    if df[x_col].duplicated().any():
        # Lakukan agregasi SUM per kategori
        df_agg = (
            df.groupby(x_col, as_index=False)[y_col]
            .sum()
            .rename(columns={y_col: y_col})
        )
        df_agg = df_agg.sort_values(y_col, ascending=False).reset_index(drop=True)
        agg_label = f"Total {y_col.replace('_', ' ').title()}"
    else:
        # Data sudah unik per kategori, tidak perlu agregasi
        df_agg = df[[x_col, y_col]].copy().reset_index(drop=True)
        agg_label = y_col.replace("_", " ").title()

    return df_agg, x_col, y_col, agg_label


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI: generate_summary_bullets
# ─────────────────────────────────────────────────────────────────────────────
def generate_summary_bullets(
    df_raw: pd.DataFrame,
    df_agg: pd.DataFrame,
    x_col: str,
    y_col: str,
    agg_label: str,
    report_title: str,
) -> List[str]:
    """
    Menghasilkan daftar bullet points untuk Executive Summary slide
    berdasarkan statistik deskriptif dari data yang teragregasi.

    Parameter:
    ----------
    df_raw : pd.DataFrame
        Data mentah asli (untuk info jumlah baris).
    df_agg : pd.DataFrame
        Data yang sudah teragregasi.
    x_col : str
        Nama kolom kategori.
    y_col : str
        Nama kolom nilai.
    agg_label : str
        Label agregasi (misal: "Total Penjualan").
    report_title : str
        Judul laporan untuk konteks kalimat.

    Returns:
    --------
    List[str]
        Daftar string bullet points siap pakai.
    """
    total_val = df_agg[y_col].sum()
    mean_val = df_agg[y_col].mean()
    max_val = df_agg[y_col].max()
    min_val = df_agg[y_col].min()
    max_cat = df_agg.loc[df_agg[y_col].idxmax(), x_col]
    min_cat = df_agg.loc[df_agg[y_col].idxmin(), x_col]
    num_categories = len(df_agg)

    bullets = [
        f"Laporan ini menganalisis {df_raw.shape[0]:,} baris data dengan "
        f"{num_categories} kategori unik berdasarkan kolom '{x_col}'.",

        f"{agg_label} keseluruhan tercatat sebesar {total_val:,.2f}, "
        f"dengan rata-rata per kategori sebesar {mean_val:,.2f}.",

        f"Kategori dengan performa tertinggi adalah '{max_cat}' "
        f"dengan nilai {max_val:,.2f}.",

        f"Kategori dengan nilai terendah adalah '{min_cat}' "
        f"dengan nilai {min_val:,.2f}.",

        f"Rentang nilai antara tertinggi dan terendah adalah "
        f"{max_val - min_val:,.2f} ({((max_val - min_val) / mean_val * 100):.1f}% dari rata-rata).",

        f"Data dalam laporan ini dapat digunakan sebagai dasar perencanaan "
        f"dan pengambilan keputusan strategis ke depannya.",
    ]

    return bullets
