"""
chart_generator.py - Modul pembuatan bar chart korporat
=========================================================
Menggunakan Matplotlib untuk menghasilkan bar chart dengan
gaya minimalis dan profesional (corporate style).

Grafik disimpan ke buffer BytesIO agar bisa digunakan
langsung di Streamlit maupun python-pptx tanpa menyimpan ke disk.
"""

import io
import matplotlib
matplotlib.use("Agg")  # Non-interactive backend untuk server/Streamlit
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.font_manager as fm
import pandas as pd
import numpy as np
from typing import Optional


# ─────────────────────────────────────────────────────────────────────────────
# KONSTANTA: Palet warna korporat & konfigurasi style
# ─────────────────────────────────────────────────────────────────────────────

CORPORATE_STYLE = {
    "bg_color": "#FFFFFF",          # Latar belakang putih bersih
    "grid_color": "#E5E9F0",        # Warna gridline halus
    "text_color": "#2C3E50",        # Warna teks utama (dark navy)
    "title_color": "#1A252F",       # Warna judul (lebih gelap)
    "subtitle_color": "#7F8C8D",    # Warna label & subtitle
    "bar_alpha": 0.88,              # Transparansi batang
    "bar_edge_color": "white",      # Border batang
    "value_label_color": "#2C3E50", # Warna label nilai di atas batang
    "figure_dpi": 150,              # DPI output gambar
}


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI: _format_value
# ─────────────────────────────────────────────────────────────────────────────
def _format_value(val: float) -> str:
    """
    Memformat angka menjadi string yang mudah dibaca (K, M, B).

    Parameter:
    ----------
    val : float
        Nilai numerik yang akan diformat.

    Returns:
    --------
    str
        String terformat, contoh: "1.2M", "450K", "1.5B".
    """
    abs_val = abs(val)
    if abs_val >= 1_000_000_000:
        return f"{val / 1_000_000_000:.1f}B"
    elif abs_val >= 1_000_000:
        return f"{val / 1_000_000:.1f}M"
    elif abs_val >= 1_000:
        return f"{val / 1_000:.1f}K"
    else:
        return f"{val:,.1f}"


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI: create_bar_chart
# ─────────────────────────────────────────────────────────────────────────────
def create_bar_chart(
    df_agg: pd.DataFrame,
    x_col: str,
    y_col: str,
    title: str = "Visualisasi Data",
    bar_color: str = "#4C72B0",
    max_categories: int = 20,
) -> bytes:
    """
    Membuat bar chart bergaya korporat minimalis dari DataFrame teragregasi.

    Fitur:
    - Warna batang dengan gradient opacity berdasarkan nilai (highlight tertinggi)
    - Label nilai di atas setiap batang
    - Gridline horizontal yang halus
    - Rotasi label X otomatis jika banyak kategori
    - Output sebagai bytes (PNG) di buffer memori

    Parameter:
    ----------
    df_agg : pd.DataFrame
        DataFrame teragregasi dengan kolom x_col dan y_col.
    x_col : str
        Nama kolom untuk sumbu X (kategori/label).
    y_col : str
        Nama kolom untuk sumbu Y (nilai numerik).
    title : str, optional
        Judul grafik. Default: "Visualisasi Data".
    bar_color : str, optional
        Warna dasar batang dalam format hex. Default: "#4C72B0".
    max_categories : int, optional
        Jumlah maksimum kategori yang ditampilkan. Default: 20.

    Returns:
    --------
    bytes
        Gambar PNG dalam format bytes (dari BytesIO buffer).
    """
    style = CORPORATE_STYLE

    # Batasi jumlah kategori agar grafik tetap terbaca
    df_plot = df_agg.head(max_categories).copy()

    # Urutkan dari terbesar ke terkecil untuk visual yang lebih baik
    df_plot = df_plot.sort_values(y_col, ascending=False).reset_index(drop=True)

    n = len(df_plot)
    labels = df_plot[x_col].astype(str).tolist()
    values = df_plot[y_col].tolist()

    # Hitung alpha (opacity) berdasarkan posisi relatif — highlight bar tertinggi
    max_val = max(values) if values else 1
    alphas = [0.5 + 0.5 * (v / max_val) for v in values]

    # ── Setup figure ──────────────────────────────────────────────────────────
    fig_width = max(10, min(n * 0.9 + 2, 20))  # Lebar dinamis
    fig, ax = plt.subplots(figsize=(fig_width, 5.5), dpi=style["figure_dpi"])
    fig.patch.set_facecolor(style["bg_color"])
    ax.set_facecolor(style["bg_color"])

    # ── Gambar batang ─────────────────────────────────────────────────────────
    x_positions = np.arange(n)
    bars = ax.bar(
        x_positions,
        values,
        color=bar_color,
        alpha=style["bar_alpha"],
        edgecolor=style["bar_edge_color"],
        linewidth=0.8,
        width=0.65,
        zorder=3,
    )

    # Terapkan alpha individual per batang untuk efek highlight
    for bar, alpha in zip(bars, alphas):
        bar.set_alpha(alpha)

    # Tambahkan highlight warna lebih terang untuk batang tertinggi
    bars[0].set_edgecolor(bar_color)
    bars[0].set_linewidth(2)

    # ── Label nilai di atas batang ────────────────────────────────────────────
    y_range = max(values) - min(values) if len(values) > 1 else max(values)
    offset = y_range * 0.02 + max(values) * 0.01

    for i, (bar, val) in enumerate(zip(bars, values)):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + offset,
            _format_value(val),
            ha="center",
            va="bottom",
            fontsize=8.5,
            fontweight="600",
            color=style["value_label_color"],
            zorder=4,
        )

    # ── Styling sumbu & grid ──────────────────────────────────────────────────
    ax.set_xticks(x_positions)
    x_rot = 35 if n > 8 else 0
    x_ha = "right" if n > 8 else "center"
    ax.set_xticklabels(
        labels,
        rotation=x_rot,
        ha=x_ha,
        fontsize=9,
        color=style["subtitle_color"],
    )

    ax.yaxis.set_major_formatter(
        mticker.FuncFormatter(lambda x, _: _format_value(x))
    )
    ax.tick_params(axis="y", labelsize=9, colors=style["subtitle_color"])

    # Gridline horizontal halus
    ax.set_axisbelow(True)
    ax.yaxis.grid(True, color=style["grid_color"], linewidth=0.8, linestyle="--")
    ax.xaxis.grid(False)

    # Hilangkan border (spine) yang tidak perlu
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    ax.spines["bottom"].set_color(style["grid_color"])

    # ── Judul & label ─────────────────────────────────────────────────────────
    ax.set_title(
        title,
        fontsize=14,
        fontweight="700",
        color=style["title_color"],
        pad=16,
        loc="left",
    )
    ax.set_xlabel(
        x_col.replace("_", " ").title(),
        fontsize=10,
        color=style["subtitle_color"],
        labelpad=8,
    )
    ax.set_ylabel(
        y_col.replace("_", " ").title(),
        fontsize=10,
        color=style["subtitle_color"],
        labelpad=8,
    )

    # Tambahkan padding atas agar label nilai tidak terpotong
    current_top = ax.get_ylim()[1]
    ax.set_ylim(0, current_top * 1.15)

    # ── Watermark / catatan sumber ─────────────────────────────────────────────
    fig.text(
        0.99, 0.01,
        "Generated by Report Automator",
        ha="right", va="bottom",
        fontsize=7,
        color="#AAAAAA",
        style="italic",
    )

    plt.tight_layout(pad=1.5)

    # ── Simpan ke buffer BytesIO ──────────────────────────────────────────────
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", facecolor=style["bg_color"])
    plt.close(fig)
    buf.seek(0)

    return buf.read()
