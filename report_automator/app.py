"""
app.py - Entry point aplikasi Streamlit untuk Automasi Laporan
=================================================================
Menggabungkan semua modul: data processing, chart generation, 
dan PPT generation ke dalam satu UI yang interaktif.
"""

import streamlit as st
import pandas as pd
import io
import json
from datetime import datetime

from data_processor import (
    load_data, aggregate_data, generate_summary_bullets,
    get_excel_info, load_excel_sheet,
)
from chart_generator import create_bar_chart
from ppt_generator import generate_pptx


# ─────────────────────────────────────────────
# Konfigurasi halaman Streamlit
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="📊 Report Automator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# Custom CSS untuk tampilan yang lebih premium
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Header utama */
    .main-header {
        background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
        color: white;
        padding: 2rem 2.5rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(0,0,0,0.3);
    }

    .main-header h1 {
        font-size: 2.2rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.5px;
    }

    .main-header p {
        margin: 0.5rem 0 0 0;
        opacity: 0.75;
        font-size: 1rem;
    }

    /* Card metric */
    .metric-card {
        background: #1a1a2e;
        border: 1px solid #16213e;
        border-radius: 12px;
        padding: 1.2rem 1.5rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }

    /* Badge status */
    .status-badge {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.78rem;
        font-weight: 600;
        background: linear-gradient(135deg, #00b09b, #96c93d);
        color: white;
    }

    /* Tombol download kustom */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #667eea, #764ba2);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.7rem 2rem;
        font-weight: 600;
        font-size: 1rem;
        width: 100%;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
    }

    .stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6);
    }

    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background: #0d1117;
    }

    section[data-testid="stSidebar"] * {
        color: #e6edf3 !important;
    }

    /* Divider */
    .section-divider {
        border: none;
        height: 1px;
        background: linear-gradient(to right, transparent, #333, transparent);
        margin: 1.5rem 0;
    }

    /* Step indicator */
    .step-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        background: linear-gradient(135deg, #667eea, #764ba2);
        color: white;
        font-weight: 700;
        font-size: 0.85rem;
        margin-right: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Header Utama
# ─────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 Report Automator</h1>
    <p>Upload data CSV → Generate Bar Chart → Export ke PowerPoint secara otomatis</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# Sidebar: Konfigurasi Laporan
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Konfigurasi Laporan")
    st.markdown("---")

    report_title = st.text_input(
        "📝 Judul Laporan",
        value="Laporan Analisis Data Bulanan",
        help="Judul ini akan muncul di Slide 1 presentasi."
    )

    report_subtitle = st.text_input(
        "📌 Sub-judul / Departemen",
        value="Divisi Business Intelligence",
        help="Sub-judul yang muncul di bawah judul utama."
    )

    author_name = st.text_input(
        "👤 Nama Pembuat",
        value="Tim Analitik",
        help="Nama yang akan tercantum di slide."
    )

    report_date = st.date_input(
        "📅 Tanggal Laporan",
        value=datetime.today(),
    )

    st.markdown("---")
    st.markdown("### 🎨 Opsi Grafik")

    chart_color = st.color_picker(
        "Warna Batang Chart",
        value="#4C72B0",
        help="Pilih warna untuk bar chart."
    )

    x_col = st.text_input(
        "Kolom Sumbu X (Kategori)",
        value="",
        placeholder="Kosongkan untuk auto-detect",
        help="Nama kolom yang akan menjadi label sumbu X."
    )

    y_col = st.text_input(
        "Kolom Sumbu Y (Nilai)",
        value="",
        placeholder="Kosongkan untuk auto-detect",
        help="Nama kolom numerik yang akan diplot."
    )

    st.markdown("---")
    st.markdown(
        "<small>🔒 Data hanya diproses di sesi lokal Anda dan tidak dikirim ke server manapun.</small>",
        unsafe_allow_html=True
    )


# ─────────────────────────────────────────────
# Area utama: 3 kolom untuk step indicator
# ─────────────────────────────────────────────
col1, col2, col3 = st.columns(3)
with col1:
    st.info("**Step 1** · Upload CSV atau Excel")
with col2:
    st.info("**Step 2** · Preview Data & Grafik")
with col3:
    st.info("**Step 3** · Download File .pptx")


st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# STEP 1: File Uploader
# ─────────────────────────────────────────────
st.markdown("### 📂 Upload File Data")

uploaded_file = st.file_uploader(
    label="Seret & lepas file di sini, atau klik untuk memilih",
    type=["csv", "xlsx", "xls"],
    help="Format yang didukung: CSV (.csv) atau Excel (.xlsx/.xls). Ukuran maksimal 200MB.",
    label_visibility="visible"
)


# ─────────────────────────────────────────────
# Proses data setelah file diupload
# ─────────────────────────────────────────────
if uploaded_file is not None:

    filename_lower = uploaded_file.name.lower()
    is_excel = filename_lower.endswith((".xlsx", ".xls"))

    # ── Baca & load data ───────────────────────────────────────────────
    if is_excel:
        file_bytes = uploaded_file.getvalue()

        # Ekstrak info workbook (sheet names, chart detection)
        try:
            excel_info = get_excel_info(file_bytes)
        except Exception as e:
            st.error(f"❌ Gagal membaca file Excel: {e}")
            st.stop()

        # ── Sheet Selector (tampilkan hanya jika lebih dari 1 sheet) ──
        if len(excel_info["sheet_names"]) > 1:
            st.markdown("### 📊 Pilih Sheet Data")
            selected_sheet = st.selectbox(
                "Workbook ini memiliki beberapa sheet. Pilih sheet yang berisi data:",
                options=excel_info["sheet_names"],
                index=excel_info["sheet_names"].index(excel_info["recommended_sheet"]),
                help=(
                    f"Sheet yang direkomendasikan: ‘{excel_info['recommended_sheet']}’ "
                    f"(sheet pertama yang terdeteksi memiliki data)"
                ),
            )
        else:
            selected_sheet = excel_info["sheet_names"][0]

        # ── Banner jika workbook memiliki chart ───────────────────────────
        if excel_info["has_charts"]:
            st.info(
                "💡 **Excel Anda memiliki chart!**  \n\n"
                "Chart Excel tidak dapat diekstrak secara otomatis karena "
                "tersimpan sebagai data vektor, bukan gambar.  \n"
                "**Solusi:** Screenshot chart dari Excel Anda → simpan sebagai PNG/JPG → "
                "upload di bagian **📎 Lampiran Gambar** di bawah.  \n"
                "Gambar tersebut akan ditambahkan sebagai slide khusus di dalam PPTX."
            )

        # ── Load sheet yang dipilih ───────────────────────────────────
        try:
            df_raw = load_excel_sheet(file_bytes, selected_sheet)
        except Exception as e:
            st.error(f"❌ Gagal membaca sheet ‘{selected_sheet}’: {e}")
            st.stop()
    else:
        # ── CSV path (alur yang sudah ada) ─────────────────────────────
        try:
            df_raw = load_data(uploaded_file)
        except Exception as e:
            st.error(f"❌ Gagal membaca file: {e}")
            st.stop()

    st.success(f"✅ File **{uploaded_file.name}** berhasil dimuat — "
               f"**{df_raw.shape[0]:,} baris** × **{df_raw.shape[1]} kolom**")

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ── Preview data (tabel penuh) ────────────────────────────
    st.markdown("### 🔍 Data Mentah Lengkap")
    with st.expander(
        f"Tampilkan seluruh data ({df_raw.shape[0]:,} baris × {df_raw.shape[1]} kolom)",
        expanded=True
    ):
        st.dataframe(
            df_raw,
            use_container_width=True,
            height=400,
        )

    # ── Deteksi kolom otomatis ────────────────────────────────
    # Gunakan input user jika tersedia, jika tidak auto-detect
    x_col_resolved = x_col.strip() if x_col.strip() else None
    y_col_resolved = y_col.strip() if y_col.strip() else None

    # ── Agregasi data ─────────────────────────────────────────
    try:
        df_agg, x_col_final, y_col_final, agg_label = aggregate_data(
            df_raw, x_col_resolved, y_col_resolved
        )
    except Exception as e:
        st.error(f"❌ Gagal mengagregas data: {e}")
        st.stop()

    # ── STEP 2: Tampilkan agregasi & metrik ──────────────────
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown("### 📊 Hasil Agregasi & Statistik")

    # Metric cards
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.metric("Total Baris", f"{df_raw.shape[0]:,}")
    with m2:
        st.metric(f"Total {y_col_final}", f"{df_agg[y_col_final].sum():,.2f}")
    with m3:
        st.metric(f"Rata-rata {y_col_final}", f"{df_agg[y_col_final].mean():,.2f}")
    with m4:
        st.metric(f"Nilai Tertinggi", f"{df_agg[y_col_final].max():,.2f}")

    # Tampilkan tabel agregasi (seluruh baris)
    with st.expander(
        f"Tampilkan tabel agregasi ({len(df_agg):,} baris)",
        expanded=True
    ):
        st.dataframe(df_agg, use_container_width=True, height=350)

    # ── Generate bar chart ────────────────────────────────────
    chart_buf = create_bar_chart(
        df_agg=df_agg,
        x_col=x_col_final,
        y_col=y_col_final,
        title=f"{agg_label} per {x_col_final}",
        bar_color=chart_color,
    )

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown("### 📈 Visualisasi Bar Chart")
    st.image(chart_buf, use_container_width=True, caption=f"{agg_label} per {x_col_final}")

    # ── Generate bullet points summary ───────────────────────────────
    bullets = generate_summary_bullets(
        df_raw=df_raw,
        df_agg=df_agg,
        x_col=x_col_final,
        y_col=y_col_final,
        agg_label=agg_label,
        report_title=report_title,
    )

    # ── LAMPIRAN GAMBAR (Opsional) ────────────────────────────────────
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown("### 📎 Lampiran Gambar (Opsional — Maks. 5 Gambar)")
    st.markdown(
        "<small>Upload gambar atau screenshot chart Excel Anda. "
        "Setiap gambar akan menjadi <strong>slide lampiran tersendiri</strong> "
        "sebelum slide Penutup di dalam PPTX.</small>",
        unsafe_allow_html=True,
    )

    uploaded_images = st.file_uploader(
        "Upload gambar (PNG / JPG / JPEG)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
        key="user_image_uploader",
        help="Maks. 5 gambar. Cocok untuk screenshot chart Excel, infografis, atau visual pendukung lainnya.",
    )

    user_images = []
    if uploaded_images:
        images_to_use = uploaded_images[:5]
        if len(uploaded_images) > 5:
            st.warning(
                f"⚠️ Hanya 5 gambar pertama yang akan digunakan "
                f"(Anda mengupload {len(uploaded_images)} gambar)."
            )

        st.markdown(
            f"**{len(images_to_use)} gambar** siap digunakan — "
            "berikan **judul** dan **keterangan singkat** untuk setiap gambar:"
        )
        for i, img_file in enumerate(images_to_use):
            img_bytes = img_file.getvalue()
            st.markdown(f"---\n**Gambar {i + 1}** — `{img_file.name}`")
            col_prev, col_form = st.columns([1, 2])
            with col_prev:
                st.image(img_bytes, use_container_width=True)
                side = "🖼️ Gambar kiri · 📝 Teks kanan" if i % 2 == 0 else "📝 Teks kiri · 🖼️ Gambar kanan"
                st.caption(f"🔄 Layout di slide: {side}")
            with col_form:
                default_cap = (
                    img_file.name.rsplit(".", 1)[0]
                    .replace("_", " ")
                    .replace("-", " ")
                )
                caption = st.text_input(
                    f"📝 Judul Gambar {i + 1}",
                    value=default_cap,
                    key=f"img_caption_{i}",
                    placeholder="Contoh: Chart Penjualan Q1 2024",
                )
                description = st.text_area(
                    f"📋 Keterangan Singkat Gambar {i + 1}",
                    value="",
                    key=f"img_desc_{i}",
                    height=80,
                    placeholder="Contoh: Penjualan Q1 naik 15% YoY, tertinggi di cluster Jabodetabek.",
                )
                st.caption(
                    "📌 Judul tampil **bold** · Keterangan singkat (1–2 kalimat) muncul di panel merah."
                )
            user_images.append({
                "bytes": img_bytes,
                "caption": caption,
                "description": description,
            })

    # ── STEP 3: Generate & Download PPTX ─────────────────────
    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)
    st.markdown("### 🚀 Generate & Download Laporan PowerPoint")

    # Tombol generate
    if st.button("⚡ Generate File .pptx Sekarang", type="primary", use_container_width=True):

        with st.spinner("Sedang membuat presentasi PowerPoint..."):
            try:
                pptx_buf = generate_pptx(
                    title=report_title,
                    subtitle=report_subtitle,
                    author=author_name,
                    report_date=report_date.strftime("%d %B %Y"),
                    bullets=bullets,
                    chart_buf=chart_buf,
                    chart_title=f"{agg_label} per {x_col_final}",
                    analysis_text=(
                        f"Grafik menampilkan {agg_label.lower()} dari kolom "
                        f"'{y_col_final}' berdasarkan '{x_col_final}'.\n\n"
                        f"Nilai tertinggi: "
                        f"{df_agg[y_col_final].max():,.2f} "
                        f"({df_agg.loc[df_agg[y_col_final].idxmax(), x_col_final]}).\n\n"
                        f"Nilai terendah: "
                        f"{df_agg[y_col_final].min():,.2f} "
                        f"({df_agg.loc[df_agg[y_col_final].idxmin(), x_col_final]}).\n\n"
                        f"Rata-rata keseluruhan: "
                        f"{df_agg[y_col_final].mean():,.2f}.\n\n"
                        f"Total kategori: {len(df_agg)} · Total baris data: {df_raw.shape[0]:,}.\n\n"
                        f"Tren ini dapat digunakan sebagai acuan perencanaan dan pengambilan keputusan strategis operasional jaringan Telkomsel."
                    ),
                    df_raw=df_raw,
                    x_col=x_col_final,
                    y_col=y_col_final,
                    user_images=user_images if user_images else None,
                )
                st.session_state["pptx_buf"] = pptx_buf
                st.session_state["filename"] = (
                    f"Laporan_{report_title.replace(' ', '_')}_{report_date.strftime('%Y%m%d')}.pptx"
                )
                st.success("✅ File PowerPoint berhasil dibuat! Klik tombol Download di bawah.")
            except Exception as e:
                st.error(f"❌ Gagal membuat PPTX: {e}")

    # Tampilkan tombol download jika file sudah digenerate
    if "pptx_buf" in st.session_state:
        st.download_button(
            label="⬇️ Download File .pptx",
            data=st.session_state["pptx_buf"],
            file_name=st.session_state["filename"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

else:
    # ── Placeholder saat belum ada file ──────────────────────
    st.markdown("""
    <div style="
        text-align: center;
        padding: 4rem 2rem;
        background: linear-gradient(135deg, #0d1117, #161b22);
        border: 1px dashed #30363d;
        border-radius: 16px;
        color: #8b949e;
        margin-top: 1rem;
    ">
        <div style="font-size: 4rem; margin-bottom: 1rem;">📁</div>
        <h3 style="color: #e6edf3; margin-bottom: 0.5rem;">Belum ada file yang diupload</h3>
        <p>Upload file <strong>CSV</strong> atau <strong>Excel (.xlsx / .xls)</strong> Anda di atas untuk memulai.</p>
        <p style="font-size:0.85rem; margin-top:1rem;">
            💡 <em>Contoh: data penjualan bulanan, laporan keuangan, atau data operasional lainnya.</em>
        </p>
    </div>
    """, unsafe_allow_html=True)

    # Tampilkan contoh format data yang didukung
    with st.expander("📖 Lihat contoh format data yang didukung"):
        st.markdown("**Format CSV:**")
        st.code("""bulan,penjualan,target
Januari,120000,100000
Februari,98000,100000
Maret,145000,120000
April,132000,120000""", language="csv")
        st.markdown("**Format Excel (.xlsx):**")
        st.info(
            "Buka Excel → isi data dengan baris pertama sebagai header kolom → simpan sebagai .xlsx.  \n"
            "Jika memiliki beberapa sheet, Anda bisa memilih sheet mana yang akan diproses."
        )
        st.info("💡 Pastikan baris pertama adalah **header/nama kolom**, dan minimal ada satu kolom berisi angka.")
