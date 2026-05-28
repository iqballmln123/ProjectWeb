"""
app.py - Entry point aplikasi Streamlit untuk Automasi Laporan Telkomsel
=========================================================================
Template: Coverage Activity Report
Struktur:
  Slide 1       : Cover
  Slide 2..N    : Satu slide per baris data (Site)
  Slide terakhir: Penutup
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime

from data_processor import (
    get_excel_info, load_excel_sheet, load_data,
)
from ppt_generator import (
    generate_pptx,
    COL_SITE_ID, COL_SITE_NAME, COL_PURPOSE, COL_CITY,
    COL_FINDING, COL_PLAN_ACTION, COL_SUPPORT, COL_GOALS, COL_INCREMENT,
)


# ─────────────────────────────────────────────────────────────────────────────
# Konfigurasi halaman
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title=" Telkomsel Report Automator",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .main-header {
        background: linear-gradient(135deg, #8B0000, #CC0000, #FF3333);
        color: white;
        padding: 1.8rem 2.5rem;
        border-radius: 16px;
        margin-bottom: 1.5rem;
        box-shadow: 0 8px 32px rgba(204,0,0,0.35);
    }
    .main-header h1 { font-size: 2rem; font-weight: 700; margin: 0; }
    .main-header p  { margin: 0.4rem 0 0 0; opacity: 0.85; font-size: 0.95rem; }

    .step-badge {
        background: linear-gradient(135deg, #CC0000, #FF3333);
        color: white;
        padding: 0.35rem 1rem;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.8rem;
        display: inline-block;
        margin-bottom: 0.5rem;
    }

    .site-card {
        background: #1a1a2e;
        border: 1px solid #CC0000;
        border-radius: 10px;
        padding: 1rem 1.2rem;
        margin-bottom: 0.8rem;
    }

    section[data-testid="stSidebar"] { background: #0d1117; }
    section[data-testid="stSidebar"] * { color: #e6edf3 !important; }

    .stDownloadButton > button {
        background: linear-gradient(135deg, #CC0000, #990000);
        color: white; border: none; border-radius: 10px;
        padding: 0.7rem 2rem; font-weight: 700; font-size: 1rem;
        width: 100%; transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(204,0,0,0.4);
    }
    .stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(204,0,0,0.6);
    }
    .section-divider {
        border: none; height: 1px;
        background: linear-gradient(to right, transparent, #CC0000, transparent);
        margin: 1.5rem 0;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📡 Telkomsel Coverage Report Automator</h1>
    <p>Upload data Excel → Preview per Site → Generate PowerPoint otomatis (Cover · Data per Site · Penutup)</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar: Konfigurasi Laporan
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("##  Konfigurasi Laporan")
    st.markdown("---")

    report_title = st.text_input(
        " Judul Laporan",
        value="Tracking Activity NOP 2026",
        help="Judul ini akan muncul di Slide Cover.",
    )
    report_subtitle = st.text_input(
        " Sub-judul / Departemen",
        value="Divisi Network Operation",
        help="Sub-judul di bawah judul utama.",
    )
    author_name = st.text_input(
        " Nama Pembuat",
        value="Tim Network Operation",
    )
    report_date = st.date_input(
        " Tanggal Laporan",
        value=datetime.today(),
    )

    st.markdown("---")
    st.markdown(
        "<small> Data hanya diproses lokal dan tidak dikirim ke server manapun.</small>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# Step indicator
# ─────────────────────────────────────────────────────────────────────────────
c1, c2, c3 = st.columns(3)
with c1: st.info("**Step 1** · Upload Excel Data")
with c2: st.info("**Step 2** · Upload Gambar per Site")
with c3: st.info("**Step 3** · Generate & Download .pptx")

st.markdown('<hr class="section-divider">', unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# STEP 1: Upload File
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("###  Step 1 · Upload File Data")
st.markdown(
    "Upload file Excel dengan kolom: **NO · SITE ID · SITE NAME · FINDING · CITY · "
    "PURPOSE HEADER · SOW · PLAN ACTION · SUPPORT NEEDED · GOALS · INCREAMENT PAYLOAD AND REVENUE · LONGITUDE · LATITUDE**"
)

uploaded_file = st.file_uploader(
    "Seret & lepas file Excel di sini, atau klik untuk memilih",
    type=["csv", "xlsx", "xls"],
    help="Format: CSV (.csv) atau Excel (.xlsx/.xls). Setiap baris = 1 slide PPT.",
)

df_raw = None

if uploaded_file is not None:
    filename_lower = uploaded_file.name.lower()
    is_excel = filename_lower.endswith((".xlsx", ".xls"))

    if is_excel:
        file_bytes = uploaded_file.getvalue()
        try:
            excel_info = get_excel_info(file_bytes)
        except Exception as e:
            st.error(f"❌ Gagal membaca file Excel: {e}")
            st.stop()

        if len(excel_info["sheet_names"]) > 1:
            st.markdown("###  Pilih Sheet Data")
            selected_sheet = st.selectbox(
                "Workbook ini memiliki beberapa sheet. Pilih sheet yang berisi data:",
                options=excel_info["sheet_names"],
                index=excel_info["sheet_names"].index(excel_info["recommended_sheet"]),
            )
        else:
            selected_sheet = excel_info["sheet_names"][0]

        try:
            df_raw = load_excel_sheet(file_bytes, selected_sheet)
        except Exception as e:
            st.error(f"❌ Gagal membaca sheet '{selected_sheet}': {e}")
            st.stop()
    else:
        try:
            df_raw = load_data(uploaded_file)
        except Exception as e:
            st.error(f"❌ Gagal membaca file: {e}")
            st.stop()

    st.success(
        f"✅ **{uploaded_file.name}** berhasil dimuat — "
        f"**{df_raw.shape[0]:,} baris** (site) × **{df_raw.shape[1]} kolom**"
    )
    st.markdown(
        f" Akan menghasilkan **{df_raw.shape[0]} slide data** + 1 Cover + 1 Penutup "
        f"= **{df_raw.shape[0] + 2} slide** total."
    )

    # Preview tabel ringkas
    with st.expander(" Preview data (10 baris pertama)", expanded=False):
        # Konversi semua ke string agar aman untuk Arrow serialization
        df_preview = df_raw.head(10).astype(str)
        st.dataframe(df_preview, use_container_width=True)


    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ─────────────────────────────────────────────────────────────────────────
    # STEP 2: Upload Gambar per Site
    # ─────────────────────────────────────────────────────────────────────────
    st.markdown("###  Step 2 · Upload Gambar per Site")
    st.markdown(
        "Untuk setiap site, upload gambar pendukung. "
        "Setiap site membutuhkan **minimum 3 gambar** (Payload Chart, Maps/COVMO, Support Image), "
        "dan **maksimal 5 gambar** (Payload Chart, Maps/COVMO, + 3 Support Images)."
    )

    # Ambil daftar Site ID
    col_site_id = COL_SITE_ID
    if col_site_id not in df_raw.columns:
        candidates = [c for c in df_raw.columns if "site" in c.lower() and "id" in c.lower()]
        col_site_id = candidates[0] if candidates else df_raw.columns[0]
        st.warning(f"⚠️ Kolom '{COL_SITE_ID}' tidak ditemukan. Menggunakan kolom '{col_site_id}' sebagai Site ID.")

    # Sanitize site IDs: ganti NaN / kosong dengan index, pastikan unik
    raw_ids = df_raw[col_site_id].tolist()
    site_ids = []
    seen = {}
    for i, sid_raw in enumerate(raw_ids):
        # Pastikan selalu string, bahkan jika nilai float NaN
        sid = str(sid_raw).strip() if sid_raw is not None else ""
        # Ganti NaN, kosong, atau 'nan' dengan fallback berbasis index
        if not sid or sid.lower() in ("nan", "none", ""):
            sid = f"site_{i+1}"
        # Hindari duplikat key dengan menambahkan suffix
        if sid in seen:
            seen[sid] += 1
            sid = f"{sid}_{seen[sid]}"
        else:
            seen[sid] = 0
        site_ids.append(sid)


    # Dict untuk menyimpan gambar per site: {site_id: [{bytes, label}, ...]}
    site_images: dict = {}

    # Pilih apakah ingin upload gambar per site atau skip
    upload_mode = st.radio(
        "Mode upload gambar:",
        ["Upload gambar per site (direkomendasikan)", "Skip — generate tanpa gambar (placeholder)"],
        index=0,
        horizontal=True,
    )

    if upload_mode.startswith("Upload"):
        # Bisa expand tiap site atau pilih site tertentu
        st.markdown("---")
        selected_sites = st.multiselect(
            "Pilih site yang ingin diisi gambarnya (kosongkan untuk semua):",
            options=site_ids,
            default=[],
            help="Jika dikosongkan, form upload akan ditampilkan untuk semua site.",
        )
        sites_to_show = selected_sites if selected_sites else site_ids[:10]  # Batasi 10 pertama

        if not selected_sites and len(site_ids) > 10:
            st.info(
                f"⚠️ Menampilkan form upload untuk 10 site pertama dari {len(site_ids)} site. "
                "Gunakan filter di atas untuk memilih site tertentu."
            )

        for i, sid in enumerate(sites_to_show):
            # Ambil nama site (berdasarkan index asli di site_ids)
            orig_idx = site_ids.index(sid)
            site_name_col = COL_SITE_NAME
            if site_name_col in df_raw.columns:
                sname = str(df_raw.iloc[orig_idx][site_name_col]).strip()
                site_label = f"{sid} — {sname}" if sname else sid
            else:
                site_label = str(sid)

            with st.expander(f" {site_label}", expanded=False):
                col_a, col_b = st.columns([1, 1])

                with col_a:
                    st.markdown("**① Payload Site Surrounding** *(wajib)*")
                    img_payload = st.file_uploader(
                        "Upload chart payload",
                        type=["png", "jpg", "jpeg"],
                        key=f"payload_{sid}",
                    )

                    st.markdown("**② Maps Preview & COVMO** *(wajib)*")
                    img_maps = st.file_uploader(
                        "Upload maps/COVMO",
                        type=["png", "jpg", "jpeg"],
                        key=f"maps_{sid}",
                    )

                with col_b:
                    st.markdown("**③–⑤ Support Images** *(min 1, maks 3)*")
                    img_supports = st.file_uploader(
                        "Upload support images (1–3 gambar)",
                        type=["png", "jpg", "jpeg"],
                        accept_multiple_files=True,
                        key=f"support_{sid}",
                    )

                # Kumpulkan gambar site ini
                imgs_this_site = []
                if img_payload:
                    imgs_this_site.append({"bytes": img_payload.getvalue(), "label": "Payload Site"})
                if img_maps:
                    imgs_this_site.append({"bytes": img_maps.getvalue(), "label": "Maps & COVMO"})
                for si_img in (img_supports or [])[:3]:
                    imgs_this_site.append({"bytes": si_img.getvalue(), "label": "Support"})

                if imgs_this_site:
                    site_images[sid] = imgs_this_site
                    st.success(f"✅ {len(imgs_this_site)} gambar siap untuk site ini.")
                else:
                    st.caption("Belum ada gambar untuk site ini — akan menggunakan placeholder.")

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ─────────────────────────────────────────────────────────────────────────
    # STEP 3: Generate & Download
    # ─────────────────────────────────────────────────────────────────────────
    st.markdown("###  Step 3 · Generate & Download Laporan PowerPoint")

    if st.button(" Generate File .pptx Sekarang", type="primary", use_container_width=True):
        with st.spinner("Sedang membuat presentasi PowerPoint..."):
            try:
                pptx_buf = generate_pptx(
                    title=report_title,
                    subtitle=report_subtitle,
                    author=author_name,
                    report_date=report_date.strftime("%d %B %Y"),
                    df_raw=df_raw,
                    site_images=site_images if site_images else {},
                )
                st.session_state["pptx_buf"]  = pptx_buf
                st.session_state["pptx_fname"] = (
                    f"CoverageReport_{report_title.replace(' ', '_')}_"
                    f"{report_date.strftime('%Y%m%d')}.pptx"
                )
                st.success(
                    f"✅ File PowerPoint berhasil dibuat! "
                    f"Total **{df_raw.shape[0] + 2} slide**. Klik tombol Download di bawah."
                )
            except Exception as e:
                st.error(f"❌ Gagal membuat PPTX: {e}")
                import traceback
                st.code(traceback.format_exc())

    if "pptx_buf" in st.session_state:
        st.download_button(
            label="⬇ Download File .pptx",
            data=st.session_state["pptx_buf"],
            file_name=st.session_state["pptx_fname"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

else:
    # Placeholder belum ada file
    st.markdown("""
    <div style="
        text-align: center;
        padding: 4rem 2rem;
        background: linear-gradient(135deg, #0d1117, #161b22);
        border: 1px dashed #CC0000;
        border-radius: 16px;
        color: #8b949e;
        margin-top: 1rem;
    ">
        <div style="font-size: 4rem; margin-bottom: 1rem;">📡</div>
        <h3 style="color: #e6edf3; margin-bottom: 0.5rem;">Belum ada file yang diupload</h3>
        <p>Upload file <strong>Excel (.xlsx / .xls)</strong> atau <strong>CSV</strong> di atas untuk memulai.</p>
        <p style="font-size:0.85rem; margin-top:1rem;">
             <em>Kolom yang dibutuhkan: NO · SITE ID · SITE NAME · FINDING · CITY ·
            PURPOSE HEADER · SOW · PLAN ACTION · SUPPORT NEEDED · GOALS · INCREAMENT PAYLOAD AND REVENUE</em>
        </p>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(" Format Excel yang dibutuhkan"):
        st.markdown("**Struktur kolom:**")
        st.dataframe(pd.DataFrame([{
            "NO": "1",
            "SITE ID": "BDG212",
            "SITE NAME": "BDG212_SETRAMURNI-DMT",
            "FINDING": "Nearest Site BDG212_SETRAMURNI-DMT\nFrom BDG212 to Red Coverage descending contour",
            "CITY": "Bandung",
            "PURPOSE HEADER": "Install EM Under BDG212 SETRAMURNI",
            "SOW": "Purpose Easy macro 9 Meter -6.869313, 107.578078",
            "PLAN ACTION": "Propose Easy Macro 9 m -6.869313, 107.578078",
            "SUPPORT NEEDED": "Material AAU dual Band\nKabel Power + KWH\nTiang Telom 9 m\nInstallation & integration",
            "GOALS": "Improvement payload & Revenue area perumahan Setramurni\nImprovement Red Coverage Kec Sukasari",
            "INCREAMENT PAYLOAD AND REVENUE": "Payload: 300 GB (Monthly)\nRevneue: 30 Jt (Monthly)",
            "LONGITUDE": "107.578078",
            "LATITUDE": "-6.869313",
        }]), use_container_width=True)
        st.info(
            " Setiap baris = 1 site = 1 slide PPT.  \n"
            "Untuk field berisi beberapa poin, pisahkan dengan baris baru (Enter di Excel)."
        )
