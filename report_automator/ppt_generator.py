"""
ppt_generator.py - Modul pembuatan presentasi PowerPoint
==========================================================
Menggunakan python-pptx untuk membuat file .pptx dengan 3 slide:
  - Slide 1: Title Slide (Judul Laporan)
  - Slide 2: Executive Summary (Bullet Points ringkasan data)
  - Slide 3: Visualisasi (Grafik + Teks Analisis)

Semua dimensi menggunakan satuan EMU (English Metric Units).
1 inch = 914400 EMU | 1 cm = 360000 EMU
"""

import io
from typing import List
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt


# ─────────────────────────────────────────────────────────────────────────────
# KONSTANTA: Palet warna tema korporat
# ─────────────────────────────────────────────────────────────────────────────

# Warna utama — biru navy korporat
COLOR_PRIMARY = RGBColor(0x0F, 0x20, 0x27)       # #0F2027 (very dark navy)
COLOR_ACCENT = RGBColor(0x2C, 0x53, 0x64)         # #2C5364 (medium blue)
COLOR_ACCENT_LIGHT = RGBColor(0x20, 0x3A, 0x43)   # #203A43 (mid navy)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_LIGHT_GRAY = RGBColor(0xF4, 0xF6, 0xF9)
COLOR_TEXT_DARK = RGBColor(0x2C, 0x3E, 0x50)
COLOR_TEXT_GRAY = RGBColor(0x7F, 0x8C, 0x8D)
COLOR_BULLET_DOT = RGBColor(0x2C, 0x53, 0x64)


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI HELPER: _add_solid_fill
# ─────────────────────────────────────────────────────────────────────────────
def _add_solid_fill(shape, rgb_color: RGBColor):
    """
    Mengisi background sebuah shape dengan warna solid.

    Parameter:
    ----------
    shape : pptx.shapes.base.BaseShape
        Shape yang akan diisi warna.
    rgb_color : RGBColor
        Objek warna RGBColor dari python-pptx.
    """
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = rgb_color


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI HELPER: _set_text
# ─────────────────────────────────────────────────────────────────────────────
def _set_text(
    tf,
    text: str,
    font_size: int,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor = COLOR_WHITE,
    alignment=PP_ALIGN.LEFT,
):
    """
    Mengatur teks pada TextFrame secara langsung (mengganti semua teks yang ada).

    Parameter:
    ----------
    tf : pptx.text.text.TextFrame
        TextFrame dari shape yang akan diisi teks.
    text : str
        Teks yang akan ditampilkan.
    font_size : int
        Ukuran font dalam satuan poin (pt).
    bold : bool
        Apakah teks dicetak tebal.
    italic : bool
        Apakah teks dicetak miring.
    color : RGBColor
        Warna teks.
    alignment : PP_ALIGN
        Perataan teks (LEFT, CENTER, RIGHT, JUSTIFY).
    """
    tf.clear()
    para = tf.paragraphs[0]
    para.alignment = alignment
    run = para.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = "Calibri"


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI HELPER: _add_rectangle
# ─────────────────────────────────────────────────────────────────────────────
def _add_rectangle(slide, left, top, width, height, color: RGBColor):
    """
    Menambahkan persegi panjang (rectangle shape) ke slide sebagai elemen dekoratif.

    Parameter:
    ----------
    slide : pptx.slide.Slide
        Slide tujuan.
    left, top, width, height : float
        Posisi dan dimensi dalam satuan Inches.
    color : RGBColor
        Warna isian.

    Returns:
    --------
    shape
        Shape yang telah ditambahkan.
    """
    from pptx.util import Inches
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    _add_solid_fill(shape, color)
    shape.line.fill.background()  # Hapus border
    return shape


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI: _build_slide_1_title
# ─────────────────────────────────────────────────────────────────────────────
def _build_slide_1_title(
    prs: Presentation,
    title: str,
    subtitle: str,
    author: str,
    report_date: str,
):
    """
    Membangun Slide 1: Title Slide dengan desain korporat berlatar gelap.

    Layout:
    - Background gelap gradient (disimulasikan dengan rectangle penuh)
    - Aksen garis horizontal berwarna biru
    - Judul besar di tengah
    - Sub-judul, nama pembuat, dan tanggal di bawahnya

    Parameter:
    ----------
    prs : Presentation
        Objek Presentation python-pptx.
    title : str
        Judul laporan utama.
    subtitle : str
        Sub-judul atau nama departemen.
    author : str
        Nama pembuat laporan.
    report_date : str
        Tanggal laporan dalam string terformat.
    """
    # Gunakan layout blank (index 6) untuk kontrol penuh
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    slide_w = prs.slide_width
    slide_h = prs.slide_height

    # Background utama (dark navy penuh)
    bg = slide.shapes.add_shape(
        1, 0, 0, slide_w, slide_h
    )
    _add_solid_fill(bg, COLOR_PRIMARY)
    bg.line.fill.background()

    # Accent rectangle bawah (garis warna biru)
    accent_bar = slide.shapes.add_shape(
        1, 0, Inches(4.8), slide_w, Inches(0.08)
    )
    _add_solid_fill(accent_bar, COLOR_ACCENT)
    accent_bar.line.fill.background()

    # Kotak semi-transparent di area judul untuk depth
    mid_bg = slide.shapes.add_shape(
        1, Inches(0.5), Inches(1.5), Inches(8.8), Inches(2.8)
    )
    _add_solid_fill(mid_bg, COLOR_ACCENT_LIGHT)
    mid_bg.line.fill.background()

    # ── Judul utama ────────────────────────────────────────────────────────────
    txb_title = slide.shapes.add_textbox(
        Inches(0.7), Inches(1.7), Inches(8.5), Inches(1.5)
    )
    _set_text(
        txb_title.text_frame, title,
        font_size=32, bold=True, color=COLOR_WHITE,
        alignment=PP_ALIGN.LEFT,
    )
    txb_title.text_frame.word_wrap = True

    # ── Sub-judul ──────────────────────────────────────────────────────────────
    txb_sub = slide.shapes.add_textbox(
        Inches(0.7), Inches(3.0), Inches(8.5), Inches(0.6)
    )
    _set_text(
        txb_sub.text_frame, subtitle,
        font_size=16, bold=False, italic=True,
        color=RGBColor(0xA0, 0xC4, 0xD8),
        alignment=PP_ALIGN.LEFT,
    )

    # ── Penulis & Tanggal ──────────────────────────────────────────────────────
    txb_author = slide.shapes.add_textbox(
        Inches(0.7), Inches(5.1), Inches(5), Inches(0.4)
    )
    _set_text(
        txb_author.text_frame, f"Disusun oleh: {author}",
        font_size=11, color=RGBColor(0xBD, 0xD5, 0xE2),
        alignment=PP_ALIGN.LEFT,
    )

    txb_date = slide.shapes.add_textbox(
        Inches(5.5), Inches(5.1), Inches(4), Inches(0.4)
    )
    _set_text(
        txb_date.text_frame, report_date,
        font_size=11, color=RGBColor(0xBD, 0xD5, 0xE2),
        alignment=PP_ALIGN.RIGHT,
    )

    # Label "LAPORAN RESMI" di pojok kiri atas
    txb_badge = slide.shapes.add_textbox(
        Inches(0.7), Inches(0.8), Inches(3), Inches(0.35)
    )
    _set_text(
        txb_badge.text_frame, "📊  LAPORAN RESMI",
        font_size=9, color=RGBColor(0x96, 0xC9, 0x3D),
        alignment=PP_ALIGN.LEFT,
    )


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI: _build_slide_2_summary
# ─────────────────────────────────────────────────────────────────────────────
def _build_slide_2_summary(
    prs: Presentation,
    bullets: List[str],
):
    """
    Membangun Slide 2: Executive Summary dengan daftar bullet points.

    Layout:
    - Header bar biru navy di atas
    - Judul "Executive Summary"
    - Bullet points dirender manual dengan ikon bullet kustom

    Parameter:
    ----------
    prs : Presentation
        Objek Presentation python-pptx.
    bullets : List[str]
        Daftar string bullet points untuk ditampilkan.
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    slide_w = prs.slide_width
    slide_h = prs.slide_height

    # Background putih bersih
    bg = slide.shapes.add_shape(1, 0, 0, slide_w, slide_h)
    _add_solid_fill(bg, COLOR_WHITE)
    bg.line.fill.background()

    # Header bar atas (navy)
    header = slide.shapes.add_shape(
        1, 0, 0, slide_w, Inches(1.1)
    )
    _add_solid_fill(header, COLOR_PRIMARY)
    header.line.fill.background()

    # Accent stripe bawah header
    accent = slide.shapes.add_shape(
        1, 0, Inches(1.1), slide_w, Inches(0.06)
    )
    _add_solid_fill(accent, COLOR_ACCENT)
    accent.line.fill.background()

    # Judul di header
    txb_title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2), Inches(9), Inches(0.7)
    )
    _set_text(
        txb_title.text_frame, "Executive Summary",
        font_size=22, bold=True, color=COLOR_WHITE,
        alignment=PP_ALIGN.LEFT,
    )

    # Slide number badge di header kanan
    txb_num = slide.shapes.add_textbox(
        Inches(8.8), Inches(0.3), Inches(1), Inches(0.5)
    )
    _set_text(
        txb_num.text_frame, "02",
        font_size=18, bold=True,
        color=RGBColor(0x50, 0x80, 0x9A),
        alignment=PP_ALIGN.RIGHT,
    )

    # Render bullet points — tiap bullet dalam textbox terpisah untuk kontrol penuh
    y_start = 1.35
    bullet_spacing = 0.68
    bullet_symbols = ["▶", "▶", "▶", "▶", "▶", "▶", "▶", "▶"]

    for i, bullet_text in enumerate(bullets[:7]):  # Maksimal 7 bullet
        y_pos = y_start + i * bullet_spacing

        # Ikon bullet (warna biru)
        txb_icon = slide.shapes.add_textbox(
            Inches(0.4), Inches(y_pos), Inches(0.3), Inches(0.5)
        )
        _set_text(
            txb_icon.text_frame,
            bullet_symbols[i % len(bullet_symbols)],
            font_size=9, bold=True,
            color=COLOR_ACCENT,
            alignment=PP_ALIGN.CENTER,
        )

        # Teks bullet
        txb_bullet = slide.shapes.add_textbox(
            Inches(0.8), Inches(y_pos), Inches(8.8), Inches(0.55)
        )
        tf = txb_bullet.text_frame
        tf.word_wrap = True
        _set_text(
            tf, bullet_text,
            font_size=11, color=COLOR_TEXT_DARK,
            alignment=PP_ALIGN.LEFT,
        )

    # Footer line
    footer_line = slide.shapes.add_shape(
        1, Inches(0.5), Inches(6.95), Inches(9), Inches(0.02)
    )
    _add_solid_fill(footer_line, COLOR_ACCENT)
    footer_line.line.fill.background()


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI: _build_slide_3_visualization
# ─────────────────────────────────────────────────────────────────────────────
def _build_slide_3_visualization(
    prs: Presentation,
    chart_buf: bytes,
    chart_title: str,
    analysis_text: str,
):
    """
    Membangun Slide 3: Visualisasi Data.

    Layout dua kolom:
    - Kolom kiri (60%): Gambar grafik bar chart
    - Kolom kanan (40%): Teks analisis dan insight

    Parameter:
    ----------
    prs : Presentation
        Objek Presentation python-pptx.
    chart_buf : bytes
        Gambar bar chart dalam format bytes PNG.
    chart_title : str
        Judul grafik yang akan ditampilkan di atas gambar.
    analysis_text : str
        Teks narasi analisis yang ditampilkan di samping grafik.
    """
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    slide_w = prs.slide_width
    slide_h = prs.slide_height

    # Background putih
    bg = slide.shapes.add_shape(1, 0, 0, slide_w, slide_h)
    _add_solid_fill(bg, COLOR_WHITE)
    bg.line.fill.background()

    # Header bar (navy)
    header = slide.shapes.add_shape(
        1, 0, 0, slide_w, Inches(1.1)
    )
    _add_solid_fill(header, COLOR_PRIMARY)
    header.line.fill.background()

    # Accent stripe bawah header
    accent = slide.shapes.add_shape(
        1, 0, Inches(1.1), slide_w, Inches(0.06)
    )
    _add_solid_fill(accent, COLOR_ACCENT)
    accent.line.fill.background()

    # Judul di header
    txb_title = slide.shapes.add_textbox(
        Inches(0.5), Inches(0.2), Inches(9), Inches(0.7)
    )
    _set_text(
        txb_title.text_frame, "Visualisasi & Analisis Data",
        font_size=22, bold=True, color=COLOR_WHITE,
        alignment=PP_ALIGN.LEFT,
    )

    # Slide number badge
    txb_num = slide.shapes.add_textbox(
        Inches(8.8), Inches(0.3), Inches(1), Inches(0.5)
    )
    _set_text(
        txb_num.text_frame, "03",
        font_size=18, bold=True,
        color=RGBColor(0x50, 0x80, 0x9A),
        alignment=PP_ALIGN.RIGHT,
    )

    # Sub-judul chart
    txb_chart_title = slide.shapes.add_textbox(
        Inches(0.3), Inches(1.25), Inches(6), Inches(0.45)
    )
    _set_text(
        txb_chart_title.text_frame, chart_title,
        font_size=11, bold=True,
        color=COLOR_ACCENT,
        alignment=PP_ALIGN.LEFT,
    )

    # ── Masukkan gambar grafik (kolom kiri, 60% lebar slide) ──────────────────
    chart_stream = io.BytesIO(chart_buf)
    chart_left = Inches(0.3)
    chart_top = Inches(1.7)
    chart_width = Inches(6.0)   # 60% dari lebar slide (10 inch)
    chart_height = Inches(4.8)

    slide.shapes.add_picture(
        chart_stream,
        chart_left, chart_top,
        chart_width, chart_height,
    )

    # ── Panel analisis kanan ───────────────────────────────────────────────────
    # Background panel abu-abu muda
    analysis_panel = slide.shapes.add_shape(
        1,
        Inches(6.5), Inches(1.25),
        Inches(3.2), Inches(5.2)
    )
    _add_solid_fill(analysis_panel, COLOR_LIGHT_GRAY)
    analysis_panel.line.color.rgb = RGBColor(0xD0, 0xD8, 0xE4)

    # Label "ANALISIS" di panel
    txb_analysis_label = slide.shapes.add_textbox(
        Inches(6.6), Inches(1.35), Inches(3.0), Inches(0.4)
    )
    _set_text(
        txb_analysis_label.text_frame, "📋  ANALISIS",
        font_size=10, bold=True,
        color=COLOR_ACCENT,
        alignment=PP_ALIGN.LEFT,
    )

    # Garis pemisah di bawah label analisis
    divider = slide.shapes.add_shape(
        1, Inches(6.6), Inches(1.75), Inches(2.8), Inches(0.02)
    )
    _add_solid_fill(divider, COLOR_ACCENT)
    divider.line.fill.background()

    # Teks analisis utama
    txb_analysis = slide.shapes.add_textbox(
        Inches(6.6), Inches(1.85), Inches(3.0), Inches(4.0)
    )
    tf_analysis = txb_analysis.text_frame
    tf_analysis.word_wrap = True
    _set_text(
        tf_analysis, analysis_text,
        font_size=10, color=COLOR_TEXT_DARK,
        alignment=PP_ALIGN.LEFT,
    )

    # Footer line
    footer_line = slide.shapes.add_shape(
        1, Inches(0.3), Inches(6.95), Inches(9.4), Inches(0.02)
    )
    _add_solid_fill(footer_line, COLOR_ACCENT)
    footer_line.line.fill.background()


# ─────────────────────────────────────────────────────────────────────────────
# FUNGSI UTAMA: generate_pptx
# ─────────────────────────────────────────────────────────────────────────────
def generate_pptx(
    title: str,
    subtitle: str,
    author: str,
    report_date: str,
    bullets: List[str],
    chart_buf: bytes,
    chart_title: str,
    analysis_text: str,
) -> bytes:
    """
    Fungsi utama yang mengkoordinasikan pembuatan seluruh presentasi PowerPoint.

    Membuat Presentation dengan 3 slide, lalu mengembalikannya sebagai bytes
    yang siap didownload via Streamlit.

    Parameter:
    ----------
    title : str
        Judul laporan utama.
    subtitle : str
        Sub-judul / nama departemen.
    author : str
        Nama pembuat laporan.
    report_date : str
        Tanggal laporan (string terformat).
    bullets : List[str]
        Daftar bullet points untuk Executive Summary.
    chart_buf : bytes
        Gambar bar chart dalam format bytes PNG.
    chart_title : str
        Judul grafik untuk slide 3.
    analysis_text : str
        Teks narasi analisis untuk slide 3.

    Returns:
    --------
    bytes
        File .pptx dalam format bytes, siap untuk didownload.
    """
    # Inisialisasi Presentation dengan rasio 16:9
    prs = Presentation()
    prs.slide_width = Inches(10)    # 25.4 cm
    prs.slide_height = Inches(7.5)  # 19.05 cm (standard 4:3) — lebih lebar
    # Untuk 16:9 gunakan Inches(10) x Inches(5.625)
    # Di sini pakai 10x7.5 agar teks lebih lega
    prs.slide_height = Inches(7.5)

    # ── Bangun setiap slide ────────────────────────────────────────────────────
    _build_slide_1_title(
        prs=prs,
        title=title,
        subtitle=subtitle,
        author=author,
        report_date=report_date,
    )

    _build_slide_2_summary(
        prs=prs,
        bullets=bullets,
    )

    _build_slide_3_visualization(
        prs=prs,
        chart_buf=chart_buf,
        chart_title=chart_title,
        analysis_text=analysis_text,
    )

    # ── Simpan ke buffer BytesIO ───────────────────────────────────────────────
    pptx_buf = io.BytesIO()
    prs.save(pptx_buf)
    pptx_buf.seek(0)

    return pptx_buf.read()
