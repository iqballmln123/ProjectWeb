"""
ppt_generator.py - Modul pembuatan presentasi PowerPoint
==========================================================
Template khusus untuk PT. Telkomsel — Industri Telekomunikasi.
5 slide presentasi:
  - Slide 1 : Cover / Title Slide
  - Slide 2 : Executive Summary (bullet points)
  - Slide 3 : Data Table (tabel penuh dari CSV)
  - Slide 4 : Visualisasi (Grafik + Analisis)
  - Slide 5 : Penutup / Conclusion

Dimensi slide: 13.33" × 7.5" (widescreen 16:9)
1 inch = 914400 EMU | 1 cm = 360000 EMU
"""

import io
import math
from typing import List, Optional
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


# ─────────────────────────────────────────────────────────────────────────────
# KONSTANTA: Palet warna tema Telkomsel
# ─────────────────────────────────────────────────────────────────────────────

# Brand Telkomsel: Merah dominan + putih
COLOR_TSEL_RED        = RGBColor(0xCC, 0x00, 0x00)   # #CC0000  — Merah Telkomsel
COLOR_TSEL_RED_DARK   = RGBColor(0x99, 0x00, 0x00)   # #990000  — Merah gelap
COLOR_TSEL_RED_LIGHT  = RGBColor(0xFF, 0x33, 0x33)   # #FF3333  — Merah terang
COLOR_TSEL_GRAY_DARK  = RGBColor(0x1A, 0x1A, 0x2E)   # #1A1A2E  — Abu gelap (header)
COLOR_TSEL_GRAY_MID   = RGBColor(0x4A, 0x4A, 0x6A)   # #4A4A6A  — Abu sedang
COLOR_TSEL_GRAY_LIGHT = RGBColor(0xF5, 0xF5, 0xF7)   # #F5F5F7  — Abu terang (bg)
COLOR_TSEL_GRAY_ROW   = RGBColor(0xF0, 0xF0, 0xF5)   # #F0F0F5  — Baris tabel genap
COLOR_WHITE           = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_TEXT_DARK       = RGBColor(0x1A, 0x1A, 0x2E)
COLOR_TEXT_GRAY       = RGBColor(0x6B, 0x7B, 0x8D)
COLOR_ACCENT_GOLD     = RGBColor(0xFF, 0xC4, 0x00)   # #FFC400  — Emas aksen

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: solid fill
# ─────────────────────────────────────────────────────────────────────────────
def _fill(shape, color: RGBColor):
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = color
    shape.line.fill.background()


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: add rectangle
# ─────────────────────────────────────────────────────────────────────────────
def _rect(slide, left, top, width, height, color: RGBColor, border_color: Optional[RGBColor] = None):
    shape = slide.shapes.add_shape(
        1,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    _fill(shape, color)
    if border_color:
        shape.line.color.rgb = border_color
    return shape


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: set text
# ─────────────────────────────────────────────────────────────────────────────
def _txt(
    tf,
    text: str,
    size: int,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor = COLOR_WHITE,
    align=PP_ALIGN.LEFT,
    font: str = "Calibri",
):
    tf.clear()
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: add textbox (returns textbox shape)
# ─────────────────────────────────────────────────────────────────────────────
def _txb(slide, left, top, width, height):
    return slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )


# ─────────────────────────────────────────────────────────────────────────────
# HELPER: shared header bar (digunakan slide 2-5)
# ─────────────────────────────────────────────────────────────────────────────
def _header_bar(slide, title_text: str, slide_num: str, subtitle_text: str = ""):
    # Background seluruh slide putih
    bg = slide.shapes.add_shape(1, 0, 0, SLIDE_W, SLIDE_H)
    _fill(bg, COLOR_TSEL_GRAY_LIGHT)

    # Header bar utama (navy gelap)
    header = slide.shapes.add_shape(
        1, 0, 0, SLIDE_W, Inches(1.2)
    )
    _fill(header, COLOR_TSEL_GRAY_DARK)

    # Stripe merah Telkomsel di bawah header
    stripe = slide.shapes.add_shape(
        1, 0, Inches(1.2), SLIDE_W, Inches(0.07)
    )
    _fill(stripe, COLOR_TSEL_RED)

    # Kotak merah kecil di kiri header sebagai aksen brand
    accent_box = slide.shapes.add_shape(
        1, 0, 0, Inches(0.35), Inches(1.2)
    )
    _fill(accent_box, COLOR_TSEL_RED)

    # Judul slide di header
    txb_title = _txb(slide, 0.55, 0.15, 11.0, 0.75)
    _txt(txb_title.text_frame, title_text, size=22, bold=True,
         color=COLOR_WHITE, align=PP_ALIGN.LEFT)

    # Sub-judul kecil di header (opsional)
    if subtitle_text:
        txb_sub = _txb(slide, 0.55, 0.82, 10.0, 0.35)
        _txt(txb_sub.text_frame, subtitle_text, size=10, italic=True,
             color=RGBColor(0xBB, 0xBB, 0xCC), align=PP_ALIGN.LEFT)

    # Nomor slide di kanan header
    txb_num = _txb(slide, 12.5, 0.25, 0.7, 0.6)
    _txt(txb_num.text_frame, slide_num, size=22, bold=True,
         color=COLOR_TSEL_RED_LIGHT, align=PP_ALIGN.RIGHT)

    # Footer bar bawah slide
    footer_bg = slide.shapes.add_shape(
        1, 0, Inches(7.15), SLIDE_W, Inches(0.35)
    )
    _fill(footer_bg, COLOR_TSEL_GRAY_DARK)

    # Teks footer: brand Telkomsel
    txb_footer = _txb(slide, 0.3, 7.17, 8.0, 0.28)
    _txt(txb_footer.text_frame,
         "PT. Telkomsel — Connecting Indonesia",
         size=8, color=RGBColor(0xAA, 0xAA, 0xBB), align=PP_ALIGN.LEFT)

    # Teks footer kanan: klasifikasi dokumen
    txb_footer_r = _txb(slide, 10.0, 7.17, 3.2, 0.28)
    _txt(txb_footer_r.text_frame, "DOKUMEN INTERNAL — RAHASIA",
         size=8, color=COLOR_TSEL_RED_LIGHT, bold=True, align=PP_ALIGN.RIGHT)

    return slide


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 1: Cover / Title Slide
# ─────────────────────────────────────────────────────────────────────────────
def _slide_1_cover(prs, title, subtitle, author, report_date):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # ── Background dua panel ──────────────────────────────────────────────────
    # Panel kiri: merah Telkomsel (55% lebar)
    left_panel = slide.shapes.add_shape(
        1, 0, 0, Inches(7.3), SLIDE_H
    )
    _fill(left_panel, COLOR_TSEL_RED_DARK)

    # Panel kanan: abu gelap (45% lebar)
    right_panel = slide.shapes.add_shape(
        1, Inches(7.3), 0, Inches(6.03), SLIDE_H
    )
    _fill(right_panel, COLOR_TSEL_GRAY_DARK)

    # Garis vertikal aksen emas antara dua panel
    divider = slide.shapes.add_shape(
        1, Inches(7.28), 0, Inches(0.06), SLIDE_H
    )
    _fill(divider, COLOR_ACCENT_GOLD)

    # Strip horizontal atas (merah terang tipis)
    top_strip = slide.shapes.add_shape(
        1, 0, 0, SLIDE_W, Inches(0.12)
    )
    _fill(top_strip, COLOR_TSEL_RED_LIGHT)

    # Strip horizontal bawah
    bot_strip = slide.shapes.add_shape(
        1, 0, Inches(7.38), SLIDE_W, Inches(0.12)
    )
    _fill(bot_strip, COLOR_TSEL_RED_LIGHT)

    # ── Logo area kiri ────────────────────────────────────────────────────────
    # Label "Telkomsel" sebagai mock logo teks (karena file logo tidak tersedia)
    txb_logo = _txb(slide, 0.5, 0.3, 5.0, 0.7)
    _txt(txb_logo.text_frame, "TELKOMSEL",
         size=28, bold=True, color=COLOR_WHITE,
         align=PP_ALIGN.LEFT, font="Calibri")

    # Tagline brand
    txb_tagline = _txb(slide, 0.5, 0.92, 5.5, 0.35)
    _txt(txb_tagline.text_frame, "Connecting Indonesia · 4G / 5G Network Excellence",
         size=9, italic=True,
         color=RGBColor(0xFF, 0xCC, 0xCC),
         align=PP_ALIGN.LEFT)

    # Garis pembatas di panel kiri
    line_h = slide.shapes.add_shape(
        1, Inches(0.5), Inches(1.5), Inches(5.5), Inches(0.05)
    )
    _fill(line_h, COLOR_ACCENT_GOLD)

    # ── Judul laporan (panel kiri) ────────────────────────────────────────────
    txb_title = _txb(slide, 0.5, 1.75, 6.5, 2.8)
    txb_title.text_frame.word_wrap = True
    _txt(txb_title.text_frame, title,
         size=30, bold=True, color=COLOR_WHITE,
         align=PP_ALIGN.LEFT)

    # Sub-judul
    txb_sub = _txb(slide, 0.5, 4.45, 6.5, 0.7)
    txb_sub.text_frame.word_wrap = True
    _txt(txb_sub.text_frame, subtitle,
         size=15, italic=True,
         color=RGBColor(0xFF, 0xCC, 0xCC),
         align=PP_ALIGN.LEFT)

    # ── Panel kanan: info metadata ────────────────────────────────────────────
    # Label "LAPORAN RESMI"
    badge_box = slide.shapes.add_shape(
        1, Inches(8.0), Inches(1.0), Inches(4.5), Inches(0.45)
    )
    _fill(badge_box, COLOR_TSEL_RED)

    txb_badge = _txb(slide, 8.05, 1.05, 4.4, 0.35)
    _txt(txb_badge.text_frame, "  ▶  LAPORAN RESMI — INTERNAL USE ONLY",
         size=9, bold=True, color=COLOR_WHITE, align=PP_ALIGN.LEFT)

    # Info cards kanan
    info_items = [
        ("📅  Tanggal Laporan", report_date),
        ("👤  Disusun Oleh", author),
        ("🏢  Unit Kerja", subtitle),
        ("📊  Jenis Laporan", "Analisis Data Operasional"),
    ]

    y_info = 1.7
    for label, value in info_items:
        # Card background
        card = slide.shapes.add_shape(
            1, Inches(7.55), Inches(y_info), Inches(5.4), Inches(0.7)
        )
        _fill(card, RGBColor(0x25, 0x25, 0x45))
        card.line.color.rgb = COLOR_TSEL_RED

        # Label teks kecil
        txb_lbl = _txb(slide, 7.65, y_info + 0.06, 5.2, 0.25)
        _txt(txb_lbl.text_frame, label, size=8,
             color=RGBColor(0xAA, 0xAA, 0xCC), align=PP_ALIGN.LEFT)

        # Nilai teks
        txb_val = _txb(slide, 7.65, y_info + 0.32, 5.2, 0.32)
        _txt(txb_val.text_frame, value, size=12, bold=True,
             color=COLOR_WHITE, align=PP_ALIGN.LEFT)

        y_info += 0.82

    # Klasifikasi dokumen (bawah panel kanan)
    txb_class = _txb(slide, 7.5, 6.8, 5.5, 0.35)
    _txt(txb_class.text_frame, "RAHASIA — TIDAK UNTUK DISEBARLUASKAN",
         size=8, bold=True, color=COLOR_TSEL_RED_LIGHT, align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 2: Executive Summary
# ─────────────────────────────────────────────────────────────────────────────
def _slide_2_summary(prs, bullets: List[str], report_title: str, report_date: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header_bar(slide, "Executive Summary", "02",
                subtitle_text=f"{report_title}  ·  {report_date}")

    # Area konten: dua kolom
    # Kolom kiri: bullets (60%)
    # Kolom kanan: panel highlight statistik (40%)

    # ── Kolom kiri: bullet points ─────────────────────────────────────────────
    y_start = 1.45
    bullet_h = 0.67

    for i, text in enumerate(bullets[:7]):
        y_pos = y_start + i * bullet_h

        # Baris background alternating
        if i % 2 == 0:
            row_bg = _rect(slide, 0.3, y_pos - 0.04, 7.5, bullet_h - 0.06,
                          COLOR_TSEL_GRAY_ROW)

        # Ikon nomor merah
        num_box = slide.shapes.add_shape(
            1, Inches(0.35), Inches(y_pos), Inches(0.38), Inches(0.42)
        )
        _fill(num_box, COLOR_TSEL_RED)

        txb_num = _txb(slide, 0.36, y_pos + 0.03, 0.36, 0.35)
        _txt(txb_num.text_frame, str(i + 1), size=11, bold=True,
             color=COLOR_WHITE, align=PP_ALIGN.CENTER)

        # Teks bullet
        txb_text = _txb(slide, 0.85, y_pos + 0.02, 7.0, 0.55)
        txb_text.text_frame.word_wrap = True
        _txt(txb_text.text_frame, text, size=11,
             color=COLOR_TEXT_DARK, align=PP_ALIGN.LEFT)

    # ── Kolom kanan: panel highlight ──────────────────────────────────────────
    right_panel = _rect(slide, 8.05, 1.38, 5.0, 5.45,
                        COLOR_TSEL_GRAY_DARK)

    # Header panel kanan
    right_header = _rect(slide, 8.05, 1.38, 5.0, 0.5, COLOR_TSEL_RED)

    txb_ph = _txb(slide, 8.15, 1.43, 4.8, 0.38)
    _txt(txb_ph.text_frame, "▶  RINGKASAN EKSEKUTIF", size=11, bold=True,
         color=COLOR_WHITE, align=PP_ALIGN.LEFT)

    # Konten panel kanan: poin-poin kunci singkat
    key_points = [
        "Laporan ini merangkum performa\noperasional jaringan Telkomsel.",
        "Data dianalisis secara otomatis\ndari sumber CSV internal.",
        "Visualisasi grafik tersedia di\nSlide 04.",
        "Tabel data lengkap tersedia di\nSlide 03.",
        "Gunakan laporan ini sebagai\ndasar keputusan strategis.",
    ]

    y_kp = 2.05
    for kp in key_points:
        dot = slide.shapes.add_shape(
            1, Inches(8.2), Inches(y_kp + 0.12), Inches(0.12), Inches(0.12)
        )
        _fill(dot, COLOR_TSEL_RED_LIGHT)

        txb_kp = _txb(slide, 8.42, y_kp, 4.4, 0.65)
        txb_kp.text_frame.word_wrap = True
        _txt(txb_kp.text_frame, kp, size=10,
             color=RGBColor(0xDD, 0xDD, 0xEE), align=PP_ALIGN.LEFT)

        y_kp += 0.82


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 3: Data Table
# ─────────────────────────────────────────────────────────────────────────────
def _slide_3_table(prs, df_raw: pd.DataFrame, x_col: str, y_col: str,
                   report_title: str, report_date: str):
    """
    Menampilkan tabel data lengkap dari CSV.
    Jika data sangat banyak, ditampilkan maksimal ROWS_PER_PAGE baris per slide
    dan slide tambahan dibuat otomatis.
    """
    MAX_COLS = 6      # Maksimal 6 kolom per slide (lebar terbatas)
    ROWS_PER_PAGE = 18  # Maksimal 18 baris data + 1 baris header

    # Batasi kolom yang tampil agar tidak melebihi lebar slide
    cols_to_show = list(df_raw.columns[:MAX_COLS])
    df_display = df_raw[cols_to_show].copy()

    # Konversi semua ke string untuk tampilan
    for col in df_display.columns:
        df_display[col] = df_display[col].astype(str)

    total_rows = len(df_display)
    total_pages = math.ceil(total_rows / ROWS_PER_PAGE)

    for page_idx in range(total_pages):
        row_start = page_idx * ROWS_PER_PAGE
        row_end = min(row_start + ROWS_PER_PAGE, total_rows)
        df_page = df_display.iloc[row_start:row_end]

        slide = prs.slides.add_slide(prs.slide_layouts[6])
        page_label = f"({page_idx + 1}/{total_pages})" if total_pages > 1 else ""
        _header_bar(slide,
                    f"Tabel Data Lengkap {page_label}",
                    "03",
                    subtitle_text=f"{report_title}  ·  Menampilkan baris {row_start + 1}–{row_end} dari {total_rows} total")

        # ── Bangun tabel pptx ─────────────────────────────────────────────────
        n_cols = len(cols_to_show)
        n_rows_table = len(df_page) + 1  # +1 untuk header

        # Lebar total area tabel
        tbl_left   = Inches(0.3)
        tbl_top    = Inches(1.42)
        tbl_width  = Inches(12.73)
        tbl_height = Inches(5.5)

        table = slide.shapes.add_table(
            n_rows_table, n_cols,
            tbl_left, tbl_top, tbl_width, tbl_height
        ).table

        # Atur lebar kolom proporsional
        col_width = tbl_width // n_cols
        for ci in range(n_cols):
            table.columns[ci].width = col_width

        # ── Baris header tabel ─────────────────────────────────────────────────
        for ci, col_name in enumerate(cols_to_show):
            cell = table.cell(0, ci)
            cell.text = col_name.replace("_", " ").upper()

            # Style header: background merah Telkomsel
            fill = cell.fill
            fill.solid()
            fill.fore_color.rgb = COLOR_TSEL_RED

            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.CENTER
            run = para.runs[0] if para.runs else para.add_run()
            run.font.size = Pt(9)
            run.font.bold = True
            run.font.color.rgb = COLOR_WHITE
            run.font.name = "Calibri"

        # ── Baris data ──────────────────────────────────────────────────────────
        for ri, (_, row) in enumerate(df_page.iterrows()):
            actual_row = ri + 1  # offset karena baris 0 = header
            is_even = (ri % 2 == 0)

            for ci, col_name in enumerate(cols_to_show):
                cell = table.cell(actual_row, ci)
                cell_val = str(row[col_name])
                cell.text = cell_val

                # Alternating row color
                fill = cell.fill
                fill.solid()
                fill.fore_color.rgb = (
                    COLOR_TSEL_GRAY_ROW if is_even else COLOR_WHITE
                )

                # Highlight kolom Y (numerik utama)
                if col_name == y_col:
                    fill.fore_color.rgb = RGBColor(0xFF, 0xF0, 0xF0)  # merah sangat pucat

                para = cell.text_frame.paragraphs[0]
                para.alignment = PP_ALIGN.CENTER
                run = para.runs[0] if para.runs else para.add_run()
                run.font.size = Pt(8)
                run.font.color.rgb = COLOR_TEXT_DARK
                run.font.name = "Calibri"


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 4: Visualisasi & Analisis
# ─────────────────────────────────────────────────────────────────────────────
def _slide_4_visualization(prs, chart_buf: bytes, chart_title: str,
                            analysis_text: str, report_title: str, report_date: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header_bar(slide, "Visualisasi & Analisis Data", "04",
                subtitle_text=f"{report_title}  ·  {report_date}")

    # ── Grafik (kiri, 60% lebar) ──────────────────────────────────────────────
    chart_stream = io.BytesIO(chart_buf)
    slide.shapes.add_picture(
        chart_stream,
        Inches(0.3), Inches(1.38),
        Inches(8.3), Inches(5.45),
    )

    # ── Panel analisis kanan (40%) ────────────────────────────────────────────
    right_panel = _rect(slide, 8.8, 1.38, 4.25, 5.45, COLOR_TSEL_GRAY_DARK)

    # Header panel analisis
    panel_header = _rect(slide, 8.8, 1.38, 4.25, 0.52, COLOR_TSEL_RED)
    txb_ah = _txb(slide, 8.9, 1.43, 4.05, 0.4)
    _txt(txb_ah.text_frame, "▶  INSIGHT & ANALISIS", size=11, bold=True,
         color=COLOR_WHITE, align=PP_ALIGN.LEFT)

    # Teks analisis
    txb_analysis = _txb(slide, 8.9, 2.0, 4.0, 4.5)
    txb_analysis.text_frame.word_wrap = True
    _txt(txb_analysis.text_frame, analysis_text, size=10.5,
         color=RGBColor(0xDD, 0xDD, 0xEE), align=PP_ALIGN.LEFT)

    # Divider gold di bawah teks analisis
    div_line = slide.shapes.add_shape(
        1, Inches(8.9), Inches(6.3), Inches(3.9), Inches(0.04)
    )
    _fill(div_line, COLOR_ACCENT_GOLD)

    # Label chart di bawah grafik
    txb_ct = _txb(slide, 0.3, 6.9, 8.3, 0.28)
    _txt(txb_ct.text_frame, f"Gambar: {chart_title}", size=8, italic=True,
         color=COLOR_TEXT_GRAY, align=PP_ALIGN.CENTER)


# ─────────────────────────────────────────────────────────────────────────────
# SLIDE LAMPIRAN: 2 Gambar + 2 Panel Deskripsi (2×2 Grid, Z-Pattern)
# ─────────────────────────────────────────────────────────────────────────────
def _draw_image_quadrant(
    slide,
    img_bytes: bytes,
    panel_x: float, panel_y: float,
    panel_w: float, panel_h: float,
    pad: float = 0.15,
):
    """
    Helper: Gambar panel putih + gambar ter-center (aspect-ratio preserved)
    dalam bounding box yang ditentukan.
    Semua parameter dalam inches.
    """
    # Background putih
    bg = slide.shapes.add_shape(
        1, Inches(panel_x), Inches(panel_y), Inches(panel_w), Inches(panel_h)
    )
    _fill(bg, COLOR_WHITE)

    avail_w = int(Inches(panel_w - 2 * pad))
    avail_h = int(Inches(panel_h - 2 * pad))

    try:
        from PIL import Image as PILImage
        img_pil = PILImage.open(io.BytesIO(img_bytes))
        wp, hp = img_pil.size
        asp = wp / hp
        if int(avail_w / asp) <= avail_h:
            fw, fh = avail_w, int(avail_w / asp)
        else:
            fh, fw = avail_h, int(avail_h * asp)
        ox = (avail_w - fw) // 2
        oy = (avail_h - fh) // 2
        pl = int(Inches(panel_x + pad)) + ox
        pt = int(Inches(panel_y + pad)) + oy
    except Exception:
        fw, fh = avail_w, avail_h
        pl = int(Inches(panel_x + pad))
        pt = int(Inches(panel_y + pad))

    slide.shapes.add_picture(io.BytesIO(img_bytes), pl, pt, fw, fh)


def _draw_desc_quadrant(
    slide,
    caption: str,
    description: str,
    label_num: int,
    panel_x: float, panel_y: float,
    panel_w: float, panel_h: float,
    pm: float = 0.25,
):
    """
    Helper: Panel deskripsi merah Telkomsel dengan nomor, judul singkat
    dan keterangan pendek (max 2-3 baris).
    Semua parameter dalam inches.
    """
    # Background merah
    desc_bg = slide.shapes.add_shape(
        1, Inches(panel_x), Inches(panel_y), Inches(panel_w), Inches(panel_h)
    )
    _fill(desc_bg, COLOR_TSEL_RED)

    # Garis aksen gelap di tepi atas panel
    acc = slide.shapes.add_shape(
        1, Inches(panel_x), Inches(panel_y), Inches(panel_w), Inches(0.05)
    )
    _fill(acc, COLOR_TSEL_RED_DARK)

    # Badge nomor (kotak kecil)
    badge_sz = 0.38
    nb = slide.shapes.add_shape(
        1,
        Inches(panel_x + pm),
        Inches(panel_y + 0.18),
        Inches(badge_sz),
        Inches(badge_sz),
    )
    _fill(nb, COLOR_TSEL_RED_DARK)
    txb_n = _txb(slide, panel_x + pm + 0.01, panel_y + 0.19, badge_sz - 0.02, badge_sz - 0.02)
    _txt(txb_n.text_frame, str(label_num),
         size=12, bold=True, color=COLOR_WHITE, align=PP_ALIGN.CENTER)

    # Caption / Judul (bold)
    txb_cap = _txb(
        slide, panel_x + pm, panel_y + 0.65,
        panel_w - pm * 2, panel_h * 0.38,
    )
    txb_cap.text_frame.word_wrap = True
    _txt(txb_cap.text_frame, caption.strip() or f"Gambar {label_num}",
         size=11, bold=True, color=COLOR_WHITE, align=PP_ALIGN.LEFT)

    # Garis divider emas tipis
    div_y = panel_y + panel_h * 0.52
    div_line = slide.shapes.add_shape(
        1,
        Inches(panel_x + pm),
        Inches(div_y),
        Inches(panel_w - pm * 2),
        Inches(0.03),
    )
    _fill(div_line, COLOR_ACCENT_GOLD)

    # Deskripsi singkat (max ~3 baris)
    desc_y = div_y + 0.12
    txb_d = _txb(
        slide, panel_x + pm, desc_y,
        panel_w - pm * 2, panel_y + panel_h - desc_y - 0.15,
    )
    txb_d.text_frame.word_wrap = True
    _txt(txb_d.text_frame, description.strip() or "—",
         size=22, color=RGBColor(0xFF, 0xE0, 0xE0), align=PP_ALIGN.LEFT)


def _slide_user_images_pair(
    prs,
    img_pair: list,
    slide_num_str: str,
    report_title: str,
    report_date: str,
    pair_index: int = 0,
):
    """
    Membuat 1 slide dengan layout 2×2 grid:

      Row 1 (atas)  : [Gambar 1 kiri]   [Deskripsi 1 kanan]  (merah)
      Row 2 (bawah) : [Deskripsi 2 kiri](merah)  [Gambar 2 kanan]

    img_pair berisi 1 atau 2 item dict:
      {"bytes": <bytes>, "caption": <str>, "description": <str>}

    Jika hanya 1 gambar: row bawah digantikan branded placeholder.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    img1 = img_pair[0]
    img2 = img_pair[1] if len(img_pair) > 1 else None

    # Judul header: ambil caption gambar pertama (+ kedua jika ada)
    hdr_title = img1.get("caption") or "Lampiran Gambar"
    if img2:
        hdr_title += f"  ·  {img2.get('caption', '')}"

    _header_bar(
        slide,
        f"Lampiran: {hdr_title[:60]}",
        slide_num_str,
        subtitle_text=f"{report_title}  ·  {report_date}",
    )

    # ── Konstanta grid ───────────────────────────────────────────────
    CT  = 1.27        # Content Top
    FT  = 7.15        # Footer Top
    CH  = FT - CT     # 5.88" total content height
    RH  = CH / 2      # Row Height = 2.94" per baris
    IW  = 8.0         # Image width  (60%)
    DW  = 13.33 - IW  # Desc  width  (40% = 5.33")

    # Nomor gambar global untuk badge
    base_num = pair_index * 2 + 1  # gambar ke-1 pada pair ini

    # ── ROW 1 (atas): Gambar KIRI | Deskripsi KANAN ───────────────────
    _draw_image_quadrant(slide, img1["bytes"],
                         panel_x=0.0, panel_y=CT,
                         panel_w=IW,  panel_h=RH)
    _draw_desc_quadrant(slide,
                        caption=img1.get("caption", ""),
                        description=img1.get("description", ""),
                        label_num=base_num,
                        panel_x=IW, panel_y=CT,
                        panel_w=DW, panel_h=RH)

    # Garis pemisah horizontal (emas tipis) antara dua baris
    sep = slide.shapes.add_shape(
        1, Inches(0), Inches(CT + RH), Inches(13.33), Inches(0.04)
    )
    _fill(sep, COLOR_ACCENT_GOLD)

    # ── ROW 2 (bawah): Deskripsi KIRI | Gambar KANAN ──────────────────
    if img2:
        _draw_desc_quadrant(slide,
                            caption=img2.get("caption", ""),
                            description=img2.get("description", ""),
                            label_num=base_num + 1,
                            panel_x=0.0, panel_y=CT + RH,
                            panel_w=DW,  panel_h=RH)
        _draw_image_quadrant(slide, img2["bytes"],
                             panel_x=DW,     panel_y=CT + RH,
                             panel_w=IW,     panel_h=RH)
    else:
        # Hanya 1 gambar → row bawah = branded placeholder
        ph = slide.shapes.add_shape(
            1, Inches(0), Inches(CT + RH), Inches(13.33), Inches(RH)
        )
        _fill(ph, COLOR_TSEL_GRAY_DARK)
        txb_ph = _txb(slide, 0.5, CT + RH + RH / 2 - 0.25, 12.33, 0.5)
        _txt(txb_ph.text_frame,
             "PT. Telkomsel  ·  Connecting Indonesia  ·  Dokumen Internal",
             size=11, italic=True,
             color=RGBColor(0x88, 0x88, 0x99),
             align=PP_ALIGN.CENTER)



# ─────────────────────────────────────────────────────────────────────────────
# SLIDE 5: Penutup / Conclusion
# ─────────────────────────────────────────────────────────────────────────────
def _slide_5_conclusion(prs, author: str, report_date: str, report_title: str,
                        total_slides: int = 5):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # ── Background split: kiri merah, kanan putih ──────────────────────────────
    # Panel kiri
    lp = slide.shapes.add_shape(1, 0, 0, Inches(6.0), SLIDE_H)
    _fill(lp, COLOR_TSEL_RED_DARK)

    # Panel kanan
    rp = slide.shapes.add_shape(1, Inches(6.0), 0, Inches(7.33), SLIDE_H)
    _fill(rp, COLOR_TSEL_GRAY_LIGHT)

    # Stripe vertikal aksen emas
    sv = slide.shapes.add_shape(1, Inches(5.97), 0, Inches(0.07), SLIDE_H)
    _fill(sv, COLOR_ACCENT_GOLD)

    # Strip atas & bawah
    su = slide.shapes.add_shape(1, 0, 0, SLIDE_W, Inches(0.12))
    _fill(su, COLOR_TSEL_RED_LIGHT)
    sb = slide.shapes.add_shape(1, 0, Inches(7.38), SLIDE_W, Inches(0.12))
    _fill(sb, COLOR_TSEL_RED_LIGHT)

    # ── Panel kiri: pesan penutup ─────────────────────────────────────────────
    txb_logo = _txb(slide, 0.4, 0.3, 4.5, 0.7)
    _txt(txb_logo.text_frame, "TELKOMSEL", size=28, bold=True,
         color=COLOR_WHITE, align=PP_ALIGN.LEFT, font="Calibri")

    txb_tg = _txb(slide, 0.4, 0.92, 5.0, 0.3)
    _txt(txb_tg.text_frame, "Connecting Indonesia · 4G / 5G Network Excellence",
         size=9, italic=True, color=RGBColor(0xFF, 0xCC, 0xCC), align=PP_ALIGN.LEFT)

    line_l = slide.shapes.add_shape(1, Inches(0.4), Inches(1.4), Inches(4.8), Inches(0.05))
    _fill(line_l, COLOR_ACCENT_GOLD)

    txb_msg = _txb(slide, 0.4, 1.65, 5.2, 2.5)
    txb_msg.text_frame.word_wrap = True
    _txt(txb_msg.text_frame,
         "Terima kasih.\n\nLaporan ini disusun secara otomatis menggunakan Report Automator berbasis data CSV operasional Telkomsel.",
         size=16, color=COLOR_WHITE, align=PP_ALIGN.LEFT)

    txb_thankyou = _txb(slide, 0.4, 4.0, 5.2, 1.0)
    _txt(txb_thankyou.text_frame, "\"Leading in Digital Connectivity\"",
         size=14, italic=True, bold=True,
         color=COLOR_ACCENT_GOLD, align=PP_ALIGN.LEFT)

    txb_disclaimer = _txb(slide, 0.4, 6.55, 5.2, 0.55)
    _txt(txb_disclaimer.text_frame,
         "Dokumen ini bersifat RAHASIA. Hanya untuk keperluan internal PT. Telkomsel.",
         size=8, italic=True, color=RGBColor(0xFF, 0xCC, 0xCC), align=PP_ALIGN.LEFT)

    # ── Panel kanan: ringkasan dokumen ────────────────────────────────────────
    txb_doc_title = _txb(slide, 6.3, 0.5, 6.7, 0.65)
    _txt(txb_doc_title.text_frame, "Ringkasan Dokumen", size=18, bold=True,
         color=COLOR_TSEL_RED_DARK, align=PP_ALIGN.LEFT)

    line_r = slide.shapes.add_shape(
        1, Inches(6.3), Inches(1.1), Inches(6.5), Inches(0.05)
    )
    _fill(line_r, COLOR_TSEL_RED)

    doc_info = [
        ("Judul Laporan", report_title),
        ("Tanggal Dikeluarkan", report_date),
        ("Disusun Oleh",          author),
        ("Diklasifikasikan",      "RAHASIA — INTERNAL"),
        ("Dibuat Dengan",         "Telkomsel Report Automator v2.0"),
        ("Total Slide",           f"{total_slides} Slide"),
    ]

    y_di = 1.35
    for key, val in doc_info:
        card = _rect(slide, 6.3, y_di, 6.6, 0.78,
                     COLOR_WHITE, border_color=RGBColor(0xE0, 0xE0, 0xE5))

        txb_k = _txb(slide, 6.42, y_di + 0.06, 6.3, 0.25)
        _txt(txb_k.text_frame, key, size=8, color=COLOR_TEXT_GRAY, align=PP_ALIGN.LEFT)

        txb_v = _txb(slide, 6.42, y_di + 0.32, 6.3, 0.35)
        _txt(txb_v.text_frame, val, size=12, bold=True,
             color=COLOR_TSEL_GRAY_DARK, align=PP_ALIGN.LEFT)

        y_di += 0.88


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
    df_raw: Optional[pd.DataFrame] = None,
    x_col: str = "",
    y_col: str = "",
    user_images: Optional[List[dict]] = None,
) -> bytes:
    """
    Fungsi utama yang mengkoordinasikan pembuatan seluruh presentasi PowerPoint
    bertema Telkomsel.

    Struktur slide:
      Slide 1       : Cover
      Slide 2       : Executive Summary
      Slide 3 (+)   : Tabel Data (bisa multi-slide jika data banyak)
      Slide 4       : Visualisasi & Analisis
      Slide 5..N    : Lampiran Gambar User (opsional, maks 5)
      Slide terakhir: Penutup

    Parameter:
    ----------
    user_images : List[dict], optional
        Daftar gambar lampiran dari user. Setiap item adalah dict:
        {"bytes": <bytes gambar>, "caption": <str judul gambar>}
        Maksimal 5 gambar. Setiap gambar menjadi 1 slide sebelum Penutup.
    """
    prs = Presentation()
    prs.slide_width  = SLIDE_W   # 13.33" — widescreen 16:9
    prs.slide_height = SLIDE_H   # 7.5"

    # Slide 1: Cover
    _slide_1_cover(prs, title, subtitle, author, report_date)

    # Slide 2: Executive Summary
    _slide_2_summary(prs, bullets, report_title=title, report_date=report_date)

    # Slide 3: Tabel Data Lengkap
    if df_raw is not None and not df_raw.empty:
        _slide_3_table(prs, df_raw, x_col=x_col, y_col=y_col,
                       report_title=title, report_date=report_date)

    # Slide 4: Visualisasi & Analisis
    _slide_4_visualization(prs, chart_buf, chart_title, analysis_text,
                           report_title=title, report_date=report_date)

    # Slide 5..N: Lampiran Gambar User (disisipkan sebelum Penutup)
    images = user_images or []
    images = images[:5]  # Batasi maksimal 5 gambar

    # Kelompokkan per pasang (2 gambar per slide)
    pairs = [images[i:i+2] for i in range(0, len(images), 2)]
    for pi, pair in enumerate(pairs):
        slide_idx = 5 + pi
        slide_num_str = f"0{slide_idx}" if slide_idx < 10 else str(slide_idx)
        _slide_user_images_pair(
            prs,
            img_pair=pair,
            slide_num_str=slide_num_str,
            report_title=title,
            report_date=report_date,
            pair_index=pi,
        )

    # Slide Penutup: total = 4 slide tetap + jumlah slide pair
    total_slides = 5 + len(pairs)
    _slide_5_conclusion(
        prs,
        author=author,
        report_date=report_date,
        report_title=title,
        total_slides=total_slides,
    )

    # Simpan ke buffer BytesIO
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()
