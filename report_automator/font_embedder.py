"""
font_embedder.py — Embed font Poppins ke dalam file PPTX
=========================================================

MASALAH:
  run.font.name = "Poppins" hanya memberi INSTRUKSI pakai Poppins.
  Jika Poppins tidak terinstall di sistem atau tidak di-embed,
  PowerPoint / Canva / Google Slides fallback ke Times New Roman.

SOLUSI:
  Embed file TTF Poppins langsung ke dalam arsip PPTX (file PPTX
  pada dasarnya adalah ZIP). Font yang di-embed akan selalu muncul
  dengan benar di platform manapun — termasuk Canva.

CARA KERJA:
  1. Download Poppins TTF dari Google Fonts (sekali saja, di-cache lokal)
  2. Setelah PPTX di-generate, buka sebagai ZIP
  3. Suntikkan font TTF ke dalam ppt/fonts/
  4. Tambahkan relationship dan deklarasi embeddedFont di presentation.xml
  5. Simpan kembali sebagai PPTX baru

PENGGUNAAN:
  from font_embedder import embed_poppins_into_pptx

  pptx_bytes = generate_pptx(...)          # bytes hasil generator
  pptx_bytes = embed_poppins_into_pptx(pptx_bytes)   # ← tambahkan baris ini
  # Sekarang kirim pptx_bytes ke user
"""

import io
import os
import zipfile
import urllib.request
from pathlib import Path
from typing import Dict, Optional
import logging

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# URL download Poppins dari GitHub (Google Fonts mirror, no auth needed)
# ─────────────────────────────────────────────────────────────────────────────
_FONT_CACHE_DIR = Path(__file__).parent / "_font_cache"

_POPPINS_URLS: Dict[str, str] = {
    "Regular": (
        "https://github.com/google/fonts/raw/main/ofl/poppins/Poppins-Regular.ttf"
    ),
    "Bold": (
        "https://github.com/google/fonts/raw/main/ofl/poppins/Poppins-Bold.ttf"
    ),
    "Italic": (
        "https://github.com/google/fonts/raw/main/ofl/poppins/Poppins-Italic.ttf"
    ),
    "BoldItalic": (
        "https://github.com/google/fonts/raw/main/ofl/poppins/Poppins-BoldItalic.ttf"
    ),
}

# Namespace XML PPTX
_NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_P   = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
_REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"


def _download_font(style: str, url: str) -> Optional[bytes]:
    """
    Download font TTF. Cache lokal di _font_cache/ agar tidak re-download.
    Returns bytes atau None jika gagal.
    """
    _FONT_CACHE_DIR.mkdir(exist_ok=True)
    cache_path = _FONT_CACHE_DIR / f"Poppins-{style}.ttf"

    if cache_path.exists():
        logger.info(f"[font_embedder] Poppins-{style}: menggunakan cache lokal")
        return cache_path.read_bytes()

    logger.info(f"[font_embedder] Mendownload Poppins-{style} dari {url} ...")
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (TelkomselReportAutomator/3.0)"},
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = resp.read()
        cache_path.write_bytes(data)
        logger.info(f"[font_embedder] Poppins-{style}: berhasil ({len(data):,} bytes)")
        return data
    except Exception as exc:
        logger.warning(f"[font_embedder] Gagal download Poppins-{style}: {exc}")
        return None


def _ensure_poppins_fonts() -> Dict[str, bytes]:
    """
    Pastikan semua varian Poppins tersedia (cache atau download).
    Returns dict {style: bytes}, bisa kosong jika semua gagal.
    """
    fonts: Dict[str, bytes] = {}
    for style, url in _POPPINS_URLS.items():
        data = _download_font(style, url)
        if data:
            fonts[style] = data
    return fonts


def embed_poppins_into_pptx(pptx_bytes: bytes) -> bytes:
    """
    Embed font Poppins ke dalam file PPTX (bytes → bytes).

    Langkah:
    1. Download/cache Poppins TTF (Regular, Bold, Italic, BoldItalic)
    2. Buka PPTX sebagai ZIP in-memory
    3. Tambahkan font TTF ke ppt/fonts/
    4. Tambahkan relationship di ppt/_rels/presentation.xml.rels
    5. Tambahkan <p:embeddedFont> di ppt/presentation.xml
    6. Return PPTX baru sebagai bytes

    Jika download gagal total, kembalikan pptx_bytes asli tanpa modifikasi
    (graceful degradation — tidak crash).
    """
    fonts = _ensure_poppins_fonts()
    if not fonts:
        logger.warning("[font_embedder] Tidak ada font Poppins yang berhasil di-load. "
                       "PPTX dikembalikan tanpa embed font.")
        return pptx_bytes

    # ── Buka PPTX sebagai ZIP ──────────────────────────────────────────────
    in_buf  = io.BytesIO(pptx_bytes)
    out_buf = io.BytesIO()

    style_to_rel_id: Dict[str, str] = {}   # "Regular" → "rId_poppins_reg" dst.

    # File yang akan kita patch — skip saat copy awal, tulis versi patched nanti
    # WAJIB include [Content_Types].xml agar python-pptx tidak crash saat
    # membuka PPTX dengan font embed (error: "no content-type for partname")
    FILES_TO_PATCH = {
        "[Content_Types].xml",
        "ppt/_rels/presentation.xml.rels",
        "ppt/presentation.xml",
    }

    with zipfile.ZipFile(in_buf, "r") as zin, \
         zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:

        # Salin semua file KECUALI yang akan di-patch (hindari duplikat)
        existing_names = set()
        for item in zin.infolist():
            existing_names.add(item.filename)
            if item.filename not in FILES_TO_PATCH:
                zout.writestr(item, zin.read(item.filename))

        # ── 1. Tambahkan TTF ke ppt/fonts/ ──────────────────────────────
        style_map = {
            "Regular":    ("regular",    "rIdPoppinsReg"),
            "Bold":       ("bold",        "rIdPoppinsBold"),
            "Italic":     ("italic",      "rIdPoppinsItalic"),
            "BoldItalic": ("boldItalic",  "rIdPoppinsBoldItalic"),
        }

        added_styles = []
        for style, ttf_bytes in fonts.items():
            attr_name, rel_id = style_map[style]
            font_path = f"ppt/fonts/Poppins-{style}.fntdata"
            if font_path not in existing_names:
                zout.writestr(font_path, ttf_bytes)
            style_to_rel_id[attr_name] = rel_id
            added_styles.append((attr_name, rel_id, font_path))

        # ── 2. Patch ppt/_rels/presentation.xml.rels ────────────────────
        rels_path = "ppt/_rels/presentation.xml.rels"
        if rels_path in existing_names:
            rels_xml = zin.read(rels_path).decode("utf-8")
        else:
            rels_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/'
                'package/2006/relationships"/>'
            )

        # Sisipkan Relationship untuk tiap font sebelum </Relationships>
        new_rels = ""
        for attr_name, rel_id, font_path in added_styles:
            target = "/" + font_path  # absolute part URI
            rel_line = (
                f'<Relationship Id="{rel_id}" '
                f'Type="{_REL_FONT}" '
                f'Target="{target}"/>'
            )
            if rel_id not in rels_xml:   # jangan duplikat
                new_rels += rel_line + "\n"

        if new_rels:
            rels_xml = rels_xml.replace(
                "</Relationships>",
                new_rels + "</Relationships>"
            )
        zout.writestr(rels_path, rels_xml.encode("utf-8"))

        # ── 3. Patch [Content_Types].xml ── daftarkan .fntdata content type ─
        # WAJIB: setiap file dalam ZIP PPTX harus terdaftar di sini.
        # Tanpa ini, python-pptx crash dengan:
        #   "no content-type for partname '/ppt/fonts/Poppins-X.fntdata'"
        ct_path = "[Content_Types].xml"
        if ct_path in existing_names:
            ct_xml = zin.read(ct_path).decode("utf-8")
        else:
            ct_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
            )

        # Tambahkan Default untuk ekstensi .fntdata jika belum ada
        fntdata_ct = 'Extension="fntdata" ContentType="application/x-fontdata"'
        if fntdata_ct not in ct_xml:
            ct_xml = ct_xml.replace(
                "</Types>",
                f'<Default {fntdata_ct}/>\n</Types>'
            )
        zout.writestr(ct_path, ct_xml.encode("utf-8"))

        # ── 4. Patch ppt/presentation.xml ── tambahkan <p:embeddedFont> ─
        prs_path = "ppt/presentation.xml"
        if prs_path in existing_names:
            prs_xml = zin.read(prs_path).decode("utf-8")

            # Bangun blok <p:embeddedFont> untuk Poppins
            # Hanya sertakan varian yang berhasil di-download
            def _font_ref(attr: str) -> str:
                rid = style_to_rel_id.get(attr)
                if rid:
                    return f'<p:font typeface="Poppins" r:{attr}="{rid}"/>'
                return ""

            font_block = (
                '<p:embeddedFont>\n'
                f'  <p:font typeface="Poppins"/>\n'
                + ("".join(
                    f'  {_font_ref(a)}\n'
                    for a in ("regular", "bold", "italic", "boldItalic")
                    if _font_ref(a)
                ))
                + '</p:embeddedFont>\n'
            )

            # Sisipkan sebelum </p:embeddedFontLst> jika sudah ada, atau buat baru
            if "</p:embeddedFontLst>" in prs_xml:
                if "Poppins" not in prs_xml:
                    prs_xml = prs_xml.replace(
                        "</p:embeddedFontLst>",
                        font_block + "</p:embeddedFontLst>"
                    )
            elif "</p:presentation>" in prs_xml:
                # Buat blok embeddedFontLst baru
                full_block = (
                    "<p:embeddedFontLst>\n"
                    + font_block
                    + "</p:embeddedFontLst>\n"
                )
                if "Poppins" not in prs_xml:
                    prs_xml = prs_xml.replace(
                        "</p:presentation>",
                        full_block + "</p:presentation>"
                    )

            zout.writestr(prs_path, prs_xml.encode("utf-8"))

    out_buf.seek(0)
    result = out_buf.read()
    logger.info(
        f"[font_embedder] Selesai. "
        f"Ukuran PPTX: {len(pptx_bytes):,} → {len(result):,} bytes "
        f"(+{len(result) - len(pptx_bytes):,} bytes font data)"
    )
    return result


# ─────────────────────────────────────────────────────────────────────────────
# CLI sederhana untuk test: python font_embedder.py input.pptx output.pptx
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("Usage: python font_embedder.py <input.pptx> <output.pptx>")
        sys.exit(1)

    logging.basicConfig(level=logging.INFO)

    src = Path(sys.argv[1]).read_bytes()
    dst = embed_poppins_into_pptx(src)
    Path(sys.argv[2]).write_bytes(dst)
    print(f"✅ Selesai: {sys.argv[2]}")
