"""
Bulk Format Converter — Streamlit App (Standalone)
---------------------------------------------------
Supports: Images, Audio, Documents (docx, doc, pdf, rtf, odt, txt, html, md, csv, xlsx)

Run:
    streamlit run app.py

Install (core):
    pip install streamlit Pillow pydub python-docx pymupdf pypandoc openpyxl odfpy html2text markdown

Install (optional but recommended):
    # Audio codec support:
    #   macOS:   brew install ffmpeg
    #   Ubuntu:  sudo apt-get install ffmpeg
    #   Windows: https://ffmpeg.org/download.html

    # Broad document conversion (pandoc):
    #   macOS:   brew install pandoc
    #   Ubuntu:  sudo apt-get install pandoc
    #   Windows: https://pandoc.org/installing.html

    # Legacy .doc support (LibreOffice):
    #   macOS:   brew install --cask libreoffice
    #   Ubuntu:  sudo apt-get install libreoffice
    #   Windows: https://www.libreoffice.org/download/download/
    #
    # Streamlit Cloud deployment note:
    #   Add a packages.txt file to your repo containing:
    #       libreoffice
    #       pandoc
    #   Streamlit Cloud will install them automatically via apt-get.

Dependency matrix for .doc files:
    .doc → .docx        LibreOffice only
    .doc → .txt         LibreOffice + python-docx  (no pandoc needed)
    .doc → .html/.md    LibreOffice + pandoc
    .doc → .pdf         LibreOffice only (direct export, no pandoc needed)
    .doc → .rtf/.odt    LibreOffice + pandoc
"""

import io
import shutil
import subprocess
import zipfile
import tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st


# ============================================================
# BACKEND — format definitions
# ============================================================

IMAGE_FORMATS = {
    ".png", ".jpg", ".jpeg", ".bmp", ".gif",
    ".tiff", ".tif", ".webp", ".ico", ".ppm",
}

AUDIO_FORMATS = {
    ".wav", ".mp3", ".ogg", ".flac", ".aac", ".wma", ".m4a",
}

DOC_FORMATS = {
    # Word
    ".docx", ".doc", ".dotx", ".dotm", ".docm",
    # PDF
    ".pdf",
    # Rich / Open
    ".rtf", ".odt",
    # Plain / markup
    ".txt", ".md", ".html", ".htm",
    # Data
    ".csv", ".xlsx",
}

VALID_IMAGE_OUTPUTS = {"png", "jpg", "jpeg", "bmp", "gif", "tiff", "webp"}
VALID_AUDIO_OUTPUTS = {"mp3", "wav", "ogg", "flac", "aac"}
VALID_DOC_OUTPUTS   = {"txt", "docx", "pdf", "html", "md", "rtf", "odt", "csv", "xlsx"}

DEFAULT_QUALITY = {
    "jpg": 85, "jpeg": 85, "webp": 80, "png": None,
    "mp3": 192, "ogg": 192, "aac": 192, "wav": None, "flac": None,
}

# Conversions that require pandoc
PANDOC_INPUTS  = {".docx", ".doc", ".rtf", ".odt", ".html", ".htm", ".md", ".txt", ".pdf"}
PANDOC_OUTPUTS = {"docx", "pdf", "rtf", "odt", "html", "md", "txt"}


def detect_file_type(ext: str):
    ext = ext.lower()
    if ext in IMAGE_FORMATS: return "image"
    if ext in AUDIO_FORMATS: return "audio"
    if ext in DOC_FORMATS:   return "doc"
    return None


def get_valid_outputs(file_type: str) -> set:
    return {
        "image": VALID_IMAGE_OUTPUTS,
        "audio": VALID_AUDIO_OUTPUTS,
        "doc":   VALID_DOC_OUTPUTS,
    }.get(file_type, set())


# ============================================================
# DEPENDENCY CHECKS
# ============================================================

def _check_libreoffice() -> tuple[bool, str]:
    """
    Locate the LibreOffice CLI binary (soffice).

    Returns (found: bool, path_or_message: str).
    Checks common install paths on all platforms so this works on
    macOS (Homebrew / .app bundle), Linux (apt / snap), and Windows.
    """
    # 1. Already on PATH?
    soffice = shutil.which("soffice")
    if soffice:
        return True, soffice

    # 2. Common hard-coded locations
    candidates = [
        # macOS app bundle
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        # Ubuntu snap
        "/snap/bin/libreoffice",
        # Ubuntu / Debian apt
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        # Windows (common install paths — works even on Linux CI)
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for path in candidates:
        if Path(path).exists():
            return True, path

    return False, (
        "LibreOffice not found. Install it and make sure `soffice` is on PATH.\n"
        "  macOS:   brew install --cask libreoffice\n"
        "  Ubuntu:  sudo apt-get install libreoffice\n"
        "  Windows: https://www.libreoffice.org/download/download/\n"
        "  Streamlit Cloud: add 'libreoffice' to packages.txt"
    )


def _check_pandoc() -> tuple[bool, str]:
    """Return (found, path_or_message) for Pandoc."""
    try:
        import pypandoc
        # pypandoc.get_pandoc_path() raises OSError when pandoc is missing
        path = pypandoc.get_pandoc_path()
        return True, path
    except OSError:
        pass
    except ImportError:
        pass

    pandoc = shutil.which("pandoc")
    if pandoc:
        return True, pandoc

    return False, (
        "Pandoc not found. Install it:\n"
        "  macOS:   brew install pandoc\n"
        "  Ubuntu:  sudo apt-get install pandoc\n"
        "  Windows: https://pandoc.org/installing.html\n"
        "  Streamlit Cloud: add 'pandoc' to packages.txt"
    )


# ============================================================
# IMAGE CONVERSION
# ============================================================

def convert_image(input_path: Path, output_path: Path, quality):
    from PIL import Image
    img = Image.open(input_path)
    out_ext = output_path.suffix.lower()
    if out_ext in (".jpg", ".jpeg") and img.mode in ("RGBA", "P", "LA"):
        img = img.convert("RGB")
    if out_ext == ".png" and img.mode == "P":
        img = img.convert("RGBA")
    save_kwargs = {}
    if quality and out_ext in (".jpg", ".jpeg", ".webp"):
        save_kwargs["quality"] = quality
    img.save(output_path, **save_kwargs)


# ============================================================
# AUDIO CONVERSION
# ============================================================

def convert_audio(input_path: Path, output_path: Path, bitrate):
    from pydub import AudioSegment
    audio = AudioSegment.from_file(str(input_path))
    export_kwargs = {"format": output_path.suffix.lstrip(".")}
    if bitrate and output_path.suffix.lower() in (".mp3", ".ogg", ".aac"):
        export_kwargs["bitrate"] = f"{bitrate}k"
    audio.export(str(output_path), **export_kwargs)


# ============================================================
# DOCUMENT HELPERS
# ============================================================

def _pdf_to_text(input_path: Path) -> str:
    import fitz  # PyMuPDF
    doc = fitz.open(str(input_path))
    return "\n".join(page.get_text() for page in doc)


def _read_docx_text(input_path: Path) -> str:
    from docx import Document
    doc = Document(str(input_path))
    return "\n".join(p.text for p in doc.paragraphs)


def _read_xlsx_as_csv(input_path: Path) -> str:
    import openpyxl, csv, io as _io
    wb = openpyxl.load_workbook(str(input_path), data_only=True)
    ws = wb.active
    buf = _io.StringIO()
    writer = csv.writer(buf)
    for row in ws.iter_rows(values_only=True):
        writer.writerow([("" if v is None else str(v)) for v in row])
    return buf.getvalue()


def _csv_to_xlsx(input_path: Path, output_path: Path):
    import openpyxl, csv
    wb = openpyxl.Workbook()
    ws = wb.active
    with open(str(input_path), newline="", encoding="utf-8-sig") as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save(str(output_path))


def _try_pandoc(input_path: Path, output_path: Path):
    """
    Convert via pypandoc. Raises a clear RuntimeError when pandoc is absent
    rather than surfacing pypandoc's raw OSError.
    """
    pandoc_ok, pandoc_msg = _check_pandoc()
    if not pandoc_ok:
        raise RuntimeError(pandoc_msg)

    import pypandoc
    out_fmt = output_path.suffix.lstrip(".")
    pypandoc.convert_file(
        str(input_path),
        to=out_fmt,
        outputfile=str(output_path),
        extra_args=["--standalone"] if out_fmt == "html" else [],
    )


# ── .doc → .docx via LibreOffice ─────────────────────────────────────────────

def _doc_to_docx_via_libreoffice(input_path: Path, work_dir: Path) -> Path:
    """
    Convert a legacy binary .doc file (or any LO-readable format) to .docx
    by shelling out to `soffice --headless`.

    LibreOffice always writes the output into the *directory* specified by
    --outdir, naming it <stem>.docx.  We point it at a dedicated temp
    subdirectory so the output file is predictable and isolated.

    Parameters
    ----------
    input_path : Path
        Absolute path to the source .doc file.
    work_dir : Path
        A writable directory used as LibreOffice's output directory.
        The caller is responsible for its lifecycle.

    Returns
    -------
    Path
        Absolute path to the freshly written .docx file.

    Raises
    ------
    RuntimeError
        If LibreOffice is not installed, or if the conversion process
        exits with a non-zero return code.
    """
    lo_ok, lo_path = _check_libreoffice()
    if not lo_ok:
        raise RuntimeError(lo_path)  # lo_path contains the helpful install message

    cmd = [
        lo_path,
        "--headless",
        "--norestore",           # skip crash-recovery dialog
        "--nofirststartwizard",  # skip the welcome wizard
        "--convert-to", "docx:MS Word 2007 XML",  # explicit filter avoids ambiguity
        "--outdir", str(work_dir),
        str(input_path),
    ]

    result = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=120,             # 2-minute hard cap; large docs can be slow
    )

    if result.returncode != 0:
        stderr_text = result.stderr.decode(errors="replace").strip()
        raise RuntimeError(
            f"LibreOffice conversion failed (exit {result.returncode}).\n{stderr_text}"
        )

    # LibreOffice names the output file after the input stem
    expected = work_dir / f"{input_path.stem}.docx"
    if not expected.exists():
        # Scan for any .docx LibreOffice may have produced
        matches = list(work_dir.glob("*.docx"))
        if not matches:
            raise RuntimeError(
                f"LibreOffice ran successfully but no .docx was found in {work_dir}. "
                "The input file may be corrupt or in an unsupported variant."
            )
        expected = matches[0]

    return expected


# ── .doc entry-point ──────────────────────────────────────────────────────────

def _convert_doc_file(input_path: Path, output_path: Path):
    """
    Full pipeline for legacy .doc input files.

    Strategy
    --------
    1.  If the caller only wants .docx → run LibreOffice and copy result.
    2.  If the caller wants .pdf   → use LibreOffice's native PDF export
        (higher fidelity than docx→pandoc→pdf).
    3.  For everything else        → LibreOffice to .docx first, then hand
        the .docx off to the normal convert_doc() path (python-docx / pandoc).

    This keeps LibreOffice concerns contained here and reuses all existing
    conversion logic for downstream formats.
    """
    lo_ok, lo_path = _check_libreoffice()
    if not lo_ok:
        raise RuntimeError(
            f"Cannot convert .doc files: {lo_path}"
        )

    out_ext = output_path.suffix.lower()

    with tempfile.TemporaryDirectory(prefix="bulk_conv_lo_") as lo_tmp:
        lo_dir = Path(lo_tmp)

        # ── Special case: .doc → .pdf (LibreOffice native, best fidelity) ──
        if out_ext == ".pdf":
            cmd = [
                lo_path,
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                "--convert-to", "pdf",
                "--outdir", str(lo_dir),
                str(input_path),
            ]
            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=120,
            )
            if result.returncode != 0:
                stderr_text = result.stderr.decode(errors="replace").strip()
                raise RuntimeError(
                    f"LibreOffice PDF export failed (exit {result.returncode}).\n{stderr_text}"
                )
            pdf_out = lo_dir / f"{input_path.stem}.pdf"
            if not pdf_out.exists():
                matches = list(lo_dir.glob("*.pdf"))
                if not matches:
                    raise RuntimeError(
                        "LibreOffice ran but produced no PDF. "
                        "The file may be corrupt or password-protected."
                    )
                pdf_out = matches[0]
            shutil.copy2(pdf_out, output_path)
            return

        # ── All other targets: .doc → .docx → target ────────────────────────
        docx_path = _doc_to_docx_via_libreoffice(input_path, lo_dir)

        # If the final target is .docx we are done
        if out_ext == ".docx":
            shutil.copy2(docx_path, output_path)
            return

        # Hand the intermediate .docx to the generic converter
        convert_doc(docx_path, output_path)


# ============================================================
# MAIN DOCUMENT CONVERSION DISPATCHER
# ============================================================

def convert_doc(input_path: Path, output_path: Path):
    """
    Convert between document formats.

    Routing logic (in priority order):
    1. Legacy .doc → handled entirely by _convert_doc_file() via LibreOffice
    2. PDF input   → PyMuPDF text extraction + optional pandoc
    3. CSV ↔ XLSX  → openpyxl (no pandoc needed)
    4. TXT → DOCX  → python-docx (no pandoc needed)
    5. DOCX → TXT  → python-docx (no pandoc needed)
    6. Everything else → pandoc via _try_pandoc()
    """
    in_ext  = input_path.suffix.lower()
    out_ext = output_path.suffix.lower()
    out_fmt = out_ext.lstrip(".")

    # ── 1. Legacy binary .doc (and variant Word formats) ─────────────────────
    #
    # Pandoc cannot read the legacy binary .doc format directly; it requires
    # LibreOffice as an intermediary.  We also catch .dotx / .dotm / .docm
    # here because LibreOffice handles those more reliably than pandoc does.
    #
    LIBREOFFICE_ONLY_INPUTS = {".doc", ".dotx", ".dotm", ".docm"}
    if in_ext in LIBREOFFICE_ONLY_INPUTS:
        _convert_doc_file(input_path, output_path)
        return

    # ── 2. PDF input ──────────────────────────────────────────────────────────
    if in_ext == ".pdf":
        text = _pdf_to_text(input_path)
        if out_ext == ".txt":
            output_path.write_text(text, encoding="utf-8")
        elif out_ext in (".docx",):
            from docx import Document
            doc = Document()
            for line in text.split("\n"):
                doc.add_paragraph(line)
            doc.save(str(output_path))
        elif out_ext == ".html":
            html = (
                "<html><body>"
                + "".join(f"<p>{l}</p>" for l in text.split("\n") if l.strip())
                + "</body></html>"
            )
            output_path.write_text(html, encoding="utf-8")
        elif out_ext in (".md", ".rtf", ".odt"):
            _try_pandoc(input_path, output_path)
        else:
            raise ValueError(f"PDF → {out_ext} not supported")
        return

    # ── 3. CSV ↔ XLSX ─────────────────────────────────────────────────────────
    if in_ext == ".csv" and out_ext == ".xlsx":
        _csv_to_xlsx(input_path, output_path)
        return

    if in_ext == ".xlsx" and out_ext == ".csv":
        output_path.write_text(_read_xlsx_as_csv(input_path), encoding="utf-8")
        return

    # ── 4. TXT → DOCX (python-docx, no pandoc) ───────────────────────────────
    if in_ext == ".txt" and out_ext == ".docx":
        from docx import Document
        text = input_path.read_text(encoding="utf-8")
        doc  = Document()
        for line in text.split("\n"):
            doc.add_paragraph(line)
        doc.save(str(output_path))
        return

    # ── 5. DOCX → TXT (python-docx, no pandoc) ───────────────────────────────
    if in_ext == ".docx" and out_ext == ".txt":
        output_path.write_text(_read_docx_text(input_path), encoding="utf-8")
        return

    # ── 6. Everything else: delegate to pandoc ────────────────────────────────
    if in_ext in PANDOC_INPUTS and out_fmt in PANDOC_OUTPUTS:
        _try_pandoc(input_path, output_path)
        return

    raise ValueError(f"Unsupported document conversion: {in_ext} → {out_ext}")


# ============================================================
# RESULT CONTAINER
# ============================================================

class ConversionResult:
    def __init__(self, source, target, status, error=""):
        self.source = source
        self.target = target
        self.status = status
        self.error  = error


# ============================================================
# STREAMLIT UI
# ============================================================

st.set_page_config(page_title="Bulk Format Converter", page_icon="⚡", layout="centered")

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Space+Mono&family=Syne:wght@700;800&display=swap');
  html, body, [class*="css"] { font-family: 'Syne', sans-serif; }
  .stApp { background-color: #0a0a0f; color: #e8e8f0; }
  #MainMenu, footer, header { visibility: hidden; }

  .main-title { font-size: 2.6rem; font-weight: 800; letter-spacing: -0.03em; line-height: 1; margin-bottom: 4px; }
  .main-title span { color: #7fff6a; }
  .subtitle { font-family: 'Space Mono', monospace; font-size: 0.68rem; letter-spacing: 0.12em; text-transform: uppercase; color: #555570; margin-bottom: 32px; }

  .stat-card { background: #111118; border: 1px solid #1e1e2e; border-radius: 2px; padding: 18px 20px; text-align: center; }
  .stat-number { font-size: 2rem; font-weight: 800; line-height: 1; }
  .stat-label { font-family: 'Space Mono', monospace; font-size: 0.62rem; letter-spacing: 0.1em; text-transform: uppercase; color: #555570; margin-top: 4px; }

  .result-success { border-left: 3px solid #7fff6a; padding: 8px 14px; background: #111118; margin: 4px 0; font-family: 'Space Mono', monospace; font-size: 0.72rem; }
  .result-skipped { border-left: 3px solid #555570; padding: 8px 14px; background: #111118; margin: 4px 0; font-family: 'Space Mono', monospace; font-size: 0.72rem; color: #555570; }
  .result-failed  { border-left: 3px solid #ff6aad; padding: 8px 14px; background: #111118; margin: 4px 0; font-family: 'Space Mono', monospace; font-size: 0.72rem; color: #ff6aad; }

  .section-sep { border: none; border-top: 1px solid #1e1e2e; margin: 28px 0 20px; }
  .section-label { font-family: 'Space Mono', monospace; font-size: 0.62rem; letter-spacing: 0.14em; text-transform: uppercase; color: #555570; margin-bottom: 12px; }

  .format-badge {
    display: inline-block; background: #1e1e2e; border: 1px solid #2e2e4e;
    border-radius: 2px; padding: 2px 8px; margin: 2px;
    font-family: 'Space Mono', monospace; font-size: 0.62rem;
    color: #7fff6a; letter-spacing: 0.06em;
  }
  .format-section { margin-bottom: 12px; }
  .format-section-title { font-family: 'Space Mono', monospace; font-size: 0.58rem; color: #555570; text-transform: uppercase; letter-spacing: 0.12em; margin-bottom: 4px; }

  .dep-ok   { color: #7fff6a; font-family: 'Space Mono', monospace; font-size: 0.7rem; }
  .dep-warn { color: #ffcc44; font-family: 'Space Mono', monospace; font-size: 0.7rem; }
  .dep-err  { color: #ff6aad; font-family: 'Space Mono', monospace; font-size: 0.7rem; }

  div[data-baseweb="select"] > div { background: #111118 !important; border-color: #1e1e2e !important; }
  .stSlider > div > div { background: #1e1e2e !important; }
  .stButton > button {
    background: #7fff6a !important; color: #0a0a0f !important;
    font-family: 'Space Mono', monospace !important; font-weight: 700 !important;
    font-size: 0.75rem !important; letter-spacing: 0.1em !important;
    text-transform: uppercase !important; border: none !important;
    border-radius: 0 !important; padding: 12px 28px !important;
  }
  .stButton > button:hover { background: #5fdf4a !important; }
  .stDownloadButton > button {
    background: #111118 !important; color: #7fff6a !important;
    font-family: 'Space Mono', monospace !important; font-size: 0.72rem !important;
    letter-spacing: 0.08em !important; border: 1px solid #7fff6a !important;
    border-radius: 0 !important;
  }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<div class="main-title">Bulk<span>Convert</span></div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Images · Audio · Documents — batch conversion</div>', unsafe_allow_html=True)

# ── Dependency status panel ───────────────────────────────────────────────────
# with st.expander("🔧 System Dependency Status"):
#     lo_ok,  lo_msg  = _check_libreoffice()
#     pan_ok, pan_msg = _check_pandoc()
#
#     lo_icon  = "✅" if lo_ok  else "❌"
#     pan_icon = "✅" if pan_ok else "❌"
#     lo_css   = "dep-ok" if lo_ok  else "dep-err"
#     pan_css  = "dep-ok" if pan_ok else "dep-err"
#
#     st.markdown(
#         f'<div class="{lo_css}">{lo_icon} LibreOffice — '
#         f'{"found at: " + lo_msg if lo_ok else lo_msg}</div>',
#         unsafe_allow_html=True,
#     )
#     st.markdown(
#         f'<div class="{pan_css}">{pan_icon} Pandoc — '
#         f'{"found at: " + pan_msg if pan_ok else pan_msg}</div>',
#         unsafe_allow_html=True,
#     )
#     st.caption(
#         "LibreOffice is required for .doc files. "
#         "Pandoc is required for .rtf, .odt, .md, .html outputs. "
#         "Both are optional for images and audio."
#     )

# ── Supported formats reference ───────────────────────────────────────────────
with st.expander("📋 Supported Input Formats"):
    st.markdown("""
    <div class="format-section">
      <div class="format-section-title">🖼️ Images</div>
      <span class="format-badge">.png</span><span class="format-badge">.jpg</span>
      <span class="format-badge">.jpeg</span><span class="format-badge">.bmp</span>
      <span class="format-badge">.gif</span><span class="format-badge">.tiff</span>
      <span class="format-badge">.webp</span><span class="format-badge">.ico</span>
      <span class="format-badge">.ppm</span>
    </div>
    <div class="format-section">
      <div class="format-section-title">🎵 Audio</div>
      <span class="format-badge">.mp3</span><span class="format-badge">.wav</span>
      <span class="format-badge">.ogg</span><span class="format-badge">.flac</span>
      <span class="format-badge">.aac</span><span class="format-badge">.wma</span>
      <span class="format-badge">.m4a</span>
    </div>
    <div class="format-section">
      <div class="format-section-title">📄 Documents — Word</div>
      <span class="format-badge">.docx</span>
      <span class="format-badge">.doc ⚠️ LibreOffice required</span>
      <span class="format-badge">.dotx</span><span class="format-badge">.dotm</span>
      <span class="format-badge">.docm</span>
    </div>
    <div class="format-section">
      <div class="format-section-title">📄 Documents — Other</div>
      <span class="format-badge">.pdf</span><span class="format-badge">.rtf</span>
      <span class="format-badge">.odt</span><span class="format-badge">.txt</span>
      <span class="format-badge">.md</span><span class="format-badge">.html</span>
      <span class="format-badge">.htm</span><span class="format-badge">.csv</span>
      <span class="format-badge">.xlsx</span>
    </div>
    """, unsafe_allow_html=True)

# ── Step 1: Upload ────────────────────────────────────────────────────────────
st.markdown('<hr class="section-sep"><div class="section-label">01 — Upload Files</div>', unsafe_allow_html=True)
uploaded_files = st.file_uploader("Drop files here", accept_multiple_files=True, label_visibility="collapsed")
if uploaded_files:
    names = ", ".join(f.name for f in uploaded_files[:5])
    st.caption(f"📎 {len(uploaded_files)} file(s) — {names}" + ("…" if len(uploaded_files) > 5 else ""))

# ── Step 2: Format ────────────────────────────────────────────────────────────
st.markdown('<hr class="section-sep"><div class="section-label">02 — Target Format</div>', unsafe_allow_html=True)

FORMAT_MAP = {
    "🖼️ Image":    ["webp", "png", "jpg", "jpeg", "bmp", "gif", "tiff"],
    "🎵 Audio":    ["mp3", "wav", "ogg", "flac", "aac"],
    "📄 Document": ["txt", "docx", "pdf", "html", "md", "rtf", "odt", "csv", "xlsx"],
}

QUALITY_FORMATS = {"jpg", "jpeg", "webp", "mp3", "ogg", "aac"}
IS_BITRATE      = {"mp3", "ogg", "aac"}

col1, col2 = st.columns([1, 2])
with col1:
    file_type_label = st.selectbox("File Type", list(FORMAT_MAP.keys()), label_visibility="collapsed")
with col2:
    target_format = st.selectbox("Target Format", FORMAT_MAP[file_type_label], label_visibility="collapsed")

# Inline warnings for formats that need specific tools
PANDOC_NEEDED_OUTPUTS = {"rtf", "odt", "html", "md"}
if file_type_label == "📄 Document":
    if target_format == "pdf":
        st.info(
            "ℹ️ PDF output from .doc files uses LibreOffice directly. "
            "Other document-to-PDF routes also require **Pandoc**.",
            icon="ℹ️",
        )
    elif target_format in PANDOC_NEEDED_OUTPUTS:
        pan_ok, _ = _check_pandoc()
        if not pan_ok:
            st.warning(
                f"⚠️ `.{target_format}` output requires **Pandoc**, which was not found. "
                "[Download Pandoc](https://pandoc.org/installing.html)",
                icon="⚠️",
            )
        else:
            st.info(f"ℹ️ `.{target_format}` output uses **Pandoc**.", icon="ℹ️")

quality = None
if target_format in QUALITY_FORMATS:
    st.markdown('<hr class="section-sep">', unsafe_allow_html=True)
    if target_format in IS_BITRATE:
        quality = st.slider("Bitrate (kbps)", 64, 320, int(DEFAULT_QUALITY.get(target_format, 192)), step=32)
        st.caption(f"Audio bitrate: **{quality} kbps** — higher = better quality, larger file")
    else:
        quality = st.slider("Image Quality", 1, 100, int(DEFAULT_QUALITY.get(target_format, 85)))
        st.caption(f"Quality: **{quality}/100** — higher = sharper image, larger file")

# ── Step 3: Convert ───────────────────────────────────────────────────────────
st.markdown('<hr class="section-sep"><div class="section-label">03 — Convert</div>', unsafe_allow_html=True)

if st.button("⚡ Convert All Files"):
    if not uploaded_files:
        st.warning("Please upload at least one file first.")
        st.stop()

    results = []
    converted_files = {}
    progress = st.progress(0, text="Starting…")
    total = len(uploaded_files)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp     = Path(tmpdir)
        out_dir = tmp / "out"
        out_dir.mkdir()

        for i, uf in enumerate(uploaded_files):
            progress.progress(i / total, text=f"Converting {uf.name}…")
            in_path = tmp / uf.name
            in_path.write_bytes(uf.read())

            ext       = in_path.suffix.lower()
            file_type = detect_file_type(ext)

            if file_type is None:
                results.append(ConversionResult(uf.name, "", "skipped", f"Unrecognised format ({ext})"))
                continue
            if ext == f".{target_format}":
                results.append(ConversionResult(uf.name, "", "skipped", "Already in target format"))
                continue
            if target_format not in get_valid_outputs(file_type):
                results.append(ConversionResult(uf.name, "", "skipped", f"Cannot convert {file_type} → .{target_format}"))
                continue

            out_path = out_dir / f"{in_path.stem}.{target_format}"
            c = 1
            while out_path.exists():
                out_path = out_dir / f"{in_path.stem}_{c}.{target_format}"
                c += 1

            try:
                if file_type == "image":
                    convert_image(in_path, out_path, quality)
                elif file_type == "audio":
                    convert_audio(in_path, out_path, quality)
                elif file_type == "doc":
                    convert_doc(in_path, out_path)

                converted_files[out_path.name] = out_path.read_bytes()
                results.append(ConversionResult(uf.name, out_path.name, "success"))
            except Exception as e:
                results.append(ConversionResult(uf.name, "", "failed", str(e)))

        progress.progress(1.0, text="Done!")

    # ── Summary cards ─────────────────────────────────────────────────────────
    success = sum(1 for r in results if r.status == "success")
    skipped = sum(1 for r in results if r.status == "skipped")
    failed  = sum(1 for r in results if r.status == "failed")

    c1, c2, c3 = st.columns(3)
    with c1: st.markdown(f'<div class="stat-card"><div class="stat-number" style="color:#7fff6a">{success}</div><div class="stat-label">Converted</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="stat-card"><div class="stat-number" style="color:#555570">{skipped}</div><div class="stat-label">Skipped</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="stat-card"><div class="stat-number" style="color:#ff6aad">{failed}</div><div class="stat-label">Failed</div></div>', unsafe_allow_html=True)

    # ── Per-file results ───────────────────────────────────────────────────────
    st.markdown("<br><div class='section-label'>Results</div>", unsafe_allow_html=True)
    for r in results:
        icon = {"success": "✅", "skipped": "⏭️", "failed": "❌"}.get(r.status, "•")
        css  = {"success": "result-success", "skipped": "result-skipped", "failed": "result-failed"}.get(r.status, "")
        note = f" — {r.error}" if r.error else (f" → {r.target}" if r.target else "")
        st.markdown(f'<div class="{css}">{icon} {r.source}{note}</div>', unsafe_allow_html=True)

    # ── Download ZIP ───────────────────────────────────────────────────────────
    if converted_files:
        st.markdown("<br>", unsafe_allow_html=True)
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, data in converted_files.items():
                zf.writestr(fname, data)
        zip_buf.seek(0)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label=f"⬇ Download {len(converted_files)} converted file(s) as ZIP",
            data=zip_buf,
            file_name=f"converted_{ts}.zip",
            mime="application/zip",
        )

    # ── Download log ───────────────────────────────────────────────────────────
    if results:
        ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
        log = [f"Conversion Log — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n" + "=" * 60]
        for r in results:
            log.append(f"\n[{r.status.upper()}]\n  Source: {r.source}")
            if r.target: log.append(f"  Output: {r.target}")
            if r.error:  log.append(f"  Note:   {r.error}")
        st.download_button(
            label="📋 Download conversion log",
            data="\n".join(log),
            file_name=f"conversion_log_{ts}.txt",
            mime="text/plain",
        )
