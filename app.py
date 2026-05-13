"""
Bulk Format Converter — Streamlit App (Standalone)
---------------------------------------------------
Supports: Images, Audio, Documents (docx, doc, pdf, rtf, odt, txt, html, md, csv, xlsx)

Run:
    streamlit run app.py

Install:
    pip install streamlit Pillow pydub python-docx pymupdf pypandoc
    pip install openpyxl odfpy html2text markdown
    # For audio: https://ffmpeg.org/download.html
    # For .doc/.rtf/.odt: https://pandoc.org/installing.html
"""

import io
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

# Which conversions are supported natively (no pandoc needed)
NATIVE_DOC_PAIRS = {
    # (input_ext, output_ext)
    (".txt",  ".docx"), (".docx", ".txt"),
    (".txt",  ".html"), (".html", ".txt"),
    (".txt",  ".md"),   (".md",   ".txt"),
    (".csv",  ".xlsx"), (".xlsx", ".csv"),
    (".pdf",  ".txt"),
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


# ── Image ─────────────────────────────────────────────────────────────────────

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


# ── Audio ─────────────────────────────────────────────────────────────────────

def convert_audio(input_path: Path, output_path: Path, bitrate):
    from pydub import AudioSegment
    audio = AudioSegment.from_file(str(input_path))
    export_kwargs = {"format": output_path.suffix.lstrip(".")}
    if bitrate and output_path.suffix.lower() in (".mp3", ".ogg", ".aac"):
        export_kwargs["bitrate"] = f"{bitrate}k"
    audio.export(str(output_path), **export_kwargs)


# ── Document helpers ──────────────────────────────────────────────────────────

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
    """Use pypandoc for broad format coverage."""
    import pypandoc
    out_fmt = output_path.suffix.lstrip(".")
    # pandoc uses 'docx' for both .docx input and output
    pypandoc.convert_file(
        str(input_path),
        to=out_fmt,
        outputfile=str(output_path),
        extra_args=["--standalone"] if out_fmt == "html" else [],
    )


def convert_doc(input_path: Path, output_path: Path):
    in_ext  = input_path.suffix.lower()
    out_ext = output_path.suffix.lower()
    out_fmt = out_ext.lstrip(".")

    # ── PDF input ────────────────────────────────────────────────────────────
    if in_ext == ".pdf":
        text = _pdf_to_text(input_path)
        if out_ext == ".txt":
            output_path.write_text(text, encoding="utf-8")
        elif out_ext in (".docx", ".doc"):
            from docx import Document
            doc = Document()
            for line in text.split("\n"):
                doc.add_paragraph(line)
            doc.save(str(output_path))
        elif out_ext == ".html":
            html = "<html><body>" + "".join(f"<p>{l}</p>" for l in text.split("\n") if l.strip()) + "</body></html>"
            output_path.write_text(html, encoding="utf-8")
        elif out_ext in (".md", ".rtf", ".odt"):
            _try_pandoc(input_path, output_path)
        else:
            raise ValueError(f"PDF → {out_ext} not supported")

    # ── CSV ↔ XLSX ───────────────────────────────────────────────────────────
    elif in_ext == ".csv" and out_ext == ".xlsx":
        _csv_to_xlsx(input_path, output_path)

    elif in_ext == ".xlsx" and out_ext == ".csv":
        csv_text = _read_xlsx_as_csv(input_path)
        output_path.write_text(csv_text, encoding="utf-8")

    # ── TXT → DOCX ──────────────────────────────────────────────────────────
    elif in_ext == ".txt" and out_ext in (".docx", ".doc"):
        from docx import Document
        text = input_path.read_text(encoding="utf-8")
        doc  = Document()
        for line in text.split("\n"):
            doc.add_paragraph(line)
        doc.save(str(output_path))

    # ── DOCX → TXT ──────────────────────────────────────────────────────────
    elif in_ext in (".docx", ".doc", ".dotx", ".docm") and out_ext == ".txt":
        if in_ext == ".docx":
            output_path.write_text(_read_docx_text(input_path), encoding="utf-8")
        else:
            _try_pandoc(input_path, output_path)

    # ── TXT / MD / HTML → plain text ────────────────────────────────────────
    elif in_ext in (".md", ".html", ".htm") and out_ext == ".txt":
        _try_pandoc(input_path, output_path)

    # ── TXT → HTML / MD ─────────────────────────────────────────────────────
    elif in_ext == ".txt" and out_ext in (".html", ".md"):
        _try_pandoc(input_path, output_path)

    # ── Everything else: delegate to pandoc ─────────────────────────────────
    elif in_ext in PANDOC_INPUTS and out_fmt in PANDOC_OUTPUTS:
        _try_pandoc(input_path, output_path)

    else:
        raise ValueError(f"Unsupported document conversion: {in_ext} → {out_ext}")


# ── Result container ──────────────────────────────────────────────────────────

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
      <span class="format-badge">.docx</span><span class="format-badge">.doc</span>
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

# Pandoc warning for formats that need it
PANDOC_NEEDED_OUTPUTS = {"pdf", "rtf", "odt", "html", "md"}
if file_type_label == "📄 Document" and target_format in PANDOC_NEEDED_OUTPUTS:
    st.info("ℹ️ This format requires **Pandoc** to be installed. [Download Pandoc](https://pandoc.org/installing.html)", icon="ℹ️")

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
