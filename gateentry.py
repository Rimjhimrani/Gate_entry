"""
AgiloGateLabel — Gate Entry Label Generator — single-file Streamlit app.

Build your label list right here in the app: each label (Vendor Name,
Vendor ID, Vehicle No.) is entered as its own set of "fill in the blank"
text fields — there is no spreadsheet-style grid to edit. You can also
upload an Excel file to bulk-prefill the fields, then keep tweaking them
by hand. No upload is required.

For every entry it draws a 100 mm x 75 mm label:

+------------------------------------------------+
| Vendor Name | Pheonix Harness                  |
| Vendor ID   | V01234                           |
| Vehicle No  | MH04AB1456                       |
| Serial No   | 260812-11:11-001                 |
+------------------------------------------------+

Serial No = YYMMDD-HH:MM-Seq. The date/time part is always the real IST
date/time stamped at the moment you click "Generate labels". The Seq part
auto-increments by default, but is also a fill-in-the-blank field per
label if you want to override it by hand.

Run:      streamlit run app.py
Deploy:   push this file + requirements.txt (+ the optional packages.txt
          included alongside it) to a repo, then deploy on
          https://share.streamlit.io pointing at app.py
"""

import base64
import io
import zipfile
from datetime import datetime

try:
    from zoneinfo import ZoneInfo
except ImportError:  # Python <3.9 fallback
    ZoneInfo = None

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

# ===========================================================================
# LABEL DRAWING (PIL)
# ===========================================================================

MM_TO_PX = 12                       # px per mm -> good screen/print quality
LABEL_W_MM, LABEL_H_MM = 100, 75
LABEL_W = LABEL_W_MM * MM_TO_PX     # 1200
LABEL_H = LABEL_H_MM * MM_TO_PX     # 900
PRINT_DPI = round(MM_TO_PX * 25.4)  # ~305 dpi -> so viewers/printers render
                                     # this image at its true 100mm x 75mm
                                     # physical size instead of guessing 72dpi
                                     # and blowing the page up huge.

# Text size is FIXED here in code (no on-screen slider / manual control).
# This is the fraction of each row's height the font STARTS at before being
# shrunk (if needed) to fit its column width/height.
TEXT_SIZE_FACTOR = 0.58

# Never let the auto-fit shrink text below this, no matter how long the
# value is (it will wrap onto more lines instead of collapsing further).
MIN_FONT_SIZE = 22

# ---------------------------------------------------------------------------
# Font loading. On some hosts (notably a fresh Streamlit Community Cloud
# container) the exact system path below does NOT exist unless you add a
# packages.txt asking apt to install it — and the old code silently fell
# back to PIL's tiny fixed-size bitmap font in that case, which is what made
# every label render like ants regardless of TEXT_SIZE_FACTOR. This version
# tries several system locations first, then falls back to the DejaVu font
# files that ship *inside* Pillow itself (always present, no system install
# needed), and only as an absolute last resort uses a scaled built-in font.
# ---------------------------------------------------------------------------
FONT_BOLD_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
    "DejaVuSans-Bold.ttf",  # resolved from Pillow's bundled fonts directory
]
FONT_REGULAR_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    "DejaVuSans.ttf",
]

_font_cache = {}


def _load_font(candidates, size):
    key = (tuple(candidates), size)
    if key in _font_cache:
        return _font_cache[key]
    for path in candidates:
        try:
            font = ImageFont.truetype(path, size)
            _font_cache[key] = font
            return font
        except OSError:
            continue
    # Absolute last resort: PIL's built-in font. Modern Pillow (>=10.1)
    # lets this be scaled; older Pillow will ignore size and stay tiny.
    try:
        font = ImageFont.load_default(size=size)
    except TypeError:
        font = ImageFont.load_default()
    _font_cache[key] = font
    return font


def _font_bold(size):
    return _load_font(FONT_BOLD_CANDIDATES, size)


def _font_regular(size):
    return _load_font(FONT_REGULAR_CANDIDATES, size)


BLACK = (0, 0, 0)
WHITE = (255, 255, 255)


def build_serial_no(dt: datetime, seq) -> str:
    """Serial format: YYMMDD-HH:MM-SSS. `seq` may be an int (zero-padded to
    3 digits) or a string override typed in by hand, used as-is."""
    if isinstance(seq, str):
        seq_part = seq.strip() or "001"
    else:
        seq_part = f"{int(seq):03d}"
    return f"{dt:%y%m%d}-{dt:%H:%M}-{seq_part}"


def _wrap_text(text, font, max_w, draw):
    """Greedy word-wrap of `text` so every line fits within max_w."""
    words = str(text).split()
    if not words:
        return [""]
    lines = []
    current = words[0]
    for w in words[1:]:
        test = f"{current} {w}"
        if draw.textbbox((0, 0), test, font=font)[2] <= max_w:
            current = test
        else:
            lines.append(current)
            current = w
    lines.append(current)
    return lines


def _fit_wrapped(text, font_loader, max_size, max_w, max_h, draw,
                  min_size=MIN_FONT_SIZE):
    """Find the largest font size (down to min_size) at which `text`,
    word-wrapped to max_w, fits within max_h. Returns (font, lines,
    line_height)."""
    size = max(max_size, min_size)
    best = None
    while size >= min_size:
        font = font_loader(size)
        lines = _wrap_text(text, font, max_w, draw)
        bbox = font.getbbox("Ag")
        line_h = (bbox[3] - bbox[1]) * 1.25
        total_h = line_h * len(lines)
        fits_w = all(draw.textbbox((0, 0), ln, font=font)[2] <= max_w for ln in lines)
        if total_h <= max_h and fits_w:
            return font, lines, line_h
        best = (font, lines, line_h)
        size -= 2
    # Hit the floor: return the smallest attempt even if it slightly
    # overflows, rather than shrinking indefinitely.
    return best


def generate_label(vendor_name: str, vendor_id: str, vehicle_no: str,
                    dt: datetime, seq) -> Image.Image:
    img = Image.new("RGB", (LABEL_W, LABEL_H), WHITE)
    draw = ImageDraw.Draw(img)

    # ---- outer double border -------------------------------------------
    margin = int(0.06 * LABEL_W)
    outer = [margin, margin, LABEL_W - margin, LABEL_H - margin]
    draw.rectangle(outer, outline=BLACK, width=4)
    inner_gap = 8
    draw.rectangle([outer[0] + inner_gap, outer[1] + inner_gap,
                    outer[2] - inner_gap, outer[3] - inner_gap],
                   outline=BLACK, width=2)

    table = [outer[0] + inner_gap, outer[1] + inner_gap,
             outer[2] - inner_gap, outer[3] - inner_gap]
    tx0, ty0, tx1, ty1 = table
    table_w = tx1 - tx0
    table_h = ty1 - ty0

    # 4 equal-height rows filling the whole table (no extra explanatory row)
    n_rows = 4
    row_h = table_h / n_rows
    row_ys = [ty0 + i * row_h for i in range(n_rows + 1)]

    # ---- fonts (large, readable; fixed in code — no manual adjustment) ------
    max_label_size = int(row_h * TEXT_SIZE_FACTOR)
    max_value_size = int(row_h * TEXT_SIZE_FACTOR)

    serial_no = build_serial_no(dt, seq)

    rows = [
        ("Vendor Name", vendor_name),
        ("Vendor ID", vendor_id),
        ("Vehicle No", vehicle_no),
        ("Serial No", serial_no),
    ]

    pad_x = int(table_w * 0.025)
    pad_y = int(row_h * 0.08)

    # Label column gets more room than before (0.34 -> 0.38) so labels like
    # "Vendor Name" can sit at a large size without being force-shrunk.
    col_split = tx0 + int(table_w * 0.38)
    label_col_w = col_split - tx0 - pad_x * 2
    value_col_w = tx1 - col_split - pad_x * 2
    value_col_h = row_h - pad_y * 2

    def _fit_label_font(text, max_size, max_w, min_size=MIN_FONT_SIZE):
        size = max(max_size, min_size)
        font = _font_bold(size)
        while draw.textbbox((0, 0), text, font=font)[2] > max_w and size > min_size:
            size -= 2
            font = _font_bold(size)
        return font

    # ---- grid lines -------------------------------------------------------
    line_w = 2
    for y in row_ys:
        draw.line([tx0, y, tx1, y], fill=BLACK, width=line_w)
    draw.line([col_split, ty0, col_split, row_ys[-1]], fill=BLACK, width=line_w)
    draw.line([tx0, ty0, tx0, row_ys[-1]], fill=BLACK, width=line_w)
    draw.line([tx1, ty0, tx1, row_ys[-1]], fill=BLACK, width=line_w)

    for i, (label, value) in enumerate(rows):
        y_center = row_ys[i] + row_h / 2

        # Headers: bold font PLUS a text stroke so they read as clearly
        # bolder than the values even on hosts where the bold font file
        # fails to load and PIL falls back to a regular-weight font.
        l_font = _fit_label_font(label, max_label_size, label_col_w)
        draw.text((tx0 + pad_x, y_center), label, font=l_font, fill=BLACK,
                   anchor="lm", stroke_width=1, stroke_fill=BLACK)

        # Values: word-wrap to the column width; only shrink font size if
        # wrapped text still doesn't fit vertically in the row.
        v_font, v_lines, v_line_h = _fit_wrapped(
            value, _font_regular, max_value_size, value_col_w, value_col_h, draw
        )
        block_h = v_line_h * len(v_lines)
        start_y = y_center - block_h / 2 + v_line_h / 2
        for li, line in enumerate(v_lines):
            draw.text((col_split + pad_x, start_y + li * v_line_h), line,
                       font=v_font, fill=BLACK, anchor="lm")

    return img


# ===========================================================================
# BRANDING (Agilomatrix logo, embedded as base64 so the app stays a single
# self-contained file — no separate image asset to deploy alongside it)
# ===========================================================================

LOGO_B64 = (
        "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"
)

# ===========================================================================
# STREAMLIT APP
# ===========================================================================

st.set_page_config(page_title="AgiloGateLabel - Gate Entry Label Generator", layout="wide")

# ---------------------------------------------------------------------------
# Styling. Background stays pure white throughout — only typography, spacing,
# borders, and a brand accent (drawn from the Agilomatrix mark's own four
# dot colors + navy wordmark) are layered on top.
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500&display=swap');

    :root {
        --amx-navy: #1B2A56;
        --amx-orange: #F2872B;
        --amx-blue: #2F80ED;
        --amx-green: #27AE60;
        --amx-pink: #EB4B98;
        --amx-surface: #F8F9FB;
        --amx-border: #E6E9EF;
        --amx-text: #262B3D;
        --amx-muted: #6B7280;
    }

    html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
        background-color: #FFFFFF !important;
        color: var(--amx-text);
        font-family: 'Inter', -apple-system, sans-serif;
    }

    h1, h2, h3, h4 {
        font-family: 'Sora', -apple-system, sans-serif !important;
        color: var(--amx-navy) !important;
        letter-spacing: -0.01em;
    }
    h1 { font-weight: 800 !important; }
    h2, h3, h4 { font-weight: 700 !important; }

    p, li { color: var(--amx-text); }

    .amx-hero-bar {
        height: 4px;
        width: 100%;
        border-radius: 4px;
        margin: 2px 0 30px 0;
        background: linear-gradient(90deg,
            var(--amx-blue) 0%, var(--amx-blue) 22%,
            var(--amx-green) 22%, var(--amx-green) 47%,
            var(--amx-pink) 47%, var(--amx-pink) 72%,
            var(--amx-orange) 72%, var(--amx-orange) 100%);
    }
    .amx-footer-bar {
        height: 3px;
        width: 100%;
        border-radius: 3px;
        margin: 4px 0 14px 0;
        opacity: 0.6;
        background: linear-gradient(90deg,
            var(--amx-blue) 0%, var(--amx-blue) 22%,
            var(--amx-green) 22%, var(--amx-green) 47%,
            var(--amx-pink) 47%, var(--amx-pink) 72%,
            var(--amx-orange) 72%, var(--amx-orange) 100%);
    }
    .amx-footer {
        text-align: center;
        color: var(--amx-muted);
        font-size: 0.76rem;
        letter-spacing: 0.09em;
        text-transform: uppercase;
        font-family: 'Inter', sans-serif;
        padding-bottom: 8px;
    }

    /* Sidebar — pure white, separated only by a hairline border */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid var(--amx-border);
    }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4 {
        color: var(--amx-navy) !important;
    }

    /* Center the main content in a comfortable reading width instead of
       stretching edge-to-edge on wide screens */
    [data-testid="stMain"] .block-container {
        max-width: 1040px;
        margin: 0 auto;
        padding-top: 2.5rem;
    }

    /* Label / field cards */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 14px !important;
        border: 1px solid var(--amx-border) !important;
        box-shadow: 0 1px 3px rgba(27, 42, 86, 0.06);
        transition: box-shadow .15s ease;
    }
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {
        box-shadow: 0 4px 14px rgba(27, 42, 86, 0.10);
    }

    /* Field labels */
    label p, label {
        font-family: 'Inter', sans-serif !important;
        font-weight: 600 !important;
        color: var(--amx-muted) !important;
        font-size: 0.78rem !important;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }

    /* Text / number inputs — monospace reads as registry/ledger data,
       fitting for gate-pass codes, plate numbers, and serials. */
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input {
        border-radius: 8px !important;
        border: 1px solid var(--amx-border) !important;
        font-family: 'IBM Plex Mono', monospace !important;
        font-size: 0.92rem !important;
        color: var(--amx-navy) !important;
    }
    div[data-testid="stTextInput"] input:focus,
    div[data-testid="stNumberInput"] input:focus {
        border-color: var(--amx-blue) !important;
        box-shadow: 0 0 0 3px rgba(47, 128, 237, 0.15) !important;
    }

    /* Buttons */
    button[kind="primary"] {
        background-color: var(--amx-navy) !important;
        border: none !important;
        border-radius: 9px !important;
        font-weight: 600 !important;
        box-shadow: 0 2px 6px rgba(27, 42, 86, 0.25);
        transition: background-color .15s ease, transform .1s ease;
    }
    button[kind="primary"]:hover {
        background-color: var(--amx-orange) !important;
        transform: translateY(-1px);
    }
    button[kind="secondary"] {
        border-radius: 9px !important;
        border: 1px solid var(--amx-navy) !important;
        color: var(--amx-navy) !important;
        font-weight: 600 !important;
        background-color: #FFFFFF !important;
    }
    button[kind="secondary"]:hover {
        background-color: var(--amx-surface) !important;
        color: var(--amx-navy) !important;
    }
    button:focus-visible {
        outline: 2px solid var(--amx-blue) !important;
        outline-offset: 2px;
    }

    /* Alerts / success banners */
    div[data-testid="stAlert"] {
        border-radius: 10px !important;
    }

    /* File uploader — white, dashed navy-tinted border rather than a tint fill */
    [data-testid="stFileUploaderDropzone"] {
        border-radius: 10px !important;
        border: 1.5px dashed var(--amx-border) !important;
        background-color: #FFFFFF !important;
    }

    /* Hero */
    .amx-hero {
        text-align: center;
        padding: 4px 0 6px 0;
    }
    .amx-hero img {
        margin-bottom: 4px;
    }
    .amx-hero h1 {
        margin: 8px 0 2px 0 !important;
        font-size: 2.6rem !important;
    }
    .amx-hero-sub {
        font-family: 'Inter', sans-serif;
        font-weight: 600;
        color: var(--amx-muted);
        text-transform: uppercase;
        letter-spacing: 0.1em;
        font-size: 0.95rem;
        margin-bottom: 18px;
    }
    .amx-hero-bar-wrap {
        display: flex;
        justify-content: center;
    }
    .amx-hero-bar-wrap .amx-hero-bar {
        max-width: 320px;
    }
    .amx-intro {
        max-width: 760px;
        margin: 18px auto 0 auto;
        text-align: center;
        color: var(--amx-muted) !important;
        font-size: 0.98rem;
        line-height: 1.6;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

REQUIRED_COLS = ["Vendor Name", "Vendor ID", "Vehicle No"]
IST = ZoneInfo("Asia/Kolkata") if ZoneInfo else None


def now_ist() -> datetime:
    """Real current date/time in India Standard Time, regardless of what
    timezone the server the app happens to be running on is set to."""
    if IST is not None:
        return datetime.now(IST).replace(tzinfo=None)
    return datetime.now()  # fallback if zoneinfo unavailable


def blank_row():
    return {"vendor_name": "", "vendor_id": "", "vehicle_no": "", "seq_override": ""}


st.markdown(
    f"""
    <div class="amx-hero">
        <img src="data:image/png;base64,{LOGO_B64}" width="170">
        <h1>AgiloGateLabel</h1>
        <div class="amx-hero-sub">Gate Entry Label Generator</div>
        <div class="amx-hero-bar-wrap"><div class="amx-hero-bar"></div></div>
        <div class="amx-intro">
            Add labels below by filling in the blanks — no spreadsheet to wrangle.
            Upload an Excel file any time to bulk-prefill the fields, then keep
            editing by hand. A 100&nbsp;mm&nbsp;x&nbsp;75&nbsp;mm label is generated
            for every entry. Serial No (YYMMDD-HH:MM-Seq): the date/time is always
            the real IST time at the moment you click <strong>Generate labels</strong>;
            the Seq number auto-increments but you can also type your own value
            into its blank.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Settings")
    start_seq = st.number_input("Starting serial sequence", min_value=1, value=1, step=1)
    seq_reset_daily = st.checkbox(
        "Restart sequence at 1 for each new date", value=False
    )

    st.markdown("---")
    st.caption(f"Current IST time: **{now_ist():%Y-%m-%d %H:%M:%S}**")

    st.markdown("---")
    st.subheader("Optional: upload to prefill")
    uploaded = st.file_uploader("Upload mastersheet (.xlsx)", type=["xlsx"])

    template_df = pd.DataFrame(
        {
            "Vendor Name": ["Pheonix Harness"],
            "Vendor ID": ["V01234"],
            "Vehicle No": ["MH04AB1456"],
        }
    )
    buf = io.BytesIO()
    template_df.to_excel(buf, index=False)
    st.download_button(
        "Download blank mastersheet template",
        data=buf.getvalue(),
        file_name="mastersheet_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------------------------------------------------------------------
# Editable label list. Each label is its own set of fill-in-the-blank text
# fields — never a spreadsheet grid. Uploading an Excel file bulk-prefills
# these fields the first time it's picked; after that the fields you've
# typed into are the source of truth. No upload is required at all.
# ---------------------------------------------------------------------------
if "rows" not in st.session_state:
    st.session_state.rows = [blank_row()]

if uploaded is not None:
    if st.session_state.get("_last_uploaded_name") != uploaded.name:
        try:
            up_df = pd.read_excel(uploaded)
        except Exception as e:
            st.error(f"Could not read the Excel file: {e}")
            st.stop()
        up_df.columns = [str(c).strip() for c in up_df.columns]
        missing = [c for c in REQUIRED_COLS if c not in up_df.columns]
        if missing:
            st.error(
                f"The uploaded file is missing required column(s): {', '.join(missing)}. "
                f"Required headers are exactly: {', '.join(REQUIRED_COLS)}"
            )
            st.stop()
        up_df = up_df[REQUIRED_COLS].dropna(how="all").fillna("").astype(str)
        st.session_state.rows = [
            {
                "vendor_name": r["Vendor Name"],
                "vendor_id": r["Vendor ID"],
                "vehicle_no": r["Vehicle No"],
                "seq_override": "",
            }
            for _, r in up_df.iterrows()
        ] or [blank_row()]
        st.session_state._last_uploaded_name = uploaded.name
        st.success(f"Loaded {len(st.session_state.rows)} label(s) from '{uploaded.name}' — edit the blanks below.")

st.subheader("Labels")
st.caption("Fill in the blanks for each label. Add more labels or remove ones you don't need.")

remove_idx = None
for i, row in enumerate(st.session_state.rows):
    with st.container(border=True):
        c1, c2, c3, c4, c5 = st.columns([3, 2, 2, 2, 1])
        row["vendor_name"] = c1.text_input("Vendor Name", value=row["vendor_name"], key=f"vn_{i}")
        row["vendor_id"] = c2.text_input("Vendor ID", value=row["vendor_id"], key=f"vid_{i}")
        row["vehicle_no"] = c3.text_input("Vehicle No", value=row["vehicle_no"], key=f"veh_{i}")
        row["seq_override"] = c4.text_input("Seq No (optional)", value=row["seq_override"],
                                             key=f"seq_{i}", placeholder="auto")
        c5.markdown("<br>", unsafe_allow_html=True)
        if c5.button("✕", key=f"rm_{i}", help="Remove this label"):
            remove_idx = i

if remove_idx is not None:
    st.session_state.rows.pop(remove_idx)
    st.rerun()

if st.button("+ Add label"):
    st.session_state.rows.append(blank_row())
    st.rerun()

active_rows = [
    r for r in st.session_state.rows
    if any(str(r[k]).strip() for k in ("vendor_name", "vendor_id", "vehicle_no"))
]

if not active_rows:
    st.info("Fill in at least one label's blanks (Vendor Name, Vendor ID, Vehicle No) to generate labels.")
else:
    st.success(f"{len(active_rows)} label(s) ready.")

    if st.button("Generate labels", type="primary"):
        # Real IST date/time at the moment of generation, used for every row.
        gen_dt = now_ist()

        rows = []
        seq_by_date = {}
        running_seq = int(start_seq)
        for r in active_rows:
            override = str(r.get("seq_override", "")).strip()
            if override:
                seq = override
            elif seq_reset_daily:
                key = gen_dt.date()
                seq_by_date[key] = seq_by_date.get(key, int(start_seq) - 1) + 1
                seq = seq_by_date[key]
            else:
                seq = running_seq
                running_seq += 1
            rows.append(
                {
                    "vendor_name": str(r["vendor_name"]),
                    "vendor_id": str(r["vendor_id"]),
                    "vehicle_no": str(r["vehicle_no"]),
                    "dt": gen_dt,
                    "seq": seq,
                    "serial_no": build_serial_no(gen_dt, seq),
                }
            )

        images = []
        progress = st.progress(0.0)
        for i, r in enumerate(rows):
            img = generate_label(r["vendor_name"], r["vendor_id"], r["vehicle_no"],
                                  r["dt"], r["seq"])
            images.append((r, img))
            progress.progress((i + 1) / len(rows))
        progress.empty()

        st.subheader("Preview")
        n_show = min(4, len(images))
        cols = st.columns(n_show) if n_show else []
        for i in range(n_show):
            r, img = images[i]
            with cols[i]:
                st.image(img, caption=r["serial_no"], use_container_width=True)
        if len(images) > n_show:
            st.caption(f"...and {len(images) - n_show} more label(s) below in the downloads.")

        # ---- Zip of individual PNGs -----------------------------------
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for r, img in images:
                png_buf = io.BytesIO()
                img.save(png_buf, format="PNG", dpi=(PRINT_DPI, PRINT_DPI))
                safe_serial = r["serial_no"].replace(":", "").replace("/", "-")
                fname = f"{r['vendor_id']}_{r['vehicle_no']}_{safe_serial}.png"
                zf.writestr(fname, png_buf.getvalue())
        zip_buf.seek(0)

        # ---- Single multi-page PDF -------------------------------------
        # resolution= tells the PDF the true physical size (100mm x 75mm per
        # page) instead of defaulting to 72dpi, which would otherwise create
        # a huge poster-sized page that viewers shrink to fit — making the
        # text look tiny even though it's drawn at full size.
        pdf_buf = io.BytesIO()
        pil_images = [img.convert("RGB") for _, img in images]
        pil_images[0].save(
            pdf_buf, format="PDF", save_all=True, append_images=pil_images[1:],
            resolution=PRINT_DPI,
        )
        pdf_buf.seek(0)

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Download all labels (ZIP of PNGs)",
                data=zip_buf.getvalue(),
                file_name="labels.zip",
                mime="application/zip",
            )
        with c2:
            st.download_button(
                "Download all labels (single PDF)",
                data=pdf_buf.getvalue(),
                file_name="labels.pdf",
                mime="application/pdf",
            )

st.markdown('<div class="amx-footer-bar"></div>', unsafe_allow_html=True)
st.markdown('<div class="amx-footer">Designed and Developed by Agilomatrix</div>', unsafe_allow_html=True)
