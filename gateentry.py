"""
Vendor Label Generator — single-file Streamlit app.

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
# STREAMLIT APP
# ===========================================================================

st.set_page_config(page_title="Vendor Label Generator", layout="wide")

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


st.title("Vendor Label Generator")
st.caption(
    "Add labels below by filling in the blanks — no spreadsheet to wrangle. "
    "Upload an Excel file any time to bulk-prefill the fields, then keep "
    "editing by hand. A 100 mm x 75 mm label is generated for every entry. "
    "Serial No (YYMMDD-HH:MM-Seq): the date/time is always the real IST "
    "time at the moment you click **Generate labels**; the Seq number "
    "auto-increments but you can also type your own value into its blank."
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
