"""
Vendor Label Generator — single-file Streamlit app.

Build (or upload) a mastersheet with columns: Vendor Name, Vendor ID, Vehicle No.
The mastersheet is a fully editable grid inside the app — you can type rows
directly, add/delete rows, or upload an Excel file to prefill it and then
keep editing by hand. No upload is required.

For every row it draws a 100 mm x 75 mm label:

+------------------------------------------------+
| Vendor Name | Pheonix Harness                  |
| Vendor ID   | V01234                           |
| Vehicle No  | MH04AB1456                       |
| Serial No   | 260812-11:11-001                 |
+------------------------------------------------+

Serial No = YYMMDD-HH:MM-Seq, stamped with the real IST date/time at the
moment you click "Generate labels" — never hand-entered, always live.

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
# shrunk (if needed) to fit its column width. Raised from the old 0.42 —
# combined with the font-loading fix below, this is what actually makes the
# text big and readable instead of ant-sized.
TEXT_SIZE_FACTOR = 0.58

# Never let the auto-fit shrink text below this, no matter how long the
# value is. Previously this floor was 14px, which is why long values could
# collapse down to a near-invisible size.
MIN_FONT_SIZE = 30

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


def build_serial_no(dt: datetime, seq: int) -> str:
    """Serial format: YYMMDD-HH:MM-SSS."""
    return f"{dt:%y%m%d}-{dt:%H:%M}-{seq:03d}"


def generate_label(vendor_name: str, vendor_id: str, vehicle_no: str,
                    dt: datetime, seq: int) -> Image.Image:
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

    # Label column gets more room than before (0.34 -> 0.38) so labels like
    # "Vendor Name" can sit at a large size without being force-shrunk.
    col_split = tx0 + int(table_w * 0.38)
    label_col_w = col_split - tx0 - pad_x * 2
    value_col_w = tx1 - col_split - pad_x * 2

    def _fit_font(text, font_loader, max_size, max_w, min_size=MIN_FONT_SIZE):
        size = max(max_size, min_size)
        font = font_loader(size)
        while draw.textbbox((0, 0), text, font=font)[2] > max_w and size > min_size:
            size -= 2
            font = font_loader(size)
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

        l_font = _fit_font(label, _font_bold, max_label_size, label_col_w)
        draw.text((tx0 + pad_x, y_center), label, font=l_font, fill=BLACK, anchor="lm")

        v_font = _fit_font(value, _font_regular, max_value_size, value_col_w)
        draw.text((col_split + pad_x, y_center), value, font=v_font, fill=BLACK, anchor="lm")

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


st.title("Vendor Label Generator")
st.caption(
    "Build your mastersheet right here — add rows, edit cells, or upload an "
    "Excel file to prefill the table below and then tweak it by hand. "
    "A 100 mm x 75 mm label is generated for every row. Serial No "
    "(YYMMDD-HH:MM-Seq) is stamped with the real IST date/time at the "
    "moment you click **Generate labels** — you never type it in."
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
# Editable mastersheet. Seeds from an uploaded file the first time it's
# picked, but from then on the in-app grid (which the user can freely add
# to, delete from, or edit) is the source of truth — no upload is required
# at all to use this app.
# ---------------------------------------------------------------------------
if "mastersheet_df" not in st.session_state:
    st.session_state.mastersheet_df = pd.DataFrame(
        {"Vendor Name": [""], "Vendor ID": [""], "Vehicle No": [""]}
    )

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
        up_df = up_df[REQUIRED_COLS].dropna(how="all").reset_index(drop=True)
        st.session_state.mastersheet_df = up_df.astype(str)
        st.session_state._last_uploaded_name = uploaded.name
        st.success(f"Loaded {len(up_df)} row(s) from '{uploaded.name}' — edit freely below.")

st.subheader("Mastersheet (editable)")
st.caption("Type directly into any cell. Use the ⋮ menu or the blank bottom row to add rows; select a row and press delete to remove it.")
edited_df = st.data_editor(
    st.session_state.mastersheet_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "Vendor Name": st.column_config.TextColumn(required=True),
        "Vendor ID": st.column_config.TextColumn(required=True),
        "Vehicle No": st.column_config.TextColumn(required=True),
    },
    key="mastersheet_editor",
)
st.session_state.mastersheet_df = edited_df

df = edited_df.copy()
df.columns = [str(c).strip() for c in df.columns]
df = df.dropna(subset=REQUIRED_COLS, how="all")
df = df[df[REQUIRED_COLS].apply(lambda r: any(str(v).strip() for v in r), axis=1)].reset_index(drop=True)

if df.empty:
    st.info("Add at least one row (Vendor Name, Vendor ID, Vehicle No) to generate labels.")
else:
    st.success(f"{len(df)} row(s) ready.")

    if st.button("Generate labels", type="primary"):
        # Real IST date/time at the moment of generation, used for every row.
        gen_dt = now_ist()

        rows = []
        seq_by_date = {}
        running_seq = int(start_seq)
        for _, r in df.iterrows():
            if seq_reset_daily:
                key = gen_dt.date()
                seq_by_date[key] = seq_by_date.get(key, int(start_seq) - 1) + 1
                seq = seq_by_date[key]
            else:
                seq = running_seq
                running_seq += 1
            rows.append(
                {
                    "vendor_name": str(r["Vendor Name"]),
                    "vendor_id": str(r["Vendor ID"]),
                    "vehicle_no": str(r["Vehicle No"]),
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
