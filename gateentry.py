"""
Vendor Label Generator — single-file Streamlit app.

Upload an Excel mastersheet with columns: Vendor Name, Vendor ID, Vehicle No.
For every row it draws a 100 mm x 75 mm label:

+------------------------------------------------+
| Vendor Name | Pheonix Harness                  |
| Vendor ID   | V01234                           |
| Vehicle No  | MH04AB1456                       |
| Serial No   | 260812-11:11-001                 |
+------------------------------------------------+

Serial No = YYMMDD-HH:MM-Seq, stamped with the real date/time at the moment
the label is generated.

Run:      streamlit run app.py
Deploy:   push this single file + requirements.txt to a repo,
          then deploy on https://share.streamlit.io pointing at app.py
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
# This is the fraction of each row's height the font is sized to.
# Raise/lower this single number to change text size everywhere.
TEXT_SIZE_FACTOR = 0.42

FONT_DIR = "/usr/share/fonts/truetype/dejavu/"
FONT_BOLD = FONT_DIR + "DejaVuSans-Bold.ttf"
FONT_REGULAR = FONT_DIR + "DejaVuSans.ttf"

BLACK = (0, 0, 0)
WHITE = (255, 255, 255)


def _font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except OSError:
        # Fallback if DejaVu isn't present on the host (e.g. some cloud images)
        return ImageFont.load_default()


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

    # Fixed column split (label column gets ~34% of the table width, like the
    # reference layout), independent of font size.
    col_split = tx0 + int(table_w * 0.34)
    label_col_w = col_split - tx0 - pad_x * 2
    value_col_w = tx1 - col_split - pad_x * 2

    def _fit_font(text, path, max_size, max_w, min_size=14):
        size = max(max_size, min_size)
        font = _font(path, size)
        while draw.textbbox((0, 0), text, font=font)[2] > max_w and size > min_size:
            size -= 2
            font = _font(path, size)
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

        l_font = _fit_font(label, FONT_BOLD, max_label_size, label_col_w)
        draw.text((tx0 + pad_x, y_center), label, font=l_font, fill=BLACK, anchor="lm")

        v_font = _fit_font(value, FONT_REGULAR, max_value_size, value_col_w)
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
    "Upload a mastersheet (Excel) with columns **Vendor Name**, **Vendor ID**, "
    "**Vehicle No**. A 100 mm x 75 mm label is generated for every row. "
    "Serial No (YYMMDD-HH:MM-Seq) is stamped with the real IST date/time at "
    "the moment you click Generate."
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
    st.subheader("Need a template?")
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

uploaded = st.file_uploader("Upload mastersheet (.xlsx)", type=["xlsx"])

if uploaded is not None:
    try:
        df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Could not read the Excel file: {e}")
        st.stop()

    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(
            f"The mastersheet is missing required column(s): {', '.join(missing)}. "
            f"Required headers are exactly: {', '.join(REQUIRED_COLS)}"
        )
        st.stop()

    df = df.dropna(subset=REQUIRED_COLS, how="all").reset_index(drop=True)
    st.success(f"Loaded {len(df)} row(s) from the mastersheet.")
    st.dataframe(df[REQUIRED_COLS], use_container_width=True)

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
else:
    st.info("Upload a mastersheet to get started, or download the template from the sidebar.")
