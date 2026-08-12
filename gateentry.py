"""
Vendor Label Generator — single-file Streamlit app.

Upload an Excel mastersheet with columns: Vendor Name, Vendor ID, Vehicle No.
For every row it draws a 100 mm x 75 mm label matching the reference design:

+------------------------------------------------+
| Vendor Name | Pheonix Harness                  |
| Vendor ID   | V01234                           |
| Vehicle No  | MH04AB1456                       |
| Serial No   | 26 08 04 -10:18 -001             |
|             |  YYYY  MM   DD   HH:MM   Serial   |
+------------------------------------------------+

Serial No is auto-generated as YYMMDD-HH:MM-Seq.

Run:      streamlit run app.py
Deploy:   push this single file + requirements.txt to a repo,
          then deploy on https://share.streamlit.io pointing at app.py
"""

import io
import zipfile
from datetime import datetime

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

FONT_DIR = "/usr/share/fonts/truetype/dejavu/"
FONT_BOLD = FONT_DIR + "DejaVuSans-Bold.ttf"
FONT_REGULAR = FONT_DIR + "DejaVuSans.ttf"

BLACK = (0, 0, 0)
RED = (220, 0, 0)
WHITE = (255, 255, 255)


def _font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except OSError:
        # Fallback if DejaVu isn't present on the host (e.g. some cloud images)
        return ImageFont.load_default()


def build_serial_no(dt: datetime, seq: int) -> str:
    """Serial format: YYMMDD-HH:MM-SSS (matches the reference label)."""
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

    col_split = tx0 + int(table_w * 0.30)   # divider: label column vs value column

    # 4 label rows + 1 taller breakdown row
    n_label_rows = 4
    row_h = int(table_h * 0.155)
    breakdown_h = table_h - row_h * n_label_rows
    row_ys = [ty0 + i * row_h for i in range(n_label_rows + 1)]
    breakdown_top = row_ys[-1]
    breakdown_bottom = ty0 + table_h

    # ---- grid lines -------------------------------------------------------
    line_w = 2
    for y in row_ys:
        draw.line([tx0, y, tx1, y], fill=BLACK, width=line_w)
    draw.line([tx0, breakdown_bottom, tx1, breakdown_bottom], fill=BLACK, width=line_w)
    draw.line([col_split, ty0, col_split, row_ys[4]], fill=BLACK, width=line_w)
    draw.line([tx0, ty0, tx0, breakdown_bottom], fill=BLACK, width=line_w)
    draw.line([tx1, ty0, tx1, breakdown_bottom], fill=BLACK, width=line_w)

    # ---- fonts --------------------------------------------------------------
    label_font = _font(FONT_BOLD, int(row_h * 0.34))
    value_font = _font(FONT_REGULAR, int(row_h * 0.34))
    serial_font = _font(FONT_REGULAR, int(row_h * 0.34))
    tag_font = _font(FONT_REGULAR, int(row_h * 0.20))

    rows = [
        ("Vendor Name", vendor_name),
        ("Vendor ID", vendor_id),
        ("Vehicle No", vehicle_no),
        ("Serial No", None),  # drawn specially below
    ]

    pad_x = int(table_w * 0.02)
    for i, (label, value) in enumerate(rows):
        y_center = row_ys[i] + row_h / 2
        draw.text((tx0 + pad_x, y_center), label, font=label_font, fill=BLACK, anchor="lm")
        if value is not None:
            draw.text((col_split + pad_x, y_center), value, font=value_font, fill=BLACK, anchor="lm")

    # ---- Serial No value with colored segment boxes -------------------------
    yyyy = f"{dt:%y}"
    mm = f"{dt:%m}"
    dd = f"{dt:%d}"
    hhmm = f"{dt:%H:%M}"
    ser = f"{seq:03d}"

    serial_row_center = row_ys[3] + row_h / 2
    cursor_x = col_split + pad_x
    box_pad = 6

    # layout: yyyy mm dd -hhmm -ser  (e.g. "26 08 04 -10:18 -001")
    seg_texts_with_sep = [yyyy, mm, dd, "-" + hhmm, "-" + ser]

    box_centers = []
    for seg in seg_texts_with_sep:
        display = seg.lstrip("-")
        prefix = "-" if seg.startswith("-") else ""
        if prefix:
            draw.text((cursor_x, serial_row_center), prefix, font=serial_font, fill=BLACK, anchor="lm")
            bbox = draw.textbbox((cursor_x, serial_row_center), prefix, font=serial_font, anchor="lm")
            cursor_x = bbox[2] + 2

        bbox = draw.textbbox((cursor_x, serial_row_center), display, font=serial_font, anchor="lm")
        box = [bbox[0] - box_pad, bbox[1] - box_pad, bbox[2] + box_pad, bbox[3] + box_pad]
        draw.rectangle(box, outline=RED, width=2)
        draw.text((cursor_x, serial_row_center), display, font=serial_font, fill=BLACK, anchor="lm")
        box_centers.append(((box[0] + box[2]) / 2, box[3]))
        cursor_x = box[2] + 10

    # ---- breakdown labels + arrows ------------------------------------------
    tags = ["YYYY", "MM", "DD", "HH:MM", "Serial No."]
    tag_y = breakdown_top + (breakdown_h * 0.62)
    arrow_top_y = breakdown_top + (breakdown_h * 0.12)

    for (bx, by), tag in zip(box_centers, tags):
        draw.line([bx, arrow_top_y, bx, tag_y - 4], fill=RED, width=2)
        ah = 6
        draw.polygon([(bx - ah, tag_y - 4 - ah), (bx + ah, tag_y - 4 - ah), (bx, tag_y - 4)], fill=RED)
        draw.text((bx, tag_y + 4), tag, font=tag_font, fill=BLACK, anchor="ma")

    return img


# ===========================================================================
# STREAMLIT APP
# ===========================================================================

st.set_page_config(page_title="Vendor Label Generator", layout="wide")

REQUIRED_COLS = ["Vendor Name", "Vendor ID", "Vehicle No"]

st.title("Vendor Label Generator")
st.caption(
    "Upload a mastersheet (Excel) with columns **Vendor Name**, **Vendor ID**, "
    "**Vehicle No**. A 100 mm x 75 mm label is generated for every row, in the "
    "same format as the reference label (Serial No = YYMMDD-HH:MM-Seq)."
)

with st.sidebar:
    st.header("Settings")
    use_now = st.checkbox("Stamp all labels with current date/time", value=True)
    if not use_now:
        d = st.date_input("Date to stamp on labels", value=datetime.now().date())
        t = st.time_input("Time to stamp on labels", value=datetime.now().time())
        chosen_dt = datetime.combine(d, t)
    else:
        chosen_dt = None

    start_seq = st.number_input("Starting serial sequence", min_value=1, value=1, step=1)
    seq_reset_daily = st.checkbox(
        "Restart sequence at 1 for each new date in the sheet", value=False
    )

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

    # ---- assign a datetime + running sequence to every row ----------------
    base_dt = chosen_dt or datetime.now()
    rows = []
    seq_by_date = {}
    running_seq = int(start_seq)
    for _, r in df.iterrows():
        row_dt = base_dt
        if seq_reset_daily:
            key = row_dt.date()
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
                "dt": row_dt,
                "seq": seq,
                "serial_no": build_serial_no(row_dt, seq),
            }
        )

    preview_df = pd.DataFrame(
        [{"Vendor Name": x["vendor_name"], "Vendor ID": x["vendor_id"],
          "Vehicle No": x["vehicle_no"], "Serial No": x["serial_no"]} for x in rows]
    )
    st.subheader("Labels to be generated")
    st.dataframe(preview_df, use_container_width=True)

    if st.button("Generate labels", type="primary"):
        images = []
        progress = st.progress(0.0)
        for i, r in enumerate(rows):
            img = generate_label(r["vendor_name"], r["vendor_id"], r["vehicle_no"], r["dt"], r["seq"])
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
                img.save(png_buf, format="PNG")
                safe_serial = r["serial_no"].replace(":", "").replace("/", "-")
                fname = f"{r['vendor_id']}_{r['vehicle_no']}_{safe_serial}.png"
                zf.writestr(fname, png_buf.getvalue())
        zip_buf.seek(0)

        # ---- Single multi-page PDF -------------------------------------
        pdf_buf = io.BytesIO()
        pil_images = [img.convert("RGB") for _, img in images]
        pil_images[0].save(pdf_buf, format="PDF", save_all=True, append_images=pil_images[1:])
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
