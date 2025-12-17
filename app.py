# ==========================
# WCS Survey Editor (v47 - Colored Borders)
# ==========================

import streamlit as st
import fitz
import pandas as pd
import io
import re
from datetime import datetime
import numpy as np

# Regex pattern
pattern = re.compile(
    r"^\s*(0\d{3})\s*?\n"
    r"^\s*1\s*?\n"
    r"^\s*(\d{2,4})\s*?\n"
    r"^\s*(\d{2,4})\s*?\n",
    re.MULTILINE
)

@st.cache_data
def extract_sales_blocks(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text("text") + "\n"

    lines = text.splitlines()
    blocks = []

    for match in pattern.finditer(text):
        sales_line = match.group(1)
        order_height = int(match.group(2))
        order_width = int(match.group(3))

        start_idx = text[:match.start()].count("\n")
        for j in range(start_idx, min(start_idx + 100, len(lines))):
            if lines[j].strip().lower().startswith("reference"):
                if j >= 3:
                    reference = lines[j-3].strip()
                    location = lines[j-2].strip()
                    system = lines[j-1].strip()
                else:
                    reference, location, system = "", "", ""
                break
        else:
            reference, location, system = "", "", ""

        blocks.append({
            "sales_line": sales_line,
            "order_width": order_width,
            "order_height": order_height,
            "reference": reference,
            "location": location,
            "system": system,
            "width": None,
            "height": None,
            "location_input": "",
            "remarks": ""
        })
    return blocks

def extract_editor_value(val):
    """Extract scalar from st.data_editor NumberColumn values."""
    if pd.isna(val) or val is None or val == "":
        return None
    
    if isinstance(val, (list, np.ndarray)):
        if len(val) > 0:
            val = val[0]
        else:
            return None
    
    if isinstance(val, dict):
        if 'value' in val:
            val = val['value']
        elif len(val) > 0:
            val = list(val.values())[0]
    
    try:
        return float(val)
    except (ValueError, TypeError):
        return None

def get_fontname_for_page(page):
    return "tiro"

def draw_text_with_colored_border(page, point, text, fontname, fontsize, color, border_color=(0,0,1), border_width=1.5):
    """
    Draw colored border → white fill → text.
    border_color: (1,0,0)=RED, (0,0,1)=BLUE
    """
    if not text:
        return

    x, y = point
    box_width = 280
    box_height = fontsize * 1.5

    # Box coordinates (y is baseline, move up for full coverage)
    rect = fitz.Rect(
        x - 2,
        y - fontsize * 1.3,
        x - 2 + box_width,
        y - fontsize * 1.3 + box_height,
    )
    
    # 1. BORDER first (red/blue line)
    page.draw_rect(rect, color=border_color, fill=None, width=border_width)
    
    # 2. WHITE FILL over border
    page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1), width=0)
    
    # 3. TEXT on top
    page.insert_text(
        (x, y),
        text,
        fontsize=fontsize,
        fontname=fontname,
        color=color,
        render_mode=0,
    )

def update_pdf(pdf_bytes, entries, surveyor_name=None):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    base_offset_x = 40
    font_size = 14
    line_spacing = int(font_size * 1.6)

    # ===== Surveyor Name (black text, no border) =====
    if surveyor_name:
        for page in doc:
            name_rects = page.search_for("Name")
            if len(name_rects) >= 2:
                name_rect = name_rects[1]
            elif len(name_rects) == 1:
                name_rect = name_rects[0]
            else:
                continue

            fontname = get_fontname_for_page(page)
            text_x = name_rect.x1 + 10
            text_y = name_rect.y1 - 3

            # Simple white bg for name (no border)
            rect = fitz.Rect(text_x-2, text_y-fontsize*1.3, text_x+150, text_y+5)
            page.draw_rect(rect, color=(1,1,1), fill=(1,1,1), width=0)
            page.insert_text((text_x, text_y), surveyor_name, fontsize=11, fontname=fontname, color=(0,0,0))
            break

    # ===== Find aperture anchor positions =====
    aperture_positions = []
    for page_num, page in enumerate(doc):
        rects1 = page.search_for("Aperture Size")
        rects2 = page.search_for("aperture size")
        rects3 = page.search_for("ApertureSize")
        for inst in rects1 + rects2 + rects3:
            aperture_positions.append((page_num, inst))

    if not aperture_positions and len(entries) > 0:
        page = doc[0]
        page_rect = page.mediabox
        fallback_x = page_rect.width - 300
        fallback_y = page_rect.height - 100
        for idx in range(len(entries)):
            aperture_positions.append((0, fitz.Rect(fallback_x, fallback_y + (idx * 40), fallback_x + 250, fallback_y + (idx * 40) + 20)))

    # ===== Write each row =====
    for idx, entry in enumerate(entries):
        if idx >= len(aperture_positions):
            break

        page_num, inst = aperture_positions[idx]
        page = doc[page_num]
        fontname = get_fontname_for_page(page)

        survey_w = extract_editor_value(entry.get("width"))
        survey_h = extract_editor_value(entry.get("height"))
        order_w = extract_editor_value(entry.get("order_width"))
        order_h = extract_editor_value(entry.get("order_height"))

        print(f"ROW {idx}: w={survey_w}, h={survey_h}, order_w={order_w}, order_h={order_h}")

        # Always create size_text
        if survey_w is not None and survey_h is not None:
            size_text = f"{survey_w:.0f} x {survey_h:.0f}"
        elif survey_w is not None:
            size_text = f"{survey_w:.0f} x --"
        elif survey_h is not None:
            size_text = f"-- x {survey_h:.0f}"
        else:
            size_text = "-- x --"

        location_input = (entry.get("location_input") or "").strip()
        remarks_text = (entry.get("remarks") or "").strip()

        insert_x = inst.x1 + base_offset_x
        insert_y = inst.y0 + 2

        line1_text = f"{location_input} : {size_text}"

        # COLOR LOGIC: Border + Text color
        text_color = (0, 0, 1)  # blue text default
        border_color = (0, 0, 1)  # blue border default
        
        if survey_w is not None and survey_h is not None:
            if (order_w is not None and abs(order_w - survey_w) > 75) or \
               (order_h is not None and abs(order_h - survey_h) > 75):
                text_color = (1, 0, 0)   # red text
                border_color = (1, 0, 0) # red border

        # Line 1: Location : Size with colored border
        draw_text_with_colored_border(
            page,
            (insert_x, insert_y),
            line1_text,
            fontname=fontname,
            fontsize=font_size,
            color=text_color,
            border_color=border_color,
            border_width=1.8
        )

        # Line 2: Remarks with SAME border color
        if remarks_text:
            draw_text_with_colored_border(
                page,
                (insert_x, insert_y + line_spacing),
                remarks_text,
                fontname=fontname,
                fontsize=font_size,
                color=text_color,
                border_color=border_color,
                border_width=1.8
            )

    out_bytes = io.BytesIO()
    doc.save(out_bytes)
    return out_bytes.getvalue()

def make_excel_safe_name(name):
    return "".join(c if c.isalnum() else "_" for c in name)[:31] or "Sheet"

def build_ref_summary(df):
    order_counts = df.groupby("reference", dropna=False).size().rename("Order")
    survey_mask = df["width"].notna() & df["height"].notna()
    survey_counts = (
        df[survey_mask]
        .groupby("reference", dropna=False)
        .size()
        .rename("Survey")
    )
    summary = pd.concat([order_counts, survey_counts], axis=1).fillna(0).astype(int)
    summary.index = summary.index.fillna("Unknown")
    total_row = pd.DataFrame(
        {"Order": [summary["Order"].sum()], "Survey": [summary["Survey"].sum()]},
        index=["Total"],
    )
    summary_with_total = pd.concat([summary, total_row])
    return summary_with_total.reset_index().rename(columns={"index": "Ref"})

def clean_dataframe_for_excel(df):
    df_clean = df.copy()
    for col in ["width", "height"]:
        if col in df_clean.columns:
            df_clean[col] = df_clean[col].apply(extract_editor_value)
    return df_clean

# ==== UI ==== (unchanged from v46)
st.set_page_config(page_title="WCS Survey Editor", layout="wide", initial_sidebar_state="expanded")

st.title("WCS Survey Editor")
st.markdown("Edit survey dimensions. **Red border/text = >75mm difference**. Blue = OK.")
st.divider()

with st.sidebar:
    st.header("Instructions")
    st.markdown("""
1. Enter lot name (optional).
2. Upload survey sheet PDFs.
3. For each PDF, enter Surveyor Name and edit Width / Height.
4. **Red border** = >75mm difference. **Blue border** = OK.
5. Download combined Excel and individual PDFs.
    """)
    st.divider()
    st.caption("Fenesta Building Systems")

lot_name = st.text_input("Lot Name (optional):", value="", placeholder="Enter lot name")

uploaded_pdfs = st.file_uploader("Upload Survey Sheet PDFs", type="pdf", accept_multiple_files=True)

per_file_data = []
pdf_results = []
used_names = set()

if uploaded_pdfs:
    st.divider()

    for i, uploaded_pdf in enumerate(uploaded_pdfs, 1):
        with st.expander(f"{uploaded_pdf.name}", expanded=(i == 1)):
            surveyor_name = st.text_input("Surveyor Name", value="", key=f"surveyor_{i}")
            
            col1, col2 = st.columns([4, 1])
            with col1:
                custom_pdf_name = st.text_input(
                    "Output PDF name",
                    value=f"{uploaded_pdf.name.replace('.pdf', '')}",
                    key=f"name_{i}",
                )
            with col2:
                if st.button("Refresh", key=f"refresh_{i}"):
                    st.cache_data.clear()
                    st.rerun()

            if custom_pdf_name in used_names:
                st.error("Duplicate filename. Please change.")
                continue
            used_names.add(custom_pdf_name)

            with st.spinner("Extracting data..."):
                sales_data = extract_sales_blocks(uploaded_pdf)

            if not sales_data:
                st.warning("No sales lines detected in this PDF")
                continue

            st.markdown("Edit dimensions:")
            base_df = pd.DataFrame(sales_data)

            edited_df = st.data_editor(
                base_df,
                num_rows="fixed",
                hide_index=True,
                use_container_width=True,
                key=f"editor_{i}",
                column_config={
                    "sales_line": st.column_config.TextColumn("Sales Line", disabled=True, width="small"),
                    "order_width": st.column_config.NumberColumn("Order W", disabled=True, width="small"),
                    "order_height": st.column_config.NumberColumn("Order H", disabled=True, width="small"),
                    "reference": st.column_config.TextColumn("Ref", disabled=True),
                    "location": st.column_config.TextColumn("Location", disabled=True),
                    "system": st.column_config.TextColumn("System", disabled=True),
                    "width": st.column_config.NumberColumn("Input Width (mm)", step=1, width="medium"),
                    "height": st.column_config.NumberColumn("Input Height (mm)", step=1, width="medium"),
                    "location_input": st.column_config.TextColumn("Location", width="medium"),
                    "remarks": st.column_config.TextColumn("Remarks"),
                },
            )

            st.markdown("Summary by Reference:")
            summary_df = build_ref_summary(edited_df)
            st.dataframe(summary_df, hide_index=True, use_container_width=False)

            edited_df_clean = clean_dataframe_for_excel(edited_df)
            sheet_name = make_excel_safe_name(custom_pdf_name)
            per_file_data.append((sheet_name, edited_df_clean, custom_pdf_name, surveyor_name))

            st.info("Check terminal for ROW debug output...")
            uploaded_pdf.seek(0)
            edited_pdf = update_pdf(
                uploaded_pdf.read(),
                edited_df.to_dict("records"),
                surveyor_name=surveyor_name,
            )
            pdf_results.append((custom_pdf_name, edited_pdf, surveyor_name))

    if per_file_data:
        st.divider()
        st.header("Download Files")

        st.subheader("Combined Excel File")
        excel_file = io.BytesIO()
        with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
            for sheet_name, df_part, _, _ in per_file_data:
                df_part.to_excel(writer, index=False, sheet_name=sheet_name)

        excel_filename = (
            f"{lot_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            if lot_name
            else f"WCS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )

        st.download_button(
            label="Download Combined Excel",
            data=excel_file.getvalue(),
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.subheader("Individual PDFs")
        for pdf_name, pdf_bytes, _ in pdf_results:
            if lot_name:
                base = f"{lot_name}_{pdf_name}"
            else:
                base = pdf_name
            pdf_filename = f"{base}.pdf"

            st.download_button(
                label=f"Download {pdf_name}",
                data=pdf_bytes,
                file_name=pdf_filename,
                mime="application/pdf",
                use_container_width=True,
            )

else:
    st.info("Upload survey sheet PDFs to start editing")

st.divider()
st.caption("🔴 Red border/text = >75mm difference | 🔵 Blue border/text = Within tolerance")
