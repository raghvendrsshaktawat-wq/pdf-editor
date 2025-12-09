# ==========================
# WCS Survey Editor
# ==========================

import streamlit as st
import fitz
import pandas as pd
import io
import re
import zipfile
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

def safe_float_convert(val):
    if pd.isna(val) or val is None or val == "":
        return None
    try:
        return float(val)
    except:
        return None

def update_pdf(pdf_bytes, entries):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    base_offset_x = 40
    base_offset_y = 10
    spacing1 = 120
    spacing2 = 80
    font_size = 14

    aperture_positions = []
    for page_num, page in enumerate(doc):
        for inst in page.search_for("Aperture Size"):
            aperture_positions.append((page_num, inst))

    for idx, entry in enumerate(entries):
        if idx >= len(aperture_positions):
            break

        page_num, inst = aperture_positions[idx]
        page = doc[page_num]

        order_w = entry.get("order_width")
        order_h = entry.get("order_height")
        input_w = safe_float_convert(entry.get("width"))
        input_h = safe_float_convert(entry.get("height"))
        
        color = (0, 0, 1)  # Blue
        if order_w is not None and input_w is not None and abs(order_w - input_w) > 75:
            color = (1, 0, 0)  # Red
        if order_h is not None and input_h is not None and abs(order_h - input_h) > 75:
            color = (1, 0, 0)  # Red

        size_text = f"{input_w:.0f} x {input_h:.0f}" if input_w and input_h else "N/A"
        location_text = f"({entry.get('location_input', '')})"
        remarks_text = entry.get("remarks", "")

        insert_x = inst.x1 + base_offset_x
        insert_y = inst.y0 + base_offset_y
        page.insert_text(
            (insert_x, insert_y),
            size_text,
            fontsize=font_size,
            fontname="helv",
            color=color,
            render_mode=2,
        )

        loc_x = insert_x + spacing1
        page.insert_text(
            (loc_x, insert_y),
            location_text,
            fontsize=font_size,
            fontname="helv",
            color=color,
            render_mode=2,
        )

        if remarks_text:
            rem_x = loc_x + spacing2
            page.insert_text(
                (rem_x, insert_y),
                remarks_text,
                fontsize=font_size,
                fontname="helv",
                color=color,
                render_mode=2,
            )

    out_bytes = io.BytesIO()
    doc.save(out_bytes)
    return out_bytes.getvalue()

def make_excel_safe_name(name):
    return "".join(c if c.isalnum() else "_" for c in name)[:31] or "Sheet"

def build_ref_summary(df):
    """Build reference summary for a single PDF: Ref | Order | Survey"""
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
        {
            "Order": [summary["Order"].sum()],
            "Survey": [summary["Survey"].sum()],
        },
        index=["Total"],
    )
    summary_with_total = pd.concat([summary, total_row])
    return summary_with_total.reset_index().rename(columns={"index": "Ref"})

# ==== UI ====
st.set_page_config(
    page_title="WCS Survey Editor",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("WCS Survey Editor")
st.markdown("Edit survey dimensions. Red text in PDF output indicates >75mm differences.")
st.divider()

# Sidebar
with st.sidebar:
    st.header("Instructions")
    st.markdown("""
    1. Upload survey sheet PDFs
    2. Edit Width and Height for each window
    3. Red text in PDF = difference > 75mm
    4. Download combined Excel and individual PDFs
    """)
    st.divider()
    st.caption("Fenesta Building Systems")

# Input: Lot name
lot_name = st.text_input("Lot Name (optional):", value="", placeholder="Enter lot name")

# File upload
uploaded_pdfs = st.file_uploader(
    "Upload Survey Sheet PDFs",
    type="pdf",
    accept_multiple_files=True,
)

per_file_data = []
pdf_results = []
used_names = set()

if uploaded_pdfs:
    st.divider()

    for i, uploaded_pdf in enumerate(uploaded_pdfs, 1):
        with st.expander(f"{uploaded_pdf.name}", expanded=(i == 1)):
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

            # Editor
            st.markdown("**Edit dimensions:**")
            base_df = pd.DataFrame(sales_data)

            edited_df = st.data_editor(
                base_df,
                num_rows="fixed",
                hide_index=True,
                use_container_width=True,
                key=f"editor_{i}",
                column_config={
                    "sales_line": st.column_config.TextColumn(
                        "Sales Line", disabled=True, width="small"
                    ),
                    "order_width": st.column_config.NumberColumn(
                        "Order W", disabled=True, width="small"
                    ),
                    "order_height": st.column_config.NumberColumn(
                        "Order H", disabled=True, width="small"
                    ),
                    "reference": st.column_config.TextColumn("Ref", disabled=True),
                    "location": st.column_config.TextColumn("Location", disabled=True),
                    "system": st.column_config.TextColumn("System", disabled=True),
                    "width": st.column_config.NumberColumn(
                        "Input Width (mm)", step=1, width="medium"
                    ),
                    "height": st.column_config.NumberColumn(
                        "Input Height (mm)", step=1, width="medium"
                    ),
                    "location_input": st.column_config.TextColumn(
                        "Location", width="medium"
                    ),
                    "remarks": st.column_config.TextColumn("Remarks"),
                },
            )

            # Summary for this PDF
            st.markdown("**Summary by Reference:**")
            summary_df = build_ref_summary(edited_df)
            st.dataframe(summary_df, hide_index=True, use_container_width=False)

            # Store for combined Excel
            sheet_name = make_excel_safe_name(custom_pdf_name)
            per_file_data.append((sheet_name, edited_df, custom_pdf_name))

            # Generate PDF
            uploaded_pdf.seek(0)
            edited_pdf = update_pdf(
                uploaded_pdf.read(), edited_df.to_dict("records")
            )
            pdf_results.append((custom_pdf_name, edited_pdf))

    # Download section
    if per_file_data:
        st.divider()
        st.header("Download Files")

        # Combined Excel
        col1, col2 = st.columns([2, 2])

        with col1:
            st.markdown("**Combined Excel File**")
            excel_file = io.BytesIO()
            with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
                for sheet_name, df_part, _ in per_file_data:
                    df_part.to_excel(writer, index=False, sheet_name=sheet_name)

            excel_filename = f"{lot_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx" if lot_name else f"WCS_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            
            st.download_button(
                label="Download Combined Excel",
                data=excel_file.getvalue(),
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col2:
            st.markdown("**Individual PDFs**")

        # Individual PDF downloads
        for pdf_name, pdf_bytes in pdf_results:
            pdf_filename = f"{lot_name}_{pdf_name}.pdf" if lot_name else f"{pdf_name}.pdf"
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
st.caption("Red text in PDF indicates >75mm difference from order sizes")
