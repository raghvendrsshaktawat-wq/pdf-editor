# ==========================
# 📝 WCS Editor (v28 - Fixed Syntax)
# ==========================

import streamlit as st
import fitz
import pandas as pd
import io
import re
import zipfile
from datetime import datetime

# Fenesta Brand Colors
FENESTA_BLUE = "#003087"
FENESTA_RED = "#D32F2F"
FENESTA_LIGHT_BLUE = "#4FC3F7"

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
            "remarks": "",
            "w_status": "➖",
            "h_status": "➖"
        })
    return blocks

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
        input_w = entry.get("width")
        input_h = entry.get("height")

        input_w_num = float(input_w) if input_w is not None else None
        input_h_num = float(input_h) if input_h is not None else None
        
        color = (0, 0, 1)
        if order_w is not None and input_w_num is not None and abs(order_w - input_w_num) > 75:
            color = (1, 0, 0)
        if order_h is not None and input_h_num is not None and abs(order_h - input_h_num) > 75:
            color = (1, 0, 0)

        size_text = f"{input_w_num:.0f} x {input_h_num:.0f}" if input_w_num and input_h_num else "N/A"
        location_text = f"({entry.get('location_input', '')})"
        remarks_text = entry.get("remarks", "")

        insert_x = inst.x1 + base_offset_x
        insert_y = inst.y0 + base_offset_y
        page.insert_text((insert_x, insert_y), size_text,
                         fontsize=font_size, fontname="helv",
                         color=color, render_mode=2)

        loc_x = insert_x + spacing1
        page.insert_text((loc_x, insert_y), location_text,
                         fontsize=font_size, fontname="helv",
                         color=color, render_mode=2)

        if remarks_text:
            rem_x = loc_x + spacing2
            page.insert_text((rem_x, insert_y), remarks_text,
                             fontsize=font_size, fontname="helv",
                             color=color, render_mode=2)

    out_bytes = io.BytesIO()
    doc.save(out_bytes)
    return out_bytes.getvalue()

def update_status_indicators(df):
    df_copy = df.copy()
    for idx, row in df_copy.iterrows():
        if pd.isna(row['width']) or row['order_width'] is None:
            df_copy.at[idx, 'w_status'] = "➖"
        else:
            diff_w = abs(row['order_width'] - row['width'])
            df_copy.at[idx, 'w_status'] = "🔴" if diff_w > 75 else "✅"
        
        if pd.isna(row['height']) or row['order_height'] is None:
            df_copy.at[idx, 'h_status'] = "➖"
        else:
            diff_h = abs(row['order_height'] - row['height'])
            df_copy.at[idx, 'h_status'] = "🔴" if diff_h > 75 else "✅"
    return df_copy

def make_excel_safe_name(name):
    safe = "".join(c if c.isalnum() else "_" for c in name)[:31]
    return safe if safe else "Sheet"

# ----------------- UI -----------------
st.set_page_config(page_title="WCS Editor", layout="wide", page_icon="📝")

# Header
st.markdown("""
<div style='text-align: center; padding: 2rem; background: linear-gradient(90deg, #003087 0%, #4FC3F7 100%); 
           color: white; border-radius: 15px; margin-bottom: 2rem;'>
    <h1 style='margin: 0; font-size: 2.5rem;'>📝 WCS Editor Pro</h1>
    <p style='margin: 0.5rem 0 0 0; font-size: 1.2rem;'>Fenesta Smart PDF Survey Sheet Editor</p>
</div>
""", unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)
col1.metric("📄 Files", 0)
col2.metric("🔴 Critical", 0)
col3.metric("✅ Valid", 100)

st.divider()

tab1, tab2 = st.tabs(["📤 Upload & Edit", "ℹ️ Guide"])

with tab1:
    uploaded_pdfs = st.file_uploader("Upload Survey Sheet PDFs", type="pdf", accept_multiple_files=True)

    if uploaded_pdfs:
        file_name_prefix = st.text_input("📂 Output prefix", value="WCS_Edited")
        per_file_data = []
        pdf_results = []
        used_names = set()

        for i, uploaded_pdf in enumerate(uploaded_pdfs, start=1):
            with st.expander(f"📄 File {i}: {uploaded_pdf.name}", expanded=True):
                col1, col2 = st.columns([3, 1])
                with col1:
                    custom_pdf_name = st.text_input(
                        f"Output PDF name",
                        value=f"{uploaded_pdf.name.replace('.pdf','')}_edited",
                        key=f"rename_{i}"
                    )
                with col2:
                    if st.button(f"🔄 Refresh", key=f"refresh_{i}"):
                        st.cache_data.clear()
                        st.rerun()

                if custom_pdf_name in used_names:
                    st.error(f"❌ Duplicate: {custom_pdf_name}")
                    continue
                used_names.add(custom_pdf_name)

                sales_data = extract_sales_blocks(uploaded_pdf)
                if not sales_data:
                    st.warning("⚠️ No sales lines found")
                    continue

                df = pd.DataFrame(sales_data)
                df = update_status_indicators(df)

                edited_df = st.data_editor(
                    df,
                    num_rows="fixed",
                    width='stretch',
                    hide_index=True,
                    key=f"editor_{i}",
                    column_config={
                        "sales_line": st.column_config.TextColumn("Sales Line", disabled=True, width="small"),
                        "order_width": st.column_config.NumberColumn("Order W", disabled=True, width="small"),
                        "w_status": st.column_config.TextColumn("W", disabled=True, width="small"),
                        "order_height": st.column_config.NumberColumn("Order H", disabled=True, width="small"),
                        "h_status": st.column_config.TextColumn("H", disabled=True, width="small"),
                        "reference": st.column_config.TextColumn("Reference", disabled=True),
                        "location": st.column_config.TextColumn("Location", disabled=True),
                        "system": st.column_config.TextColumn("System", disabled=True, width="small"),
                        "width": st.column_config.NumberColumn("Input Width (mm)", step=1),
                        "height": st.column_config.NumberColumn("Input Height (mm)", step=1),
                        "location_input": st.column_config.TextColumn("Location (Input)"),
                        "remarks": st.column_config.TextColumn("Remarks")
                    }
                )

                w_mismatches = ((abs(edited_df['order_width'] - edited_df['width'].fillna(0)) > 75)).sum()
                h_mismatches = ((abs(edited_df['order_height'] - edited_df['height'].fillna(0)) > 75)).sum()
                
                col1, col2, col3 = st.columns(3)
                col1.metric("🔴 Width Issues", w_mismatches)
                col2.metric("🔴 Height Issues", h_mismatches)
                col3.metric("✅ Total OK", len(edited_df) * 2 - (w_mismatches + h_mismatches))

                sheet_name = make_excel_safe_name(custom_pdf_name)
                per_file_data.append((sheet_name, edited_df))
                uploaded_pdf.seek(0)
                edited_pdf = update_pdf(uploaded_pdf.read(), edited_df.to_dict("records"))
                pdf_results.append((custom_pdf_name, edited_pdf))

        if per_file_data:
            st.subheader("📥 Download")
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📦 ZIP All Files", type="primary"):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        excel_file = io.BytesIO()
                        with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
                            for sheet_name, df_part in per_file_data:
                                df_part.to_excel(writer, index=False, sheet_name=sheet_name)
                        zf.writestr(f"{file_name_prefix}.xlsx", excel_file.getvalue())
                        
                        for pdf_name, pdf_bytes in pdf_results:
                            zf.writestr(f"{pdf_name}.pdf", pdf_bytes)
                    
                    zip_buffer.seek(0)
                    st.download_button(
                        "⬇️ Download ZIP",
                        zip_buffer.getvalue(),
                        f"{file_name_prefix}_{datetime.now().strftime('%Y%m%d')}.zip",
                        "application/zip"
                    )
            
            with col2:
                for pdf_name, pdf_bytes in pdf_results:
                    st.download_button(f"📄 {pdf_name}.pdf", pdf_bytes, f"{pdf_name}.pdf")

with tab2:
    st.markdown("""
## 🎯 Quick Guide

**🔴 Red Indicators** = Difference >75mm from order sizes  
**✅ Green** = Within tolerance (±75mm)  
**➖ Grey** = No input yet

### Workflow:
1. Upload Survey PDFs
2. Edit Width/Height values  
3. Watch indicators update live (🔴/✅)
4. Download edited PDFs + Excel
    """)

st.markdown("---")
st.caption("© Fenesta Building Systems | Streamlit 1.52.1 | v28")
