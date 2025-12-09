# ==========================
# 📝 WCS Editor (v31 - Yellow Highlighting)
# ==========================

import streamlit as st
import fitz
import pandas as pd
import io
import re
import zipfile
from datetime import datetime
import numpy as np

# Minimal color palette
FENESTA_BLUE = "#1E3A8A"
FENESTA_GRAY = "#F8FAFC"
WHITE = "#FFFFFF"

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
    """Safe conversion for Streamlit editor inputs"""
    if pd.isna(val) or val is None or val == "":
        return None
    try:
        return float(val)
    except:
        return None

def create_styled_editor(df):
    """Create styled DataFrame with yellow highlighting for >75mm differences"""
    def highlight_width_mismatch(val, row):
        if pd.isna(val) or row.get('order_width') is None:
            return ''
        diff = abs(row['order_width'] - val)
        return 'background-color: #FFF9C4; color: #000' if diff > 75 else ''
    
    def highlight_height_mismatch(val, row):
        if pd.isna(val) or row.get('order_height') is None:
            return ''
        diff = abs(row['order_height'] - val)
        return 'background-color: #FFF9C4; color: #000' if diff > 75 else ''
    
    styled_df = (df.style
        .map(highlight_width_mismatch, subset=['width'])
        .map(highlight_height_mismatch, subset=['height']))
    
    return styled_df

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

def calculate_mismatches(df):
    w_issues = 0
    h_issues = 0
    
    for idx, row in df.iterrows():
        w_val = safe_float_convert(row['width'])
        h_val = safe_float_convert(row['height'])
        
        if w_val is not None and row['order_width'] is not None:
            if abs(row['order_width'] - w_val) > 75:
                w_issues += 1
        if h_val is not None and row['order_height'] is not None:
            if abs(row['order_height'] - h_val) > 75:
                h_issues += 1
    
    return w_issues, h_issues

def make_excel_safe_name(name):
    safe = "".join(c if c.isalnum() else "_" for c in name)[:31]
    return safe if safe else "Sheet"

# ----------------- UI -----------------
st.set_page_config(
    page_title="WCS Editor", 
    layout="wide", 
    page_icon="📏",
    initial_sidebar_state="expanded"
)

# Sidebar Instructions
with st.sidebar:
    st.markdown("""
    # 📋 Quick Start
    
    **1. Upload** Survey Sheet PDF(s)
    **2. Edit** Width/Height values  
    **3. Check** 🟡 yellow highlights
    **4. Download** edited PDF + Excel
    
    ---
    **🟡 Yellow cells** = >75mm difference  
    **⚪ White cells** = Within tolerance  
    **💡 Tip:** Tab between cells for instant highlighting
    """)
    
    st.markdown("---")
    st.caption("© Fenesta Building Systems")

# Main Header - Minimal
st.title("📏 WCS Editor")
st.markdown("*Edit survey sheet dimensions accurately*")

# File Upload
uploaded_pdfs = st.file_uploader(
    "Choose PDF files", 
    type="pdf", 
    accept_multiple_files=True,
    help="Upload one or more survey sheet PDFs"
)

if uploaded_pdfs:
    st.divider()
    
    file_name_prefix = st.text_input("Output prefix", value="WCS_Edited")
    per_file_data = []
    pdf_results = []
    used_names = set()

    for i, uploaded_pdf in enumerate(uploaded_pdfs, start=1):
        with st.expander(f"📄 {uploaded_pdf.name}", expanded=(i==1)):
            col1, col2 = st.columns([4, 1])
            
            with col1:
                custom_pdf_name = st.text_input(
                    "Output name",
                    value=f"{uploaded_pdf.name.replace('.pdf','')}_edited",
                    key=f"name_{i}"
                )
            
            with col2:
                if st.button("🔄 Refresh", key=f"refresh_{i}"):
                    st.cache_data.clear()
                    st.rerun()

            if custom_pdf_name in used_names:
                st.error(f"❌ Duplicate name")
                continue
            used_names.add(custom_pdf_name)

            with st.spinner(f"Extracting data..."):
                sales_data = extract_sales_blocks(uploaded_pdf)
            
            if not sales_data:
                st.warning("⚠️ No sales lines detected")
                continue

            # Create styled editor
            df = pd.DataFrame(sales_data)
            styled_df = create_styled_editor(df)
            
            edited_df = st.data_editor(
                styled_df,
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
                    "remarks": st.column_config.TextColumn("Remarks")
                }
            )

            # Metrics
            w_issues, h_issues = calculate_mismatches(edited_df)
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Width Issues", w_issues, delta=None)
            col2.metric("Height Issues", h_issues, delta=None)
            col3.metric("Records", len(edited_df))

            # Store data
            sheet_name = make_excel_safe_name(custom_pdf_name)
            per_file_data.append((sheet_name, edited_df))
            uploaded_pdf.seek(0)
            edited_pdf = update_pdf(uploaded_pdf.read(), edited_df.to_dict("records"))
            pdf_results.append((custom_pdf_name, edited_pdf))

    if per_file_data:
        st.divider()
        st.subheader("📥 Download Files")
        
        col1, col2 = st.columns([1, 3])
        
        with col1:
            if st.button("📦 ZIP All", use_container_width=True):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    # Excel
                    excel_file = io.BytesIO()
                    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
                        for sheet_name, df_part in per_file_data:
                            df_part.to_excel(writer, index=False, sheet_name=sheet_name)
                    zf.writestr(f"{file_name_prefix}.xlsx", excel_file.getvalue())
                    
                    # PDFs
                    for pdf_name, pdf_bytes in pdf_results:
                        zf.writestr(f"{pdf_name}.pdf", pdf_bytes)
                
                zip_buffer.seek(0)
                st.download_button(
                    label="⬇️ Download ZIP",
                    data=zip_buffer.getvalue(),
                    file_name=f"{file_name_prefix}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                    mime="application/zip"
                )
        
        with col2:
            for pdf_name, pdf_bytes in pdf_results:
                st.download_button(
                    label=f"📄 {pdf_name}.pdf",
                    data=pdf_bytes,
                    file_name=f"{pdf_name}.pdf",
                    use_container_width=True
                )

else:
    st.info("👆 Upload survey sheet PDFs to get started")
