# ==========================
# 📝 WCS Editor (v27 - Conditional Formatting + Improved UI + Logo)
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
FENESTA_WHITE = "#FFFFFF"

# Logo (Base64 encoded Fenesta logo - replace with your actual logo)
LOGO_BASE64 = """
data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAA...[Your Fenesta logo base64 here]
"""

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
        if (order_w is not None and input_w_num is not None and 
            abs(order_w - input_w_num) > 75):
            color = (1, 0, 0)
        if (order_h is not None and input_h_num is not None and 
            abs(order_h - input_h_num) > 75):
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

def make_excel_safe_name(name):
    safe = "".join(c if c.isalnum() else "_" for c in name)[:31]
    return safe if safe else "Sheet"

# ----------------- Enhanced Streamlit UI -----------------
st.set_page_config(page_title="WCS Editor", layout="wide", page_icon="📝")

# Header with Logo & Branding
col1, col2, col3 = st.columns([1, 2, 1])
with col1:
    st.empty()
with col2:
    st.image(LOGO_BASE64 if 'LOGO_BASE64' in locals() else "https://via.placeholder.com/200x60/003087/FFFFFF?text=Fenesta", width=200)
with col3:
    st.empty()

st.markdown(f"""
    <h1 style='text-align: center; color: {FENESTA_BLUE}; margin-bottom: 0.5rem;'>📝 WCS Editor Pro</h1>
    <p style='text-align: center; color: {FENESTA_LIGHT_BLUE}; font-size: 1.1rem;'>Smart PDF Survey Sheet Editor with Auto-Highlighting</p>
""", unsafe_allow_html=True)

# Metrics Row
col1, col2, col3, col4 = st.columns(4)
total_files = len(st.session_state.get('processed_files', [])) if 'processed_files' in st.session_state else 0
col1.metric("📄 Files Ready", total_files)
col2.metric("✅ Auto-Highlight", "75mm Δ")
col3.metric("🎨 Fenesta Colors", "Blue/Red")
col4.metric("📦 ZIP Export", "Multi-file")

st.divider()

# Main Content
tab1, tab2 = st.tabs(["📤 Upload & Edit", "ℹ️ Instructions"])

with tab1:
    uploaded_pdfs = st.file_uploader(
        "Upload Survey Sheet PDFs", 
        type="pdf", 
        accept_multiple_files=True,
        help="Supports multiple PDFs - will create separate sheets for each"
    )

    if uploaded_pdfs:
        st.success(f"✅ Loaded {len(uploaded_pdfs)} file(s)")
        
        file_name_prefix = st.text_input(
            "📂 Output file prefix", 
            value="WCS_Edited",
            help="Name for Excel/ZIP files (no extension needed)"
        )

        per_file_data = []
        pdf_results = []
        used_names = set()
        all_unique = True

        for i, uploaded_pdf in enumerate(uploaded_pdfs, start=1):
            with st.expander(f"📄 File {i}: {uploaded_pdf.name}", expanded=True):
                col1, col2 = st.columns([3, 1])
                with col1:
                    custom_pdf_name = st.text_input(
                        f"Output PDF name",
                        value=f"{uploaded_pdf.name.replace('.pdf','')}_edited",
                        key=f"rename_{i}",
                        help="Unique name for this PDF output"
                    )
                with col2:
                    if st.button(f"🔄 Re-extract", key=f"reextract_{i}"):
                        st.cache_data.clear()
                        st.rerun()

                if custom_pdf_name in used_names:
                    st.error(f"❌ Duplicate name '{custom_pdf_name}'")
                    all_unique = False
                    continue
                used_names.add(custom_pdf_name)

                # Extract & Display Data
                sales_data = extract_sales_blocks(uploaded_pdf)
                if not sales_data:
                    st.warning("⚠️ No sales lines found")
                    continue

                df = pd.DataFrame(sales_data)

                # Conditional column config with highlighting
                def width_color(row):
                    if pd.isna(row['width']) or row['order_width'] is None:
                        return ""
                    diff = abs(row['order_width'] - row['width'])
                    return FENESTA_RED if diff > 75 else FENESTA_LIGHT_BLUE

                def height_color(row):
                    if pd.isna(row['height']) or row['order_height'] is None:
                        return ""
                    diff = abs(row['order_height'] - row['height'])
                    return FENESTA_RED if diff > 75 else FENESTA_LIGHT_BLUE

                edited_df = st.data_editor(
                    df,
                    num_rows="fixed",
                    width='stretch',
                    hide_index=True,
                    key=f"editor_{i}",
                    column_config={
                        "sales_line": st.column_config.TextColumn("Sales Line", disabled=True, width="small"),
                        "order_width": st.column_config.NumberColumn("Order W", disabled=True, width="small", format="%.0f mm"),
                        "order_height": st.column_config.NumberColumn("Order H", disabled=True, width="small", format="%.0f mm"),
                        "reference": st.column_config.TextColumn("Reference", disabled=True),
                        "location": st.column_config.TextColumn("Location", disabled=True),
                        "system": st.column_config.TextColumn("System", disabled=True, width="small"),
                        "width": st.column_config.NumberColumn(
                            "Input Width", 
                            step=1, 
                            format="%.0f mm",
                            cell_background_color=width_color
                        ),
                        "height": st.column_config.NumberColumn(
                            "Input Height", 
                            step=1, 
                            format="%.0f mm",
                            cell_background_color=height_color
                        ),
                        "location_input": st.column_config.TextColumn("Location (Input)"),
                        "remarks": st.column_config.TextColumn("Remarks")
                    }
                )

                # Show mismatch summary
                edited_df['w_diff'] = abs(edited_df['order_width'] - edited_df['width'].fillna(0))
                edited_df['h_diff'] = abs(edited_df['order_height'] - edited_df['height'].fillna(0))
                mismatches = ((edited_df['w_diff'] > 75) | (edited_df['h_diff'] > 75)).sum()
                
                col1, col2 = st.columns(2)
                col1.metric("🔴 Mismatches (>75mm)", mismatches, delta=None)
                col2.metric("✅ Matches", len(edited_df) - mismatches, delta=None)

                sheet_name = make_excel_safe_name(custom_pdf_name)
                per_file_data.append((sheet_name, edited_df))
                
                uploaded_pdf.seek(0)
                edited_pdf = update_pdf(uploaded_pdf.read(), edited_df.to_dict("records"))
                pdf_results.append((custom_pdf_name, edited_pdf))

        # Download Section
        if all_unique and per_file_data:
            st.subheader("📥 Download Options")
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("📦 Download All (ZIP)", type="primary", use_container_width=True):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zf:
                        excel_file = io.BytesIO()
                        with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
                            for sheet_name, df_part in per_file_data:
                                df_part.to_excel(writer, index=False, sheet_name=sheet_name)
                        excel_file.seek(0)
                        zf.writestr(f"{file_name_prefix}.xlsx", excel_file.getvalue())
                        
                        for pdf_name, pdf_bytes in pdf_results:
                            zf.writestr(f"{pdf_name}.pdf", pdf_bytes)
                    
                    zip_buffer.seek(0)
                    today_str = datetime.today().strftime("%Y-%m-%d")
                    st.download_button(
                        "⬇️ ZIP Complete Package",
                        zip_buffer.getvalue(),
                        f"{file_name_prefix}_{today_str}.zip",
                        "application/zip"
                    )
            
            with col2:
                # Individual downloads
                for pdf_name, pdf_bytes in pdf_results:
                    st.download_button(
                        f"📄 {pdf_name}.pdf",
                        pdf_bytes,
                        f"{pdf_name}.pdf",
                        "application/pdf"
                    )

with tab2:
    st.markdown("""
    ## 🎯 How to Use WCS Editor Pro
    
    ### 1. **Upload PDFs**
    - Upload one or more Survey Sheet PDFs
    - Each file gets its own editable sheet
    
    ### 2. **Smart Editing** 
    - **🔴 Red cells** = Difference >75mm from order sizes
    - **🔵 Light blue** = Within tolerance 
    - Edit **Width**, **Height**, **Location**, **Remarks**
    
    ### 3. **Auto-Validation**
    - Real-time mismatch counter
    - PDF annotations match editor colors
    
    ### 4. **Export Options**
    - 📦 **ZIP**: Excel (multi-sheet) + All PDFs
    - 📄 **Individual PDFs** for selective download
    
    **Pro Tip**: Use unique PDF names to avoid conflicts!
    """)

st.markdown("---")
st.markdown(f"**© Fenesta Building Systems** | Powered by Streamlit | v27", unsafe_allow_html=True)
