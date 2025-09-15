# ==========================
# 📝 Raghvendra's PDF Editor (v26 - Add direct download buttons)
# ==========================

import streamlit as st
import fitz
import pandas as pd
import io
import re
import zipfile
from datetime import datetime

# Regex: sales line = 0xxx, qty = 1, then height + width
pattern = re.compile(
    r"^\s*(0\d{3})\s*?\n"
    r"^\s*1\s*?\n"
    r"^\s*(\d{2,4})\s*?\n"
    r"^\s*(\d{2,4})\s*?\n",
    re.MULTILINE
)

def extract_sales_blocks(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text("text") + "\n"

    lines = text.splitlines()
    blocks = []

    for match in pattern.finditer(text):
        sales_line = match.group(1)
        order_height = int(match.group(2))  # swapped
        order_width = int(match.group(3))   # swapped

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
    font_size = 14   # locked font size

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

        color = (0, 0, 1)
        if order_w and input_w and abs(order_w - input_w) > 75:
            color = (1, 0, 0)
        if order_h and input_h and abs(order_h - input_h) > 75:
            color = (1, 0, 0)

        size_text = f"{input_w} x {input_h}"
        location_text = f"({entry['location_input']})"
        remarks_text = f"{entry['remarks']}" if entry.get("remarks") else ""

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


# ----------------- Streamlit UI -----------------
st.title("📝 Raghvendra's PDF Editor")

uploaded_pdfs = st.file_uploader("Upload one or more Survey Sheet PDFs", type="pdf", accept_multiple_files=True)

if uploaded_pdfs:
    file_name_prefix = st.text_input("📂 Excel/ZIP file name prefix (without extension)", value="output")

    per_file_data = []
    pdf_results = []
    used_names = set()
    all_unique = True

    # Loop over each uploaded PDF
    for i, uploaded_pdf in enumerate(uploaded_pdfs, start=1):
        st.write(f"---")
        st.write(f"### 📄 File {i}: {uploaded_pdf.name}")

        # User renames output PDF
        custom_pdf_name = st.text_input(
            f"Rename output PDF for file {i}",
            value=f"{uploaded_pdf.name.replace('.pdf','')}_edited",
            key=f"rename_{i}"
        )

        if custom_pdf_name in used_names:
            st.error(f"❌ The name '{custom_pdf_name}' is already used. Please choose a unique name.")
            all_unique = False
        else:
            used_names.add(custom_pdf_name)

            # Extract sales data
            sales_data = extract_sales_blocks(uploaded_pdf)
            if len(sales_data) == 0:
                st.warning(f"⚠️ No sales lines found in {uploaded_pdf.name}")
                continue

            df = pd.DataFrame(sales_data)[[
                "sales_line","order_width","order_height","reference","location","system",
                "width","height","location_input","remarks"
            ]]

            edited_df = st.data_editor(
                df,
                num_rows="fixed",
                use_container_width=True,
                hide_index=True,   # hide index
                key=f"editor_{i}",
                column_config={
                    "sales_line": st.column_config.TextColumn(disabled=True),
                    "order_width": st.column_config.NumberColumn(disabled=True),
                    "order_height": st.column_config.NumberColumn(disabled=True),
                    "reference": st.column_config.TextColumn(disabled=True),
                    "location": st.column_config.TextColumn(disabled=True),
                    "system": st.column_config.TextColumn(disabled=True),
                    "width": st.column_config.NumberColumn("Width (input)", step=1),
                    "height": st.column_config.NumberColumn("Height (input)", step=1),
                    "location_input": st.column_config.TextColumn("Location (input)"),
                    "remarks": st.column_config.TextColumn("Remarks")
                }
            )

            sheet_name = make_excel_safe_name(custom_pdf_name)
            per_file_data.append((sheet_name, edited_df))

            uploaded_pdf.seek(0)
            edited_pdf = update_pdf(
                uploaded_pdf.read(),
                edited_df.to_dict("records")
            )
            pdf_results.append((custom_pdf_name, edited_pdf))

    # ========== DOWNLOAD SECTION ==========

    # 1. ZIP download
    if all_unique and st.button("📦 Download All Results (ZIP)"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            # Excel with multiple sheets
            if per_file_data:
                excel_file = io.BytesIO()
                with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
                    for sheet_name, df_part in per_file_data:
                        df_part.to_excel(writer, index=False, sheet_name=sheet_name)
                excel_file.seek(0)
                zf.writestr(f"{file_name_prefix}.xlsx", excel_file.getvalue())
            # PDFs
            for pdf_name, pdf_bytes in pdf_results:
                zf.writestr(f"{pdf_name}.pdf", pdf_bytes)

        zip_buffer.seek(0)
        today_str = datetime.today().strftime("%Y-%m-%d")
        zip_name = f"{file_name_prefix}_{today_str}_all.zip"
        st.success("✅ ZIP file ready for download")
        st.download_button("⬇️ Download All Files (ZIP)",
                           data=zip_buffer, file_name=zip_name, mime="application/zip")

    # 2. Direct download section
    if all_unique and per_file_data:
        st.write("---")
        st.subheader("📥 Direct Downloads (No ZIP)")

        # Excel direct
        excel_file = io.BytesIO()
        with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
            for sheet_name, df_part in per_file_data:
                df_part.to_excel(writer, index=False, sheet_name=sheet_name)
        excel_file.seek(0)
        st.download_button("⬇️ Download Excel Only",
                           data=excel_file, file_name=f"{file_name_prefix}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDFs direct
        for pdf_name, pdf_bytes in pdf_results:
            st.download_button(f"⬇️ Download {pdf_name}.pdf",
                               data=pdf_bytes, file_name=f"{pdf_name}.pdf", mime="application/pdf")
