# ==========================
# 📝 WCS Editor (v28 - Fixed for Streamlit 1.52.1)
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
            "w_status": "➖",  # Neutral
            "h_status": "➖"   # Neutral
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

def update_status_indicators(df):
    """Update status indicators for width/height columns"""
    df_copy = df.copy()
    
    for idx, row in df_copy.iterrows():
        # Width status
        if pd.isna(row['width']) or row['order_width'] is None:
            df_copy.at[idx, 'w_status'] = "➖"
        else:
            diff_w = abs(row['order_width'] - row['width'])
            df_copy.at[idx, 'w_status'] = "🔴" if diff_w > 75 else "✅"
        
        # Height status
        if pd.isna(row['height']) or row['order_height'] is None:
            df_copy.at[idx, 'h_status'] = "➖"
        else:
            diff_h = abs(row['order_height'] - row['height'])
            df_copy.at[idx, 'h_status'] = "🔴" if diff_h > 75 else "✅"
    
    return df_copy

def make_excel_safe_name(name):
    safe = "".join(c if c.isalnum() else "_" for c in name)[:31]
    return safe if safe else "Sheet"

# ----------------- Enhanced UI -----------------
st.set_page_config(page_title="WCS Editor", layout="wide", page_icon="📝")

# Header
st.markdown("""
