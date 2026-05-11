# ==========================
# Fenesta WCS Survey Editor
# ==========================
import streamlit as st
import fitz
import pandas as pd
import io
import re
from datetime import datetime
import numpy as np

st.set_page_config(
    page_title="Fenesta WCS Survey Editor",
    page_icon="🪟",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ── Custom CSS ──
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Fenesta brand colors: Orange #F47920, Dark Navy #1A1A2E, White */
.stApp { background: #f8f6f2; }

/* Top header bar */
.fen-header {
    background: linear-gradient(135deg, #1A1A2E 0%, #16213E 100%);
    padding: 18px 32px;
    border-radius: 16px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 24px;
    box-shadow: 0 4px 20px rgba(26,26,46,0.18);
}
.fen-header .brand { color: #F47920; font-size: 22px; font-weight: 800; letter-spacing: -0.5px; }
.fen-header .tagline { color: #94a3b8; font-size: 13px; margin-top: 2px; }
.fen-header .badge { background: #F47920; color: white; border-radius: 8px; padding: 6px 14px; font-size: 12px; font-weight: 700; }

/* Metric cards */
.metric-row { display: flex; gap: 12px; margin: 16px 0; flex-wrap: wrap; }
.metric-card {
    background: white;
    border-radius: 12px;
    padding: 14px 20px;
    flex: 1;
    min-width: 130px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    border-left: 4px solid #F47920;
}
.metric-card.ok { border-left-color: #22c55e; }
.metric-card.warn { border-left-color: #f59e0b; }
.metric-card.danger { border-left-color: #ef4444; }
.metric-card .mk { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: #94a3b8; }
.metric-card .mv { font-size: 26px; font-weight: 800; color: #1A1A2E; margin-top: 2px; }
.metric-card .ms { font-size: 12px; color: #94a3b8; }

/* Section headers */
.section-title {
    font-size: 15px; font-weight: 700; color: #1A1A2E;
    border-left: 4px solid #F47920;
    padding-left: 10px;
    margin: 20px 0 10px;
}

/* Tolerance legend */
.legend {
    display: flex; gap: 16px; align-items: center;
    background: white; border-radius: 10px; padding: 10px 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06); margin-bottom: 16px;
    flex-wrap: wrap;
}
.legend-item { display: flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 500; }
.dot { width: 12px; height: 12px; border-radius: 50%; }
.dot-ok { background: #22c55e; }
.dot-warn { background: #f59e0b; }
.dot-danger { background: #ef4444; }
.dot-empty { background: #cbd5e1; }

/* Upload area */
[data-testid="stFileUploader"] {
    background: white !important;
    border-radius: 14px !important;
    padding: 8px !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06) !important;
}

/* Buttons */
.stDownloadButton > button, .stButton > button {
    background: linear-gradient(135deg, #F47920, #e8671a) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    padding: 10px 20px !important;
    box-shadow: 0 2px 8px rgba(244,121,32,0.3) !important;
    transition: all 0.2s !important;
}
.stDownloadButton > button:hover, .stButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 14px rgba(244,121,32,0.4) !important;
}

/* Data editor styling */
[data-testid="stDataEditor"] { border-radius: 12px !important; overflow: hidden !important; box-shadow: 0 2px 12px rgba(0,0,0,0.08) !important; }

/* Expander */
[data-testid="stExpander"] {
    background: white !important;
    border-radius: 14px !important;
    border: 1px solid #e2e8f0 !important;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important;
    margin-bottom: 14px !important;
}

/* Inputs */
[data-testid="stTextInput"] input, [data-testid="stNumberInput"] input {
    border-radius: 8px !important;
    border: 1.5px solid #e2e8f0 !important;
}
[data-testid="stTextInput"] input:focus, [data-testid="stNumberInput"] input:focus {
    border-color: #F47920 !important;
    box-shadow: 0 0 0 3px rgba(244,121,32,0.1) !important;
}

/* Divider */
hr { border-color: #e2e8f0 !important; }

/* Tooltip pill */
.tol-pill {
    display: inline-block; border-radius: 999px; padding: 2px 10px;
    font-size: 11px; font-weight: 700;
}
.tol-ok    { background: #dcfce7; color: #166534; }
.tol-warn  { background: #fef9c3; color: #854d0e; }
.tol-crit  { background: #fee2e2; color: #991b1b; }
.tol-empty { background: #f1f5f9; color: #64748b; }
</style>
""", unsafe_allow_html=True)

# ── Header ──
st.markdown("""
<div class="fen-header">
  <div>
    <div class="brand">🪟 Fenesta WCS Survey Editor</div>
    <div class="tagline">Fenesta Building Systems · Survey Data Overlay Tool</div>
  </div>
  <div class="badge">v2.0</div>
</div>
""", unsafe_allow_html=True)

# ── Tolerance legend ──
st.markdown("""
<div class="legend">
  <span style="font-size:13px;font-weight:600;color:#1A1A2E;">Tolerance Guide:</span>
  <span class="legend-item"><span class="dot dot-ok"></span> ≤75 mm — Within tolerance</span>
  <span class="legend-item"><span class="dot dot-warn"></span> 76–200 mm — Review required</span>
  <span class="legend-item"><span class="dot dot-danger"></span> &gt;200 mm — Critical — Re-survey</span>
  <span class="legend-item"><span class="dot dot-empty"></span> Not surveyed yet</span>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════
# PARSER (PyMuPDF)
# ══════════════════════════════════════
@st.cache_data
def parse_wcs_pdf(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    full_text = ""
    for page in doc:
        full_text += page.get_text("text") + "\n"

    meta = {}
    m = re.search(r'\b(W\d{7,})\b', full_text)
    if m: meta['order_no'] = m.group(1)
    m = re.search(r'\b(9\d{9,})\b', full_text)
    if m: meta['msc_no'] = m.group(1)
    m = re.search(r'(HYDERABAD|DELHI|GURGAON|BENGALURU|MUMBAI|CHENNAI|PUNE|KOLKATA)', full_text)
    if m: meta['zone'] = m.group(1)
    m = re.search(r'(?i)customer\s*[:\-]?\s*\n?([A-Z][A-Z ]{3,})\n', full_text)
    if m: meta['customer'] = m.group(1).strip()

    lines = full_text.splitlines()
    rows = []

    # Primary pattern: 4-digit line, qty on next line, then sizes
    i = 0
    while i < len(lines):
        ln = lines[i].strip()
        # Match sales line like "0001" or "0001 1 Bay 2 Facet 3550 2145 SG..."
        sl_match = re.match(r'^(0\d{3})$', ln)
        inline_match = re.match(r'^(0\d{3})\s+(\d)\s+(.+?)\s{2,}(\d{3,4})\s+(\d{3,4})\s+(.+)$', ln)

        if inline_match:
            row = {
                'sales_line': inline_match.group(1),
                'qty': inline_match.group(2),
                'description': inline_match.group(3).strip(),
                'order_width': int(inline_match.group(4)),
                'order_height': int(inline_match.group(5)),
                'glazing': inline_match.group(6).strip(),
                'reference': '', 'location': '', 'system': '',
                'survey_width': None, 'survey_height': None,
                'room': '', 'remarks': ''
            }
            # look ahead for Reference/Location/System
            for j in range(i+1, min(i+20, len(lines))):
                lj = lines[j].strip()
                if re.match(r'^0\d{3}$', lj): break
                if lj.lower() == 'reference' and j+1 < len(lines): row['reference'] = lines[j+1].strip()
                if lj.lower() == 'location'  and j+1 < len(lines): row['location']  = lines[j+1].strip()
                if lj.lower() == 'system'    and j+1 < len(lines): row['system']    = lines[j+1].strip()
                m2 = re.match(r'(?i)^Reference\s+([A-Z0-9\/]+)$', lj)
                if m2: row['reference'] = m2.group(1)
                m2 = re.match(r'(?i)^Location\s+([A-Z0-9\/]+)$', lj)
                if m2: row['location'] = m2.group(1)
                m2 = re.match(r'(?i)^System\s+(SY\d+.+)$', lj)
                if m2: row['system'] = m2.group(1).strip()
            rows.append(row)
            i += 1
            continue

        if sl_match:
            sales_line = sl_match.group(1)
            # collect next few lines to reconstruct the block
            block = lines[i:min(i+12, len(lines))]
            nums = [x.strip() for x in block if re.match(r'^\d{3,4}$', x.strip())]
            descs = [x.strip() for x in block if x.strip() and not re.match(r'^\d', x.strip()) and len(x.strip()) > 1]
            qty = ''
            for x in block[1:4]:
                if re.match(r'^\d$', x.strip()): qty = x.strip(); break
            row = {
                'sales_line': sales_line,
                'qty': qty,
                'description': descs[0] if descs else '',
                'order_width': int(nums[0]) if len(nums) > 0 else 0,
                'order_height': int(nums[1]) if len(nums) > 1 else 0,
                'glazing': '',
                'reference': '', 'location': '', 'system': '',
                'survey_width': None, 'survey_height': None,
                'room': '', 'remarks': ''
            }
            for j in range(i+1, min(i+20, len(lines))):
                lj = lines[j].strip()
                if re.match(r'^0\d{3}$', lj): break
                if lj.lower() == 'reference' and j+1 < len(lines): row['reference'] = lines[j+1].strip()
                if lj.lower() == 'location'  and j+1 < len(lines): row['location']  = lines[j+1].strip()
                if lj.lower() == 'system'    and j+1 < len(lines): row['system']    = lines[j+1].strip()
                gl = re.match(r'^(SG|DG|TG|Louver)\s+.+', lj, re.I)
                if gl and not row['glazing']: row['glazing'] = lj
            rows.append(row)
            i += 1
            continue
        i += 1

    return meta, rows

# ══════════════════════════════════════
# TOLERANCE HELPER
# ══════════════════════════════════════
def get_tolerance(order_val, survey_val):
    if survey_val is None or pd.isna(survey_val): return 'empty'
    diff = abs(float(order_val) - float(survey_val))
    if diff <= 75: return 'ok'
    if diff <= 200: return 'warn'
    return 'danger'

def tolerance_label(order_w, order_h, sw, sh):
    tw = get_tolerance(order_w, sw)
    th = get_tolerance(order_h, sh)
    if tw == 'empty' or th == 'empty': return 'empty'
    if tw == 'danger' or th == 'danger': return 'danger'
    if tw == 'warn' or th == 'warn': return 'warn'
    return 'ok'

# ══════════════════════════════════════
# PDF OVERLAY
# ══════════════════════════════════════
def overlay_survey_on_pdf(pdf_bytes, rows_data, surveyor_name=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    # Write surveyor name
    if surveyor_name:
        for page in doc:
            hits = page.search_for("Name")
            if len(hits) >= 2:
                r = hits[1]
                tx, ty = r.x1 + 8, r.y1 - 2
                page.draw_rect(fitz.Rect(tx-2, ty-12, tx+160, ty+4), color=(1,1,1), fill=(1,1,1), width=0)
                page.insert_text((tx, ty), surveyor_name, fontsize=10, fontname="helv", color=(0,0,0))
                break

    # Find Aperture Size anchors (one per sales line page block)
    aperture_anchors = []
    for pn, page in enumerate(doc):
        for inst in page.search_for("Aperture Size"):
            aperture_anchors.append((pn, inst))

    color_map = {'ok': (0.05, 0.6, 0.25), 'warn': (0.8, 0.5, 0.0), 'danger': (0.85, 0.1, 0.1), 'empty': (0.5, 0.5, 0.5)}
    fs = 10
    gap = int(fs * 1.8)

    for idx, row in enumerate(rows_data):
        if idx >= len(aperture_anchors): break
        pn, anchor = aperture_anchors[idx]
        page = doc[pn]

        sw = row.get('survey_width')
        sh = row.get('survey_height')
        ow = row.get('order_width', 0)
        oh = row.get('order_height', 0)
        room = (row.get('room') or '').strip()
        remarks = (row.get('remarks') or '').strip()

        tol = tolerance_label(ow, oh, sw, sh)
        col = color_map[tol]

        if sw is not None and sh is not None:
            size_txt = f"{int(sw)} x {int(sh)}"
        elif sw is not None:
            size_txt = f"{int(sw)} x --"
        elif sh is not None:
            size_txt = f"-- x {int(sh)}"
        else:
            size_txt = "Not surveyed"

        line1 = f"{room} : {size_txt}" if room else size_txt
        ix = anchor.x1 + 30
        iy = anchor.y0 + 2

        def write_line(x, y, text, c):
            bw, bh = 320, fs * 1.65
            rect = fitz.Rect(x-3, y-fs*1.25, x-3+bw, y-fs*1.25+bh)
            page.draw_rect(rect, color=c, fill=None, width=1.5)
            page.draw_rect(fitz.Rect(rect.x0+1.5, rect.y0+1.5, rect.x1-1.5, rect.y1-1.5), color=(1,1,1), fill=(1,1,1), width=0)
            page.insert_text((x, y), text[:60], fontsize=fs, fontname="helv", color=c)

        write_line(ix, iy, line1, col)
        if remarks:
            write_line(ix, iy + gap, remarks, col)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════
c1, c2 = st.columns([3, 1])
with c1:
    uploaded_pdfs = st.file_uploader(
        "📎 Upload Fenesta WCS PDF Reports",
        type="pdf",
        accept_multiple_files=True,
        help="Upload one or more Fenesta WCS system-generated PDF reports"
    )
with c2:
    lot_name = st.text_input("Lot / Project Name", placeholder="e.g. Zaheerabad Phase 1")
    surveyor_global = st.text_input("Surveyor Name", placeholder="Enter surveyor name")

if not uploaded_pdfs:
    st.markdown("""
    <div style="text-align:center;padding:60px 20px;background:white;border-radius:16px;border:2px dashed #e2e8f0;margin-top:20px">
      <div style="font-size:48px">📄</div>
      <div style="font-size:18px;font-weight:700;color:#1A1A2E;margin:12px 0 6px">Upload WCS PDF Reports to Begin</div>
      <div style="color:#94a3b8;font-size:14px">Supports multiple Fenesta WCS PDFs · Auto-extracts all sales line items</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

all_file_data = []
pdf_outputs = []

# ── Per-file metrics for top summary
total_items_all = 0
surveyed_all = 0
ok_all = 0
warn_all = 0
danger_all = 0

for i, pdf_file in enumerate(uploaded_pdfs, 1):
    file_bytes = pdf_file.read()
    with st.expander(f"📄  {pdf_file.name}", expanded=(i == 1)):
        with st.spinner(f"Parsing {pdf_file.name}..."):
            meta, rows = parse_wcs_pdf(file_bytes)

        if not rows:
            st.error("⚠ No sales line items detected. This WCS may be image-based/scanned.")
            continue

        # Per-file header info
        hc1, hc2, hc3 = st.columns(3)
        hc1.markdown(f"**Order:** `{meta.get('order_no','—')}`")
        hc2.markdown(f"**MSC:** `{meta.get('msc_no','—')}`")
        hc3.markdown(f"**Zone:** `{meta.get('zone','—')}`")

        df = pd.DataFrame(rows)
        df['survey_width']  = df['survey_width'].astype(object).where(df['survey_width'].notna(), None)
        df['survey_height'] = df['survey_height'].astype(object).where(df['survey_height'].notna(), None)

        st.markdown('<div class="section-title">Line Item Editor</div>', unsafe_allow_html=True)

        edited = st.data_editor(
            df,
            num_rows="fixed",
            hide_index=True,
            use_container_width=True,
            key=f"editor_{i}",
            column_config={
                "sales_line":   st.column_config.TextColumn("Sales Line",   disabled=True, width="small"),
                "qty":          st.column_config.TextColumn("Qty",          disabled=True, width="small"),
                "description":  st.column_config.TextColumn("Description",  disabled=True, width="medium"),
                "system":       st.column_config.TextColumn("System",       disabled=True, width="medium"),
                "order_width":  st.column_config.NumberColumn("Order W",    disabled=True, width="small", format="%d mm"),
                "order_height": st.column_config.NumberColumn("Order H",    disabled=True, width="small", format="%d mm"),
                "reference":    st.column_config.TextColumn("Reference",    disabled=True, width="small"),
                "location":     st.column_config.TextColumn("Location",     disabled=True, width="small"),
                "glazing":      st.column_config.TextColumn("Glazing",      disabled=True, width="medium"),
                "survey_width": st.column_config.NumberColumn("Survey W ✎", width="small", step=1, format="%d mm"),
                "survey_height":st.column_config.NumberColumn("Survey H ✎", width="small", step=1, format="%d mm"),
                "room":         st.column_config.TextColumn("Room ✎",       width="medium"),
                "remarks":      st.column_config.TextColumn("Remarks ✎",    width="large"),
            }
        )

        # ── Inline tolerance summary ──
        tols = [tolerance_label(r['order_width'], r['order_height'], r['survey_width'], r['survey_height']) for _, r in edited.iterrows()]
        n_ok     = tols.count('ok')
        n_warn   = tols.count('warn')
        n_danger = tols.count('danger')
        n_empty  = tols.count('empty')
        n_total  = len(tols)
        n_done   = n_ok + n_warn + n_danger

        ok_all     += n_ok
        warn_all   += n_warn
        danger_all += n_danger
        total_items_all += n_total
        surveyed_all    += n_done

        st.markdown(f"""
        <div class="metric-row">
          <div class="metric-card"><div class="mk">Total Items</div><div class="mv">{n_total}</div><div class="ms">in this WCS</div></div>
          <div class="metric-card ok"><div class="mk">✅ Within Tolerance</div><div class="mv">{n_ok}</div><div class="ms">≤ 75 mm diff</div></div>
          <div class="metric-card warn"><div class="mk">⚠️ Review Required</div><div class="mv">{n_warn}</div><div class="ms">76–200 mm diff</div></div>
          <div class="metric-card danger"><div class="mk">🔴 Critical</div><div class="mv">{n_danger}</div><div class="ms">&gt; 200 mm diff</div></div>
          <div class="metric-card"><div class="mk">Not Surveyed</div><div class="mv">{n_empty}</div><div class="ms">pending</div></div>
        </div>
        """, unsafe_allow_html=True)

        # PDF overlay
        overlaid = overlay_survey_on_pdf(file_bytes, edited.to_dict('records'), surveyor_name=surveyor_global)
        base = pdf_file.name.replace('.pdf','')
        out_name = f"{lot_name}_{base}.pdf" if lot_name else f"{base}_surveyed.pdf"
        pdf_outputs.append((out_name, overlaid))
        all_file_data.append((meta.get('order_no', pdf_file.name), edited))

        st.download_button(
            f"⬇ Download {out_name}",
            data=overlaid,
            file_name=out_name,
            mime="application/pdf",
            key=f"dl_pdf_{i}"
        )

# ══════════════════════════════════════
# GLOBAL SUMMARY + EXCEL EXPORT
# ══════════════════════════════════════
if all_file_data:
    st.markdown("---")
    st.markdown('<div class="section-title">📊 Overall Survey Progress</div>', unsafe_allow_html=True)
    pct = round(surveyed_all / total_items_all * 100) if total_items_all else 0
    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="mk">Total Line Items</div><div class="mv">{total_items_all}</div><div class="ms">across all WCS</div></div>
      <div class="metric-card ok"><div class="mk">✅ Within Tolerance</div><div class="mv">{ok_all}</div></div>
      <div class="metric-card warn"><div class="mk">⚠️ Review Required</div><div class="mv">{warn_all}</div></div>
      <div class="metric-card danger"><div class="mk">🔴 Critical</div><div class="mv">{danger_all}</div></div>
      <div class="metric-card"><div class="mk">Completion</div><div class="mv">{pct}%</div><div class="ms">{surveyed_all} of {total_items_all} surveyed</div></div>
    </div>
    """, unsafe_allow_html=True)

    # Progress bar
    st.progress(pct / 100, text=f"Survey completion: {pct}%")

    # Excel export
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        for sheet_label, df_part in all_file_data:
            safe = "".join(c if c.isalnum() else "_" for c in str(sheet_label))[:31] or "Sheet"
            df_part.to_excel(writer, index=False, sheet_name=safe)
    ts = datetime.now().strftime('%Y%m%d_%H%M')
    xl_name = f"{lot_name}_{ts}.xlsx" if lot_name else f"WCS_Survey_{ts}.xlsx"

    st.download_button(
        "⬇ Download Combined Excel (All WCS)",
        data=excel_buf.getvalue(),
        file_name=xl_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.markdown("---")
st.caption("🪟 Fenesta Building Systems · WCS Survey Editor · 🔴 >200mm Critical &nbsp;|&nbsp; 🟡 76–200mm Review &nbsp;|&nbsp; 🟢 ≤75mm OK")
