# ==========================
# Fenesta WCS Survey Editor v3.1
# ==========================
import streamlit as st
import fitz
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="Fenesta WCS Survey Editor", page_icon="🪟", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif}
.stApp{background:#f8f6f2}
.fen-header{background:linear-gradient(135deg,#1A1A2E 0%,#16213E 100%);padding:18px 32px;border-radius:16px;display:flex;align-items:center;justify-content:space-between;margin-bottom:24px;box-shadow:0 4px 20px rgba(26,26,46,.18)}
.brand{color:#F47920;font-size:22px;font-weight:800;letter-spacing:-.5px}
.tagline{color:#94a3b8;font-size:13px;margin-top:2px}
.badge{background:#F47920;color:white;border-radius:8px;padding:6px 14px;font-size:12px;font-weight:700}
.metric-row{display:flex;gap:12px;margin:16px 0;flex-wrap:wrap}
.metric-card{background:white;border-radius:12px;padding:14px 20px;flex:1;min-width:130px;box-shadow:0 2px 8px rgba(0,0,0,.06);border-left:4px solid #F47920}
.metric-card.ok{border-left-color:#22c55e}.metric-card.warn{border-left-color:#f59e0b}.metric-card.danger{border-left-color:#ef4444}
.mk{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#94a3b8}
.mv{font-size:26px;font-weight:800;color:#1A1A2E;margin-top:2px}.ms{font-size:12px;color:#94a3b8}
.section-title{font-size:15px;font-weight:700;color:#1A1A2E;border-left:4px solid #F47920;padding-left:10px;margin:20px 0 10px}
.legend{display:flex;gap:16px;align-items:center;background:white;border-radius:10px;padding:10px 16px;box-shadow:0 1px 4px rgba(0,0,0,.06);margin-bottom:16px;flex-wrap:wrap}
.legend-item{display:flex;align-items:center;gap:6px;font-size:13px;font-weight:500}
.dot{width:12px;height:12px;border-radius:50%}
.dot-ok{background:#22c55e}.dot-warn{background:#f59e0b}.dot-danger{background:#ef4444}.dot-empty{background:#cbd5e1}
[data-testid="stFileUploader"]{background:white!important;border-radius:14px!important;padding:8px!important;box-shadow:0 2px 8px rgba(0,0,0,.06)!important}
.stDownloadButton>button,.stButton>button{background:linear-gradient(135deg,#F47920,#e8671a)!important;color:white!important;border:none!important;border-radius:10px!important;font-weight:700!important;box-shadow:0 2px 8px rgba(244,121,32,.3)!important}
[data-testid="stExpander"]{background:white!important;border-radius:14px!important;border:1px solid #e2e8f0!important;box-shadow:0 2px 8px rgba(0,0,0,.05)!important;margin-bottom:14px!important}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="fen-header">
  <div><div class="brand">🪟 Fenesta WCS Survey Editor</div><div class="tagline">Fenesta Building Systems · Survey Data Overlay Tool</div></div>
  <div class="badge">v3.1</div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="legend">
  <span style="font-size:13px;font-weight:600;color:#1A1A2E">Tolerance Guide:</span>
  <span class="legend-item"><span class="dot dot-ok"></span> ≤75 mm — Within tolerance</span>
  <span class="legend-item"><span class="dot dot-warn"></span> 76–200 mm — Review required</span>
  <span class="legend-item"><span class="dot dot-danger"></span> &gt;200 mm — Critical — Re-survey</span>
  <span class="legend-item"><span class="dot dot-empty"></span> Not surveyed yet</span>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════
# PARSER
# ══════════════════════════════════════
@st.cache_data
def parse_wcs_pdf(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    full_text = ""
    for page in doc:
        full_text += page.get_text("text") + "\n"
    return parse_wcs_lines(full_text)

def parse_wcs_lines(full_text):
    lines = [l.rstrip() for l in full_text.splitlines()]

    # ── Meta ──
    meta = {}
    for ln in lines:
        m = re.search(r'\b(W\d{7,})\b', ln)
        if m and not meta.get('order_no'): meta['order_no'] = m.group(1)
    for ln in lines:
        m = re.search(r'\b(9\d{9})\b', ln)
        if m and 'Tel' not in ln: meta['msc_no'] = m.group(1); break
    m = re.search(r'\d{10}\s+[\d/]+\s+([A-Z]{4,})\s+\w+', full_text)
    if m: meta['zone'] = m.group(1)
    m = re.search(r'INDIA\n([A-Z][A-Z ]+)\n', full_text)
    if m: meta['customer'] = m.group(1).strip()

    # ── Sales line patterns ──
    # Pattern A: 0NNN 1 <text description> WIDTH HEIGHT GLAZING
    patA = re.compile(r'^(0\d{3})\s+1\s+([A-Za-z].+?)\s+(\d{3,4})\s+(\d{3,4})\s+((?:SG|DG|TG|Louver)[\w\s_]+?)\s*$')
    # Pattern B: 0NNN 1 WIDTH HEIGHT GLAZING (description on previous lines)
    patB = re.compile(r'^(0\d{3})\s+1\s+(\d{3,4})\s+(\d{3,4})\s+((?:SG|DG|TG|Louver)[\w\s_]+?)\s*$')

    sl_hits = []
    for i, raw_ln in enumerate(lines):
        ln = raw_ln.strip()
        mA = patA.match(ln)
        mB = patB.match(ln)
        if mA:
            sl_hits.append({'line_idx': i, 'sales_line': mA.group(1),
                            'inline_desc': mA.group(2).strip(),
                            'order_width': int(mA.group(3)), 'order_height': int(mA.group(4)),
                            'glazing': mA.group(5).strip()})
        elif mB:
            sl_hits.append({'line_idx': i, 'sales_line': mB.group(1),
                            'inline_desc': None,
                            'order_width': int(mB.group(2)), 'order_height': int(mB.group(3)),
                            'glazing': mB.group(4).strip()})

    SKIP = re.compile(r'^(RE Remarks|Customer Remarks|Configuration Changed|Remarks|Sales Line|'
                      r'Order No\.|SURVEY CHECKLIST|INSTALLATION).*$', re.I)

    rows = []
    for hit in sl_hits:
        i = hit['line_idx']

        # ── Description ──
        if hit['inline_desc']:
            description = hit['inline_desc']
        else:
            desc = ''
            for back in range(i - 1, max(i - 8, 0), -1):
                candidate = lines[back].strip()
                if not candidate: continue
                if re.match(r'^\d{2,4}$', candidate): continue
                if SKIP.match(candidate): continue
                desc = candidate
                break
            description = desc

        # ── Reference / Location / System ──
        ref = loc = sys = ''
        arch_idx = None
        for j in range(i + 1, min(i + 90, len(lines))):
            if 'Arch Height(mm)' in lines[j]:
                arch_idx = j
                break

        if arch_idx is not None:
            candidates = []
            stop_words = re.compile(
                r'^(Reference|Location|System|Colour|Foil|Corner|Frame|Sash|Bug|Handle|'
                r'Aperture|Hinge|Masonry|CW |BS |T-Join|Louver|Facet|Coupling)', re.I)
            for j in range(arch_idx + 1, min(arch_idx + 20, len(lines))):
                stripped = lines[j].strip()
                if not stripped: continue
                if stop_words.match(stripped): break
                candidates.append(stripped)
            if len(candidates) >= 1: ref = candidates[0]
            if len(candidates) >= 2: loc = candidates[1]
            if len(candidates) >= 3: sys = candidates[2]

        rows.append({
            'sales_line':    hit['sales_line'],
            'description':   description,
            'system':        sys,
            'order_width':   hit['order_width'],
            'order_height':  hit['order_height'],
            'reference':     ref,
            'location':      loc,
            'survey_width':  None,
            'survey_height': None,
            'room':          '',
            'remarks':       ''
        })

    return meta, rows

# ══════════════════════════════════════
# TOLERANCE
# ══════════════════════════════════════
def get_tol(order_val, survey_val):
    if survey_val is None or pd.isna(survey_val): return 'empty'
    diff = abs(float(order_val) - float(survey_val))
    if diff <= 75:  return 'ok'
    if diff <= 200: return 'warn'
    return 'danger'

def row_tol(ow, oh, sw, sh):
    tw = get_tol(ow, sw); th = get_tol(oh, sh)
    if 'empty'  in (tw, th): return 'empty'
    if 'danger' in (tw, th): return 'danger'
    if 'warn'   in (tw, th): return 'warn'
    return 'ok'

# ══════════════════════════════════════
# PDF OVERLAY
# ══════════════════════════════════════
def overlay_pdf(pdf_bytes, rows_data, surveyor_name=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    if surveyor_name:
        for page in doc:
            hits = page.search_for("Name")
            if len(hits) >= 2:
                r = hits[1]; tx, ty = r.x1 + 8, r.y1 - 2
                page.draw_rect(fitz.Rect(tx-2, ty-12, tx+160, ty+4), color=(1,1,1), fill=(1,1,1), width=0)
                page.insert_text((tx, ty), surveyor_name, fontsize=10, fontname="helv", color=(0,0,0))
                break

    anchors = []
    for pn, page in enumerate(doc):
        for inst in page.search_for("Aperture Size"):
            anchors.append((pn, inst))

    col_map = {'ok':(0.05,.6,.25), 'warn':(.8,.5,0), 'danger':(.85,.1,.1), 'empty':(.5,.5,.5)}
    fs, gap = 10, 18

    for idx, row in enumerate(rows_data):
        if idx >= len(anchors): break
        pn, anchor = anchors[idx]
        page = doc[pn]
        sw = row.get('survey_width'); sh = row.get('survey_height')
        ow = row.get('order_width', 0); oh = row.get('order_height', 0)
        room    = (row.get('room') or '').strip()
        remarks = (row.get('remarks') or '').strip()
        tol = row_tol(ow, oh, sw, sh)
        col = col_map[tol]

        if sw is not None and sh is not None:   size_txt = f"{int(sw)} x {int(sh)}"
        elif sw is not None:                    size_txt = f"{int(sw)} x --"
        elif sh is not None:                    size_txt = f"-- x {int(sh)}"
        else:                                   size_txt = "Not surveyed"

        line1 = f"{room} : {size_txt}" if room else size_txt
        ix, iy = anchor.x1 + 30, anchor.y0 + 2

        def write(x, y, txt, c):
            bw, bh = 320, fs * 1.65
            rect = fitz.Rect(x-3, y-fs*1.25, x-3+bw, y-fs*1.25+bh)
            page.draw_rect(rect, color=c, fill=None, width=1.5)
            page.draw_rect(fitz.Rect(rect.x0+1.5, rect.y0+1.5, rect.x1-1.5, rect.y1-1.5),
                           color=(1,1,1), fill=(1,1,1), width=0)
            page.insert_text((x, y), txt[:60], fontsize=fs, fontname="helv", color=c)

        write(ix, iy, line1, col)
        if remarks: write(ix, iy + gap, remarks, col)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════
c1, c2, c3 = st.columns([3, 1, 1])
with c1:
    uploaded_pdfs = st.file_uploader("📎 Upload Fenesta WCS PDF Reports", type="pdf", accept_multiple_files=True)
with c2:
    lot_name = st.text_input("Lot / Project Name", placeholder="e.g. Zaheerabad Phase 1")
with c3:
    surveyor_global = st.text_input("Surveyor Name", placeholder="Surveyor name")

if not uploaded_pdfs:
    st.markdown("""
    <div style="text-align:center;padding:60px 20px;background:white;border-radius:16px;
                border:2px dashed #e2e8f0;margin-top:20px">
      <div style="font-size:48px">📄</div>
      <div style="font-size:18px;font-weight:700;color:#1A1A2E;margin:12px 0 6px">
        Upload WCS PDF Reports to Begin</div>
      <div style="color:#94a3b8;font-size:14px">
        Supports multiple Fenesta WCS PDFs · Auto-extracts all sales line items</div>
    </div>""", unsafe_allow_html=True)
    st.stop()

all_file_data = []
total_all = surveyed_all = ok_all = warn_all = danger_all = 0

for i, pdf_file in enumerate(uploaded_pdfs, 1):
    file_bytes = pdf_file.read()
    with st.expander(f"📄  {pdf_file.name}", expanded=(i == 1)):
        with st.spinner(f"Parsing {pdf_file.name}..."):
            meta, rows = parse_wcs_pdf(file_bytes)

        if not rows:
            st.error("⚠ No sales line items detected in this WCS PDF.")
            continue

        hc1, hc2, hc3, hc4 = st.columns(4)
        hc1.metric("Order No.", meta.get('order_no', '—'))
        hc2.metric("MSC No.",   meta.get('msc_no',   '—'))
        hc3.metric("Zone",      meta.get('zone',     '—'))
        hc4.metric("Customer",  meta.get('customer', '—'))

        df = pd.DataFrame(rows)
        st.markdown('<div class="section-title">Line Item Editor</div>', unsafe_allow_html=True)

        edited = st.data_editor(
            df, num_rows="fixed", hide_index=True,
            use_container_width=True, key=f"editor_{i}",
            column_config={
                "sales_line":    st.column_config.TextColumn("Sales Line",   disabled=True, width="small"),
                "description":   st.column_config.TextColumn("Description",  disabled=True, width="medium"),
                "system":        st.column_config.TextColumn("System",       disabled=True, width="large"),
                "order_width":   st.column_config.NumberColumn("Order W",    disabled=True, width="small", format="%d mm"),
                "order_height":  st.column_config.NumberColumn("Order H",    disabled=True, width="small", format="%d mm"),
                "reference":     st.column_config.TextColumn("Reference",    disabled=True, width="small"),
                "location":      st.column_config.TextColumn("Location",     disabled=True, width="small"),
                "survey_width":  st.column_config.NumberColumn("Survey W ✎", width="small", step=1, format="%d mm"),
                "survey_height": st.column_config.NumberColumn("Survey H ✎", width="small", step=1, format="%d mm"),
                "room":          st.column_config.TextColumn("Room ✎",       width="medium"),
                "remarks":       st.column_config.TextColumn("Remarks ✎",    width="large"),
            }
        )

        tols = [row_tol(r['order_width'], r['order_height'], r['survey_width'], r['survey_height'])
                for _, r in edited.iterrows()]
        n_ok = tols.count('ok'); n_warn = tols.count('warn')
        n_danger = tols.count('danger'); n_empty = tols.count('empty')
        n_total = len(tols)
        ok_all += n_ok; warn_all += n_warn; danger_all += n_danger
        total_all += n_total; surveyed_all += n_ok + n_warn + n_danger

        st.markdown(f"""
        <div class="metric-row">
          <div class="metric-card"><div class="mk">Total Items</div><div class="mv">{n_total}</div></div>
          <div class="metric-card ok"><div class="mk">✅ Within Tolerance</div><div class="mv">{n_ok}</div><div class="ms">≤ 75 mm</div></div>
          <div class="metric-card warn"><div class="mk">⚠️ Review Required</div><div class="mv">{n_warn}</div><div class="ms">76–200 mm</div></div>
          <div class="metric-card danger"><div class="mk">🔴 Critical</div><div class="mv">{n_danger}</div><div class="ms">&gt; 200 mm</div></div>
          <div class="metric-card"><div class="mk">Pending</div><div class="mv">{n_empty}</div><div class="ms">not filled</div></div>
        </div>""", unsafe_allow_html=True)

        overlaid = overlay_pdf(file_bytes, edited.to_dict('records'), surveyor_name=surveyor_global)
        base = pdf_file.name.replace('.pdf', '')
        out_name = f"{lot_name}_{base}.pdf" if lot_name else f"{base}_surveyed.pdf"
        all_file_data.append((meta.get('order_no', base), edited))

        st.download_button(f"⬇ Download {out_name}", data=overlaid,
                           file_name=out_name, mime="application/pdf", key=f"dl_{i}")

if all_file_data:
    st.markdown("---")
    st.markdown('<div class="section-title">📊 Overall Survey Progress</div>', unsafe_allow_html=True)
    pct = round(surveyed_all / total_all * 100) if total_all else 0
    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="mk">All Line Items</div><div class="mv">{total_all}</div></div>
      <div class="metric-card ok"><div class="mk">✅ Within Tolerance</div><div class="mv">{ok_all}</div></div>
      <div class="metric-card warn"><div class="mk">⚠️ Review</div><div class="mv">{warn_all}</div></div>
      <div class="metric-card danger"><div class="mk">🔴 Critical</div><div class="mv">{danger_all}</div></div>
      <div class="metric-card"><div class="mk">Completion</div><div class="mv">{pct}%</div></div>
    </div>""", unsafe_allow_html=True)
    st.progress(pct / 100, text=f"Survey completion: {pct}%")

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        for label, df_part in all_file_data:
            safe = "".join(c if c.isalnum() else "_" for c in str(label))[:31] or "Sheet"
            df_part.to_excel(writer, index=False, sheet_name=safe)
    ts = datetime.now().strftime('%Y%m%d_%H%M')
    xl_name = f"{lot_name}_{ts}.xlsx" if lot_name else f"WCS_Survey_{ts}.xlsx"
    st.download_button("⬇ Download Combined Excel (All WCS)", data=excel_buf.getvalue(),
                       file_name=xl_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

st.markdown("---")
st.caption("🪟 Fenesta Building Systems · 🔴 >200mm Critical | 🟡 76–200mm Review | 🟢 ≤75mm OK")
