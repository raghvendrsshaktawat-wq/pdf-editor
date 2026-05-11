# ==========================
# Fenesta WCS Survey Editor v4.2
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
.stApp{background:#f0f4f9}

/* ── Header ── */
.fen-header{
  background:linear-gradient(135deg,#005BAC 0%,#0047a0 100%);
  padding:18px 32px;border-radius:16px;
  display:flex;align-items:center;justify-content:space-between;
  margin-bottom:24px;box-shadow:0 4px 20px rgba(0,91,172,.25)}
.brand{color:#ffffff;font-size:24px;font-weight:800;letter-spacing:-.5px}
.brand span{color:#E8212E}
.tagline{color:#a8c8f0;font-size:13px;margin-top:3px}
.badge{background:#E8212E;color:white;border-radius:8px;padding:6px 14px;font-size:12px;font-weight:700}

/* ── Metric cards ── */
.metric-row{display:flex;gap:12px;margin:16px 0;flex-wrap:wrap}
.metric-card{background:white;border-radius:12px;padding:14px 20px;flex:1;min-width:130px;
  box-shadow:0 2px 8px rgba(0,91,172,.08);border-left:4px solid #005BAC}
.metric-card.ok    {border-left-color:#22c55e}
.metric-card.warn  {border-left-color:#f59e0b}
.metric-card.danger{border-left-color:#E8212E}
.mk{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#94a3b8}
.mv{font-size:26px;font-weight:800;color:#005BAC;margin-top:2px}
.ms{font-size:12px;color:#94a3b8}

/* ── Section title ── */
.section-title{font-size:15px;font-weight:700;color:#005BAC;
  border-left:4px solid #E8212E;padding-left:10px;margin:20px 0 10px}

/* ── Legend ── */
.legend{display:flex;gap:16px;align-items:center;background:white;border-radius:10px;
  padding:10px 16px;box-shadow:0 1px 4px rgba(0,91,172,.08);margin-bottom:16px;flex-wrap:wrap}
.legend-item{display:flex;align-items:center;gap:6px;font-size:13px;font-weight:500}
.dot{width:12px;height:12px;border-radius:50%}
.dot-ok    {background:#22c55e}
.dot-warn  {background:#f59e0b}
.dot-danger{background:#E8212E}
.dot-empty {background:#cbd5e1}

/* ── File uploader ── */
[data-testid="stFileUploader"]{background:white!important;border-radius:14px!important;
  padding:8px!important;box-shadow:0 2px 8px rgba(0,91,172,.08)!important}

/* ── Buttons ── */
.stDownloadButton>button,.stButton>button{
  background:linear-gradient(135deg,#005BAC,#0047a0)!important;
  color:white!important;border:none!important;border-radius:10px!important;
  font-weight:700!important;box-shadow:0 2px 8px rgba(0,91,172,.3)!important;
  transition:all .2s!important}
.stDownloadButton>button:hover,.stButton>button:hover{
  transform:translateY(-1px)!important;
  box-shadow:0 4px 14px rgba(0,91,172,.4)!important}

/* ── Expander ── */
[data-testid="stExpander"]{background:white!important;border-radius:14px!important;
  border:1px solid #dde8f5!important;box-shadow:0 2px 8px rgba(0,91,172,.06)!important;
  margin-bottom:14px!important}

/* ── Metric value color override ── */
[data-testid="stMetricValue"]{color:#005BAC!important;font-weight:800!important}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="fen-header">
  <div>
    <div class="brand">🪟 Fenesta <span>WCS</span> Survey Editor</div>
    <div class="tagline">Fenesta Building Systems · Better by Design · Survey Data Overlay Tool</div>
  </div>
  <div class="badge">v4.2</div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="legend">
  <span style="font-size:13px;font-weight:600;color:#005BAC">Tolerance Guide:</span>
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
    all_lines = []
    for page in doc:
        for ln in page.get_text("text").splitlines():
            all_lines.append(ln)
    return _parse(all_lines)

def _parse(all_lines):
    lines = [l.rstrip() for l in all_lines]

    meta = {}
    for ln in lines:
        m = re.search(r'\b(W\d{7,})\b', ln)
        if m and not meta.get('order_no'): meta['order_no'] = m.group(1)
    for ln in lines:
        m = re.match(r'^(9\.\d+e\+\d+)$', ln.strip())
        if m:
            try: meta['msc_no'] = str(int(float(m.group(1))))
            except: pass
            break
    for i, ln in enumerate(lines):
        if i > 0 and re.match(r'^\d+/\d+/\d+$', lines[i-1].strip()) and re.match(r'^[A-Z]{4,}$', ln.strip()):
            meta['zone'] = ln.strip(); break
    for i, ln in enumerate(lines):
        if ln.strip() == 'INDIA' and i+1 < len(lines):
            cand = lines[i+1].strip()
            if re.match(r'^[A-Z][A-Z ]{2,}$', cand) and 'ZAHEERABAD' not in cand:
                meta['customer'] = cand; break

    SKIP_DESC = re.compile(
        r'^(RE Remarks|Customer Remarks|Configuration Changed|Remarks|Sales Line|'
        r'Size \(w x h\)|Glazing|Description|Qty|^N$|^Y$|Cust\. Initials|'
        r'Order No\.|Zone|Quote No\.|Date|MSC No\.|Print Date|Segment|Retail|'
        r'Viewed from Inside|Window Position|Mechanical Join|Aperture Finish|'
        r'Opening|Orientation|Grill|B/S|Sash Handle).*$', re.I)

    BOILERPLATE_VALS = re.compile(
        r'^(Foil 2S|Walnut|Feature|118mm|65mm|Fibre|^Yes$|Sleek|^Mechanical$|^Black$|'
        r'^Plaster$|Easy Clean|^Brick$|^Center$|Luxury|Fixed New|A65)', re.I)

    sales_line_re = re.compile(r'^0\d{3}$')
    glazing_re    = re.compile(r'^(SG|DG|TG|Louver)\s*', re.I)
    rows = []

    for i, raw in enumerate(lines):
        ln = raw.strip()
        if not sales_line_re.match(ln): continue
        if i+1 >= len(lines) or lines[i+1].strip() != '1': continue
        if i+3 >= len(lines): continue
        h_str = lines[i+2].strip()
        w_str = lines[i+3].strip()
        if not (re.match(r'^\d{3,4}$', h_str) and re.match(r'^\d{3,4}$', w_str)): continue

        order_height = int(h_str)
        order_width  = int(w_str)

        glazing = ''
        if i+4 < len(lines) and glazing_re.match(lines[i+4].strip()):
            glazing = lines[i+4].strip()

        desc = ''
        for back in range(i-1, max(i-6, -1), -1):
            c = lines[back].strip()
            if not c: continue
            if re.match(r'^\d{2,4}$', c): continue
            if SKIP_DESC.match(c): continue
            desc = c; break

        ref = loc = sys = ''
        arch_idx = None
        for j in range(i+1, min(i+120, len(lines))):
            if 'Arch Height(mm)' in lines[j]:
                arch_idx = j; break

        if arch_idx is not None:
            cands = []
            for j in range(arch_idx+1, min(arch_idx+12, len(lines))):
                s = lines[j]; stripped = s.strip()
                if not stripped: continue
                if s.startswith('  ') and stripped:
                    if BOILERPLATE_VALS.match(stripped): break
                    cands.append(stripped)
                else: break
            if len(cands) >= 1: ref = cands[0]
            if len(cands) >= 2: loc = cands[1]
            if len(cands) >= 3: sys = cands[2]

        rows.append({
            'sales_line':    ln,
            'description':   desc,
            'system':        sys,
            'order_width':   order_width,
            'order_height':  order_height,
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
def get_tol(ov, sv):
    if sv is None or pd.isna(sv): return 'empty'
    d = abs(float(ov) - float(sv))
    if d <= 75:  return 'ok'
    if d <= 200: return 'warn'
    return 'danger'

def row_tol(ow, oh, sw, sh):
    tw, th = get_tol(ow, sw), get_tol(oh, sh)
    if 'empty'  in (tw, th): return 'empty'
    if 'danger' in (tw, th): return 'danger'
    if 'warn'   in (tw, th): return 'warn'
    return 'ok'

# ══════════════════════════════════════
# PDF OVERLAY — pixel-perfect cell alignment
# Cell X: 78.25 → 459.65 (measured from actual WCS PDF drawing paths)
# ══════════════════════════════════════
CELL_X0  = 78.25
CELL_X1  = 459.65
CELL_PAD = 4.0

def overlay_pdf(pdf_bytes, rows_data, surveyor_name=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    if surveyor_name:
        for page in doc:
            hits = page.search_for("Name")
            if len(hits) >= 2:
                r = hits[1]; tx, ty = r.x1+8, r.y1-2
                page.draw_rect(fitz.Rect(tx-2, ty-12, tx+160, ty+4), color=(1,1,1), fill=(1,1,1), width=0)
                page.insert_text((tx, ty), surveyor_name, fontsize=10, fontname="helv", color=(0,0,0))
                break

    col_map = {
        'ok':     (0.0, 0.6, 0.2),
        'warn':   (0.8, 0.5, 0.0),
        'danger': (0.91, 0.13, 0.18),   # Fenesta Red #E8212E
        'empty':  (0.0, 0.36, 0.67)     # Fenesta Blue #005BAC
    }

    cell_list = []
    for pn in range(len(doc)):
        pg = doc[pn]
        ap_hits = pg.search_for("Aperture Size")
        drawings = pg.get_drawings()
        h_lines = sorted(
            [p["rect"] for p in drawings
             if p["rect"] and abs(p["rect"].y1 - p["rect"].y0) < 3
             and p["rect"].x1 - p["rect"].x0 > 100],
            key=lambda r: r.y0
        )
        for hit in ap_hits:
            above = [r for r in h_lines if r.y1 <= hit.y0 + 2]
            below = [r for r in h_lines if r.y0 >= hit.y1 - 2]
            top_y = above[-1].y0 if above else hit.y0 - 13
            bot_y = below[0].y0  if below  else hit.y1 + 3
            cell_list.append((pn, top_y, bot_y))

    for idx, row in enumerate(rows_data):
        if idx >= len(cell_list): break
        pn, top_y, bot_y = cell_list[idx]
        page = doc[pn]

        sw = row.get('survey_width');   sh = row.get('survey_height')
        ow = row.get('order_width', 0); oh = row.get('order_height', 0)
        room    = (row.get('room')    or '').strip()
        remarks = (row.get('remarks') or '').strip()
        tol = row_tol(ow, oh, sw, sh)
        col = col_map[tol]

        if sw is not None and sh is not None: size_txt = f"{int(sw)} x {int(sh)}"
        elif sw is not None:                  size_txt = f"{int(sw)} x --"
        elif sh is not None:                  size_txt = f"-- x {int(sh)}"
        else:                                 size_txt = "Not surveyed"

        line1 = f"{room} : {size_txt}" if room else size_txt
        row_h = bot_y - top_y
        fs    = min(9.0, row_h * 0.62)

        # ── Aperture Size row ──
        cell = fitz.Rect(CELL_X0+0.5, top_y+0.5, CELL_X1-0.5, bot_y-0.5)
        page.draw_rect(cell, color=col, fill=(1,1,1), width=1.5)
        page.insert_text((CELL_X0+CELL_PAD, top_y+row_h*0.73),
                         line1[:72], fontsize=fs, fontname="helv", color=col)

        # ── Production Size row — remarks ──
        if remarks:
            pt = bot_y; pb = bot_y + row_h
            pc = fitz.Rect(CELL_X0+0.5, pt+0.5, CELL_X1-0.5, pb-0.5)
            page.draw_rect(pc, color=col, fill=(1,1,1), width=1.5)
            page.insert_text((CELL_X0+CELL_PAD, pt+row_h*0.73),
                             remarks[:72], fontsize=fs, fontname="helv", color=col)

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
                border:2px dashed #dde8f5;margin-top:20px">
      <div style="font-size:48px">📄</div>
      <div style="font-size:18px;font-weight:700;color:#005BAC;margin:12px 0 6px">
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
            st.error("⚠ No sales line items detected. Ensure this is a text-based Fenesta WCS PDF.")
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
        n_danger = tols.count('danger'); n_empty = tols.count('empty'); n_total = len(tols)
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
st.caption("🪟 Fenesta Building Systems · Better by Design · 🔴 >200mm Critical | 🟡 76–200mm Review | 🟢 ≤75mm OK")
