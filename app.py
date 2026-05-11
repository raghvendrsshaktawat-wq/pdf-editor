# ══════════════════════════════════════
# PDF OVERLAY — aligned to actual table cell lines
# ══════════════════════════════════════
CELL_X0 = 78.25   # left edge of Aperture/Production Size cells (constant across all WCS pages)
CELL_X1 = 459.65  # right edge
PADDING = 4.0

def overlay_pdf(pdf_bytes, rows_data, surveyor_name=""):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    if surveyor_name:
        for page in doc:
            hits = page.search_for("Name")
            if len(hits) >= 2:
                r = hits[1]; tx, ty = r.x1+8, r.y1-2
                page.draw_rect(fitz.Rect(tx-2,ty-12,tx+160,ty+4), color=(1,1,1), fill=(1,1,1), width=0)
                page.insert_text((tx,ty), surveyor_name, fontsize=10, fontname="helv", color=(0,0,0))
                break

    col_map = {'ok':(0.05,.6,.25), 'warn':(.8,.5,0), 'danger':(.85,.1,.1), 'empty':(.4,.4,.4)}

    # Build precise cell list by finding h-lines bracketing each "Aperture Size" anchor
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
            bot_y = below[0].y0 if below else hit.y1 + 3
            cell_list.append((pn, top_y, bot_y))

    for idx, row in enumerate(rows_data):
        if idx >= len(cell_list): break
        pn, top_y, bot_y = cell_list[idx]
        page = doc[pn]
        sw = row.get('survey_width'); sh = row.get('survey_height')
        ow = row.get('order_width', 0); oh = row.get('order_height', 0)
        room    = (row.get('room') or '').strip()
        remarks = (row.get('remarks') or '').strip()
        tol = row_tol(ow, oh, sw, sh)
        col = col_map[tol]

        if sw is not None and sh is not None: size_txt = f"{int(sw)} x {int(sh)}"
        elif sw is not None:                  size_txt = f"{int(sw)} x --"
        elif sh is not None:                  size_txt = f"-- x {int(sh)}"
        else:                                 size_txt = "Not surveyed"

        line1 = f"{room} : {size_txt}" if room else size_txt
        row_h = bot_y - top_y
        fs = min(9.0, row_h * 0.62)

        # ── Aperture Size cell — draw border + text exactly inside the cell ──
        cell = fitz.Rect(CELL_X0 + 0.5, top_y + 0.5, CELL_X1 - 0.5, bot_y - 0.5)
        page.draw_rect(cell, color=col, fill=(1, 1, 1), width=1.5)
        page.insert_text(
            (CELL_X0 + PADDING, top_y + row_h * 0.73),
            line1[:72], fontsize=fs, fontname="helv", color=col
        )

        # ── Production Size cell — remarks ──
        if remarks:
            prod_top = bot_y
            prod_bot = bot_y + row_h
            prod_cell = fitz.Rect(CELL_X0 + 0.5, prod_top + 0.5, CELL_X1 - 0.5, prod_bot - 0.5)
            page.draw_rect(prod_cell, color=col, fill=(1, 1, 1), width=1.5)
            page.insert_text(
                (CELL_X0 + PADDING, prod_top + row_h * 0.73),
                remarks[:72], fontsize=fs, fontname="helv", color=col
            )

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
