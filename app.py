fontname = get_fontname_for_page(page)
anchor_x = inst.x1         # right edge of "Aperture Size"
anchor_y = inst.y          # baseline of "Aperture Size"

line1_text = f"{location_input} : {size_text}"

# Decide colors based on dimension difference (as you already do)
text_color = (0, 0, 1)      # blue
border_color = (0, 0, 1)
if survey_w is not None and survey_h is not None:
    if (order_w is not None and abs(order_w - survey_w) > 75) or \
       (order_h is not None and abs(order_h - survey_h) > 75):
        text_color = (1, 0, 0)
        border_color = (1, 0, 0)

# --- Adjustable settings ---
BOX_WIDTH = 280          # change this
BOX_HEIGHT_MULT = 1.7    # change this
BORDER_WIDTH = 2.0       # change this
OFFSET_X = 40            # shift right from anchor
OFFSET_Y = 0             # shift up/down from anchor baseline
LINE_SPACING = int(fontsize * 1.6)

# Line 1
draw_text_box(
    page,
    (anchor_x, anchor_y),
    line1_text,
    fontname=fontname,
    fontsize=font_size,
    text_color=text_color,
    border_color=border_color,
    border_width=BORDER_WIDTH,
    box_width=BOX_WIDTH,
    box_height_mult=BOX_HEIGHT_MULT,
    offset_x=OFFSET_X,
    offset_y=OFFSET_Y,
)

# Line 2 (remarks), same box settings but lower Y
if remarks_text:
    draw_text_box(
        page,
        (anchor_x, anchor_y + LINE_SPACING),
        remarks_text,
        fontname=fontname,
        fontsize=font_size,
        text_color=text_color,
        border_color=border_color,
        border_width=BORDER_WIDTH,
        box_width=BOX_WIDTH,
        box_height_mult=BOX_HEIGHT_MULT,
        offset_x=OFFSET_X,
        offset_y=OFFSET_Y,
    )
