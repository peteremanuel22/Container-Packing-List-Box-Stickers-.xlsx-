
import io
from datetime import date, datetime
import openpyxl
import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

st.set_page_config(page_title="Container Packing List → Box Stickers", layout="wide")

# ======================== Parsing helpers ========================

def find_first_sheet(wb):
    for ws in wb.worksheets:
        if ws.sheet_state == "visible":
            return ws
    return wb.active

def find_header_row(ws, header_candidates):
    """
    Detect header row in the uploaded In.xlsx by matching known header labels.
    Returns (header_row_idx, col_map) where col_map maps logical keys → 1-based column index.

    NOTE (forced columns by index):
      - Box code       -> column B (index 2)
      - Component code -> column E (index 5)
      - Box type       -> column G (index 7, Arabic text)
    """
    header_row_idx = None
    best_match_count = -1
    best_col_map = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        texts = [str(c.value).strip() if c.value is not None else "" for c in row]
        lower_texts = [t.lower() for t in texts]
        match_count = 0
        col_map = {}
        for key, variants in header_candidates.items():
            found_col = None
            for v in variants:
                v_low = v.lower()
                for idx, cell_text in enumerate(lower_texts):
                    if v_low and v_low in cell_text and texts[idx] != "":
                        found_col = idx + 1
                        break
                if found_col:
                    break
            if found_col:
                match_count += 1
                col_map[key] = found_col
        if match_count > best_match_count and match_count >= 3:
            best_match_count = match_count
            header_row_idx = row[0].row
            best_col_map = col_map
    return header_row_idx, best_col_map

def read_rows(ws, header_row, col_map):
    """
    Read rows under the header. Stop on 2 consecutive empty rows.

    Fields read:
      - sn        : from header mapping
      - code_box  : from column B (index 2) **forced**
      - comp_ar   : from header mapping
      - comp_en   : from header mapping
      - comp_code : from column E (index 5) **forced**
      - qty       : from header mapping
      - box_type  : from column G (index 7) **forced**
    """
    rows = []
    r = header_row + 1
    empties = 0
    while r <= ws.max_row:
        row_vals = {}

        # Header-mapped fields
        for k in ["sn", "comp_ar", "comp_en", "qty"]:
            c = col_map.get(k)
            v = ws.cell(row=r, column=c).value if c else None
            row_vals[k] = v

        # Forced columns
        row_vals["code_box"]  = ws.cell(row=r, column=2).value  # B
        row_vals["comp_code"] = ws.cell(row=r, column=5).value  # E
        row_vals["box_type"]  = ws.cell(row=r, column=7).value  # G

        # Emptiness check across core fields
        core = [row_vals.get(k) for k in ["sn","code_box","comp_ar","comp_en","comp_code","qty","box_type"]]
        all_empty = all(v in (None, "") for v in core)
        if all_empty:
            empties += 1
            if empties >= 2:
                break
        else:
            empties = 0
            rows.append(row_vals)

        r += 1
    return rows

def _text(v):
    return "" if v is None else str(v).strip()

def _sn(v):
    s = _text(v)
    return None if s == "" else s

def _code_box(v):
    s = _text(v)
    return None if s == "" else s

def group_boxes(data_rows):
    """
    Group components into boxes.
    Rule: A component WITHOUT S.N OR WITHOUT Box code follows the previous component in the same box.
    Box identity = S.N + Box code. A row with both starts a new box.
    """
    groups = []
    current_box = None
    for r in data_rows:
        sn = _sn(r.get("sn"))
        code_box = _code_box(r.get("code_box"))  # from column B
        comp_ar = _text(r.get("comp_ar"))
        comp_en = _text(r.get("comp_en"))
        comp_code = _text(r.get("comp_code"))    # from column E
        qty = r.get("qty")
        box_type = _text(r.get("box_type"))      # from column G (Arabic)

        # New box when both present
        if sn and code_box:
            if current_box is not None:
                groups.append(current_box)
            current_box = {
                "sn": sn,
                "code_box": code_box,
                "box_type": box_type,
                "items": []
            }

        # Attach to current box
        if current_box is None:
            current_box = {
                "sn": "(UNKNOWN)",
                "code_box": "(UNKNOWN)",
                "box_type": box_type,
                "items": []
            }

        current_box["items"].append({
            "comp_ar": comp_ar,
            "comp_en": comp_en,
            "comp_code": comp_code,
            "qty": qty
        })

        if box_type:
            current_box["box_type"] = box_type

    if current_box is not None:
        groups.append(current_box)

    return groups

# ======================== Styling helpers ========================

B_THIN = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrapText=True)
ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center", wrapText=True)
ALIGN_RIGHT  = Alignment(horizontal="right",  vertical="center", wrapText=True)

TITLE_FONT = Font(bold=True, size=12)
LABEL_FONT = Font(bold=True, size=11)
TEXT_FONT  = Font(size=10)

HEADER_FILL       = PatternFill("solid", fgColor="D9E1F2")   # light blue
TABLE_HEADER_FILL = PatternFill("solid", fgColor="F2F2F2")   # light gray

ROW_HEIGHT_FROM_TO_PT = 64.5  # 86 px
ROW_HEIGHT_LABELS_PT  = 15.0  # 20 px

COL_A_WIDTH_UNITS = 12.2      # ≈ 90 px

# ======================== Sticker layout ========================

def set_default_widths(ws):
    ws.column_dimensions["A"].width = COL_A_WIDTH_UNITS  # ~90 px
    width_map = {
        "B": 20, "C": 20, "D": 20, "E": 20, "F": 20, "G": 16,
        "H": 16, "I": 16, "J": 16
    }
    for col, w in width_map.items():
        ws.column_dimensions[col].width = w

def draw_header(ws, r0, c0, title):
    """
    Header band merges A..G (previously A..F).
    """
    ws.merge_cells(start_row=r0, start_column=c0, end_row=r0, end_column=c0+6)  # A..G
    cell = ws.cell(row=r0, column=c0)
    cell.value = title
    cell.font = TITLE_FONT
    cell.alignment = ALIGN_CENTER
    cell.fill = HEADER_FILL
    cell.border = B_THIN
    for c in range(c0, c0+7):  # apply border across A..G
        ws.cell(row=r0, column=c).border = B_THIN
    return r0 + 1

def merge_value_B_to_G(ws, r, value, set_label_height=False):
    """
    Merge and write 'value' across columns B..G on row r.
    Adds borders and left alignment with wrap.
    Optionally set the row height for single-line label rows (20 px ≈ 15 pt).
    """
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)  # B..G
    vcell = ws.cell(row=r, column=2)
    vcell.value = value
    vcell.font = TEXT_FONT
    vcell.alignment = ALIGN_LEFT
    vcell.border = B_THIN
    for cc in range(2, 8):  # B..G inclusive
        ws.cell(row=r, column=cc).border = B_THIN
    if set_label_height:
        ws.row_dimensions[r].height = ROW_HEIGHT_LABELS_PT

def draw_label(ws, r, c_label, label, set_label_height=False):
    lab = ws.cell(row=r, column=c_label)
    lab.value = label
    lab.font = LABEL_FONT
    lab.alignment = ALIGN_RIGHT
    lab.border = B_THIN
    if set_label_height:
        ws.row_dimensions[r].height = ROW_HEIGHT_LABELS_PT

def draw_multiline_value(ws, r, label, value):
    draw_label(ws, r, 1, label, set_label_height=False)       # column A
    merge_value_B_to_G(ws, r, value, set_label_height=False)  # columns B..G
    ws.row_dimensions[r].height = ROW_HEIGHT_FROM_TO_PT
    return r + 1

def draw_table_header(ws, r, c0):
    """
    Components table header (7 columns):
      A: Box S.N   (merged vertically per box)
      B: Box code  (merged vertically per box)
      C: Component (Arabic)
      D: Component (English)
      E: Code   (from column E of In.xlsx)
      F: Qty
      G: Box type (Arabic from column G of In.xlsx)
    """
    headers = ["Box S.N", "Box code", "Component (Arabic)", "Component (English)", "Code", "Qty", "Box type"]
    widths  = [12, 14, 28, 28, 16, 8, 12]
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(64 + c0 + i)].width = w
    for i, h in enumerate(headers):
        cell = ws.cell(row=r, column=c0 + i)
        cell.value = h
        cell.font = LABEL_FONT
        cell.fill = TABLE_HEADER_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = B_THIN
    return r + 1

def draw_components_table_with_merged_sn_and_code(ws, r_start, c0, box):
    """
    Draw the table header + component rows.
    Merge column A with Box S.N and column B with Box code, vertically across the component rows.
    Returns the first free row after the table.
    """
    r = draw_table_header(ws, r_start, c0)
    first_comp_row = r
    items = box["items"]

    # Write rows; leave A (Box S.N) and B (Box code) blank for now
    for it in items:
        values = [
            None,                        # A (Box S.N merged later)
            None,                        # B (Box code merged later)
            it.get("comp_ar", ""),       # C
            it.get("comp_en", ""),       # D
            it.get("comp_code", ""),     # E (from column E in In.xlsx)
            it.get("qty", ""),           # F
            box.get("box_type", "")      # G (Arabic from column G)
        ]
        aligns = [ALIGN_CENTER, ALIGN_CENTER, ALIGN_LEFT, ALIGN_LEFT, ALIGN_CENTER, ALIGN_CENTER, ALIGN_CENTER]

        for i, val in enumerate(values):
            cell = ws.cell(row=r, column=c0 + i)
            if val is not None:
                cell.value = val
            cell.font = TEXT_FONT
            if i in (2, 3):  # wrap Arabic/English name columns
                cell.alignment = Alignment(wrapText=True, horizontal="left", vertical="center")
            else:
                cell.alignment = aligns[i]
            cell.border = B_THIN

        ws.row_dimensions[r].height = 18
        r += 1

    # Merge A (Box S.N) and B (Box code)
    if items:
        last_comp_row = r - 1
        if last_comp_row >= first_comp_row:
            # A: Box S.N
            ws.merge_cells(start_row=first_comp_row, start_column=c0, end_row=last_comp_row, end_column=c0)
            cell_sn = ws.cell(row=first_comp_row, column=c0)
            cell_sn.value = box.get("sn", "")
            cell_sn.font = TEXT_FONT
            cell_sn.alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
            cell_sn.border = B_THIN

            # B: Box code
            ws.merge_cells(start_row=first_comp_row, start_column=c0+1, end_row=last_comp_row, end_column=c0+1)
            cell_code = ws.cell(row=first_comp_row, column=c0+1)
            cell_code.value = box.get("code_box", "")
            cell_code.font = TEXT_FONT
            cell_code.alignment = Alignment(horizontal="center", vertical="center", wrapText=True)
            cell_code.border = B_THIN

    return r

def draw_sticker(ws, top_row, box, inputs):
    r = top_row
    c = 1  # start at column A

    # Header (A..G merged)
    r = draw_header(ws, r, c, "Packing List — Box Sticker")

    # Single-line wide rows (20 px; merged B..G)
    draw_label(ws, r, c, "Packing List No.", set_label_height=True)
    merge_value_B_to_G(ws, r, inputs["packing_list_no"], set_label_height=True)
    r += 1

    draw_label(ws, r, c, "Order No.", set_label_height=True)
    merge_value_B_to_G(ws, r, inputs["order_no"], set_label_height=True)
    r += 1

    draw_label(ws, r, c, "Date of Shipment", set_label_height=True)
    merge_value_B_to_G(ws, r, inputs["date_str"], set_label_height=True)
    r += 1

    draw_label(ws, r, c, "Modele", set_label_height=True)
    merge_value_B_to_G(ws, r, inputs["modele"], set_label_height=True)
    r += 1

    # Multi-line rows (86 px) merged B..G
    r = draw_multiline_value(ws, r, "From", inputs["from_addr"])
    r = draw_multiline_value(ws, r, "To",   inputs["to_addr"])

    # Box identifiers (20 px; merged B..G)
    draw_label(ws, r, c, "Box S.N", set_label_height=True)
    merge_value_B_to_G(ws, r, box["sn"], set_label_height=True)
    r += 1

    draw_label(ws, r, c, "Box code", set_label_height=True)
    merge_value_B_to_G(ws, r, box["code_box"], set_label_height=True)
    r += 1

    # Components table with merged Box S.N (A) and Box code (B)
    r = draw_components_table_with_merged_sn_and_code(ws, r, c, box)

    # Spacer
    r += 1
    return r

# ======================== Streamlit UI ========================

st.title("🎫 Container Packing List → Box Stickers (.xlsx)")
st.caption("Upload the packing list (.xlsx). I’ll generate one sticker per box, including all components and the required shipment details, arranged in a single worksheet (RTL) as a repeatable pattern.")

in_file = st.file_uploader("Upload packing list (In.xlsx)", type=["xlsx"], accept_multiple_files=False)

with st.expander("Required shipment details", expanded=True):
    packing_list_no = st.text_input("Packing List No.:", value="")
    order_no        = st.text_input("Order number:", value="")
    shipment_dt     = st.date_input("Date of shipment:", value=date.today())
    modele          = st.text_input("Modele:", value="")
    from_addr       = st.text_area("From address (multi-line):", value="Fresh Electric for Home Appliances\n10th of Ramadan City, Egypt.\nP.O.Box: 122")
    to_addr         = st.text_area("To address (multi-line):",   value="Customer / Plant\nCity, Country\nContact / Phone")

with st.expander("Advanced options", expanded=False):
    rtl = st.checkbox("Right-to-left worksheet (recommended for Arabic)", value=True)
    spacer_rows = st.number_input("Blank rows after each sticker:", min_value=0, max_value=5, value=1, step=1)

st.divider()

if in_file and st.button("Generate stickers", type="primary", use_container_width=True):
    try:
        wb_in = load_workbook(in_file, data_only=True)
        ws_in = find_first_sheet(wb_in)

        # Header mapping (forced: B for box code, E for component code, G for box type)
        header_candidates = {
            "sn":       ["S.N", "sn", "s.n", "serial", "box sn"],
            "comp_ar":  ["component in arabic", "arabic", "arabic name"],
            "comp_en":  ["component in english", "component in e", "english name", "english"],
            "qty":      ["Qut.", "Qu.", "Qty", "Quantity", "QTY"]
        }
        header_row, col_map = find_header_row(ws_in, header_candidates)
        if header_row is None:
            st.error("Could not detect the header row. Ensure the sheet includes: S.N, component in arabic, component in english, Qty/Qut.")
            st.stop()

        # Read & group
        data_rows = read_rows(ws_in, header_row, col_map)
        boxes = group_boxes(data_rows)

        if not boxes:
            st.error("No boxes/components found after the header row.")
            st.stop()

        # Build output workbook
        wb_out = Workbook()
        ws_out = wb_out.active
        ws_out.title = "Stickers"
        ws_out.sheet_view.rightToLeft = rtl

        set_default_widths(ws_out)

        next_top = 1
        for b in boxes:
            inputs = {
                "packing_list_no": packing_list_no.strip(),
                "order_no":        order_no.strip(),
                "date_str":        shipment_dt.strftime("%Y-%m-%d"),
                "modele":          modele.strip(),
                "from_addr":       from_addr.strip(),
                "to_addr":         to_addr.strip()
            }
            next_top = draw_sticker(ws_out, next_top, b, inputs)
            next_top += spacer_rows

        # Save to bytes and offer download
        buff = io.BytesIO()
        wb_out.save(buff)
        buff.seek(0)

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"box_stickers_{ts}.xlsx"
        st.success(f"✅ Generated {len(boxes)} stickers.")
        st.download_button(
            label="⬇️ Download Stickers Excel",
            data=buff,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # Optional preview
        with st.expander("Preview (first 10 rows from your In.xlsx)", expanded=False):
            st.dataframe(pd.DataFrame(data_rows[:10]), use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Generation failed: {e}")

else:
    st.info("Upload your packing list (In.xlsx), fill the details, then click **Generate stickers**.")


# ==== Centered footer ====
footer_css = """
<style>
.app-footer {
  position: fixed;
  left: 50%;
  bottom: 12px;
  transform: translateX(-50%);
  z-index: 9999;
  background: rgba(255,255,255,0.85);
  border: 1px solid #e6e6e6;
  border-radius: 14px;
  padding: 8px 14px;
  font-weight: 600;
  font-size: 14px;
}
</style>
"""
footer_html = """
<div class="app-footer">✨ تم التنفيذ بواسطة م / بيتر عمانوئيل – جميع الحقوق محفوظة © 2025 ✨</div>
"""
st.markdown(footer_css, unsafe_allow_html=True)
st.markdown(footer_html, unsafe_allow_html=True)

