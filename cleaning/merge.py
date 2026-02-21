# cleaning/merge.py
# -*- coding: utf-8 -*-

import re
from io import BytesIO
from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ==========================================================
# CONFIG (same sheet names as your originals)
# ==========================================================
priority_sheets = [
    "TAS Vehicle List",
    "Peugeot Vehicle List",
    "Citroen Vehicle List",
]

vehicle_list_sheets = {
    "TAS": "TAS Vehicle List",
    "Peugeot": "Peugeot Vehicle List",
    "Citroen": "Citroen Vehicle List",
}

vehicle_pattern = re.compile(r"^\d{2}-\d{6}$")


# =========================
# STYLE HELPERS (your look)
# =========================
THIN = Side(style="thin", color="FF000000")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FILL_LIST_HEADER = PatternFill("solid", fgColor="FF8B2E2E")   # dark red/maroon
FILL_MAIN_HEADER = PatternFill("solid", fgColor="FFF4C7A1")   # peach
FILL_PIECE_HEADER = PatternFill("solid", fgColor="FFCFE2F3")  # light blue
FILL_TOTAL = PatternFill("solid", fgColor="FFC6E0B4")         # light green

FONT_WHITE_BOLD = Font(bold=True, color="FFFFFFFF")
FONT_BOLD = Font(bold=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _apply_border_range(ws, min_row, max_row, min_col, max_col):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(r, c).border = BORDER_THIN


def _normalize(s):
    return str(s).strip().lower().replace("\u00a0", " ")


def style_vehicle_list(ws):
    max_row = ws.max_row
    max_col = ws.max_column

    header_row = None
    for r in range(1, min(15, max_row) + 1):
        row_vals = [_normalize(ws.cell(r, c).value) for c in range(1, min(10, max_col) + 1)]
        if any("vehicule id" in v or "véhicule id" in v for v in row_vals):
            header_row = r
            break
    if header_row is None:
        return

    table_max_col = min(max_col, 3) if max_col >= 3 else max_col

    for c in range(1, table_max_col + 1):
        cell = ws.cell(header_row, c)
        cell.fill = FILL_LIST_HEADER
        cell.font = FONT_WHITE_BOLD
        cell.alignment = ALIGN_CENTER

    last = header_row
    for r in range(header_row + 1, max_row + 1):
        if ws.cell(r, 1).value is None and ws.cell(r, 2).value is None and ws.cell(r, 3).value is None:
            if r > header_row + 2:
                break
        else:
            last = r

    _apply_border_range(ws, header_row, last, 1, table_max_col)

    for r in range(header_row + 1, last + 1):
        ws.cell(r, 1).font = FONT_BOLD
        ws.cell(r, 1).alignment = ALIGN_LEFT
        ws.cell(r, 2).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(r, 3).alignment = ALIGN_LEFT


def _find_row_contains(ws, text, col=1, search_rows=200):
    t = _normalize(text)
    for r in range(1, min(search_rows, ws.max_row) + 1):
        v = ws.cell(r, col).value
        if v is None:
            continue
        if t in _normalize(v):
            return r
    return None


def _style_section(ws, title_row, header_fill):
    if title_row is None:
        return

    header_row = title_row + 1
    max_col = ws.max_column

    last_col = 0
    for c in range(1, max_col + 1):
        if ws.cell(header_row, c).value is not None:
            last_col = c
    if last_col == 0:
        return

    for c in range(1, last_col + 1):
        cell = ws.cell(header_row, c)
        cell.fill = header_fill
        cell.font = FONT_BOLD
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_THIN

    end_row = header_row
    for r in range(header_row + 1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is not None:
            an = _normalize(a)
            if "piece" in an or "pièce" in an or "main d'oeuvre" in an or "main d'œuvre" in an:
                break
        if any(ws.cell(r, c).value is not None for c in range(1, last_col + 1)):
            end_row = r
        else:
            if r > header_row + 1:
                break

    _apply_border_range(ws, header_row, end_row, 1, last_col)

    for r in range(header_row + 1, end_row + 1):
        for c in range(1, last_col + 1):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)):
                ws.cell(r, c).alignment = Alignment(horizontal="right", vertical="center")
            else:
                ws.cell(r, c).alignment = ALIGN_LEFT


def style_vehicle_detail(ws):
    main_row = _find_row_contains(ws, "main d", col=1, search_rows=200)
    piece_row = _find_row_contains(ws, "pièce", col=1, search_rows=400) or _find_row_contains(ws, "piece", col=1, search_rows=400)

    if main_row:
        ws.cell(main_row, 1).font = Font(size=14, bold=True)
    if piece_row:
        ws.cell(piece_row, 1).font = Font(size=14, bold=True)

    _style_section(ws, main_row, FILL_MAIN_HEADER)
    _style_section(ws, piece_row, FILL_PIECE_HEADER)

    total_row = _find_row_contains(ws, "total htva", col=1, search_rows=800) or _find_row_contains(ws, "total htva", col=2, search_rows=800)
    if total_row:
        val_col = None
        for c in range(2, min(ws.max_column, 12) + 1):
            v = ws.cell(total_row, c).value
            if isinstance(v, (int, float)):
                val_col = c
                break

        ws.cell(total_row, 1).font = Font(size=12, bold=True)
        ws.cell(total_row, 1).alignment = ALIGN_LEFT

        if val_col:
            vc = ws.cell(total_row, val_col)
            vc.fill = FILL_TOTAL
            vc.font = Font(bold=True)
            vc.alignment = Alignment(horizontal="center", vertical="center")
            vc.border = BORDER_THIN


def copy_sheet(source_sheet, master_wb, new_title):
    new_sheet = master_wb.create_sheet(title=new_title)

    for col, dim in source_sheet.column_dimensions.items():
        new_sheet.column_dimensions[col] = copy(dim)

    for row, dim in source_sheet.row_dimensions.items():
        new_sheet.row_dimensions[row] = copy(dim)

    for merged in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged))

    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)

            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.alignment = copy(cell.alignment)
                new_cell.border = copy(cell.border)

            if cell.hyperlink:
                new_cell.hyperlink = copy(cell.hyperlink)

            if cell.number_format:
                new_cell.number_format = cell.number_format

    if source_sheet.freeze_panes:
        new_sheet.freeze_panes = source_sheet.freeze_panes

    new_sheet.sheet_properties.tabColor = source_sheet.sheet_properties.tabColor

    if "vehicle list" in (new_title or "").lower():
        style_vehicle_list(new_sheet)
    else:
        style_vehicle_detail(new_sheet)

    return new_sheet


# =========================
# GLOBAL SHEET STYLES
# =========================
thin2 = Side(style="thin", color="1F1F1F")
border2 = Border(left=thin2, right=thin2, top=thin2, bottom=thin2)

header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor="1F4E79")

section_font = Font(bold=True, color="FFFFFF", size=12)
section_fill = PatternFill("solid", fgColor="2F5597")

subheader_font = Font(bold=True, color="FFFFFF")
subheader_fill = PatternFill("solid", fgColor="1F4E79")

label_font = Font(bold=True)

wrap_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)


def set_col_widths(ws, widths):
    for col_idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w


def write_row(ws, row_idx, values, start_col=1, *,
              font=None, fill=None, alignment=None, apply_border=False):
    for j, v in enumerate(values, start=start_col):
        c = ws.cell(row_idx, j, v)
        if font:
            c.font = font
        if fill:
            c.fill = fill
        if alignment:
            c.alignment = alignment
        if apply_border:
            c.border = border2


# =========================
# Vehicle list reader
# =========================
def find_header_row_and_cols(ws):
    max_row = min(ws.max_row, 50)
    max_col = min(ws.max_column, 30)
    for r in range(1, max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, max_col + 1)]
        norm = [_normalize(v) for v in vals]
        if any("véhicule id" in v or "vehicule id" in v for v in norm):
            col_vid = None
            col_total = None
            col_link = None
            for c, v in enumerate(norm, start=1):
                if col_vid is None and ("véhicule id" in v or "vehicule id" in v):
                    col_vid = c
                if col_total is None and ("total htva" in v):
                    col_total = c
                if col_link is None and ("open sheet" in v or v == "open"):
                    col_link = c
            if col_vid is None:
                col_vid = 1
            if col_total is None:
                col_total = 2
            if col_link is None:
                col_link = 3 if ws.max_column >= 3 else None
            return r, col_vid, col_total, col_link
    return None, None, None, None


def read_vehicle_list(ws, brand):
    header_row, col_vid, col_total, _ = find_header_row_and_cols(ws)
    if header_row is None:
        return []

    data = []
    started = False
    for r in range(header_row + 1, ws.max_row + 1):
        vid = ws.cell(r, col_vid).value
        total = ws.cell(r, col_total).value

        if vid is None and total is None:
            if started:
                break
            else:
                continue

        started = True
        vid_s = str(vid).strip() if vid is not None else ""
        if not vehicle_pattern.match(vid_s):
            continue

        try:
            total_f = float(total) if total is not None else 0.0
        except Exception:
            total_f = 0.0

        data.append((vid_s, total_f, brand))
    return data


def relink_vehicle_list_to_global_details(ws):
    header_row, col_vid, _, col_link = find_header_row_and_cols(ws)
    if header_row is None or col_link is None:
        return

    started = False
    for r in range(header_row + 1, ws.max_row + 1):
        vid = ws.cell(r, col_vid).value
        link_cell = ws.cell(r, col_link)

        if vid is None and (link_cell.value is None):
            if started:
                break
            else:
                continue

        started = True
        vid_s = str(vid).strip() if vid is not None else ""
        if not vehicle_pattern.match(vid_s):
            continue

        target = f"V_{vid_s}"
        link_cell.value = link_cell.value or "Open"
        link_cell.hyperlink = f"#'{target}'!A1"
        link_cell.style = "Hyperlink"
        link_cell.alignment = ALIGN_LEFT


# ==========================================================
# TABLE EXTRACT FOR V_<VID>
# ==========================================================
def _norm_label(x):
    s = "" if x is None else str(x)
    s = s.strip().lower().replace("\u00a0", " ")
    s = s.replace("œ", "oe")
    s = s.replace("’", "'")
    s = re.sub(r"\s+", " ", s)
    return s


def extract_tables(ws):
    rows = [[cell.value for cell in r] for r in ws.iter_rows()]

    def find_contains(*needles):
        needles = [_norm_label(n) for n in needles]
        for i, r in enumerate(rows):
            a = _norm_label(r[0] if len(r) else None)
            if any(n in a for n in needles):
                return i
        return None

    i_main = find_contains("main d", "main d'oeuvre", "main d’œuvre", "main d'œuvre")
    i_piece = find_contains("pièce", "piece")

    end_total = len(rows)
    for i, r in enumerate(rows):
        a = _norm_label(r[0] if len(r) else None)
        if "total htva" in a or "total htv" in a:
            end_total = i
            break

    def extract(start, end):
        if start is None:
            return [], []
        header_row = start + 1
        headers = rows[header_row][:8]
        data = []
        for rr in rows[header_row + 1:end]:
            if any(v is not None for v in rr[:8]):
                data.append(rr[:8])
        return headers, data

    main_headers, main_data = extract(i_main, i_piece if i_piece is not None else end_total)
    piece_headers, piece_data = extract(i_piece, end_total)
    return (main_headers, main_data), (piece_headers, piece_data)


def write_table_block(ws, start_row, title, headers, data):
    r = start_row
    ws.cell(r, 1, title).font = label_font
    ws.cell(r, 1).alignment = wrap_left
    r += 1

    if not headers:
        ws.cell(r, 1, "(table not found)").alignment = wrap_left
        return r + 2

    write_row(ws, r, headers, start_col=1, font=subheader_font, fill=subheader_fill,
              alignment=center, apply_border=True)
    r += 1

    for rowvals in data:
        write_row(ws, r, rowvals, start_col=1, alignment=top_left, apply_border=True)
        r += 1

    return r + 2


# ==========================================================
# PUBLIC API: MERGE FROM BYTES
# ==========================================================
def build_one_dataset_from_bytes(brand_files: dict[str, bytes]) -> bytes:
    """
    brand_files example:
      {
        "TAS": tas_workbook_bytes,
        "Citroen": citroen_workbook_bytes,
        "Peugeot": peugeot_workbook_bytes
      }
    Returns merged workbook as bytes.
    """

    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    # 1) Global Vehicle List
    main_ws = out_wb.create_sheet("Global Vehicle List")

    # Load each brand wb from bytes
    brand_wbs = {}
    for brand, b in brand_files.items():
        brand_wbs[brand] = load_workbook(BytesIO(b), data_only=False)

    # 2) Copy priority sheets first
    for priority_name in priority_sheets:
        for brand, wb in brand_wbs.items():
            if priority_name in wb.sheetnames:
                copy_sheet(wb[priority_name], out_wb, priority_name)
                break

    # 3) Copy all remaining sheets + map vehicle detail sheets
    brand_sheet_map = {}  # (vid, brand) -> copied_sheet_name

    for brand, wb in brand_wbs.items():
        for sheet_name in wb.sheetnames:
            if sheet_name in priority_sheets:
                continue

            new_name = sheet_name
            counter = 1
            while new_name in out_wb.sheetnames:
                new_name = f"{sheet_name}_{counter}"
                counter += 1

            new_ws = copy_sheet(wb[sheet_name], out_wb, new_name)

            base = sheet_name.split("_")[0]
            if vehicle_pattern.match(base):
                brand_sheet_map[(base, brand)] = new_ws.title

    # 4) Read vehicle lists -> aggregate
    rows = []
    for brand, sheet_name in vehicle_list_sheets.items():
        if sheet_name not in out_wb.sheetnames:
            continue
        rows.extend(read_vehicle_list(out_wb[sheet_name], brand))

    agg = {}
    for vid, total, brand in rows:
        agg.setdefault(vid, {"total": 0.0, "brands": set()})
        agg[vid]["total"] += float(total)
        agg[vid]["brands"].add(brand)

    # 5) Write Global Vehicle List
    headers = ["Véhicule ID", "Total HTVA", "Go to sheet", "Brand(s)"]
    main_ws.append(headers)

    for col, h in enumerate(headers, 1):
        c = main_ws.cell(1, col, h)
        c.font = header_font
        c.fill = header_fill
        c.border = border2
        c.alignment = center

    for vid in sorted(agg.keys()):
        total = agg[vid]["total"]
        brands = sorted(agg[vid]["brands"])
        detail_sheet = f"V_{vid}"

        r = main_ws.max_row + 1
        main_ws.cell(r, 1, vid).border = border2
        main_ws.cell(r, 2, round(total, 3)).border = border2

        link_cell = main_ws.cell(r, 3, "Open")
        link_cell.hyperlink = f"#'{detail_sheet}'!A1"
        link_cell.style = "Hyperlink"
        link_cell.border = border2

        main_ws.cell(r, 4, ", ".join(brands)).border = border2

        for cidx in range(1, 5):
            main_ws.cell(r, cidx).alignment = Alignment(vertical="top", wrap_text=True)

    main_ws.freeze_panes = "A2"
    set_col_widths(main_ws, [14, 12, 12, 24])

    # 6) Create V_<VID> sheets
    for vid in sorted(agg.keys()):
        total = agg[vid]["total"]
        brands = sorted(agg[vid]["brands"])

        ws = out_wb.create_sheet(f"V_{vid}")

        ws["A1"] = "Back To Global List"
        ws["A1"].hyperlink = "#'Global Vehicle List'!A1"
        ws["A1"].style = "Hyperlink"

        ws["A3"] = f"Total HTVA: {round(total, 3)}"
        ws["A3"].font = Font(bold=True, size=12)

        current_row = 5
        copied_any = False

        for brand in brands:
            source_sheet_name = brand_sheet_map.get((vid, brand))
            if not source_sheet_name:
                continue

            src_ws = out_wb[source_sheet_name]
            (mh, md), (ph, pd_) = extract_tables(src_ws)

            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=8)
            ws.cell(current_row, 1, brand).font = section_font
            ws.cell(current_row, 1).fill = section_fill
            ws.cell(current_row, 1).alignment = wrap_left
            current_row += 2

            if (mh and md) or (ph and pd_):
                if mh:
                    current_row = write_table_block(ws, current_row, "Main d'œuvre", mh, md)
                if ph:
                    current_row = write_table_block(ws, current_row, "Pièces", ph, pd_)
                copied_any = True

        if not copied_any:
            ws.cell(5, 1, "No detail data found for this vehicle.").font = Font(bold=True, color="FF0000")

        set_col_widths(ws, [10, 14, 14, 8, 46, 6, 12, 12])
        ws.freeze_panes = "A5"

    # 7) Relink brand lists to V_<VID>
    for sheet_name in priority_sheets:
        if sheet_name in out_wb.sheetnames:
            relink_vehicle_list_to_global_details(out_wb[sheet_name])
            style_vehicle_list(out_wb[sheet_name])

    # Save to bytes
    bio = BytesIO()
    out_wb.save(bio)
    return bio.getvalue()