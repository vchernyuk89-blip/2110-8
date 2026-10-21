# -*- coding: utf-8 -*-
"""Ship Cont CAD CV v2.6 — maritime-themed container stowage planning and EDI/Excel workflow.
Requires Python 3.10+ and openpyxl for development builds. GitHub Actions builds the standalone Windows EXE.
"""
from __future__ import annotations
import re, sys, traceback, os, json, shutil, time, threading
from copy import copy
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.dimensions import ColumnDimension
except Exception as e:
    raise SystemExit("Не найдена библиотека openpyxl. Установите зависимости из requirements.txt.\n" + str(e))

APP_NAME = "Ship Cont CAD CV 2.4"
BLUE = "0077C8"
DARK_BLUE = "003B70"
PALE_BLUE = "DCEAF7"
GREY = "E7E6E6"
WHITE = "FFFFFF"
THIN = Side(style="thin", color="A6B4C0")

def clean(v):
    return "" if v is None else str(v).strip()

def num(v, default=0):
    try:
        if v is None or clean(v) == "":
            return default
        return int(float(str(v).replace(",", ".")))
    except Exception:
        return default

def parse_edi(path):
    raw = Path(path).read_text(encoding="utf-8-sig", errors="replace")
    segments = [s.strip() for s in raw.replace("\r", "").replace("\n", "").split("'") if s.strip()]
    records = []
    pending = {}
    header = {}
    active_record = None
    for seg in segments:
        p = seg.split("+")
        tag = p[0]
        if tag == "UNH" and len(p) > 2:
            header["message_type"] = p[2]
        elif tag == "BGM" and len(p) > 1:
            header["message_ref"] = p[1]
        elif tag == "DTM" and len(p) > 1:
            header["date"] = p[1]
        elif tag == "TDT" and len(p) > 5:
            if len(p) > 2 and p[2]: header["voyage_no"] = p[2]
            header["vessel_name"] = p[-1].split(":")[-1] if p[-1] else ""
            # In BAPLIE TDT, the vessel identifier is carried in C222:
            # identification + code-list qualifier (e.g. 11 = IMO number).
            for field in p[3:]:
                parts = field.split(":")
                if len(parts) >= 2 and parts[0].strip():
                    ident = parts[0].strip()
                    qualifier = parts[2].strip() if len(parts) > 2 else ""
                    if qualifier == "11" and re.fullmatch(r"\d{7}", ident):
                        header["vessel_id_type"] = "IMO"
                        header["vessel_id"] = ident
                        break
                    if qualifier == "ZZZ" and re.fullmatch(r"\d{8}", ident):
                        header["vessel_id_type"] = "ENI"
                        header["vessel_id"] = ident
                        break
        elif tag == "UNB" and len(p) > 3:
            header["sender_id"] = p[2]
            header["receiver_id"] = p[3]
        elif tag == "LOC" and len(p) > 2:
            code = p[2].split(":")[0]
            if p[1] == "5": header["pol"] = code
            elif p[1] == "61": header["pod"] = code
            elif p[1] == "147":
                digits = re.sub(r"\D", "", code)
                if len(digits) >= 6:
                    d = digits[-6:]
                    pending["LOC+147"] = code
                    pending["BAY"] = int(d[0:2])
                    pending["ROW"] = int(d[2:4])
                    pending["TIER"] = int(d[4:6])
            elif p[1] == "11": pending["POL"] = code
            elif p[1] == "13": pending["POT"] = code
            elif p[1] == "83": pending["POD"] = code
        elif tag == "MEA" and len(p) > 3:
            m = p[3].split(":")
            if len(m) > 1 and m[0] == "KGM": pending["GROSS_KG"] = num(m[1])
        elif tag == "RFF" and len(p) > 1:
            ref = p[1].split(":")
            if ref[0] == "BM" and len(ref) > 1: pending["BOOKING_NO"] = ref[1]
            elif ref[0] == "VON" and len(ref) > 1: header["rff_von"] = ref[1]
        elif tag == "EQD" and len(p) > 3:
            active_record = {
                "CONTAINER_NO": p[2], "ISO_TYPE": p[3], "LINE": "", "F_E": "F",
                "FINAL_DEST": "", "BOOKING_NO": pending.get("BOOKING_NO", ""),
                "GROSS_KG": pending.get("GROSS_KG", 0), "BAY": pending.get("BAY", 0),
                "ROW": pending.get("ROW", 0), "TIER": pending.get("TIER", 0),
                "POL": pending.get("POL", header.get("pol", "")),
                "POT": pending.get("POT", ""),
                "POD": pending.get("POD", header.get("pod", "")),
                "LOC+147": pending.get("LOC+147", "")
            }
            records.append(active_record)
            pending = {}
        elif tag == "NAD" and len(p) > 2 and p[1] == "CA":
            if active_record is not None:
                active_record["LINE"] = p[2].split(":")[0]
            else:
                pending["LINE"] = p[2].split(":")[0]
    return raw, header, records

def style_sheet(ws, freeze="A2"):
    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
            if c.value is not None:
                c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def container_label(c):
    weight = c.get("GROSS_KG")
    wt = f"{weight/1000:.1f}".replace(".", ",") + " t" if weight not in (None, "") else "—"
    return "\n".join(x for x in [clean(c.get("CONTAINER_NO")), clean(c.get("ISO_TYPE")), wt, clean(c.get("POD"))] if x)

def make_header(ws, title, row=1, end_col=12):
    ws.cell(row, 1, title)
    ws.cell(row, 1).font = Font(bold=True, size=14, color=WHITE)
    ws.cell(row, 1).fill = PatternFill("solid", fgColor=DARK_BLUE)
    ws.cell(row, 1).alignment = Alignment(horizontal="left", vertical="center")
    if end_col > 1:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
        for col in range(1, end_col+1):
            ws.cell(row, col).fill = PatternFill("solid", fgColor=DARK_BLUE)
    ws.row_dimensions[row].height = 26

def write_summary(wb, source, records, mode):
    ws = wb.create_sheet("SUMMARY", 0)
    make_header(ws, f"BAPLIE STOWAGE PLAN — {mode.upper()}", 1, 3)
    rows = [
        ("Исходный файл", Path(source).name),
        ("Контейнеров всего", len(records)),
        ("20 ft — количество", sum(1 for r in records if teu_for_record(r) == 1)),
        ("40/45 ft — количество", sum(1 for r in records if teu_for_record(r) == 2)),
        ("TEU всего", sum(teu_for_record(r) for r in records)),
        ("Суммарный вес, кг", sum(num(r.get("GROSS_KG")) for r in records)),
        ("Создано", datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]
    for i,(k,v) in enumerate(rows,3):
        ws.cell(i,1,k).font=Font(bold=True)
        ws.cell(i,2,v)
    ws.column_dimensions["A"].width=28; ws.column_dimensions["B"].width=34
    style_sheet(ws, "A3")

def write_containers(wb, records):
    ws = wb.create_sheet("CONTAINERS")
    headers = ["CONTAINER_NO","ISO_TYPE","GROSS_KG","WEIGHT_T","BAY","ROW","TIER","POL","POT","POD","LOC+147","LINE","BOOKING_NO"]
    ws.append(headers)
    for c in records:
        kg = num(c.get("GROSS_KG"))
        ws.append([c.get("CONTAINER_NO"),c.get("ISO_TYPE"),kg,round(kg/1000,1),c.get("BAY"),c.get("ROW"),c.get("TIER"),c.get("POL"),c.get("POT"),c.get("POD"),c.get("LOC+147"),c.get("LINE"),c.get("BOOKING_NO")])
    for c in ws[1]:
        c.font=Font(bold=True,color=WHITE); c.fill=PatternFill("solid",fgColor=DARK_BLUE)
    for col,w in {"A":18,"B":12,"C":13,"D":12,"E":8,"F":8,"G":8,"H":12,"I":12,"J":12,"K":14,"L":18,"M":16}.items():
        ws.column_dimensions[col].width=w
    ws.auto_filter.ref=ws.dimensions
    style_sheet(ws,"A2")

def tier_list(records):
    vals = sorted({num(r.get("TIER")) for r in records if num(r.get("TIER")) > 0})
    return vals or [2]

def odd_bays(records):
    vals = {num(r.get("BAY")) for r in records if num(r.get("BAY")) > 0}
    # Show all odd columns around even 40-foot positions too.
    expanded = set()
    for b in vals:
        if b % 2 == 0: expanded.update([b-1,b+1])
        else: expanded.add(b)
    return sorted((b for b in expanded if b > 0 and b % 2 == 1), reverse=True)

def row_order(records):
    vals = {num(r.get("ROW")) for r in records}
    # Match prior plan orientation: even rows on one side, then center, then odd rows.
    evens = sorted([v for v in vals if v > 0 and v % 2 == 0], reverse=True)
    zeros = [0] if 0 in vals else []
    odds = sorted([v for v in vals if v % 2 == 1])
    order = evens + zeros + odds
    return order or [0]

def write_tier_sheet(ws, tier, records, title_prefix="TIER", start_row=1):
    subset = [r for r in records if num(r.get("TIER")) == tier]
    bays = odd_bays(records)
    rows = row_order(records)
    title_row = start_row
    header_row = start_row + 1
    data_start = start_row + 2
    make_header(ws, f"{title_prefix} {tier:02d} — TOP-DOWN CONTAINER LAYOUT", title_row, len(bays)+1)
    ws.cell(header_row,1,"Row / Bay")
    ws.cell(header_row,1).font=Font(bold=True,color=WHITE); ws.cell(header_row,1).fill=PatternFill("solid",fgColor=BLUE)
    for j,b in enumerate(bays,2):
        ws.cell(header_row,j,f"Bay {b:02d}")
        ws.cell(header_row,j).font=Font(bold=True,color=WHITE); ws.cell(header_row,j).fill=PatternFill("solid",fgColor=BLUE)
        ws.column_dimensions[get_column_letter(j)].width=22
    for i,row in enumerate(rows,data_start):
        ws.cell(i,1,f"Row {row:02d}")
        ws.cell(i,1).font=Font(bold=True); ws.cell(i,1).fill=PatternFill("solid",fgColor=PALE_BLUE)
        ws.row_dimensions[i].height=68
        for j,b in enumerate(bays,2):
            ws.cell(i,j,"")
            ws.column_dimensions[get_column_letter(j)].width=22
    rowmap={v:i+data_start for i,v in enumerate(rows)}
    baymap={v:i+2 for i,v in enumerate(bays)}
    for c in subset:
        b=num(c.get("BAY")); r=num(c.get("ROW"))
        if r not in rowmap: continue
        rr=rowmap[r]
        if b % 2 == 1:
            col=baymap.get(b)
            if col:
                cell=ws.cell(rr,col,container_label(c))
                cell.fill=PatternFill("solid",fgColor=PALE_BLUE if clean(c.get("ISO_TYPE"))=="22G1" else "FCE4D6")
                cell.font=Font(size=8,bold=True)
        else:
            left_col=baymap.get(b+1); right_col=baymap.get(b-1)
            if left_col and right_col:
                c1=min(left_col,right_col); c2=max(left_col,right_col)
                try: ws.merge_cells(start_row=rr,start_column=c1,end_row=rr,end_column=c2)
                except Exception: pass
                cell=ws.cell(rr,c1,container_label(c))
                cell.fill=PatternFill("solid",fgColor="FCE4D6")
                cell.font=Font(size=8,bold=True)
    style_sheet(ws,f"B{data_start}")

def section_for_bay(b):
    if b % 2 == 0: return b
    return ((b-1)//4)*4+2

def write_bay_section(ws, center, records, compact_title=True, start_row=1):
    left_bay, right_bay = center+1, center-1
    tiers = sorted({num(r.get("TIER")) for r in records if num(r.get("TIER"))>0}, reverse=True)
    rows = row_order(records)
    # Each section has left and right plans with each row across.
    make_header(ws, f"BAY {center:02d} — ПОПЕРЕЧНАЯ СХЕМА", start_row, len(rows)*2+1)
    headerrow=start_row+1
    ws.cell(headerrow,1,"Tier")
    ws.merge_cells(start_row=headerrow,start_column=2,end_row=headerrow,end_column=1+len(rows))
    ws.cell(headerrow,2,f"BAY {left_bay:02d}  (левая сторона)")
    ws.merge_cells(start_row=headerrow,start_column=2+len(rows),end_row=headerrow,end_column=1+2*len(rows))
    ws.cell(headerrow,2+len(rows),f"BAY {right_bay:02d}  (правая сторона)")
    for col in range(1,2*len(rows)+2):
        ws.cell(headerrow,col).fill=PatternFill("solid",fgColor=BLUE)
        ws.cell(headerrow,col).font=Font(bold=True,color=WHITE)
    rowhead=headerrow+1
    ws.cell(rowhead,1,"Tier")
    for side_offset in (1,1+len(rows)):
        for j,r in enumerate(rows):
            cell=ws.cell(rowhead,side_offset+j+1,f"Row {r:02d}")
            cell.fill=PatternFill("solid",fgColor=PALE_BLUE); cell.font=Font(bold=True)
    relevant=[]
    for c in records:
        b=num(c.get("BAY"))
        if section_for_bay(b)==center:
            relevant.append(c)
    # Tier levels include observed levels, with a sensible descending range to preserve empty slots.
    if tiers:
        high=max(tiers); low=min(tiers)
        if high >= 80: levels=list(range(high, max(79,low)-1, -2)) + list(range(6, max(0,low)-1,-2)) if low < 80 else list(range(high,low-1,-2))
        else: levels=list(range(high,low-1,-2))
        levels=sorted(set(levels)|set(tiers), reverse=True)
    else: levels=[2]
    rowmap={}
    base= rowhead+1
    for i,t in enumerate(levels):
        rr=base+i; rowmap[t]=rr
        ws.cell(rr,1,f"Tier {t:02d}")
        ws.cell(rr,1).fill=PatternFill("solid",fgColor=PALE_BLUE); ws.cell(rr,1).font=Font(bold=True)
        ws.row_dimensions[rr].height=66
        for col in range(2,2*len(rows)+2): ws.cell(rr,col,"")
    rowmap_row={v:i for i,v in enumerate(rows)}
    for c in relevant:
        b=num(c.get("BAY")); r=num(c.get("ROW")); t=num(c.get("TIER"))
        if t not in rowmap or r not in rowmap_row: continue
        rr=rowmap[t]
        is40 = b%2==0 or clean(c.get("ISO_TYPE")) in ("42G1","45G1","L5G1","45R1")
        if is40:
            # 40/45ft: display container in the right-hand Bay and X at corresponding left-hand position.
            pos=rowmap_row.get(r)
            if pos is None: continue
            rightcol=2+len(rows)+pos
            leftcol=2+pos
            ws.cell(rr,rightcol,container_label(c)).fill=PatternFill("solid",fgColor="FCE4D6")
            ws.cell(rr,rightcol).font=Font(size=8,bold=True)
            ws.cell(rr,leftcol,"X").font=Font(size=20,bold=True,color="C00000")
            ws.cell(rr,leftcol).fill=PatternFill("solid",fgColor="F4CCCC")
        else:
            side = 0 if b==left_bay else 1
            pos=rowmap_row.get(r)
            col=2+side*len(rows)+pos
            ws.cell(rr,col,container_label(c))
            ws.cell(rr,col).fill=PatternFill("solid",fgColor=PALE_BLUE)
            ws.cell(rr,col).font=Font(size=8,bold=True)
    for col in range(2,2*len(rows)+2):
        ws.column_dimensions[get_column_letter(col)].width=19
    ws.column_dimensions["A"].width=12
    style_sheet(ws, f"B{base}")

# ------------------------- Ship Cont CAD CV v1.07 -------------------------

FIELDS = [
    "CONTAINER_NO", "LINE", "ISO_TYPE", "GROSS_KG", "BAY", "ROW", "TIER",
    "POL", "POT", "POD", "F_E", "FINAL_DEST", "BOOKING_NO", "REEFER_TEMP_C",
    "UN_NUMBER", "IMDG_CLASS", "OVERDIMENSION_NOTE", "SEAL_NO", "COMMENTS"
]
FIELD_LABELS = {
    "CONTAINER_NO": "Номер контейнера", "LINE": "Линия / перевозчик", "ISO_TYPE": "ISO-код",
    "GROSS_KG": "Вес брутто, кг", "BAY": "Bay", "ROW": "Row", "TIER": "Tier",
    "POL": "Порт погрузки (POL)", "POT": "Порт перевалки (POT)", "POD": "Порт выгрузки (POD)", "F_E": "F/E (гружёный/порожний)",
    "FINAL_DEST": "Конечный пункт", "BOOKING_NO": "Booking No.", "REEFER_TEMP_C": "Температура рефрижератора, °C",
    "UN_NUMBER": "UN Number", "IMDG_CLASS": "IMDG Class", "OVERDIMENSION_NOTE": "Негабарит / примечание",
    "SEAL_NO": "Номер пломбы", "COMMENTS": "Комментарии"
}
ALIASES = {
    "CONTAINER_NO": {"CONTAINER_NO", "CONTAINER NO", "CONTAINER NO.", "CONTAINER NUMBER", "CONTAINER"},
    "LINE": {"LINE", "CARRIER", "SHIPPING LINE", "NAD+CA"},
    "ISO_TYPE": {"ISO_TYPE", "ISO TYPE", "TYPE", "ISO CODE"},
    "GROSS_KG": {"GROSS_KG", "GROSS KG", "WEIGHT KG", "WEIGHT (KG)", "WEIGHT"},
    "BAY": {"BAY", "BAY (EDI)", "BAY NO"},
    "ROW": {"ROW", "ROW (EDI)", "ROW NO"},
    "TIER": {"TIER", "TIER (EDI)", "TIER NO"},
    "POL": {"POL", "LOAD PORT", "PORT OF LOADING", "LOAD PORT (LOC+11)"},
    "POT": {"POT", "TRANSHIPMENT PORT", "TRANSSHIPMENT PORT", "PORT OF TRANSHIPMENT", "PORT OF TRANSSHIPMENT", "TRANSFER PORT"},
    "POD": {"POD", "DISCHARGE PORT", "PORT OF DISCHARGE", "POD (LOC+83)", "DISCHARGE PORT (LOC+83)"},
    "F_E": {"F_E", "F/E", "F E", "FULL EMPTY"},
    "FINAL_DEST": {"FINAL_DEST", "FINAL DEST", "FINAL DESTINATION"},
    "BOOKING_NO": {"BOOKING_NO", "BOOKING NO", "BOOKING", "BOOKING NUMBER"},
    "REEFER_TEMP_C": {"REEFER_TEMP_C", "REEFER TEMP C", "TEMPERATURE", "TEMP C"},
    "UN_NUMBER": {"UN_NUMBER", "UN NUMBER"},
    "IMDG_CLASS": {"IMDG_CLASS", "IMDG CLASS"},
    "OVERDIMENSION_NOTE": {"OVERDIMENSION_NOTE", "OVERDIMENSION NOTE", "OOG NOTE"},
    "SEAL_NO": {"SEAL_NO", "SEAL NO", "SEAL NUMBER"},
    "COMMENTS": {"COMMENTS", "COMMENT", "NOTES"}
}
DEFAULT_VESSEL = {
    "VESSEL_NAME": "TBN", "IMO": "", "VESSEL_ID_TYPE": "IMO", "VESSEL_ID": "", "VOYAGE_NO": "1", "VESSEL_OPERATOR": "",
    "SENDER_ID": "", "RECEIVER_ID": "", "MESSAGE_VERSION": "D:95B:UN:SMDG20",
    "MESSAGE_FUNCTION": "", "PORT_OF_DEPARTURE": "", "PORT_OF_ARRIVAL": "",
    "PORT_TERMINAL": "", "MESSAGE_DATE_TIME": "", "CURRENCY/UNITS_NOTE": "Масса в кг; температура °C"
}
SUPPORTED_OPEN = [("Поддерживаемые файлы", "*.edi *.txt *.baplie *.xlsx *.xlsm"), ("EDI", "*.edi *.txt *.baplie"), ("Excel", "*.xlsx *.xlsm")]


def app_home():
    # Stable writable storage, outside the executable so updates don't remove projects.
    # Migrate the old app data folder once so upgrading does not lose user files.
    docs = Path.home() / "Documents"
    if not docs.exists():
        docs = Path.home()
    new_home = docs / "Ship Cont CAD CV"
    old_home = docs / "CV Container Ship"
    try:
        if old_home.exists() and not new_home.exists():
            shutil.copytree(old_home, new_home)
    except Exception:
        # If migration is not possible, keep using the new path; old files remain untouched.
        pass
    return new_home

APP_HOME = app_home()
PROJECTS_DIR = APP_HOME / "Projects"
EXPORTS_DIR = APP_HOME / "Exports"
BACKUPS_DIR = APP_HOME / "Backups"
SKELETONS_DIR = APP_HOME / "VesselSkeletons"
RECOVERY_DIR = APP_HOME / "Recovery"
ACTIVE_PROJECTS_DIR = APP_HOME / "ActiveProjects"
ARCHIVED_PROJECTS_DIR = APP_HOME / "ArchivedProjects"
SAMPLES_DIR = APP_HOME / "Samples"
RECOVERY_PATH = RECOVERY_DIR / "autosave_recovery.json"
CONFIG_PATH = APP_HOME / "config.json"


def resource_path(relative):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative


def ensure_folders():
    APP_HOME.mkdir(parents=True, exist_ok=True)
    PROJECTS_DIR.mkdir(parents=True, exist_ok=True)
    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    BACKUPS_DIR.mkdir(parents=True, exist_ok=True)
    SKELETONS_DIR.mkdir(parents=True, exist_ok=True)
    RECOVERY_DIR.mkdir(parents=True, exist_ok=True)
    ACTIVE_PROJECTS_DIR.mkdir(parents=True, exist_ok=True)
    ARCHIVED_PROJECTS_DIR.mkdir(parents=True, exist_ok=True)
    SAMPLES_DIR.mkdir(parents=True, exist_ok=True)


def clean_header_key(value):
    return re.sub(r"\s+", " ", clean(value).strip().upper().replace("_", " ").replace(".", "")).strip()


def field_for_header(value):
    key = clean_header_key(value)
    for field, aliases in ALIASES.items():
        if key in {clean_header_key(a) for a in aliases}:
            return field
    # Preserve exact canonical header names with underscores.
    canonical = clean(value).strip().upper()
    return canonical if canonical in FIELDS else None


def empty_record():
    return {k: "" for k in FIELDS}


def normalize_record(source):
    rec = empty_record()
    for k, v in (source or {}).items():
        key = clean(k).strip().upper()
        if key in rec:
            rec[key] = "" if v is None else v
    for k in ("BAY", "ROW", "TIER", "GROSS_KG"):
        if clean(rec.get(k)) != "":
            rec[k] = num(rec[k], 0)
    return rec


def read_excel_document(path):
    wb = load_workbook(path, data_only=True)
    header = dict(DEFAULT_VESSEL)
    if "VESSEL" in wb.sheetnames:
        vs = wb["VESSEL"]
        for row in vs.iter_rows(min_row=2, values_only=True):
            if row and row[0] is not None:
                key = clean(row[0])
                if len(row) > 1 and row[1] is not None:
                    value = row[1]
                    text_value = clean(value).lower()
                    if "по требованиям получателя" in text_value or "заполните, если известно" in text_value:
                        value = ""
                    if text_value == "не знаю":
                        value = DEFAULT_VESSEL.get(key, "")
                    header[key] = value
                elif key not in header:
                    header[key] = ""
    records = []
    ws = wb["CONTAINERS"] if "CONTAINERS" in wb.sheetnames else None
    if ws is None:
        # Old conversion outputs or user workbooks may have a translated sheet title.
        for candidate in wb.worksheets:
            hdrs = [field_for_header(c.value) for c in candidate[1]]
            if "CONTAINER_NO" in hdrs and ("BAY" in hdrs or "ROW" in hdrs):
                ws = candidate
                break
    if ws is None:
        wb.close()
        raise ValueError("Не найден лист CONTAINERS или таблица с колонками контейнеров.")
    colmap = {}
    for i, cell in enumerate(ws[1], 1):
        field = field_for_header(cell.value)
        if field:
            colmap[field] = i
    if "CONTAINER_NO" not in colmap:
        wb.close()
        raise ValueError("На листе CONTAINERS не найден столбец CONTAINER_NO (номер контейнера).")
    for row in ws.iter_rows(min_row=2, values_only=True):
        cn_i = colmap.get("CONTAINER_NO", 0) - 1
        if cn_i < 0 or cn_i >= len(row) or not clean(row[cn_i]):
            continue
        rec = empty_record()
        for field, col in colmap.items():
            if col <= len(row):
                val = row[col-1]
                if val is not None:
                    rec[field] = val
        records.append(normalize_record(rec))
    wb.close()
    return header, records


def clone_row_style(ws, src_row, dst_row, max_col):
    for col in range(1, max_col+1):
        source = ws.cell(src_row, col)
        dest = ws.cell(dst_row, col)
        if source.has_style:
            dest._style = copy(source._style)
        if source.number_format:
            dest.number_format = copy(source.number_format)
        if source.alignment:
            dest.alignment = copy(source.alignment)


def find_template_path(config):
    p = clean(config.get("template_path", ""))
    if p and Path(p).exists():
        return Path(p)
    bundled = resource_path("EDI_Input_Template.xlsx")
    if bundled.exists(): return bundled
    nested = resource_path("templates/EDI_Input_Template.xlsx")
    return nested if nested.exists() else None


def config_load():
    ensure_folders()
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def config_save(config):
    ensure_folders()
    CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def write_input_workbook(path, vessel, records, template_path=None):
    path = Path(path)
    template = Path(template_path) if template_path and Path(template_path).exists() else (resource_path("EDI_Input_Template.xlsx") if resource_path("EDI_Input_Template.xlsx").exists() else resource_path("templates/EDI_Input_Template.xlsx"))
    if path.exists() and path.suffix.lower() in (".xlsx", ".xlsm"):
        # Preserve extra vessel sheets only for a workbook that is itself editable input format.
        try:
            temp_wb = load_workbook(path, read_only=True, data_only=True)
            has_input = "VESSEL" in temp_wb.sheetnames and "CONTAINERS" in temp_wb.sheetnames
            temp_wb.close()
        except Exception:
            has_input = False
    else:
        has_input = False
    base = path if has_input else template
    if not base or not Path(base).exists():
        wb = Workbook()
        vs = wb.active; vs.title = "VESSEL"
        vs.append(["Поле", "Значение", "Пояснение"])
        for key, val in DEFAULT_VESSEL.items(): vs.append([key, vessel.get(key, val), "Заполните, если известно"])
        cs = wb.create_sheet("CONTAINERS")
        cs.append(FIELDS)
    else:
        keep_vba = str(base).lower().endswith(".xlsm")
        wb = load_workbook(base, keep_vba=keep_vba)
        if "VESSEL" not in wb.sheetnames:
            vs = wb.create_sheet("VESSEL", 0)
            vs.append(["Поле", "Значение", "Пояснение"])
        if "CONTAINERS" not in wb.sheetnames:
            cs = wb.create_sheet("CONTAINERS")
            cs.append(FIELDS)
        vs = wb["VESSEL"]
        for row in vs.iter_rows(min_row=2):
            key = clean(row[0].value) if row else ""
            if key in vessel and len(row) > 1:
                row[1].value = vessel[key]
        cs = wb["CONTAINERS"]
        colmap = {}
        for col in range(1, cs.max_column+1):
            f = field_for_header(cs.cell(1, col).value)
            if f:
                colmap[f] = col
        # Upgrade older input sheets safely when new canonical fields (such as POT) are added.
        for field in FIELDS:
            if field not in colmap:
                nonempty_headers = [col for col in range(1, cs.max_column + 1) if clean(cs.cell(1, col).value)]
                new_col = (max(nonempty_headers) + 1) if nonempty_headers else 1
                cs.cell(1, new_col, field)
                if cs.cell(1, 1).has_style:
                    cs.cell(1, new_col)._style = copy(cs.cell(1, 1)._style)
                for rr in range(2, cs.max_row + 1):
                    if cs.cell(rr, 1).has_style:
                        cs.cell(rr, new_col)._style = copy(cs.cell(rr, 1)._style)
                cs.column_dimensions[get_column_letter(new_col)].width = 16
                colmap[field] = new_col
        if "CONTAINER_NO" not in colmap:
            # Rebuild just the container sheet if the source layout isn't compatible.
            idx = wb.sheetnames.index("CONTAINERS")
            del wb["CONTAINERS"]
            cs = wb.create_sheet("CONTAINERS", idx)
            for j, f in enumerate(FIELDS, 1):
                cs.cell(1, j, f)
            colmap = {f: j for j, f in enumerate(FIELDS, 1)}
        maxrow = max(cs.max_row, len(records)+1)
        for r in range(2, maxrow+1):
            for c in range(1, max(cs.max_column, len(FIELDS))+1):
                cs.cell(r, c).value = None
        if cs.max_row >= 2:
            style_source = 2
        else:
            style_source = 1
        for i, rec0 in enumerate(records, 2):
            rec = normalize_record(rec0)
            if i > 2:
                clone_row_style(cs, style_source, i, max(cs.max_column, max(colmap.values(), default=1)))
            for field, col in colmap.items():
                val = rec.get(field, "")
                cs.cell(i, col, val if clean(val) != "" else None)
        if "VESSEL" in wb.sheetnames:
            wb["VESSEL"].sheet_view.showGridLines = False
        cs.freeze_panes = "A2"
        if cs.max_column:
            cs.auto_filter.ref = f"A1:{get_column_letter(cs.max_column)}{max(1, len(records)+1)}"
    # Save .xlsx/.xlsm based on selected destination. SaveAs might point to .edi; guard caller.
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()


def create_edi_text(vessel, records):
    def getv(key, default=""):
        val = vessel.get(key, default)
        return default if val is None else clean(val)
    now = datetime.now().strftime("%y%m%d%H%M")
    ref = now + "20"
    msg = now + "21"
    operator = getv("VESSEL_OPERATOR", "ITERIS") or "ITERIS"
    sender = getv("SENDER_ID", "")
    receiver = getv("RECEIVER_ID", operator) or operator
    msgver = getv("MESSAGE_VERSION", "D:95B:UN:SMDG20")
    if not msgver or msgver.lower() == "не знаю" or not msgver.startswith("D:"):
        msgver = "D:95B:UN:SMDG20"
    vesselname = getv("VESSEL_NAME", "TBN") or "TBN"
    id_type = getv("VESSEL_ID_TYPE", "IMO").upper() or "IMO"
    vessel_id = getv("VESSEL_ID", "") or getv("IMO", "")
    if vessel_id:
        if id_type == "IMO":
            if not re.fullmatch(r"\d{7}", vessel_id):
                raise ValueError("IMO должен содержать ровно 7 цифр без префикса IMO.")
            checksum = sum(int(vessel_id[i]) * (7 - i) for i in range(6)) % 10
            if checksum != int(vessel_id[6]):
                raise ValueError("Контрольная цифра IMO неверна. Проверьте номер судна.")
        elif id_type == "ENI":
            if not re.fullmatch(r"\d{8}", vessel_id):
                raise ValueError("ENI должен содержать ровно 8 цифр.")
        else:
            raise ValueError("Выберите тип идентификатора IMO или ENI.")
    voyage = getv("VOYAGE_NO", "1") or "1"
    pol_default = getv("POL", getv("PORT_OF_DEPARTURE", "")) or getv("PORT_OF_DEPARTURE", "")
    pod_default = getv("POD", getv("PORT_OF_ARRIVAL", "")) or getv("PORT_OF_ARRIVAL", "")
    segments = [
        f"UNB+UNOA:2+{sender}+{receiver}+{now[:6]}:{now[6:10]}+{ref}+++++{operator}",
        f"UNH+{msg}+BAPLIE:{msgver}",
        f"BGM++{ref}+9",
        f"DTM+137:{now}:201",
        (f"TDT+20+{voyage}+++{operator}:LINES:306+++{vessel_id}::11:{vesselname}" if vessel_id and id_type == "IMO" else
         f"TDT+20+{voyage}+++{operator}:LINES:306+++{vessel_id}::ZZZ:{vesselname}" if vessel_id and id_type == "ENI" else
         f"TDT+20+{voyage}+1+1++:172:20+++-:103:ZZZ:{vesselname}"),
    ]
    if pol_default: segments.append(f"LOC+5+{pol_default}:139:6")
    if pod_default: segments.append(f"LOC+61+{pod_default}:139:6")
    segments.append(f"RFF+VON:{operator}")
    for raw in records:
        c = normalize_record(raw)
        if not clean(c.get("CONTAINER_NO")):
            continue
        bay, row, tier = num(c.get("BAY")), num(c.get("ROW")), num(c.get("TIER"))
        if bay < 0 or row < 0 or tier < 0:
            raise ValueError(f"У контейнера {clean(c.get('CONTAINER_NO'))} отрицательные координаты BAY/ROW/TIER.")
        loc = f"{bay:02d}{row:02d}{tier:02d}"
        kg = num(c.get("GROSS_KG"), 0)
        pol = clean(c.get("POL")) or pol_default
        pod = clean(c.get("POD")) or pod_default
        line = clean(c.get("LINE")) or operator
        booking = clean(c.get("BOOKING_NO"))
        iso = clean(c.get("ISO_TYPE")) or "22G1"
        segments.append(f"LOC+147+0{loc}::5")
        segments.append(f"MEA+WT++KGM:{kg}")
        pot = clean(c.get("POT"))
        if pol: segments.append(f"LOC+11+{pol}")
        if pot: segments.append(f"LOC+13+{pot}")
        if pod: segments.append(f"LOC+83+{pod}")
        if booking: segments.append(f"RFF+BM:{booking}")
        segments.append(f"EQD+CN+{clean(c.get('CONTAINER_NO'))}+{iso}+++5")
        segments.append(f"NAD+CA+{line}:172:ZZZ")
    unh_index = next(i for i, s in enumerate(segments) if s.startswith("UNH+"))
    segments.append(f"UNT+{len(segments)-unh_index+1}+{msg}")
    segments.append(f"UNZ+1+{ref}")
    return "'".join(segments) + "'"


def make_export_workbook(source_name, records, mode, vessel=None, raw_edi=""):
    wb = Workbook()
    wb.remove(wb.active)
    write_summary(wb, source_name, records, mode)
    write_containers(wb, records)
    if mode == "Tier":
        for t in tier_list(records):
            ws = wb.create_sheet(f"Tier {t:02d}")
            write_tier_sheet(ws, t, records)
    elif mode == "Bay":
        centers = sorted({section_for_bay(num(c.get("BAY"))) for c in records if num(c.get("BAY")) > 0})
        for center in centers:
            ws = wb.create_sheet(f"Bay {center:02d}")
            write_bay_section(ws, center, records)
    else:
        ws = wb.create_sheet("TIER Version")
        cursor = 1
        for t in tier_list(records):
            write_tier_sheet(ws, t, records, title_prefix="TIER", start_row=cursor)
            cursor = ws.max_row + 2
        ws = wb.create_sheet("BAY Version")
        centers = sorted({section_for_bay(num(c.get("BAY"))) for c in records if num(c.get("BAY")) > 0})
        cursor = 1
        for center in centers:
            write_bay_section(ws, center, records, start_row=cursor)
            cursor = ws.max_row + 2
    rawws = wb.create_sheet("RAW EDI")
    rawws.append(["Segment No.", "Segment Tag", "Raw segment"])
    for cell in rawws[1]:
        cell.font = Font(bold=True, color=WHITE); cell.fill = PatternFill("solid", fgColor=DARK_BLUE)
    for n, segment in enumerate([s for s in raw_edi.split("'") if s.strip()], 1):
        rawws.append([n, segment.split("+", 1)[0], segment.strip()])
    if not raw_edi:
        rawws.append([1, "INFO", "RAW EDI появится после импорта или экспорта EDI."])
    rawws.column_dimensions["A"].width = 13
    rawws.column_dimensions["B"].width = 16
    rawws.column_dimensions["C"].width = 120
    rawws.freeze_panes = "A2"
    rawws.auto_filter.ref = rawws.dimensions
    return wb


def decode_edi_document(path):
    raw, header, parsed = parse_edi(path)
    records = [normalize_record(r) for r in parsed]
    # Carry global ports into rows where per-container values were not in the EDI.
    pol = clean(header.get("pol", ""))
    pod = clean(header.get("pod", ""))
    for r in records:
        if not clean(r.get("POL")): r["POL"] = pol
        if not clean(r.get("POD")): r["POD"] = pod
    vessel = dict(DEFAULT_VESSEL)
    msgver = clean(header.get("message_type", "D:95B:UN:SMDG20"))
    if msgver.startswith("BAPLIE:"): msgver = msgver.split(":", 1)[1]
    vessel.update({"VESSEL_NAME": header.get("vessel_name", "TBN"), "VOYAGE_NO": header.get("voyage_no", "1"),
                   "VESSEL_OPERATOR": header.get("rff_von", ""), "SENDER_ID": header.get("sender_id", ""),
                   "RECEIVER_ID": header.get("receiver_id", ""),
                   "PORT_OF_DEPARTURE": pol, "PORT_OF_ARRIVAL": pod,
                   "VESSEL_ID_TYPE": header.get("vessel_id_type", "IMO"),
                   "VESSEL_ID": header.get("vessel_id", header.get("imo", "")),
                   "IMO": header.get("vessel_id", header.get("imo", "")) if header.get("vessel_id_type", "IMO") == "IMO" else "",
                   "MESSAGE_VERSION": msgver})
    return raw, vessel, records


def validation_issues(records):
    """Compatibility wrapper returning human-readable plan issues."""
    return [item["message"] for item in detailed_validation_issues(records)]


TEU_FIELD_OPTIONS = ["Общая сводка", "По линиям", "По TEU / размеру", "По Bay", "По Tier", "По POD", "По POT", "По типам контейнеров"]
COLOR_MODE_OPTIONS = ["Тип контейнера", "Порт назначения (POD)", "Порт перевалки (POT)", "Порт погрузки (POL)", "Линия", "Статус проверки"]
COLOR_MODE_FIELDS = {
    "Тип контейнера": "ISO_TYPE", "Порт назначения (POD)": "POD",
    "Порт перевалки (POT)": "POT", "Порт погрузки (POL)": "POL",
    "Линия": "LINE", "Статус проверки": "__VALIDATION__"
}
COLOR_PALETTE = ["#DCEAF7", "#FCE4D6", "#D9EAD3", "#FFF2CC", "#D9D2E9", "#CFE2F3", "#F4CCCC", "#D0E0E3", "#EAD1DC", "#E2F0D9", "#FCE5CD", "#D9EAD3", "#C9DAF8", "#E6B8AF", "#D9D2E9", "#CFE2F3"]

# Editable application reference data. The user's entries are persisted in config.json.
DEFAULT_REFERENCE_DATA = {
    "Порты": [],
    "Типы ISO": [
        {"key":"22G1","short":"DV","description":"20 ft standard dry","group":"Dry"},
        {"key":"42G1","short":"DV","description":"40 ft standard dry","group":"Dry"},
        {"key":"45G1","short":"HC","description":"40 ft High Cube dry","group":"Dry"},
        {"key":"22R1","short":"RF","description":"20 ft refrigerated","group":"Reefer"},
        {"key":"45R1","short":"HR","description":"40 ft High Cube Reefer","group":"Reefer"},
        {"key":"22U1","short":"OT","description":"20 ft Open Top","group":"Open Top"},
        {"key":"45U1","short":"OT","description":"40 ft Open Top","group":"Open Top"},
        {"key":"22T1","short":"TK","description":"20 ft tank container","group":"Tank"},
        {"key":"22P3","short":"FR","description":"20 ft Flat Rack","group":"Flat Rack"},
        {"key":"45P3","short":"FR","description":"40 ft Flat Rack","group":"Flat Rack"},
        {"key":"42U1","short":"OS","description":"40 ft Open Top Standard","group":"Open Top"},
        {"key":"42P3","short":"FS","description":"40 ft Flat Rack Standard","group":"Flat Rack"}
    ],
    "Другие значения": []
}

def short_container_type(iso, reference_data=None):
    iso = clean(iso).upper().replace(" ", "")
    data = (reference_data or DEFAULT_REFERENCE_DATA).get("Типы ISO", [])
    for item in data:
        if clean(item.get("key")).upper() == iso and clean(item.get("short")):
            return clean(item.get("short")).upper()
    return iso[:2] if iso else "?"



def teu_for_record(record):
    """Project convention: 20-foot = 1 TEU; 40/45-foot = 2 TEU."""
    iso = clean(record.get("ISO_TYPE")).upper().replace(" ", "")
    if iso.startswith("2") or iso.startswith("20"):
        return 1
    if iso.startswith(("4", "L5", "45", "40")):
        return 2
    # Fall back to explicitly stated size tokens found in atypical descriptions.
    if "20" in iso and "40" not in iso and "45" not in iso:
        return 1
    if "40" in iso or "45" in iso:
        return 2
    return 0


def is_twenty_foot(record):
    return teu_for_record(record) == 1


def position_footprint(record):
    """Return occupied odd-Bay slots using Ship Cont CAD CV's 20/40ft convention."""
    rec = normalize_record(record)
    bay, row, tier = num(rec.get("BAY"), -1), num(rec.get("ROW"), -1), num(rec.get("TIER"), -1)
    if bay < 1 or row < 0 or tier < 0:
        return []
    if teu_for_record(rec) == 2 and bay % 2 == 0:
        return [(bay - 1, row, tier), (bay + 1, row, tier)]
    return [(bay, row, tier)]


def detailed_validation_issues(records, available_bays=None, available_rows=None, available_tiers=None, reference_data=None):
    """Check records for missing data, invalid coordinates/ISO codes and occupied-slot overlaps."""
    issues = []
    seen_numbers = {}
    slot_owners = {}
    available_bays = set(available_bays or [])
    available_rows = set(available_rows or [])
    available_tiers = set(available_tiers or [])
    iso_reference = {clean(x.get("key")).upper() for x in (reference_data or DEFAULT_REFERENCE_DATA).get("Типы ISO", []) if clean(x.get("key"))}

    def add(severity, message, index, record=None):
        rec = normalize_record(record or {})
        try:
            location = f"Bay {int(float(rec.get('BAY'))):02d} / Row {int(float(rec.get('ROW'))):02d} / Tier {int(float(rec.get('TIER'))):02d}"
        except (TypeError, ValueError):
            location = "Координаты не заданы"
        issues.append({"severity": severity, "message": message, "index": index,
                       "container_no": clean(rec.get("CONTAINER_NO")) or "(без номера)", "location": location})

    for i, raw in enumerate(records):
        rec = normalize_record(raw)
        cn = clean(rec.get("CONTAINER_NO"))
        prefix = cn or f"Строка {i + 1}"
        if not cn:
            add("Ошибка", f"Строка {i + 1}: не указан номер контейнера.", i, rec)
        else:
            key = cn.upper()
            if key in seen_numbers:
                add("Ошибка", f"Дубликат номера контейнера {cn}: строки {seen_numbers[key] + 1} и {i + 1}.", i, rec)
            else:
                seen_numbers[key] = i
            if not re.fullmatch(r"[A-Z]{3}[UJZ][0-9]{7}", key):
                add("Предупреждение", f"{cn}: номер отличается от стандартного шаблона ISO 6346 (4 буквы + 7 цифр); проверь формат.", i, rec)

        iso = clean(rec.get("ISO_TYPE")).upper().replace(" ", "")
        if not iso:
            add("Ошибка", f"{prefix}: не указан ISO-код контейнера.", i, rec)
        elif not re.fullmatch(r"[0-9]{2}[A-Z0-9]{2}", iso):
            add("Ошибка", f"{cn}: некорректный формат ISO-кода «{iso}» (ожидается 4 символа, например 22G1).", i, rec)
        elif iso_reference and iso not in iso_reference:
            add("Предупреждение", f"{cn}: ISO-код {iso} не найден в справочнике приложения. Проверь код или добавь его в разделе «Данные». Он может быть допустимым, но не распознан текущим справочником.", i, rec)

        fe = clean(rec.get("F_E")).upper()
        if fe and fe not in {"F", "E"}:
            add("Предупреждение", f"{prefix}: неподдерживаемое значение F_E={fe!r}; ожидается F (гружёный) или E (порожний).", i, rec)

        coords = {}
        for field in ("BAY", "ROW", "TIER"):
            value = clean(rec.get(field))
            if not value:
                add("Ошибка", f"{prefix}: не заполнено обязательное поле {field}.", i, rec)
                coords[field] = None
                continue
            try:
                n = int(float(value.replace(",", ".")))
                coords[field] = n
                if (field == "BAY" and n <= 0) or (field != "BAY" and n < 0):
                    add("Ошибка", f"{prefix}: недопустимое значение координаты {field}={n}.", i, rec)
            except (ValueError, TypeError):
                coords[field] = None
                add("Ошибка", f"{prefix}: координата {field} должна быть целым числом.", i, rec)

        weight = clean(rec.get("GROSS_KG"))
        if not weight:
            add("Предупреждение", f"{prefix}: не указан вес брутто (GROSS_KG).", i, rec)
        else:
            try:
                w = float(weight.replace(",", "."))
                if w < 0:
                    add("Ошибка", f"{prefix}: вес брутто не может быть отрицательным.", i, rec)
                elif w == 0:
                    add("Предупреждение", f"{prefix}: вес брутто равен 0 кг.", i, rec)
            except (ValueError, TypeError):
                add("Ошибка", f"{prefix}: GROSS_KG должен быть числом в килограммах.", i, rec)
        if not clean(rec.get("POL")):
            add("Предупреждение", f"{prefix}: не указан порт погрузки POL.", i, rec)
        if not clean(rec.get("POD")):
            add("Предупреждение", f"{prefix}: не указан порт назначения/выгрузки POD.", i, rec)

        bay, row, tier = coords.get("BAY"), coords.get("ROW"), coords.get("TIER")
        if bay is not None and bay > 0:
            size = teu_for_record(rec)
            if size == 1 and bay % 2 == 0:
                add("Ошибка", f"{cn}: 20-футовый контейнер должен быть записан на нечётном Bay, а не Bay {bay:02d}.", i, rec)
            elif size == 2 and bay % 2 == 1:
                add("Ошибка", f"{cn}: 40/45-футовый контейнер должен быть записан на чётном центральном Bay; проверь Bay {bay:02d}.", i, rec)
            elif size == 0 and iso:
                add("Предупреждение", f"{cn}: по ISO-коду {iso} не удалось определить размер для проверки размещения.", i, rec)
            if available_bays and size in (1, 2):
                expected = [bay] if size == 1 else ([bay - 1, bay + 1] if bay % 2 == 0 else [bay])
                missing = [b for b in expected if b not in available_bays]
                if missing:
                    add("Предупреждение", f"{cn}: Bay-позиция(и) {', '.join(f'{b:02d}' for b in missing)} отсутствуют в загруженном скелете судна.", i, rec)
        if row is not None and available_rows and row not in available_rows:
            add("Предупреждение", f"{cn}: Row {row:02d} отсутствует в загруженном скелете судна.", i, rec)
        if tier is not None and available_tiers and tier not in available_tiers:
            add("Предупреждение", f"{cn}: Tier {tier:02d} отсутствует в загруженном скелете судна.", i, rec)

        for slot in position_footprint(rec):
            if slot in slot_owners:
                other_i, other_cn = slot_owners[slot]
                if other_i != i:
                    add("Ошибка", f"Конфликт занятости позиции Bay {slot[0]:02d} / Row {slot[1]:02d} / Tier {slot[2]:02d}: {other_cn} и {cn or '(без номера)'}.", i, rec)
            else:
                slot_owners[slot] = (i, cn or f"строка {i + 1}")
    return issues


def statistics_rows(records, group_field=None):
    """Return grouped statistics: label, container count, TEU, gross kg, 20ft, 40/45ft."""
    groups = {}
    for rec0 in records:
        rec = normalize_record(rec0)
        if not clean(rec.get("CONTAINER_NO")):
            continue
        if group_field is None:
            key = "Все контейнеры"
        elif group_field == "__SIZE__":
            size_teu = teu_for_record(rec)
            key = "20 ft (1 TEU)" if size_teu == 1 else ("40/45 ft (2 TEU)" if size_teu == 2 else "Тип не определён")
        elif group_field == "BAY":
            key = f"Bay {num(rec.get('BAY')):02d}" if clean(rec.get("BAY")) else "(Bay не указан)"
        elif group_field == "TIER":
            key = f"Tier {num(rec.get('TIER')):02d}" if clean(rec.get("TIER")) else "(Tier не указан)"
        else:
            key = clean(rec.get(group_field)) or "(не указан)"
        g = groups.setdefault(key, {"count":0, "teu":0, "kg":0, "20":0, "40":0})
        g["count"] += 1
        g["teu"] += teu_for_record(rec)
        g["kg"] += num(rec.get("GROSS_KG"), 0)
        if teu_for_record(rec) == 1: g["20"] += 1
        elif teu_for_record(rec) == 2: g["40"] += 1
    if not groups and group_field is None:
        groups["Все контейнеры"] = {"count":0, "teu":0, "kg":0, "20":0, "40":0}
    result = [(k, v["count"], v["teu"], v["kg"], v["20"], v["40"]) for k, v in groups.items()]
    if group_field in ("BAY", "TIER"):
        def numeric_tail(label):
            m = re.search(r"(\d+)", label)
            return int(m.group(1)) if m else 999999
        result.sort(key=lambda x: numeric_tail(x[0]))
    elif group_field:
        result.sort(key=lambda x: x[0].casefold())
    return result


def container_validation_status(record, records):
    rec = normalize_record(record)
    cn = clean(rec.get("CONTAINER_NO"))
    if not cn or not clean(rec.get("ISO_TYPE")):
        return "Есть ошибки"
    coords = tuple(num(rec.get(k), -1) for k in ("BAY", "ROW", "TIER"))
    if min(coords) < 0 or any(clean(rec.get(k)) == "" for k in ("BAY", "ROW", "TIER")):
        return "Есть ошибки"
    for other in records:
        if other is record:
            continue
        oth = normalize_record(other)
        if cn == clean(oth.get("CONTAINER_NO")):
            return "Есть ошибки"
        if coords == tuple(num(oth.get(k), -2) for k in ("BAY", "ROW", "TIER")) and cn:
            return "Есть ошибки"
    return "Проверка пройдена"


class ContainerDialog(tk.Toplevel):
    def __init__(self, master, record=None, title="Контейнер"):
        super().__init__(master)
        self.title(title)
        self.result = None
        self.transient(master)
        self.grab_set()
        self.resizable(True, False)
        self.vars = {}
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)
        for i, field in enumerate(FIELDS):
            row = i // 2
            col = (i % 2) * 2
            ttk.Label(frame, text=FIELD_LABELS.get(field, field)).grid(row=row, column=col, sticky="w", padx=(0, 8), pady=4)
            var = tk.StringVar(value=clean((record or {}).get(field, "")))
            self.vars[field] = var
            entry = ttk.Entry(frame, textvariable=var, width=27)
            entry.grid(row=row, column=col+1, sticky="ew", padx=(0, 18), pady=4)
        for c in range(4): frame.columnconfigure(c, weight=1)
        buttons = ttk.Frame(frame); buttons.grid(row=(len(FIELDS)+1)//2, column=0, columnspan=4, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="Отмена", command=self.destroy).pack(side="right", padx=(8,0))
        ttk.Button(buttons, text="Сохранить", command=self.save).pack(side="right")
        self.bind("<Escape>", lambda _e: self.destroy())
        self.update_idletasks()
        self.geometry(f"900x{min(700, self.winfo_reqheight()+10)}")
        self.focus_force()

    def save(self):
        data = {k: v.get().strip() for k, v in self.vars.items()}
        if not data["CONTAINER_NO"]:
            messagebox.showwarning("Проверка", "Укажи номер контейнера.", parent=self); return
        for k in ("BAY", "ROW", "TIER", "GROSS_KG"):
            if data[k] != "":
                try: int(float(data[k].replace(",", ".")))
                except ValueError:
                    messagebox.showwarning("Проверка", f"Поле {k} должно быть числом.", parent=self); return
        self.result = normalize_record(data)
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        ensure_folders()
        self.config_data = config_load()
        # In-memory caches avoid reopening and reparsing vessel Excel files on every redraw.
        self._skeleton_positions_cache = {}
        self._skeleton_legacy_cache = {}
        self._skeleton_metadata_cache = {}
        if not isinstance(self.config_data.get("reference_data"), dict):
            self.config_data["reference_data"] = json.loads(json.dumps(DEFAULT_REFERENCE_DATA, ensure_ascii=False))
        else:
            for category, defaults in DEFAULT_REFERENCE_DATA.items():
                if category not in self.config_data["reference_data"]:
                    self.config_data["reference_data"][category] = json.loads(json.dumps(defaults, ensure_ascii=False))
        # Migrate a previously selected single skeleton into the persistent vessel registry.
        existing_skeleton = clean(self.config_data.get("skeleton_path", ""))
        if existing_skeleton and Path(existing_skeleton).exists():
            current_entries = self.config_data.get("skeletons", [])
            if not isinstance(current_entries, list): current_entries = []
            if not any(isinstance(e, dict) and clean(e.get("path", "")) and Path(clean(e.get("path", ""))).exists() and Path(clean(e.get("path", ""))).resolve() == Path(existing_skeleton).resolve() for e in current_entries):
                current_entries.append({"name": Path(existing_skeleton).stem, "imo": "", "path": str(Path(existing_skeleton).resolve()), "added_at": datetime.now().isoformat(timespec="seconds")})
            self.config_data["skeletons"] = current_entries
        config_save(self.config_data)
        self.vessel = dict(DEFAULT_VESSEL)
        self.records = []
        self.source_records = []
        self.source_vessel = dict(DEFAULT_VESSEL)
        self.source_label = "Исходный план"
        self.current_path = None
        self.active_project_id = ""
        self.active_project_name = ""
        self._creating_project = False
        self._loading_active_project = False
        self.raw_edi = ""
        self.dirty = False
        self._search_match_indices = set()
        self.move_mode_active = False
        self.move_record_index = None
        self.search_field = tk.StringVar(value="Все поля")
        self.search_value = tk.StringVar(value="")
        self.move_mode_label = tk.StringVar(value="Перемещение: выключено")
        self.mode = tk.StringVar(value="Complex")
        self.status = tk.StringVar(value="Готово к работе")
        self.current_display = tk.StringVar(value="Документ не открыт")
        self.tier_choice = tk.StringVar(value="")
        self.plan_view_mode = tk.StringVar(value="Tier")
        self.plan_color_by = tk.StringVar(value=COLOR_MODE_OPTIONS[0])
        self.plan_detail_mode = tk.StringVar(value="Full Plan")
        self.skeleton_only_mode = tk.BooleanVar(value=False)
        self.plan_short_top = tk.StringVar(value="Тип контейнера")
        self.reference_category = tk.StringVar(value="Типы ISO")
        self.reference_editor_vars = {}
        self._plan_color_lookup = {}
        self.filter_field = tk.StringVar(value="Все поля")
        self.filter_value = tk.StringVar(value="")
        self._active_filter = ("", "")
        self.stats_report_type = tk.StringVar(value="Общая сводка")
        self.stats_summary = tk.StringVar(value="Контейнеров: 0   •   TEU: 0   •   Вес: 0 кг")
        self.plan_selection_label = tk.StringVar(value="Tier:")
        self.export_source_display = tk.StringVar(value="Нет открытого документа. Импортируйте EDI/Excel или создайте документ.")
        self._suppress_tree_event = False
        self.title(APP_NAME)
        self.geometry("1450x900")
        self.minsize(1100, 720)
        self.configure(bg="#EAF3F9")
        try:
            self.iconbitmap(str(resource_path("app.ico")))
        except Exception:
            pass
        self._setup_style()
        self._build_ui()
        self.refresh_projects()
        self.refresh_active_projects()
        self.refresh_samples()
        self.refresh_editor()
        self.refresh_plan()
        self.refresh_stats()
        self.protocol("WM_DELETE_WINDOW", self.close_app)
        self.after(800, self.offer_recovery)
        self.after(10000, self.autosave_tick)

    def _setup_style(self):
        style = ttk.Style(self)
        try: style.theme_use("clam")
        except Exception: pass
        bg = "#EAF3F9"
        navy = "#063B66"
        blue = "#087EBD"
        teal = "#13A6B5"
        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg, foreground="#163B59", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 23, "bold"), foreground=navy, background=bg)
        style.configure("SubTitle.TLabel", font=("Segoe UI", 10), foreground="#52738B", background=bg)
        style.configure("TButton", padding=(11, 8), font=("Segoe UI", 9), background="#DCECF7", foreground=navy, borderwidth=0)
        style.map("TButton", background=[("active", "#C4E4F5"), ("pressed", "#A8D9EF")], foreground=[("disabled", "#91A8B7")])
        style.configure("Accent.TButton", padding=(14, 9), font=("Segoe UI", 9, "bold"), background=blue, foreground="#FFFFFF")
        style.map("Accent.TButton", background=[("active", "#0B98D8"), ("pressed", "#05699F")], foreground=[("active", "#FFFFFF")])
        style.configure("TLabelframe", background="#F8FCFF", bordercolor="#C5DDEB", padding=12)
        style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground=navy, background=bg)
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=30, background="#FFFFFF", fieldbackground="#FFFFFF", foreground="#193B53", borderwidth=0)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#D6EAF6", foreground=navy, padding=(8, 7))
        style.map("Treeview", background=[("selected", "#B9E3F6")], foreground=[("selected", navy)])
        style.configure("TNotebook", background=bg, borderwidth=0, tabmargins=(2, 4, 2, 0))
        style.configure("TNotebook.Tab", padding=(16, 10), font=("Segoe UI", 10, "bold"), background="#D9EAF5", foreground="#355A73")
        style.map("TNotebook.Tab", background=[("selected", "#FFFFFF"), ("active", "#C7E7F6")], foreground=[("selected", navy)])
        style.configure("TCombobox", padding=5)
        style.configure("TEntry", padding=6)

    def _build_ui(self):
        head = ttk.Frame(self, padding=(22, 16, 22, 10))
        head.pack(fill="x")
        left = ttk.Frame(head); left.pack(side="left", fill="x", expand=True)
        title_row = ttk.Frame(left); title_row.pack(anchor="w")
        try:
            self.logo_photo = tk.PhotoImage(file=str(resource_path("app_logo.png")))
            ttk.Label(title_row, image=self.logo_photo).pack(side="left", padx=(0, 10))
        except Exception:
            ttk.Label(title_row, text="🚢", font=("Segoe UI Emoji", 25), foreground="#087EBD").pack(side="left", padx=(0, 10))
        ttk.Label(title_row, text="SHIP CONT CAD CV", style="Title.TLabel").pack(side="left")
        ttk.Label(left, text="v2.3  •  Maritime workspace for container stowage, EDI & Excel", style="SubTitle.TLabel").pack(anchor="w", pady=(3,0))
        ttk.Label(head, textvariable=self.current_display, style="SubTitle.TLabel", wraplength=420, justify="right").pack(side="right", anchor="e")
        globalbar = ttk.Frame(self, padding=(20,0,20,8)); globalbar.pack(fill="x")
        ttk.Button(globalbar, text="＋ Новый документ", style="Accent.TButton", command=self.new_document).pack(side="left")
        ttk.Button(globalbar, text="Открыть…", command=self.open_dialog).pack(side="left", padx=(8,0))
        ttk.Button(globalbar, text="Сохранить", command=self.save_current).pack(side="left", padx=(8,0))
        ttk.Button(globalbar, text="Сохранить как…", command=self.save_as).pack(side="left", padx=(8,0))
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=20, pady=(0,10))
        self.tab_projects = ttk.Frame(self.nb, padding=14)
        self.tab_home = ttk.Frame(self.nb, padding=14)
        self.tab_samples = ttk.Frame(self.nb, padding=14)
        self.tab_plan = ttk.Frame(self.nb, padding=14)
        self.tab_editor = ttk.Frame(self.nb, padding=14)
        self.tab_import = ttk.Frame(self.nb, padding=14)
        self.tab_export = ttk.Frame(self.nb, padding=14)
        self.tab_stats = ttk.Frame(self.nb, padding=14)
        self.tab_data = ttk.Frame(self.nb, padding=14)
        self.tab_vessels = ttk.Frame(self.nb, padding=14)
        self.nb.add(self.tab_projects, text="Проекты")
        self.nb.add(self.tab_home, text="Загрузки")
        self.nb.add(self.tab_samples, text="Образцы")
        self.nb.add(self.tab_plan, text="План судна")
        self.nb.add(self.tab_vessels, text="Судна")
        self.nb.add(self.tab_editor, text="Редактор")
        self.nb.add(self.tab_import, text="Импорт")
        self.nb.add(self.tab_stats, text="Статистика")
        self.nb.add(self.tab_export, text="Сохранение")
        self.nb.add(self.tab_data, text="Справочник (Data)")
        self._build_active_projects()
        self._build_home()
        self._build_samples()
        self._build_plan()
        self._build_vessels()
        self._build_editor()
        self._build_import()
        self._build_stats()
        self._build_export()
        self._build_data()
        foot = ttk.Frame(self, padding=(22,4,22,12)); foot.pack(fill="x", side="bottom")
        ttk.Label(foot, textvariable=self.status, foreground="#176DA5").pack(side="left", fill="x", expand=True)
        ttk.Label(foot, text=f"Рабочая папка: {APP_HOME}", foreground="#526477").pack(side="right")

    def _build_active_projects(self):
        ttk.Label(self.tab_projects, text="Активные проекты судов", font=("Segoe UI", 16, "bold"), foreground="#063B66").pack(anchor="w")
        ttk.Label(self.tab_projects, text="Создавайте отдельный рабочий проект для каждого судна. Изменения контейнеров, данные судна и выбранный скелет автоматически сохраняются в проекте.", wraplength=1200).pack(anchor="w", pady=(4, 12))
        toolbar = ttk.Frame(self.tab_projects); toolbar.pack(fill="x", pady=(0, 10))
        ttk.Button(toolbar, text="＋ Создать проект", style="Accent.TButton", command=self.create_active_project).pack(side="left")
        ttk.Button(toolbar, text="Открыть / продолжить", command=self.open_active_project).pack(side="left", padx=(8,0))
        ttk.Button(toolbar, text="Сохранить проект", style="Accent.TButton", command=self.save_active_project_now).pack(side="left", padx=(8,0))
        ttk.Button(toolbar, text="Завершить проект", command=self.archive_active_project).pack(side="left", padx=(8,0))
        ttk.Button(toolbar, text="Удалить проект", command=self.delete_active_project).pack(side="left", padx=(8,0))
        ttk.Button(toolbar, text="Обновить", command=self.refresh_active_projects).pack(side="right")
        frame=ttk.Frame(self.tab_projects); frame.pack(fill="both", expand=True)
        cols=("name","source","vessel","containers","updated")
        self.active_project_tree=ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        for col,title,width in [("name","Название проекта",250),("source","Исходный файл",270),("vessel","Судно",210),("containers","Контейнеров",100),("updated","Последнее сохранение",190)]:
            self.active_project_tree.heading(col,text=title); self.active_project_tree.column(col,width=width,anchor="w")
        sy=ttk.Scrollbar(frame,orient="vertical",command=self.active_project_tree.yview)
        self.active_project_tree.configure(yscrollcommand=sy.set)
        self.active_project_tree.grid(row=0,column=0,sticky="nsew"); sy.grid(row=0,column=1,sticky="ns")
        frame.rowconfigure(0,weight=1); frame.columnconfigure(0,weight=1)
        self.active_project_tree.bind("<Double-1>", lambda _e:self.open_active_project())
        self.active_project_tree.bind("<<TreeviewSelect>>", self._on_active_project_selected)
        self.active_project_hint=tk.StringVar(value="Нет активных проектов. Нажмите «Создать проект», чтобы начать работу.")
        ttk.Label(self.tab_projects,textvariable=self.active_project_hint,foreground="#087EBD").pack(anchor="w",pady=(10,0))

    def _build_samples(self):
        ttk.Label(self.tab_samples,text="Хранилище образцов",font=("Segoe UI",16,"bold"),foreground="#063B66").pack(anchor="w")
        ttk.Label(self.tab_samples,text="Добавляйте готовые файлы для повторного использования. Файлы хранятся отдельно от активных проектов и не изменяются автоматически.",wraplength=1200).pack(anchor="w",pady=(4,12))
        frame=ttk.Frame(self.tab_samples); frame.pack(fill="both",expand=True)
        cols=("name","type","modified","size")
        self.samples_tree=ttk.Treeview(frame,columns=cols,show="headings",selectmode="browse")
        for col,title,width in [("name","Файл образца",400),("type","Формат",100),("modified","Добавлен / изменён",180),("size","Размер",110)]:
            self.samples_tree.heading(col,text=title); self.samples_tree.column(col,width=width,anchor="w")
        sy=ttk.Scrollbar(frame,orient="vertical",command=self.samples_tree.yview)
        self.samples_tree.configure(yscrollcommand=sy.set)
        self.samples_tree.grid(row=0,column=0,sticky="nsew"); sy.grid(row=0,column=1,sticky="ns")
        frame.rowconfigure(0,weight=1); frame.columnconfigure(0,weight=1)
        self.samples_tree.bind("<Double-1>",lambda _e:self.open_selected_sample())
        controls=ttk.Frame(self.tab_samples); controls.pack(fill="x",pady=(10,0))
        ttk.Button(controls,text="＋ Добавить файл…",style="Accent.TButton",command=self.add_sample).pack(side="left")
        ttk.Button(controls,text="Открыть / использовать",command=self.open_selected_sample).pack(side="left",padx=(8,0))
        ttk.Button(controls,text="Открыть папку",command=self.open_samples_folder).pack(side="left",padx=(8,0))
        ttk.Button(controls,text="Удалить из образцов",command=self.delete_selected_sample).pack(side="left",padx=(8,0))
        ttk.Button(controls,text="Обновить",command=self.refresh_samples).pack(side="right")

    def refresh_active_projects(self):
        if not hasattr(self,"active_project_tree"): return
        for item in self.active_project_tree.get_children(): self.active_project_tree.delete(item)
        ACTIVE_PROJECTS_DIR.mkdir(parents=True,exist_ok=True)
        entries=[]
        for path in ACTIVE_PROJECTS_DIR.glob("*.json"):
            try:
                data=json.loads(path.read_text(encoding="utf-8"))
                if data.get("archived"): continue
                entries.append((path,data))
            except Exception: continue
        entries.sort(key=lambda x: clean(x[1].get("updated_at")),reverse=True)
        for path,data in entries:
            vessel=data.get("vessel",{}) or {}
            updated=clean(data.get("updated_at")) or "—"
            try: updated=datetime.fromisoformat(updated).strftime("%Y-%m-%d %H:%M")
            except Exception: pass
            self.active_project_tree.insert("","end",iid=str(path),values=(data.get("project_name",path.stem),data.get("source_name",Path(clean(data.get("current_path"))).name if data.get("current_path") else "Новый документ"),vessel.get("VESSEL_NAME","TBN"),len(data.get("records",[])),updated))
        current_path = self._active_project_path() if self.active_project_id else None
        if current_path and current_path.exists() and str(current_path) in self.active_project_tree.get_children():
            self.active_project_tree.selection_set(str(current_path)); self.active_project_tree.focus(str(current_path))
        if hasattr(self,"active_project_hint"):
            self.active_project_hint.set(f"Активных проектов: {len(entries)}  •  Текущий рабочий проект: {self.active_project_name or 'Нулевой проект (без сохранения)'}  •  Состояние сохраняется автоматически")

    def _on_active_project_selected(self, _event=None):
        if getattr(self, "_loading_active_project", False): return
        selected = self.active_project_tree.selection() if hasattr(self, "active_project_tree") else ()
        if selected:
            try:
                data=json.loads(Path(selected[0]).read_text(encoding="utf-8"))
                self.active_project_hint.set(f"Выбран проект: {clean(data.get('project_name')) or Path(selected[0]).stem}  •  Нажмите «Открыть / продолжить», чтобы переключить рабочий проект.")
            except Exception: pass

    def _active_project_path(self, project_id=None):
        pid=project_id or self.active_project_id
        return ACTIVE_PROJECTS_DIR / f"{pid}.json" if pid else None

    def _write_active_project_snapshot(self):
        """Persist the complete editable state for the current active project."""
        if not self.active_project_id:
            return False
        try:
            ensure_folders()
            path = self._active_project_path()
            if path is None:
                return False
            current_tab = "План судна"
            try:
                current_tab = self.nb.tab(self.nb.select(), "text")
            except Exception:
                pass
            payload = {
                "app_version": "2.3",
                "project_id": self.active_project_id,
                "project_name": self.active_project_name or self.active_project_id,
                "source_name": self.current_path.name if self.current_path else self.source_label,
                "current_path": str(self.current_path) if self.current_path else "",
                "vessel": self.vessel,
                "records": [normalize_record(r) for r in self.records],
                "source_records": [normalize_record(r) for r in self.source_records],
                "source_vessel": self.source_vessel,
                "source_label": self.source_label,
                "raw_edi": self.raw_edi,
                "skeleton_path": clean(self.config_data.get("skeleton_path", "")),
                "last_tab": current_tab,
                "plan_view_mode": self.plan_view_mode.get(),
                "plan_detail_mode": self.plan_detail_mode.get(),
                "tier_choice": self.tier_choice.get(),
                "plan_color_by": self.plan_color_by.get(),
                "skeleton_only_mode": bool(self.skeleton_only_mode.get()),
                "search_field": self.search_field.get(),
                "search_value": self.search_value.get(),
                "filter_field": self.filter_field.get(),
                "filter_value": self.filter_value.get(),
                "updated_at": datetime.now().isoformat(timespec="seconds"),
                "archived": False,
            }
            tmp = path.with_suffix(".tmp")
            tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
            tmp.replace(path)
            self.refresh_active_projects()
            return True
        except Exception as e:
            self.set_status(f"ОШИБКА сохранения проекта: {e}")
            try:
                messagebox.showerror("Сохранение проекта", f"Не удалось сохранить рабочее состояние проекта:\n{e}", parent=self)
            except Exception:
                pass
            return False

    def save_active_project_now(self):
        """Explicit save command for the active project's resumable state."""
        if not self.active_project_id:
            messagebox.showinfo("Сохранение проекта", "Сначала создайте или откройте проект на вкладке «Проекты».", parent=self)
            return
        if self._write_active_project_snapshot():
            self.dirty = False
            self.write_recovery_snapshot()
            self.set_status(f"Проект «{self.active_project_name}» сохранён. Можно закрыть программу и продолжить позже.")

    def create_active_project(self):
        name=simpledialog.askstring("Новый проект","Введите название проекта (например, Barge UGL — рейс 24):",parent=self)
        if not name or not name.strip(): return
        use_file=messagebox.askyesno("Исходный файл","Создать проект на основе существующего EDI/Excel-файла?\n\nДа — выбрать файл.\nНет — создать пустой проект и загрузить файл позже.",parent=self)
        source_path=None
        if use_file:
            chosen=filedialog.askopenfilename(title="Выберите исходный EDI или Excel-файл",initialdir=str(PROJECTS_DIR),filetypes=SUPPORTED_OPEN)
            if chosen: source_path=Path(chosen)
            elif not messagebox.askyesno("Без исходного файла","Файл не выбран. Создать пустой проект?",parent=self): return
        if self.active_project_id and not self._write_active_project_snapshot(): return
        if not self.active_project_id and self.dirty:
            if not messagebox.askyesno("Несохранённые изменения","Текущий нулевой проект содержит изменения. Создать отдельный проект и оставить текущие данные без привязки?",parent=self): return
        self._creating_project=True
        try:
            if source_path:
                self.dirty=False; self.load_document(source_path)
                if not getattr(self,"_last_load_succeeded",False): return
            else:
                self.vessel=dict(DEFAULT_VESSEL); self.records=[]; self.source_records=[]
                self.source_vessel=dict(DEFAULT_VESSEL); self.source_label="Пустой проект"
                self.raw_edi=""; self.current_path=None; self.dirty=False
                self._search_match_indices=set(); self.move_mode_active=False; self.move_record_index=None
                self.config_data["skeleton_path"]=""; config_save(self.config_data)
            slug=re.sub(r"[^A-Za-z0-9_-]+","_",name.strip()).strip("_")[:48] or "Project"
            self.active_project_id=f"{slug}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"; self.active_project_name=name.strip(); self.dirty=True
            if not self._write_active_project_snapshot():
                self.set_status("Проект создан в памяти, но состояние не удалось сохранить. Проверьте папку ActiveProjects."); return
            self.dirty=False; self.refresh_editor(); self.refresh_plan(); self.refresh_stats(); self.update_current_display(); self.refresh_active_projects()
            self.nb.select(self.tab_editor); self.set_status(f"Создан и сохранён активный проект: {self.active_project_name}")
        finally:
            self._creating_project=False; self.refresh_active_projects()

    def open_active_project(self):
        selected=self.active_project_tree.selection() if hasattr(self,"active_project_tree") else ()
        if selected: path=Path(selected[0])
        elif self.active_project_id: path=self._active_project_path()
        else:
            messagebox.showinfo("Проекты","Выберите проект в списке или создайте новый.",parent=self); return
        if not path or not path.exists():
            messagebox.showerror("Открытие проекта","Файл состояния проекта не найден. Обновите список или проверьте папку ActiveProjects.",parent=self); return
        target_id=path.stem
        if self.active_project_id and self.active_project_id != target_id:
            if not self._write_active_project_snapshot(): return
        elif self.active_project_id and self.active_project_id == target_id and self.dirty:
            if not self._write_active_project_snapshot(): return
        elif self.dirty:
            if not messagebox.askyesno("Несохранённые изменения","Текущее состояние не привязано к проекту. Продолжить открытие проекта?",parent=self): return
        try:
            data=json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict) or not isinstance(data.get("records", []), list):
                raise ValueError("Файл проекта повреждён или имеет неподдерживаемый формат.")
            self._loading_active_project=True
            self.active_project_id=clean(data.get("project_id")) or path.stem
            self.active_project_name=clean(data.get("project_name")) or path.stem
            self.vessel=dict(DEFAULT_VESSEL); self.vessel.update(data.get("vessel",{}))
            self.records=[normalize_record(r) for r in data.get("records",[])]
            self.source_records=[normalize_record(r) for r in data.get("source_records",self.records)]
            self.source_vessel=dict(data.get("source_vessel",self.vessel)); self.source_label=clean(data.get("source_label")) or "Сохранённое состояние проекта"
            self.raw_edi=clean(data.get("raw_edi"))
            cp=clean(data.get("current_path")); self.current_path=Path(cp) if cp and Path(cp).exists() else None
            sp=clean(data.get("skeleton_path"))
            self.config_data["skeleton_path"]=sp if sp and Path(sp).exists() else ""
            config_save(self.config_data)
            self.plan_view_mode.set(clean(data.get("plan_view_mode")) or "Tier")
            self.plan_detail_mode.set(clean(data.get("plan_detail_mode")) or "Full Plan")
            self.tier_choice.set(clean(data.get("tier_choice")))
            self.plan_color_by.set(clean(data.get("plan_color_by")) or COLOR_MODE_OPTIONS[0])
            self.skeleton_only_mode.set(bool(data.get("skeleton_only_mode",False)))
            if data.get("search_field"): self.search_field.set(data.get("search_field"))
            self.search_value.set(clean(data.get("search_value")))
            if data.get("filter_field"): self.filter_field.set(data.get("filter_field"))
            self.filter_value.set(clean(data.get("filter_value")))
            self.dirty=False; self._search_match_indices=set(); self.move_mode_active=False; self.move_record_index=None
            self.refresh_editor(); self.refresh_plan(); self.refresh_stats(); self.update_current_display(); self.refresh_projects(); self.refresh_active_projects()
            self.nb.select(self.tab_projects)
            wanted=clean(data.get("last_tab"))
            for tab in self.nb.tabs():
                if self.nb.tab(tab,"text")==wanted:
                    self.nb.select(tab); break
            self.set_status(f"Продолжена работа над проектом: {self.active_project_name}")
        except Exception as e:
            messagebox.showerror("Открытие проекта",f"Не удалось открыть проект:\n{e}",parent=self)
        finally:
            self._loading_active_project=False
            self.refresh_active_projects()

    def archive_active_project(self):
        if not self.active_project_id:
            messagebox.showinfo("Проекты","Сначала откройте проект, который нужно завершить.",parent=self); return
        self._write_active_project_snapshot()
        src=self._active_project_path()
        if src and src.exists():
            try:
                data=json.loads(src.read_text(encoding="utf-8")); data["archived"]=True; data["updated_at"]=datetime.now().isoformat(timespec="seconds")
                archived=ARCHIVED_PROJECTS_DIR/f"{self.active_project_id}.json"
                archived.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8")
                src.unlink()
                name=self.active_project_name; self.active_project_id=""; self.active_project_name=""
                self.set_status(f"Проект завершён и перемещён в архив: {name}"); self.refresh_active_projects()
            except Exception as e: messagebox.showerror("Завершение проекта",str(e),parent=self)

    def delete_active_project(self):
        selected=self.active_project_tree.selection() if hasattr(self,"active_project_tree") else ()
        if not selected:
            messagebox.showinfo("Проекты","Выберите проект в списке.",parent=self); return
        path=Path(selected[0])
        if not messagebox.askyesno("Удалить проект",f"Удалить сохранённое состояние проекта «{path.stem}»? Исходный EDI/Excel-файл не удаляется.",parent=self): return
        try:
            path.unlink(missing_ok=True)
            if self.active_project_id and path.stem==self.active_project_id:
                self.active_project_id=""; self.active_project_name=""
            self.refresh_active_projects()
        except Exception as e: messagebox.showerror("Удаление проекта",str(e),parent=self)

    def refresh_samples(self):
        if not hasattr(self,"samples_tree"): return
        for item in self.samples_tree.get_children(): self.samples_tree.delete(item)
        SAMPLES_DIR.mkdir(parents=True,exist_ok=True)
        for p in sorted((x for x in SAMPLES_DIR.iterdir() if x.is_file()),key=lambda x:x.name.lower()):
            try:
                st=p.stat(); modified=datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"); size=f"{st.st_size/1024:.1f} KB"
            except Exception: modified="—"; size="—"
            self.samples_tree.insert("","end",iid=str(p),values=(p.name,p.suffix.lower().lstrip(".") or "файл",modified,size))

    def add_sample(self):
        p=filedialog.askopenfilename(title="Добавить файл в хранилище образцов",filetypes=[("Все файлы","*.*")])
        if not p: return
        src=Path(p).resolve(); target=SAMPLES_DIR/src.name
        if src==target.resolve(): self.refresh_samples(); return
        if target.exists():
            if not messagebox.askyesno("Файл уже существует",f"Файл {target.name} уже есть. Заменить его?",parent=self): return
        try:
            shutil.copy2(src,target); self.refresh_samples(); self.set_status(f"Файл добавлен в образцы: {target.name}")
        except Exception as e: messagebox.showerror("Добавление образца",str(e),parent=self)

    def _selected_sample_path(self):
        sel=self.samples_tree.selection() if hasattr(self,"samples_tree") else ()
        return Path(sel[0]) if sel else None

    def open_selected_sample(self):
        p=self._selected_sample_path()
        if not p or not p.exists():
            messagebox.showinfo("Образцы","Выберите файл в списке.",parent=self); return
        if p.suffix.lower() in (".edi",".txt",".baplie",".xlsx",".xlsm"):
            self.load_document(p)
        else:
            try:
                if os.name=="nt": os.startfile(str(p))
                elif sys.platform=="darwin": os.system(f'open "{p}"')
                else: os.system(f'xdg-open "{p}"')
            except Exception as e: messagebox.showerror("Открытие образца",str(e),parent=self)

    def delete_selected_sample(self):
        p=self._selected_sample_path()
        if not p or not p.exists():
            messagebox.showinfo("Образцы","Выберите файл.",parent=self); return
        if not messagebox.askyesno("Удалить образец",f"Удалить копию «{p.name}» из хранилища? Исходный файл не изменится.",parent=self): return
        try: p.unlink(); self.refresh_samples()
        except Exception as e: messagebox.showerror("Удаление образца",str(e),parent=self)

    def open_samples_folder(self):
        try:
            if os.name=="nt": os.startfile(str(SAMPLES_DIR))
            elif sys.platform=="darwin": os.system(f'open "{SAMPLES_DIR}"')
            else: os.system(f'xdg-open "{SAMPLES_DIR}"')
        except Exception as e: messagebox.showerror("Папка образцов",f"{SAMPLES_DIR}\n\n{e}",parent=self)

    def _build_home(self):
        ttk.Label(self.tab_home, text="Рабочие документы", font=("Segoe UI", 15, "bold"), foreground="#17365D").pack(anchor="w")
        ttk.Label(self.tab_home, text="Список файлов хранится отдельно от EXE. Открой проект отсюда или через Проводник Windows.").pack(anchor="w", pady=(3,10))
        split = ttk.Frame(self.tab_home); split.pack(fill="both", expand=True)
        cols = ("name", "ext", "modified", "size")
        self.project_tree = ttk.Treeview(split, columns=cols, show="headings", selectmode="browse")
        for col, title, width in [("name","Файл",420),("ext","Формат",90),("modified","Изменён",170),("size","Размер",110)]:
            self.project_tree.heading(col, text=title); self.project_tree.column(col, width=width, anchor="w")
        sy = ttk.Scrollbar(split, orient="vertical", command=self.project_tree.yview)
        sx = ttk.Scrollbar(split, orient="horizontal", command=self.project_tree.xview)
        self.project_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.project_tree.grid(row=0,column=0,sticky="nsew"); sy.grid(row=0,column=1,sticky="ns"); sx.grid(row=1,column=0,sticky="ew")
        split.rowconfigure(0,weight=1); split.columnconfigure(0,weight=1)
        self.project_tree.bind("<Double-1>", lambda _e: self.open_selected_project())
        controls = ttk.Frame(self.tab_home); controls.pack(fill="x", pady=(10,14))
        ttk.Button(controls, text="Открыть выбранный", command=self.open_selected_project).pack(side="left")
        ttk.Button(controls, text="Обновить список", command=self.refresh_projects).pack(side="left", padx=(8,0))
        ttk.Button(controls, text="Удалить выбранный", command=self.delete_selected_project).pack(side="left", padx=(8,0))
        ttk.Button(controls, text="Восстановить резервную копию…", command=self.restore_backup_dialog).pack(side="left", padx=(8,0))
        ttk.Button(controls, text="Открыть папку резервных копий", command=self.open_backups_folder).pack(side="left", padx=(8,0))
        settings = ttk.LabelFrame(self.tab_home, text="Настройки шаблонов и скелетов", padding=12); settings.pack(fill="x")
        tpath = clean(self.config_data.get("template_path", "")) or "Встроенный шаблон Excel (VESSEL + CONTAINERS)"
        spath = clean(self.config_data.get("skeleton_path", "")) or "Скелет судна не выбран — общий план по Row / Tier"
        self.template_label = tk.StringVar(value=tpath)
        self.skeleton_label = tk.StringVar(value=spath)
        row1=ttk.Frame(settings); row1.pack(fill="x", pady=3)
        ttk.Button(row1, text="Выбрать образец Excel…", command=self.choose_template).pack(side="left")
        ttk.Label(row1, textvariable=self.template_label, wraplength=800).pack(side="left", padx=10)
        row2=ttk.Frame(settings); row2.pack(fill="x", pady=3)
        ttk.Button(row2, text="Загрузить скелет судна…", command=self.choose_skeleton).pack(side="left")
        ttk.Label(row2, textvariable=self.skeleton_label, wraplength=800).pack(side="left", padx=10)
        ttk.Label(settings, text="Excel-образец используется для импорта и пустого шаблона. Скелет хранится отдельно и не перезаписывается при работе над рейсом.", foreground="#526477", wraplength=1100).pack(anchor="w", pady=(8,0))

    def _build_plan(self):
        top=ttk.Frame(self.tab_plan); top.pack(fill="x", pady=(0,6))
        row1=ttk.Frame(top); row1.pack(fill="x")
        ttk.Label(row1, text="Графика:").pack(side="left")
        self.plan_mode_combo=ttk.Combobox(row1, textvariable=self.plan_view_mode, state="readonly", width=8, values=["Tier", "Bay"])
        self.plan_mode_combo.pack(side="left", padx=(6,12))
        self.plan_mode_combo.bind("<<ComboboxSelected>>", self.on_plan_mode_change)
        ttk.Label(row1, text="Режим:").pack(side="left", padx=(2,3))
        self.plan_detail_combo=ttk.Combobox(row1, textvariable=self.plan_detail_mode, state="readonly", width=12, values=["Full Plan", "Stowage Plan"])
        self.plan_detail_combo.pack(side="left", padx=(0,8))
        self.plan_detail_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_plan())
        ttk.Label(row1, text="Верхняя строка:").pack(side="left")
        self.plan_short_top_combo=ttk.Combobox(row1, textvariable=self.plan_short_top, state="readonly", width=12, values=["Тип контейнера", "Порт выгрузки"])
        self.plan_short_top_combo.pack(side="left", padx=(5,8))
        self.plan_short_top_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_plan())
        ttk.Label(row1, textvariable=self.plan_selection_label).pack(side="left")
        self.tier_combo=ttk.Combobox(row1, textvariable=self.tier_choice, state="readonly", width=11, values=["Все"])
        self.tier_combo.pack(side="left", padx=(6,10)); self.tier_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_plan())
        ttk.Label(row1, text="Раскраска:").pack(side="left", padx=(4,4))
        self.plan_color_combo = ttk.Combobox(row1, textvariable=self.plan_color_by, state="readonly", width=21, values=COLOR_MODE_OPTIONS)
        self.plan_color_combo.pack(side="left")
        self.plan_color_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_plan())
        ttk.Button(row1, text="Обновить", command=self.refresh_plan).pack(side="left", padx=(8,0))
        ttk.Button(row1, text="Показать скелет / контейнеры", command=self.choose_skeleton_for_plan).pack(side="left", padx=(6,0))

        row2=ttk.Frame(self.tab_plan); row2.pack(fill="x", pady=(5,8))
        ttk.Label(row2,text="Найти:").pack(side="left")
        self.plan_search_combo=ttk.Combobox(row2,textvariable=self.search_field,state="readonly",width=19,
            values=["Все поля","CONTAINER_NO","POD","POT","POL","LINE","ISO_TYPE","GROSS_KG","BAY","ROW","TIER","BOOKING_NO"])
        self.plan_search_combo.pack(side="left",padx=(6,4))
        self.plan_search_entry=ttk.Entry(row2,textvariable=self.search_value,width=26)
        self.plan_search_entry.pack(side="left")
        self.plan_search_entry.bind("<Return>",lambda _e:self.search_plan())
        ttk.Button(row2,text="Найти на схеме",command=self.search_plan).pack(side="left",padx=(6,0))
        ttk.Button(row2,text="Сбросить поиск",command=self.clear_plan_search).pack(side="left",padx=(6,0))
        ttk.Button(row2,textvariable=self.move_mode_label,command=self.toggle_move_mode).pack(side="left",padx=(14,0))
        ttk.Button(row2,text="Сравнить с исходным…",command=self.compare_with_source).pack(side="right")
        ttk.Label(self.tab_plan, text="Нажмите на контейнер, чтобы открыть параметры. Для перемещения включите режим, выберите контейнер, затем нажмите пустую позицию. Контейнер 40/45 ft в Bay 02 занимает позиции 01 и 03.", wraplength=1200, foreground="#526477").pack(anchor="w", pady=(0,6))
        self.plan_legend = ttk.Frame(self.tab_plan)
        self.plan_legend.pack(fill="x", pady=(0,6))
        frame=ttk.Frame(self.tab_plan); frame.pack(fill="both",expand=True)
        self.plan_canvas=tk.Canvas(frame, background="#FFFFFF", highlightthickness=1, highlightbackground="#C5D0D9")
        v=ttk.Scrollbar(frame,orient="vertical",command=self.plan_canvas.yview); h=ttk.Scrollbar(frame,orient="horizontal",command=self.plan_canvas.xview)
        self.plan_canvas.configure(yscrollcommand=v.set,xscrollcommand=h.set)
        self.plan_canvas.grid(row=0,column=0,sticky="nsew"); v.grid(row=0,column=1,sticky="ns"); h.grid(row=1,column=0,sticky="ew")
        frame.rowconfigure(0,weight=1); frame.columnconfigure(0,weight=1)

    def _build_vessels(self):
        ttk.Label(self.tab_vessels, text="Библиотека судов и скелетов", font=("Segoe UI", 15, "bold"), foreground="#17365D").pack(anchor="w")
        ttk.Label(self.tab_vessels, text="Загруженные скелеты копируются в рабочую папку приложения и остаются в перечне после перезапуска или обновления EXE.", wraplength=1100).pack(anchor="w", pady=(4, 12))
        frame = ttk.Frame(self.tab_vessels); frame.pack(fill="both", expand=True)
        cols = ("vessel", "imo", "positions", "file", "path")
        self.skeleton_tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        for col, title, width in [("vessel", "Судно", 220), ("imo", "IMO", 110), ("positions", "Позиций", 90), ("file", "Файл скелета", 260), ("path", "Путь хранения", 440)]:
            self.skeleton_tree.heading(col, text=title); self.skeleton_tree.column(col, width=width, anchor="w")
        sy = ttk.Scrollbar(frame, orient="vertical", command=self.skeleton_tree.yview)
        sx = ttk.Scrollbar(frame, orient="horizontal", command=self.skeleton_tree.xview)
        self.skeleton_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.skeleton_tree.grid(row=0, column=0, sticky="nsew"); sy.grid(row=0, column=1, sticky="ns"); sx.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1); frame.columnconfigure(0, weight=1)
        self.skeleton_tree.bind("<Double-1>", lambda _e: self.activate_selected_skeleton())
        controls = ttk.Frame(self.tab_vessels); controls.pack(fill="x", pady=(12, 8))
        ttk.Button(controls, text="＋ Загрузить скелет Excel…", command=self.choose_skeleton).pack(side="left")
        ttk.Button(controls, text="Использовать выбранное судно", command=self.activate_selected_skeleton).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Удалить из перечня…", command=self.remove_selected_skeleton).pack(side="left", padx=(8, 0))
        ttk.Button(controls, text="Обновить перечень", command=self.refresh_skeleton_registry).pack(side="left", padx=(8, 0))
        self.skeleton_registry_status = tk.StringVar(value="")
        ttk.Label(self.tab_vessels, textvariable=self.skeleton_registry_status, foreground="#526477", wraplength=1100).pack(anchor="w", pady=(4, 0))
        self.refresh_skeleton_registry()

    def _skeleton_fingerprint(self, path):
        """Cheap identity for cache invalidation when a skeleton file is replaced."""
        try:
            p = Path(path).resolve()
            st = p.stat()
            return (str(p), st.st_mtime_ns, st.st_size)
        except Exception:
            return (str(path), None, None)

    def _skeleton_metadata(self, path):
        p = Path(path)
        name = p.stem
        imo = ""
        positions = 0
        fingerprint = self._skeleton_fingerprint(p)
        cached = self._skeleton_metadata_cache.get(fingerprint)
        if cached is not None:
            return dict(cached)
        # Persisted metadata avoids opening every vessel workbook at each startup.
        for entry in self.config_data.get("skeletons", []):
            if not isinstance(entry, dict) or not clean(entry.get("path")):
                continue
            try:
                if Path(clean(entry["path"])).resolve() != p.resolve():
                    continue
                st = p.stat()
                if (entry.get("meta_mtime_ns") == st.st_mtime_ns and
                        entry.get("meta_size") == st.st_size and
                        all(k in entry for k in ("meta_name", "meta_imo", "meta_positions"))):
                    result = {"name": clean(entry.get("meta_name")) or name,
                              "imo": clean(entry.get("meta_imo")),
                              "positions": num(entry.get("meta_positions"), 0)}
                    self._skeleton_metadata_cache[fingerprint] = result
                    return dict(result)
            except Exception:
                pass
        try:
            wb = load_workbook(p, read_only=True, data_only=True)
            if "VESSEL" in wb.sheetnames:
                for row in wb["VESSEL"].iter_rows(values_only=True):
                    if not row or row[0] is None or len(row) < 2: continue
                    key = clean(row[0]).lower()
                    value = clean(row[1])
                    if key in ("vessel name", "vessel_name", "название судна", "судно") and value and not value.upper().startswith("DEMO"):
                        name = value
                    elif key in ("imo number", "imo", "imo номер") and value and value.upper() not in ("НЕ ВВЕДЁН", "НЕ ВВЕДЕН", "UNKNOWN"):
                        imo = value
            if "POSITIONS" in wb.sheetnames:
                ws = wb["POSITIONS"]
                header_row = None
                header_vals = None
                for rr, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 8), values_only=True), 1):
                    vals = [clean(v).lower() for v in row]
                    if "bay" in vals and "row" in vals and "tier" in vals:
                        header_row, header_vals = rr, vals
                        break
                if header_row is not None:
                    header = {v: i for i, v in enumerate(header_vals) if v}
                    c_id = header.get("position id")
                    c_bay, c_row, c_tier = header.get("bay"), header.get("row"), header.get("tier")
                    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
                        def val(idx): return clean(row[idx]) if idx is not None and idx < len(row) else ""
                        pid, bv, rv, tv = val(c_id), val(c_bay), val(c_row), val(c_tier)
                        if pid.upper().startswith("DEMO") or "DEMO-REMOVE" in pid.upper(): continue
                        if not (bv and rv and tv): continue
                        try:
                            int(float(bv)); int(float(rv)); int(float(tv)); positions += 1
                        except Exception:
                            pass
            wb.close()
        except Exception:
            pass
        result = {"name": name, "imo": imo, "positions": positions}
        self._skeleton_metadata_cache[fingerprint] = dict(result)
        return result

    def _skeleton_entries(self):
        entries = self.config_data.get("skeletons", [])
        if not isinstance(entries, list): entries = []
        valid = []
        for entry in entries:
            if isinstance(entry, dict) and clean(entry.get("path")) and Path(clean(entry.get("path"))).exists():
                valid.append(entry)
        self.config_data["skeletons"] = valid
        return valid

    def refresh_skeleton_registry(self):
        if not hasattr(self, "skeleton_tree"): return
        for item in self.skeleton_tree.get_children(): self.skeleton_tree.delete(item)
        active = Path(clean(self.config_data.get("skeleton_path", ""))).resolve() if clean(self.config_data.get("skeleton_path", "")) else None
        entries = self._skeleton_entries()
        no_skeleton_active = active is None
        self.skeleton_tree.insert("", "end", iid="__no_skeleton__", values=("Без скелета" + (" ✓" if no_skeleton_active else ""), "—", "—", "Общий план", "(без файла скелета)"))
        for i, entry in enumerate(entries):
            path = Path(clean(entry.get("path", "")))
            meta = self._skeleton_metadata(path)
            try:
                st = path.stat()
                entry.update({"meta_name": meta["name"], "meta_imo": meta["imo"],
                              "meta_positions": meta["positions"], "meta_mtime_ns": st.st_mtime_ns,
                              "meta_size": st.st_size})
            except Exception:
                pass
            mark = " ✓" if active and path.resolve() == active else ""
            self.skeleton_tree.insert("", "end", iid=str(i), values=(meta["name"] + mark, meta["imo"], meta["positions"], path.name, str(path)))
        config_save(self.config_data)
        if hasattr(self, "skeleton_registry_status"):
            self.skeleton_registry_status.set(f"Скелетов в перечне: {len(entries)}. Активное судно отмечено галочкой. Рабочая папка: {SKELETONS_DIR}")

    def _selected_skeleton_entry(self):
        if not hasattr(self, "skeleton_tree"): return None
        selection = self.skeleton_tree.selection()
        if not selection: return None
        if selection[0] == "__no_skeleton__":
            return {"name": "Без скелета", "imo": "", "positions": 0, "path": ""}
        try: return self._skeleton_entries()[int(selection[0])]
        except Exception: return None

    def activate_selected_skeleton(self):
        entry = self._selected_skeleton_entry()
        if not entry:
            messagebox.showinfo("Судна", "Сначала выберите скелет в перечне.", parent=self); return
        raw_path = clean(entry.get("path", ""))
        if not raw_path:
            self.config_data["skeleton_path"] = ""
            config_save(self.config_data)
            if hasattr(self, "skeleton_label"): self.skeleton_label.set("Скелет не выбран — общий план по данным контейнеров")
            self.skeleton_only_mode.set(False)
            self.refresh_plan(); self.refresh_skeleton_registry()
            self.set_status("Режим без скелета: отображается общий план контейнеров.")
            return
        path = Path(raw_path)
        self.config_data["skeleton_path"] = str(path.resolve())
        config_save(self.config_data)
        meta = self._skeleton_metadata(path)
        if hasattr(self, "skeleton_label"): self.skeleton_label.set(f"{meta['name']} — {path.name}")
        self.skeleton_only_mode.set(False)
        self.refresh_plan()
        self.refresh_skeleton_registry()
        self.set_status(f"Активное судно: {meta['name']} ({path.name}). План обновлён.")

    def remove_selected_skeleton(self):
        entry = self._selected_skeleton_entry()
        if not entry:
            messagebox.showinfo("Судна", "Сначала выберите скелет в перечне.", parent=self); return
        raw_path = clean(entry.get("path", ""))
        if not raw_path:
            messagebox.showinfo("Без скелета", "Это системный пункт для общего плана, его нельзя удалить.", parent=self); return
        path = Path(raw_path)
        meta = self._skeleton_metadata(path)
        if not messagebox.askyesno("Удалить скелет", f"Убрать «{meta['name']}» из перечня?\nФайл в рабочей папке приложения также будет удалён, если он хранится в библиотеке скелетов.", parent=self): return
        self.config_data["skeletons"] = [e for e in self._skeleton_entries() if Path(clean(e.get("path", ""))).resolve() != path.resolve()]
        if clean(self.config_data.get("skeleton_path", "")) and Path(clean(self.config_data.get("skeleton_path"))).resolve() == path.resolve():
            self.config_data["skeleton_path"] = ""
            self.skeleton_only_mode.set(False)
        try:
            if path.resolve().parent == SKELETONS_DIR.resolve() and path.exists(): path.unlink()
        except Exception as e:
            messagebox.showwarning("Удаление файла", f"Запись удалена из перечня, но файл не удалось удалить:\n{e}", parent=self)
        config_save(self.config_data)
        self.refresh_skeleton_registry(); self.refresh_plan()
        self.set_status(f"Скелет «{meta['name']}» удалён из перечня.")

    def choose_skeleton_for_plan(self):
        entries = self._skeleton_entries()
        if not entries:
            if messagebox.askyesno("Нет загруженных судов", "В библиотеке пока нет скелетов. Загрузить Excel-скелет сейчас?", parent=self): self.choose_skeleton()
            return
        win = tk.Toplevel(self); win.title("Выбрать судно для плана"); win.transient(self); win.grab_set(); win.geometry("720x390")
        ttk.Label(win, text="Выберите скелет, который нужно показать на плане судна:", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=(14, 8))
        tree = ttk.Treeview(win, columns=("vessel", "imo", "positions", "file"), show="headings", selectmode="browse", height=10)
        for col, title, width in [("vessel", "Судно", 190), ("imo", "IMO", 100), ("positions", "Позиций", 80), ("file", "Файл скелета", 260)]:
            tree.heading(col, text=title); tree.column(col, width=width, anchor="w")
        tree.insert("", "end", iid="__no_skeleton__", values=("Без скелета (общий план)", "—", "—", "—"))
        for i, entry in enumerate(entries):
            path = Path(clean(entry.get("path", ""))); meta = self._skeleton_metadata(path)
            tree.insert("", "end", iid=str(i), values=(meta["name"], meta["imo"], meta["positions"], path.name))
        tree.pack(fill="both", expand=True, padx=14, pady=4)
        active_path = clean(self.config_data.get("skeleton_path", ""))
        if not active_path:
            tree.selection_set("__no_skeleton__"); tree.focus("__no_skeleton__")
        for i, entry in enumerate(entries):
            if active_path and Path(clean(entry.get("path", ""))).resolve() == Path(active_path).resolve():
                tree.selection_set(str(i)); tree.focus(str(i)); break
        def show(empty):
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Выбор судна", "Выберите судно из списка.", parent=win); return
            if sel[0] == "__no_skeleton__":
                self.config_data["skeleton_path"] = ""; config_save(self.config_data)
                if hasattr(self, "skeleton_label"): self.skeleton_label.set("Скелет не выбран — общий план по данным контейнеров")
                self.skeleton_only_mode.set(False); self.refresh_plan(); self.refresh_skeleton_registry()
                self.set_status("Показан общий план без скелета судна."); win.destroy(); return
            entry = entries[int(sel[0])]; path = Path(clean(entry.get("path", "")))
            self.config_data["skeleton_path"] = str(path.resolve()); config_save(self.config_data)
            meta = self._skeleton_metadata(path)
            if hasattr(self, "skeleton_label"): self.skeleton_label.set(f"{meta['name']} — {path.name}")
            if not self.extract_skeleton_positions() and not (self.extract_skeleton_bays() and self.extract_skeleton_rows() and self.extract_skeleton_tiers()):
                messagebox.showwarning("Нет координат скелета", "В выбранном файле не найдены координаты Bay / Row / Tier на листе POSITIONS или в старой графической схеме.", parent=win)
                return
            self.skeleton_only_mode.set(bool(empty)); self.refresh_plan(); self.refresh_skeleton_registry()
            self.set_status(f"На плане выбрано судно: {meta['name']}. Режим: {'пустой скелет' if empty else 'контейнеры поверх скелета'}.")
            win.destroy()
        buttons = ttk.Frame(win); buttons.pack(fill="x", padx=14, pady=(8, 14))
        ttk.Button(buttons, text="Показать пустой скелет", command=lambda: show(True)).pack(side="left")
        ttk.Button(buttons, text="Показать с контейнерами", command=lambda: show(False)).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="Отмена", command=win.destroy).pack(side="right")
        tree.bind("<Double-1>", lambda _e: show(True))

    def toggle_skeleton_only(self):
        """Toggle between the empty vessel skeleton and the same plan with EDI/container records overlaid."""
        if not clean(self.config_data.get("skeleton_path", "")) or not Path(clean(self.config_data.get("skeleton_path", ""))).exists():
            messagebox.showwarning("Скелет судна не выбран", "Сначала нажмите «Загрузить скелет судна…» и выберите Excel-файл скелета.", parent=self)
            return
        if not self.extract_skeleton_positions() and not (self.extract_skeleton_bays() and self.extract_skeleton_rows() and self.extract_skeleton_tiers()):
            messagebox.showwarning("Нет координат скелета", "В файле не найдены координаты Bay / Row / Tier. Для нового шаблона заполните лист POSITIONS реальными позициями из Manual.", parent=self)
            return
        self.skeleton_only_mode.set(not self.skeleton_only_mode.get())
        self.refresh_plan()
        if self.skeleton_only_mode.get():
            self.set_status("Показан пустой скелет судна; данные контейнеров сохранены и не удалены.")
        else:
            self.set_status("Показан план с контейнерами поверх скелета судна.")

    def extract_skeleton_positions(self):
        """Return cached structured positions; parse each unchanged Excel skeleton once."""
        path = clean(self.config_data.get("skeleton_path", ""))
        if not path or not Path(path).exists():
            return []
        fingerprint = self._skeleton_fingerprint(path)
        if fingerprint in self._skeleton_positions_cache:
            return [dict(pos) for pos in self._skeleton_positions_cache[fingerprint]]
        positions = []
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            if "POSITIONS" not in wb.sheetnames:
                wb.close()
                self._skeleton_positions_cache[fingerprint] = []
                return []
            ws = wb["POSITIONS"]
            header_row = None
            headers = {}
            for rr, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 8), values_only=True), 1):
                vals = [clean(v) for v in row]
                if "Bay" in vals and "Row" in vals and "Tier" in vals:
                    header_row = rr
                    headers = {v.strip().lower(): i for i, v in enumerate(vals) if v}
                    break
            if header_row:
                def col(*names):
                    for name in names:
                        if name.lower() in headers: return headers[name.lower()]
                    return None
                c_id=col("Position ID"); c_bay=col("Bay"); c_row=col("Row"); c_tier=col("Tier")
                c_area=col("Area (Hold/Deck)", "Area"); c_status=col("Position status", "Status")
                for row_data in ws.iter_rows(min_row=header_row+1, values_only=True):
                    def val(idx): return clean(row_data[idx]) if idx is not None and idx < len(row_data) else ""
                    pid=val(c_id)
                    if pid.upper().startswith("DEMO") or "DEMO-REMOVE" in pid.upper(): continue
                    bv,rv,tv=val(c_bay),val(c_row),val(c_tier)
                    if not (bv and rv and tv): continue
                    try: bay=int(float(bv)); row=int(float(rv)); tier=int(float(tv))
                    except Exception: continue
                    positions.append({"id":pid,"bay":bay,"row":row,"tier":tier,
                        "area":val(c_area),"status":val(c_status)})
            wb.close()
        except Exception:
            positions = []
        self._skeleton_positions_cache[fingerprint] = [dict(pos) for pos in positions]
        return [dict(pos) for pos in positions]

    def on_plan_mode_change(self, _event=None):
        # Clear the old selector value because Tier and Bay lists use different domains.
        self.tier_choice.set("")
        self.refresh_plan()

    def _build_editor(self):
        top=ttk.Frame(self.tab_editor); top.pack(fill="x", pady=(0,8))
        ttk.Label(top, text="Контейнеры", font=("Segoe UI",14,"bold"), foreground="#17365D").pack(side="left")
        ttk.Button(top, text="＋ Добавить", command=self.add_container).pack(side="right")
        ttk.Button(top, text="Данные судна", command=self.edit_vessel).pack(side="right", padx=(8,0))
        ttk.Button(top, text="Удалить", command=self.delete_container).pack(side="right", padx=(8,0))
        ttk.Button(top, text="Изменить", command=self.edit_container).pack(side="right", padx=(8,0))
        ttk.Button(top, text="Проверить план", command=self.check_plan).pack(side="right", padx=(8,0))
        ttk.Button(top, text="Сравнить с исходным", command=self.compare_with_source).pack(side="right", padx=(8,0))
        ttk.Button(top, text="Массовое изменение…", command=self.mass_edit_selected).pack(side="right", padx=(8,0))
        filterbar=ttk.Frame(self.tab_editor); filterbar.pack(fill="x", pady=(0,8))
        ttk.Label(filterbar,text="Фильтр по столбцу / координате:").pack(side="left")
        self.filter_combo=ttk.Combobox(filterbar,textvariable=self.filter_field,state="readonly",width=20,values=["Все поля"]+FIELDS)
        self.filter_combo.pack(side="left",padx=(8,6))
        filter_entry=ttk.Entry(filterbar,textvariable=self.filter_value,width=26); filter_entry.pack(side="left")
        filter_entry.bind("<Return>",lambda _e:self.apply_editor_filter())
        ttk.Button(filterbar,text="Применить",command=self.apply_editor_filter).pack(side="left",padx=(6,0))
        ttk.Button(filterbar,text="Сбросить",command=self.clear_editor_filter).pack(side="left",padx=(6,0))
        cols=FIELDS
        container=ttk.Frame(self.tab_editor); container.pack(fill="both",expand=True)
        self.editor_tree=ttk.Treeview(container,columns=cols,show="headings",selectmode="extended")
        widths={"CONTAINER_NO":135,"LINE":110,"ISO_TYPE":80,"GROSS_KG":90,"BAY":60,"ROW":60,"TIER":60,"POL":75,"POT":75,"POD":75,"F_E":55,"FINAL_DEST":95,"BOOKING_NO":115,"REEFER_TEMP_C":80,"UN_NUMBER":75,"IMDG_CLASS":75,"OVERDIMENSION_NOTE":160,"SEAL_NO":110,"COMMENTS":160}
        for f in cols:
            self.editor_tree.heading(f,text=f)
            self.editor_tree.column(f,width=widths.get(f,95),minwidth=55,stretch=False,anchor="center" if f in ("GROSS_KG","BAY","ROW","TIER","F_E") else "w")
        sy=ttk.Scrollbar(container,orient="vertical",command=self.editor_tree.yview); sx=ttk.Scrollbar(container,orient="horizontal",command=self.editor_tree.xview)
        self.editor_tree.configure(yscrollcommand=sy.set,xscrollcommand=sx.set)
        self.editor_tree.grid(row=0,column=0,sticky="nsew"); sy.grid(row=0,column=1,sticky="ns"); sx.grid(row=1,column=0,sticky="ew")
        container.rowconfigure(0,weight=1); container.columnconfigure(0,weight=1)
        self.editor_tree.bind("<Double-1>", lambda _e: self.edit_container())
        bottom=ttk.Frame(self.tab_editor); bottom.pack(fill="x",pady=(8,0))
        self.editor_count=tk.StringVar(value="Контейнеров: 0")
        ttk.Label(bottom,textvariable=self.editor_count).pack(side="left")
        ttk.Label(bottom,text="Ctrl/Shift — выделение нескольких строк. Двойной щелчок — карточка контейнера.",foreground="#526477").pack(side="right")

    def _build_import(self):
        ttk.Label(self.tab_import, text="Импорт и создание документа", font=("Segoe UI",15,"bold"), foreground="#17365D").pack(anchor="w")
        ttk.Label(self.tab_import, text="Загрузи EDI/Excel или начни с пустой таблицы. Данные затем доступны в редакторе.", wraplength=1100).pack(anchor="w",pady=(4,16))
        cards=ttk.Frame(self.tab_import); cards.pack(fill="x")
        self._action_card(cards,"Excel → Редактор","Открыть Excel с листами VESSEL и CONTAINERS или совместимую таблицу.","Загрузить Excel…",self.import_excel,0)
        self._action_card(cards,"EDI → Редактор","Разобрать EDI и открыть контейнеры для просмотра и изменения.","Загрузить EDI…",self.import_edi,1)
        self._action_card(cards,"Новый документ","Создать пустую таблицу контейнеров и заполнить её в программе.","Создать пустой документ",self.new_document,2)
        info=ttk.LabelFrame(self.tab_import,text="Текущий документ",padding=14); info.pack(fill="x",pady=16)
        ttk.Label(info,textvariable=self.current_display,wraplength=1000).pack(anchor="w")
        ttk.Button(info,text="Перейти к редактору",command=lambda:self.nb.select(self.tab_editor)).pack(anchor="w",pady=(10,0))

    def _action_card(self,parent,title,desc,button,callback,column):
        card=ttk.LabelFrame(parent,text=title,padding=14); card.grid(row=0,column=column,sticky="nsew",padx=(0 if column==0 else 8,0))
        parent.columnconfigure(column,weight=1)
        ttk.Label(card,text=desc,wraplength=310,justify="left").pack(anchor="w",pady=(0,14))
        ttk.Button(card,text=button,command=callback).pack(anchor="w")

    def _build_data(self):
        ttk.Label(self.tab_data, text="Справочники приложения", font=("Segoe UI", 15, "bold"), foreground="#17365D").pack(anchor="w")
        ttk.Label(self.tab_data, text="Данные сохраняются локально и используются при отображении краткого Stowage Plan. Списки можно пополнять по мере необходимости.", wraplength=1100).pack(anchor="w", pady=(3,10))
        toolbar=ttk.Frame(self.tab_data); toolbar.pack(fill="x", pady=(0,8))
        ttk.Label(toolbar,text="Раздел:").pack(side="left")
        self.reference_category_combo=ttk.Combobox(toolbar,textvariable=self.reference_category,state="readonly",width=22,values=list(DEFAULT_REFERENCE_DATA.keys()))
        self.reference_category_combo.pack(side="left",padx=(6,12)); self.reference_category_combo.bind("<<ComboboxSelected>>",lambda _e:self.refresh_reference_table())
        ttk.Button(toolbar,text="＋ Добавить",command=self.add_reference_item).pack(side="left")
        ttk.Button(toolbar,text="Изменить",command=self.edit_reference_item).pack(side="left",padx=(6,0))
        ttk.Button(toolbar,text="Удалить",command=self.delete_reference_item).pack(side="left",padx=(6,0))
        ttk.Button(toolbar,text="Вернуть встроенный справочник типов",command=self.reset_type_reference).pack(side="right")
        frame=ttk.Frame(self.tab_data); frame.pack(fill="both",expand=True)
        cols=("key","short","description","group")
        self.reference_tree=ttk.Treeview(frame,columns=cols,show="headings",selectmode="browse")
        if self.reference_category.get()=="Порты":
            self.reference_tree.configure(displaycolumns=("group","description","key","short"))
            headings=[("key","Код обозначения 5 символов",210),("short","Код обозначения 2 символа",190),("description","Порт",260),("group","Страна",180)]
        else:
            self.reference_tree.configure(displaycolumns=("key","short","description","group"))
            headings=[("key","Код / ISO / значение",180),("short","Краткое обозначение",170),("description","Описание / название",430),("group","Группа",180)]
        for col,title,width in headings:
            self.reference_tree.heading(col,text=title); self.reference_tree.column(col,width=width,anchor="w")
        sy=ttk.Scrollbar(frame,orient="vertical",command=self.reference_tree.yview); self.reference_tree.configure(yscrollcommand=sy.set)
        self.reference_tree.grid(row=0,column=0,sticky="nsew"); sy.grid(row=0,column=1,sticky="ns")
        frame.rowconfigure(0,weight=1); frame.columnconfigure(0,weight=1)
        self.reference_tree.bind("<Double-1>",lambda _e:self.edit_reference_item())
        self.reference_hint=tk.StringVar(value="")
        ttk.Label(self.tab_data,textvariable=self.reference_hint,foreground="#526477",wraplength=1100).pack(anchor="w",pady=(8,0))
        self.refresh_reference_table()

    def _reference_items(self):
        data=self.config_data.setdefault("reference_data",{})
        category=self.reference_category.get()
        if not isinstance(data.get(category),list): data[category]=[]
        return data[category]

    def refresh_reference_table(self):
        if not hasattr(self,"reference_tree"): return
        for iid in self.reference_tree.get_children(): self.reference_tree.delete(iid)
        for i,item in enumerate(self._reference_items()):
            self.reference_tree.insert("","end",iid=str(i),values=(clean(item.get("key")),clean(item.get("short")),clean(item.get("description")),clean(item.get("group"))))
        hints={"Порты":"Вноси вручную: страна, название порта, код из 5 символов и соответствующий код из 2 символов. Неизвестные коды не преобразуются автоматически.","Типы ISO":"ISO-код → короткое обозначение для Stowage Plan. Подтверждённые пользователем правила включены; справочник можно редактировать.","Другие значения":"Универсальный список для дополнительных кодов, статусов и других значений."}
        self.reference_hint.set(hints.get(self.reference_category.get(),""))

    def _reference_item_dialog(self,title,initial=None):
        initial=initial or {}
        win=tk.Toplevel(self); win.title(title); win.transient(self); win.grab_set(); win.resizable(False,False)
        body=ttk.Frame(win,padding=14); body.pack(fill="both",expand=True)
        if self.reference_category.get()=="Порты":
            fields=[("group","Страна"),("description","Порт"),("key","Код обозначения 5 символов"),("short","Код обозначения 2 символа")]
        else:
            fields=[("key","Код / ISO / значение"),("short","Краткое обозначение"),("description","Описание / название"),("group","Группа")]
        vars_={}
        for row,(key,label) in enumerate(fields):
            ttk.Label(body,text=label).grid(row=row,column=0,sticky="w",padx=(0,10),pady=5)
            v=tk.StringVar(value=clean(initial.get(key)))
            ttk.Entry(body,textvariable=v,width=44).grid(row=row,column=1,sticky="ew",pady=5)
            vars_[key]=v
        result={"value":None}
        def save():
            item={k:v.get().strip() for k,v in vars_.items()}
            if not item["key"]:
                messagebox.showwarning("Справочник","Поле кода обязательно.",parent=win); return
            if self.reference_category.get()=="Порты":
                if len(item["key"].strip())!=5 or len(item["short"].strip())!=2 or not item["description"].strip() or not item["group"].strip():
                    messagebox.showwarning("Справочник портов","Заполни страну, порт, 5-символьный код и 2-символьный код.",parent=win); return
            result["value"]=item; win.destroy()
        buttons=ttk.Frame(body); buttons.grid(row=len(fields),column=0,columnspan=2,sticky="e",pady=(10,0))
        ttk.Button(buttons,text="Отмена",command=win.destroy).pack(side="right")
        ttk.Button(buttons,text="Сохранить",command=save).pack(side="right",padx=(0,8))
        body.columnconfigure(1,weight=1); win.wait_window()
        return result["value"]

    def _save_reference_data(self):
        self.config_data["reference_data"] = self.config_data.get("reference_data",{})
        config_save(self.config_data)
        self.refresh_reference_table()
        if hasattr(self,"plan_canvas"): self.refresh_plan()
        self.set_status("Справочник сохранён локально.")

    def add_reference_item(self):
        item=self._reference_item_dialog("Добавить значение в справочник")
        if not item:return
        items=self._reference_items()
        if any(clean(x.get("key")).casefold()==item["key"].casefold() for x in items):
            messagebox.showwarning("Справочник","Такой код уже существует в выбранном разделе.",parent=self); return
        items.append(item); items.sort(key=lambda x:clean(x.get("key")).casefold()); self._save_reference_data()

    def edit_reference_item(self):
        sel=self.reference_tree.selection()
        if not sel:return
        idx=int(sel[0]); items=self._reference_items()
        if idx<0 or idx>=len(items):return
        item=self._reference_item_dialog("Изменить значение справочника",items[idx])
        if not item:return
        if any(i!=idx and clean(x.get("key")).casefold()==item["key"].casefold() for i,x in enumerate(items)):
            messagebox.showwarning("Справочник","Такой код уже существует в выбранном разделе.",parent=self); return
        items[idx]=item; items.sort(key=lambda x:clean(x.get("key")).casefold()); self._save_reference_data()

    def delete_reference_item(self):
        sel=self.reference_tree.selection()
        if not sel:return
        idx=int(sel[0]); items=self._reference_items()
        if idx<0 or idx>=len(items):return
        if not messagebox.askyesno("Удаление значения",f"Удалить «{clean(items[idx].get('key'))}» из раздела «{self.reference_category.get()}»?",parent=self):return
        del items[idx]; self._save_reference_data()

    def reset_type_reference(self):
        if not messagebox.askyesno("Встроенный справочник","Заменить текущий список типов ISO встроенным стартовым списком? Пользовательские изменения в разделе «Типы ISO» будут потеряны.",parent=self):return
        self.config_data.setdefault("reference_data",{})["Типы ISO"]=json.loads(json.dumps(DEFAULT_REFERENCE_DATA["Типы ISO"],ensure_ascii=False))
        self._save_reference_data()

    def _build_stats(self):
        ttk.Label(self.tab_stats,text="Статистика загрузки",font=("Segoe UI",15,"bold"),foreground="#17365D").pack(anchor="w")
        ttk.Label(self.tab_stats,text="Отчёты строятся по текущему содержимому редактора. TEU рассчитывается по принятому правилу: 20 ft = 1 TEU, 40/45 ft = 2 TEU.",wraplength=1150).pack(anchor="w",pady=(4,10))
        controls=ttk.Frame(self.tab_stats); controls.pack(fill="x",pady=(0,8))
        ttk.Label(controls,text="Отчёт:").pack(side="left")
        combo=ttk.Combobox(controls,textvariable=self.stats_report_type,state="readonly",width=24,values=TEU_FIELD_OPTIONS)
        combo.pack(side="left",padx=(8,8)); combo.bind("<<ComboboxSelected>>",lambda _e:self.refresh_stats())
        ttk.Button(controls,text="Обновить",command=self.refresh_stats).pack(side="left")
        ttk.Button(controls,text="Открыть папку отчётов",command=self.open_exports_folder).pack(side="right",padx=(0,8))
        ttk.Button(controls,text="Экспортировать все отчёты в Excel…",style="Accent.TButton",command=self.export_statistics).pack(side="right")
        summary=ttk.LabelFrame(self.tab_stats,text="Общая сводка",padding=10); summary.pack(fill="x",pady=(0,10))
        ttk.Label(summary,textvariable=self.stats_summary,font=("Segoe UI",11,"bold"),foreground="#17365D").pack(anchor="w")
        frame=ttk.Frame(self.tab_stats); frame.pack(fill="both",expand=True)
        cols=("group","count","teu","kg","twenty","forty")
        self.stats_tree=ttk.Treeview(frame,columns=cols,show="headings")
        for col,title,width,anchor in [("group","Группа / показатель",270,"w"),("count","Контейнеров",110,"e"),("teu","TEU",90,"e"),("kg","Вес, кг",130,"e"),("twenty","20 ft",90,"e"),("forty","40/45 ft",100,"e")]:
            self.stats_tree.heading(col,text=title); self.stats_tree.column(col,width=width,anchor=anchor,stretch=col=="group")
        sy=ttk.Scrollbar(frame,orient="vertical",command=self.stats_tree.yview); self.stats_tree.configure(yscrollcommand=sy.set)
        self.stats_tree.grid(row=0,column=0,sticky="nsew"); sy.grid(row=0,column=1,sticky="ns")
        frame.rowconfigure(0,weight=1); frame.columnconfigure(0,weight=1)

    def _build_export(self):
        ttk.Label(self.tab_export,text="Сохранение и экспорт файлов",font=("Segoe UI",15,"bold"),foreground="#17365D").pack(anchor="w")
        ttk.Label(self.tab_export,text="Экспорт создаётся из текущего состояния редактора, включая внесённые изменения.",wraplength=1100).pack(anchor="w",pady=(4,8))
        plan_box=ttk.LabelFrame(self.tab_export,text="Сохранение графического плана судна",padding=14); plan_box.pack(fill="x",pady=(0,12))
        ttk.Label(plan_box,text="Выбери режим Tier или Bay, подробный Full Plan или компактный Stowage Plan, а также область сохранения.",wraplength=1100).pack(anchor="w")
        ttk.Button(plan_box,text="Настроить и сохранить план…",style="Accent.TButton",command=self.export_plan_jpg).pack(anchor="e",pady=(8,0))
        source=ttk.LabelFrame(self.tab_export,text="Текущий источник экспорта",padding=10); source.pack(fill="x",pady=(0,12))
        ttk.Label(source,textvariable=self.export_source_display,wraplength=1100).pack(anchor="w")
        edi=ttk.LabelFrame(self.tab_export,text="EDI",padding=14); edi.pack(fill="x",pady=(0,10))
        ttk.Label(edi,text="Сформировать EDI-файл из текущего плана контейнеров.").pack(side="left",fill="x",expand=True)
        ttk.Button(edi,text="Экспорт EDI…",style="Accent.TButton",command=self.export_edi).pack(side="right")
        excel=ttk.LabelFrame(self.tab_export,text="Графический Excel",padding=14); excel.pack(fill="x",pady=(0,10))
        ttk.Label(excel,text="Сохраняет SUMMARY, CONTAINERS, RAW EDI и расположение контейнеров по выбранному режиму.").pack(anchor="w")
        modes=ttk.Frame(excel); modes.pack(fill="x",pady=(8,0))
        for m in ("Tier","Bay","Complex"):
            ttk.Radiobutton(modes,text=m,variable=self.mode,value=m).pack(side="left",padx=(0,18))
        ttk.Button(modes,text="Экспорт Excel…",style="Accent.TButton",command=self.export_graph_excel).pack(side="right")
        template=ttk.LabelFrame(self.tab_export,text="Входной шаблон Excel для создания EDI",padding=14); template.pack(fill="x",pady=(0,10))
        ttk.Label(template,text="Выгружает незаполненный шаблон того же образца: структура VESSEL / CONTAINERS и листы с объяснениями сохраняются, записи контейнеров очищаются.",wraplength=1000).pack(anchor="w")
        row=ttk.Frame(template); row.pack(fill="x",pady=(10,0))
        ttk.Button(row,text="Скачать пустой шаблон Excel…",command=self.export_blank_template).pack(side="left")
        ttk.Button(row,text="Выбрать/запомнить образец…",command=self.choose_template).pack(side="left",padx=(8,0))
        project=ttk.LabelFrame(self.tab_export,text="Рабочий файл",padding=14); project.pack(fill="x")
        ttk.Label(project,text="Сохранить редактируемый документ в обычном .edi или .xlsx, чтобы его можно было открыть через Проводник Windows.").pack(anchor="w")
        row2=ttk.Frame(project); row2.pack(fill="x",pady=(10,0))
        ttk.Button(row2,text="Сохранить",command=self.save_current).pack(side="left")
        ttk.Button(row2,text="Сохранить как…",command=self.save_as).pack(side="left",padx=(8,0))
        ttk.Button(row2,text="Открыть папку проектов",command=self.open_projects_folder).pack(side="left",padx=(8,0))

    def set_status(self,text):
        self.status.set(text)

    def update_current_display(self):
        suffix = "  •  есть несохранённые изменения" if self.dirty else ""
        if self.current_path:
            self.current_display.set(f"Текущий документ: {self.current_path.name}{suffix}")
            source_name = self.current_path.name
        else:
            self.current_display.set(f"Новый документ — ещё не сохранён{suffix}")
            source_name = "Новый документ (ещё не сохранён)"
        if hasattr(self, "export_source_display"):
            self.export_source_display.set(f"Источник: {source_name}  |  Контейнеров в редакторе: {len(self.records)}{suffix}. При экспорте используются текущие данные редактора, даже если рабочий файл ещё не сохранён.")

    def refresh_projects(self):
        if not hasattr(self,"project_tree"): return
        for item in self.project_tree.get_children(): self.project_tree.delete(item)
        ensure_folders()
        files=sorted([p for p in PROJECTS_DIR.iterdir() if p.is_file() and p.suffix.lower() in (".edi",".txt",".baplie",".xlsx",".xlsm")], key=lambda p:p.stat().st_mtime, reverse=True)
        for p in files:
            try:
                st=p.stat(); modified=datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M"); size=f"{st.st_size/1024:.1f} KB"
            except Exception: modified="—"; size="—"
            self.project_tree.insert("","end",iid=str(p),values=(p.name,p.suffix.lower().lstrip("."),modified,size))
        self.set_status(f"Папка проектов обновлена: {len(files)} файл(ов)")

    def get_selected_project(self):
        if not hasattr(self,"project_tree"): return None
        sel=self.project_tree.selection()
        return Path(sel[0]) if sel else None

    def open_selected_project(self):
        p=self.get_selected_project()
        if not p:
            messagebox.showinfo("Проекты","Сначала выбери файл в списке.",parent=self); return
        self.load_document(p)

    def delete_selected_project(self):
        p=self.get_selected_project()
        if not p or not p.exists():
            messagebox.showinfo("Проекты","Сначала выбери файл в списке.",parent=self); return
        if not messagebox.askyesno("Подтверждение удаления",f"Удалить файл «{p.name}» из папки проектов?\n\nЕсли доступна Корзина Windows, файл будет перемещён туда.",parent=self): return
        try:
            try:
                from send2trash import send2trash
                send2trash(str(p))
            except ImportError:
                p.unlink()
            if self.current_path and self.current_path.resolve()==p.resolve():
                self.current_path=None; self.records=[]; self.vessel=dict(DEFAULT_VESSEL); self.raw_edi=""; self.dirty=False; self.refresh_editor(); self.refresh_plan(); self.update_current_display()
            self.refresh_projects(); self.set_status(f"Удалён файл: {p.name}")
        except Exception as e: messagebox.showerror("Ошибка удаления",str(e),parent=self)

    def open_projects_folder(self):
        ensure_folders()
        try:
            if os.name=="nt": os.startfile(str(PROJECTS_DIR))
            elif sys.platform=="darwin": os.system(f'open "{PROJECTS_DIR}"')
            else: os.system(f'xdg-open "{PROJECTS_DIR}"')
        except Exception as e: messagebox.showerror("Папка проектов",f"Путь:\n{PROJECTS_DIR}\n\n{e}",parent=self)

    def open_dialog(self):
        p=filedialog.askopenfilename(title="Открыть документ",filetypes=SUPPORTED_OPEN,initialdir=str(PROJECTS_DIR))
        if p: self.load_document(Path(p))

    def load_document(self,path):
        self._last_load_succeeded=False
        if self.dirty and self.active_project_id:
            # Active projects have their own autosaved snapshot, so switching files does not discard work.
            self._write_active_project_snapshot()
            self.dirty=False
        elif self.dirty:
            if not messagebox.askyesno("Несохранённые изменения", "В текущем документе есть несохранённые изменения. Открыть другой документ и отказаться от них?", parent=self):
                return
        path=Path(path)
        if not path.exists():
            messagebox.showerror("Не найден файл",str(path),parent=self); return
        try:
            if path.suffix.lower() in (".edi",".txt",".baplie"):
                raw,vessel,records=decode_edi_document(path)
            elif path.suffix.lower() in (".xlsx",".xlsm"):
                vessel,records=read_excel_document(path)
                raw=""
            else:
                raise ValueError("Поддерживаются EDI (.edi/.txt/.baplie) и Excel (.xlsx/.xlsm). Формат .xls не поддерживается библиотекой чтения Excel.")
            # Imported files outside the work folder are copied into Projects so edits are managed by the app
            # and the original source stays untouched. Existing files inside Projects are opened in place.
            work_path = path
            try:
                projects_resolved = PROJECTS_DIR.resolve()
                if path.resolve().parent != projects_resolved:
                    target = PROJECTS_DIR / path.name
                    suffix = 2
                    while target.exists():
                        target = PROJECTS_DIR / f"{path.stem} ({suffix}){path.suffix}"
                        suffix += 1
                    shutil.copy2(path, target)
                    work_path = target
            except Exception:
                # If copying fails, keep the source open and show its actual location.
                work_path = path
            self.vessel=dict(DEFAULT_VESSEL); self.vessel.update(vessel)
            self.records=[normalize_record(r) for r in records]
            self.source_records=[normalize_record(r) for r in records]
            self.source_vessel=dict(self.vessel)
            self.source_label=path.name
            self._search_match_indices=set(); self.move_mode_active=False; self.move_record_index=None
            self.raw_edi=raw; self.current_path=work_path; self.dirty=False
            if not self._creating_project and not self._loading_active_project:
                self.active_project_id=""; self.active_project_name=""
            self.refresh_editor(); self.refresh_plan(); self.refresh_stats(); self.update_current_display(); self.refresh_projects(); self.refresh_active_projects()
            self.set_status(f"Открыт {path.name}: контейнеров {len(records)}. Рабочий файл: {work_path}")
            self._last_load_succeeded=True
            self.nb.select(self.tab_editor)
        except Exception as e:
            messagebox.showerror("Ошибка открытия",f"Не удалось открыть файл:\n{path}\n\n{e}",parent=self)

    def new_document(self):
        if self.active_project_id:
            self._write_active_project_snapshot()
            self.dirty=False
        elif self.dirty:
            if not messagebox.askyesno("Несохранённые изменения","Текущие изменения будут потеряны. Создать новый документ?",parent=self): return
        self.vessel=dict(DEFAULT_VESSEL); self.records=[]; self.source_records=[]; self.source_vessel=dict(DEFAULT_VESSEL); self.source_label="Пустой документ"
        self._search_match_indices=set(); self.move_mode_active=False; self.move_record_index=None
        self.raw_edi=""; self.current_path=None; self.dirty=False
        self.active_project_id=""; self.active_project_name=""
        self.refresh_editor(); self.refresh_plan(); self.refresh_stats(); self.update_current_display(); self.set_status("Нулевой проект: можно работать без создания активного проекта; сохранение доступно через «Сохранить как…».")
        self.nb.select(self.tab_editor)

    def import_excel(self):
        template=find_template_path(self.config_data)
        p=filedialog.askopenfilename(title="Импортировать Excel",initialdir=str(Path(template).parent) if template else str(PROJECTS_DIR),filetypes=[("Excel","*.xlsx *.xlsm")])
        if p: self.load_document(Path(p))

    def import_edi(self):
        p=filedialog.askopenfilename(title="Импортировать EDI",initialdir=str(PROJECTS_DIR),filetypes=[("EDI","*.edi *.txt *.baplie"),("Все файлы","*.*")])
        if p: self.load_document(Path(p))

    def choose_template(self):
        p=filedialog.askopenfilename(title="Выбрать образец Excel для импорта/шаблона",filetypes=[("Excel","*.xlsx *.xlsm")])
        if not p:return
        try:
            wb=load_workbook(p,read_only=True,data_only=True)
            ok="VESSEL" in wb.sheetnames and "CONTAINERS" in wb.sheetnames
            wb.close()
            if not ok:
                messagebox.showwarning("Неподходящий образец","В книге должны быть листы VESSEL и CONTAINERS.",parent=self); return
            self.config_data["template_path"]=str(Path(p).resolve()); config_save(self.config_data)
            if hasattr(self,"template_label"): self.template_label.set(str(p))
            self.set_status(f"Запомнен Excel-образец: {Path(p).name}")
        except Exception as e: messagebox.showerror("Ошибка образца",str(e),parent=self)

    def choose_skeleton(self):
        p = filedialog.askopenfilename(title="Добавить скелет судна в библиотеку (Excel)", filetypes=[("Ship Creator CV / Excel skeleton", "*.scv *.xlsx *.xlsm"), ("All files", "*.*")])
        if not p: return
        try:
            source = Path(p).resolve()
            if source.suffix.lower()==".scv":
                with open(source,"r",encoding="utf-8") as f: scv=json.load(f)
                if scv.get("format")!="ShipCreatorCV" or not isinstance(scv.get("bays"),list):
                    raise ValueError("Файл не является скелетом Ship Creator CV (.scv).")
                SKELETONS_DIR.mkdir(parents=True,exist_ok=True)
                target= SKELETONS_DIR / source.name
                if source != target.resolve():
                    if target.exists():
                        stem,suffix=source.stem,source.suffix;n=2
                        while target.exists(): target=SKELETONS_DIR/f"{stem}_{n}{suffix}";n+=1
                    shutil.copy2(source,target)
                # Generate a CAD-readable workbook alongside the original .scv, retaining the .scv in the vessel library.
                converted=target.with_suffix(".xlsx")
                wb_out=Workbook(); meta=wb_out.active;meta.title="VESSEL"
                meta.append(["Vessel name",scv.get("vessel_name",source.stem)]);meta.append(["IMO number",scv.get("identifier","")])
                pos=wb_out.create_sheet("POSITIONS");pos.append(["Position ID","Bay","Row","Tier","Area (Hold/Deck)","Position status"])
                for bay in scv.get("bays",[]):
                    for area in ("hold","deck"):
                        for rr,tt in bay.get(area+"_cells",[]):
                            pos.append([f"{int(bay['number']):02d}{int(rr):02d}{int(tt):02d}",int(bay["number"]),int(rr),int(tt),"Hold" if area=="hold" else "Deck","Allowed"])
                wb_out.save(converted)
                source=converted
            wb = load_workbook(source, read_only=True, data_only=True)
            sheets = wb.sheetnames
            has_positions = "POSITIONS" in sheets
            wb.close()
            if not has_positions:
                if not messagebox.askyesno("Нет листа POSITIONS", "В файле нет листа POSITIONS. Добавить его в библиотеку всё равно? Старые графические форматы могут поддерживаться частично.", parent=self): return
            SKELETONS_DIR.mkdir(parents=True, exist_ok=True)
            # Copy imported skeleton into the app-owned library so the registry remains valid if the original is moved.
            target = SKELETONS_DIR / source.name
            if source != target.resolve():
                if target.exists() and source != target.resolve():
                    stem, suffix = source.stem, source.suffix
                    n = 2
                    while target.exists():
                        target = SKELETONS_DIR / f"{stem}_{n}{suffix}"; n += 1
                shutil.copy2(source, target)
            meta = self._skeleton_metadata(target)
            entries = self._skeleton_entries()
            resolved = target.resolve()
            if not any(Path(clean(e.get("path", ""))).resolve() == resolved for e in entries):
                entries.append({"name": meta["name"], "imo": meta["imo"], "path": str(resolved), "added_at": datetime.now().isoformat(timespec="seconds")})
            self.config_data["skeletons"] = entries
            self.config_data["skeleton_path"] = str(resolved)
            config_save(self.config_data)
            if hasattr(self, "skeleton_label"): self.skeleton_label.set(f"{meta['name']} — {target.name} — листы: {', '.join(sheets[:4])}")
            self.refresh_skeleton_registry()
            self.set_status(f"Скелет «{meta['name']}» добавлен в библиотеку: {target.name}")
            self.nb.select(self.tab_vessels)
        except Exception as e:
            messagebox.showerror("Ошибка скелета", str(e), parent=self)

    def extract_skeleton_bays(self):
        structured = self.extract_skeleton_positions()
        if structured:
            odd=set()
            for pos in structured:
                b=pos["bay"]
                if b>0:
                    if b%2==0: odd.update([b-1,b+1])
                    else: odd.add(b)
            return sorted([b for b in odd if b>0 and b%2==1], reverse=True)
        path=clean(self.config_data.get("skeleton_path",""))
        if not path or not Path(path).exists(): return []
        fingerprint = self._skeleton_fingerprint(path)
        cache_key = ("bays", fingerprint)
        if cache_key in self._skeleton_legacy_cache:
            return list(self._skeleton_legacy_cache[cache_key])
        try:
            wb=load_workbook(path,data_only=True)
            vals=[]
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        if isinstance(cell.value,str):
                            m=re.search(r"\bBAY\s*0*(\d+)\b",cell.value.upper())
                            if m: vals.append(int(m.group(1)))
            wb.close()
            odd=set()
            for b in vals:
                if b>0:
                    if b%2==0: odd.update([b-1,b+1])
                    else: odd.add(b)
            result = sorted([b for b in odd if b>0 and b%2==1],reverse=True)
            self._skeleton_legacy_cache[cache_key] = list(result)
            return result
        except Exception:
            return []

    def extract_skeleton_rows(self):
        """Read rows from structured POSITIONS, falling back to legacy graphical sheets."""
        structured = self.extract_skeleton_positions()
        if structured: return sorted({p["row"] for p in structured})
        path=clean(self.config_data.get("skeleton_path",""))
        if not path or not Path(path).exists(): return []
        fingerprint = self._skeleton_fingerprint(path)
        cache_key = ("rows", fingerprint)
        if cache_key in self._skeleton_legacy_cache:
            return list(self._skeleton_legacy_cache[cache_key])
        try:
            wb=load_workbook(path,data_only=True)
            result=set()
            for ws in wb.worksheets:
                for header_row in range(1, ws.max_row+1):
                    has_bay_label=any(isinstance(ws.cell(header_row,col).value,str) and re.search(r"\bBAY\s*0*\d+\b",ws.cell(header_row,col).value.upper()) for col in range(1,ws.max_column+1))
                    if not has_bay_label: continue
                    # In the agreed Bay Plan layout, the first row with several small numeric values under Bay headings contains Row positions.
                    for rr in range(header_row+1, min(ws.max_row,header_row+5)+1):
                        values=[]
                        for col in range(1,ws.max_column+1):
                            v=ws.cell(rr,col).value
                            if isinstance(v,(int,float)) and 0 <= int(v) <= 60: values.append(int(v))
                        if len(values)>=3 and len(set(values))>=2:
                            result.update(values)
                            break
            wb.close()
            result = sorted(result)
            self._skeleton_legacy_cache[cache_key] = list(result)
            return result
        except Exception:
            return []

    def extract_skeleton_tiers(self):
        """Read tiers from structured POSITIONS, falling back to legacy graphical sheets."""
        structured = self.extract_skeleton_positions()
        if structured: return sorted({p["tier"] for p in structured})
        path=clean(self.config_data.get("skeleton_path",""))
        if not path or not Path(path).exists(): return []
        fingerprint = self._skeleton_fingerprint(path)
        cache_key = ("tiers", fingerprint)
        if cache_key in self._skeleton_legacy_cache:
            return list(self._skeleton_legacy_cache[cache_key])
        try:
            wb=load_workbook(path,data_only=True)
            result=set()
            for ws in wb.worksheets:
                for header_row in range(1, ws.max_row+1):
                    has_bay_label=any(isinstance(ws.cell(header_row,col).value,str) and re.search(r"\bBAY\s*0*\d+\b",ws.cell(header_row,col).value.upper()) for col in range(1,ws.max_column+1))
                    if not has_bay_label: continue
                    # Look for numeric tier labels in column A following the row-position markers.
                    for rr in range(header_row+2, min(ws.max_row,header_row+8)+1):
                        value=ws.cell(rr,1).value
                        if isinstance(value,(int,float)) and 0 <= int(value) <= 100:
                            result.add(int(value))
                        elif isinstance(value,str) and re.search(r"\bBAY\s*0*\d+\b",value.upper()):
                            break
            wb.close()
            result = sorted(result)
            self._skeleton_legacy_cache[cache_key] = list(result)
            return result
        except Exception:
            return []

    def refresh_editor(self):
        if not hasattr(self,"editor_tree"):return
        for item in self.editor_tree.get_children(): self.editor_tree.delete(item)
        field, needle = getattr(self, "_active_filter", ("", ""))
        shown = 0
        for index,rec in enumerate(self.records):
            rec=normalize_record(rec)
            if needle:
                if field and field != "__ALL__":
                    haystack = clean(rec.get(field, ""))
                else:
                    haystack = " ".join(clean(rec.get(f, "")) for f in FIELDS)
                if needle.casefold() not in haystack.casefold():
                    continue
            vals=[]
            for f in FIELDS:
                val=rec.get(f,"")
                if f in ("BAY","ROW","TIER","GROSS_KG") and clean(val)!="": val=num(val)
                vals.append(val)
            self.editor_tree.insert("","end",iid=str(index),values=vals)
            shown += 1
        total = len(self.records)
        self.editor_count.set(f"Показано: {shown} из {total} контейнеров")
        self.update_current_display()

    def selected_record_index(self):
        sel=self.editor_tree.selection() if hasattr(self,"editor_tree") else ()
        if not sel:return None
        try:return int(sel[0])
        except ValueError:return None

    def add_container(self):
        dlg=ContainerDialog(self, title="Добавить контейнер")
        self.wait_window(dlg)
        if dlg.result:
            self.records.append(dlg.result); self.dirty=True; self._write_active_project_snapshot(); self.refresh_editor(); self.refresh_plan(); self.refresh_stats(); self.set_status("Добавлен контейнер " + clean(dlg.result.get("CONTAINER_NO")))

    def edit_vessel(self):
        keys=["VESSEL_NAME","VESSEL_ID_TYPE","VESSEL_ID","VOYAGE_NO","VESSEL_OPERATOR","SENDER_ID","RECEIVER_ID","MESSAGE_VERSION","MESSAGE_FUNCTION","PORT_OF_DEPARTURE","PORT_OF_ARRIVAL","PORT_TERMINAL","MESSAGE_DATE_TIME"]
        labels={"VESSEL_NAME":"Название судна","VESSEL_ID_TYPE":"Тип идентификатора","VESSEL_ID":"IMO (7 цифр) / ENI (8 цифр)","VOYAGE_NO":"Номер рейса","VESSEL_OPERATOR":"Оператор / линия","SENDER_ID":"EDI Sender ID","RECEIVER_ID":"EDI Receiver ID","MESSAGE_VERSION":"Версия BAPLIE","MESSAGE_FUNCTION":"Функция сообщения","PORT_OF_DEPARTURE":"Порт отправления (POL)","PORT_OF_ARRIVAL":"Порт прибытия (POD)","PORT_TERMINAL":"Терминал","MESSAGE_DATE_TIME":"Дата/время сообщения"}
        win=tk.Toplevel(self); win.title("Данные судна и EDI"); win.transient(self); win.grab_set(); win.resizable(True,False)
        frame=ttk.Frame(win,padding=16); frame.pack(fill="both",expand=True)
        vars_={}
        for i,key in enumerate(keys):
            row=i//2; col=(i%2)*2
            ttk.Label(frame,text=labels.get(key,key)).grid(row=row,column=col,sticky="w",padx=(0,8),pady=5)
            initial = clean(self.vessel.get(key,""))
            if key == "VESSEL_ID_TYPE": initial = initial or "IMO"
            if key == "VESSEL_ID" and not initial: initial = clean(self.vessel.get("IMO", ""))
            var=tk.StringVar(value=initial); vars_[key]=var
            if key == "VESSEL_ID_TYPE":
                ttk.Combobox(frame,textvariable=var,values=("IMO","ENI"),state="readonly",width=31).grid(row=row,column=col+1,sticky="ew",padx=(0,16),pady=5)
            else:
                ttk.Entry(frame,textvariable=var,width=34).grid(row=row,column=col+1,sticky="ew",padx=(0,16),pady=5)
        for c in range(4): frame.columnconfigure(c,weight=1)
        buttons=ttk.Frame(frame); buttons.grid(row=(len(keys)+1)//2,column=0,columnspan=4,sticky="e",pady=(12,0))
        def commit():
            new_id_type = vars_["VESSEL_ID_TYPE"].get().strip().upper() or "IMO"
            new_id = vars_["VESSEL_ID"].get().strip()
            if new_id and new_id_type == "IMO":
                if not re.fullmatch(r"\d{7}", new_id):
                    messagebox.showerror("Неверный IMO", "IMO должен содержать ровно 7 цифр без префикса IMO.", parent=win); return
                checksum = sum(int(new_id[i]) * (7 - i) for i in range(6)) % 10
                if checksum != int(new_id[6]):
                    messagebox.showerror("Неверный IMO", "Контрольная цифра IMO неверна. Проверьте номер судна.", parent=win); return
            elif new_id and new_id_type == "ENI" and not re.fullmatch(r"\d{8}", new_id):
                messagebox.showerror("Неверный ENI", "ENI должен содержать ровно 8 цифр.", parent=win); return
            for key,var in vars_.items(): self.vessel[key]=var.get().strip()
            self.vessel["VESSEL_ID_TYPE"] = new_id_type
            self.vessel["VESSEL_ID"] = new_id
            self.vessel["IMO"] = new_id if new_id_type == "IMO" else ""
            self.dirty=True; self._write_active_project_snapshot(); self.update_current_display(); self.set_status("Данные судна изменены"); win.destroy()
        ttk.Button(buttons,text="Отмена",command=win.destroy).pack(side="right",padx=(8,0))
        ttk.Button(buttons,text="Сохранить изменения",command=commit).pack(side="right")
        win.update_idletasks(); win.geometry(f"920x{min(650,win.winfo_reqheight()+10)}"); win.focus_force()

    def edit_container(self):
        idx=self.selected_record_index()
        if idx is None:
            messagebox.showinfo("Редактор","Выбери контейнер для изменения.",parent=self);return
        dlg=ContainerDialog(self,self.records[idx],"Изменить контейнер")
        self.wait_window(dlg)
        if dlg.result:
            self.records[idx]=dlg.result; self.dirty=True; self._write_active_project_snapshot(); self.refresh_editor(); self.refresh_plan(); self.refresh_stats()
            if self.editor_tree.exists(str(idx)): self.editor_tree.selection_set(str(idx))
            self.set_status("Изменения внесены")

    def delete_container(self):
        idx=self.selected_record_index()
        if idx is None:
            messagebox.showinfo("Редактор","Выбери контейнер для удаления.",parent=self);return
        cn=clean(self.records[idx].get("CONTAINER_NO"))
        if not messagebox.askyesno("Удалить контейнер",f"Удалить контейнер {cn} из текущего документа?",parent=self):return
        del self.records[idx]; self.dirty=True; self._write_active_project_snapshot(); self.refresh_editor(); self.refresh_plan(); self.refresh_stats(); self.set_status(f"Удалён контейнер {cn}")

    def _detailed_issues_for_current_plan(self, records=None):
        records = self.records if records is None else records
        return detailed_validation_issues(
            records,
            available_bays=self.extract_skeleton_bays(),
            available_rows=self.extract_skeleton_rows(),
            available_tiers=self.extract_skeleton_tiers(),
            reference_data=self.config_data.get("reference_data", DEFAULT_REFERENCE_DATA),
        )

    def check_plan(self):
        issues=self._detailed_issues_for_current_plan()
        if not issues:
            messagebox.showinfo("Проверка плана",f"Комплексная проверка завершена. Замечаний не найдено.\nКонтейнеров: {len(self.records)}\n\nПроверены номера и ISO-коды, обязательные поля, координаты, масса, занятость позиций, правило 20/40/45 ft и выбранный скелет судна. Проверка не заменяет подтверждение EDI по правилам терминала.",parent=self)
            return
        self.show_validation_dialog(issues,for_export=False)

    def show_validation_dialog(self, issues, for_export=False):
        result={"continue":False}
        win=tk.Toplevel(self); win.title("Результаты проверки плана"); win.transient(self); win.grab_set(); win.geometry("1050x560")
        frame=ttk.Frame(win,padding=14); frame.pack(fill="both",expand=True)
        errors=sum(1 for issue in issues if issue["severity"]=="Ошибка")
        warnings=len(issues)-errors
        ttk.Label(frame,text=f"Найдено замечаний: {len(issues)}   •   Ошибки: {errors}   •   Предупреждения: {warnings}",font=("Segoe UI",11,"bold"),foreground="#17365D").pack(anchor="w",pady=(0,8))
        ttk.Label(frame,text="Выбери замечание и нажми «Перейти к контейнеру», чтобы открыть запись и подсветить её позицию на плане.",wraplength=1000).pack(anchor="w",pady=(0,8))
        cols=("severity","container","location","message")
        tree=ttk.Treeview(frame,columns=cols,show="headings",selectmode="browse")
        for col,title,width in [("severity","Уровень",100),("container","Контейнер",155),("location","Позиция",210),("message","Описание замечания",520)]:
            tree.heading(col,text=title); tree.column(col,width=width,anchor="w",stretch=col=="message")
        sy=ttk.Scrollbar(frame,orient="vertical",command=tree.yview); sx=ttk.Scrollbar(frame,orient="horizontal",command=tree.xview)
        tree.configure(yscrollcommand=sy.set,xscrollcommand=sx.set)
        tree.pack(fill="both",expand=True)
        tree.bind("<Double-1>",lambda _e:jump_to_issue())
        for i,issue in enumerate(issues):
            tree.insert("","end",iid=str(i),values=(issue["severity"],issue["container_no"],issue["location"],issue["message"]))
        if issues: tree.selection_set("0")
        def selected_issue():
            sel=tree.selection()
            if not sel:return None
            try:return issues[int(sel[0])]
            except (ValueError,IndexError):return None
        def jump_to_issue():
            item=selected_issue()
            if item is None:return
            idx=item.get("index",-1)
            win.destroy()
            if idx is not None and 0<=idx<len(self.records):
                self.navigate_to_record(idx,open_editor=True)
        buttons=ttk.Frame(frame); buttons.pack(fill="x",pady=(10,0))
        ttk.Button(buttons,text="Перейти к контейнеру",command=jump_to_issue).pack(side="left")
        if for_export:
            def continue_anyway():
                if errors and not messagebox.askyesno("Экспорт с ошибками",f"Найдено {errors} ошибок и {warnings} предупреждений. Экспорт всё равно может создать некорректный файл. Продолжить?",parent=win):
                    return
                result["continue"]=True
                win.destroy()
            ttk.Button(buttons,text="Экспортировать несмотря на замечания…",command=continue_anyway).pack(side="right")
        ttk.Button(buttons,text="Закрыть",command=win.destroy).pack(side="right",padx=(0,8))
        self.wait_window(win)
        return result["continue"]

    def validate_before_edi_export(self, records):
        issues=self._detailed_issues_for_current_plan(records)
        if not issues:
            self.set_status("Проверка перед экспортом EDI завершена без замечаний")
            return True
        return bool(self.show_validation_dialog(issues,for_export=True))

    def _short_port_code(self, code):
        code=clean(code).upper()
        for item in self.config_data.get("reference_data",{}).get("Порты",[]):
            if clean(item.get("key")).upper()==code:
                return clean(item.get("short")).upper() or "--"
        return "--"

    def _plan_container_label(self, record, include_tier=False):
        """Return the full or compact two-line label for a plan cell."""
        mode = self.plan_detail_mode.get() if hasattr(self, "plan_detail_mode") else "Full Plan"
        kg = num(record.get("GROSS_KG"), 0)
        if mode == "Stowage Plan":
            if self.plan_short_top.get() == "Порт выгрузки":
                top = self._short_port_code(record.get("POD"))
            else:
                top = short_container_type(record.get("ISO_TYPE"), self.config_data.get("reference_data"))
            bottom = str(int((kg + 500) // 1000)) if kg > 0 else "--"
            label = f"{top}\n{bottom}"
        else:
            cn = clean(record.get("CONTAINER_NO")); iso = clean(record.get("ISO_TYPE"))
            weight = f"{kg/1000:.1f} t" if kg else ""
            label = "\n".join(z for z in [cn, iso, weight, clean(record.get("POD"))] if z)
        if include_tier:
            label = f"TIER {num(record.get('TIER')):02d}\n" + label
        return label

    def refresh_plan(self):
        if not hasattr(self, "plan_canvas"):
            return
        try:
            self.plan_canvas.delete("all")
        except Exception:
            return

        mode = self.plan_view_mode.get() if hasattr(self, "plan_view_mode") else "Tier"
        self._rebuild_plan_color_map(self.records)
        self._update_plan_legend()
        existing_tiers = {num(r.get("TIER")) for r in self.records if clean(r.get("TIER")) != "" and num(r.get("TIER")) >= 0}
        skeleton_tiers = set(self.extract_skeleton_tiers())
        all_tiers = set(existing_tiers) | skeleton_tiers
        # Without a vessel skeleton offer standard even Tier values as editable target positions.
        if not skeleton_tiers:
            all_tiers.update(range(2, 100, 2))
        tiers = sorted(all_tiers, reverse=True) or [2]

        if mode == "Tier":
            self.plan_selection_label.set("Tier:")
            tier_values = [str(t) for t in sorted(all_tiers)] or ["2"]
            combo_values = ["Все"] + tier_values
            self.tier_combo["values"] = combo_values
            selected = self.tier_choice.get()
            if selected not in combo_values:
                actual_tiers = sorted(existing_tiers | skeleton_tiers)
                selected = str(actual_tiers[0]) if actual_tiers else "2"
                self.tier_choice.set(selected)
            shown = [] if self.skeleton_only_mode.get() else (self.records[:] if selected == "Все" else [r for r in self.records if str(num(r.get("TIER"))) == selected])

            skeleton_bays = set(self.extract_skeleton_bays())
            bays = set(skeleton_bays)
            for rec in shown:
                b = num(rec.get("BAY"))
                if b > 0:
                    if b % 2 == 0:
                        bays.update([b - 1, b + 1])
                    else:
                        bays.add(b)
            if not skeleton_bays:
                max_bay = max(bays or {19})
                bays.update(range(1, max(26, max_bay + 8), 2))
            if not bays:
                bays = set(range(19, 0, -2))
            bays = sorted([b for b in bays if b > 0 and b % 2 == 1], reverse=True)
            rows = set(num(r.get("ROW")) for r in shown if clean(r.get("ROW")) != "")
            skeleton_rows = set(self.extract_skeleton_rows())
            rows.update(skeleton_rows)
            if not skeleton_rows:
                max_row = max(rows or {12})
                rows.update(range(0, max(13, max_row + 5)))
            rows = row_order([{"ROW": r} for r in rows]) if rows else [0]
            cell_w, cell_h, x0, y0 = 146, 75, 86, 52
            self.plan_canvas.create_text(16, 18, anchor="w", text=f"{'ПУСТОЙ СКЕЛЕТ' if self.skeleton_only_mode.get() else 'План судна'} — Tier {selected} | контейнеров в отображении: {len(shown)}", font=("Segoe UI", 12, "bold"), fill="#17365D")
            self.plan_canvas.create_text(16, 36, anchor="w", text="Дважды щёлкните по контейнеру, чтобы открыть его в редакторе", font=("Segoe UI", 9), fill="#526477")
            for j, b in enumerate(bays):
                x = x0 + j * cell_w
                self.plan_canvas.create_rectangle(x, y0 - cell_h + 10, x + cell_w, y0, fill="#176DA5", outline="#FFFFFF")
                self.plan_canvas.create_text(x + cell_w / 2, y0 - cell_h / 2 + 10, text=f"BAY {b:02d}", fill="white", font=("Segoe UI", 10, "bold"))
            rowmap = {row: i for i, row in enumerate(rows)}
            baymap = {b: i for i, b in enumerate(bays)}
            structured_positions = self.extract_skeleton_positions()
            for i, row in enumerate(rows):
                y = y0 + i * cell_h
                self.plan_canvas.create_rectangle(5, y, x0 - 3, y + cell_h, fill="#DCEAF7", outline="#C5D0D9")
                self.plan_canvas.create_text((x0 + 5) / 2, y + cell_h / 2, text=f"{row:02d}", fill="#17365D", font=("Segoe UI", 9, "bold"))
                for j, b in enumerate(bays):
                    x = x0 + j * cell_w
                    target_tier = None if selected == "Все" else num(selected)
                    matches = [p for p in structured_positions if p["row"] == row and p["tier"] == target_tier and (p["bay"] == b or (p["bay"] % 2 == 0 and b in (p["bay"]-1,p["bay"]+1)))] if target_tier is not None else []
                    status = (matches[0].get("status", "") if matches else "")
                    if structured_positions and target_tier is not None and not matches:
                        cell_fill = "#E7E6E6"  # no listed position at this coordinate
                    else:
                        cell_fill = "#E2F0D9" if status.lower() in ("allowed", "допустима", "разрешена") else ("#FCE4D6" if status.lower() in ("prohibited", "запрещена", "forbidden") else ("#FFF2CC" if status.lower() in ("restricted", "ограничена") else "#FFFFFF"))
                    cell = self.plan_canvas.create_rectangle(x, y, x + cell_w, y + cell_h, fill=cell_fill, outline="#D2DCE4")
                    self._bind_plan_slot(cell, "Tier", b, row, target_tier)
            for rec in shown:
                b, r = num(rec.get("BAY")), num(rec.get("ROW"))
                if r not in rowmap:
                    continue
                if b % 2 == 0:
                    pair = [baymap.get(b - 1), baymap.get(b + 1)]
                    if None in pair:
                        continue
                    j = min(pair); span = 2
                    x = x0 + j * cell_w + 2; w = span * cell_w - 4
                    fill = "#FCE4D6"
                else:
                    j = baymap.get(b)
                    if j is None:
                        continue
                    x = x0 + j * cell_w + 2; w = cell_w - 4; fill = "#DCEAF7"
                y = y0 + rowmap[r] * cell_h + 2
                label = self._plan_container_label(rec, include_tier=(selected == "Все"))
                self._bind_plan_container(rec, x, y, w, cell_h - 4, label, fill)
            total_w = x0 + len(bays) * cell_w + 30
            total_h = y0 + len(rows) * cell_h + 30
            self.plan_canvas.configure(scrollregion=(0, 0, max(total_w, 800), max(total_h, 400)))
            return

        # Bay mode: show the selected transverse bay section with tiers vertically and rows horizontally.
        self.plan_selection_label.set("Секция Bay:")
        skeleton_bays = set(self.extract_skeleton_bays())
        raw_bays = set(skeleton_bays)
        raw_bays.update(num(r.get("BAY")) for r in self.records if num(r.get("BAY")) > 0)
        if not skeleton_bays:
            max_bay = max(raw_bays or {25})
            max_bay = max(26, max_bay + 8)
            raw_bays.update(range(1, max_bay, 2))
        centers = set(section_for_bay(b) for b in raw_bays if b > 0)
        # Make empty centre-Bay positions selectable as targets, too.
        if not skeleton_bays:
            centers.update(range(2, max(raw_bays or {25}) + 1, 2))
        centers = sorted(centers)
        if not centers:
            centers = [2]
        combo_values = [f"{b:02d}" for b in centers]
        self.tier_combo["values"] = combo_values
        selected_value = self.tier_choice.get()
        if selected_value not in combo_values:
            selected_value = combo_values[0]
            self.tier_choice.set(selected_value)
        center = int(selected_value)
        shown = [] if self.skeleton_only_mode.get() else [r for r in self.records if num(r.get("BAY")) > 0 and section_for_bay(num(r.get("BAY"))) == center]
        left_bay, right_bay = center + 1, center - 1
        rows = set(num(r.get("ROW")) for r in shown if clean(r.get("ROW")) != "")
        skeleton_rows = set(self.extract_skeleton_rows())
        rows.update(skeleton_rows)
        if not skeleton_rows:
            max_row = max([num(r.get("ROW")) for r in self.records if clean(r.get("ROW")) != ""] or [12])
            rows.update(range(0, max(13, max_row + 5)))
        rows = row_order([{"ROW": r} for r in rows]) if rows else [0]
        section_tiers = {num(r.get("TIER")) for r in self.records if clean(r.get("TIER")) != "" and num(r.get("TIER")) >= 0}
        skeleton_tiers = set(self.extract_skeleton_tiers())
        section_tiers.update(skeleton_tiers)
        if not skeleton_tiers:
            max_tier = max(section_tiers or {12})
            section_tiers.update(range(2, max(14, max_tier + 5), 2))
        levels = sorted(section_tiers, reverse=True) or tiers
        # Compact square cells are only for abbreviated Stowage Plan.
        # Full Plan keeps wide cells so full container numbers and details fit.
        if self.plan_detail_mode.get() == "Stowage Plan":
            cell_w = cell_h = 34
        else:
            cell_w, cell_h = 126, 66
        x0, y0 = 82, 106
        left_width = len(rows) * cell_w
        full_width = x0 + left_width * 2 + 28
        self.plan_canvas.create_text(16, 18, anchor="w", text=f"{'ПУСТОЙ СКЕЛЕТ' if self.skeleton_only_mode.get() else 'План судна'} — Bay {center:02d} (позиции {left_bay:02d} и {right_bay:02d}) | контейнеров: {len(shown)}", font=("Segoe UI", 12, "bold"), fill="#17365D")
        self.plan_canvas.create_text(16, 38, anchor="w", text="Поперечная схема: Tier по вертикали, Row по горизонтали. Для 40/45 ft отметка X показывает вторую занятую позицию.", font=("Segoe UI", 9), fill="#526477")
        self.plan_canvas.create_rectangle(x0, 54, x0 + left_width, 78, fill="#176DA5", outline="#FFFFFF")
        self.plan_canvas.create_text(x0 + left_width / 2, 66, text=f"BAY {left_bay:02d} — левая сторона", fill="white", font=("Segoe UI", 10, "bold"))
        self.plan_canvas.create_rectangle(x0 + left_width, 54, x0 + left_width * 2, 78, fill="#176DA5", outline="#FFFFFF")
        self.plan_canvas.create_text(x0 + left_width * 1.5, 66, text=f"BAY {right_bay:02d} — правая сторона", fill="white", font=("Segoe UI", 10, "bold"))
        rowmap = {r: i for i, r in enumerate(rows)}
        tiermap = {t: i for i, t in enumerate(levels)}
        self.plan_canvas.create_text(36, y0 + 12, text="TIER", fill="#17365D", font=("Segoe UI", 9, "bold"))
        for side_offset in (0, len(rows)):
            for i, row in enumerate(rows):
                x = x0 + (side_offset + i) * cell_w
                self.plan_canvas.create_rectangle(x, 80, x + cell_w, y0, fill="#DCEAF7", outline="#C5D0D9")
                self.plan_canvas.create_text(x + cell_w / 2, 93, text=f"{row:02d}", fill="#17365D", font=("Segoe UI", 9, "bold"))
        for i, tier in enumerate(levels):
            y = y0 + i * cell_h
            self.plan_canvas.create_rectangle(5, y, x0 - 4, y + cell_h, fill="#DCEAF7", outline="#C5D0D9")
            self.plan_canvas.create_text((x0 + 5) / 2, y + cell_h / 2, text=f"{tier:02d}", fill="#17365D", font=("Segoe UI", 10, "bold"))
            for j in range(len(rows) * 2):
                x = x0 + j * cell_w
                cell = self.plan_canvas.create_rectangle(x, y, x + cell_w, y + cell_h, fill="#FFFFFF", outline="#000000", width=1)
                side_bay = left_bay if j < len(rows) else right_bay
                row_value = rows[j % len(rows)]
                self._bind_plan_slot(cell, "Bay", side_bay, row_value, tier)
        # Thick hatch-cover boundary between deck tiers (80+) and hold tiers (<80).
        deck_indices = [i for i, tier in enumerate(levels) if tier >= 80]
        hold_indices = [i for i, tier in enumerate(levels) if tier < 80]
        if deck_indices and hold_indices:
            boundary_i = max(deck_indices) + 1 if max(deck_indices) < min(hold_indices) else max(hold_indices) + 1
            boundary_y = y0 + boundary_i * cell_h
            self.plan_canvas.create_line(x0, boundary_y, x0 + len(rows) * 2 * cell_w, boundary_y, fill="#000000", width=5)
            self.plan_canvas.create_text(x0 + 6, boundary_y - 5, text="HATCH COVER / DECK-HOLD BOUNDARY", anchor="sw", fill="#000000", font=("Segoe UI", 8, "bold"))
        # Dark-grey divider between the two sides of the selected Bay section.
        separator_x = x0 + len(rows) * cell_w
        self.plan_canvas.create_line(separator_x, 54, separator_x, y0 + len(levels) * cell_h, fill="#555F68", width=3)
        for rec in shown:
            b, r, t = num(rec.get("BAY")), num(rec.get("ROW")), num(rec.get("TIER"))
            if r not in rowmap or t not in tiermap:
                continue
            ri, ti = rowmap[r], tiermap[t]
            if b % 2 == 0:
                # A 40/45 ft entry at even Bay occupies the two adjacent odd Bay positions.
                right_x = x0 + (len(rows) + ri) * cell_w + 2
                left_x = x0 + ri * cell_w + 2
                y = y0 + ti * cell_h + 2
                label = self._plan_container_label(rec)
                self._bind_plan_container(rec, right_x, y, cell_w - 4, cell_h - 4, label, "#FCE4D6")
                x_fill = self._record_color(rec, fallback="#FCE4D6")
                rec_idx=self._record_index_by_identity(rec)
                hi=self._is_search_match(rec) or (self.move_mode_active and self.move_record_index==rec_idx)
                edge="#168A4A" if self.move_mode_active and self.move_record_index==rec_idx else ("#F28C28" if self._is_search_match(rec) else "#8FA9BC")
                tags=("plan_container",f"record_{rec_idx}")
                x_rect = self.plan_canvas.create_rectangle(left_x, y, left_x + cell_w - 4, y + cell_h - 4, fill=x_fill, outline=edge, width=4 if hi else 1, tags=tags)
                xmark = self.plan_canvas.create_text(left_x + (cell_w - 4) / 2, y + (cell_h - 4) / 2, text="X", fill="#C00000", font=("Segoe UI", 24, "bold"), tags=tags)
                self.plan_canvas.tag_bind(x_rect, "<Button-1>", lambda _e, record=rec: self.on_plan_container_click(record))
                self.plan_canvas.tag_bind(xmark, "<Button-1>", lambda _e, record=rec: self.on_plan_container_click(record))
            else:
                side_index = None
                if b == left_bay:
                    side_index = ri
                elif b == right_bay:
                    side_index = len(rows) + ri
                if side_index is None:
                    continue
                x = x0 + side_index * cell_w + 2; y = y0 + ti * cell_h + 2
                label = self._plan_container_label(rec)
                self._bind_plan_container(rec, x, y, cell_w - 4, cell_h - 4, label, "#DCEAF7")
        total_h = y0 + len(levels) * cell_h + 30
        self.plan_canvas.configure(scrollregion=(0, 0, max(full_width, 800), max(total_h, 400)))

    def _record_index_by_identity(self, record):
        for i, item in enumerate(self.records):
            if item is record:
                return i
        return None

    def _is_search_match(self, record):
        idx=self._record_index_by_identity(record)
        return idx is not None and idx in self._search_match_indices

    def _bind_plan_slot(self, canvas_item, view_mode, bay, row, tier):
        self.plan_canvas.tag_bind(canvas_item,"<Button-1>",
            lambda _e, mode=view_mode, b=bay, r=row, t=tier: self.move_to_slot(mode,b,r,t))

    def _bind_plan_container(self, record, x, y, width, height, label, fill):
        fill = self._record_color(record, fallback=fill)
        idx=self._record_index_by_identity(record)
        matched=self._is_search_match(record)
        moving=(self.move_mode_active and self.move_record_index is not None and idx==self.move_record_index)
        if getattr(self, "plan_view_mode", None) is not None and self.plan_view_mode.get() == "Bay":
            outline="#168A4A" if moving else ("#F28C28" if matched else "#000000")
            line_width=4 if (moving or matched) else 1
        else:
            outline="#168A4A" if moving else ("#F28C28" if matched else "#8FA9BC")
            line_width=4 if (moving or matched) else 1
        tags=("plan_container",f"record_{idx}") if idx is not None else ("plan_container",)
        rect = self.plan_canvas.create_rectangle(x, y, x + width, y + height, fill=fill, outline=outline, width=line_width, tags=tags)
        font_size = 7 if self.plan_detail_mode.get() == "Stowage Plan" and self.plan_view_mode.get() == "Bay" else 8
        text = self.plan_canvas.create_text(x + width / 2, y + height / 2, text=label, fill="#18344D", font=("Segoe UI", font_size, "bold"), width=width - 6, justify="center", tags=tags)
        self.plan_canvas.tag_bind(rect, "<Button-1>", lambda _e, rec=record: self.on_plan_container_click(rec))
        self.plan_canvas.tag_bind(text, "<Button-1>", lambda _e, rec=record: self.on_plan_container_click(rec))

    def on_plan_container_click(self, record):
        idx=self._record_index_by_identity(record)
        if idx is None:return
        if self.move_mode_active:
            if self.move_record_index is None:
                self.move_record_index=idx
                self.set_status(f"Выбран {clean(record.get('CONTAINER_NO'))}. Нажми свободную позицию для перемещения.")
                self.refresh_plan()
            elif self.move_record_index==idx:
                self.move_record_index=None
                self.set_status("Выбор контейнера для перемещения снят")
                self.refresh_plan()
            else:
                messagebox.showwarning("Позиция занята","Для перемещения выбери пустую позицию. Чтобы выбрать другой контейнер, нажми текущий выбранный контейнер для отмены выбора.",parent=self)
            return
        self.open_record_from_plan(record)

    def toggle_move_mode(self):
        self.move_mode_active=not self.move_mode_active
        self.move_record_index=None
        self.move_mode_label.set("Перемещение: ВКЛ" if self.move_mode_active else "Перемещение: выключено")
        self.set_status("Режим перемещения включён: выбери контейнер и затем свободную позицию." if self.move_mode_active else "Режим перемещения выключен")
        self.refresh_plan()

    def _record_slots_available(self, index, bay, row, tier):
        candidate=dict(self.records[index]); candidate.update({"BAY":bay,"ROW":row,"TIER":tier})
        target=set(position_footprint(candidate))
        if not target:return False
        for i,other in enumerate(self.records):
            if i==index:continue
            if target.intersection(position_footprint(other)):return False
        return True

    def move_to_slot(self, view_mode, bay, row, tier):
        if not self.move_mode_active:return
        idx=self.move_record_index
        if idx is None:
            self.set_status("Сначала выбери контейнер, который нужно переместить.")
            return
        if not (0<=idx<len(self.records)):
            self.move_record_index=None; return
        source=normalize_record(self.records[idx])
        dest_row=num(row,-1)
        dest_tier=num(source.get("TIER"),0) if tier is None else num(tier,-1)
        size=teu_for_record(source)
        candidates=[]
        if view_mode=="Bay":
            center=num(self.tier_choice.get(),0)
            target_bay=center if size==2 else num(bay,0)
            candidates=[target_bay] if target_bay>0 else []
        elif size==2:
            clicked=num(bay,0)
            candidates=sorted(set([clicked-1,clicked+1]))
            candidates=[b for b in candidates if b>0 and b%2==0]
        else:
            target_bay=num(bay,0)
            candidates=[target_bay] if target_bay>0 and target_bay%2==1 else []
        if dest_row<0 or dest_tier<0 or not candidates:
            messagebox.showwarning("Недопустимая позиция","Не удалось определить корректную Bay / Row / Tier для выбранного контейнера.",parent=self); return
        viable=[b for b in candidates if self._record_slots_available(idx,b,dest_row,dest_tier)]
        if not viable:
            messagebox.showwarning("Позиция занята","Целевая позиция или часть занимаемых слотов занята другим контейнером. Выбери другую позицию.",parent=self); return
        target_bay=viable[0]
        if len(viable)>1:
            options="\n".join(f"Bay {b:02d} (позиции {b-1:02d} + {b+1:02d})" for b in viable)
            first=viable[0]
            answer=messagebox.askyesno("Выбор центрального Bay",f"Для 40/45 ft доступны две пары:\n{options}\n\nНажми «Да», чтобы выбрать Bay {first:02d}; «Нет» — выбрать Bay {viable[1]:02d}.",parent=self)
            target_bay=first if answer else viable[1]
            if not self._record_slots_available(idx,target_bay,dest_row,dest_tier):
                messagebox.showwarning("Позиция занята",f"Bay {target_bay:02d} уже занят. Выбери другую позицию.",parent=self);return
        old=(num(source.get("BAY")),num(source.get("ROW")),num(source.get("TIER")))
        new=(target_bay,dest_row,dest_tier)
        if old==new:
            self.set_status("Контейнер уже находится в этой позиции.");return
        updated=dict(source); updated.update({"BAY":target_bay,"ROW":dest_row,"TIER":dest_tier})
        self.records[idx]=normalize_record(updated)
        self.dirty=True; self._write_active_project_snapshot(); self.move_record_index=None
        self.refresh_editor(); self.refresh_plan(); self.refresh_stats(); self.update_current_display()
        self.set_status(f"Контейнер {clean(source.get('CONTAINER_NO'))} перемещён: Bay/Row/Tier {old} → {new}")

    def open_record_from_plan(self, record):
        idx=self._record_index_by_identity(record)
        if idx is None:return
        self.navigate_to_record(idx,open_editor=True)

    def navigate_to_record(self, idx, open_editor=False, force_bay=False):
        if idx is None or not (0<=idx<len(self.records)):return
        rec=self.records[idx]
        if force_bay:
            self.plan_view_mode.set("Bay")
        self.move_mode_active=False; self.move_record_index=None
        self.move_mode_label.set("Перемещение: выключено")
        self._search_match_indices={idx}
        self.nb.select(self.tab_plan)
        if self.plan_view_mode.get()=="Tier":
            self.tier_choice.set(str(num(rec.get("TIER"))))
        else:
            self.tier_choice.set(f"{section_for_bay(num(rec.get('BAY'))):02d}")
        self.refresh_plan(); self.update_idletasks()
        try:
            bbox=self.plan_canvas.bbox(f"record_{idx}")
            region=self.plan_canvas.cget("scrollregion").split()
            if bbox and len(region)==4:
                width=max(1,float(region[2])-float(region[0])); height=max(1,float(region[3])-float(region[1]))
                self.plan_canvas.xview_moveto(max(0,(bbox[0]-60)/width))
                self.plan_canvas.yview_moveto(max(0,(bbox[1]-50)/height))
        except Exception:pass
        self.set_status(f"На схеме выделен контейнер {clean(rec.get('CONTAINER_NO'))}")
        if open_editor:
            # Keep the highlighted plan visible behind the container parameter dialog.
            self.refresh_editor()
            if not self.editor_tree.exists(str(idx)):
                self.filter_field.set("Все поля"); self.filter_value.set(""); self._active_filter=("", "")
                self.refresh_editor()
            if self.editor_tree.exists(str(idx)):
                self.editor_tree.selection_set(str(idx)); self.editor_tree.see(str(idx))
            self.edit_container()

    def search_plan(self):
        query=self.search_value.get().strip()
        if not query:
            messagebox.showinfo("Поиск","Введи номер контейнера, порт, линию, тип или координату.",parent=self);return
        chosen=self.search_field.get()
        search_fields=FIELDS if chosen=="Все поля" else [chosen]
        matches=[]; needle=query.casefold()
        for idx,record in enumerate(self.records):
            if any(needle in clean(record.get(field,"")).casefold() for field in search_fields):
                matches.append(idx)
        if not matches:
            self._search_match_indices=set(); self.refresh_plan()
            messagebox.showinfo("Поиск",f"По запросу «{query}» ничего не найдено.",parent=self);return
        self.navigate_to_record(matches[0],open_editor=False,force_bay=True)
        self._search_match_indices=set(matches)
        if len(matches)>1 and self.plan_view_mode.get()=="Tier":
            self.tier_choice.set("Все")
        self.refresh_plan()
        self.set_status(f"Найдено {len(matches)} контейнер(ов) по запросу «{query}». Совпадения выделены оранжевой рамкой.")
        if len(matches)==1:return
        win=tk.Toplevel(self); win.title(f"Результаты поиска: {len(matches)}"); win.transient(self); win.geometry("940x460")
        frame=ttk.Frame(win,padding=12); frame.pack(fill="both",expand=True)
        tree=ttk.Treeview(frame,columns=("no","iso","bay","row","tier","pol","pot","pod","line"),show="headings",selectmode="browse")
        for col,title,width in [("no","Контейнер",145),("iso","ISO",70),("bay","Bay",55),("row","Row",55),("tier","Tier",55),("pol","POL",80),("pot","POT",80),("pod","POD",80),("line","Линия",100)]:
            tree.heading(col,text=title); tree.column(col,width=width,anchor="w")
        tree.pack(fill="both",expand=True); mapping={}
        for idx in matches:
            rec=self.records[idx]
            tree.insert("","end",iid=str(idx),values=(clean(rec.get("CONTAINER_NO")),clean(rec.get("ISO_TYPE")),clean(rec.get("BAY")),clean(rec.get("ROW")),clean(rec.get("TIER")),clean(rec.get("POL")),clean(rec.get("POT")),clean(rec.get("POD")),clean(rec.get("LINE"))))
            mapping[str(idx)]=idx
        if matches:tree.selection_set(str(matches[0]))
        def jump():
            sel=tree.selection()
            if not sel:return
            target=mapping[sel[0]]; win.destroy(); self.navigate_to_record(target,open_editor=False,force_bay=True)
            self._search_match_indices=set(matches); self.refresh_plan()
        def open_params():
            sel=tree.selection()
            if not sel:return
            target=mapping[sel[0]]; win.destroy(); self.navigate_to_record(target,open_editor=True,force_bay=True)
            self._search_match_indices=set(matches); self.refresh_plan()
        tree.bind("<Double-1>",lambda _e:jump())
        buttons=ttk.Frame(frame); buttons.pack(fill="x",pady=(8,0))
        ttk.Button(buttons,text="Показать на схеме",command=jump).pack(side="left")
        ttk.Button(buttons,text="Открыть параметры",command=open_params).pack(side="left",padx=(8,0))
        ttk.Button(buttons,text="Закрыть",command=win.destroy).pack(side="right")

    def clear_plan_search(self):
        self.search_value.set(""); self._search_match_indices=set(); self.refresh_plan()
        self.set_status("Поиск и выделение сброшены")

    def _normal_compare_value(self, field, value):
        if field in ("BAY","ROW","TIER","GROSS_KG") and clean(value)!="":
            try:return num(value)
            except Exception:pass
        return clean(value)

    def compare_with_source(self):
        original=[normalize_record(r) for r in self.source_records]
        current=[normalize_record(r) for r in self.records]
        before={}; after={}
        for i,rec in enumerate(original):
            key=clean(rec.get("CONTAINER_NO")).upper() or f"__EMPTY_BEFORE_{i}"
            before.setdefault(key,[]).append((i,rec))
        for i,rec in enumerate(current):
            key=clean(rec.get("CONTAINER_NO")).upper() or f"__EMPTY_NOW_{i}"
            after.setdefault(key,[]).append((i,rec))
        diff=[]
        for key in dict.fromkeys(list(before.keys())+list(after.keys())):
            left=before.get(key,[]); right=after.get(key,[]); matched=min(len(left),len(right))
            for n in range(matched):
                old_i,old=left[n]; new_i,new=right[n]; cn=clean(new.get("CONTAINER_NO")) or "(без номера)"
                old_coords=tuple(self._normal_compare_value(f,old.get(f)) for f in ("BAY","ROW","TIER"))
                new_coords=tuple(self._normal_compare_value(f,new.get(f)) for f in ("BAY","ROW","TIER"))
                if old_coords!=new_coords:
                    old_pos=f"Bay {num(old.get('BAY')):02d} / Row {num(old.get('ROW')):02d} / Tier {num(old.get('TIER')):02d}"
                    new_pos=f"Bay {num(new.get('BAY')):02d} / Row {num(new.get('ROW')):02d} / Tier {num(new.get('TIER')):02d}"
                    diff.append({"status":"Перемещён","container":cn,"field":"Bay / Row / Tier","old":old_pos,"new":new_pos,"index":new_i})
                for field in FIELDS:
                    if field in ("CONTAINER_NO","BAY","ROW","TIER"):continue
                    if self._normal_compare_value(field,old.get(field))!=self._normal_compare_value(field,new.get(field)):
                        diff.append({"status":"Изменён","container":cn,"field":field,"old":clean(old.get(field)),"new":clean(new.get(field)),"index":new_i})
            for _old_i,old in left[matched:]:
                diff.append({"status":"Удалён","container":clean(old.get("CONTAINER_NO")) or "(без номера)","field":"Запись","old":"Есть в исходном плане","new":"Удалён из текущего плана","index":None})
            for new_i,new in right[matched:]:
                diff.append({"status":"Добавлен","container":clean(new.get("CONTAINER_NO")) or "(без номера)","field":"Запись","old":"—","new":"Добавлен в текущий план","index":new_i})
        win=tk.Toplevel(self); win.title("Сравнение с исходным планом"); win.transient(self); win.geometry("1080x560")
        frame=ttk.Frame(win,padding=12); frame.pack(fill="both",expand=True)
        ttk.Label(frame,text=f"Источник сравнения: {self.source_label}   •   Изменений: {len(diff)}",font=("Segoe UI",11,"bold"),foreground="#17365D").pack(anchor="w",pady=(0,8))
        if not diff:ttk.Label(frame,text="Различий не обнаружено. Текущий план совпадает с исходным.").pack(anchor="w",pady=(0,8))
        tree=ttk.Treeview(frame,columns=("status","container","field","old","new"),show="headings",selectmode="browse")
        for col,title,width in [("status","Изменение",100),("container","Контейнер",150),("field","Поле",150),("old","Исходное значение",300),("new","Текущее значение",300)]:
            tree.heading(col,text=title); tree.column(col,width=width,anchor="w",stretch=col in ("old","new"))
        tree.pack(fill="both",expand=True); mapping={}
        for n,item in enumerate(diff):
            iid=str(n); mapping[iid]=item
            tree.insert("","end",iid=iid,values=(item["status"],item["container"],item["field"],item["old"],item["new"]))
        def jump_diff():
            sel=tree.selection()
            if not sel:return
            idx=mapping[sel[0]].get("index")
            if idx is None:
                messagebox.showinfo("Сравнение","Удалённая запись отсутствует в текущем плане; исходные значения показаны в отчёте.",parent=win);return
            win.destroy(); self.navigate_to_record(idx,open_editor=False)
        tree.bind("<Double-1>",lambda _e:jump_diff())
        buttons=ttk.Frame(frame); buttons.pack(fill="x",pady=(8,0))
        ttk.Button(buttons,text="Перейти к текущей записи",command=jump_diff).pack(side="left")
        ttk.Button(buttons,text="Закрыть",command=win.destroy).pack(side="right")

    def _font_for_image(self, size, bold=False):
        from PIL import ImageFont
        candidates = []
        if os.name == "nt":
            candidates += [r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf", r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf"]
        candidates += ["/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]
        for candidate in candidates:
            try: return ImageFont.truetype(candidate,size)
            except Exception: pass
        return ImageFont.load_default()

    def export_plan_jpg(self, print_after=False, _options=None, _return_image=False):
        """Export plan graphics with independent print options; supports multi-page PDFs."""
        try:
            from PIL import Image, ImageDraw
        except Exception:
            messagebox.showerror("Сохранение графики", "Для сохранения графики нужна библиотека Pillow.", parent=self); return
        if not self.records and not self.skeleton_only_mode.get():
            messagebox.showinfo("Сохранение графики", "Сначала открой документ с контейнерами или выбери пустой скелет.", parent=self); return

        if _options is None:
            win=tk.Toplevel(self); win.title("Параметры сохранения плана"); win.transient(self); win.grab_set(); win.resizable(False,False)
            box=ttk.Frame(win,padding=16); box.pack(fill="both",expand=True)
            view_var=tk.StringVar(value=self.plan_view_mode.get())
            detail_var=tk.StringVar(value=self.plan_detail_mode.get())
            scope_default="current"
            scope_var=tk.StringVar(value="Текущий Bay / Tier")
            ttk.Label(box,text="Графика").grid(row=0,column=0,sticky="w",pady=5)
            ttk.Combobox(box,textvariable=view_var,state="readonly",values=["Tier","Bay"],width=18).grid(row=0,column=1,sticky="ew",pady=5)
            ttk.Label(box,text="Режим плана").grid(row=1,column=0,sticky="w",pady=5)
            ttk.Combobox(box,textvariable=detail_var,state="readonly",values=["Full Plan","Stowage Plan"],width=18).grid(row=1,column=1,sticky="ew",pady=5)
            ttk.Label(box,text="Область сохранения").grid(row=2,column=0,sticky="w",pady=5)
            scope_combo=ttk.Combobox(box,textvariable=scope_var,state="readonly",values=["Текущий Bay / Tier","Все Bay судна","Все Tier"],width=22)
            scope_combo.grid(row=2,column=1,sticky="ew",pady=5)
            scope_labels={"Текущий Bay / Tier":"current","Все Bay судна":"all_bays","Все Tier":"all_tiers"}
            def update_scopes(*_):
                vals=["Текущий Bay / Tier","Все Bay судна"] if view_var.get()=="Bay" else ["Текущий Bay / Tier","Все Tier"]
                scope_combo.configure(values=vals)
                if scope_combo.get() not in vals: scope_var.set(vals[0])
            view_var.trace_add("write",update_scopes)
            ttk.Label(box,text="Все Bay будут размещены на отдельных страницах PDF (или объединены в длинное изображение для JPG/PNG).",wraplength=400,foreground="#526477").grid(row=3,column=0,columnspan=2,sticky="w",pady=(8,4))
            result={}
            def accept():
                result.update({"mode":view_var.get(),"detail":detail_var.get(),"scope":scope_labels.get(scope_var.get(),"current")})
                win.destroy()
            buttons=ttk.Frame(box); buttons.grid(row=4,column=0,columnspan=2,sticky="ew",pady=(10,0))
            ttk.Button(buttons,text="Сохранить…",style="Accent.TButton",command=accept).pack(side="right")
            ttk.Button(buttons,text="Отмена",command=win.destroy).pack(side="right",padx=(0,8))
            box.columnconfigure(1,weight=1)
            self.wait_window(win)
            if not result:return
            _options=result

        if not _return_image:
            mode=_options.get("mode",self.plan_view_mode.get())
            detail=_options.get("detail",self.plan_detail_mode.get())
            scope=_options.get("scope","current")
            old_mode=self.plan_view_mode.get(); old_detail=self.plan_detail_mode.get(); old_selection=self.tier_choice.get()
            self.plan_view_mode.set(mode); self.plan_detail_mode.set(detail)
            try:
                if mode=="Bay":
                    raw_bays=set(self.extract_skeleton_bays()); raw_bays.update(num(r.get("BAY")) for r in self.records if num(r.get("BAY"))>0)
                    centers=sorted({section_for_bay(b) for b in raw_bays if b>0}) or [2]
                    current=self.tier_choice.get()
                    if current not in [f"{b:02d}" for b in centers]: current=f"{centers[0]:02d}"
                    selections=[f"{b:02d}" for b in centers] if scope=="all_bays" else [current]
                else:
                    existing={num(r.get("TIER")) for r in self.records if clean(r.get("TIER"))!="" and num(r.get("TIER"))>=0}
                    existing.update(self.extract_skeleton_tiers())
                    tier_values=sorted(existing) or [2]
                    current=self.tier_choice.get()
                    if scope=="all_tiers": selections=["Все"]
                    else: selections=[current if current in ["Все"]+[str(t) for t in tier_values] else str(tier_values[0])]
                images=[]
                for selection in selections:
                    self.tier_choice.set(selection)
                    result=self.export_plan_jpg(_options={"mode":mode,"detail":detail,"scope":"current"},_return_image=True)
                    if result is not None: images.append(result)
                if not images:return
                stem=f"{self.current_path.stem if self.current_path else 'CVCS'}_{mode}_{detail.replace(' ','_')}"
                if len(images)>1: stem += "_ALL"
                p=filedialog.asksaveasfilename(title="Сохранить графический план",initialdir=str(EXPORTS_DIR),initialfile=f"{stem}.pdf",defaultextension=".pdf",filetypes=[("Документ PDF","*.pdf"),("Изображение JPEG","*.jpg *.jpeg"),("Изображение PNG","*.png")])
                if not p:return
                suffix=Path(p).suffix.lower()
                try:
                    if suffix==".pdf":
                        if mode=="Bay" and scope=="all_bays" and len(images)>1:
                            # Put up to six compact Bay views on each landscape page.
                            tile_w, tile_h = 1400, 900
                            pages=[]
                            for base in range(0,len(images),6):
                                page=Image.new("RGB",(tile_w*2, tile_h*3),"white")
                                for j,im in enumerate(images[base:base+6]):
                                    thumb=im.convert("RGB"); thumb.thumbnail((tile_w-24,tile_h-24),Image.Resampling.LANCZOS)
                                    x=(j%2)*tile_w+(tile_w-thumb.width)//2
                                    y=(j//2)*tile_h+(tile_h-thumb.height)//2
                                    page.paste(thumb,(x,y))
                                pages.append(page)
                            pages[0].save(p,"PDF",resolution=150.0,save_all=True,append_images=pages[1:])
                        else:
                            rgb=[im.convert("RGB") for im in images]
                            rgb[0].save(p,"PDF",resolution=150.0,save_all=True,append_images=rgb[1:])
                    elif suffix in (".png", ".jpg", ".jpeg"):
                        if len(images)==1:
                            images[0].save(p,"PNG" if suffix==".png" else "JPEG",optimize=True,**({} if suffix==".png" else {"quality":92}))
                        else:
                            tile_w, tile_h = 1400, 900
                            cols = 2
                            rows_n = (len(images) + cols - 1) // cols
                            merged=Image.new("RGB",(tile_w*cols,tile_h*rows_n),"white")
                            for i,im in enumerate(images):
                                thumb=im.convert("RGB"); thumb.thumbnail((tile_w-24,tile_h-24),Image.Resampling.LANCZOS)
                                x=(i%cols)*tile_w+(tile_w-thumb.width)//2
                                y=(i//cols)*tile_h+(tile_h-thumb.height)//2
                                merged.paste(thumb,(x,y))
                            merged.save(p,"PNG" if suffix==".png" else "JPEG",optimize=True,**({} if suffix==".png" else {"quality":92}))
                except Exception as e:
                    messagebox.showerror("Сохранение графики",f"Не удалось сохранить файл:\n{e}",parent=self); return
                self.set_status(f"Графический план сохранён: {Path(p).name}")
                messagebox.showinfo("Сохранение графики",f"План сохранён:\n{p}\nСтраниц / видов: {len(images)}",parent=self)
                return
            finally:
                self.plan_view_mode.set(old_mode); self.plan_detail_mode.set(old_detail); self.tier_choice.set(old_selection)
                if not _return_image:
                    self.refresh_plan()

        # Render one view to a raster image; called recursively above with _return_image=True.
        self._rebuild_plan_color_map(self.records)
        mode=self.plan_view_mode.get(); selected=self.tier_choice.get(); detail=self.plan_detail_mode.get()
        if mode=="Tier":
            tiers={num(r.get("TIER")) for r in self.records if clean(r.get("TIER"))!="" and num(r.get("TIER"))>=0}; tiers.update(self.extract_skeleton_tiers())
            if selected not in ["Все"]+[str(t) for t in sorted(tiers)]: selected=(str(sorted(tiers)[0]) if tiers else "2")
            shown=[] if self.skeleton_only_mode.get() else (self.records[:] if selected=="Все" else [r for r in self.records if str(num(r.get("TIER")))==selected])
            bays=set(self.extract_skeleton_bays())
            for rec in shown:
                b=num(rec.get("BAY"))
                if b>0: bays.update([b-1,b+1] if b%2==0 else [b])
            bays=sorted([b for b in bays if b>0 and b%2==1],reverse=True) or list(range(19,0,-2))
            rows=set(num(r.get("ROW")) for r in shown if clean(r.get("ROW"))!=""); rows.update(self.extract_skeleton_rows())
            rows=row_order([{"ROW":r} for r in rows]) if rows else list(range(13))
            cell_w,cell_h,x0,y0=146,75,86,160
            width=x0+len(bays)*cell_w+30; height=y0+len(rows)*cell_h+40
        else:
            raw_bays=set(self.extract_skeleton_bays()); raw_bays.update(num(r.get("BAY")) for r in self.records if num(r.get("BAY"))>0)
            centers=sorted({section_for_bay(b) for b in raw_bays if b>0}) or [2]
            center=int(selected) if selected.isdigit() else centers[0]
            shown=[] if self.skeleton_only_mode.get() else [r for r in self.records if num(r.get("BAY"))>0 and section_for_bay(num(r.get("BAY")))==center]
            left_bay,right_bay=center+1,center-1
            rows=set(num(r.get("ROW")) for r in shown if clean(r.get("ROW"))!=""); rows.update(self.extract_skeleton_rows())
            rows=row_order([{"ROW":r} for r in rows]) if rows else list(range(13))
            tiers={num(r.get("TIER")) for r in shown if clean(r.get("TIER"))!="" and num(r.get("TIER"))>=0}; tiers.update(self.extract_skeleton_tiers())
            levels=sorted(tiers,reverse=True) or [2]
            if detail=="Stowage Plan": cell_w=cell_h=34
            else: cell_w,cell_h=126,66
            x0,y0=82,142; left_width=len(rows)*cell_w
            width=x0+left_width*2+28; height=y0+len(levels)*cell_h+40
        scale=min(1.0,14000/max(width,1),10000/max(height,1)); scale=max(scale,0.45)
        sw,sh=max(800,int(width*scale)),max(400,int(height*scale))
        image=Image.new("RGB",(sw,sh),"white"); draw=ImageDraw.Draw(image)
        font_title=self._font_for_image(max(12,int(18*scale)),True); font_head=self._font_for_image(max(9,int(13*scale)),True); font_small=self._font_for_image(max(7,int((8 if detail=="Stowage Plan" and mode=="Bay" else 10)*scale)),True)
        def box(coords,fill,outline="#D2DCE4",w=1):
            coords=tuple(int(v*scale) for v in coords); draw.rectangle(coords,fill=fill,outline=outline,width=max(1,int(w*scale)))
        def txt(pos,text,fill="#17365D",font=None,anchor="mm",align="center"):
            pos=tuple(int(v*scale) for v in pos); draw.multiline_text(pos,str(text),fill=fill,font=font or font_small,anchor=anchor,align=align,spacing=max(1,int(2*scale)))
        title=f"Ship Cont CAD CV 2.3 — {mode} {selected} — {detail}"
        txt((16,18),title,font=font_title,anchor="la")
        txt((16,43),f"Контейнеров в отображении: {len(shown)}  |  Раскраска: {self.plan_color_by.get()}",fill="#526477",font=font_small,anchor="la")
        if mode=="Tier":
            for j,b in enumerate(bays):
                x=x0+j*cell_w; box((x,y0-cell_h+10,x+cell_w,y0),"#176DA5","#FFFFFF"); txt((x+cell_w/2,y0-cell_h/2+10),f"BAY {b:02d}",fill="white",font=font_head)
            rowmap={r:i for i,r in enumerate(rows)}; baymap={b:i for i,b in enumerate(bays)}
            for i,row in enumerate(rows):
                y=y0+i*cell_h; box((5,y,x0-3,y+cell_h),"#DCEAF7"); txt(((x0+5)/2,y+cell_h/2),f"ROW {row:02d}")
                for j,b in enumerate(bays): box((x0+j*cell_w,y,x0+(j+1)*cell_w,y+cell_h),"#FFFFFF")
            for rec in shown:
                b,r=num(rec.get("BAY")),num(rec.get("ROW"))
                if r not in rowmap: continue
                if b%2==0:
                    pair=[baymap.get(b-1),baymap.get(b+1)]
                    if None in pair: continue
                    x=x0+min(pair)*cell_w+2; w=2*cell_w-4
                else:
                    j=baymap.get(b)
                    if j is None: continue
                    x=x0+j*cell_w+2; w=cell_w-4
                y=y0+rowmap[r]*cell_h+2; fill=self._record_color(rec); label=self._plan_container_label(rec,include_tier=(selected=="Все"))
                box((x,y,x+w,y+cell_h-4),fill,"#8FA9BC"); txt((x+w/2,y+(cell_h-4)/2),label,fill="#18344D",font=font_small)
        else:
            full_width=len(rows)*cell_w; box((x0,84,x0+full_width,108),"#176DA5","#FFFFFF"); txt((x0+full_width/2,96),f"BAY {center+1:02d} — левая сторона",fill="white",font=font_head)
            box((x0+full_width,54,x0+2*full_width,78),"#176DA5","#FFFFFF"); txt((x0+full_width*1.5,66),f"BAY {center-1:02d} — правая сторона",fill="white",font=font_head)
            rowmap={r:i for i,r in enumerate(rows)}; tiermap={t:i for i,t in enumerate(levels)}
            for side in (0,len(rows)):
                for i,row in enumerate(rows):
                    x=x0+(side+i)*cell_w; box((x,110,x+cell_w,y0),"#DCEAF7"); txt((x+cell_w/2,126),f"ROW {row:02d}",font=font_small)
            for i,tier in enumerate(levels):
                y=y0+i*cell_h; box((5,y,x0-4,y+cell_h),"#DCEAF7"); txt(((x0+5)/2,y+cell_h/2),f"{tier:02d}",font=font_head)
                for j in range(len(rows)*2): box((x0+j*cell_w,y,x0+(j+1)*cell_w,y+cell_h),"#FFFFFF","#000000" if mode=="Bay" else "#D2DCE4")
            for rec in shown:
                b,r,t=num(rec.get("BAY")),num(rec.get("ROW")),num(rec.get("TIER"))
                if r not in rowmap or t not in tiermap: continue
                ri,ti=rowmap[r],tiermap[t]; y=y0+ti*cell_h+2; fill=self._record_color(rec); label=self._plan_container_label(rec)
                if b%2==0:
                    for side in (0,1):
                        x=x0+(side*len(rows)+ri)*cell_w+2; box((x,y,x+cell_w-4,y+cell_h-4),fill,"#000000")
                        txt((x+(cell_w-4)/2,y+(cell_h-4)/2),label if side else "X",fill="#18344D" if side else "#C00000",font=font_small)
                else:
                    si=ri if b==center+1 else (len(rows)+ri if b==center-1 else None)
                    if si is not None:
                        x=x0+si*cell_w+2; box((x,y,x+cell_w-4,y+cell_h-4),fill,"#000000"); txt((x+(cell_w-4)/2,y+(cell_h-4)/2),label,fill="#18344D",font=font_small)
            # The Bay export uses the same separators as the on-screen plan.
            sep=x0+len(rows)*cell_w
            draw.line((int(sep*scale),int(54*scale),int(sep*scale),int((y0+len(levels)*cell_h)*scale)),fill="#555F68",width=max(2,int(3*scale)))
            deck=[i for i,t in enumerate(levels) if t>=80]; hold=[i for i,t in enumerate(levels) if t<80]
            if deck and hold:
                boundary_i=max(deck)+1 if max(deck)<min(hold) else max(hold)+1
                by=int((y0+boundary_i*cell_h)*scale)
                draw.line((int(x0*scale),by,int((x0+len(rows)*2*cell_w)*scale),by),fill="#000000",width=max(3,int(5*scale)))
        return image

    def apply_editor_filter(self):
        field_label = self.filter_field.get()
        needle = self.filter_value.get().strip()
        field = "" if field_label == "Все поля" else field_label
        self._active_filter = (field, needle)
        self.refresh_editor()
        self.set_status(f"Фильтр применён: {field_label} = «{needle}»" if needle else "Фильтр очищен")

    def clear_editor_filter(self):
        self.filter_field.set("Все поля")
        self.filter_value.set("")
        self._active_filter = ("", "")
        self.refresh_editor()
        self.set_status("Фильтр редактора сброшен")

    def mass_edit_selected(self):
        selected = []
        for iid in self.editor_tree.selection():
            try:
                idx = int(iid)
                if 0 <= idx < len(self.records): selected.append(idx)
            except (ValueError, TypeError): pass
        selected = list(dict.fromkeys(selected))
        if not selected:
            messagebox.showinfo("Массовое изменение", "Выдели одну или несколько строк в редакторе. Для нескольких строк используй Ctrl или Shift.", parent=self)
            return
        win = tk.Toplevel(self); win.title(f"Массовое изменение — {len(selected)} контейнер(ов)"); win.transient(self); win.grab_set(); win.geometry("760x560")
        frame = ttk.Frame(win, padding=14); frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Выбранные контейнеры будут изменены в одном поле. Перед применением обязательно проверь предварительный просмотр.", wraplength=700).pack(anchor="w", pady=(0,10))
        line = ttk.Frame(frame); line.pack(fill="x", pady=(0,8))
        ttk.Label(line,text="Поле:").pack(side="left")
        field_var = tk.StringVar(value="POD")
        combo = ttk.Combobox(line,textvariable=field_var,state="readonly",values=FIELDS,width=22); combo.pack(side="left",padx=(6,18))
        ttk.Label(line,text="Новое значение:").pack(side="left")
        value_var = tk.StringVar()
        entry = ttk.Entry(line,textvariable=value_var,width=26); entry.pack(side="left",padx=(6,0),fill="x",expand=True)
        cols=("container","old","new")
        preview=ttk.Treeview(frame,columns=cols,show="headings",height=15)
        for col,title,width in [("container","Контейнер",190),("old","Текущее значение",220),("new","Новое значение",220)]:
            preview.heading(col,text=title); preview.column(col,width=width,anchor="w")
        preview.pack(fill="both",expand=True,pady=(4,10))
        btns=ttk.Frame(frame); btns.pack(fill="x")
        state={"previewed":False}
        def make_preview():
            for item in preview.get_children(): preview.delete(item)
            field=field_var.get(); new_value=value_var.get().strip()
            for idx in selected:
                rec=self.records[idx]
                preview.insert("","end",values=(clean(rec.get("CONTAINER_NO")),clean(rec.get(field)),new_value))
            state["previewed"]=True
            apply_btn.state(["!disabled"])
        def apply_changes():
            if not state["previewed"]:
                messagebox.showinfo("Предпросмотр","Сначала нажми «Предпросмотр изменений».",parent=win); return
            field=field_var.get(); raw=value_var.get().strip()
            if field in ("BAY","ROW","TIER","GROSS_KG") and raw:
                try: newval=int(float(raw.replace(",",".")))
                except ValueError:
                    messagebox.showwarning("Проверка значения",f"Поле {field} должно быть числом.",parent=win); return
            else: newval=raw
            if field=="CONTAINER_NO" and len(selected)>1 and raw:
                if not messagebox.askyesno("Внимание", "Будет присвоен одинаковый номер нескольким контейнерам, что создаст ошибки в плане. Продолжить?", parent=win): return
            if not messagebox.askyesno("Подтверждение",f"Применить изменение поля {field} к {len(selected)} контейнерам?",parent=win):return
            for idx in selected:
                data=dict(self.records[idx]); data[field]=newval; self.records[idx]=normalize_record(data)
            self.dirty=True; self._write_active_project_snapshot(); self.refresh_editor(); self.refresh_plan(); self.refresh_stats(); self.update_current_display()
            self.set_status(f"Массовое изменение: {field}, контейнеров: {len(selected)}")
            win.destroy()
        ttk.Button(btns,text="Предпросмотр изменений",command=make_preview).pack(side="left")
        apply_btn=ttk.Button(btns,text="Применить изменения",command=apply_changes); apply_btn.pack(side="right"); apply_btn.state(["disabled"])
        ttk.Button(btns,text="Отмена",command=win.destroy).pack(side="right",padx=(0,8))
        combo.bind("<<ComboboxSelected>>",lambda _e: state.update(previewed=False))
        entry.bind("<KeyRelease>",lambda _e: state.update(previewed=False))

    def _stats_group_field(self):
        return {"По линиям":"LINE", "По TEU / размеру":"__SIZE__", "По Bay":"BAY", "По Tier":"TIER", "По POD":"POD", "По POT":"POT", "По типам контейнеров":"ISO_TYPE"}.get(self.stats_report_type.get())

    def refresh_stats(self):
        if not hasattr(self,"stats_tree"): return
        for item in self.stats_tree.get_children(): self.stats_tree.delete(item)
        records=[normalize_record(r) for r in self.records if clean(r.get("CONTAINER_NO"))]
        total_count=len(records); total_teu=sum(teu_for_record(r) for r in records); total_kg=sum(num(r.get("GROSS_KG"),0) for r in records)
        self.stats_summary.set(f"Контейнеров: {total_count:,}   •   TEU: {total_teu:,}   •   Суммарный вес: {total_kg:,} кг".replace(",", " "))
        group=self._stats_group_field()
        rows=statistics_rows(records,group)
        if self.stats_report_type.get()=="Общая сводка":
            rows=[("Все контейнеры",total_count,total_teu,total_kg,sum(1 for r in records if teu_for_record(r)==1),sum(1 for r in records if teu_for_record(r)==2))]
            rows.extend([
                ("20 ft — 1 TEU",sum(1 for r in records if teu_for_record(r)==1),sum(teu_for_record(r) for r in records if teu_for_record(r)==1),sum(num(r.get("GROSS_KG"),0) for r in records if teu_for_record(r)==1),sum(1 for r in records if teu_for_record(r)==1),0),
                ("40/45 ft — 2 TEU",sum(1 for r in records if teu_for_record(r)==2),sum(teu_for_record(r) for r in records if teu_for_record(r)==2),sum(num(r.get("GROSS_KG"),0) for r in records if teu_for_record(r)==2),0,sum(1 for r in records if teu_for_record(r)==2)),
                ("Тип не определён",sum(1 for r in records if teu_for_record(r)==0),0,sum(num(r.get("GROSS_KG"),0) for r in records if teu_for_record(r)==0),0,0),
            ])
        for label,count,teu,kg,twenty,forty in rows:
            self.stats_tree.insert("","end",values=(label,count,teu,f"{kg:,}".replace(","," "),twenty,forty))

    def export_statistics(self):
        try:
            records=[normalize_record(r) for r in self.records if clean(r.get("CONTAINER_NO"))]
            if not records:
                messagebox.showinfo("Статистика","В текущем документе нет контейнеров для отчёта.",parent=self); return
            path=filedialog.asksaveasfilename(title="Экспорт статистики в Excel",initialdir=str(EXPORTS_DIR),initialfile=f"{self.current_path.stem if self.current_path else 'CVCS'}_Statistics.xlsx",defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
            if not path:return
            wb=Workbook(); summary=wb.active; summary.title="SUMMARY"
            total_count=len(records); total_teu=sum(teu_for_record(r) for r in records); total_kg=sum(num(r.get("GROSS_KG"),0) for r in records)
            summary.append(["Показатель","Значение"])
            summary.append(["Источник",self.current_path.name if self.current_path else "Новый документ"])
            summary.append(["Контейнеров",total_count]); summary.append(["TEU",total_teu]); summary.append(["Суммарный вес, кг",total_kg])
            reports=[("BY_LINE","LINE"),("BY_BAY","BAY"),("BY_TIER","TIER"),("BY_POD","POD"),("BY_POT","POT"),("BY_TYPE","ISO_TYPE")]
            for sheet_name,field in reports:
                ws=wb.create_sheet(sheet_name)
                ws.append(["Группа","Контейнеров","TEU","Вес, кг","20 ft","40/45 ft"])
                for row in statistics_rows(records,field): ws.append(list(row))
                for cell in ws[1]: cell.font=Font(bold=True,color=WHITE); cell.fill=PatternFill("solid",fgColor=DARK_BLUE)
                for col,width in {"A":30,"B":16,"C":12,"D":16,"E":12,"F":14}.items(): ws.column_dimensions[col].width=width
                ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions; ws.sheet_view.showGridLines=False
            for cell in summary[1]: cell.font=Font(bold=True,color=WHITE); cell.fill=PatternFill("solid",fgColor=DARK_BLUE)
            summary.column_dimensions["A"].width=28; summary.column_dimensions["B"].width=34; summary.sheet_view.showGridLines=False
            wb.save(path); wb.close(); self.refresh_projects()
            self.set_status(f"Статистика выгружена: {Path(path).name}")
            messagebox.showinfo("Экспорт статистики",f"Создан отчёт:\n{path}\n\nВключены общая сводка и отчёты по линии, Bay, Tier, POD, POT и типам контейнеров.",parent=self)
        except Exception as e: messagebox.showerror("Ошибка экспорта статистики",str(e),parent=self)

    def _color_value(self, record):
        mode=self.plan_color_by.get()
        field=COLOR_MODE_FIELDS.get(mode,"ISO_TYPE")
        if field=="__VALIDATION__": return container_validation_status(record,self.records)
        return clean(record.get(field)) or "(не указан)"

    def _rebuild_plan_color_map(self, records):
        values=sorted({self._color_value(r) for r in records},key=lambda v:v.casefold())
        self._plan_color_lookup={value:COLOR_PALETTE[i%len(COLOR_PALETTE)] for i,value in enumerate(values)}

    def _record_color(self, record, fallback="#DCEAF7"):
        if not self._plan_color_lookup:
            self._rebuild_plan_color_map(self.records)
        return self._plan_color_lookup.get(self._color_value(record),fallback)

    def _update_plan_legend(self):
        if not hasattr(self,"plan_legend"): return
        for child in self.plan_legend.winfo_children(): child.destroy()
        if not self.records:
            ttk.Label(self.plan_legend,text="Легенда раскраски появится после загрузки контейнеров.",foreground="#526477").pack(side="left")
            return
        ttk.Label(self.plan_legend,text="Легенда:",font=("Segoe UI",9,"bold")).pack(side="left",padx=(0,6))
        items=list(self._plan_color_lookup.items())
        for value,color in items[:8]:
            chip=tk.Label(self.plan_legend,text="  ",background=color,relief="solid",borderwidth=1,width=2)
            chip.pack(side="left",padx=(4,2))
            ttk.Label(self.plan_legend,text=value[:24]).pack(side="left")
        if len(items)>8: ttk.Label(self.plan_legend,text=f"+ ещё {len(items)-8} значений",foreground="#526477").pack(side="left",padx=8)

    def _create_backup(self, path):
        path=Path(path)
        if not path.exists() or path.suffix.lower() not in (".edi",".txt",".baplie",".xlsx",".xlsm"):
            return None
        ensure_folders()
        stamp=datetime.now().strftime("%Y%m%d_%H%M%S")
        dest=BACKUPS_DIR/f"{stamp}__{path.stem}{path.suffix}"
        suffix=2
        while dest.exists():
            dest=BACKUPS_DIR/f"{stamp}_{suffix}__{path.stem}{path.suffix}"; suffix+=1
        shutil.copy2(path,dest)
        # Keep the last 25 backups per source filename to control disk use.
        related=sorted(BACKUPS_DIR.glob(f"*__{path.stem}{path.suffix}"),key=lambda q:q.stat().st_mtime,reverse=True)
        for old in related[25:]:
            try: old.unlink()
            except OSError: pass
        return dest

    def open_exports_folder(self):
        ensure_folders()
        try:
            if os.name=="nt": os.startfile(str(EXPORTS_DIR))
            elif sys.platform=="darwin": os.system(f'open "{EXPORTS_DIR}"')
            else: os.system(f'xdg-open "{EXPORTS_DIR}"')
        except Exception as e: messagebox.showerror("Папка экспорта",f"Путь:\n{EXPORTS_DIR}\n\n{e}",parent=self)

    def open_backups_folder(self):
        ensure_folders()
        try:
            if os.name=="nt": os.startfile(str(BACKUPS_DIR))
            elif sys.platform=="darwin": os.system(f'open "{BACKUPS_DIR}"')
            else: os.system(f'xdg-open "{BACKUPS_DIR}"')
        except Exception as e: messagebox.showerror("Резервные копии",f"Путь:\n{BACKUPS_DIR}\n\n{e}",parent=self)

    def restore_backup_dialog(self):
        ensure_folders()
        backups=sorted([p for p in BACKUPS_DIR.iterdir() if p.is_file() and p.suffix.lower() in (".edi",".txt",".baplie",".xlsx",".xlsm")],key=lambda p:p.stat().st_mtime,reverse=True)
        if not backups:
            messagebox.showinfo("Резервные копии","Пока нет резервных копий. Они появятся после сохранения изменений в существующий файл.",parent=self); return
        win=tk.Toplevel(self); win.title("Восстановление резервной копии"); win.transient(self); win.grab_set(); win.geometry("760x440")
        frame=ttk.Frame(win,padding=14); frame.pack(fill="both",expand=True)
        ttk.Label(frame,text="Выбери резервную копию. Программа откроет её как отдельный рабочий файл и не перезапишет текущий проект.",wraplength=720).pack(anchor="w",pady=(0,8))
        tree=ttk.Treeview(frame,columns=("name","date","size"),show="headings",selectmode="browse")
        for col,title,width in [("name","Резервная копия",420),("date","Создана",160),("size","Размер",90)]: tree.heading(col,text=title); tree.column(col,width=width,anchor="w")
        mapping={}
        for path in backups:
            mapping[path.name]=path
            tree.insert("","end",iid=path.name,values=(path.name,datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S"),f"{path.stat().st_size/1024:.1f} KB"))
        tree.pack(fill="both",expand=True)
        def restore():
            sel=tree.selection()
            if not sel:return
            backup=mapping[sel[0]]
            win.destroy(); self.load_document(backup)
        buttons=ttk.Frame(frame); buttons.pack(fill="x",pady=(10,0))
        ttk.Button(buttons,text="Восстановить выбранную",command=restore).pack(side="right")
        ttk.Button(buttons,text="Отмена",command=win.destroy).pack(side="right",padx=(0,8))
        tree.bind("<Double-1>",lambda _e:restore())

    def write_recovery_snapshot(self):
        ensure_folders()
        try:
            if not self.dirty:
                if RECOVERY_PATH.exists(): RECOVERY_PATH.unlink()
                return
            payload={"app_version":"1.06","saved_at":datetime.now().isoformat(timespec="seconds"),"current_path":str(self.current_path) if self.current_path else "","raw_edi":self.raw_edi,"vessel":self.vessel,"records":[normalize_record(r) for r in self.records],"source_records":[normalize_record(r) for r in self.source_records],"source_label":self.source_label,"dirty":True}
            tmp=RECOVERY_PATH.with_suffix(".tmp")
            tmp.write_text(json.dumps(payload,ensure_ascii=False,indent=2,default=str),encoding="utf-8")
            tmp.replace(RECOVERY_PATH)
        except Exception as e:
            self.set_status(f"Не удалось создать автосохранение: {e}")

    def autosave_tick(self):
        try:
            if self.dirty:
                self.write_recovery_snapshot()
            if self.active_project_id:
                if self._write_active_project_snapshot():
                    self.set_status(f"Проект «{self.active_project_name}» автоматически сохранён: {datetime.now().strftime('%H:%M:%S')}")
            elif self.dirty:
                self.set_status(f"Автосохранение состояния выполнено: {datetime.now().strftime('%H:%M:%S')}")
        finally:
            try: self.after(10000,self.autosave_tick)
            except tk.TclError: pass

    def offer_recovery(self):
        if not RECOVERY_PATH.exists(): return
        try:
            payload=json.loads(RECOVERY_PATH.read_text(encoding="utf-8"))
            saved_at=clean(payload.get("saved_at")) or "время неизвестно"
            if not messagebox.askyesno("Найдено автосохранение",f"Обнаружено состояние с несохранёнными изменениями от {saved_at}. Восстановить его?",parent=self):
                RECOVERY_PATH.unlink(missing_ok=True); return
            self.vessel=dict(DEFAULT_VESSEL); self.vessel.update(payload.get("vessel",{}))
            self.records=[normalize_record(r) for r in payload.get("records",[])]
            self.source_records=[normalize_record(r) for r in payload.get("source_records", payload.get("records",[]))]
            self.source_vessel=dict(self.vessel); self.source_label=clean(payload.get("source_label")) or "Восстановленное автосохранение"
            self.raw_edi=payload.get("raw_edi","")
            old_path=clean(payload.get("current_path","")); self.current_path=Path(old_path) if old_path and Path(old_path).exists() else None
            self.dirty=True; self._write_active_project_snapshot(); self.refresh_editor(); self.refresh_plan(); self.refresh_stats(); self.update_current_display()
            self.set_status(f"Восстановлено автосохранение от {saved_at}"); self.nb.select(self.tab_editor)
        except Exception as e:
            try: RECOVERY_PATH.rename(RECOVERY_PATH.with_name("autosave_recovery_broken.json"))
            except Exception: pass
            messagebox.showwarning("Автосохранение",f"Не удалось восстановить автоматическое сохранение:\n{e}",parent=self)

    def get_template(self):
        template=find_template_path(self.config_data)
        if not template:
            raise FileNotFoundError("Не найден встроенный или выбранный Excel-шаблон. Выбери образец на вкладке «Загрузки».")
        wb=load_workbook(template,read_only=True,data_only=True)
        ok="VESSEL" in wb.sheetnames and "CONTAINERS" in wb.sheetnames
        wb.close()
        if not ok: raise ValueError("Excel-образец должен содержать листы VESSEL и CONTAINERS.")
        return Path(template)

    def save_current(self):
        if not self.current_path:
            return self.save_as()
        p=self.current_path
        master=clean(self.config_data.get("template_path",""))
        if master and Path(master).exists():
            try:
                if p.resolve()==Path(master).resolve():
                    return self.save_as()
            except Exception:pass
        try:
            self._create_backup(p)
            if p.suffix.lower() in (".edi",".txt",".baplie"):
                text=create_edi_text(self.vessel,self.records); p.write_text(text,encoding="utf-8")
                self.raw_edi=text
            elif p.suffix.lower() in (".xlsx",".xlsm"):
                template=self.get_template()
                write_input_workbook(p,self.vessel,self.records,template)
            else:
                raise ValueError("Формат сохранения должен быть .edi или .xlsx.")
            self.dirty=False; self.write_recovery_snapshot(); self._write_active_project_snapshot(); self.set_status(f"Сохранено: {p.name}"); self.refresh_projects(); self.refresh_editor(); self.refresh_plan(); self.refresh_stats()
        except Exception as e: messagebox.showerror("Ошибка сохранения",str(e),parent=self)

    def save_as(self):
        ensure_folders()
        initial=(self.current_path.stem if self.current_path else "CVCS_Project")
        p=filedialog.asksaveasfilename(title="Сохранить рабочий файл как",initialdir=str(PROJECTS_DIR),initialfile=initial,
              defaultextension=".xlsx",filetypes=[("Рабочая таблица Excel","*.xlsx"),("EDI","*.edi"),("Все файлы","*.*")])
        if not p:return
        path=Path(p)
        try:
            self._create_backup(path)
            if path.suffix.lower()==".edi":
                text=create_edi_text(self.vessel,self.records); path.write_text(text,encoding="utf-8"); self.raw_edi=text
            elif path.suffix.lower() in (".xlsx",".xlsm"):
                template=self.get_template(); write_input_workbook(path,self.vessel,self.records,template)
            else:
                messagebox.showwarning("Формат файла","Выбери расширение .edi или .xlsx.",parent=self);return
            self.current_path=path; self.dirty=False; self.write_recovery_snapshot(); self._write_active_project_snapshot(); self.update_current_display(); self.refresh_projects(); self.refresh_plan(); self.refresh_stats(); self.set_status(f"Рабочий файл сохранён: {path.name}")
        except Exception as e: messagebox.showerror("Ошибка сохранения",str(e),parent=self)

    def _current_editor_snapshot(self):
        """Take a consistent export snapshot from the active editor document."""
        snapshot = [normalize_record(r) for r in self.records]
        if hasattr(self, "editor_tree"):
            edited = False
            synchronized = [normalize_record(r) for r in self.records]
            # Only visible rows are represented by the Treeview when a filter is active;
            # update those original record indices and keep hidden rows intact.
            for item_id in self.editor_tree.get_children(""):
                try: index = int(item_id)
                except (ValueError, TypeError): continue
                if index < 0 or index >= len(synchronized): continue
                values = self.editor_tree.item(item_id, "values")
                visible = {field: values[i] if i < len(values) else "" for i, field in enumerate(FIELDS)}
                current = dict(synchronized[index])
                for field in FIELDS:
                    if clean(visible.get(field, "")) != clean(current.get(field, "")):
                        current[field] = visible.get(field, "")
                        edited = True
                synchronized[index] = normalize_record(current)
            if edited:
                self.records = synchronized
                snapshot = synchronized
                self.dirty = True
                self._write_active_project_snapshot()
                self.update_current_display()
                self.refresh_plan()
                self.refresh_stats()
        return dict(self.vessel), [normalize_record(r) for r in snapshot]

    def export_edi(self):
        try:
            vessel, all_records = self._current_editor_snapshot()
            if not any(clean(r.get("CONTAINER_NO")) for r in all_records):
                messagebox.showwarning("Нет данных для экспорта", "В текущем документе редактора нет контейнеров с заполненным номером. Сначала откройте EDI/Excel или добавьте контейнеры в редактор.", parent=self)
                return
            if not self.validate_before_edi_export(all_records):
                return
            records = [normalize_record(r) for r in all_records if clean(r.get("CONTAINER_NO"))]
            stem = self.current_path.stem if self.current_path else "BAPLIE"
            initial = f"{stem}_export.edi" if self.current_path and self.current_path.suffix.lower() in (".edi", ".txt", ".baplie") else f"{stem}.edi"
            p=filedialog.asksaveasfilename(title="Экспорт EDI из текущего редактора",initialdir=str(PROJECTS_DIR),initialfile=initial,defaultextension=".edi",filetypes=[("EDI","*.edi"),("Text","*.txt")])
            if not p:return
            text=create_edi_text(vessel,records)
            Path(p).write_text(text,encoding="utf-8")
            self.raw_edi=text
            self.set_status(f"EDI выгружен из текущего редактора: {Path(p).name}; контейнеров: {len(records)}")
            self.refresh_projects()
            messagebox.showinfo("Экспорт завершён", f"Создан EDI из текущего содержимого редактора:\n{p}\nКонтейнеров выгружено: {len(records)}\n\nПеред отправкой проверь EDI по требованиям принимающего терминала/линии.", parent=self)
        except Exception as e:messagebox.showerror("Ошибка экспорта EDI",str(e),parent=self)

    def export_graph_excel(self):
        try:
            vessel, records = self._current_editor_snapshot()
            records = [r for r in records if clean(r.get("CONTAINER_NO"))]
            if not records:
                messagebox.showwarning("Нет данных для экспорта", "В текущем документе редактора нет контейнеров с заполненным номером. Сначала откройте EDI/Excel или добавьте контейнеры в редактор.", parent=self)
                return
            if not self.validate_before_edi_export(records):
                return
            mode=self.mode.get()
            stem=self.current_path.stem if self.current_path else "BAPLIE"
            p=filedialog.asksaveasfilename(title="Экспорт Excel из текущего редактора",initialdir=str(PROJECTS_DIR),initialfile=f"{stem}_{mode}.xlsx",defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
            if not p:return
            raw=create_edi_text(vessel,records)
            wb=make_export_workbook(Path(p).name,records,mode,vessel,raw)
            wb.save(p);wb.close()
            self.set_status(f"Excel выгружен из текущего редактора: режим {mode}; контейнеров: {len(records)}")
            self.refresh_projects()
            messagebox.showinfo("Экспорт завершён", f"Создан Excel из текущего содержимого редактора:\n{p}\nРежим: {mode}\nКонтейнеров: {len(records)}", parent=self)
        except Exception as e:messagebox.showerror("Ошибка экспорта Excel",str(e)+"\n\n"+traceback.format_exc(),parent=self)

    def export_blank_template(self):
        try:
            template=self.get_template()
            default_name="EDI_Input_Template.xlsx"
            p=filedialog.asksaveasfilename(title="Сохранить пустой Excel-шаблон",initialdir=str(PROJECTS_DIR),initialfile=default_name,defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
            if not p:return
            wb=load_workbook(template)
            cs=wb["CONTAINERS"]
            max_col=cs.max_column
            for r in range(2,cs.max_row+1):
                for c in range(1,max_col+1): cs.cell(r,c).value=None
            # Blank user-provided vessel data while keeping the field labels and explanations.
            if "VESSEL" in wb.sheetnames:
                vs=wb["VESSEL"]
                for row in vs.iter_rows(min_row=2):
                    key=clean(row[0].value) if row else ""
                    if len(row)>1 and key in DEFAULT_VESSEL:
                        row[1].value = "TBN" if key=="VESSEL_NAME" else ("D:95B:UN:SMDG20" if key=="MESSAGE_VERSION" else ("Масса в кг; температура °C" if key=="CURRENCY/UNITS_NOTE" else ""))
            if "README" not in wb.sheetnames:
                ins=wb.create_sheet("README",0)
                ins.append(["ИНСТРУКЦИЯ К ШАБЛОНУ EDI"])
                ins.append(["1. Заполните данные судна на листе VESSEL."])
                ins.append(["2. На листе CONTAINERS добавляйте по одному контейнеру в каждую строку, начиная со строки 2."])
                ins.append(["3. Не изменяйте названия обязательных столбцов: CONTAINER_NO, ISO_TYPE, GROSS_KG, BAY, ROW, TIER."])
                ins.append(["4. GROSS_KG указывается в килограммах; BAY/ROW/TIER должны соответствовать плану судна."])
                ins.append(["5. POL/POD указывайте согласованными кодами портов. Неизвестные поля оставляйте пустыми."])
                ins.append(["6. POT — порт перевалки, если он есть в данных рейса; не путайте с POL и POD."])
            else:
                ws=wb["README"]
                ws["B5"]="В листе CONTAINERS демонстрационных записей нет. Добавляйте реальные контейнеры начиная со строки 2."
                # Replace any legacy note that tells users to remove sample records.
                for row in ws.iter_rows():
                    for cell in row:
                        if isinstance(cell.value,str) and "Удалите три демонстрационные строки" in cell.value:
                            cell.value="В листе CONTAINERS нет демонстрационных записей. Добавляйте реальные контейнеры начиная со строки 2."
            for ws in wb.worksheets:
                ws.sheet_view.showGridLines=False
            wb.save(p);wb.close(); self.refresh_projects(); self.set_status(f"Пустой шаблон сохранён: {Path(p).name}")
            messagebox.showinfo("Шаблон готов",f"Создан пустой Excel-шаблон:\n{p}\n\nСтруктура и пояснения сохранены, контейнерные строки очищены.",parent=self)
        except Exception as e: messagebox.showerror("Ошибка шаблона",str(e),parent=self)

    def close_app(self):
        if self.active_project_id:
            if not self._write_active_project_snapshot():
                if not messagebox.askyesno("Не удалось сохранить проект", "Состояние проекта не удалось сохранить. Всё равно закрыть программу?", parent=self):
                    return
            # Project state is persisted independently from the source file.
            self.dirty=False
        if self.dirty:
            answer=messagebox.askyesnocancel("Несохранённые изменения", "Сохранить изменения перед закрытием? Если выбрать «Нет», состояние останется в автосохранении для восстановления при следующем запуске.", parent=self)
            if answer is None: return
            if answer:
                self.save_current()
                if self.dirty: return
            else:
                self.write_recovery_snapshot()
        else:
            self.write_recovery_snapshot()
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
