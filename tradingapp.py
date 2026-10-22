from datetime import date, datetime
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import streamlit as st
import re
from html import escape
from uuid import uuid4
from zoneinfo import ZoneInfo

# ==========================================
# 1. ตั้งค่าหน้าเว็บ Streamlit (Wide Mode & CSS)
# ==========================================
st.set_page_config(
    page_title="Thai Stick Trading Portal", page_icon="🌱", layout="wide"
)

st.markdown(
    """
    <style>
    .main { background-color: #f7f9fc; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .stMetric { background-color: #ffffff; padding: 18px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.02); border: 1px solid #e2e8f0; }
    .stMetric label { color: #4a5568 !important; font-weight: 600; }
    .stButton>button { width: 100%; background-color: #68d391; color: #1a202c; border-radius: 8px; font-weight: bold; border: none; padding: 0.5rem 1rem; box-shadow: 0 2px 4px rgba(0,0,0,0.05); transition: all 0.2s ease; }
    .stButton>button:hover { background-color: #48bb78; color: white; }
    .stFormSubmitButton>button { width: 100%; background-color: #4fd1c5; color: #1a202c; border-radius: 8px; font-weight: bold; border: none; }
    .stFormSubmitButton>button:hover { background-color: #319795; color: white; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { background-color: #edf2f7; border-radius: 8px 8px 0px 0px; padding: 10px 20px; color: #4a5568; font-weight: 600; }
    .stTabs [aria-selected="true"] { background-color: #4fd1c5 !important; color: #1a202c !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# 2. ฟังก์ชัน Cache ที่ปลอดภัย (แยกจำแต่ละ Sheet เด็ดขาด)
# ==========================================
@st.cache_data(ttl=30, show_spinner=False)
def get_cached_data(sheet_name, _sheet, include_row_index=False):
    try:
        rows = _sheet.get_all_values()
        if not rows or len(rows) <= 1:
            return []
        
        headers = [str(h).strip() for h in rows[0]]
        data = []
        
        for sheet_row, r in enumerate(rows[1:], start=2):
            # กรองบรรทัดที่ว่างเปล่าทิ้ง ป้องกันตารางบวม
            if not any(str(cell).strip() for cell in r):
                continue
                
            while len(r) < len(headers):
                r.append("")
            
            row_dict = {headers[i]: r[i] for i in range(len(headers))}
            if include_row_index:
                row_dict["Original_Row_Index"] = sheet_row
            data.append(row_dict)
            
        return data
    except Exception as e:
        return []

# ==========================================
# 3. ฟังก์ชันเชื่อมต่อ Google Sheets (โหลดครั้งเดียว ประหยัดโควตา API)
# ==========================================
SHEET_NAME = "Trading Data Log"

@st.cache_resource
def init_google_sheets(sheet_name):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    sh = client.open(sheet_name)
    
    main_sheet = sh.get_worksheet(0)
    
    def get_or_create_worksheet(title, rows, cols, headers):
        try:
            ws = sh.worksheet(title)
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title=title, rows=rows, cols=cols)
            ws.append_row(headers, table_range="A1")
        return ws

    # แทรก Supplier Name เข้าไปเป็นตำแหน่งที่ 3 (Index 2)
    packing_records_sheet = get_or_create_worksheet(
        "Packing_Records", "1000", "15", 
        ["Timestamp", "Source WIP Lot", "Supplier Name", "Cultivar Name", "Export Name", "Generated FG Lot No.", "Grade", "Quantity", "Waste Qty", "Packing Date", "Note"]
    )
    packing_log_sheet = get_or_create_worksheet(
        "Packing_Log", "1000", "10", 
        ["Timestamp", "Source WIP Lot", "Supplier Name", "Cultivar", "Packed FG Name", "Grade & Qty Breakdown", "Total Packed Qty", "Waste Qty", "Packing Date", "Note"]
    )
    shipped_sheet = get_or_create_worksheet(
        "Shipped_Out_Log", "1000", "12", 
        ["Timestamp", "Category", "Lot No.", "Cultivar / Product", "Grade", "Shipped Qty", "Consignee (ปลายทาง)", "Client (ลูกค้า)", "Shipping Reason", "Shipment Date", "Note"]
    )

    EXPECTED_HEADERS = [
        "Cultivar Name", "Supplier Name", "Supplier Lot. No.", "Export Name", "THAI STICK Lot. No.", 
        "Category", "Grade", "Receive date", "Quantity", "Remark Qty", "Inspection Req.", "IR Req.", 
        "Shipment date", "Type", "Status", "Packing date"
    ]
    current_headers = main_sheet.row_values(1)
    if not current_headers or len(current_headers) < len(EXPECTED_HEADERS):
        main_sheet.update(range_name="A1:P1", values=[EXPECTED_HEADERS])
        
    return client, sh, main_sheet, packing_records_sheet, packing_log_sheet, shipped_sheet

client, sh, sheet, packing_records_sheet, packing_sheet, shipped_sheet = init_google_sheets(SHEET_NAME)

# ==========================================
# 4. ฟังก์ชันตัวช่วย (Helper Functions)
# ==========================================
def clean_quantity(val):
    try: return float(str(val).replace(",", "").strip())
    except: return 0.0

DELIVERY_PENDING_VALUES = {"tbc", "to be confirmed", "tbd", "to be decided", "", "nan", "nat", "none"}


def parse_delivery_date(value):
    text = str(value).strip()
    if text.casefold() in DELIVERY_PENDING_VALUES:
        return pd.NaT
    if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:\D|$)", text):
        parsed = pd.to_datetime(text, yearfirst=True, errors="coerce")
    else:
        parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    return pd.Timestamp(parsed.date()) if pd.notna(parsed) else pd.NaT


def format_df_quantity(df):
    if df is not None and not df.empty:
        df_copy = df.copy()
        for col in df_copy.columns:
            col_str = str(col).lower()
            
            # ✨ 1. จัดการ Format วันที่แบบขั้นสูง (ดักจับและแก้ปัญหาวัน/เดือนสลับกัน)
            date_keywords = ["date", "วันที่", "timestamp"]
            if any(k in col_str for k in date_keywords):
                def format_date(x):
                    if pd.isna(x) or str(x).strip().casefold() in DELIVERY_PENDING_VALUES:
                        return x
                    try:
                        d_str = str(x).strip().split(" ")[0] # ตัดเวลาทิ้งก่อนเอามาวิเคราะห์
                        
                        import re
                        # กรณี 1: ขึ้นต้นด้วยปี (เช่น 2026-10-24 หรือ 2026/10/24)
                        if re.match(r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}$', d_str):
                            dt = pd.to_datetime(d_str)
                        
                        # กรณี 2: ขึ้นต้นด้วยวันหรือเดือน (เช่น 24/10/2026 หรือ 10/24/2026)
                        elif re.match(r'^\d{1,2}[-/]\d{1,2}[-/]\d{4}$', d_str):
                            parts = re.split(r'[-/]', d_str)
                            p1, p2, p3 = int(parts[0]), int(parts[1]), int(parts[2])
                            
                            if p1 > 12: 
                                # ตัวเลขแรกเกิน 12 แปลว่าวันขึ้นก่อนแน่นอน (DD/MM/YYYY)
                                dt = pd.to_datetime(f"{p3}-{p2:02d}-{p1:02d}")
                            elif p2 > 12: 
                                # ตัวเลขที่สองเกิน 12 แปลว่าเดือนขึ้นก่อนแน่นอน (MM/DD/YYYY)
                                dt = pd.to_datetime(f"{p3}-{p1:02d}-{p2:02d}")
                            else:
                                # ถ้าคลุมเครือ (เช่น 05/06/2026) ให้ยึด DD/MM/YYYY เป็นมาตรฐานไทย
                                dt = pd.to_datetime(f"{p3}-{p2:02d}-{p1:02d}")
                        else:
                            # กรณีอื่นๆ ที่รูปแบบแปลกไป ให้ Pandas ลองเดา
                            dt = pd.to_datetime(d_str, dayfirst=True, errors='coerce')
                        
                        if pd.notna(dt):
                            return dt.strftime('%d/%m/%Y')
                        return x
                    except:
                        return x
                df_copy[col] = df_copy[col].apply(format_date)

            # 📦 2. จัดการลูกน้ำ (,) สำหรับคอลัมน์ตัวเลขและจำนวน
            keywords = ["qty", "quantity", "amount", "total", "price", "ยอด", "จำนวน", "ราคา", "น้ำหนัก", "weight", "มูลค่า", "volume", "ปริมาณ", "available", "wip", "fg"]
            is_target_col = any(k in col_str for k in keywords)
            ignore_kws = ["year", "ปี", "lot", "id", "เลข", "no.", "รหัส"]
            is_ignore = any(ig in col_str for ig in ignore_kws)
            
            if is_target_col and not is_ignore:
                def format_mixed(x):
                    if pd.isna(x) or str(x).strip() == "": return x
                    clean_str = str(x).replace(",", "")
                    try:
                        val = float(clean_str)
                        return f"{val:,.2f}" if val % 1 != 0 else f"{int(val):,}"
                    except ValueError:
                        def repl(match):
                            try:
                                num_val = float(match.group(0))
                                return f"{num_val:,.2f}" if num_val % 1 != 0 else f"{int(num_val):,}"
                            except: return match.group(0)
                        return re.sub(r'\b\d+(?:\.\d+)?\b', repl, clean_str)
                df_copy[col] = df_copy[col].apply(format_mixed)
                
        return df_copy
    return df

import math

PICKING_BASE_HEADERS = [
    "Request ID", "Line ID", "Created At", "Cultivar Name", "Export Name",
    "Source Categories", "Grade", "Volume", "Consignee", "True Consignee",
    "Expected Shipping Date", "Required Category", "Note",
]
PICKING_HEADERS = PICKING_BASE_HEADERS + ["Project Owner", "Type", "Supplier Name"]
PICKING_TYPES = ["Inspection & Packing Service", "Trading Goods (TS)", "Consignment", "In House Product", "To be decided"]


def picking_text(value):
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.casefold() in {"nan", "nat", "none", "<na>"} else text


def picking_key(value):
    return " ".join(picking_text(value).split()).casefold()


def picking_quantity(value):
    try:
        number = float(picking_text(value).replace(",", ""))
    except (ValueError, TypeError):
        return None
    return number if math.isfinite(number) else None


def build_picking_catalog(main_records, packing_records):
    catalog = {}

    def add(record, category, packed=False):
        quantity = picking_quantity(record.get("Quantity", ""))
        if quantity is None or quantity <= 0:
            return
        if picking_key(record.get("Status")) == "delivered":
            return
        cultivar = picking_text(record.get("Cultivar Name"))
        export_name = picking_text(record.get("Export Name"))
        if not cultivar:
            if not packed:
                return
            cultivar = export_name
        if not cultivar:
            return
        item = catalog.setdefault(picking_key(cultivar), {
            "cultivar": cultivar, "export_names": [], "sources": [],
            "wip_qty": 0.0, "fg_by_grade": {}, "suppliers": [],
        })
        if export_name and export_name not in item["export_names"]:
            item["export_names"].append(export_name)
        if category not in item["sources"]:
            item["sources"].append(category)
        supplier = picking_text(record.get("Supplier Name"))
        if supplier and supplier not in item["suppliers"]:
            item["suppliers"].append(supplier)
        grade = picking_text(record.get("Grade")).upper() or "GENERAL"
        if category == "FG":
            item["fg_by_grade"][grade] = item["fg_by_grade"].get(grade, 0.0) + quantity
        else:
            item["wip_qty"] += quantity

    for record in main_records:
        category = picking_text(record.get("Category")).upper()
        if "WIP" in category:
            add(record, "WIP")
        elif "FG" in category:
            add(record, "FG")
    for record in packing_records:
        add(record, "FG", packed=True)
    return catalog


def prepare_picking_lines(records, catalog, default_consignee="", default_true_consignee="", default_project_owner="", default_type=""):
    lines = []
    for index, record in enumerate(records, start=1):
        cultivar = picking_text(record.get("Cultivar Name"))
        grade = picking_text(record.get("Grade")).upper()
        raw_quantity = record.get("Volume", "")
        consignee = picking_text(record.get("Consignee")) or picking_text(default_consignee)
        true_consignee = picking_text(record.get("True Consignee")) or picking_text(default_true_consignee)
        project_owner = picking_text(record.get("Project Owner")) or picking_text(default_project_owner)
        product_type = picking_text(record.get("Type")) or picking_text(default_type)
        # Ignore only completely empty editor rows.
        if not any(picking_text(record.get(name)) for name in
                   ["Cultivar Name", "Grade", "Volume", "Consignee", "True Consignee", "Project Owner", "Type", "Note"]):
            continue
        item = catalog.get(picking_key(cultivar))
        quantity = picking_quantity(raw_quantity)
        if item is None:
            raise ValueError(f"รายการที่ {index}: กรุณาเลือกสายพันธุ์จากรายการ WIP / FG")
        if not grade:
            raise ValueError(f"รายการที่ {index}: กรุณาระบุเกรด FG ที่ต้องจัดส่ง")
        if quantity is None or quantity <= 0:
            raise ValueError(f"รายการที่ {index}: จำนวนต้องเป็นตัวเลขมากกว่า 0")
        if not consignee or not true_consignee:
            raise ValueError(f"รายการที่ {index}: กรุณาระบุ Consignee และ True Consignee")
        if not project_owner or not product_type:
            raise ValueError(f"รายการที่ {index}: กรุณาระบุ Project Owner และ Type ผลิตภัณฑ์")
        lines.append({
            "Cultivar Name": item["cultivar"],
            "Export Name": " / ".join(item["export_names"]),
            "Source Categories": " / ".join(item["sources"]),
            "Grade": grade,
            "Volume": quantity,
            "Consignee": consignee,
            "True Consignee": true_consignee,
            "Project Owner": project_owner, "Type": product_type,
            "Supplier Name": " / ".join(item.get("suppliers", [])),
            "Required Category": "FG (Finished Goods)",
            "Note": picking_text(record.get("Note")),
        })
    if not lines:
        raise ValueError("กรุณาเพิ่มสินค้าอย่างน้อย 1 รายการ")
    return lines


def project_picking_rows(rows, catalog):
    # Aggregate demand inside one request only; different requests do not reserve stock.
    demand = {}
    for row in rows:
        quantity = picking_quantity(row.get("Volume"))
        if quantity is None or quantity <= 0:
            continue
        key = (
            picking_text(row.get("Request ID")) or "DRAFT",
            picking_key(row.get("Cultivar Name")),
            picking_text(row.get("Grade")).upper(),
        )
        demand[key] = demand.get(key, 0.0) + quantity

    projected = []
    for row in rows:
        result = dict(row)
        quantity = picking_quantity(row.get("Volume"))
        grade = picking_text(row.get("Grade")).upper()
        key = (
            picking_text(row.get("Request ID")) or "DRAFT",
            picking_key(row.get("Cultivar Name")),
            grade,
        )
        item = catalog.get(key[1], {})
        fg_quantity = item.get("fg_by_grade", {}).get(grade, 0.0)
        wip_quantity = item.get("wip_qty", 0.0)
        request_demand = demand.get(key, 0.0)
        if quantity is None or quantity <= 0 or not grade:
            status = "⚠️ กรุณาตรวจสอบข้อมูลคำขอ"
        elif fg_quantity >= request_demand:
            status = "✅ พร้อมส่ง"
        elif wip_quantity > 0:
            status = "📦 รอแพ็ค"
        else:
            status = "📥 รอรับเข้าสินค้า"
        result.update({
            "Required Category": "FG (Finished Goods)",
            "Available FG": fg_quantity,
            "Available WIP": wip_quantity,
            "Request Demand": request_demand,
            "FG Shortfall": max(0.0, request_demand - fg_quantity),
            "Status (ความพร้อม)": status,
        })
        projected.append(result)
    return projected


PICKING_SHEET_NAME = "Picking_Requests"


def read_picking_records(worksheet):
    values = worksheet.get_all_values()
    if not values:
        return []
    headers = [str(value).strip() for value in values[0]]
    missing = [header for header in PICKING_BASE_HEADERS if header not in headers]
    if missing:
        raise ValueError("หัวตาราง Picking_Requests ไม่ครบ: " + ", ".join(missing))
    if len(headers) != len(set(headers)):
        raise ValueError("พบชื่อคอลัมน์ซ้ำใน Picking_Requests กรุณาตรวจสอบหัวตาราง")
    records = []
    for row_number, values_row in enumerate(values[1:], start=2):
        if not any(picking_text(value) for value in values_row):
            continue
        record = {header: values_row[index] if index < len(values_row) else ""
                  for index, header in enumerate(headers)}
        if not any(picking_text(record.get(header)) for header in PICKING_HEADERS):
            continue
        for header in PICKING_HEADERS:
            record.setdefault(header, "")
        record["Original_Row_Index"] = row_number
        records.append(record)
    return records


@st.cache_data(ttl=30, show_spinner=False)
def get_cached_picking_rows(_worksheet):
    return read_picking_records(_worksheet)


def load_picking_requests(spreadsheet):
    try:
        worksheet = spreadsheet.worksheet(PICKING_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        return None, []
    return worksheet, get_cached_picking_rows(worksheet)


def picking_same_payload(existing, expected):
    if len(existing) != len(expected):
        return False
    by_id = {picking_text(row.get("Line ID")): row for row in existing}
    if len(by_id) != len(existing):
        return False
    for wanted in expected:
        actual = by_id.get(wanted["Line ID"])
        if actual is None:
            return False
        for header in PICKING_HEADERS:
            if header == "Volume":
                actual_quantity = picking_quantity(actual.get(header))
                if actual_quantity is None or picking_quantity(wanted.get(header)) is None or not math.isclose(actual_quantity, picking_quantity(wanted[header]), rel_tol=1e-12, abs_tol=1e-9):
                    return False
            elif picking_text(actual.get(header)) != picking_text(wanted.get(header)):
                return False
    return True


def save_picking_request(spreadsheet, request_id, created_at, lines, shipping_date):
    request_id = picking_text(request_id)
    created_at = picking_text(created_at)
    shipping_date = picking_text(shipping_date)
    if not request_id or not created_at or not lines:
        raise ValueError("ข้อมูลคำขอไม่ครบ กรุณาตรวจสอบก่อนบันทึก")
    if shipping_date.upper() != "TBC":
        try:
            shipping_date = date.fromisoformat(shipping_date).isoformat()
        except ValueError:
            raise ValueError("กรุณาระบุวันจัดส่งที่ถูกต้อง หรือเลือก TBC") from None
    else:
        shipping_date = "TBC"
    payload = []
    for index, line in enumerate(lines, start=1):
        quantity = picking_quantity(line.get("Volume"))
        if quantity is None or quantity <= 0:
            raise ValueError(f"รายการที่ {index}: จำนวนต้องเป็นตัวเลขมากกว่า 0")
        for required in ["Cultivar Name", "Grade", "Consignee", "True Consignee", "Project Owner", "Type"]:
            if not picking_text(line.get(required)):
                raise ValueError(f"รายการที่ {index}: กรุณาระบุ {required}")
        record = {header: picking_text(line.get(header)) for header in PICKING_HEADERS}
        record.update({
            "Request ID": request_id, "Line ID": f"{request_id}-{index:03d}",
            "Created At": created_at, "Grade": picking_text(line.get("Grade")).upper(),
            "Volume": quantity, "Expected Shipping Date": shipping_date,
            "Required Category": "FG (Finished Goods)",
        })
        payload.append(record)
    try:
        worksheet = spreadsheet.worksheet(PICKING_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        try:
            worksheet = spreadsheet.add_worksheet(title=PICKING_SHEET_NAME, rows=1000, cols=len(PICKING_HEADERS))
        except gspread.exceptions.APIError as creation_error:
            # Another session may have created the same worksheet meanwhile.
            try:
                worksheet = spreadsheet.worksheet(PICKING_SHEET_NAME)
            except gspread.exceptions.WorksheetNotFound:
                raise creation_error
    raw = worksheet.get_all_values()
    if not raw or not any(any(picking_text(value) for value in row) for row in raw):
        worksheet.update(range_name=f"A1:{gspread.utils.rowcol_to_a1(1, len(PICKING_HEADERS))}", values=[PICKING_HEADERS], value_input_option="RAW")
    records = read_picking_records(worksheet)
    headers = [str(header).strip() for header in worksheet.row_values(1)]

    def check_existing(current_records):
        matches = [row for row in current_records if picking_text(row.get("Request ID")) == request_id]
        if not matches:
            return False
        if not picking_same_payload(matches, payload):
            raise ValueError("พบ Request ID นี้แล้วแต่ข้อมูลไม่ตรงหรือไม่ครบ กรุณาตรวจประวัติก่อนเริ่มคำขอใหม่")
        return True

    already_saved = check_existing(records)
    if not already_saved:
        headers, requests = picking_schema_requests(worksheet, headers)
        requests.append(picking_append_cells(worksheet, headers, payload))
        try:
            spreadsheet.batch_update({"requests": requests})
        except Exception:
            # A timeout can happen after Sheets has accepted the append.
            # Confirm the stable IDs before allowing a retry to append again.
            if not check_existing(read_picking_records(worksheet)):
                raise
            already_saved = True
    get_cached_picking_rows.clear()
    get_cached_data.clear()
    return {"request_id": request_id, "line_count": len(payload), "already_saved": already_saved}


def picking_schema_requests(worksheet, headers):
    """Add missing optional Picking columns without rewriting existing headers.

    Return the physical header order and requests to include in the same
    spreadsheet batch as the successful create/edit mutation.
    """
    extended_headers = [str(header).strip() for header in headers]
    if len(extended_headers) != len(set(extended_headers)):
        raise ValueError("พบชื่อคอลัมน์ซ้ำใน Picking_Requests กรุณาตรวจสอบหัวตาราง")
    missing_base = [header for header in PICKING_BASE_HEADERS if header not in extended_headers]
    if missing_base:
        raise ValueError("หัวตาราง Picking_Requests ไม่ครบ: " + ", ".join(missing_base))

    missing_optional = [header for header in PICKING_HEADERS
                        if header not in PICKING_BASE_HEADERS and header not in extended_headers]
    first_new_column = len(extended_headers)
    extended_headers.extend(missing_optional)
    requests = []
    missing_columns = len(extended_headers) - int(worksheet.col_count)
    if missing_columns > 0:
        requests.append({"appendDimension": {
            "sheetId": worksheet.id, "dimension": "COLUMNS", "length": missing_columns,
        }})
    if missing_optional:
        requests.append({"updateCells": {
            "range": {"sheetId": worksheet.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": first_new_column, "endColumnIndex": len(extended_headers)},
            "rows": [{"values": [
                {"userEnteredValue": {"stringValue": header}} for header in missing_optional
            ]}],
            "fields": "userEnteredValue",
        }})
    return extended_headers, requests


def picking_append_cells(worksheet, headers, records):
    """Build a typed RAW-style append after the last occupied sheet row."""
    headers = [str(header).strip() for header in headers]
    if len(headers) != len(set(headers)):
        raise ValueError("พบชื่อคอลัมน์ซ้ำใน Picking_Requests กรุณาตรวจสอบหัวตาราง")
    missing = [header for header in PICKING_HEADERS if header not in headers]
    if missing:
        raise ValueError("หัวตาราง Picking_Requests ไม่ครบ: " + ", ".join(missing))
    rows = []
    for record in records:
        cells = []
        for header in headers:
            if header not in PICKING_HEADERS or header not in record:
                cells.append({})
            elif header == "Volume":
                quantity = picking_quantity(record[header])
                if quantity is None or quantity <= 0:
                    raise ValueError("จำนวนต้องเป็นตัวเลขมากกว่า 0")
                cells.append({"userEnteredValue": {"numberValue": quantity}})
            else:
                cells.append({"userEnteredValue": {"stringValue": picking_text(record[header])}})
        rows.append({"values": cells})
    if not rows:
        raise ValueError("กรุณาเพิ่มสินค้าอย่างน้อย 1 รายการ")
    return {"appendCells": {
        "sheetId": worksheet.id, "rows": rows, "fields": "userEnteredValue",
    }}
def picking_shipping_date(value):
    text = picking_text(value)
    if text.upper() == "TBC":
        return "TBC"
    for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, pattern).date().isoformat()
        except ValueError:
            continue
    raise ValueError("วันจัดส่งต้องเป็น YYYY-MM-DD, DD/MM/YYYY หรือ TBC")


def picking_edit_catalog(catalog, existing_rows):
    # Saved cultivars remain editable even after their stock reaches zero.
    result = dict(catalog)
    for row in existing_rows:
        cultivar = picking_text(row.get("Cultivar Name"))
        if cultivar and picking_key(cultivar) not in result:
            result[picking_key(cultivar)] = {
                "cultivar": cultivar,
                "export_names": [picking_text(row.get("Export Name"))] if picking_text(row.get("Export Name")) else [],
                "sources": [picking_text(row.get("Source Categories"))] if picking_text(row.get("Source Categories")) else [],
                "suppliers": [picking_text(row.get("Supplier Name"))] if picking_text(row.get("Supplier Name")) else [],
                "wip_qty": 0.0, "fg_by_grade": {},
            }
    return result


def prepare_picking_changes(records, catalog, existing_rows):
    edit_catalog = picking_edit_catalog(catalog, existing_rows)
    lines = []
    for index, record in enumerate(records, start=1):
        if not any(picking_text(record.get(header)) for header in
                   ["Line ID", "Cultivar Name", "Grade", "Volume", "Consignee", "True Consignee", "Project Owner", "Type", "Expected Shipping Date", "Note"]):
            continue
        try:
            line = prepare_picking_lines([record], edit_catalog)[0]
            line["Expected Shipping Date"] = picking_shipping_date(record.get("Expected Shipping Date"))
        except ValueError as error:
            raise ValueError(f"แถวแก้ไขที่ {index}: {error}") from None
        line["Line ID"] = picking_text(record.get("Line ID"))
        # Keep saved reference metadata when only quantity/recipients/date change.
        original = next((row for row in existing_rows if picking_text(row.get("Line ID")) == line["Line ID"]), None)
        if original and picking_key(original.get("Cultivar Name")) == picking_key(line["Cultivar Name"]):
            for header in ["Cultivar Name", "Export Name", "Source Categories", "Supplier Name"]:
                line[header] = picking_text(original.get(header))
            if not line["Supplier Name"]:
                line["Supplier Name"] = " / ".join(edit_catalog[picking_key(line["Cultivar Name"])].get("suppliers", []))
        lines.append(line)
    if not lines:
        raise ValueError("ต้องมีสินค้าอย่างน้อย 1 รายการ หากต้องการยกเลิกทั้งคำขอให้ใช้ปุ่มลบ Request")
    return lines


def picking_owned_cell_updates(worksheet, headers, row_number, record=None):
    # Touch only our columns. Custom columns and formulas remain in place.
    requests = []
    owned = sorted(headers.index(header) for header in PICKING_HEADERS if header in headers)
    groups = []
    for index in owned:
        if not groups or index != groups[-1][-1] + 1:
            groups.append([])
        groups[-1].append(index)
    for group in groups:
        cells = []
        for index in group:
            if record is None:
                cells.append({})  # fields=userEnteredValue clears this cell only.
            else:
                header = headers[index]
                value = record.get(header, "")
                entered = {"numberValue": float(value)} if header == "Volume" else {"stringValue": picking_text(value)}
                cells.append({"userEnteredValue": entered})
        requests.append({"updateCells": {
            "range": {"sheetId": worksheet.id, "startRowIndex": row_number - 1,
                      "endRowIndex": row_number, "startColumnIndex": group[0],
                      "endColumnIndex": group[-1] + 1},
            "rows": [{"values": cells}], "fields": "userEnteredValue",
        }})
    return requests


def update_picking_request(spreadsheet, request_id, expected_rows, lines):
    request_id = picking_text(request_id)
    if not request_id or not expected_rows or not lines:
        raise ValueError("ข้อมูลแก้ไขไม่ครบ กรุณาเปิด Request ใหม่อีกครั้ง")
    originals = {picking_text(row.get("Line ID")): row for row in expected_rows}
    if "" in originals or len(originals) != len(expected_rows) or any(
            picking_text(row.get("Request ID")) != request_id for row in expected_rows):
        raise ValueError("ข้อมูลอ้างอิง Request ไม่ถูกต้อง กรุณาโหลดข้อมูลใหม่")
    used_ids = set(originals)
    suffixes = [int(line_id.rsplit("-", 1)[-1]) for line_id in used_ids if line_id.rsplit("-", 1)[-1].isdigit()]
    next_number = max(suffixes, default=0) + 1
    payload = []
    retained_ids = set()
    for index, line in enumerate(lines, start=1):
        line_id = picking_text(line.get("Line ID"))
        if line_id:
            if line_id not in originals or line_id in retained_ids:
                raise ValueError("Line ID ซ้ำหรือไม่อยู่ใน Request นี้ กรุณาเปิดแก้ไขใหม่")
            retained_ids.add(line_id)
        else:
            line_id = f"{request_id}-{next_number:03d}"
            while line_id in used_ids:
                next_number += 1
                line_id = f"{request_id}-{next_number:03d}"
            used_ids.add(line_id)
            next_number += 1
        quantity = picking_quantity(line.get("Volume"))
        if quantity is None or quantity <= 0:
            raise ValueError(f"รายการที่ {index}: จำนวนต้องเป็นตัวเลขมากกว่า 0")
        for required in ["Cultivar Name", "Grade", "Consignee", "True Consignee", "Project Owner", "Type"]:
            if not picking_text(line.get(required)):
                raise ValueError(f"รายการที่ {index}: กรุณาระบุ {required}")
        record = {header: picking_text(line.get(header)) for header in PICKING_HEADERS}
        record.update({
            "Request ID": request_id, "Line ID": line_id,
            "Created At": picking_text(originals.get(line_id, expected_rows[0]).get("Created At")),
            "Grade": picking_text(line.get("Grade")).upper(), "Volume": quantity,
            "Expected Shipping Date": picking_shipping_date(line.get("Expected Shipping Date")),
            "Required Category": "FG (Finished Goods)",
        })
        payload.append(record)
    worksheet = spreadsheet.worksheet(PICKING_SHEET_NAME)
    records = read_picking_records(worksheet)
    current = [row for row in records if picking_text(row.get("Request ID")) == request_id]
    already_saved = picking_same_payload(current, payload)
    if not already_saved:
        if not picking_same_payload(current, expected_rows):
            raise ValueError("Request ถูกแก้ไขหรือลบไปแล้ว กรุณายกเลิกการแก้ไขและเปิดข้อมูลล่าสุดอีกครั้ง")
        headers, requests = picking_schema_requests(worksheet, [str(header).strip() for header in worksheet.row_values(1)])
        by_id = {picking_text(row.get("Line ID")): row for row in payload}
        for row in current:
            requests.extend(picking_owned_cell_updates(
                worksheet, headers, row["Original_Row_Index"], by_id.get(picking_text(row.get("Line ID"))),
            ))
        new_rows = [row for row in payload if row["Line ID"] not in originals]
        if new_rows:
            requests.append(picking_append_cells(worksheet, headers, new_rows))
        try:
            spreadsheet.batch_update({"requests": requests})
        except Exception:
            confirmed = [row for row in read_picking_records(worksheet) if picking_text(row.get("Request ID")) == request_id]
            if not picking_same_payload(confirmed, payload):
                raise
            already_saved = True
    get_cached_picking_rows.clear()
    get_cached_data.clear()
    return {"request_id": request_id, "line_count": len(payload), "already_saved": already_saved}


def delete_picking_request(spreadsheet, request_id, expected_rows):
    request_id = picking_text(request_id)
    if not request_id or not expected_rows or any(picking_text(row.get("Request ID")) != request_id for row in expected_rows):
        raise ValueError("ข้อมูลอ้างอิง Request ไม่ถูกต้อง กรุณาเปิดข้อมูลใหม่")
    try:
        worksheet = spreadsheet.worksheet(PICKING_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = None
    current = [row for row in read_picking_records(worksheet) if picking_text(row.get("Request ID")) == request_id] if worksheet is not None else []
    already_deleted = not current
    if current:
        if not picking_same_payload(current, expected_rows):
            raise ValueError("Request ถูกแก้ไขแล้ว กรุณายกเลิกการลบและเปิดข้อมูลล่าสุดเพื่อตรวจสอบอีกครั้ง")
        headers = [str(header).strip() for header in worksheet.row_values(1)]
        requests = [request for row in current for request in
                    picking_owned_cell_updates(worksheet, headers, row["Original_Row_Index"])]
        try:
            spreadsheet.batch_update({"requests": requests})
        except Exception:
            if any(picking_text(row.get("Request ID")) == request_id for row in read_picking_records(worksheet)):
                raise
            already_deleted = True
    get_cached_picking_rows.clear()
    get_cached_data.clear()
    return {"request_id": request_id, "line_count": len(expected_rows), "already_deleted": already_deleted}


def render_picking_manager(spreadsheet, catalog):
    manager = st.session_state.get("picking_manage")
    if not manager:
        return
    request_id = manager["request_id"]
    token = manager["token"]
    with st.container(border=True, key="picking_management"):
        if manager["mode"] == "edit":
            st.markdown(f"**แก้ไข Request: {request_id}**")
            st.caption("แก้สายพันธุ์ เกรด จำนวน ผู้รับ Project Owner, Type และวันส่งรายรายการได้ เพิ่มหรือลบแถวได้ในตาราง โดยคงหมายเลข Request เดิม")
            st.caption("วันจัดส่ง: YYYY-MM-DD หรือ DD/MM/YYYY · ใส่ TBC เมื่อยังไม่มีกำหนดส่ง แต่ละรายการเก็บวันส่งแยกกัน")
            edit_catalog = picking_edit_catalog(catalog, manager["expected"])
            options = sorted({item["cultivar"] for item in edit_catalog.values()} |
                             {picking_text(row.get("Cultivar Name")) for row in manager["expected"]
                              if picking_text(row.get("Cultivar Name"))}, key=str.casefold)
            records = st.data_editor(
                manager["editor_input"], width="stretch", hide_index=True, num_rows="dynamic",
                key=f"picking_manage_editor_{token}",
                column_config={
                    "Line ID": None,
                    "Cultivar Name": st.column_config.SelectboxColumn("สายพันธุ์", options=options, required=True, width="medium"),
                    "Grade": st.column_config.TextColumn("เกรด FG", required=True, width="small"),
                    "Volume": st.column_config.NumberColumn("จำนวนส่ง (Volume)", min_value=0.0, step=0.01, format="%.2f", required=True),
                    "Consignee": st.column_config.TextColumn("Consignee", required=True, width="medium"),
                    "True Consignee": st.column_config.TextColumn("True Consignee", required=True, width="medium"),
                    "Project Owner": st.column_config.TextColumn("Project Owner", required=True, width="medium"),
                    "Type": st.column_config.TextColumn("Type ผลิตภัณฑ์", required=True, width="medium"),
                    "Expected Shipping Date": st.column_config.TextColumn("วันจัดส่ง / TBC", required=True, width="medium"),
                    "Note": st.column_config.TextColumn("หมายเหตุ", width="medium"),
                },
            ).to_dict("records")
            try:
                lines = prepare_picking_changes(records, catalog, manager["expected"])
            except ValueError as error:
                lines = None
                st.caption(str(error))
            else:
                preview = project_picking_rows([dict(line, **{"Request ID": request_id}) for line in lines], catalog)
                st.markdown("**ความพร้อมหลังแก้ไข เทียบ On hand ปัจจุบัน**")
                st.dataframe(
                    format_df_quantity(pd.DataFrame(preview)), width="stretch", hide_index=True,
                    column_order=["Cultivar Name", "Grade", "Volume", "Supplier Name", "Project Owner", "Type", "Expected Shipping Date", "Available FG", "Available WIP", "FG Shortfall", "Status (ความพร้อม)"],
                    column_config={"Cultivar Name": "สายพันธุ์", "Grade": "เกรด FG", "Volume": "จำนวนส่ง",
                                   "Expected Shipping Date": "วันจัดส่ง / TBC", "Available FG": "FG ตรงเกรด",
                                   "Available WIP": "WIP รอ Process", "FG Shortfall": "FG ที่ยังขาด"},
                )
            save_col, cancel_col = st.columns([2, 1])
            with save_col:
                if st.button("บันทึกการแก้ไข", key="picking_update_request", type="primary", width="stretch"):
                    try:
                        if lines is None:
                            lines = prepare_picking_changes(records, catalog, manager["expected"])
                        with st.spinner("กำลังบันทึกการแก้ไข..."):
                            result = update_picking_request(spreadsheet, request_id, manager["expected"], lines)
                    except Exception as error:
                        st.error(f"ยังยืนยันการแก้ไขไม่ได้: {error}")
                        st.caption("ข้อมูลที่กรอกยังอยู่ หากการเชื่อมต่อขัดข้องให้กดบันทึกซ้ำด้วยข้อมูลเดิม")
                    else:
                        st.session_state["picking_manage_flash"] = f"แก้ไข {request_id} แล้ว · {result['line_count']:,} รายการ"
                        st.session_state["picking_manage_draft_result"] = result
                        del st.session_state["picking_manage"]
                        st.rerun()
            with cancel_col:
                if st.button("ยกเลิกการแก้ไข", key="picking_manage_cancel", width="stretch"):
                    del st.session_state["picking_manage"]
                    get_cached_picking_rows.clear()
                    st.rerun()
        else:
            st.markdown(f"**ลบ Request: {request_id}**")
            volume = sum(picking_quantity(row.get("Volume")) or 0 for row in manager["expected"])
            st.warning(f"จะลบทั้งคำขอ {len(manager['expected']):,} รายการ รวม Volume {volume:,.2f} และนำรายการเหล่านี้ออกจาก Delivery Issue การลบไม่สามารถย้อนกลับได้")
            confirmed = st.checkbox(f"ยืนยันว่าต้องการลบ {request_id}", key=f"picking_delete_confirm_{token}")
            delete_col, cancel_col = st.columns([2, 1])
            with delete_col:
                if st.button("ยืนยันลบ Request", key="picking_confirm_delete", type="primary", disabled=not confirmed, width="stretch"):
                    try:
                        with st.spinner("กำลังลบ Picking Request..."):
                            result = delete_picking_request(spreadsheet, request_id, manager["expected"])
                    except Exception as error:
                        st.error(f"ยังยืนยันการลบไม่ได้: {error}")
                    else:
                        st.session_state["picking_manage_flash"] = f"ลบ {request_id} และนำรายการออกจาก Delivery Issue แล้ว"
                        st.session_state["picking_manage_draft_result"] = dict(result, deleted=True)
                        del st.session_state["picking_manage"]
                        st.rerun()
            with cancel_col:
                if st.button("ยกเลิกการลบ", key="picking_manage_cancel", width="stretch"):
                    del st.session_state["picking_manage"]
                    get_cached_picking_rows.clear()
                    st.rerun()


def find_delivery_column(columns, field):
    aliases = {
        "Cultivar Name": ["Cultivar Name", "Cultivar", "Cultivar / Product", "Product", "ชื่อ"],
        "Grade": ["Grade", "Size"],
        "Consignee": ["Consignee", "Consignee (ปลายทาง)"],
        "True Consignee": ["True Consignee", "True_Consignee", "True Consignee (ปลายทางจริง)"],
        "Supplier Name": ["Supplier Name", "Supplier"],
        "Project Owner": ["Project Owner", "Project_Owner", "Project owner"],
        "Type": ["Type", "Product Type", "Type ผลิตภัณฑ์"],
        "Expected Shipping Date": ["Expected Shipping Date", "Shipment date", "Delivery date", "Shipment Date", "Delivery Date"],
    }
    if field in {"Volume", "Available FG", "Available WIP"}:
        return next((column for column in columns if picking_key(field) in picking_key(column)), field)
    for alias in aliases.get(field, [field]):
        for column in columns:
            if picking_key(column) == picking_key(alias):
                return column
    return field


def merge_delivery_requests(legacy_records, picking_records, catalog):
    columns = list(dict.fromkeys(key for row in legacy_records for key in row))
    rows = [dict(row, _Delivery_Source="legacy") for row in legacy_records]
    mappings = {field: find_delivery_column(columns, field) for field in [
        "Cultivar Name", "Grade", "Volume", "Consignee", "True Consignee",
        "Expected Shipping Date", "Available FG", "Available WIP", "Project Owner", "Type", "Supplier Name",
    ]}
    for record in project_picking_rows(picking_records, catalog):
        row = {mappings.get(key, key): value for key, value in record.items()}
        row["_Delivery_Source"] = "picking"
        rows.append(row)
    return rows


def update_delivery_dates(spreadsheet, targets, new_value):
    grouped = {}
    live_picking = {}
    for target in targets:
        worksheet = target["worksheet"]
        headers = [str(header).strip() for header in worksheet.row_values(1)]
        date_header = target["date_header"]
        if date_header not in headers:
            raise ValueError(f"ไม่พบคอลัมน์ {date_header} ใน {worksheet.title}")
        row_number = int(target["row_idx"])
        if target.get("line_id"):
            if worksheet.title not in live_picking:
                live_picking[worksheet.title] = read_picking_records(worksheet)
            matches = [row for row in live_picking[worksheet.title]
                       if picking_text(row.get("Line ID")) == picking_text(target["line_id"])
                       and picking_text(row.get("Request ID")) == picking_text(target.get("request_id"))]
            if len(matches) != 1:
                raise ValueError("รายการ Picking ถูกเปลี่ยนหรือลบแล้ว กรุณาโหลด Delivery Issue ใหม่")
            row_number = matches[0]["Original_Row_Index"]
            if "expected_date" in target and picking_text(matches[0].get("Expected Shipping Date")) != picking_text(target["expected_date"]):
                raise ValueError("วันส่งรายการนี้ถูกแก้ไขแล้ว กรุณาโหลด Delivery Issue ใหม่")
        if row_number < 2:
            raise ValueError("เลขแถวรายการไม่ถูกต้อง")
        group = grouped.setdefault(worksheet.title, {"worksheet": worksheet, "cells": []})
        group["cells"].append(gspread.Cell(row=row_number, col=headers.index(date_header) + 1, value=new_value))
    if len(grouped) == 1:
        group = next(iter(grouped.values()))
        group["worksheet"].update_cells(group["cells"])
    elif grouped:
        requests = []
        for group in grouped.values():
            for cell in group["cells"]:
                requests.append({"updateCells": {
                    "range": {"sheetId": group["worksheet"].id, "startRowIndex": cell.row - 1,
                              "endRowIndex": cell.row, "startColumnIndex": cell.col - 1,
                              "endColumnIndex": cell.col},
                    "rows": [{"values": [{"userEnteredValue": {"stringValue": new_value}}]}],
                    "fields": "userEnteredValue",
                }})
        spreadsheet.batch_update({"requests": requests})



def render_portal_ui_styles():
    st.markdown("""<style>
    .st-key-delivery_header h1, .st-key-stock_header h1, .st-key-picking_header h1 {
        color: #0f172a; font-size: clamp(27px, 2.8vw, 36px); line-height: 1.3;
        letter-spacing: -.035em; padding-bottom: 8px; }
    .st-key-delivery_header [data-testid="stCaptionContainer"],
    .st-key-stock_header [data-testid="stCaptionContainer"] { color: #64748b; font-size: 14px; }
    .st-key-delivery_header [data-testid="stCaptionContainer"] p,
    .st-key-stock_header [data-testid="stCaptionContainer"] p { color: #64748b; }
    .st-key-picking_form [data-testid="stDataFrame"] { border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; }
    .st-key-picking_form [data-testid="stButton"] button { background: #1d4ed8; color: #fff; border: 1px solid #1d4ed8; min-height: 44px; border-radius: 10px; }
    .st-key-picking_form [data-testid="stButton"] button:hover { background: #1e40af; }
    .portal-section { --section-accent: #2563eb; margin: 12px 0 18px; }
    .portal-section--slate { --section-accent: #475569; }
    .portal-section--amber { --section-accent: #b45309; }
    .portal-section--teal { --section-accent: #0f766e; }
    .portal-section__eyebrow { color: var(--section-accent); font-size: 11px;
        font-weight: 750; letter-spacing: .12em; margin-bottom: 6px; }
    .portal-section h2 { margin: 0; padding: 0; color: #0f172a;
        font-size: 21px; line-height: 1.45; font-weight: 750; }
    .portal-section p { margin: 6px 0 0; color: #64748b; font-size: 14px; line-height: 1.7; }
    .st-key-delivery_overview [data-testid="stHorizontalBlock"],
    .st-key-stock_overview [data-testid="stHorizontalBlock"] {
        display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 14px; }
    .st-key-stock_overview [data-testid="stHorizontalBlock"] { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    .st-key-delivery_overview [data-testid="stColumn"],
    .st-key-stock_overview [data-testid="stColumn"] { min-width: 0 !important; width: auto !important; }
    .st-key-delivery_overview [data-testid="stMetric"],
    .st-key-stock_overview [data-testid="stMetric"] {
        padding: 18px 20px; background: #fff; border: 1px solid #e2e8f0;
        border-top: 3px solid #2563eb; border-radius: 14px; box-shadow: 0 3px 12px #0f172a04; }
    .st-key-delivery_overview [data-testid="stMetricLabel"] p,
    .st-key-stock_overview [data-testid="stMetricLabel"] p { color: #475569; font-size: 13px; line-height: 1.5;
        white-space: normal !important; overflow: visible !important; text-overflow: clip !important; }
    .st-key-delivery_overview [data-testid="stMetricValue"],
    .st-key-stock_overview [data-testid="stMetricValue"] {
        color: #0f172a; font-size: 30px; font-weight: 750; font-variant-numeric: tabular-nums; }
    .st-key-delivery_overview [data-testid="stColumn"]:nth-child(3) [data-testid="stMetric"] { border-top-color: #0f766e; }
    .st-key-delivery_overview [data-testid="stColumn"]:nth-child(4) [data-testid="stMetric"] { border-top-color: #94a3b8; }
    .st-key-stock_overview [data-testid="stColumn"]:first-child [data-testid="stMetric"] { border-top-color: #334155; }
    .st-key-stock_overview [data-testid="stColumn"]:nth-child(2) [data-testid="stMetric"] { border-top-color: #d97706; background: #fffcf5; }
    .st-key-stock_overview [data-testid="stColumn"]:nth-child(2) [data-testid="stMetricValue"] { color: #92400e; }
    .st-key-stock_overview [data-testid="stColumn"]:nth-child(3) [data-testid="stMetric"] { border-top-color: #0f766e; background: #f6fefc; }
    .st-key-stock_overview [data-testid="stColumn"]:nth-child(3) [data-testid="stMetricValue"] { color: #115e59; }
    .portal-balance { margin: 8px 0 24px; padding: 16px 20px; background: #f8fafc;
        border: 1px solid #e2e8f0; border-radius: 12px; }
    .portal-balance__labels { display: flex; justify-content: space-between; gap: 12px;
        margin-bottom: 10px; color: #475569; font-size: 12px; }
    .portal-balance__labels strong { font-weight: 700; }
    .portal-balance__wip { color: #92400e; }
    .portal-balance__fg { color: #115e59; }
    .portal-balance__track { display: flex; height: 8px; overflow: hidden;
        border-radius: 99px; background: #e2e8f0; }
    .portal-balance__segment--wip { background: #d97706; }
    .portal-balance__segment--fg { background: #0f766e; }
    .st-key-delivery_reschedule { margin: 12px 0 24px; }
    .st-key-delivery_reschedule [data-testid="stForm"] { background: #f8fafc;
        border: 1px solid #e2e8f0; border-radius: 14px; padding: 20px; }
    .st-key-delivery_reschedule [data-testid="stForm"] label p { font-size: 14px; color: #334155; }
    .st-key-delivery_reschedule [data-testid="stFormSubmitButton"] button {
        background: #1d4ed8; color: #fff; border: 1px solid #1d4ed8;
        border-radius: 10px; min-height: 44px; font-weight: 650; }
    .st-key-delivery_reschedule [data-testid="stFormSubmitButton"] button:hover { background: #1e40af; border-color: #1e40af; }
    .portal-day { --day-accent: #0f766e; --day-tint: #f0fdfa; --day-line: #99f6e4;
        display: flex; align-items: center; justify-content: space-between;
        flex-wrap: wrap; gap: 18px; padding: 20px; margin-bottom: 12px;
        border-left: 4px solid var(--day-accent); border-radius: 12px;
        background: linear-gradient(110deg, var(--day-tint), #fff 75%); }
    .portal-day--overdue { --day-accent: #be123c; --day-tint: #fff1f2; --day-line: #fecdd3; }
    .portal-day--today { --day-accent: #b45309; --day-tint: #fffbeb; --day-line: #fde68a; }
    .portal-day--soon { --day-accent: #1d4ed8; --day-tint: #eff6ff; --day-line: #bfdbfe; }
    .portal-day--pending { --day-accent: #475569; --day-tint: #f8fafc; --day-line: #cbd5e1; }
    .portal-day__main { min-width: 0; }
    .portal-day__eyebrow { color: var(--day-accent); font-size: 11px; font-weight: 750;
        letter-spacing: .09em; margin-bottom: 6px; }
    .portal-day__title { color: #0f172a; font-size: 21px; line-height: 1.5; font-weight: 750; }
    .portal-day__meta { display: flex; align-items: center; flex-wrap: wrap; gap: 10px;
        color: #64748b; font-size: 13px; line-height: 1.7; margin-top: 10px; }
    .portal-day__badge { display: inline-block; padding: 4px 10px; border-radius: 99px;
        margin-top: 0; color: var(--day-accent); background: var(--day-tint);
        border: 1px solid var(--day-line); font-size: 12px; font-weight: 650; }
    .portal-day-volume { min-width: 180px; padding: 12px 18px; background: #fff;
        border: 1px solid var(--day-line); border-radius: 12px; }
    .portal-day-volume__label { color: #475569; font-size: 12px; margin-bottom: 4px; }
    .portal-day-volume__value { color: var(--day-accent); font-size: 28px;
        font-weight: 750; font-variant-numeric: tabular-nums; line-height: 1.4; }
    .st-key-delivery_list [data-testid="stDataFrame"],
    .st-key-stock_details_wip [data-testid="stDataFrame"],
    .st-key-stock_details_fg [data-testid="stDataFrame"] {
        border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; }
    .st-key-stock_details_wip, .st-key-stock_details_fg { padding: 22px;
        border: 1px solid #fde68a; border-top: 3px solid #d97706;
        border-radius: 16px; background: #fff; margin-bottom: 20px; }
    .st-key-stock_details_fg { border-color: #99f6e4; border-top-color: #0f766e; }
    .portal-grade-grid { --grade-tint: #fffbeb; --grade-line: #fde68a; --grade-text: #92400e;
        display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 10px; margin: 0 0 18px; }
    .portal-grade-grid--fg { --grade-tint: #f0fdfa; --grade-line: #99f6e4; --grade-text: #115e59; }
    .portal-grade { padding: 12px 14px; border: 1px solid #e2e8f0; border-radius: 10px; background: #f8fafc; }
    .portal-grade dt { margin: 0 0 6px; color: #64748b; font-size: 11px; font-weight: 700; letter-spacing: .05em; }
    .portal-grade dd { margin: 0; color: #334155; font-size: 18px; font-weight: 700; font-variant-numeric: tabular-nums; }
    .portal-grade dd small { font-size: 11px; font-weight: 500; }
    .portal-grade--total { background: var(--grade-tint); border-color: var(--grade-line); }
    .portal-grade--total dt, .portal-grade--total dd { color: var(--grade-text); }
    .portal-empty { color: #64748b; background: #f8fafc; border: 1px dashed #cbd5e1;
        padding: 28px 20px; border-radius: 12px; text-align: center; font-size: 14px; line-height: 1.7; }
    @media (max-width: 760px) {
        .portal-grade-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        .portal-grade--total { grid-column: 1 / -1; }
        .portal-day { align-items: stretch; flex-direction: column; gap: 14px; padding: 16px; }
        .portal-day__title { font-size: 19px; }
        .portal-day-volume { display: flex; align-items: center; justify-content: space-between;
            flex-wrap: wrap; gap: 8px; min-width: 0; padding: 10px 14px; }
        .portal-day-volume__label { margin: 0; }
        .portal-day-volume__value { font-size: 25px; }
    }
    @media (max-width: 640px) {
        .st-key-delivery_overview [data-testid="stHorizontalBlock"],
        .st-key-stock_overview [data-testid="stHorizontalBlock"] { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
        .st-key-stock_overview [data-testid="stColumn"]:first-child { grid-column: 1 / -1; }
        .st-key-delivery_overview [data-testid="stMetric"],
        .st-key-stock_overview [data-testid="stMetric"] { padding: 14px; }
        .st-key-delivery_overview [data-testid="stMetricValue"],
        .st-key-stock_overview [data-testid="stMetricValue"] { font-size: 21px; }
        .st-key-delivery_overview [data-testid="stMetricLabel"],
        .st-key-stock_overview [data-testid="stMetricLabel"] { height: auto; min-height: 40px; }
        .st-key-stock_overview [data-testid="stColumn"]:first-child [data-testid="stMetricValue"] { font-size: 28px; }
        .st-key-stock_details_wip, .st-key-stock_details_fg { padding: 16px; }
        .portal-section h2 { font-size: 19px; }
        .portal-section p { font-size: 13px; }
        .portal-balance { padding: 14px; }
    }
    </style>""", unsafe_allow_html=True)


def render_section_heading(title, description, eyebrow="", tone="blue"):
    if tone not in {"blue", "slate", "amber", "teal"}:
        tone = "blue"
    st.markdown(
        f'<header class="portal-section portal-section--{tone}">'
        f'<div class="portal-section__eyebrow">{escape(eyebrow)}</div>'
        f'<h2>{escape(title)}</h2><p>{escape(description)}</p></header>',
        unsafe_allow_html=True,
    )


def render_grade_summary(stock_df, total, stock_type):
    grade_names = stock_df["Grade"].astype(str).str.strip().str.upper()
    items = [(grade, stock_df.loc[grade_names.eq(grade), "Clean_Qty"].sum())
             for grade in ("BIG", "SMALL", "MIX", "F")]
    items.append((f"รวม {stock_type}", total))
    html = [f'<dl class="portal-grade-grid portal-grade-grid--{stock_type.lower()}">']
    for index, (label, quantity) in enumerate(items):
        total_class = " portal-grade--total" if index == len(items) - 1 else ""
        html.append(
            f'<div class="portal-grade{total_class}"><dt>{escape(label)}</dt>'
            f'<dd>{quantity:,.2f} <small>g</small></dd></div>'
        )
    html.append('</dl>')
    st.markdown("".join(html), unsafe_allow_html=True)


def render_registration_ui_styles():
    st.markdown("""<style>
    .st-key-registration_create, .st-key-registration_edit {
        --reg-accent: #0f766e; --reg-tint: #f0fdfa; --reg-line: #b8e4dc;
        max-width: 1140px; margin-inline: auto; color: #334155; }
    .st-key-registration_edit { --reg-accent: #1d4ed8; --reg-tint: #eff6ff; --reg-line: #c7dafa; }
    :is(.st-key-registration_create, .st-key-registration_edit) .st-key-registration_header h1 {
        color: #0f172a; font-size: clamp(27px, 2.6vw, 35px); line-height: 1.35;
        letter-spacing: -.025em; padding: 0 0 8px; }
    :is(.st-key-registration_create, .st-key-registration_edit) .st-key-registration_header [data-testid="stCaptionContainer"] p {
        color: #64748b; font-size: 14px; line-height: 1.8; }
    .registration-eyebrow { margin: 0 0 8px; color: var(--reg-accent);
        font-size: 11px; letter-spacing: .12em; font-weight: 750; }
    .registration-steps { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 10px; margin: 2px 0 16px; }
    .registration-step { display: flex; align-items: center; gap: 10px; padding: 12px 14px;
        border: 1px solid #e2e8f0; border-radius: 12px; background: #f8fafc;
        color: #475569; font-size: 13px; line-height: 1.5; }
    .registration-step span { display: grid; place-items: center; flex: 0 0 26px;
        width: 26px; height: 26px; background: #fff; border: 1px solid #dbe4ee;
        border-radius: 8px; color: var(--reg-accent); font-size: 11px; font-weight: 750; }
    :is(.st-key-registration_create, .st-key-registration_edit) [data-testid="stForm"] {
        background: transparent; border: 0; padding: 0; }
    :is(.st-key-registration_create, .st-key-registration_edit) [class*="st-key-registration_card_"] {
        padding: 24px; border: 1px solid #e2e8f0; border-radius: 16px;
        background: #fff; margin-bottom: 6px; box-shadow: 0 3px 12px #0f172a03; }
    .registration-card-heading { display: flex; align-items: flex-start; gap: 12px; margin: 0 0 8px; }
    .registration-card-heading__number { display: grid; place-items: center; width: 34px; height: 34px;
        flex: 0 0 34px; border: 1px solid var(--reg-line); background: var(--reg-tint);
        color: var(--reg-accent); font-size: 12px; font-weight: 750; border-radius: 10px; }
    .registration-card-heading h2 { color: #0f172a; font-size: 19px; line-height: 1.5;
        font-weight: 700; padding: 0; margin: 0 0 4px; }
    .registration-card-heading p { color: #64748b; font-size: 13px; line-height: 1.7; margin: 0; }
    :is(.st-key-registration_create, .st-key-registration_edit) [data-testid="stWidgetLabel"] p {
        color: #334155; font-size: 14px; font-weight: 600; line-height: 1.6; }
    :is(.st-key-registration_create, .st-key-registration_edit) [data-baseweb="input"],
    :is(.st-key-registration_create, .st-key-registration_edit) [data-baseweb="select"] > div {
        border: 1px solid #dbe3ec; border-radius: 10px; background: #f8fafc;
        min-height: 45px; box-shadow: none; }
    :is(.st-key-registration_create, .st-key-registration_edit) [data-baseweb="input"]:focus-within,
    :is(.st-key-registration_create, .st-key-registration_edit) [data-baseweb="select"]:focus-within > div {
        border-color: var(--reg-accent); box-shadow: 0 0 0 3px var(--reg-tint); }
    :is(.st-key-registration_create, .st-key-registration_edit) input {
        color: #1e293b; font-size: 15px; }
    :is(.st-key-registration_create, .st-key-registration_edit) [data-testid="stCaptionContainer"] p {
        color: #64748b; font-size: 13px; line-height: 1.75; }
    :is(.st-key-registration_create, .st-key-registration_edit) [data-testid="stFormSubmitButton"] button {
        color: #fff; background: var(--reg-accent); border: 1px solid var(--reg-accent);
        border-radius: 11px; min-height: 48px; padding: 12px 24px; font-weight: 650; box-shadow: none; }
    :is(.st-key-registration_create, .st-key-registration_edit) [data-testid="stFormSubmitButton"] button:hover {
        filter: brightness(.93); color: #fff; }
    .st-key-registration_filter { padding: 20px; border: 1px solid #dbe4ee;
        background: #f8fafc; border-radius: 14px; margin: 0 0 4px; }
    .registration-selected { padding: 18px 22px; border: 1px solid var(--reg-line);
        border-left: 4px solid var(--reg-accent); border-radius: 14px; background: var(--reg-tint);
        display: flex; justify-content: space-between; align-items: center; gap: 18px; margin: 4px 0 12px; }
    .registration-selected__label { color: var(--reg-accent); font-size: 11px; font-weight: 700; margin-bottom: 5px; }
    .registration-selected h2 { margin: 0 0 6px; padding: 0; font-size: 21px; line-height: 1.5;
        color: #0f172a; overflow-wrap: anywhere; }
    .registration-selected p { color: #475569; margin: 0; font-size: 13px; line-height: 1.8; overflow-wrap: anywhere; }
    .registration-selected__quantity { flex-shrink: 0; text-align: right; }
    .registration-selected__quantity strong { font-size: 25px; color: var(--reg-accent); font-variant-numeric: tabular-nums; }
    .registration-selected__quantity small { display: block; color: #64748b; font-size: 12px; margin-top: 3px; }
    .registration-submit-note { color: #64748b; font-size: 13px; line-height: 1.8; padding: 4px 0 8px; }
    @media (max-width: 740px) {
        :is(.st-key-registration_create, .st-key-registration_edit) [data-testid="stHorizontalBlock"] {
            flex-direction: column; gap: 16px; }
        :is(.st-key-registration_create, .st-key-registration_edit) [data-testid="stColumn"] {
            width: 100% !important; min-width: 0 !important; flex: 1 1 100% !important; }
        .registration-steps { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }
        .registration-step { padding: 10px; font-size: 12px; gap: 8px; }
        :is(.st-key-registration_create, .st-key-registration_edit) [class*="st-key-registration_card_"] { padding: 18px 16px; }
        .registration-card-heading { gap: 10px; }
        .registration-card-heading h2 { font-size: 18px; }
        .registration-selected { align-items: flex-start; flex-direction: column; gap: 12px; padding: 16px; }
        .registration-selected__quantity { text-align: left; }
        .registration-selected__quantity strong { font-size: 23px; }
        .st-key-registration_filter { padding: 16px; }
    }
    </style>""", unsafe_allow_html=True)


def render_registration_steps():
    labels = ["ข้อมูลสินค้า", "Supplier และ Lot", "รับเข้าและจำนวน", "ตรวจสอบและแผนส่ง"]
    st.markdown('<nav class="registration-steps" aria-label="ส่วนของแบบฟอร์ม">' + ''.join(
        f'<div class="registration-step"><span>{index:02d}</span>{label}</div>'
        for index, label in enumerate(labels, start=1)
    ) + '</nav>', unsafe_allow_html=True)


def registration_card_heading(number, title, description):
    st.markdown(
        f'<header class="registration-card-heading"><span class="registration-card-heading__number">{number:02d}</span>'
        f'<div><h2>{escape(title)}</h2><p>{escape(description)}</p></div></header>',
        unsafe_allow_html=True,
    )


def registration_date(value):
    parsed = parse_delivery_date(picking_text(value))
    return parsed.date() if pd.notna(parsed) else datetime.now(ZoneInfo("Asia/Bangkok")).date()


def render_registration_form(editing=False, current_data=None, row_number=None):
    current = dict(current_data) if current_data is not None else {}
    prefix = f"edit_{row_number}_" if editing else "reg_"
    if not editing and st.session_state.pop("registration_create_reset", False):
        for key in list(st.session_state):
            if key.startswith("reg_"):
                del st.session_state[key]

    def text(field):
        return picking_text(current.get(field))

    def select(label, field, suffix, options):
        old = text(field) if editing else options[0]
        choices = list(options)
        if old and old not in choices:
            choices.append(old)
        return st.selectbox(label, choices, index=choices.index(old) if old in choices else None,
                            key=prefix + suffix, placeholder="เลือกข้อมูล") or ""

    render_registration_steps()
    with st.form("edit_form" if editing else "trading_form"):
        with st.container(key="registration_card_product"):
            registration_card_heading(1, "ข้อมูลสินค้า", "ระบุชื่อสายพันธุ์ ประเภทสินค้า และเกรดที่ใช้ติดตามสต็อก")
            left, right = st.columns(2, gap="large")
            with left:
                cultivar = st.text_input("Cultivar Name · ชื่อสายพันธุ์", value=text("Cultivar Name"),
                                         placeholder="เช่น Blue Dream", key=prefix + "cultivar")
            with right:
                export = st.text_input("Export Name · ชื่อส่งออก", value=text("Export Name"),
                                       placeholder="ชื่อผลิตภัณฑ์สำหรับส่งออก", key=prefix + "export")
            left, right = st.columns(2, gap="large")
            with left:
                category = select("Category · ประเภทสต็อก", "Category", "category", ["FG (Finished Goods)", "WIP (Work in Progress)"])
            with right:
                grade = select("Grade · เกรด", "Grade", "grade", ["Big", "Small", "Mix", "F", "Grade A", "Premium"])
            left, right = st.columns(2, gap="large")
            with left:
                product_type = select("Type · รูปแบบผลิตภัณฑ์", "Type", "type", [
                    "Inspection & Packing Service", "Trading Goods (TS)", "Consignment", "In House Product", "To be decided",
                ])
            with right:
                status = select("Status · สถานะสินค้า", "Status", "status", [
                    "Trading Goods Received to Stock", "Inspection Completed", "Pending Packing", "Packing done", "Delivered", "Waiting for Return",
                ])
        with st.container(key="registration_card_supplier"):
            registration_card_heading(2, "Supplier และเลข Lot", "เชื่อมโยงสินค้ากับผู้จัดหาและเลขล็อตสำหรับตรวจสอบย้อนกลับ")
            left, right = st.columns(2, gap="large")
            with left:
                supplier = st.text_input("Supplier Name · ผู้จัดหา", value=text("Supplier Name"),
                                         placeholder="ระบุชื่อผู้จัดหา", key=prefix + "supplier")
            with right:
                supplier_lot = st.text_input("Supplier Lot No. · ล็อตผู้จัดหา", value=text("Supplier Lot. No."),
                                             placeholder="เลข Lot จาก Supplier", key=prefix + "supplier_lot")
            thai_lot = st.text_input("THAI STICK Lot No. · เลขล็อตภายใน", value=text("THAI STICK Lot. No."),
                                     placeholder="ระบุเลข Lot ที่ใช้ติดตามในระบบ", key=prefix + "thai_lot")
        with st.container(key="registration_card_inventory"):
            registration_card_heading(3, "ข้อมูลรับเข้าและจำนวน", "ระบุวันที่รับสินค้าและน้ำหนักเป็นกรัม พร้อมรายละเอียดจำนวนแพ็กถ้ามี")
            left, right = st.columns(2, gap="large")
            with left:
                receive_date = st.date_input("Receive date · วันที่รับเข้า", key=prefix + "receive_date",
                    value=registration_date(current.get("Receive date", current.get("Date receive", current.get("Date Receive", "")))))
            with right:
                quantity = st.number_input("Quantity · จำนวน (g)", min_value=0.0, step=0.1, format="%.2f",
                                           value=max(0.0, picking_quantity(current.get("Quantity")) or 0.0), key=prefix + "quantity")
            remark = st.text_input("Remark Qty · รายละเอียดจำนวน", value=text("Remark Qty"),
                                    placeholder="เช่น 5 × 1,000 g หรือ 10 แพ็ก", key=prefix + "remark_qty")
        with st.container(key="registration_card_shipping"):
            registration_card_heading(4, "การตรวจสอบและแผนส่ง", "กำหนดเงื่อนไขการตรวจสอบและวันจัดส่ง หรือเลือก To be decided หากยังไม่ยืนยัน")
            left, right = st.columns(2, gap="large")
            with left:
                inspection = select("Inspection Req. · ต้องตรวจสอบ", "Inspection Req.", "inspection", ["Yes", "No", "To be decided"])
            with right:
                ir = select("IR Req. · ข้อกำหนด IR", "IR Req.", "ir", ["Yes", "No", "To be decided"])
            left, right = st.columns(2, gap="large")
            with left:
                shipment = st.date_input("Shipment date · วันที่จัดส่ง", value=registration_date(current.get("Shipment date")),
                                          key=prefix + "shipment_date")
            with right:
                ship_tbd = st.checkbox("ยังไม่กำหนดวันจัดส่ง", key=prefix + "ship_tbd",
                    value=editing and picking_text(current.get("Shipment date")).casefold() in DELIVERY_PENDING_VALUES)
                st.caption("เมื่อเลือก ระบบจะบันทึก To be decided แทนวันที่ในช่องจัดส่ง")
        st.markdown('<p class="registration-submit-note">ตรวจชื่อสายพันธุ์ เลข Lot และจำนวนให้ครบก่อนบันทึก</p>', unsafe_allow_html=True)
        submitted = st.form_submit_button("บันทึกการแก้ไข" if editing else "บันทึกสินค้าเข้าระบบ", type="primary", width="stretch")

    return submitted, {
        "Cultivar Name": cultivar, "Supplier Name": supplier, "Supplier Lot. No.": supplier_lot,
        "Export Name": export, "THAI STICK Lot. No.": thai_lot, "Category": category, "Grade": grade,
        "Receive date": str(receive_date), "Quantity": float(quantity), "Remark Qty": remark,
        "Inspection Req.": inspection, "IR Req.": ir, "Shipment date": "To be decided" if ship_tbd else str(shipment),
        "Type": product_type, "Status": status, "Packing date": text("Packing date") if editing else "",
    }


def render_registration_selected(current):
    cultivar = picking_text(current.get("Cultivar Name")) or "ไม่ระบุสายพันธุ์"
    lot = picking_text(current.get("THAI STICK Lot. No.")) or picking_text(current.get("Supplier Lot. No.")) or "ไม่ระบุ Lot"
    supplier = picking_text(current.get("Supplier Name")) or "ไม่ระบุ Supplier"
    category = picking_text(current.get("Category")) or "ไม่ระบุประเภท"
    grade = picking_text(current.get("Grade")) or "ไม่ระบุเกรด"
    quantity = picking_quantity(current.get("Quantity")) or 0.0
    st.markdown(
        '<section class="registration-selected"><div><div class="registration-selected__label">ล็อตที่กำลังแก้ไข</div>'
        f'<h2>{escape(cultivar)}</h2><p>Lot {escape(lot)} · {escape(supplier)}</p>'
        f'<p>{escape(category)} · เกรด {escape(grade)}</p></div>'
        f'<div class="registration-selected__quantity"><strong>{quantity:,.2f} g</strong><small>จำนวนคงเหลือปัจจุบัน</small></div></section>',
        unsafe_allow_html=True,
    )


# ==========================================
# 5. Sidebar เมนูด้านซ้าย
# ==========================================
st.sidebar.title("🌱 Thai Stick Trading")
st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "เลือกเมนูการใช้งาน",
    ["🚨 Delivery Issue", "📋 Picking Request", "📊 Stock Summary & Details", "📈 Monthly Report", "📝 Trading Goods Registration", "✏️ Edit Registration", "📦 Packing Module", "🚚 Shipped Out Module"],
)
st.sidebar.markdown("---")
st.sidebar.info("ระบบจัดการสินค้า Trading และกระบวนการ Packing (Pastel Theme)")


# ==========================================
# หน้าที่ 1: Delivery Issue
# ==========================================
if menu == "🚨 Delivery Issue":
    render_portal_ui_styles()
    with st.container(key="delivery_header"):
        col_head1, col_head2 = st.columns([4, 1])
        with col_head1:
            st.title("Delivery Issue")
            st.caption("ติดตามกำหนดส่ง ความพร้อมของสินค้า และยอดจัดส่งในแต่ละวัน")
        with col_head2:
            with st.expander("🔗 เปิด Google Sheets"):
                pwd = st.text_input("รหัสผ่าน:", type="password", key="gsheet_shortcut_pwd_del")
                if pwd == "tstd": st.markdown("[🟢 **เปิด Google Sheets**](https://docs.google.com/spreadsheets/d/1RLtRn7gojLLLxrsm9mSFRLy7eGmCYRn5bgsoCmuju7Y/edit)")

    if "delivery_reschedule_success" in st.session_state:
        st.success(st.session_state.pop("delivery_reschedule_success"))

    del_sheet = None
    picking_delivery_sheet = None
    legacy_delivery_rows = []
    try:
        del_sheet = sh.worksheet("Delivery_schedule")
        legacy_delivery_rows = get_cached_data("delivery", del_sheet, include_row_index=True)
    except gspread.exceptions.WorksheetNotFound:
        st.info("ยังไม่มีแท็บ Delivery_schedule รายการ Picking Request ที่สร้างแล้วจะแสดงด้านล่างได้")
    except Exception as error:
        st.error(f"❌ โหลดรายการ Delivery เดิมไม่สำเร็จ: {error}")
    delivery_rows = [dict(row, _Delivery_Source="legacy") for row in legacy_delivery_rows]
    try:
        picking_delivery_sheet, picking_delivery_rows = load_picking_requests(sh)
        if picking_delivery_rows:
            current_catalog = build_picking_catalog(
                get_cached_data("main", sheet), get_cached_data("packing", packing_records_sheet),
            )
            delivery_rows = merge_delivery_requests(legacy_delivery_rows, picking_delivery_rows, current_catalog)
    except Exception as error:
        st.warning(f"ยังโหลด Picking Request ไม่สำเร็จ รายการ Delivery เดิมยังแสดงตามข้อมูลที่โหลดได้: {error}")
    df_del = pd.DataFrame(delivery_rows) if delivery_rows else pd.DataFrame()

    if not df_del.empty:
        df_del.columns = [str(c).strip() for c in df_del.columns]

        vol_col = next((c for c in df_del.columns if "volume" in str(c).lower()), None)
        fg_col = next((c for c in df_del.columns if "available fg" in str(c).lower()), None)
        wip_col = next((c for c in df_del.columns if "available wip" in str(c).lower()), None)

        if vol_col and fg_col:
            def check_readiness(row):
                if row.get("_Delivery_Source") == "picking":
                    return row["Status (ความพร้อม)"]
                vol = clean_quantity(row.get(vol_col, 0))
                onhand = clean_quantity(row.get(fg_col, 0))
                wip_val = str(row.get(wip_col, "")).replace(",", "") if wip_col else "0"
                wip_numbers = re.findall(r"[-+]?\d*\.\d+|\d+", wip_val)
                total_wip = sum([float(n) for n in wip_numbers]) if wip_numbers else 0.0

                if vol > 0 and onhand >= vol: return "✅ พร้อมส่ง"
                elif vol > 0 and total_wip > 0: return "📦 รอแพ็ค"
                elif vol > 0: return "📥 รอรับเข้าสินค้า"
                else: return "➖ N/A"
            df_del["Status (ความพร้อม)"] = df_del.apply(check_readiness, axis=1)

        date_col = find_delivery_column(df_del.columns, "Expected Shipping Date")
        if date_col not in df_del.columns:
            date_col = None

        delivery_internal_columns = ["Original_Row_Index", "Is_Unscheduled", "Parsed_Delivery_Date", "_Delivery_Source",
                                     "Request ID", "Line ID", "Created At", "Export Name", "Source Categories",
                                     "Required Category", "Note", "Request Demand", "FG Shortfall"]
        if date_col:
            df_del[date_col] = df_del[date_col].astype(str).str.strip()
            df_del["Is_Unscheduled"] = df_del[date_col].str.casefold().isin(DELIVERY_PENDING_VALUES)
            df_del["Parsed_Delivery_Date"] = df_del[date_col].map(parse_delivery_date)
            df_del = df_del.sort_values(
                by=["Is_Unscheduled", "Parsed_Delivery_Date"],
                ascending=True, kind="stable", na_position="last",
            )

        scheduled_rows = df_del[df_del["Parsed_Delivery_Date"].notna()] if date_col else df_del.iloc[0:0]
        with st.container(key="delivery_overview"):
            render_section_heading(
                "แผนจัดส่งสินค้า", "ดูยอดส่งและความพร้อมของรายการที่มีวันจัดส่ง พร้อมติดตามรายการรอยืนยัน",
                eyebrow="DELIVERY PLANNING", tone="blue",
            )
            summary_cols = st.columns(4)
            with summary_cols[0]:
                st.metric("วันจัดส่งที่กำหนดแล้ว", f"{scheduled_rows['Parsed_Delivery_Date'].nunique():,}" if date_col else "—", help="จำนวนวันส่งที่มีวันที่ถูกต้อง รวมวันที่ผ่านมาแล้ว")
            with summary_cols[1]:
                st.metric("ยอดส่งตามกำหนด (Volume)", f"{scheduled_rows[vol_col].apply(clean_quantity).sum():,.2f}" if date_col and vol_col else "—", help="ยอดรวม Volume ของรายการที่มีวันจัดส่ง ไม่รวม TBC")
            with summary_cols[2]:
                ready_count = int(scheduled_rows["Status (ความพร้อม)"].eq("✅ พร้อมส่ง").sum()) if "Status (ความพร้อม)" in scheduled_rows else None
                st.metric("พร้อมส่งตามกำหนด · รายการ", f"{ready_count:,}" if date_col and ready_count is not None else "—")
            with summary_cols[3]:
                st.metric("รอยืนยันวันส่ง (TBC) · รายการ", f"{int(df_del['Is_Unscheduled'].sum()):,}" if date_col else "—")

        def display_delivery_table(frame):
            display_frame = format_df_quantity(frame)
            if "Status (ความพร้อม)" in display_frame.columns:
                status_styles = {
                    "✅ พร้อมส่ง": "background-color: #ecfdf5; color: #047857; font-weight: 600",
                    "📦 รอแพ็ค": "background-color: #fffbeb; color: #92400e; font-weight: 600",
                    "📥 รอรับเข้าสินค้า": "background-color: #eff6ff; color: #1d4ed8; font-weight: 600",
                    "➖ N/A": "background-color: #f1f5f9; color: #475569",
                }
                display_frame = display_frame.style.map(lambda value: status_styles.get(str(value), ""), subset=["Status (ความพร้อม)"])
            st.dataframe(
                display_frame, width="stretch", hide_index=True,
                column_config={
                    "Cultivar Name": st.column_config.TextColumn("สายพันธุ์ / สินค้า", width="medium"),
                    "Grade": st.column_config.TextColumn("เกรด", width="small"),
                    "Status (ความพร้อม)": st.column_config.TextColumn("ความพร้อม", width="medium"),
                },
            )

        with st.container(key="delivery_reschedule"):
            with st.expander("✏️ เลื่อน/ปรับเปลี่ยนกำหนดการจัดส่ง (Reschedule)"):
                with st.form("reschedule_form"):
                    st.caption("เลือกได้หลายรายการ จากนั้นระบุวันส่งใหม่ หรือเลือก TBC หากยังไม่มีกำหนด")
                    reschedule_opts = []
                    reschedule_map = {}
                    for _, r in df_del.iterrows():
                        row_idx = r["Original_Row_Index"]
                        curr_dt = str(r.get(date_col, "")) if date_col else ""
                        vol = r.get(vol_col if vol_col else "Volume", 0)

                        cultivar_column = find_delivery_column(df_del.columns, "Cultivar Name")
                        export_column = find_delivery_column(df_del.columns, "Export Name")
                        item_name = picking_text(r.get(cultivar_column)) or picking_text(r.get(export_column)) or "Unknown Item"
                        grade_column = find_delivery_column(df_del.columns, "Grade")
                        item_grade = picking_text(r.get(grade_column)) or "-"

                        try: vol_disp = f"{float(vol):,.2f}" if float(vol) % 1 != 0 else f"{int(vol):,}"
                        except: vol_disp = str(vol)

                        disp_txt = f"บรรทัด {row_idx} | {item_name} (Grade: {item_grade}) | วันเดิม: {curr_dt} | Volume: {vol_disp}"
                        row_source = picking_text(r.get("_Delivery_Source")) or "legacy"
                        if row_source == "picking":
                            disp_txt += f" | Request: {r['Request ID']}"
                        reschedule_opts.append(disp_txt)
                        reschedule_map[disp_txt] = {"row_idx": row_idx, "curr_date": curr_dt, "source": row_source,
                                                    "line_id": picking_text(r.get("Line ID")) if row_source == "picking" else "",
                                                    "request_id": picking_text(r.get("Request ID")) if row_source == "picking" else ""}

                    selected_items = st.multiselect(
                        "เลือกรายการที่ต้องการแก้ไข (เลือกกี่รายการก็ได้)", reschedule_opts,
                        key="delivery_reschedule_items",
                    )
                    mark_tbc = st.checkbox(
                        "TBC — ยกเลิกวันส่งเดิม / รอยืนยันกำหนดใหม่",
                        key="delivery_reschedule_tbc",
                        help="เลือกเมื่อยกเลิกวันส่งเดิมและยังไม่มีวันใหม่ ระบบจะบันทึก TBC แทนวันที่ที่เลือกด้านล่าง โดยเก็บรายการสินค้าไว้",
                    )
                    new_date = st.date_input(
                        "กำหนดการใหม่ (New Expected Date)", value=date.today(),
                        key="delivery_reschedule_date",
                        help="ใช้วันที่นี้เมื่อไม่ได้เลือก TBC",
                    )

                    if st.form_submit_button("💾 บันทึกกำหนดการใหม่"):
                        if not selected_items:
                            st.error("❌ กรุณาเลือกรายการอย่างน้อย 1 รายการ")
                        elif not date_col:
                            st.error("❌ ไม่พบคอลัมน์วันที่จัดส่ง จึงยังบันทึกกำหนดการใหม่ไม่ได้")
                        elif not mark_tbc and new_date is None:
                            st.error("❌ กรุณาระบุวันจัดส่งใหม่ หรือเลือก TBC หากยังไม่มีกำหนด")
                        else:
                            try:
                                new_value = "TBC" if mark_tbc else new_date.isoformat()
                                targets = []
                                success_rows = []
                                for selected in selected_items:
                                    item = reschedule_map[selected]
                                    is_picking = item["source"] == "picking"
                                    target_sheet = picking_delivery_sheet if is_picking else del_sheet
                                    if target_sheet is None:
                                        raise ValueError("ไม่พบแหล่งข้อมูลของรายการที่เลือก กรุณาโหลดหน้าใหม่")
                                    target_row = int(item["row_idx"])
                                    targets.append({
                                        "worksheet": target_sheet, "row_idx": target_row,
                                        "date_header": "Expected Shipping Date" if is_picking else date_col,
                                        "line_id": item["line_id"], "request_id": item["request_id"],
                                        "expected_date": item["curr_date"],
                                    })
                                    success_rows.append(f"{target_sheet.title} แถว {target_row}")
                                update_delivery_dates(sh, targets, new_value)
                            except Exception as error:
                                st.error(f"❌ บันทึกกำหนดการไม่สำเร็จ กรุณาลองใหม่: {error}")
                            else:
                                get_cached_data.clear()
                                get_cached_picking_rows.clear()
                                schedule_label = "TBC (รอยืนยันวันจัดส่งใหม่)" if mark_tbc else new_date.strftime("%d/%m/%Y")
                                st.session_state["delivery_reschedule_success"] = (
                                    f"✅ อัปเดตกำหนดส่งเป็น {schedule_label} สำหรับ {', '.join(success_rows)} เรียบร้อยแล้ว"
                                )
                                st.rerun()

        render_section_heading(
            "กำหนดการจัดส่ง", "เรียงวันส่งจากก่อนถึงหลัง พร้อมยอดรวมส่งของแต่ละวัน",
            eyebrow="DELIVERY SCHEDULE", tone="blue",
        )
        with st.container(key="delivery_list"):
            if date_col:
                dated_df = df_del[df_del["Parsed_Delivery_Date"].notna()]
                invalid_df = df_del[~df_del["Is_Unscheduled"] & df_del["Parsed_Delivery_Date"].isna()]
                tbc_df = df_del[df_del["Is_Unscheduled"]]

                for parsed_d, sub_df in dated_df.groupby("Parsed_Delivery_Date", sort=True):
                    sub_display = sub_df.drop(columns=delivery_internal_columns, errors="ignore")
                    d_date = parsed_d.strftime("%d/%m/%Y")
                    thai_days = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]
                    thai_day_str = thai_days[parsed_d.weekday()]
                    diff_days = (parsed_d.date() - date.today()).days
                    if diff_days < 0:
                        day_tone, diff_text = "overdue", f"เลยกำหนด {abs(diff_days)} วัน"
                    elif diff_days == 0:
                        day_tone, diff_text = "today", "จัดส่งวันนี้"
                    elif diff_days <= 3:
                        day_tone, diff_text = "soon", f"อีก {diff_days} วัน"
                    else:
                        day_tone, diff_text = "scheduled", f"อีก {diff_days} วัน"
                    daily_volume_display = (
                        f"{sub_df[vol_col].apply(clean_quantity).sum():,.2f}"
                        if vol_col else "ไม่พบคอลัมน์ Volume"
                    )
                    with st.container(border=True):
                        st.markdown(
                            f'<div class="delivery-day-header portal-day portal-day--{day_tone}">'
                            '<div class="portal-day__main"><div class="portal-day__eyebrow">กำหนดส่ง</div>'
                            f'<div class="portal-day__title">กำหนดส่ง: วัน{thai_day_str}ที่ {d_date}</div>'
                            f'<div class="portal-day__meta"><span class="portal-day__badge">{diff_text}</span><span>{len(sub_df):,} รายการ</span></div></div>'
                            '<div class="delivery-day-volume portal-day-volume">'
                            '<div class="portal-day-volume__label">ยอดรวมส่ง (Volume)</div>'
                            f'<div class="portal-day-volume__value">{daily_volume_display}</div>'
                            '</div></div>', unsafe_allow_html=True,
                        )
                        display_delivery_table(sub_display)

                if not invalid_df.empty:
                    with st.container(border=True):
                        st.warning("⚠️ รายการต่อไปนี้มีรูปแบบวันที่ไม่ถูกต้อง กรุณาระบุวันจัดส่งใหม่หรือเลือก TBC")
                        display_delivery_table(invalid_df.drop(columns=delivery_internal_columns, errors="ignore"))

                if not tbc_df.empty:
                    with st.container(key="delivery_pending", border=True):
                        st.markdown(
                            '<div class="portal-day portal-day--pending"><div class="portal-day__main">'
                            '<div class="portal-day__eyebrow">PENDING CONFIRMATION</div>'
                            '<div class="portal-day__title">TBC · รอยืนยันวันจัดส่งใหม่</div>'
                            f'<div class="portal-day__meta"><span class="portal-day__badge">ยังไม่มีกำหนดส่ง</span><span>{len(tbc_df):,} รายการ</span></div></div></div>',
                            unsafe_allow_html=True,
                        )
                        st.caption("กำหนดวันใหม่ได้ในส่วนปรับเปลี่ยนกำหนดการด้านบน รายการจะกลับไปเรียงตามวันส่งโดยอัตโนมัติ")
                        display_delivery_table(tbc_df.drop(columns=delivery_internal_columns, errors="ignore"))
            else:
                st.info("ℹ️ ไม่พบคอลัมน์วันที่จัดส่งในแท็บนี้")
                display_delivery_table(df_del.drop(columns=delivery_internal_columns, errors="ignore"))
    else:
        st.info("ยังไม่มีข้อมูลในแท็บ 'Delivery_schedule'")

elif menu == "📋 Picking Request":
    render_portal_ui_styles()
    with st.container(key="picking_header"):
        st.title("Picking Request")
        st.caption("สร้างคำขอจัดส่งหลายสายพันธุ์จาก WIP / FG และเทียบความพร้อมของ FG กับสต็อกปัจจุบัน")
    st.info("คำขอนี้ไม่จองหรือตัดสต็อก สินค้าที่ส่งต้องเป็น FG เท่านั้น ส่วน WIP ใช้แสดงว่าต้องรอ Process")

    def reset_picking_draft():
        next_epoch = int(st.session_state.get("picking_draft_epoch", 0)) + 1
        for state_key in list(st.session_state):
            if state_key.startswith("picking_draft_"):
                del st.session_state[state_key]
        st.session_state["picking_draft_epoch"] = next_epoch

    managed_result = st.session_state.pop("picking_manage_draft_result", None)
    if managed_result and managed_result["request_id"] == st.session_state.get("picking_draft_request_id"):
        reset_picking_draft()

    epoch = int(st.session_state.get("picking_draft_epoch", 0))
    if "picking_draft_request_id" not in st.session_state:
        created_at = datetime.now(ZoneInfo("Asia/Bangkok"))
        st.session_state["picking_draft_request_id"] = f"PR-{created_at:%Y%m%d}-{uuid4().hex[:12].upper()}"
        st.session_state["picking_draft_created_at"] = created_at.isoformat(timespec="seconds")
    request_id = st.session_state["picking_draft_request_id"]
    saved_result = st.session_state.get("picking_draft_saved")
    draft_saved = bool(saved_result)
    request_col, reset_col = st.columns([3, 1])
    with request_col:
        st.caption(f"หมายเลขคำขอ: {request_id}")
    with reset_col:
        st.button("เริ่มคำขอใหม่", key="picking_new_request", on_click=reset_picking_draft, width="stretch")
    if draft_saved:
        st.success(f"บันทึก {saved_result['request_id']} จำนวน {saved_result['line_count']:,} รายการแล้ว สามารถติดตามและแก้วันส่งใน Delivery Issue")

    catalog = build_picking_catalog(
        get_cached_data("main", sheet),
        get_cached_data("packing", packing_records_sheet),
    )
    catalog_by_name = {item["cultivar"]: item for item in catalog.values()}
    cultivar_options = sorted(catalog_by_name, key=str.casefold)
    # Widget state is cleaned when another module renders. Keep the draft in
    # ordinary session state and restore missing widgets before creating them.
    draft_fields = st.session_state.get("picking_draft_fields", {})
    for field_name in ["consignee", "true_consignee", "project_owner", "type", "tbc", "date", "cultivars"]:
        widget_key = f"picking_draft_{epoch}_{field_name}"
        if widget_key not in st.session_state and field_name in draft_fields:
            st.session_state[widget_key] = draft_fields[field_name]
    st.session_state.setdefault(f"picking_draft_{epoch}_tbc", True)
    st.session_state.setdefault(f"picking_draft_{epoch}_date", datetime.now(ZoneInfo("Asia/Bangkok")).date())
    selection_key = f"picking_draft_{epoch}_cultivars"
    restored_selection = st.session_state.get(selection_key, [])
    valid_selection = [name for name in restored_selection if name in catalog_by_name]
    if valid_selection != restored_selection:
        st.session_state[selection_key] = valid_selection
    render_section_heading(
        "สร้างคำขอจัดส่ง FG", "กำหนดผู้รับและวันส่งร่วมกัน แล้วแก้จำนวน เกรด และผู้รับแยกแต่ละรายการได้",
        eyebrow="NEW PICKING REQUEST", tone="blue",
    )
    with st.container(key="picking_form", border=True):
        common_left, common_right = st.columns(2)
        with common_left:
            default_consignee = st.text_input(
                "Consignee เริ่มต้น", key=f"picking_draft_{epoch}_consignee", disabled=draft_saved,
                help="ใช้กับรายการที่เว้น Consignee ในตารางไว้ สามารถแก้แยกแต่ละรายการได้",
            )
        with common_right:
            default_true_consignee = st.text_input(
                "True Consignee เริ่มต้น", key=f"picking_draft_{epoch}_true_consignee", disabled=draft_saved,
                help="ใช้กับรายการที่เว้น True Consignee ในตารางไว้ สามารถแก้แยกแต่ละรายการได้",
            )
        owner_col, type_col = st.columns(2)
        with owner_col:
            default_project_owner = st.text_input("Project Owner (ค่าเริ่มต้น)", key=f"picking_draft_{epoch}_project_owner", disabled=draft_saved)
        with type_col:
            default_type = st.selectbox("Type ผลิตภัณฑ์ (ค่าเริ่มต้น)", PICKING_TYPES, index=None, placeholder="เลือก Type ผลิตภัณฑ์", key=f"picking_draft_{epoch}_type", disabled=draft_saved)
        st.caption("Supplier ดึงจาก WIP / FG ของสายพันธุ์ที่เลือกโดยอัตโนมัติ หากมีหลาย Supplier จะแสดงชื่อทั้งหมด และไม่มีการจอง Lot")
        date_left, date_right = st.columns(2)
        with date_left:
            request_tbc = st.checkbox(
                "TBC — ยังไม่กำหนดวันจัดส่ง", key=f"picking_draft_{epoch}_tbc",
                disabled=draft_saved,
            )
        with date_right:
            request_date = st.date_input(
                "วันจัดส่ง", value="today",
                key=f"picking_draft_{epoch}_date", disabled=draft_saved or request_tbc,
            )
        shipping_date = "TBC" if request_tbc else request_date.isoformat() if request_date else ""
        selected_cultivars = st.multiselect(
            "เลือกสายพันธุ์จาก WIP / FG", cultivar_options,
            format_func=lambda name: f"{name} · {' / '.join(catalog_by_name[name]['sources'])}",
            key=f"picking_draft_{epoch}_cultivars", disabled=draft_saved,
        )
        st.session_state["picking_draft_fields"] = {
            "consignee": default_consignee, "true_consignee": default_true_consignee,
            "project_owner": default_project_owner, "type": default_type,
            "tbc": request_tbc, "date": request_date, "cultivars": list(selected_cultivars),
        }
        st.caption("กรอกเกรด FG ที่ต้องส่ง เช่น BIG / SMALL / MIX / F และจำนวนที่ต้องการ ผู้รับที่เว้นว่างจะใช้ค่าเริ่มต้นด้านบน")
        editor_columns = ["Cultivar Name", "Grade", "Volume", "Consignee", "True Consignee", "Project Owner", "Type", "Note"]
        selected_tuple = tuple(selected_cultivars)
        active_editor_key = f"picking_draft_{epoch}_editor_{st.session_state.get('picking_draft_editor_version', 0)}"
        editor_needs_restore = bool(selected_cultivars) and active_editor_key not in st.session_state
        selection_changed = st.session_state.get("picking_draft_selection") != selected_tuple
        if selection_changed or editor_needs_restore:
            previous_lines = st.session_state.get("picking_draft_lines", [])
            editor_rows = [dict(row) for row in previous_lines if row.get("Cultivar Name") in selected_cultivars]
            existing_cultivars = {row.get("Cultivar Name") for row in editor_rows}
            for cultivar in selected_cultivars:
                if selection_changed and cultivar not in existing_cultivars:
                    fg_grades = list(catalog_by_name[cultivar]["fg_by_grade"])
                    editor_rows.append({
                        "Cultivar Name": cultivar, "Grade": fg_grades[0] if len(fg_grades) == 1 else "",
                        "Volume": 0.0, "Consignee": "", "True Consignee": "", "Project Owner": "", "Type": "", "Note": "",
                    })
            st.session_state["picking_draft_editor_input"] = pd.DataFrame(editor_rows, columns=editor_columns)
            st.session_state["picking_draft_editor_version"] = int(st.session_state.get("picking_draft_editor_version", 0)) + 1
            st.session_state["picking_draft_selection"] = selected_tuple
        edited_records = []
        if selected_cultivars:
            edited_frame = st.data_editor(
                st.session_state["picking_draft_editor_input"], width="stretch", hide_index=True,
                num_rows="dynamic", disabled=draft_saved,
                key=f"picking_draft_{epoch}_editor_{st.session_state['picking_draft_editor_version']}",
                column_config={
                    "Cultivar Name": st.column_config.SelectboxColumn("สายพันธุ์", options=selected_cultivars, required=True, width="medium"),
                    "Grade": st.column_config.TextColumn("เกรด FG ที่ต้องส่ง", required=True, width="small"),
                    "Volume": st.column_config.NumberColumn("จำนวนส่ง (Volume)", min_value=0.0, step=0.01, format="%.2f", required=True),
                    "Consignee": st.column_config.TextColumn("Consignee", help="เว้นว่างเพื่อใช้ค่าเริ่มต้นด้านบน", width="medium"),
                    "True Consignee": st.column_config.TextColumn("True Consignee", help="เว้นว่างเพื่อใช้ค่าเริ่มต้นด้านบน", width="medium"),
                    "Project Owner": st.column_config.TextColumn("Project Owner", help="เว้นว่างเพื่อใช้ค่าเริ่มต้นด้านบน", width="medium"),
                    "Type": st.column_config.TextColumn("Type ผลิตภัณฑ์", help="เว้นว่างเพื่อใช้ค่าเริ่มต้นด้านบน", width="medium"),
                    "Note": st.column_config.TextColumn("หมายเหตุ", width="medium"),
                },
            )
            edited_records = edited_frame.to_dict("records")
            st.session_state["picking_draft_lines"] = edited_records
            st.caption("เพิ่มแถวได้เมื่อต้องการสายพันธุ์เดิมต่างเกรด หากสายพันธุ์และเกรดซ้ำ ระบบจะเทียบยอดรวมภายใน Request เดียวกัน")
        else:
            st.caption("เลือกอย่างน้อย 1 สายพันธุ์เพื่อเริ่มกรอกรายการ")

        if edited_records:
            try:
                preview_lines = prepare_picking_lines(
                    edited_records, catalog, default_consignee or "—", default_true_consignee or "—", default_project_owner or "—", default_type or "—",
                )
                preview_rows = project_picking_rows([
                    dict(line, **{"Request ID": request_id, "Expected Shipping Date": shipping_date})
                    for line in preview_lines
                ], catalog)
            except ValueError:
                st.caption("กรอกเกรดและจำนวนที่มากกว่า 0 ให้ครบเพื่อดูความพร้อมของ FG")
            else:
                st.markdown("**ความพร้อมตาม On hand ปัจจุบัน**")
                st.dataframe(
                    format_df_quantity(pd.DataFrame(preview_rows)), width="stretch", hide_index=True,
                    column_order=["Cultivar Name", "Grade", "Volume", "Supplier Name", "Project Owner", "Type", "Request Demand", "Available FG", "Available WIP", "FG Shortfall", "Status (ความพร้อม)"],
                    column_config={
                        "Cultivar Name": st.column_config.TextColumn("สายพันธุ์", width="medium"),
                        "Grade": "เกรด FG", "Volume": "จำนวนส่ง", "Request Demand": "ยอดรวม Request / เกรด",
                        "Available FG": "FG ตรงเกรด", "Available WIP": "WIP รอ Process",
                        "FG Shortfall": "FG ที่ยังขาด", "Status (ความพร้อม)": "ความพร้อม",
                    },
                )
                st.caption("WIP ไม่ถูกนับเป็น FG พร้อมส่ง และไม่สมมติผลผลิตหลัง Process ความพร้อมของแต่ละ Request เทียบ On hand แยกกันโดยไม่กันสต็อก")

        if st.button("สร้าง Picking Request", key="picking_save_request", type="primary", disabled=draft_saved or not selected_cultivars):
            try:
                if not shipping_date:
                    raise ValueError("กรุณาระบุวันจัดส่งหรือเลือก TBC")
                lines_to_save = prepare_picking_lines(edited_records, catalog, default_consignee, default_true_consignee, default_project_owner, default_type)
                with st.spinner("กำลังบันทึก Picking Request..."):
                    result = save_picking_request(
                        sh, request_id, st.session_state["picking_draft_created_at"], lines_to_save, shipping_date,
                    )
            except Exception as error:
                st.error(f"ยังยืนยันการบันทึกไม่ได้: {error}")
                st.caption("คำขอยังคงใช้หมายเลขเดิม หากการเชื่อมต่อขัดข้องให้กดบันทึกซ้ำโดยไม่แก้ข้อมูล ระบบจะตรวจคำขอเดิมก่อนเพื่อป้องกันรายการซ้ำ")
            else:
                st.session_state["picking_draft_saved"] = result
                st.session_state["picking_history_request"] = request_id
                get_cached_data.clear()
                get_cached_picking_rows.clear()
                st.rerun()

    render_section_heading(
        "ติดตาม Picking Request", "แสดงคำขอที่สร้างแล้ว พร้อมเทียบสต็อก FG ปัจจุบัน ส่วนรายการ Delivery เดิมยังคงแสดงใน Delivery Issue",
        eyebrow="REQUEST HISTORY", tone="blue",
    )
    if "picking_manage_flash" in st.session_state:
        st.success(st.session_state.pop("picking_manage_flash"))
    try:
        _, picking_history_rows = load_picking_requests(sh)
        projected_history = project_picking_rows(picking_history_rows, catalog)
    except Exception as error:
        st.error(f"โหลดประวัติ Picking Request ไม่สำเร็จ: {error}")
    else:
        if projected_history:
            request_ids = list(dict.fromkeys(str(row.get("Request ID", "")) for row in projected_history if row.get("Request ID")))
            request_ids.reverse()
            history_id = st.selectbox("เลือก Request เพื่อดูรายละเอียด", request_ids, key="picking_history_request", disabled=bool(st.session_state.get("picking_manage")))
            selected_history = [row for row in projected_history if str(row.get("Request ID", "")) == history_id]
            st.dataframe(
                format_df_quantity(pd.DataFrame(selected_history)), width="stretch", hide_index=True,
                column_order=["Request ID", "Cultivar Name", "Grade", "Volume", "Consignee", "True Consignee", "Project Owner", "Type", "Supplier Name", "Expected Shipping Date", "Available FG", "Available WIP", "FG Shortfall", "Status (ความพร้อม)", "Required Category", "Source Categories", "Note"],
                column_config={
                    "Original_Row_Index": None, "Line ID": None,
                    "Cultivar Name": "สายพันธุ์", "Grade": "เกรด FG", "Volume": "จำนวนส่ง",
                    "Expected Shipping Date": "วันจัดส่ง", "Available FG": "FG ตรงเกรด",
                    "Available WIP": "WIP รอ Process", "FG Shortfall": "FG ที่ยังขาด",
                    "Required Category": "ประเภทที่จัดส่ง", "Source Categories": "แหล่งอ้างอิง",
                },
            )
            edit_col, delete_col = st.columns(2)
            with edit_col:
                edit_clicked = st.button("แก้ไข Request", key="picking_edit_request", disabled=bool(st.session_state.get("picking_manage")), width="stretch")
            with delete_col:
                delete_clicked = st.button("ลบ Request", key="picking_delete_request", disabled=bool(st.session_state.get("picking_manage")), width="stretch")
            if edit_clicked or delete_clicked:
                try:
                    current_rows = [row for row in read_picking_records(sh.worksheet(PICKING_SHEET_NAME))
                                    if picking_text(row.get("Request ID")) == history_id]
                    if not current_rows:
                        raise ValueError("ไม่พบ Request นี้แล้ว กรุณาโหลดข้อมูลใหม่")
                    edit_columns = ["Line ID", "Cultivar Name", "Grade", "Volume", "Consignee", "True Consignee", "Project Owner", "Type", "Expected Shipping Date", "Note"]
                    editor_rows = [{column: row.get(column, "") for column in edit_columns} for row in current_rows]
                    for row in editor_rows:
                        row["Volume"] = picking_quantity(row["Volume"])
                    st.session_state["picking_manage"] = {
                        "request_id": history_id, "mode": "edit" if edit_clicked else "delete",
                        "expected": current_rows, "token": uuid4().hex[:12],
                        "editor_input": pd.DataFrame(editor_rows, columns=edit_columns),
                    }
                    get_cached_picking_rows.clear()
                except Exception as error:
                    st.error(f"เปิดข้อมูล Request ไม่สำเร็จ: {error}")
                else:
                    st.rerun()
        else:
            st.caption("ยังไม่มี Picking Request เมื่อสร้างคำขอใหม่ รายการจะแสดงที่นี่และใน Delivery Issue")
    render_picking_manager(sh, catalog)

# ==========================================
# หน้าที่ 2: Stock Summary & Details
# ==========================================
elif menu == "📊 Stock Summary & Details":
    render_portal_ui_styles()
    with st.container(key="stock_header"):
        st.title("Stock Summary & Details")
        st.caption("ติดตามสต็อกคงเหลือ สัดส่วน WIP / FG และรายละเอียดสินค้าแต่ละ Lot")

    data_main = get_cached_data("main", sheet)
    df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()

    records_data = get_cached_data("packing", packing_records_sheet)
    df_packing_fg = pd.DataFrame()
    if records_data:
        df_pack_raw = pd.DataFrame(records_data)
        converted_rows = []
        for _, row in df_pack_raw.iterrows():
            converted_rows.append({
                "Cultivar Name": row.get("Cultivar Name", ""), "Supplier Name": row.get("Supplier Name", ""), "Supplier Lot. No.": row.get("Source WIP Lot", ""),
                "Export Name": row.get("Export Name", ""), "THAI STICK Lot. No.": row.get("Generated FG Lot No.", ""),
                "Category": "FG (Finished Goods)", "Grade": row.get("Grade", ""), "Receive date": row.get("Packing Date", ""),
                "Quantity": clean_quantity(row.get("Quantity", 0)), "Remark Qty": "", "Inspection Req.": "No", "IR Req.": "No",
                "Shipment date": "To be decided", "Type": "In House Product", "Status": "Packing done", "Packing date": row.get("Packing Date", ""),
            })
        df_packing_fg = pd.DataFrame(converted_rows)

    if not df_main.empty and not df_packing_fg.empty: df = pd.concat([df_main, df_packing_fg], ignore_index=True)
    elif not df_main.empty: df = df_main
    elif not df_packing_fg.empty: df = df_packing_fg
    else: df = pd.DataFrame()

    if not df.empty:
        df["Clean_Qty"] = df["Quantity"].apply(clean_quantity) if "Quantity" in df.columns else 0.0
        filtered_df = df.copy()

        if "Status" in filtered_df.columns:
            filtered_df = filtered_df[~filtered_df["Status"].astype(str).str.strip().str.lower().eq("delivered")]
        filtered_df = filtered_df[filtered_df["Clean_Qty"] > 0]

        if "Category" in filtered_df.columns:
            wip_stock_df = filtered_df[filtered_df["Category"].astype(str).str.contains("WIP", case=False, na=False)]
            fg_stock_df = filtered_df[filtered_df["Category"].astype(str).str.contains("FG", case=False, na=False)]
        else:
            wip_stock_df = pd.DataFrame(); fg_stock_df = pd.DataFrame()

        total_wip = wip_stock_df["Clean_Qty"].sum()
        total_fg = fg_stock_df["Clean_Qty"].sum()
        total_on_hand = total_wip + total_fg

        render_section_heading(
            "ภาพรวมสต็อกคงเหลือ", "ยอดคงเหลือจากสินค้าในคลังและรายการ Packing รวมเฉพาะสินค้าที่ยังมีสต็อก",
            eyebrow="ON-HAND STOCK", tone="slate",
        )
        with st.container(key="stock_overview"):
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                st.metric("สต็อกคงเหลือทั้งหมด", f"{total_on_hand:,.2f} g")
            with col_m2:
                st.metric("งานระหว่างกระบวนการ · WIP", f"{total_wip:,.2f} g")
            with col_m3:
                st.metric("สินค้าสำเร็จรูป · FG", f"{total_fg:,.2f} g")

        wip_share = float(total_wip / total_on_hand * 100) if total_on_hand else 0.0
        fg_share = float(total_fg / total_on_hand * 100) if total_on_hand else 0.0
        st.markdown(
            '<div class="portal-balance"><div class="portal-balance__labels">'
            f'<strong class="portal-balance__wip">WIP · {wip_share:.1f}%</strong>'
            f'<strong class="portal-balance__fg">FG · {fg_share:.1f}%</strong></div>'
            f'<div class="portal-balance__track" role="img" aria-label="สัดส่วนสต็อก WIP {wip_share:.1f}% และ FG {fg_share:.1f}%">'
            f'<div class="portal-balance__segment--wip" style="width: {wip_share:.4f}%"></div>'
            f'<div class="portal-balance__segment--fg" style="width: {fg_share:.4f}%"></div>'
            '</div></div>', unsafe_allow_html=True,
        )

        from html import escape

        # ใช้สีและรูปแบบเฉพาะ Matrix เพื่อแยก WIP / FG อย่างชัดเจน
        st.markdown("""<style>
        .stock-matrix { --accent: #b45309; --deep: #92400e; --tint: #fffbeb;
            --soft: #fef3c7; --line: #fde68a; --rgb: 217, 119, 6;
            margin: 8px 0 28px; overflow: hidden; border: 1px solid var(--line);
            border-top: 4px solid var(--accent); border-radius: 18px;
            background: #fff; color: #1e293b; box-shadow: 0 6px 24px #0f172a06; }
        .stock-matrix--fg { --accent: #0f766e; --deep: #115e59; --tint: #f0fdfa;
            --soft: #ccfbf1; --line: #99f6e4; --rgb: 13, 148, 136; }
        .stock-matrix__header { display: flex; justify-content: space-between; gap: 24px;
            align-items: center; padding: 24px; background: linear-gradient(110deg, var(--tint), #fff); }
        .stock-matrix__badge { display: inline-block; padding: 5px 10px; border-radius: 6px;
            background: var(--soft); color: var(--deep); font-size: 11px;
            font-weight: 750; letter-spacing: .07em; }
        .stock-matrix h3 { margin: 10px 0 5px; padding: 0; color: #0f172a;
            font-size: 23px; font-weight: 750; line-height: 1.4; }
        .stock-matrix__description { margin: 0; color: #64748b; font-size: 13px; line-height: 1.7; }
        .stock-matrix__summary { flex-shrink: 0; text-align: right; }
        .stock-matrix__label { color: #475569; font-size: 12px; }
        .stock-matrix__quantity { color: var(--deep); font-size: 32px; font-weight: 800;
            line-height: 1.45; font-variant-numeric: tabular-nums; }
        .stock-matrix__quantity small { font-size: 14px; font-weight: 500; }
        .stock-matrix__counts { color: #64748b; font-size: 12px; }
        .stock-matrix__scroll { position: relative; overflow: auto; max-height: 480px;
            border-top: 1px solid var(--line); scrollbar-color: var(--accent) var(--tint); }
        .stock-matrix__scroll:focus-visible { outline: 3px solid var(--accent); outline-offset: -3px; }
        .stock-matrix table { width: 100%; border-collapse: separate; border-spacing: 0;
            margin: 0; font-size: 14px; font-variant-numeric: tabular-nums; }
        .stock-matrix th, .stock-matrix td { padding: 14px 18px; border: 0;
            border-right: 1px solid #f1f5f9; border-bottom: 1px solid #e2e8f0; }
        .stock-matrix thead th { position: sticky; top: 0; z-index: 2;
            min-width: 145px; max-width: 220px; background: var(--tint);
            color: var(--deep); text-align: center; font-size: 12px;
            font-weight: 700; overflow-wrap: anywhere; }
        .stock-matrix thead th.stock-matrix__cultivar { min-width: 200px; max-width: 260px; font-size: 20px; line-height: 1.5; }
        .stock-matrix th:first-child { position: sticky; left: 0; z-index: 1;
            min-width: 125px; background: #fff; color: #475569;
            text-align: left; white-space: nowrap; box-shadow: 3px 0 8px #0f172a05; }
        .stock-matrix thead th:first-child { z-index: 3; background: var(--tint); color: var(--deep); }
        .stock-matrix td { text-align: right; white-space: nowrap; color: #334155; }
        .stock-matrix td.stock-matrix__value { color: var(--deep); font-weight: 600; }
        .stock-matrix td.stock-matrix__zero { color: #64748b; }
        .stock-matrix .stock-matrix__row-total { background: var(--tint);
            color: var(--deep); font-weight: 750; }
        .stock-matrix tbody tr:hover td, .stock-matrix tbody tr:hover th { box-shadow: inset 0 0 0 999px #0f172a06; }
        .stock-matrix tfoot th, .stock-matrix tfoot td { border-bottom: 0;
            border-top: 2px solid var(--line); background: var(--soft); color: var(--deep); font-weight: 750; }
        .stock-matrix tfoot th:first-child { background: var(--soft); color: var(--deep); }
        .stock-matrix tfoot td:last-child { background: var(--accent); color: #fff; }
        .stock-matrix__footer { display: flex; justify-content: space-between; flex-wrap: wrap;
            gap: 8px 20px; padding: 12px 20px; background: #f8fafc;
            border-top: 1px solid #e2e8f0; color: #64748b; font-size: 11px; line-height: 1.7; }
        .stock-matrix__legend { display: inline-flex; align-items: center; gap: 7px; }
        .stock-matrix__swatch { width: 36px; height: 8px; border-radius: 4px;
            background: linear-gradient(90deg, rgba(var(--rgb), .04), rgba(var(--rgb), .24)); }
        .stock-matrix__empty { padding: 30px 24px; border-top: 1px solid var(--line);
            color: #64748b; text-align: center; font-size: 14px; }
        .stock-matrix__sr-only { position: absolute; width: 1px; height: 1px;
            padding: 0; overflow: hidden; clip-path: inset(50%); white-space: nowrap; }
        @media (max-width: 640px) {
            .stock-matrix__header { align-items: flex-start; flex-direction: column; gap: 16px; padding: 18px; }
            .stock-matrix__summary { text-align: left; }
            .stock-matrix h3 { font-size: 21px; }
            .stock-matrix__quantity { font-size: 28px; }
            .stock-matrix th, .stock-matrix td { padding: 12px 14px; }
            .stock-matrix th:first-child { min-width: 110px; }
            .stock-matrix__footer { padding: 12px 16px; }
        }
        </style>""", unsafe_allow_html=True)

        def render_matrix_table(target_df, stock_type):
            is_wip = stock_type == "WIP"
            theme = "wip" if is_wip else "fg"
            category_name = "WORK IN PROGRESS" if is_wip else "FINISHED GOODS"
            description = "สต็อกงานระหว่างกระบวนการ" if is_wip else "สต็อกสินค้าสำเร็จรูป"
            pivot_df = pd.DataFrame()
            active_grades = 0

            if not target_df.empty:
                matrix_data = target_df.copy()
                matrix_data["Cultivar Name"] = matrix_data["Cultivar Name"].fillna("").astype(str).str.strip().replace("", "ไม่ระบุชื่อสายพันธุ์")
                matrix_data["Grade"] = matrix_data["Grade"].fillna("").astype(str).str.strip().str.upper().replace("", "GENERAL")
                pivot_df = matrix_data.pivot_table(
                    index="Grade", columns="Cultivar Name", values="Clean_Qty",
                    aggfunc="sum", fill_value=0.0,
                )
                active_grades = len(pivot_df.index)
                standard_sizes = ["BIG", "SMALL", "MIX", "F"]
                other_sizes = [grade for grade in pivot_df.index if grade not in standard_sizes]
                pivot_df = pivot_df.reindex(standard_sizes + other_sizes, fill_value=0.0)

            section_total = float(pivot_df.to_numpy().sum()) if not pivot_df.empty else 0.0
            cultivar_count = len(pivot_df.columns)
            html = [
                f'<section class="stock-matrix stock-matrix--{theme}" aria-labelledby="matrix-{theme}-title">',
                '<header class="stock-matrix__header"><div>',
                f'<span class="stock-matrix__badge">{stock_type} · {category_name}</span>',
                f'<h3 id="matrix-{theme}-title">{stock_type} Matrix Summary</h3>',
                f'<p class="stock-matrix__description">{description} แยกตามสายพันธุ์และขนาด / เกรด</p></div>',
                f'<div class="stock-matrix__summary"><div class="stock-matrix__label">ยอดคงเหลือ {stock_type}</div>',
                f'<div class="stock-matrix__quantity">{section_total:,.2f} <small>g</small></div>',
                f'<div class="stock-matrix__counts">{cultivar_count:,} สายพันธุ์ &nbsp;·&nbsp; {active_grades:,} เกรดที่มีสต็อก</div></div></header>',
            ]
            if pivot_df.empty:
                html.append(f'<div class="stock-matrix__empty">ยังไม่มีสต็อก {stock_type} คงเหลือในขณะนี้</div></section>')
                st.markdown("".join(html), unsafe_allow_html=True)
                return

            html.extend([
                f'<div class="stock-matrix__scroll" role="region" tabindex="0" aria-label="ตารางสต็อก {stock_type} เลื่อนเพื่อดูสายพันธุ์ทั้งหมด">',
                f'<table><caption class="stock-matrix__sr-only">ปริมาณสต็อก {stock_type} แยกตามเกรดและสายพันธุ์ หน่วยกรัม</caption>',
                '<thead><tr><th scope="col">Size / Grade</th>',
            ])
            for cultivar in pivot_df.columns:
                html.append(f'<th scope="col" class="stock-matrix__cultivar">{escape(str(cultivar))}</th>')
            html.append('<th scope="col" class="stock-matrix__row-total">รวมตามเกรด (g)</th></tr></thead><tbody>')

            max_quantity = float(pivot_df.to_numpy().max())
            for grade, quantities in pivot_df.iterrows():
                html.append(f'<tr><th scope="row">{escape(str(grade))}</th>')
                for quantity in quantities:
                    if quantity > 0:
                        intensity = 0.04 + 0.20 * float(quantity) / max_quantity
                        html.append(f'<td class="stock-matrix__value" style="background: rgba(var(--rgb), {intensity:.3f})">{quantity:,.2f}</td>')
                    else:
                        html.append('<td class="stock-matrix__zero"><span aria-hidden="true">—</span><span class="stock-matrix__sr-only">0.00</span></td>')
                html.append(f'<td class="stock-matrix__row-total">{quantities.sum():,.2f}</td></tr>')

            html.append('</tbody><tfoot><tr><th scope="row">รวมทั้งหมด</th>')
            for quantity in pivot_df.sum(axis=0):
                html.append(f'<td>{quantity:,.2f}</td>')
            html.extend([
                f'<td>{section_total:,.2f}</td></tr></tfoot></table></div>',
                '<footer class="stock-matrix__footer">',
                '<span class="stock-matrix__legend"><span class="stock-matrix__swatch" aria-hidden="true"></span>สีเข้ม = ปริมาณมาก เทียบภายในแต่ละตาราง</span>',
                '<span>หน่วย: กรัม (g) &nbsp;·&nbsp; — = ไม่มีสต็อก &nbsp;·&nbsp; เลื่อนตารางเพื่อดูทุกสายพันธุ์</span>',
                '</footer></section>',
            ])
            st.markdown("".join(html), unsafe_allow_html=True)

        render_matrix_table(wip_stock_df, "WIP")
        render_matrix_table(fg_stock_df, "FG")

        st.divider()

        render_section_heading(
            "รายละเอียดสต็อกทั้งหมด", "ตรวจสอบยอดแต่ละเกรดและรายละเอียดสินค้าแต่ละ Lot แยกตามประเภท",
            eyebrow="STOCK DETAILS", tone="slate",
        )
        for stock_type, stock_df, stock_total, tone, title in (
            ("WIP", wip_stock_df, total_wip, "amber", "รายละเอียด WIP"),
            ("FG", fg_stock_df, total_fg, "teal", "รายละเอียด FG"),
        ):
            with st.container(key=f"stock_details_{stock_type.lower()}"):
                description = "งานระหว่างกระบวนการ · Work in Progress" if stock_type == "WIP" else "สินค้าสำเร็จรูป · Finished Goods"
                render_section_heading(title, description, eyebrow=f"{stock_type} INVENTORY", tone=tone)
                if not stock_df.empty:
                    render_grade_summary(stock_df, stock_total, stock_type)
                    st.caption(f"{len(stock_df):,} รายการ · ยอดรวมทุกเกรด · เลื่อนตารางเพื่อดูรายละเอียดทั้งหมด")
                    preferred_columns = ["Cultivar Name", "Grade", "Quantity", "THAI STICK Lot. No.", "Status", "Supplier Name"]
                    display_columns = [column for column in preferred_columns if column in stock_df.columns]
                    display_columns += [column for column in stock_df.columns if column not in preferred_columns and column != "Clean_Qty"]
                    st.dataframe(
                        format_df_quantity(stock_df), width="stretch", hide_index=True,
                        column_order=display_columns,
                        height=min(460, 38 * (len(stock_df) + 1) + 4), row_height=38,
                        column_config={
                            "Clean_Qty": None,
                            "Cultivar Name": st.column_config.TextColumn("สายพันธุ์", width="medium"),
                            "Grade": st.column_config.TextColumn("เกรด", width="small"),
                            "Quantity": st.column_config.TextColumn("จำนวน (g)", width="small"),
                            "Supplier Name": st.column_config.TextColumn("ผู้จัดหา", width="medium"),
                            "THAI STICK Lot. No.": st.column_config.TextColumn("THAI STICK Lot No.", width="medium"),
                        },
                    )
                else:
                    st.markdown(f'<div class="portal-empty">ยังไม่มีรายการสินค้า {stock_type} คงเหลือในขณะนี้</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="portal-empty">ยังไม่มีข้อมูลสต็อกในระบบ เมื่อบันทึกรับสินค้าแล้ว ยอดคงเหลือจะแสดงที่หน้านี้</div>', unsafe_allow_html=True)
# ==========================================
# หน้าที่ 3: Monthly Report
# ==========================================
elif menu == "📈 Monthly Report":
    st.title("📈 Monthly Report (รายงานสรุปผลการดำเนินงานรายเดือน)")
    st.markdown("เลือกเดือนและปีที่ต้องการตรวจสอบรายงานสต็อกเคลื่อนไหวและประวัติการส่งออก")

    today = date.today()
    col_y, col_m = st.columns(2)
    with col_y:
        year_options = [2026, 2025, 2024, 2023]
        default_year_idx = year_options.index(today.year) if today.year in year_options else 0
        selected_year = st.selectbox("เลือกปี (Year)", year_options, index=default_year_idx, key="rep_year")
        
    with col_m:
        months_dict = {
            "มกราคม (01)": 1, "กุมภาพันธ์ (02)": 2, "มีนาคม (03)": 3, "เมษายน (04)": 4,
            "พฤษภาคม (05)": 5, "มิถุนายน (06)": 6, "กรกฎาคม (07)": 7, "สิงหาคม (08)": 8,
            "กันยายน (09)": 9, "ตุลาคม (10)": 10, "พฤศจิกายน (11)": 11, "ธันวาคม (12)": 12,
        }
        default_month_idx = today.month - 1
        selected_month_label = st.selectbox("เลือกเดือน (Month)", list(months_dict.keys()), index=default_month_idx, key="rep_month")
        selected_month_num = months_dict[selected_month_label]

    st.divider()

    rep_tab1, rep_tab2 = st.tabs([
        "📊 1. Monthly On-Hand Stock Report (WIP & FG)",
        "🚚 2. Shipped Out Report & Reasons",
    ])

    def parse_robust_date(date_val):
        if not date_val or not str(date_val).strip(): return pd.NaT
        d_str = str(date_val).strip()
        try:
            dt = pd.to_datetime(d_str, errors='coerce')
            if pd.isna(dt): dt = pd.to_datetime(d_str, format='%d/%m/%Y', errors='coerce')
            return dt
        except: return pd.NaT

    shipped_data = get_cached_data("shipped", shipped_sheet)
    df_shipped = pd.DataFrame(shipped_data) if shipped_data else pd.DataFrame()
    if not df_shipped.empty:
        df_shipped.columns = [str(c).strip() for c in df_shipped.columns]
        df_shipped["Parsed_Date"] = df_shipped["Shipment Date"].apply(parse_robust_date)
        df_shipped["Clean_Shipped_Qty"] = df_shipped["Shipped Qty"].apply(clean_quantity)
        monthly_shipped = df_shipped[
            (df_shipped["Parsed_Date"].dt.year == selected_year) & 
            (df_shipped["Parsed_Date"].dt.month == selected_month_num)
        ].copy()
    else:
        monthly_shipped = pd.DataFrame()

    with rep_tab1:
        st.subheader(f"📅 รายงานความเคลื่อนไหวสต็อกประจำเดือน {selected_month_label} {selected_year}")

        main_data = get_cached_data("main", sheet)
        df_main_rep = pd.DataFrame(main_data) if main_data else pd.DataFrame()
        pack_data = get_cached_data("packing", packing_records_sheet)
        df_pack_rep = pd.DataFrame(pack_data) if pack_data else pd.DataFrame()

        events = []
        lot_addbacks = {}

        if not df_shipped.empty:
            for _, r in df_shipped.iterrows():
                lot_no = str(r.get("Lot No.", "")).strip()
                cat = str(r.get("Category", "")).lower()
                qty = clean_quantity(r.get("Shipped Qty", 0))
                s_date = str(r.get("Shipment Date", ""))
                
                lot_addbacks[lot_no] = lot_addbacks.get(lot_no, 0.0) + qty
                events.append({"date": s_date, "type": "outbound", "cat": "wip" if "wip" in cat else "fg", "qty": qty})

        if not df_pack_rep.empty:
            df_pack_rep.columns = [str(c).strip() for c in df_pack_rep.columns]
            for _, r in df_pack_rep.iterrows():
                wip_lot = str(r.get("Source WIP Lot", "")).strip()
                qty = clean_quantity(r.get("Quantity", 0))
                waste = clean_quantity(r.get("Waste Qty", 0))
                p_date = str(r.get("Packing Date", ""))
                
                withdraw_qty = qty + waste
                lot_addbacks[wip_lot] = lot_addbacks.get(wip_lot, 0.0) + withdraw_qty
                
                events.append({"date": p_date, "type": "outbound", "cat": "wip", "qty": withdraw_qty})
                events.append({"date": p_date, "type": "inbound", "cat": "fg", "qty": qty})

        if not df_main_rep.empty:
            df_main_rep.columns = [str(c).strip() for c in df_main_rep.columns]
            for _, r in df_main_rep.iterrows():
                lot_no = str(r.get("THAI STICK Lot. No.", "")).strip()
                if not lot_no: lot_no = str(r.get("Supplier Lot. No.", "")).strip()
                    
                cat = str(r.get("Category", "")).lower()
                curr_qty = clean_quantity(r.get("Quantity", 0))
                rec_date = str(r.get("Receive date", r.get("Date receive", r.get("Date Receive", ""))))
                
                original_qty = curr_qty + lot_addbacks.get(lot_no, 0.0)
                events.append({"date": rec_date, "type": "inbound", "cat": "wip" if "wip" in cat else "fg", "qty": original_qty})

        month_start = pd.to_datetime(f"{selected_year}-{selected_month_num:02d}-01")
        if selected_month_num == 12: month_end = pd.to_datetime(f"{selected_year+1}-01-01")
        else: month_end = pd.to_datetime(f"{selected_year}-{selected_month_num+1:02d}-01")

        wip_beginning, fg_beginning = 0.0, 0.0
        wip_inbound, fg_inbound = 0.0, 0.0
        wip_outbound, fg_outbound = 0.0, 0.0

        for ev in events:
            d = parse_robust_date(ev["date"])
            if pd.isna(d): continue
            
            q = ev["qty"]
            cat = ev["cat"]
            
            if d < month_start:
                if ev["type"] == "inbound":
                    if cat == "wip": wip_beginning += q
                    else: fg_beginning += q
                elif ev["type"] == "outbound":
                    if cat == "wip": wip_beginning -= q
                    else: fg_beginning -= q
            elif month_start <= d < month_end:
                if ev["type"] == "inbound":
                    if cat == "wip": wip_inbound += q
                    else: fg_inbound += q
                elif ev["type"] == "outbound":
                    if cat == "wip": wip_outbound += q
                    else: fg_outbound += q

        wip_beginning = max(0.0, wip_beginning)
        fg_beginning = max(0.0, fg_beginning)
        wip_ending = max(0.0, wip_beginning + wip_inbound - wip_outbound)
        fg_ending = max(0.0, fg_beginning + fg_inbound - fg_outbound)

        has_wip_activity = (wip_inbound > 0) or (wip_outbound > 0) or (wip_beginning > 0)
        has_fg_activity = (fg_inbound > 0) or (fg_outbound > 0) or (fg_beginning > 0)

        if not has_wip_activity: st.warning(f"⚠️ **WIP:** ไม่มียอดยกมาและไม่มีกิจกรรมในเดือน {selected_month_label} {selected_year}")
        if not has_fg_activity: st.warning(f"⚠️ **FG:** ไม่มียอดยกมาและไม่มีกิจกรรมในเดือน {selected_month_label} {selected_year}")

        summary_report_data = [
            {
                "Category": "WIP (Work in Progress)",
                "Status Activity": "🟢 Active" if (wip_inbound > 0 or wip_outbound > 0) else "⚪ No Activity (นิ่ง)",
                "Beginning Stock (ยอดยกมา)": wip_beginning,
                "Received / Inbound": wip_inbound,
                "Shipped / Packed (Outbound)": wip_outbound,
                "Ending On-Hand (คงเหลือสิ้นเดือน)": wip_ending,
            },
            {
                "Category": "FG (Finished Goods)",
                "Status Activity": "🟢 Active" if (fg_inbound > 0 or fg_outbound > 0) else "⚪ No Activity (นิ่ง)",
                "Beginning Stock (ยอดยกมา)": fg_beginning,
                "Received / Inbound": fg_inbound,
                "Shipped / Packed (Outbound)": fg_outbound,
                "Ending On-Hand (คงเหลือสิ้นเดือน)": fg_ending,
            },
        ]
        
        df_summary_rep = pd.DataFrame(summary_report_data)
        st.markdown("")
        st.dataframe(format_df_quantity(df_summary_rep), use_container_width=True)

    with rep_tab2:
        st.subheader(f"🚚 Shipped Out Report ประจำเดือน {selected_month_label} {selected_year}")
        if not monthly_shipped.empty:
            reason_summary = monthly_shipped.groupby("Shipping Reason")["Clean_Shipped_Qty"].sum().reset_index()
            reason_summary.columns = ["Shipping Reason (เหตุผลการส่ง)", "Total Shipped Qty (จำนวนรวม)"]
            st.markdown("##### 📌 สรุปยอดจัดส่งแยกตามเหตุผล")
            st.dataframe(format_df_quantity(reason_summary), use_container_width=True)
            st.markdown("##### 📋 รายละเอียดรายการส่งออกทั้งหมด")
            st.dataframe(format_df_quantity(monthly_shipped), use_container_width=True)
        else:
            st.info(f"ℹ️ ไม่พบประวัติการส่งออกในเดือน {selected_month_label} {selected_year}")

# ==========================================
# หน้าที่ 4: Trading Goods Registration
# ==========================================
elif menu == "📝 Trading Goods Registration":
    render_registration_ui_styles()
    with st.container(key="registration_create"):
        header, shortcut = st.columns([4, 1.4], gap="large")
        with header:
            with st.container(key="registration_header"):
                st.markdown('<p class="registration-eyebrow">NEW REGISTRATION</p>', unsafe_allow_html=True)
                st.title("Trading Goods Registration")
                st.caption("ลงทะเบียน WIP / FG พร้อมข้อมูลผู้จัดหา เลข Lot และเงื่อนไขการจัดส่ง")
        with shortcut:
            with st.expander("🔗 Google Sheets"):
                pwd = st.text_input("รหัสผ่าน:", type="password", key="gsheet_shortcut_pwd", placeholder="กรอกรหัสผ่าน")
                if pwd == "tstd":
                    st.markdown("[เปิด Google Sheets](https://docs.google.com/spreadsheets/d/1RLtRn7gojLLLxrsm9mSFRLy7eGmCYRn5bgsoCmuju7Y/edit?gid=0#gid=0)")
                elif pwd:
                    st.error("รหัสผ่านไม่ถูกต้อง")
        flash = st.session_state.get("registration_flash")
        if flash and flash["mode"] == "create":
            st.success(st.session_state.pop("registration_flash")["message"])
        submitted, registration = render_registration_form()
        if submitted:
            try:
                sheet.append_row(values=list(registration.values()), value_input_option="USER_ENTERED")
            except Exception as error:
                st.error(f"บันทึกยังไม่สำเร็จ กรุณาลองใหม่: {error}")
            else:
                get_cached_data.clear()
                st.session_state["registration_create_reset"] = True
                st.session_state["registration_flash"] = {
                    "mode": "create", "message": "บันทึกสินค้าเข้าระบบเรียบร้อยแล้ว พร้อมลงทะเบียนรายการถัดไป",
                }
                st.rerun()

# ==========================================
# หน้าที่ 5: Edit Registration
# ==========================================
elif menu == "✏️ Edit Registration":
    render_registration_ui_styles()
    with st.container(key="registration_edit"):
        with st.container(key="registration_header"):
            st.markdown('<p class="registration-eyebrow">EDIT REGISTRATION</p>', unsafe_allow_html=True)
            st.title("Edit Registration")
            st.caption("ค้นหาล็อตสินค้า ตรวจสอบข้อมูลปัจจุบัน และแก้ไขรายละเอียดในแบบฟอร์มด้านล่าง")
        flash = st.session_state.get("registration_flash")
        if flash and flash["mode"] == "edit":
            st.success(st.session_state.pop("registration_flash")["message"])
        data = get_cached_data("main", sheet, include_row_index=True)
        if data:
            df_all = pd.DataFrame(data)
            df_all.columns = [str(column).strip() for column in df_all.columns]
            df_all["Clean_Qty"] = df_all["Quantity"].apply(lambda value: picking_quantity(value) or 0.0) if "Quantity" in df_all else 0.0
            active = df_all["Clean_Qty"].gt(0)
            if "Status" in df_all:
                active &= ~df_all["Status"].astype(str).str.strip().str.casefold().eq("delivered")
            editable_df = df_all[active].copy()
            if editable_df.empty:
                st.info("ไม่มีล็อตสินค้าที่สามารถแก้ไขได้ในขณะนี้")
            else:
                with st.container(key="registration_filter"):
                    st.markdown("**ค้นหาล็อตที่ต้องการแก้ไข**")
                    search_col, category_col = st.columns([3, 1], gap="large")
                    with search_col:
                        query = st.text_input("ค้นหาชื่อสายพันธุ์, Supplier หรือ Lot", placeholder="พิมพ์ชื่อหรือเลข Lot ที่ต้องการค้นหา", key="registration_edit_search")
                    with category_col:
                        selected_category = st.selectbox("ประเภทสต็อก", ["ทุกประเภท", "WIP", "FG"], key="registration_edit_category")
                    filtered = editable_df
                    if query.strip():
                        fields = [field for field in ["Cultivar Name", "Export Name", "Supplier Name", "Supplier Lot. No.", "THAI STICK Lot. No."] if field in filtered]
                        search_text = filtered[fields].fillna("").astype(str).agg(" ".join, axis=1)
                        filtered = filtered[search_text.str.casefold().str.contains(query.strip().casefold(), regex=False)]
                    if selected_category != "ทุกประเภท":
                        categories = filtered["Category"].astype(str) if "Category" in filtered else pd.Series("", index=filtered.index)
                        filtered = filtered[categories.str.contains(selected_category, case=False, na=False, regex=False)]
                    st.caption(f"พบ {len(filtered):,} รายการ จากล็อตที่แก้ไขได้ทั้งหมด {len(editable_df):,} รายการ · แสดงเฉพาะสต็อกคงเหลือที่ยังไม่ Delivered")
                    records_by_row = {int(row["Original_Row_Index"]): row.to_dict() for _, row in filtered.iterrows()}
                    labels = {}
                    for index, (row_number, record) in enumerate(records_by_row.items(), start=1):
                        cultivar = picking_text(record.get("Cultivar Name")) or "ไม่ระบุสายพันธุ์"
                        lot = picking_text(record.get("THAI STICK Lot. No.")) or picking_text(record.get("Supplier Lot. No.")) or "ไม่ระบุ Lot"
                        grade = picking_text(record.get("Grade")) or "—"
                        labels[row_number] = f"{index:02d} · {cultivar} | Lot {lot} | {grade} | {record['Clean_Qty']:,.2f} g"
                    if records_by_row:
                        selected_row = st.selectbox("เลือกล็อตสินค้า", list(records_by_row), format_func=labels.get, key="registration_edit_lot")
                if not records_by_row:
                    st.info("ไม่พบล็อตที่ตรงกับการค้นหา ลองเปลี่ยนคำค้นหรือเลือกทุกประเภท")
                else:
                    current = records_by_row[selected_row]
                    render_registration_selected(current)
                    submitted, registration = render_registration_form(editing=True, current_data=current, row_number=selected_row)
                    if submitted:
                        try:
                            sheet.update(range_name=f"A{selected_row}:P{selected_row}", values=[[str(value) for value in registration.values()]])
                        except Exception as error:
                            st.error(f"ยังบันทึกการแก้ไขไม่สำเร็จ กรุณาลองใหม่: {error}")
                        else:
                            get_cached_data.clear()
                            lot = registration["THAI STICK Lot. No."] or registration["Supplier Lot. No."] or registration["Cultivar Name"]
                            st.session_state["registration_flash"] = {
                                "mode": "edit", "message": f"บันทึกการแก้ไข Lot {lot} เรียบร้อยแล้ว",
                            }
                            st.rerun()
        else:
            st.info("ยังไม่มีข้อมูลในระบบให้แก้ไข")


# ==========================================
# หน้าที่ 6: Packing Module
# ==========================================
elif menu == "📦 Packing Module":
    st.title("📦 Packing Module (WIP Conversion & Dedicated FG Tracking)")
    
    data = get_cached_data("main", sheet)
    if data:
        df = pd.DataFrame(data)
        df.columns = [str(c).strip() for c in df.columns]

        if "Category" in df.columns:
            wip_df = df[df["Category"].astype(str).str.contains("WIP", case=False, na=False)].copy()
        else: wip_df = pd.DataFrame()

        if wip_df.empty:
            st.warning("⚠️ ยังไม่พบข้อมูลสินค้าประเภท 'WIP' ในระบบ")
        else:
            wip_df["Clean_Qty"] = wip_df["Quantity"].apply(clean_quantity)
            available_wip = wip_df[wip_df["Clean_Qty"] > 0].copy()

            if available_wip.empty: 
                st.warning("⚠️ มีข้อมูล WIP แต่ทุก Lot มีจำนวนคงเหลือเป็น 0")
            else:
                available_wip["Original_Row_Index"] = available_wip.index + 2 

                unique_cultivars = sorted([str(c) for c in available_wip["Cultivar Name"].unique() if str(c).strip() != ""])
                selected_cultivar = st.selectbox("🔍 1. ค้นหาและกรองสายพันธุ์ (Cultivar Filter)", ["-- แสดงทั้งหมด --"] + unique_cultivars)

                if selected_cultivar != "-- แสดงทั้งหมด --":
                    filtered_wip = available_wip[available_wip["Cultivar Name"] == selected_cultivar]
                else:
                    filtered_wip = available_wip

                wip_options = []
                wip_map = {}

                for _, r in filtered_wip.iterrows():
                    lot_no = str(r.get("THAI STICK Lot. No.", "")).strip()
                    if not lot_no: lot_no = str(r.get("Supplier Lot. No.", "")).strip()
                    
                    cultivar = str(r.get("Cultivar Name", ""))
                    grade = str(r.get("Grade", ""))
                    qty = float(r["Clean_Qty"])
                    row_idx = r["Original_Row_Index"]
                    
                    # ✨ ดึง Supplier Name
                    supplier_name = str(r.get("Supplier Name", "")).strip()
                    if not supplier_name: supplier_name = "ไม่ระบุ"
                    
                    disp = f"[บรรทัด {row_idx}] Supplier: {supplier_name} | Lot: {lot_no} | Size: {grade} | {cultivar} | Qty: {qty:,.2f}"

                    wip_options.append(disp)
                    wip_map[disp] = {
                        "lot_no": lot_no,
                        "cultivar": cultivar,
                        "grade": grade,
                        "qty": qty,
                        "supplier": supplier_name, 
                        "row_idx": row_idx,
                        "row_data": r.to_dict()
                    }

                selected_wip_disps = st.multiselect("📦 2. เลือก WIP ที่ต้องการนำมาแพ็ก (เลือกได้มากกว่า 1 รายการ)", wip_options)

                if selected_wip_disps:
                    with st.form("packing_form", clear_on_submit=True):
                        st.markdown("#### 📥 กำหนดจำนวนเบิกของแต่ละ Lot (Withdraw Quantities)")
                        wip_withdrawals = {}
                        
                        for disp in selected_wip_disps:
                            info = wip_map[disp]
                            mq = info["qty"]
                            w_qty = st.number_input(
                                f"เบิกจาก {disp}", 
                                min_value=0.0, 
                                max_value=float(mq), 
                                step=0.1, 
                                value=float(mq), 
                                key=f"pack_w_{info['row_idx']}"
                            )
                            wip_withdrawals[disp] = w_qty

                        st.divider()
                        default_fg_name = wip_map[selected_wip_disps[0]]["cultivar"]
                        
                        col_p1, col_p2 = st.columns(2)
                        with col_p1: packing_date_str = str(st.date_input("Packing date", value=date.today()))
                        with col_p2: packed_fg_name = st.text_input("ชื่อผลิตภัณฑ์หลังแพ็ก", value=default_fg_name)

                        st.divider()
                        st.markdown("#### 📤 ผลผลิตที่ได้จากการแพ็ก (Finished Goods Output)")
                        grade_options = ["Big", "Small", "Mix", "F", "Grade A", "Premium"]
                        
                        c1, c2 = st.columns(2)
                        with c1: g1_type = st.selectbox("เกรดช่องที่ 1", grade_options, index=0)
                        with c2: g1_qty = st.number_input("จำนวนช่องที่ 1 (กรอกเป็นตัวเลข)", min_value=0.0, step=0.1, key="g1")

                        c3, c4 = st.columns(2)
                        with c3: g2_type = st.selectbox("เกรดช่องที่ 2", grade_options, index=1)
                        with c4: g2_qty = st.number_input("จำนวนช่องที่ 2 (กรอกเป็นตัวเลข)", min_value=0.0, step=0.1, key="g2")

                        c5, c6 = st.columns(2)
                        with c5: g3_type = st.selectbox("เกรดช่องที่ 3", grade_options, index=2)
                        with c6: g3_qty = st.number_input("จำนวนช่องที่ 3 (กรอกเป็นตัวเลข)", min_value=0.0, step=0.1, key="g3")

                        c7, c8 = st.columns(2)
                        with c7: g4_type = st.selectbox("เกรดช่องที่ 4", grade_options, index=3)
                        with c8: g4_qty = st.number_input("จำนวนช่องที่ 4 (กรอกเป็นตัวเลข)", min_value=0.0, step=0.1, key="g4")

                        st.divider()
                        waste_qty = st.number_input("จำนวนของเสีย (Waste Qty)", min_value=0.0, step=0.1, key="waste_pack")
                        note = st.text_area("หมายเหตุเพิ่มเติม (Note)", key="note_pack")

                        if st.form_submit_button("🚀 ยืนยันการแพ็กและบันทึกลง Google Sheets", key="packing_submit_btn"):
                            total_withdraw = sum(wip_withdrawals.values())
                            total_packed_output = g1_qty + g2_qty + g3_qty + g4_qty
                          
                            if total_withdraw <= 0: 
                                st.error("❌ กรุณาระบุจำนวนที่เบิกมาแพ็กมากกว่า 0 อย่างน้อย 1 รายการ")
                            elif (total_packed_output + waste_qty) > total_withdraw: 
                                st.warning(f"⚠️ คำเตือน: ผลผลิตรวม ({total_packed_output + waste_qty:,.2f}) มากกว่าจำนวนที่เบิกมาทั้งหมด ({total_withdraw:,.2f})")
                            else:
                                with st.spinner("⏳ กำลังบันทึกข้อมูล..."):
                                    latest_headers = [
                                        "Cultivar Name", "Supplier Name", "Supplier Lot. No.", "Export Name", "THAI STICK Lot. No.", 
                                        "Category", "Grade", "Receive date", "Quantity", "Remark Qty", "Inspection Req.", "IR Req.", 
                                        "Shipment date", "Type", "Status", "Packing date"
                                    ]
                                    col_end_letter = chr(64 + len(latest_headers)).upper()
                                    
                                    combined_source_lots = [] 
                                    combined_suppliers = [] # ✨ เก็บชื่อ Supplier
                                    
                                    for disp, w_qty in wip_withdrawals.items():
                                        if w_qty > 0:
                                            info = wip_map[disp]
                                            row_idx = info["row_idx"]
                                            lot_no = info["lot_no"]
                                            supplier = info.get("supplier", "ไม่ระบุ")
                                            
                                            combined_source_lots.append(lot_no)
                                            if supplier not in combined_suppliers:
                                                combined_suppliers.append(supplier)
                                            
                                            r_dict = info["row_data"]
                                            remaining_wip = info["qty"] - w_qty
                                            new_status = "Packing done" if remaining_wip <= 0 else "Pending Packing"
                                            remaining_wip = max(0.0, remaining_wip)
                                            
                                            r_dict["Quantity"] = float(remaining_wip)
                                            r_dict["Status"] = new_status
                                            
                                            row_values = [str(r_dict.get(h, "")) for h in latest_headers]
                                            sheet.update(range_name=f"A{row_idx}:{col_end_letter}{row_idx}", values=[row_values], value_input_option='USER_ENTERED')

                                    source_lots_str = ", ".join(combined_source_lots) if combined_source_lots else "TS-MIX"
                                    base_lot = combined_source_lots[0] if combined_source_lots else "TS-MIX"
                                    suppliers_str = ", ".join(combined_suppliers) if combined_suppliers else "ไม่ระบุ" # ✨ รวมชื่อ Supplier

                                    breakdown_str = f"{g1_type}: {g1_qty} | {g2_type}: {g2_qty} | {g3_type}: {g3_qty} | {g4_type}: {g4_qty}"
                                    log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    
                                    # ✨ บันทึก Log พร้อม Supplier Name
                                    packing_log_row = [
                                        log_time, source_lots_str, suppliers_str, default_fg_name, packed_fg_name, breakdown_str, 
                                        float(total_packed_output), float(waste_qty), packing_date_str, note
                                    ]
                                    packing_sheet.append_row(packing_log_row, value_input_option='USER_ENTERED')

                                    fg_items = [(g1_type, g1_qty), (g2_type, g2_qty), (g3_type, g3_qty), (g4_type, g4_qty)]
                                    created_count = 0
                                    
                                    for idx_fg, (g_type, g_q) in enumerate(fg_items):
                                        if g_q > 0:
                                            created_count += 1
                                            custom_fg_lot = f"{base_lot}-FG{created_count}-{g_type}"
                                            
                                            # ✨ บันทึก Record พร้อม Supplier Name
                                            rec_row = [
                                                log_time, source_lots_str, suppliers_str, default_fg_name, packed_fg_name, custom_fg_lot, 
                                                g_type, float(g_q), float(waste_qty) if idx_fg == 0 else 0.0, packing_date_str, note
                                            ]
                                            packing_records_sheet.append_row(rec_row, value_input_option='USER_ENTERED')

                                    get_cached_data.clear()
                                    st.success(f"✅ บันทึกสำเร็จ! นำ WIP จำนวน {total_withdraw:,.2f} g ไปแพ็ก และสร้าง Lot FG ใหม่ {created_count} รายการ 🎉")
                                    time.sleep(1)
                                    st.rerun()

    st.divider()
    st.subheader("📑 ประวัติการผลิตและรหัส Lot FG แยกเกรด (Packing Records Sheet)")
    records_data = get_cached_data("packing", packing_records_sheet)
    if records_data: st.dataframe(format_df_quantity(pd.DataFrame(records_data)), use_container_width=True)

# ==========================================
# หน้าที่ 7: Shipped Out Module
# ==========================================
elif menu == "🚚 Shipped Out Module":
    st.title("🚚 Shipped Out Module (Multi-Product Shipping & Stock Deduction)")
    st.markdown("บันทึกการส่งสินค้าออกทั้ง WIP และ FG (ดึงข้อมูลจากทั้งหน้า Trading Data และ Packing Records)")

    data_main = get_cached_data("main", sheet)
    df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()

    records_data = get_cached_data("packing", packing_records_sheet)
    df_packing_fg = pd.DataFrame(records_data) if records_data else pd.DataFrame()

    combined_options = []
    inventory_map = {}

    if not df_main.empty and "Category" in df_main.columns:
        df_main["Original_Row_Index"] = df_main.index + 2 
        df_main["Clean_Qty"] = df_main["Quantity"].apply(clean_quantity)
        main_avail = df_main[(df_main["Clean_Qty"] > 0) & (~df_main["Status"].astype(str).str.strip().str.lower().eq("delivered"))]
        
        for _, r in main_avail.iterrows():
            cat = str(r.get("Category", ""))
            lot_no = str(r.get("THAI STICK Lot. No.", "")).strip()
            if not lot_no: lot_no = str(r.get("Supplier Lot. No.", "")).strip()
            if not lot_no: lot_no = "No-Lot"

            prod_name = str(r.get("Cultivar Name", ""))
            grade_val = str(r.get("Grade", ""))
            max_q = float(r["Clean_Qty"])

            if "WIP" in cat.upper():
                disp = f"[WIP] {lot_no} | {prod_name} | Grade: {grade_val} | Qty: {max_q:,.1f}"
                source = "main"
                c = "WIP (Work in Progress)"
            elif "FG" in cat.upper():
                disp = f"[FG - Stock] {lot_no} | {prod_name} | Grade: {grade_val} | Qty: {max_q:,.1f}"
                source = "main"
                c = "FG (Finished Goods)"
            else: continue

            combined_options.append(disp)
            inventory_map[disp] = {
                "category": c, "lot_no": lot_no, "product": prod_name, "grade": grade_val,
                "max_qty": max_q, "source": source, "row_idx": r["Original_Row_Index"], "row_data": r.to_dict()
            }

    if not df_packing_fg.empty:
        df_packing_fg["Original_Row_Index"] = df_packing_fg.index + 2
        df_packing_fg["Clean_Qty"] = df_packing_fg["Quantity"].apply(clean_quantity)
        fg_avail = df_packing_fg[df_packing_fg["Clean_Qty"] > 0]
        
        for _, r in fg_avail.iterrows():
            lot_no = str(r.get("Generated FG Lot No.", "")).strip()
            if not lot_no: lot_no = str(r.get("THAI STICK Lot. No.", "")).strip()
            
            prod_name = str(r.get("Export Name", ""))
            grade_val = str(r.get("Grade", ""))
            max_q = float(r["Clean_Qty"])

            disp = f"[FG - Packed] {lot_no} | {prod_name} | Grade: {grade_val} | Qty: {max_q:,.1f}"
            combined_options.append(disp)
            inventory_map[disp] = {
                "category": "FG (Finished Goods)", "lot_no": lot_no, "product": prod_name,
                "grade": grade_val, "max_qty": max_q, "source": "packing_records", "row_idx": r["Original_Row_Index"]
            }

    if "ship_row_count" not in st.session_state: st.session_state.ship_row_count = 3

    col_cnt1, col_cnt2, _ = st.columns([2, 2, 6])
    with col_cnt1: 
        if st.button("➕ เพิ่มช่อง"): st.session_state.ship_row_count = min(10, st.session_state.ship_row_count + 1); st.rerun()
    with col_cnt2: 
        if st.button("➖ ลดช่อง"): st.session_state.ship_row_count = max(1, st.session_state.ship_row_count - 1); st.rerun()

    with st.form("shipped_out_multi_form"):
        st.subheader("1. ข้อมูลการจัดส่งทั่วไป")
        col_s1, col_s2, col_s3 = st.columns(3)
        with col_s1: consignee = st.text_input("ชื่อปลายทาง (Consignee)")
        with col_s2: client_name = st.text_input("ลูกค้า (Client)")
        with col_s3: shipping_reason = st.selectbox("เหตุผลในการส่ง (Shipping Reason)", ["SAMPLE", "EXPORT", "RETURN", "REJECT", "DOMESTIC SALE"])

        shipment_date_str = str(st.date_input("วันที่จัดส่ง (Shipment Date)", value=date.today()))

        st.divider()
        st.subheader(f"2. รายการผลิตภัณฑ์ที่ต้องการส่งออก (แสดง {st.session_state.ship_row_count} ช่อง)")
        shipment_items = []

        for i in range(st.session_state.ship_row_count):
            row_c1, row_c2 = st.columns([4, 2])
            with row_c1: chosen_display = st.selectbox(f"เลือกรายการสินค้า #{i+1}", ["-- กรุณาเลือก --"] + combined_options, key=f"ship_item_{i}")
            
            max_q = inventory_map.get(chosen_display, {}).get("max_qty", 0.0) if chosen_display and not chosen_display.startswith("--") else 0.0
            with row_c2: ship_q = st.number_input(f"จำนวนส่งออก (Max: {max_q:,.1f})", min_value=0.0, max_value=max_q if max_q > 0 else 1000000.0, step=0.1, key=f"ship_qty_{i}")

            if chosen_display and not chosen_display.startswith("--") and ship_q > 0:
                item_info = inventory_map[chosen_display]
                shipment_items.append({**item_info, "qty": ship_q})

        note_ship = st.text_area("หมายเหตุ (Note)")
        if st.form_submit_button("🚚 ยืนยันการส่งสินค้าออกทั้งหมด"):
            if not consignee.strip(): st.error("❌ กรุณาระบุชื่อปลายทาง (Consignee)")
            elif not client_name.strip(): st.error("❌ กรุณาระบุชื่อลูกค้า (Client)")
            elif not shipment_items: st.error("❌ กรุณาเลือก Lot สินค้าและระบุจำนวนที่ต้องการส่งมากกว่า 0 อย่างน้อย 1 รายการ")
            else:
                log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                success_count = 0

                for item in shipment_items:
                    row_idx = item["row_idx"]
                    source = item["source"]
                    qty = item["qty"]

                    if source == "main":
                        # ✨ แก้ไขใหม่: ดึง Header จาก Google Sheets จริงมาเทียบชื่อคอลัมน์ เพื่อป้องกันคอลัมน์เพี้ยน
                        sheet_headers = [str(h).strip() for h in sheet.row_values(1)]
                        current_row_values = sheet.row_values(row_idx)
                        
                        # เติมช่องว่างให้ครบตามจำนวน Header
                        while len(current_row_values) < len(sheet_headers):
                            current_row_values.append("")

                        row_dict = {sheet_headers[i]: current_row_values[i] for i in range(len(sheet_headers))}
                        
                        current_qty = clean_quantity(row_dict.get("Quantity", 0))
                        remaining_qty = max(0.0, current_qty - qty)
                        new_status = "Delivered" if remaining_qty <= 0 else ("Pending Packing" if "WIP" in item["category"] else "Trading Goods Received to Stock")

                        # อัปเดตค่าเฉพาะคอลัมน์ Quantity และ Status ตามชื่อจริง
                        row_dict["Quantity"] = str(remaining_qty)
                        row_dict["Status"] = new_status

                        # จัดเรียงข้อมูลกลับลง List ตามตำแหน่ง Header ปัจจุบันในชีต
                        updated_row_values = [str(row_dict.get(h, "")) for h in sheet_headers]
                        col_end_letter = chr(64 + len(sheet_headers)).upper()
                        
                        sheet.update(range_name=f"A{row_idx}:{col_end_letter}{row_idx}", values=[updated_row_values], value_input_option='USER_ENTERED')

                    elif source == "packing_records":
                        # ✨ แก้ไขใหม่: ค้นหาตำแหน่งคอลัมน์ Quantity จากชื่อ Header จริงใน Packing_Records
                        pack_headers = [str(h).strip() for h in packing_records_sheet.row_values(1)]
                        qty_col_idx = pack_headers.index("Quantity") + 1 if "Quantity" in pack_headers else 8
                        
                        all_pack_records = get_cached_data("packing", packing_records_sheet)
                        if row_idx and (row_idx - 2) < len(all_pack_records):
                            current_qty = clean_quantity(all_pack_records[row_idx - 2].get("Quantity", 0))
                            packing_records_sheet.update_cell(row_idx, qty_col_idx, max(0.0, current_qty - qty))

                    shipped_sheet.append_row([
                        log_time, item["category"], item["lot_no"], item["product"], item["grade"], str(qty),
                        consignee, client_name, shipping_reason, shipment_date_str, note_ship
                    ], value_input_option='USER_ENTERED')
                    success_count += 1

                get_cached_data.clear()
                st.success(f"✅ บันทึกการส่งออกสำเร็จทั้งหมด {success_count} รายการและตัดสต็อกต้นทางเรียบร้อยแล้ว! 🎉")
                time.sleep(1)
                st.rerun()

    st.divider()
    st.subheader("📑 ประวัติการส่งสินค้าออก (Shipped Out Log)")
    shipped_data = get_cached_data("shipped", shipped_sheet)
    if shipped_data: st.dataframe(format_df_quantity(pd.DataFrame(shipped_data)), use_container_width=True)
