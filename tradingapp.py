from datetime import date, datetime
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import streamlit as st
import re
from html import escape

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

def render_portal_ui_styles():
    st.markdown("""<style>
    .st-key-delivery_header h1, .st-key-stock_header h1 {
        color: #0f172a; font-size: clamp(27px, 2.8vw, 36px); line-height: 1.3;
        letter-spacing: -.035em; padding-bottom: 8px; }
    .st-key-delivery_header [data-testid="stCaptionContainer"],
    .st-key-stock_header [data-testid="stCaptionContainer"] { color: #64748b; font-size: 14px; }
    .st-key-delivery_header [data-testid="stCaptionContainer"] p,
    .st-key-stock_header [data-testid="stCaptionContainer"] p { color: #64748b; }
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


# ==========================================
# 5. Sidebar เมนูด้านซ้าย
# ==========================================
st.sidebar.title("🌱 Thai Stick Trading")
st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "เลือกเมนูการใช้งาน",
    ["🚨 Delivery Issue", "📊 Stock Summary & Details", "📈 Monthly Report", "📝 Trading Goods Registration", "✏️ Edit Registration", "📦 Packing Module", "🚚 Shipped Out Module"],
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

    try:
        del_sheet = sh.worksheet("Delivery_schedule")
        data_del = get_cached_data("delivery", del_sheet, include_row_index=True)
        df_del = pd.DataFrame(data_del) if data_del else pd.DataFrame()
    except gspread.exceptions.WorksheetNotFound:
        st.error("❌ ไม่พบแท็บชื่อ 'Delivery_schedule' ใน Google Sheets")
        df_del = pd.DataFrame()
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการดึงข้อมูล: {e}")
        df_del = pd.DataFrame()

    if not df_del.empty:
        df_del.columns = [str(c).strip() for c in df_del.columns]

        vol_col = next((c for c in df_del.columns if "volume" in str(c).lower()), None)
        fg_col = next((c for c in df_del.columns if "available fg" in str(c).lower()), None)
        wip_col = next((c for c in df_del.columns if "available wip" in str(c).lower()), None)

        if vol_col and fg_col:
            def check_readiness(row):
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

        date_col = next((c for c in ["Expected Shipping Date", "Shipment date", "Delivery date", "Shipment Date", "Delivery Date"] if c in df_del.columns), None)

        delivery_internal_columns = ["Original_Row_Index", "Is_Unscheduled", "Parsed_Delivery_Date"]
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

                        item_name = "Unknown Item"
                        for nk in ["Cultivar Name", "Cultivar name", "Cultivar", "Export Name", "Product", "ชื่อ"]:
                            if nk in df_del.columns and str(r.get(nk, "")).strip() and str(r.get(nk, "")).strip().lower() != 'nan':
                                item_name = str(r.get(nk, "")).strip(); break

                        item_grade = "-"
                        for gk in ["Grade", "grade", "Size", "size"]:
                            if gk in df_del.columns and str(r.get(gk, "")).strip() and str(r.get(gk, "")).strip().lower() != 'nan':
                                item_grade = str(r.get(gk, "")).strip(); break

                        try: vol_disp = f"{float(vol):,.2f}" if float(vol) % 1 != 0 else f"{int(vol):,}"
                        except: vol_disp = str(vol)

                        disp_txt = f"บรรทัด {row_idx} | {item_name} (Grade: {item_grade}) | วันเดิม: {curr_dt} | Volume: {vol_disp}"
                        reschedule_opts.append(disp_txt)
                        reschedule_map[disp_txt] = {"row_idx": row_idx, "curr_date": curr_dt}

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
                                headers = [str(header).strip() for header in del_sheet.row_values(1)]
                                if date_col not in headers:
                                    st.error(f"ไม่พบคอลัมน์ {date_col} ใน Google Sheets")
                                else:
                                    col_idx = headers.index(date_col) + 1
                                    new_value = "TBC" if mark_tbc else new_date.isoformat()
                                    cells_to_update = []
                                    success_rows = []
                                    for selected in selected_items:
                                        target_row = int(reschedule_map[selected]["row_idx"])
                                        cells_to_update.append(gspread.Cell(row=target_row, col=col_idx, value=new_value))
                                        success_rows.append(str(target_row))
                                    del_sheet.update_cells(cells_to_update)
                            except Exception as error:
                                st.error(f"❌ บันทึกกำหนดการไม่สำเร็จ กรุณาลองใหม่: {error}")
                            else:
                                if date_col in headers:
                                    get_cached_data.clear()
                                    schedule_label = "TBC (รอยืนยันวันจัดส่งใหม่)" if mark_tbc else new_date.strftime("%d/%m/%Y")
                                    st.session_state["delivery_reschedule_success"] = (
                                        f"✅ อัปเดตกำหนดส่งเป็น {schedule_label} สำหรับบรรทัดที่ {', '.join(success_rows)} เรียบร้อยแล้ว"
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
    col_head1, col_head2 = st.columns([7, 3])
    with col_head1:
        st.title("📝 Trading Goods Registration")
        st.markdown("กรอกข้อมูลผลิตภัณฑ์และประเภทสินค้า (WIP หรือ FG) ลงในระบบ")
    with col_head2:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("🔗 Direct Google Sheets (Multiple Batch)"):
            pwd = st.text_input("รหัสผ่าน:", type="password", key="gsheet_shortcut_pwd", placeholder="กรอกรหัสผ่าน")
            if pwd == "tstd": st.markdown("[🟢 **คลิกเพื่อเปิด Google Sheets**](https://docs.google.com/spreadsheets/d/1RLtRn7gojLLLxrsm9mSFRLy7eGmCYRn5bgsoCmuju7Y/edit?gid=0#gid=0)")
            elif pwd != "": st.error("รหัสผ่านไม่ถูกต้อง ❌")

    with st.form("trading_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            cultivar_name = st.text_input("1. Cultivar Name")
            supplier_name = st.text_input("2. Supplier Name")
            supplier_lot = st.text_input("3. Supplier Lot. No.")
            export_name = st.text_input("4. Export Name")
            thai_stick_lot = st.text_input("5. THAI STICK Lot. No.")
            category = st.selectbox("6. Category", ["FG (Finished Goods)", "WIP (Work in Progress)"])
            grade = st.selectbox("7. Grade", ["Big", "Small", "Mix", "F", "Grade A", "Premium"])
        with col2:
            receive_date_str = str(st.date_input("8. Receive date", value=date.today(), key="reg_rec_dt"))
            quantity = st.number_input("9. Quantity (g)", min_value=0.0, step=0.1)
            remark_qty = st.text_input("Remark Qty (เช่น 5x1000)", placeholder="ระบุรายละเอียดจำนวนถุง/แพ็ก")

            insp_con = st.container()
            is_tbd_insp = st.checkbox("To be decided", key="reg_tbd_insp")
            inspection_req_val = insp_con.selectbox("10. Inspection Req.", ["Yes", "No"], disabled=is_tbd_insp, key="reg_insp")
            inspection_req = "To be decided" if is_tbd_insp else inspection_req_val

            ir_con = st.container()
            is_tbd_ir = st.checkbox("To be decided", key="reg_tbd_ir")
            ir_req_val = ir_con.selectbox("11. IR Req.", ["Yes", "No"], disabled=is_tbd_ir, key="reg_ir")
            ir_req = "To be decided" if is_tbd_ir else ir_req_val

            ship_con = st.container()
            is_tbd_ship = st.checkbox("To be decided", key="reg_tbd_ship")
            shipment_date_val = ship_con.date_input("12. Shipment date", value=date.today(), disabled=is_tbd_ship, key="reg_ship")
            shipment_date = "To be decided" if is_tbd_ship else str(shipment_date_val)

            type_con = st.container()
            is_tbd_type = st.checkbox("To be decided", key="reg_tbd_type")
            type_opts = ["Inspection & Packing Service", "Trading Goods (TS)", "Consignment", "In House Product"]
            prod_type_val = type_con.selectbox("13. Type", type_opts, disabled=is_tbd_type, key="reg_type")
            prod_type = "To be decided" if is_tbd_type else prod_type_val

            status = st.selectbox(
                "14. Status",
                ["Trading Goods Received to Stock", "Inspection Completed", "Pending Packing", "Packing done", "Delivered", "Waiting for Return"],
            )
        
        submitted = st.form_submit_button("💾 บันทึกข้อมูลลง Google Sheets")
        if submitted:
            latest_headers = [
                "Cultivar Name", "Supplier Name", "Supplier Lot. No.", "Export Name", "THAI STICK Lot. No.", 
                "Category", "Grade", "Receive date", "Quantity", "Remark Qty", "Inspection Req.", "IR Req.", 
                "Shipment date", "Type", "Status", "Packing date"
            ]
            row_dict = {
                "Cultivar Name": cultivar_name, "Supplier Name": supplier_name, "Supplier Lot. No.": supplier_lot,
                "Export Name": export_name, "THAI STICK Lot. No.": thai_stick_lot, "Category": category, "Grade": grade,
                "Receive date": receive_date_str, "Quantity": float(quantity), "Remark Qty": remark_qty,
                "Inspection Req.": inspection_req, "IR Req.": ir_req, "Shipment date": shipment_date,
                "Type": prod_type, "Status": status, "Packing date": "",
            }
            row_to_insert = [row_dict.get(h, "") for h in latest_headers]
            
            sheet.append_row(values=row_to_insert, value_input_option='USER_ENTERED')
            get_cached_data.clear()
            
            # ✨ เพิ่มส่วนนี้: ล้างค่า Checkbox 'To be decided' ที่ค้างใน Session State ทิ้ง
            keys_to_clear = ["reg_tbd_insp", "reg_insp", "reg_tbd_ir", "reg_ir", "reg_tbd_ship", "reg_ship", "reg_tbd_type", "reg_type", "reg_rec_dt"]
            for k in keys_to_clear:
                if k in st.session_state:
                    del st.session_state[k]
                    
            st.success("บันทึกข้อมูลสำเร็จเรียบร้อยแล้ว! 🎉")
            time.sleep(1)
            st.rerun()

# ==========================================
# หน้าที่ 5: Edit Registration
# ==========================================
elif menu == "✏️ Edit Registration":
    st.title("✏️ Edit Registration Data")
    st.markdown("ค้นหาและแก้ไขข้อมูลล็อตสินค้าที่เคยบันทึกไว้")

    data = get_cached_data("main", sheet)
    if data:
        df_all = pd.DataFrame(data)
        df_all.columns = [str(c).strip() for c in df_all.columns]
        lot_column = "THAI STICK Lot. No." if "THAI STICK Lot. No." in df_all.columns else df_all.columns[4]

        df_all["Clean_Qty"] = df_all["Quantity"].apply(clean_quantity) if "Quantity" in df_all.columns else 0.0
        df_all["Original_Row_Index"] = df_all.index + 2

        if "Status" in df_all.columns:
            editable_df = df_all[(df_all["Clean_Qty"] > 0) & (~df_all["Status"].astype(str).str.strip().str.lower().eq("delivered"))].copy()
        else:
            editable_df = df_all[df_all["Clean_Qty"] > 0].copy()

        if editable_df.empty:
            st.info("ℹ️ ไม่มีล็อตสินค้าที่สามารถแก้ไขได้ในขณะนี้")
        else:
            edit_options = []
            edit_map = {}

            for _, r in editable_df.iterrows():
                lot_no = str(r.get(lot_column, "")).strip()
                if not lot_no: lot_no = str(r.get("Supplier Lot. No.", "")).strip()
                
                cultivar = str(r.get("Cultivar Name", ""))
                grade = str(r.get("Grade", ""))
                qty = float(r["Clean_Qty"])
                row_idx = r["Original_Row_Index"]
                df_idx = r.name
                
                disp = f"Lot: {lot_no} | Size: {grade} | {cultivar} | Qty: {qty:,.2f}"
                edit_options.append(disp)
                edit_map[disp] = {"row_idx": row_idx, "lot_no": lot_no, "df_index": df_idx}

            selected_disp = st.selectbox("🔍 เลือกข้อมูลที่ต้องการแก้ไข (Lot No. & Size)", edit_options)

            if selected_disp:
                selected_info = edit_map[selected_disp]
                matched_row_idx = selected_info["row_idx"]
                selected_lot = selected_info["lot_no"]
                current_data = df_all.iloc[selected_info["df_index"]]

                with st.form("edit_form"):
                    col1, col2 = st.columns(2)
                    with col1:
                        e_cultivar = st.text_input("1. Cultivar Name", value=str(current_data.get("Cultivar Name", "")))
                        e_supplier = st.text_input("2. Supplier Name", value=str(current_data.get("Supplier Name", "")))
                        e_supp_lot = st.text_input("3. Supplier Lot. No.", value=str(current_data.get("Supplier Lot. No.", "")))
                        e_export = st.text_input("4. Export Name", value=str(current_data.get("Export Name", "")))
                        e_thai_lot = st.text_input("5. THAI STICK Lot. No.", value=str(current_data.get("THAI STICK Lot. No.", "")))

                        cat_opt = ["FG (Finished Goods)", "WIP (Work in Progress)"]
                        old_cat = str(current_data.get("Category", "FG (Finished Goods)"))
                        e_category = st.selectbox("6. Category", cat_opt, index=cat_opt.index(old_cat) if old_cat in cat_opt else 0)

                        grades_opt = ["Big", "Small", "Mix", "F", "Grade A", "Premium"]
                        old_grade = str(current_data.get("Grade", "Big"))
                        e_grade = st.selectbox("7. Grade", grades_opt, index=grades_opt.index(old_grade) if old_grade in grades_opt else 0)

                    with col2:
                        old_rec_date = str(current_data.get("Receive date", current_data.get("Date receive", current_data.get("Date Receive", ""))))
                        try: default_rec_dt = pd.to_datetime(old_rec_date, dayfirst=True).date() if old_rec_date and str(old_rec_date).strip() else date.today()
                        except: default_rec_dt = date.today()
                        e_receive_date_str = str(st.date_input("8. Receive date", value=default_rec_dt, key=f"edit_receive_date_{matched_row_idx}"))

                        old_qty = float(current_data.get("Clean_Qty", 0.0))
                        e_qty = st.number_input("9. Quantity (g)", min_value=0.0, step=0.1, value=old_qty)
                        e_remark_qty = st.text_input("Remark Qty", value=str(current_data.get("Remark Qty", "")), key=f"edit_remark_qty_{matched_row_idx}")

                        old_insp = str(current_data.get("Inspection Req.", "No"))
                        insp_con = st.container()
                        # ✨ เติม _{matched_row_idx} ท้ายคีย์
                        e_is_tbd_insp = st.checkbox("To be decided", value=(old_insp == "To be decided"), key=f"edit_tbd_insp_{matched_row_idx}")
                        insp_opts = ["Yes", "No"]
                        e_insp_val = insp_con.selectbox("10. Inspection Req.", insp_opts, index=insp_opts.index(old_insp) if old_insp in insp_opts else 0, disabled=e_is_tbd_insp, key=f"edit_insp_{matched_row_idx}")
                        e_insp = "To be decided" if e_is_tbd_insp else e_insp_val

                        old_ir = str(current_data.get("IR Req.", "No"))
                        ir_con = st.container()
                        # ✨ เติม _{matched_row_idx} ท้ายคีย์
                        e_is_tbd_ir = st.checkbox("To be decided", value=(old_ir == "To be decided"), key=f"edit_tbd_ir_{matched_row_idx}")
                        ir_opts = ["Yes", "No"]
                        e_ir_val = ir_con.selectbox("11. IR Req.", ir_opts, index=ir_opts.index(old_ir) if old_ir in ir_opts else 0, disabled=e_is_tbd_ir, key=f"edit_ir_{matched_row_idx}")
                        e_ir = "To be decided" if e_is_tbd_ir else e_ir_val

                        old_date = str(current_data.get("Shipment date", ""))
                        ship_con = st.container()
                        # ✨ เติม _{matched_row_idx} ท้ายคีย์
                        e_is_tbd_ship = st.checkbox("To be decided", value=(old_date == "To be decided"), key=f"edit_tbd_ship_{matched_row_idx}")
                        try: default_dt = pd.to_datetime(old_date, dayfirst=True).date() if old_date != "To be decided" and old_date.strip() else date.today()
                        except: default_dt = date.today()
                        e_ship_val = ship_con.date_input("12. Shipment date", value=default_dt, disabled=e_is_tbd_ship, key=f"edit_ship_{matched_row_idx}")
                        e_shipment_date = "To be decided" if e_is_tbd_ship else str(e_ship_val)

                        old_type = str(current_data.get("Type", ""))
                        type_con = st.container()
                        # ✨ เติม _{matched_row_idx} ท้ายคีย์
                        e_is_tbd_type = st.checkbox("To be decided", value=(old_type == "To be decided"), key=f"edit_tbd_type_{matched_row_idx}")
                        type_opts = ["Inspection & Packing Service", "Trading Goods (TS)", "Consignment", "In House Product"]
                        e_type_val = type_con.selectbox("13. Type", type_opts, index=type_opts.index(old_type) if old_type in type_opts else 0, disabled=e_is_tbd_type, key=f"edit_type_{matched_row_idx}")
                        e_type = "To be decided" if e_is_tbd_type else e_type_val

                        status_opt = ["Trading Goods Received to Stock", "Inspection Completed", "Pending Packing", "Packing done", "Delivered", "Waiting for Return"]
                        old_status = str(current_data.get("Status", ""))
                        e_status = st.selectbox("14. Status", status_opt, index=status_opt.index(old_status) if old_status in status_opt else 0)

                    st.markdown("")
                    if st.form_submit_button("🔄 อัปเดตข้อมูล (Update Google Sheets)"):
                        latest_headers = [
                            "Cultivar Name", "Supplier Name", "Supplier Lot. No.", "Export Name", "THAI STICK Lot. No.", 
                            "Category", "Grade", "Receive date", "Quantity", "Remark Qty", "Inspection Req.", "IR Req.", 
                            "Shipment date", "Type", "Status", "Packing date"
                        ]
                        updated_dict = {
                            "Cultivar Name": e_cultivar, "Supplier Name": e_supplier, "Supplier Lot. No.": e_supp_lot,
                            "Export Name": e_export, "THAI STICK Lot. No.": e_thai_lot, "Category": e_category, "Grade": e_grade,
                            "Receive date": e_receive_date_str, "Quantity": str(e_qty), "Remark Qty": e_remark_qty,
                            "Inspection Req.": e_insp, "IR Req.": e_ir, "Shipment date": e_shipment_date,
                            "Type": e_type, "Status": e_status, "Packing date": "",
                        }
                        updated_row_data = [str(updated_dict.get(h, "")) for h in latest_headers]
                        col_end_letter = chr(64 + len(latest_headers)).upper()
                        
                        sheet.update(range_name=f"A{matched_row_idx}:{col_end_letter}{matched_row_idx}", values=[updated_row_data])
                        get_cached_data.clear()
                        st.success(f"อัปเดตข้อมูลของ Lot: {selected_lot} (Size: {e_grade}) เรียบร้อยแล้ว! 🎉")
                        time.sleep(1)
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
