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
# 2. ฟังก์ชัน Cache
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
            if not any(str(cell).strip() for cell in r):
                continue
            while len(r) < len(headers):
                r.append("")
            row_dict = {headers[i]: r[i] for i in range(len(headers))}
            if include_row_index:
                row_dict["Original_Row_Index"] = sheet_row
            data.append(row_dict)
        return data
    except Exception:
        return []

# ==========================================
# 3. ฟังก์ชันเชื่อมต่อ Google Sheets
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
    adjustment_log_sheet = get_or_create_worksheet(
        "Adjustment_Log", "1000", "8",
        ["Timestamp", "Source", "Lot No.", "Product Name", "Old Quantity", "New Quantity", "Reason", "Adjusted Date"]
    )
    picking_request_sheet = get_or_create_worksheet(
        "Picking_Request_Log", "1000", "14",
        ["Timestamp", "Source Category", "Source Lot", "Cultivar Name", "Target Grade", "Request Qty", 
         "Consignee", "True Consignee", "Shipping Reason", "Expected Shipment Date", "Status", "Note", "Packed FG Lot", "Shipped Out Ref"]
    )

    EXPECTED_HEADERS = [
        "Cultivar Name", "Supplier Name", "Supplier Lot. No.", "Export Name", "THAI STICK Lot. No.", 
        "Category", "Grade", "Receive date", "Quantity", "Remark Qty", "Inspection Req.", "IR Req.", 
        "Shipment date", "Type", "Status", "Packing date"
    ]
    current_headers = main_sheet.row_values(1)
    if not current_headers or len(current_headers) < len(EXPECTED_HEADERS):
        main_sheet.update(range_name="A1:P1", values=[EXPECTED_HEADERS])
        
    return client, sh, main_sheet, packing_records_sheet, packing_log_sheet, shipped_sheet, adjustment_log_sheet, picking_request_sheet

client, sh, sheet, packing_records_sheet, packing_sheet, shipped_sheet, adjustment_sheet, picking_request_sheet = init_google_sheets(SHEET_NAME)

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
            date_keywords = ["date", "วันที่", "timestamp"]
            if any(k in col_str for k in date_keywords):
                def format_date(x):
                    if pd.isna(x) or str(x).strip().casefold() in DELIVERY_PENDING_VALUES:
                        return x
                    try:
                        d_str = str(x).strip().split(" ")[0]
                        if re.match(r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}$', d_str): dt = pd.to_datetime(d_str)
                        else: dt = pd.to_datetime(d_str, dayfirst=True, errors='coerce')
                        if pd.notna(dt): return dt.strftime('%d/%m/%Y')
                        return x
                    except: return x
                df_copy[col] = df_copy[col].apply(format_date)

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

def render_section_heading(title, description, eyebrow="", tone="blue"):
    st.markdown(
        f'<div style="margin-bottom: 15px;"><span style="color: #64748b; font-size: 12px; font-weight: bold; letter-spacing: 1px;">{eyebrow}</span>'
        f'<h3 style="margin: 0; padding: 0; color: #1e293b;">{title}</h3>'
        f'<p style="margin: 0; color: #64748b; font-size: 14px;">{description}</p></div>',
        unsafe_allow_html=True,
    )

def render_grade_summary(stock_df, total, stock_type):
    grade_names = stock_df["Grade"].astype(str).str.strip().str.upper()
    items = [(grade, stock_df.loc[grade_names.eq(grade), "Clean_Qty"].sum()) for grade in ("BIG", "SMALL", "MIX", "F")]
    items.append((f"รวม {stock_type}", total))
    html = ['<div style="display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 15px;">']
    for label, quantity in items:
        html.append(f'<div style="padding: 10px 15px; border: 1px solid #e2e8f0; border-radius: 8px; background: #f8fafc;"><div style="font-size: 12px; color: #64748b; font-weight: bold;">{escape(label)}</div><div style="font-size: 18px; font-weight: bold; color: #334155;">{quantity:,.2f} g</div></div>')
    html.append('</div>')
    st.markdown("".join(html), unsafe_allow_html=True)


# ==========================================
# 5. Sidebar เมนูด้านซ้าย
# ==========================================
st.sidebar.title("🌱 Thai Stick Trading")
st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "เลือกเมนูการใช้งาน",
    ["🚚 Delivery & Picking Request", "📊 Stock Summary & Details", "📈 Monthly Report", "📝 Trading Goods Registration", "✏️ Edit Registration", "🔧 Adjustment Module", "📦 Packing Module", "🚚 Shipped Out Module"],
)
st.sidebar.markdown("---")


# ==========================================
# หน้าที่ 1: Delivery & Picking Request
# ==========================================
if menu == "🚚 Delivery & Picking Request":
    st.title("🚚 Delivery & Picking Request")
    st.caption("ระบบจัดการใบเบิกสินค้า (Picking Request) และติดตามกำหนดการจัดส่ง")

    tab_dash, tab_create, tab_legacy = st.tabs([
        "📋 1. Active Delivery Dashboard", 
        "📝 2. สร้าง Picking Request", 
        "🗄️ 3. Legacy Delivery Schedule"
    ])

    with tab_dash:
        st.subheader("📋 สถานะใบเบิกสินค้าและการจัดส่ง (Active Requests)")
        pr_data = get_cached_data("picking_request", picking_request_sheet, include_row_index=True)
        if pr_data:
            df_pr = pd.DataFrame(pr_data)
            df_pr["Clean_Qty"] = df_pr["Request Qty"].apply(clean_quantity)
            
            active_df = df_pr[~df_pr["Status"].astype(str).str.contains("Delivered|จัดส่งแล้ว", case=False, na=False)].copy()
            
            if not active_df.empty:
                active_df["Parsed_Date"] = active_df["Expected Shipment Date"].apply(parse_delivery_date)
                active_df = active_df.sort_values(by="Parsed_Date", na_position="last")
                
                col1, col2, col3 = st.columns(3)
                with col1: st.metric("ใบงานที่รอแพ็ก (WIP)", len(active_df[active_df["Status"].str.contains("รอแพ็ค")]))
                with col2: st.metric("ใบงานพร้อมส่ง (FG)", len(active_df[active_df["Status"].str.contains("พร้อมส่ง")]))
                with col3: st.metric("ปริมาณรอส่งรวม (g)", f"{active_df['Clean_Qty'].sum():,.2f}")
                
                st.dataframe(format_df_quantity(active_df.drop(columns=["Original_Row_Index", "Parsed_Date", "Clean_Qty"], errors='ignore')), use_container_width=True)
            else:
                st.info("ไม่มีรายการจัดส่งที่ค้างอยู่ (Active) ในขณะนี้")
        else:
            st.info("ยังไม่มีข้อมูล Picking Request ในระบบ")

    with tab_create:
        st.subheader("📝 สร้างใบเบิกสินค้า (Create Picking Request)")
        st.markdown("ดึงรายการสต็อกคงเหลือ (WIP/FG) มากะบุเป้าหมายการแพ็ก/จัดส่ง โดยสต็อกจริงจะยังไม่ถูกหักจนกว่าจะทำ Packing หรือ Shipped Out")
        
        data_main = get_cached_data("main", sheet)
        df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()
        records_data = get_cached_data("packing", packing_records_sheet)
        df_packing_fg = pd.DataFrame(records_data) if records_data else pd.DataFrame()

        avail_options = []
        avail_map = {}

        if not df_main.empty:
            df_main["Clean_Qty"] = df_main["Quantity"].apply(clean_quantity) if "Quantity" in df_main.columns else 0.0
            main_avail = df_main[(df_main["Clean_Qty"] > 0) & (~df_main["Status"].astype(str).str.strip().str.lower().eq("delivered"))]
            for _, r in main_avail.iterrows():
                cat = str(r.get("Category", ""))
                lot_no = str(r.get("THAI STICK Lot. No.", "")).strip() or str(r.get("Supplier Lot. No.", "")).strip() or "No-Lot"
                prod = str(r.get("Cultivar Name", ""))
                disp = f"[{cat}] Lot: {lot_no} | {prod} | Qty: {r['Clean_Qty']:,.2f}g"
                avail_options.append(disp)
                avail_map[disp] = {"lot": lot_no, "cat": cat, "prod": prod}

        if not df_packing_fg.empty:
            df_packing_fg["Clean_Qty"] = df_packing_fg["Quantity"].apply(clean_quantity) if "Quantity" in df_packing_fg.columns else 0.0
            fg_avail = df_packing_fg[df_packing_fg["Clean_Qty"] > 0]
            for _, r in fg_avail.iterrows():
                lot_no = str(r.get("Generated FG Lot No.", "")).strip() or str(r.get("THAI STICK Lot. No.", "")).strip()
                prod = str(r.get("Export Name", ""))
                disp = f"[FG Packed] Lot: {lot_no} | {prod} | Qty: {r['Clean_Qty']:,.2f}g"
                avail_options.append(disp)
                avail_map[disp] = {"lot": lot_no, "cat": "FG", "prod": prod}

        selected_sources = st.multiselect("🔍 1. เลือก WIP / FG ต้นทาง (เลือกได้หลายรายการ)", avail_options)
        
        if selected_sources:
            with st.form("create_picking_request_form"):
                st.markdown("##### 2. ระบุรายละเอียดของแต่ละรายการที่เลือก")
                request_items = []
                for disp in selected_sources:
                    info = avail_map[disp]
                    st.markdown(f"**📍 {disp}**")
                    c1, c2 = st.columns(2)
                    with c1: t_grade = st.selectbox(f"เกรดที่ต้องการ (Target Grade)", ["Big", "Small", "Mix", "F", "Grade A", "Premium", "ระบุเอง"], key=f"tg_{disp}")
                    with c2: req_qty = st.number_input(f"จำนวนที่ต้องการเบิก (Request Qty)", min_value=0.0, step=0.1, key=f"rq_{disp}")
                    if t_grade == "ระบุเอง": t_grade = st.text_input("ระบุเกรด:", key=f"tg_man_{disp}")
                    request_items.append({"source_lot": info["lot"], "cat": info["cat"], "prod": info["prod"], "t_grade": t_grade, "req_qty": req_qty})
                
                st.divider()
                st.markdown("##### 3. ข้อมูลการจัดส่ง")
                co1, co2 = st.columns(2)
                with co1: 
                    consignee = st.text_input("ชื่อปลายทาง (Consignee) *")
                    shipping_reason = st.selectbox("เหตุผลในการส่ง", ["EXPORT", "DOMESTIC SALE", "SAMPLE", "RETURN", "REJECT"])
                with co2: 
                    true_consignee = st.text_input("ลูกค้าตัวจริง (True Consignee / Client) *")
                    exp_date = st.date_input("กำหนดการจัดส่ง (Expected Shipment Date)", value=date.today())
                
                note = st.text_area("หมายเหตุเพิ่มเติม (Note)")
                
                if st.form_submit_button("💾 สร้างใบ Picking Request"):
                    if not consignee.strip() or not true_consignee.strip():
                        st.error("❌ กรุณาระบุ Consignee และ True Consignee ให้ครบถ้วน")
                    else:
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        rows_to_insert = []
                        for item in request_items:
                            if item["req_qty"] <= 0: continue
                            status = "📦 รอแพ็ค (Pending Pack)" if "WIP" in item["cat"].upper() else "✅ พร้อมส่ง (Ready to Ship)"
                            rows_to_insert.append([
                                timestamp, item["cat"], item["source_lot"], item["prod"], item["t_grade"], item["req_qty"],
                                consignee, true_consignee, shipping_reason, str(exp_date), status, note, "", ""
                            ])
                        
                        if rows_to_insert:
                            picking_request_sheet.append_rows(rows_to_insert, value_input_option='USER_ENTERED')
                            get_cached_data.clear()
                            st.success(f"✅ สร้าง Picking Request สำเร็จ {len(rows_to_insert)} รายการ!")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.warning("⚠️ กรุณาระบุจำนวนที่ต้องการเบิก (Request Qty) ให้มากกว่า 0")

    with tab_legacy:
        st.subheader("🗄️ ข้อมูลจากชีต Delivery_schedule เดิม")
        try:
            del_sheet = sh.worksheet("Delivery_schedule")
            data_del = get_cached_data("delivery", del_sheet, include_row_index=True)
            df_del = pd.DataFrame(data_del) if data_del else pd.DataFrame()
            if not df_del.empty: st.dataframe(format_df_quantity(df_del.drop(columns=["Original_Row_Index"], errors='ignore')), use_container_width=True)
            else: st.info("ไม่มีข้อมูลในชีตนี้")
        except: st.error("❌ ไม่พบแท็บชื่อ 'Delivery_schedule'")


# ==========================================
# หน้าที่ 2: Stock Summary & Details
# ==========================================
elif menu == "📊 Stock Summary & Details":
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

        wip_stock_df = filtered_df[filtered_df["Category"].astype(str).str.contains("WIP", case=False, na=False)] if "Category" in filtered_df.columns else pd.DataFrame()
        fg_stock_df = filtered_df[filtered_df["Category"].astype(str).str.contains("FG", case=False, na=False)] if "Category" in filtered_df.columns else pd.DataFrame()

        total_wip = wip_stock_df["Clean_Qty"].sum()
        total_fg = fg_stock_df["Clean_Qty"].sum()
        total_on_hand = total_wip + total_fg

        render_section_heading("ภาพรวมสต็อกคงเหลือ", "ยอดคงเหลือจากสินค้าในคลังและรายการ Packing", eyebrow="ON-HAND STOCK", tone="slate")
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1: st.metric("สต็อกคงเหลือทั้งหมด", f"{total_on_hand:,.2f} g")
        with col_m2: st.metric("งานระหว่างกระบวนการ · WIP", f"{total_wip:,.2f} g")
        with col_m3: st.metric("สินค้าสำเร็จรูป · FG", f"{total_fg:,.2f} g")

        st.divider()

        for stock_type, stock_df, stock_total, tone, title in (("WIP", wip_stock_df, total_wip, "amber", "รายละเอียด WIP"), ("FG", fg_stock_df, total_fg, "teal", "รายละเอียด FG")):
            render_section_heading(title, f"สต็อกคงเหลือของ {stock_type}", eyebrow=f"{stock_type} INVENTORY", tone=tone)
            if not stock_df.empty:
                render_grade_summary(stock_df, stock_total, stock_type)
                st.dataframe(format_df_quantity(stock_df), width="stretch", hide_index=True)
            else:
                st.info(f"ยังไม่มีรายการสินค้า {stock_type} คงเหลือในขณะนี้")
    else:
        st.info("ยังไม่มีข้อมูลสต็อกในระบบ")

# ==========================================
# หน้าที่ 3: Monthly Report
# ==========================================
elif menu == "📈 Monthly Report":
    st.title("📈 Monthly Report (รายงานสรุปผลการดำเนินงานรายเดือน)")
    
    today = date.today()
    col_y, col_m = st.columns(2)
    with col_y:
        year_options = [2026, 2025, 2024, 2023]
        selected_year = st.selectbox("เลือกปี (Year)", year_options, index=year_options.index(today.year) if today.year in year_options else 0)
    with col_m:
        months_dict = {"มกราคม": 1, "กุมภาพันธ์": 2, "มีนาคม": 3, "เมษายน": 4, "พฤษภาคม": 5, "มิถุนายน": 6, "กรกฎาคม": 7, "สิงหาคม": 8, "กันยายน": 9, "ตุลาคม": 10, "พฤศจิกายน": 11, "ธันวาคม": 12}
        month_names = list(months_dict.keys())
        selected_month_label = st.selectbox("เลือกเดือน (Month)", month_names, index=today.month - 1)
        selected_month_num = months_dict[selected_month_label]

    st.divider()

    shipped_data = get_cached_data("shipped", shipped_sheet)
    df_shipped = pd.DataFrame(shipped_data) if shipped_data else pd.DataFrame()
    if not df_shipped.empty:
        df_shipped["Parsed_Date"] = df_shipped["Shipment Date"].apply(lambda x: pd.to_datetime(str(x), errors='coerce'))
        df_shipped["Clean_Shipped_Qty"] = df_shipped["Shipped Qty"].apply(clean_quantity)
        monthly_shipped = df_shipped[(df_shipped["Parsed_Date"].dt.year == selected_year) & (df_shipped["Parsed_Date"].dt.month == selected_month_num)]
    else:
        monthly_shipped = pd.DataFrame()

    st.subheader(f"🚚 Shipped Out Report ประจำเดือน {selected_month_label} {selected_year}")
    if not monthly_shipped.empty:
        reason_summary = monthly_shipped.groupby("Shipping Reason")["Clean_Shipped_Qty"].sum().reset_index()
        reason_summary.columns = ["Shipping Reason (เหตุผลการส่ง)", "Total Shipped Qty (จำนวนรวม)"]
        st.dataframe(format_df_quantity(reason_summary), use_container_width=True)
        st.dataframe(format_df_quantity(monthly_shipped), use_container_width=True)
    else:
        st.info(f"ไม่พบประวัติการส่งออกในเดือน {selected_month_label} {selected_year}")

# ==========================================
# หน้าที่ 4: Trading Goods Registration
# ==========================================
elif menu == "📝 Trading Goods Registration":
    st.title("📝 Trading Goods Registration")
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
            receive_date_str = str(st.date_input("8. Receive date", value=date.today()))
            quantity = st.number_input("9. Quantity (g)", min_value=0.0, step=0.1)
            remark_qty = st.text_input("Remark Qty")
            inspection_req = st.selectbox("10. Inspection Req.", ["Yes", "No", "To be decided"])
            ir_req = st.selectbox("11. IR Req.", ["Yes", "No", "To be decided"])
            shipment_date = str(st.date_input("12. Shipment date", value=date.today()))
            prod_type = st.selectbox("13. Type", ["Trading Goods (TS)", "Inspection & Packing Service", "Consignment", "In House Product"])
            status = st.selectbox("14. Status", ["Trading Goods Received to Stock", "Inspection Completed", "Pending Packing", "Packing done", "Delivered", "Waiting for Return"])
        
        if st.form_submit_button("💾 บันทึกข้อมูลลง Google Sheets"):
            row_dict = {
                "Cultivar Name": cultivar_name, "Supplier Name": supplier_name, "Supplier Lot. No.": supplier_lot,
                "Export Name": export_name, "THAI STICK Lot. No.": thai_stick_lot, "Category": category, "Grade": grade,
                "Receive date": receive_date_str, "Quantity": float(quantity), "Remark Qty": remark_qty,
                "Inspection Req.": inspection_req, "IR Req.": ir_req, "Shipment date": shipment_date,
                "Type": prod_type, "Status": status, "Packing date": "",
            }
            latest_headers = [str(header).strip() for header in sheet.row_values(1)]
            row_to_insert = [row_dict.get(h, "") for h in latest_headers]
            sheet.append_row(values=row_to_insert, value_input_option='USER_ENTERED')
            get_cached_data.clear()
            st.success("บันทึกข้อมูลสำเร็จเรียบร้อยแล้ว! 🎉")
            time.sleep(1)
            st.rerun()

# ==========================================
# หน้าที่ 5: Edit Registration
# ==========================================
elif menu == "✏️ Edit Registration":
    st.title("✏️ Edit Registration Data")
    data = get_cached_data("main", sheet, include_row_index=True)
    if data:
        df_all = pd.DataFrame(data)
        df_all["Clean_Qty"] = df_all["Quantity"].apply(clean_quantity) if "Quantity" in df_all.columns else 0.0
        editable_df = df_all[(df_all["Clean_Qty"] > 0) & (~df_all["Status"].astype(str).str.strip().str.lower().eq("delivered"))].copy()

        if editable_df.empty:
            st.info("ไม่มีล็อตสินค้าที่สามารถแก้ไขได้ในขณะนี้")
        else:
            edit_options = []
            edit_map = {}
            for _, r in editable_df.iterrows():
                lot_no = str(r.get("THAI STICK Lot. No.", "")) or str(r.get("Supplier Lot. No.", ""))
                disp = f"Lot: {lot_no} | Size: {r.get('Grade','')} | {r.get('Cultivar Name','')} | Qty: {r['Clean_Qty']:,.2f}"
                edit_options.append(disp)
                edit_map[disp] = {"row_idx": r["Original_Row_Index"], "lot_no": lot_no, "data": r}

            selected_disp = st.selectbox("🔍 เลือกข้อมูลที่ต้องการแก้ไข", edit_options)

            if selected_disp:
                matched_row_idx = edit_map[selected_disp]["row_idx"]
                current_data = edit_map[selected_disp]["data"]

                with st.form("edit_form"):
                    col1, col2 = st.columns(2)
                    with col1:
                        e_cultivar = st.text_input("1. Cultivar Name", value=str(current_data.get("Cultivar Name", "")))
                        e_supplier = st.text_input("2. Supplier Name", value=str(current_data.get("Supplier Name", "")))
                        e_thai_lot = st.text_input("5. THAI STICK Lot. No.", value=str(current_data.get("THAI STICK Lot. No.", "")))
                        e_category = st.selectbox("6. Category", ["FG (Finished Goods)", "WIP (Work in Progress)"])
                        e_grade = st.selectbox("7. Grade", ["Big", "Small", "Mix", "F", "Grade A", "Premium"])
                    with col2:
                        e_qty = st.number_input("9. Quantity (g)", min_value=0.0, step=0.1, value=float(current_data.get("Clean_Qty", 0.0)))
                        e_status = st.selectbox("14. Status", ["Trading Goods Received to Stock", "Inspection Completed", "Pending Packing", "Packing done", "Delivered", "Waiting for Return"])

                    if st.form_submit_button("🔄 อัปเดตข้อมูล"):
                        latest_headers = [str(header).strip() for header in sheet.row_values(1)]
                        updated_dict = current_data.to_dict()
                        updated_dict.update({"Cultivar Name": e_cultivar, "Supplier Name": e_supplier, "THAI STICK Lot. No.": e_thai_lot, "Category": e_category, "Grade": e_grade, "Quantity": str(e_qty), "Status": e_status})
                        updated_row_data = [str(updated_dict.get(h, "")) for h in latest_headers]
                        col_end_letter = chr(64 + len(latest_headers)).upper()
                        sheet.update(range_name=f"A{matched_row_idx}:{col_end_letter}{matched_row_idx}", values=[updated_row_data])
                        get_cached_data.clear()
                        st.success("อัปเดตข้อมูลเรียบร้อยแล้ว!")
                        time.sleep(1)
                        st.rerun()

# ==========================================
# หน้าที่ 5.5: Adjustment Module
# ==========================================
elif menu == "🔧 Adjustment Module":
    st.title("🔧 Adjustment Module (WIP & FG)")
    data_main = get_cached_data("main", sheet, include_row_index=True)
    df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()
    
    records_data = get_cached_data("packing", packing_records_sheet, include_row_index=True)
    df_packing_fg = pd.DataFrame(records_data) if records_data else pd.DataFrame()

    adj_options = []
    adj_map = {}

    if not df_main.empty:
        df_main["Clean_Qty"] = df_main["Quantity"].apply(clean_quantity) if "Quantity" in df_main.columns else 0.0
        for _, r in df_main.iterrows():
            lot_no = str(r.get("THAI STICK Lot. No.", "")).strip() or "No-Lot"
            disp = f"[Main] Lot: {lot_no} | {r.get('Cultivar Name','')} | Qty: {r['Clean_Qty']:,.2f}"
            adj_options.append(disp)
            adj_map[disp] = {"source": "main", "row_idx": r["Original_Row_Index"], "lot_no": lot_no, "current_qty": r["Clean_Qty"], "product": r.get('Cultivar Name','')}

    if not df_packing_fg.empty:
        df_packing_fg["Clean_Qty"] = df_packing_fg["Quantity"].apply(clean_quantity) if "Quantity" in df_packing_fg.columns else 0.0
        for _, r in df_packing_fg.iterrows():
            lot_no = str(r.get("Generated FG Lot No.", "")).strip() or "No-Lot"
            disp = f"[Packed FG] Lot: {lot_no} | {r.get('Export Name','')} | Qty: {r['Clean_Qty']:,.2f}"
            adj_options.append(disp)
            adj_map[disp] = {"source": "packing_records", "row_idx": r["Original_Row_Index"], "lot_no": lot_no, "current_qty": r["Clean_Qty"], "product": r.get('Export Name','')}

    selected_item = st.selectbox("🔍 เลือกรายการที่ต้องการปรับแก้ยอด", ["-- กรุณาเลือก --"] + adj_options)
    
    if selected_item != "-- กรุณาเลือก --":
        info = adj_map[selected_item]
        with st.form("adjustment_form"):
            new_qty = st.number_input("จำนวนยอดที่ถูกต้อง (New Quantity)", min_value=0.0, step=0.1, value=float(info["current_qty"]))
            remark = st.text_input("เหตุผลในการปรับยอด")
            if st.form_submit_button("💾 ยืนยันการปรับแก้ยอด"):
                try:
                    if info["source"] == "main":
                        headers = [str(h).strip() for h in sheet.row_values(1)]
                        qty_idx = headers.index("Quantity") + 1
                        sheet.update_cell(info["row_idx"], qty_idx, new_qty)
                    elif info["source"] == "packing_records":
                        headers = [str(h).strip() for h in packing_records_sheet.row_values(1)]
                        qty_idx = headers.index("Quantity") + 1
                        packing_records_sheet.update_cell(info["row_idx"], qty_idx, new_qty)

                    adjustment_sheet.append_row([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), info["source"], info["lot_no"], info["product"], info["current_qty"], new_qty, remark, str(date.today())], value_input_option='USER_ENTERED')
                    get_cached_data.clear()
                    st.success("ปรับแก้ยอดสำเร็จ!")
                    time.sleep(1)
                    st.rerun()
                except Exception as e: st.error(f"Error: {e}")

# ==========================================
# หน้าที่ 6: Packing Module (อัปเดตสถานะ Auto)
# ==========================================
elif menu == "📦 Packing Module":
    st.title("📦 Packing Module (WIP Conversion & Status Sync)")
    
    data = get_cached_data("main", sheet, include_row_index=True)
    if data:
        df = pd.DataFrame(data)
        wip_df = df[df["Category"].astype(str).str.contains("WIP", case=False, na=False)].copy() if "Category" in df.columns else pd.DataFrame()

        if not wip_df.empty:
            wip_df["Clean_Qty"] = wip_df["Quantity"].apply(clean_quantity)
            available_wip = wip_df[wip_df["Clean_Qty"] > 0].copy()
            
            wip_options = []
            wip_map = {}
            for _, r in available_wip.iterrows():
                lot_no = str(r.get("THAI STICK Lot. No.", "")).strip() or str(r.get("Supplier Lot. No.", "")).strip()
                disp = f"Lot: {lot_no} | {r.get('Cultivar Name','')} | Qty: {r['Clean_Qty']:,.2f}"
                wip_options.append(disp)
                wip_map[disp] = {"lot_no": lot_no, "qty": r["Clean_Qty"], "row_idx": r["Original_Row_Index"], "supplier": r.get("Supplier Name", "")}

            selected_wip_disps = st.multiselect("📦 1. เลือก WIP ที่ต้องการแพ็ก", wip_options)
            
            if selected_wip_disps:
                with st.form("packing_form", clear_on_submit=True):
                    wip_withdrawals = {disp: st.number_input(f"เบิกจาก {disp}", min_value=0.0, max_value=wip_map[disp]["qty"], value=float(wip_map[disp]["qty"])) for disp in selected_wip_disps}
                    
                    st.divider()
                    packed_fg_name = st.text_input("ชื่อผลิตภัณฑ์หลังแพ็ก")
                    g1_type = st.selectbox("เกรด", ["Big", "Small", "Mix", "F", "Grade A"]); g1_qty = st.number_input("จำนวนแพ็ก", min_value=0.0)
                    waste_qty = st.number_input("จำนวนของเสีย (Waste Qty)", min_value=0.0)
                    packing_date_str = str(date.today())
                    
                    if st.form_submit_button("🚀 ยืนยันการแพ็ก"):
                        combined_source_lots = []
                        sheet_headers = [str(h).strip() for h in sheet.row_values(1)]
                        col_end_letter = chr(64 + len(sheet_headers)).upper()
                        
                        for disp, w_qty in wip_withdrawals.items():
                            if w_qty > 0:
                                info = wip_map[disp]
                                combined_source_lots.append(info["lot_no"])
                                
                                current_row_values = sheet.row_values(info["row_idx"])
                                while len(current_row_values) < len(sheet_headers): current_row_values.append("")
                                r_dict = {sheet_headers[i]: current_row_values[i] for i in range(len(sheet_headers))}
                                
                                remaining_wip = max(0.0, info["qty"] - w_qty)
                                r_dict["Quantity"] = str(float(remaining_wip))
                                r_dict["Status"] = "Packing done" if remaining_wip <= 0 else "Pending Packing"
                                
                                row_values = [str(r_dict.get(h, "")) for h in sheet_headers]
                                sheet.update(range_name=f"A{info['row_idx']}:{col_end_letter}{info['row_idx']}", values=[row_values], value_input_option='USER_ENTERED')
                        
                        source_lots_str = ",".join(combined_source_lots)
                        fg_lot = f"{source_lots_str}-FG-{datetime.now().strftime('%S')}"
                        
                        log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        packing_records_sheet.append_row([log_time, source_lots_str, "", "", packed_fg_name, fg_lot, g1_type, float(g1_qty), float(waste_qty), packing_date_str, ""], value_input_option='USER_ENTERED')
                        
                        # Auto-Update Picking Request Status
                        pr_data = get_cached_data("picking_request", picking_request_sheet, include_row_index=True)
                        if pr_data:
                            pr_headers = [str(h).strip() for h in picking_request_sheet.row_values(1)]
                            status_col_idx = pr_headers.index("Status") + 1 if "Status" in pr_headers else 11
                            for pr_row in pr_data:
                                if pr_row.get("Source Lot", "") in combined_source_lots and "รอแพ็ค" in str(pr_row.get("Status", "")):
                                    picking_request_sheet.update_cell(pr_row["Original_Row_Index"], status_col_idx, "✅ พร้อมส่ง (Ready to Ship)")

                        get_cached_data.clear()
                        st.success("✅ แพ็กสำเร็จ และอัปเดตสถานะใบเบิก (ถ้ามี) เป็นพร้อมส่งแล้ว!")
                        time.sleep(1)
                        st.rerun()

# ==========================================
# หน้าที่ 7: Shipped Out Module
# ==========================================
elif menu == "🚚 Shipped Out Module":
    st.title("🚚 Shipped Out Module")
    data_main = get_cached_data("main", sheet, include_row_index=True)
    df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()
    records_data = get_cached_data("packing", packing_records_sheet, include_row_index=True)
    df_packing_fg = pd.DataFrame(records_data) if records_data else pd.DataFrame()

    combined_options = []
    inventory_map = {}

    if not df_main.empty and "Category" in df_main.columns:
        df_main["Clean_Qty"] = df_main["Quantity"].apply(clean_quantity)
        main_avail = df_main[(df_main["Clean_Qty"] > 0) & (~df_main["Status"].astype(str).str.strip().str.lower().eq("delivered"))]
        for _, r in main_avail.iterrows():
            cat = str(r.get("Category", ""))
            lot_no = str(r.get("THAI STICK Lot. No.", "")).strip() or "No-Lot"
            disp = f"[Main - {cat}] {lot_no} | Qty: {r['Clean_Qty']:,.1f}"
            combined_options.append(disp)
            inventory_map[disp] = {"category": cat, "lot_no": lot_no, "max_qty": r["Clean_Qty"], "source": "main", "row_idx": r["Original_Row_Index"], "product": r.get("Cultivar Name",""), "grade": r.get("Grade","")}

    if not df_packing_fg.empty:
        df_packing_fg["Clean_Qty"] = df_packing_fg["Quantity"].apply(clean_quantity)
        fg_avail = df_packing_fg[df_packing_fg["Clean_Qty"] > 0]
        for _, r in fg_avail.iterrows():
            lot_no = str(r.get("Generated FG Lot No.", "")).strip() or "No-Lot"
            disp = f"[FG - Packed] {lot_no} | Qty: {r['Clean_Qty']:,.1f}"
            combined_options.append(disp)
            inventory_map[disp] = {"category": "FG (Finished Goods)", "lot_no": lot_no, "max_qty": r["Clean_Qty"], "source": "packing_records", "row_idx": r["Original_Row_Index"], "product": r.get("Export Name",""), "grade": r.get("Grade","")}

    with st.form("shipped_out_multi_form"):
        consignee = st.text_input("ชื่อปลายทาง (Consignee)")
        client_name = st.text_input("ลูกค้า (Client)")
        shipping_reason = st.selectbox("เหตุผลในการส่ง", ["EXPORT", "DOMESTIC SALE", "SAMPLE"])
        shipment_date_str = str(st.date_input("วันที่จัดส่ง", value=date.today()))

        chosen_display = st.selectbox(f"เลือกรายการสินค้า", ["-- กรุณาเลือก --"] + combined_options)
        max_q = inventory_map.get(chosen_display, {}).get("max_qty", 0.0) if chosen_display and not chosen_display.startswith("--") else 0.0
        ship_q = st.number_input(f"จำนวนส่งออก (Max: {max_q:,.1f})", min_value=0.0, max_value=max_q if max_q > 0 else 100000.0, step=0.1)
        
        if st.form_submit_button("🚚 ยืนยันการส่งสินค้าออก"):
            if not consignee or ship_q <= 0:
                st.error("กรุณากรอกข้อมูลให้ครบถ้วนและจำนวนส่งออกต้องมากกว่า 0")
            else:
                item = inventory_map[chosen_display]
                if item["source"] == "main":
                    sheet_headers = [str(h).strip() for h in sheet.row_values(1)]
                    current_row_values = sheet.row_values(item["row_idx"])
                    while len(current_row_values) < len(sheet_headers): current_row_values.append("")
                    row_dict = {sheet_headers[i]: current_row_values[i] for i in range(len(sheet_headers))}
                    remaining = max(0.0, clean_quantity(row_dict.get("Quantity", 0)) - ship_q)
                    row_dict["Quantity"] = str(remaining)
                    row_dict["Status"] = "Delivered" if remaining <= 0 else row_dict["Status"]
                    sheet.update(range_name=f"A{item['row_idx']}:{chr(64 + len(sheet_headers)).upper()}{item['row_idx']}", values=[[str(row_dict.get(h, "")) for h in sheet_headers]], value_input_option='USER_ENTERED')
                
                elif item["source"] == "packing_records":
                    pack_headers = [str(h).strip() for h in packing_records_sheet.row_values(1)]
                    qty_col_idx = pack_headers.index("Quantity") + 1
                    all_pack_records = get_cached_data("packing", packing_records_sheet)
                    current_qty = clean_quantity(all_pack_records[item["row_idx"] - 2].get("Quantity", 0))
                    packing_records_sheet.update_cell(item["row_idx"], qty_col_idx, max(0.0, current_qty - ship_q))
                
                shipped_sheet.append_row([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), item["category"], item["lot_no"], item["product"], item["grade"], str(ship_q), consignee, client_name, shipping_reason, shipment_date_str, ""], value_input_option='USER_ENTERED')
                
                # Auto update Picking Request Log to Delivered
                pr_data = get_cached_data("picking_request", picking_request_sheet, include_row_index=True)
                if pr_data:
                    pr_headers = [str(h).strip() for h in picking_request_sheet.row_values(1)]
                    status_col_idx = pr_headers.index("Status") + 1
                    for pr_row in pr_data:
                        if pr_row.get("Source Lot", "") == item["lot_no"] and pr_row.get("Consignee", "") == consignee:
                            picking_request_sheet.update_cell(pr_row["Original_Row_Index"], status_col_idx, "📦 Delivered (จัดส่งแล้ว)")
                            
                get_cached_data.clear()
                st.success("บันทึกการส่งออกและตัดสต็อกสำเร็จ!")
                time.sleep(1)
                st.rerun()
