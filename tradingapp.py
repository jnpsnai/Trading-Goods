from datetime import date, datetime
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import streamlit as st
import re

# --- 1. ตั้งค่าหน้าเว็บ Streamlit ให้เป็น Wide Mode ---
st.set_page_config(
    page_title="Thai Stick Trading Portal", page_icon="🌱", layout="wide"
)

# --- 2. Custom CSS ---
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

# --- 3. ฟังก์ชัน Cache การดึงข้อมูลให้ปลอดภัย (ดึงเป็น List of Dict) ---
@st.cache_data(ttl=30)
def get_cached_data(_sheet):
    try:
        rows = _sheet.get_all_values()
        if not rows or len(rows) <= 1:
            return []
        headers = [str(h).strip() for h in rows[0]]
        data = []
        for r in rows[1:]:
            while len(r) < len(headers):
                r.append("")
            row_dict = {headers[i]: r[i] for i in range(len(headers))}
            data.append(row_dict)
        return data
    except Exception as e:
        return []

# --- 4. ฟังก์ชันเชื่อมต่อและตรวจสอบโครงสร้าง Google Sheets แบบเรียกครั้งเดียว (ประหยัด API) ---
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
    
    # ดึงชีตทั้งหมดไว้ใน Cache
    main_sheet = sh.get_worksheet(0)
    
    # ฟังก์ชันตัวช่วยสำหรับสร้างชีตถ้ายังไม่มี
    def get_or_create_worksheet(title, rows, cols, headers):
        try:
            ws = sh.worksheet(title)
        except gspread.exceptions.WorksheetNotFound:
            ws = sh.add_worksheet(title=title, rows=rows, cols=cols)
            ws.append_row(headers, table_range="A1")
        return ws

    packing_records_sheet = get_or_create_worksheet(
        "Packing_Records", "1000", "15", 
        ["Timestamp", "Source WIP Lot", "Cultivar Name", "Export Name", "Generated FG Lot No.", "Grade", "Quantity", "Waste Qty", "Packing Date", "Note"]
    )
    packing_log_sheet = get_or_create_worksheet(
        "Packing_Log", "1000", "10", 
        ["Timestamp", "Source WIP Lot", "Cultivar", "Packed FG Name", "Grade & Qty Breakdown", "Total Packed Qty", "Waste Qty", "Packing Date", "Note"]
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

# เรียกใช้งานแค่บรรทัดเดียว ประหยัดโควตา API สุดๆ
client, sh, sheet, packing_records_sheet, packing_sheet, shipped_sheet = init_google_sheets(SHEET_NAME)

# --- ฟังก์ชันตัวช่วย (Helper Functions) ---
def clean_quantity(val):
    try: return float(str(val).replace(",", "").strip())
    except: return 0.0

def format_df_quantity(df):
    if df is not None and not df.empty:
        df_copy = df.copy()
        for col in df_copy.columns:
            col_str = str(col).lower()
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

# --- 5. Sidebar เมนูด้านซ้าย ---
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
    col_head1, col_head2 = st.columns([7, 3])
    with col_head1:
        st.title("🚨 Delivery Issue - รายการที่ต้องจัดส่ง")
        st.markdown("แสดงรายการและสถานะความพร้อมในการจัดส่ง (ดึงจากแท็บ **Delivery_schedule**)")
    with col_head2:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("🔗 Direct Google Sheets (Delivery_schedule)"):
            pwd = st.text_input("รหัสผ่าน:", type="password", key="gsheet_shortcut_pwd_del", placeholder="กรอกรหัสผ่าน")
            if pwd == "tstd": st.markdown("[🟢 **คลิกที่นี่เพื่อเปิด Google Sheets**](https://docs.google.com/spreadsheets/d/1RLtRn7gojLLLxrsm9mSFRLy7eGmCYRn5bgsoCmuju7Y/edit)")
            elif pwd != "": st.error("รหัสผ่านไม่ถูกต้อง ❌")

    try:
        del_sheet = sh.worksheet("Delivery_schedule")
        data_del = get_cached_data(del_sheet) # ใช้ cache แทนการยิง API
        df_del = pd.DataFrame(data_del) if data_del else pd.DataFrame()
    except gspread.exceptions.WorksheetNotFound:
        st.error("❌ ไม่พบแท็บชื่อ 'Delivery_schedule' ใน Google Sheets")
        df_del = pd.DataFrame()
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการดึงข้อมูล: {e}")
        df_del = pd.DataFrame()

    if not df_del.empty:
        df_del.columns = [str(c).strip() for c in df_del.columns]
        df_del["Original_Row_Index"] = df_del.index + 2

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

        with st.expander("✏️ เลื่อน/ปรับเปลี่ยนกำหนดการจัดส่ง (Reschedule)"):
            with st.form("reschedule_form"):
                st.markdown("เลือกรายการที่เลยกำหนด หรือต้องการเลื่อนวันส่ง **(สามารถเลือกได้หลายรายการพร้อมกัน)**")
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
                    
                selected_items = st.multiselect("เลือกรายการที่ต้องการแก้ไข (เลือกกี่รายการก็ได้)", reschedule_opts)
                new_date = st.date_input("กำหนดการใหม่ (New Expected Date)", value=date.today())
                
                if st.form_submit_button("💾 บันทึกกำหนดการใหม่"):
                    if not selected_items:
                        st.error("❌ กรุณาเลือกรายการอย่างน้อย 1 รายการก่อนทำการบันทึก")
                    elif date_col:
                        headers = del_sheet.row_values(1)
                        if date_col in headers:
                            col_idx = headers.index(date_col) + 1 
                            # 🚀 อัปเดตแบบ Batch (รวมเป็น Request เดียว ลดโควตา API)
                            cells_to_update = []
                            success_rows = []
                            for sel in selected_items:
                                target_row = reschedule_map[sel]["row_idx"]
                                cells_to_update.append(gspread.Cell(row=target_row, col=col_idx, value=str(new_date)))
                                success_rows.append(str(target_row))
                            
                            del_sheet.update_cells(cells_to_update)
                            get_cached_data.clear() # เคลียร์ cache
                            st.success(f"✅ อัปเดตวันจัดส่งใหม่เป็น {new_date} สำหรับบรรทัดที่ {', '.join(success_rows)} เรียบร้อยแล้ว!")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error(f"ไม่พบคอลัมน์ {date_col} ใน Google Sheets")

        st.divider()

        if date_col:
            df_del[date_col] = df_del[date_col].astype(str).str.strip()
            df_del["Is_TBD"] = df_del[date_col].str.lower().isin(["to be decided", "tbd", "", "nan", "nat", "none"])
            tbd_df = df_del[df_del["Is_TBD"]].copy()
            dated_df = df_del[~df_del["Is_TBD"]].copy()
            dated_df = dated_df.sort_values(by=date_col, ascending=True)
            unique_dates = dated_df[date_col].unique()

            for d_date in unique_dates:
                sub_df = dated_df[dated_df[date_col] == d_date]
                sub_display = sub_df.drop(columns=["Is_TBD", "Original_Row_Index"], errors="ignore")
                
                parsed_d = pd.to_datetime(d_date, dayfirst=True, errors='coerce')
                date_display = str(d_date)
                
                if pd.notna(parsed_d):
                    thai_days = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]
                    thai_day_str = thai_days[parsed_d.weekday()]
                    diff_days = (parsed_d.date() - date.today()).days
                    
                    if diff_days < 0: bg_color, border_color, text_color, urgency_icon, diff_text = "#fecaca", "#f87171", "#7f1d1d", "🚨", f"เลยกำหนดมาแล้ว {abs(diff_days)} วัน"
                    elif diff_days == 0: bg_color, border_color, text_color, urgency_icon, diff_text = "#fed7aa", "#fb923c", "#7c2d12", "🔥", "วันนี้"
                    elif diff_days <= 3: bg_color, border_color, text_color, urgency_icon, diff_text = "#fef08a", "#facc15", "#713f12", "⚠️", f"อีก {diff_days} วัน"
                    else: bg_color, border_color, text_color, urgency_icon, diff_text = "#bbf7d0", "#4ade80", "#14532d", "📅", f"อีก {diff_days} วัน"
                        
                    date_display = f"{urgency_icon} กำหนดส่ง: วัน{thai_day_str}ที่ {d_date} ({diff_text})"
                else:
                    bg_color, border_color, text_color = "#e2e8f0", "#a0aec0", "#2d3748"

                st.markdown(f'<div style="background-color: {bg_color}; color: {text_color}; padding: 15px 20px; border-radius: 8px; margin-top: 20px; margin-bottom: 12px; font-weight: bold; font-size: 22px; border-left: 8px solid {border_color}; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">{date_display}</div>', unsafe_allow_html=True)
                st.dataframe(format_df_quantity(sub_display), use_container_width=True)

            if not tbd_df.empty:
                tbd_display = tbd_df.drop(columns=["Is_TBD", "Original_Row_Index"], errors="ignore")
                st.markdown('<div style="background-color: #f1f5f9; color: #334155; padding: 15px 20px; border-radius: 8px; margin-top: 30px; margin-bottom: 12px; font-weight: bold; font-size: 22px; border-left: 8px solid #94a3b8; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">⏳ กำหนดส่ง: To be decided (รอยืนยันกำหนดการ)</div>', unsafe_allow_html=True)
                st.dataframe(format_df_quantity(tbd_display), use_container_width=True)
        else:
            st.info("ℹ️ ไม่พบคอลัมน์วันที่จัดส่งในแท็บนี้")
            st.dataframe(format_df_quantity(df_del.drop(columns=["Original_Row_Index"], errors="ignore")), use_container_width=True)
    else:
        st.info("ยังไม่มีข้อมูลในแท็บ 'Delivery_schedule'")

# ==========================================
# หน้าที่ 2: Stock Summary & Details
# ==========================================
elif menu == "📊 Stock Summary & Details":
    st.title("📊 Stock Summary & Details")
    st.markdown("ภาพรวมสต็อกสินค้าคงเหลือ (On-hand Stock) โดยดึงข้อมูล FG จากทั้งหน้า Goods Registration และ Packing Module")

    # 🚀 ใช้ get_cached_data แทน get_all_records() เพื่อไม่ให้กินโควตา
    data_main = get_cached_data(sheet)
    df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()

    records_data = get_cached_data(packing_records_sheet)
    df_packing_fg = pd.DataFrame()
    if records_data:
        df_pack_raw = pd.DataFrame(records_data)
        converted_rows = []
        for _, row in df_pack_raw.iterrows():
            converted_rows.append({
                "Cultivar Name": row.get("Cultivar Name", ""), "Supplier Name": "", "Supplier Lot. No.": row.get("Source WIP Lot", ""),
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

        st.subheader("📌 สรุปยอดรวมสต็อกในคลังทั้งหมด (Total Stock Summary)")
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1: st.metric(label="✨ ยอดรวมสต็อกทั้งหมด", value=f"{total_on_hand:,.2f}")
        with col_m2: st.metric(label="🧪 ยอดรวมประเภท WIP", value=f"{total_wip:,.2f}")
        with col_m3: st.metric(label="📦 ยอดรวมประเภท FG", value=f"{total_fg:,.2f}")

        st.divider()
        st.subheader("🧪 1. รายละเอียดสต็อกสินค้า WIP (Work in Progress)")
        if not wip_stock_df.empty:
            wip_big = wip_stock_df[wip_stock_df["Grade"].astype(str).str.strip().str.lower() == "big"]["Clean_Qty"].sum()
            wip_small = wip_stock_df[wip_stock_df["Grade"].astype(str).str.strip().str.lower() == "small"]["Clean_Qty"].sum()
            wip_mix = wip_stock_df[wip_stock_df["Grade"].astype(str).str.strip().str.lower() == "mix"]["Clean_Qty"].sum()
            wip_f = wip_stock_df[wip_stock_df["Grade"].astype(str).str.strip().str.lower() == "f"]["Clean_Qty"].sum()

            wc1, wc2, wc3, wc4, wc5 = st.columns(5)
            with wc1: st.metric(label="WIP: Big", value=f"{wip_big:,.2f}")
            with wc2: st.metric(label="WIP: Small", value=f"{wip_small:,.2f}")
            with wc3: st.metric(label="WIP: Mix", value=f"{wip_mix:,.2f}")
            with wc4: st.metric(label="WIP: F", value=f"{wip_f:,.2f}")
            with wc5: st.metric(label="WIP Total", value=f"{total_wip:,.2f}")
            st.dataframe(format_df_quantity(wip_stock_df), use_container_width=True)
        else: st.info("ไม่มีรายการสินค้า WIP ในคลังขณะนี้")

        st.divider()
        st.subheader("📦 2. รายละเอียดสต็อกสินค้า FG (Finished Goods)")
        if not fg_stock_df.empty:
            fg_big = fg_stock_df[fg_stock_df["Grade"].astype(str).str.strip().str.lower() == "big"]["Clean_Qty"].sum()
            fg_small = fg_stock_df[fg_stock_df["Grade"].astype(str).str.strip().str.lower() == "small"]["Clean_Qty"].sum()
            fg_mix = fg_stock_df[fg_stock_df["Grade"].astype(str).str.strip().str.lower() == "mix"]["Clean_Qty"].sum()
            fg_f = fg_stock_df[fg_stock_df["Grade"].astype(str).str.strip().str.lower() == "f"]["Clean_Qty"].sum()

            fc1, fc2, fc3, fc4, fc5 = st.columns(5)
            with fc1: st.metric(label="FG: Big", value=f"{fg_big:,.2f}")
            with fc2: st.metric(label="FG: Small", value=f"{fg_small:,.2f}")
            with fc3: st.metric(label="FG: Mix", value=f"{fg_mix:,.2f}")
            with fc4: st.metric(label="FG: F", value=f"{fg_f:,.2f}")
            with fc5: st.metric(label="FG Total", value=f"{total_fg:,.2f}")
            st.dataframe(format_df_quantity(fg_stock_df), use_container_width=True)
        else: st.info("ไม่มีรายการสินค้า FG ในคลังขณะนี้")
    else: st.warning("ยังไม่มีข้อมูลในระบบ")

# ==========================================
# หน้าที่ 3: Monthly Report
# ==========================================
elif menu == "📈 Monthly Report":
    st.title("📈 Monthly Report (รายงานสรุปผลการดำเนินงานรายเดือน)")
    st.markdown(
        "เลือกเดือนและปีที่ต้องการตรวจสอบรายงานสต็อกเคลื่อนไหวและประวัติการส่งออก (ค่าเริ่มต้นจะแสดงข้อมูลของเดือนปัจจุบัน)"
    )

    today = date.today()
    current_year = today.year
    current_month = today.month

    col_y, col_m = st.columns(2)
    with col_y:
        year_options = [2026, 2025, 2024, 2023]
        default_year_idx = year_options.index(current_year) if current_year in year_options else 0
        selected_year = st.selectbox("เลือกปี (Year)", year_options, index=default_year_idx, key="rep_year")
        
    with col_m:
        months_dict = {
            "มกราคม (01)": 1, "กุมภาพันธ์ (02)": 2, "มีนาคม (03)": 3, "เมษายน (04)": 4,
            "พฤษภาคม (05)": 5, "มิถุนายน (06)": 6, "กรกฎาคม (07)": 7, "สิงหาคม (08)": 8,
            "กันยายน (09)": 9, "ตุลาคม (10)": 10, "พฤศจิกายน (11)": 11, "ธันวาคม (12)": 12,
        }
        default_month_idx = current_month - 1
        selected_month_label = st.selectbox("เลือกเดือน (Month)", list(months_dict.keys()), index=default_month_idx, key="rep_month")
        selected_month_num = months_dict[selected_month_label]

    st.divider()

    rep_tab1, rep_tab2 = st.tabs([
        "📊 1. Monthly On-Hand Stock Report (WIP & FG)",
        "🚚 2. Shipped Out Report & Reasons",
    ])

    def parse_robust_date(date_val):
        if not date_val or not str(date_val).strip():
            return pd.NaT
        d_str = str(date_val).strip()
        try:
            dt = pd.to_datetime(d_str, errors='coerce')
            if pd.isna(dt):
                dt = pd.to_datetime(d_str, format='%d/%m/%Y', errors='coerce')
            return dt
        except:
            return pd.NaT

    # 🚀 ใช้ get_cached_data แทนการดึงสด เพื่อป้องกันติด API Limit
    shipped_data = get_cached_data(shipped_sheet)
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

        main_data = get_cached_data(sheet)
        df_main_rep = pd.DataFrame(main_data) if main_data else pd.DataFrame()
        pack_data = get_cached_data(packing_records_sheet)
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
                if not lot_no:
                    lot_no = str(r.get("Supplier Lot. No.", "")).strip()
                    
                cat = str(r.get("Category", "")).lower()
                curr_qty = clean_quantity(r.get("Quantity", 0))
                rec_date = str(r.get("Receive date", r.get("Date receive", r.get("Date Receive", ""))))
                
                original_qty = curr_qty + lot_addbacks.get(lot_no, 0.0)
                events.append({"date": rec_date, "type": "inbound", "cat": "wip" if "wip" in cat else "fg", "qty": original_qty})

        month_start = pd.to_datetime(f"{selected_year}-{selected_month_num:02d}-01")
        if selected_month_num == 12:
            month_end = pd.to_datetime(f"{selected_year+1}-01-01")
        else:
            month_end = pd.to_datetime(f"{selected_year}-{selected_month_num+1:02d}-01")

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

        if not has_wip_activity:
            st.warning(f"⚠️ **WIP:** ไม่มียอดยกมาและไม่มีกิจกรรมในเดือน {selected_month_label} {selected_year}")
        if not has_fg_activity:
            st.warning(f"⚠️ **FG:** ไม่มียอดยกมาและไม่มีกิจกรรมในเดือน {selected_month_label} {selected_year}")

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
            pwd = st.text_input("รหัสผ่าน:", type="password", key="gsheet_shortcut_pwd", placeholder="กรอกรหัสผ่านเพื่อแสดงลิงก์")
            if pwd == "tstd":
                st.markdown("[🟢 **คลิกที่นี่เพื่อเปิด Google Sheets**](https://docs.google.com/spreadsheets/d/1RLtRn7gojLLLxrsm9mSFRLy7eGmCYRn5bgsoCmuju7Y/edit?gid=0#gid=0)")
            elif pwd != "":
                st.error("รหัสผ่านไม่ถูกต้อง ❌")

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
                "Cultivar Name": cultivar_name, 
                "Supplier Name": supplier_name, 
                "Supplier Lot. No.": supplier_lot,
                "Export Name": export_name, 
                "THAI STICK Lot. No.": thai_stick_lot, 
                "Category": category, 
                "Grade": grade,
                "Receive date": receive_date_str, 
                "Quantity": float(quantity),
                "Remark Qty": remark_qty,
                "Inspection Req.": inspection_req, 
                "IR Req.": ir_req, 
                "Shipment date": shipment_date,
                "Type": prod_type, 
                "Status": status, 
                "Packing date": "",
            }
            
            row_to_insert = [row_dict.get(h, "") for h in latest_headers]
            
            sheet.append_row(values=row_to_insert, value_input_option='USER_ENTERED')
            get_cached_data.clear() # เคลียร์ cache ข้อมูลทันทีหลังเขียนเสร็จ
            st.success("บันทึกข้อมูลสำเร็จเรียบร้อยแล้ว! 🎉")
            time.sleep(1)
            st.rerun()

# ==========================================
# หน้าที่ 5: Edit Registration
# ==========================================
elif menu == "✏️ Edit Registration":
    st.title("✏️ Edit Registration Data")
    st.markdown("ค้นหาและแก้ไขข้อมูลล็อตสินค้าที่เคยบันทึกไว้")

    data = get_cached_data(sheet)
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
                edit_map[disp] = {
                    "row_idx": row_idx,
                    "lot_no": lot_no,
                    "df_index": df_idx
                }

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
                        e_receive_date_str = str(st.date_input("8. Receive date", value=default_rec_dt, key="edit_receive_date"))

                        old_qty = float(current_data.get("Clean_Qty", 0.0))
                        e_qty = st.number_input("9. Quantity (g)", min_value=0.0, step=0.1, value=old_qty)
                        e_remark_qty = st.text_input("Remark Qty", value=str(current_data.get("Remark Qty", "")), key="edit_remark_qty")

                        old_insp = str(current_data.get("Inspection Req.", "No"))
                        insp_con = st.container()
                        e_is_tbd_insp = st.checkbox("To be decided", value=(old_insp == "To be decided"), key=f"edit_tbd_insp")
                        insp_opts = ["Yes", "No"]
                        e_insp_val = insp_con.selectbox("10. Inspection Req.", insp_opts, index=insp_opts.index(old_insp) if old_insp in insp_opts else 0, disabled=e_is_tbd_insp, key=f"edit_insp")
                        e_insp = "To be decided" if e_is_tbd_insp else e_insp_val

                        old_ir = str(current_data.get("IR Req.", "No"))
                        ir_con = st.container()
                        e_is_tbd_ir = st.checkbox("To be decided", value=(old_ir == "To be decided"), key=f"edit_tbd_ir")
                        ir_opts = ["Yes", "No"]
                        e_ir_val = ir_con.selectbox("11. IR Req.", ir_opts, index=ir_opts.index(old_ir) if old_ir in ir_opts else 0, disabled=e_is_tbd_ir, key=f"edit_ir")
                        e_ir = "To be decided" if e_is_tbd_ir else e_ir_val

                        old_date = str(current_data.get("Shipment date", ""))
                        ship_con = st.container()
                        e_is_tbd_ship = st.checkbox("To be decided", value=(old_date == "To be decided"), key=f"edit_tbd_ship")
                        try: default_dt = pd.to_datetime(old_date, dayfirst=True).date() if old_date != "To be decided" and old_date.strip() else date.today()
                        except: default_dt = date.today()
                        e_ship_val = ship_con.date_input("12. Shipment date", value=default_dt, disabled=e_is_tbd_ship, key=f"edit_ship")
                        e_shipment_date = "To be decided" if e_is_tbd_ship else str(e_ship_val)

                        old_type = str(current_data.get("Type", ""))
                        type_con = st.container()
                        e_is_tbd_type = st.checkbox("To be decided", value=(old_type == "To be decided"), key=f"edit_tbd_type")
                        type_opts = ["Inspection & Packing Service", "Trading Goods (TS)", "Consignment", "In House Product"]
                        e_type_val = type_con.selectbox("13. Type", type_opts, index=type_opts.index(old_type) if old_type in type_opts else 0, disabled=e_is_tbd_type, key=f"edit_type")
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
                        get_cached_data.clear() # เคลียร์ cache
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
    
    data = get_cached_data(sheet)
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
                    
                    disp = f"[บรรทัด {row_idx}] Lot: {lot_no} | Size: {grade} | {cultivar} | Qty: {qty:,.2f}"

                    wip_options.append(disp)
                    wip_map[disp] = {
                        "lot_no": lot_no,
                        "cultivar": cultivar,
                        "grade": grade,
                        "qty": qty,
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
                                    
                                    # หักลบสต็อกโดยดึงข้อมูลจาก Memory (ไม่ต้องยิงไปขอข้อมูล Sheet ซ้ำ)
                                    for disp, w_qty in wip_withdrawals.items():
                                        if w_qty > 0:
                                            info = wip_map[disp]
                                            row_idx = info["row_idx"]
                                            lot_no = info["lot_no"]
                                            combined_source_lots.append(lot_no)
                                            
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

                                    breakdown_str = f"{g1_type}: {g1_qty} | {g2_type}: {g2_qty} | {g3_type}: {g3_qty} | {g4_type}: {g4_qty}"
                                    log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    
                                    packing_log_row = [
                                        log_time, source_lots_str, default_fg_name, packed_fg_name, breakdown_str, 
                                        float(total_packed_output), float(waste_qty), packing_date_str, note
                                    ]
                                    packing_sheet.append_row(packing_log_row, value_input_option='USER_ENTERED')

                                    fg_items = [(g1_type, g1_qty), (g2_type, g2_qty), (g3_type, g3_qty), (g4_type, g4_qty)]
                                    created_count = 0
                                    
                                    for idx_fg, (g_type, g_q) in enumerate(fg_items):
                                        if g_q > 0:
                                            created_count += 1
                                            custom_fg_lot = f"{base_lot}-FG{created_count}-{g_type}"
                                            
                                            rec_row = [
                                                log_time, source_lots_str, default_fg_name, packed_fg_name, custom_fg_lot, 
                                                g_type, float(g_q), float(waste_qty) if idx_fg == 0 else 0.0, packing_date_str, note
                                            ]
                                            packing_records_sheet.append_row(rec_row, value_input_option='USER_ENTERED')

                                    get_cached_data.clear() # เคลียร์ cache
                                    st.success(f"✅ บันทึกสำเร็จ! นำ WIP จำนวน {total_withdraw:,.2f} g ไปแพ็ก และสร้าง Lot FG ใหม่ {created_count} รายการ 🎉")
                                    time.sleep(1)
                                    st.rerun()

    st.divider()
    st.subheader("📑 ประวัติการผลิตและรหัส Lot FG แยกเกรด (Packing Records Sheet)")
    records_data = get_cached_data(packing_records_sheet)
    if records_data: st.dataframe(format_df_quantity(pd.DataFrame(records_data)), use_container_width=True)

# ==========================================
# หน้าที่ 7: Shipped Out Module
# ==========================================
elif menu == "🚚 Shipped Out Module":
    st.title("🚚 Shipped Out Module (Multi-Product Shipping & Stock Deduction)")
    st.markdown("บันทึกการส่งสินค้าออกทั้ง WIP และ FG (ดึงข้อมูลจากทั้งหน้า Trading Data และ Packing Records)")

    data_main = get_cached_data(sheet)
    df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()

    records_data = get_cached_data(packing_records_sheet)
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
                        latest_headers = [
                            "Cultivar Name", "Supplier Name", "Supplier Lot. No.", "Export Name", "THAI STICK Lot. No.", 
                            "Category", "Grade", "Receive date", "Quantity", "Remark Qty", "Inspection Req.", "IR Req.", 
                            "Shipment date", "Type", "Status", "Packing date"
                        ]
                        r_dict = item["row_data"]
                        current_qty = clean_quantity(r_dict.get("Quantity", 0))

                        remaining_qty = max(0.0, current_qty - qty)
                        new_status = "Delivered" if remaining_qty <= 0 else ("Pending Packing" if "WIP" in item["category"] else "Trading Goods Received to Stock")

                        r_dict["Quantity"] = str(remaining_qty)
                        r_dict["Status"] = new_status

                        row_values = [str(r_dict.get(h, "")) for h in latest_headers]
                        col_end_letter = chr(64 + len(latest_headers)).upper()
                        sheet.update(range_name=f"A{row_idx}:{col_end_letter}{row_idx}", values=[row_values])

                    elif source == "packing_records":
                        all_pack_records = get_cached_data(packing_records_sheet)
                        if row_idx and (row_idx - 2) < len(all_pack_records):
                            current_qty = clean_quantity(all_pack_records[row_idx - 2].get("Quantity", 0))
                            packing_records_sheet.update_cell(row_idx, 7, max(0.0, current_qty - qty))

                    shipped_sheet.append_row([
                        log_time, item["category"], item["lot_no"], item["product"], item["grade"], str(qty),
                        consignee, client_name, shipping_reason, shipment_date_str, note_ship
                    ], table_range="A1")
                    success_count += 1

                get_cached_data.clear() # เคลียร์ cache
                st.success(f"✅ บันทึกการส่งออกสำเร็จทั้งหมด {success_count} รายการและตัดสต็อกต้นทางเรียบร้อยแล้ว! 🎉")
                time.sleep(1)
                st.rerun()

    st.divider()
    st.subheader("📑 ประวัติการส่งสินค้าออก (Shipped Out Log)")
    shipped_data = get_cached_data(shipped_sheet)
    if shipped_data: st.dataframe(format_df_quantity(pd.DataFrame(shipped_data)), use_container_width=True)
