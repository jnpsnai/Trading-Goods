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
# หน้าที่ 3-7: รบกวนกด "Continue" หรือตอบ "ทำต่อ" ในแชทครับ ผมจะให้โค้ดส่วนที่เหลือ (เนื่องจากพื้นที่ตอบมีจำกัด)
# การแก้ไขทั้งหมดหลักๆ ใช้ get_cached_data(_sheet) แทนการดึงสดครับ
# ==========================================
