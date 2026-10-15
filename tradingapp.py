from datetime import date, datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import streamlit as st

# --- 1. ตั้งค่าหน้าเว็บ Streamlit ให้เป็น Wide Mode ---
st.set_page_config(
    page_title="Thai Stick Trading Portal", page_icon="🌱", layout="wide"
)

# --- 2. Custom CSS สำหรับตกแต่ง Template โทนสี Pastel อ่านง่ายสบายตา ---
st.markdown(
    """
    <style>
    .main {
        background-color: #f7f9fc;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .stMetric {
        background-color: #ffffff;
        padding: 18px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.02);
        border: 1px solid #e2e8f0;
    }
    .stMetric label {
        color: #4a5568 !important;
        font-weight: 600;
    }
    .stButton>button {
        width: 100%;
        background-color: #68d391;
        color: #1a202c;
        border-radius: 8px;
        font-weight: bold;
        border: none;
        padding: 0.5rem 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        background-color: #48bb78;
        color: white;
    }
    .stFormSubmitButton>button {
        width: 100%;
        background-color: #4fd1c5;
        color: #1a202c;
        border-radius: 8px;
        font-weight: bold;
        border: none;
    }
    .stFormSubmitButton>button:hover {
        background-color: #319795;
        color: white;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #edf2f7;
        border-radius: 8px 8px 0px 0px;
        padding: 10px 20px;
        color: #4a5568;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #4fd1c5 !important;
        color: #1a202c !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- 3. ฟังก์ชันเชื่อมต่อ Google Sheets และตรวจสอบ Header อัตโนมัติ ---
@st.cache_resource
def init_connection():
  scope = [
      "https://spreadsheets.google.com/feeds",
      "https://www.googleapis.com/auth/drive",
  ]
  creds_dict = dict(st.secrets["gcp_service_account"])
  creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
  client = gspread.authorize(creds)
  return client

client = init_connection()
SHEET_NAME = "Trading Data Log"

sh = client.open(SHEET_NAME)
sheet = sh.get_worksheet(0)

# กำหนดโครงสร้าง Column Header
EXPECTED_HEADERS = [
    "Cultivar Name",
    "Supplier Name",
    "Supplier Lot. No.",
    "Export Name",
    "THAI STICK Lot. No.",
    "Category",
    "Grade",
    "Receive date",
    "Quantity",
    "Remark Qty",
    "Inspection Req.",
    "IR Req.",
    "Shipment date",
    "Type",
    "Status",
    "Packing date",
]

current_headers = sheet.row_values(1)
if not current_headers or len(current_headers) < len(EXPECTED_HEADERS):
  sheet.update("A1:P1", [EXPECTED_HEADERS])

try:
  packing_records_sheet = sh.worksheet("Packing_Records")
except:
  packing_records_sheet = sh.add_worksheet(
      title="Packing_Records", rows="1000", cols="15"
  )
  packing_records_sheet.append_row([
      "Timestamp",
      "Source WIP Lot",
      "Cultivar Name",
      "Export Name",
      "Generated FG Lot No.",
      "Grade",
      "Quantity",
      "Waste Qty",
      "Packing Date",
      "Note",
  ])

try:
  packing_sheet = sh.worksheet("Packing_Log")
except:
  packing_sheet = sh.add_worksheet(title="Packing_Log", rows="1000", cols="10")
  packing_sheet.append_row([
      "Timestamp",
      "Source WIP Lot",
      "Cultivar",
      "Packed FG Name",
      "Grade & Qty Breakdown",
      "Total Packed Qty",
      "Waste Qty",
      "Packing Date",
      "Note",
  ])

try:
  shipped_sheet = sh.worksheet("Shipped_Out_Log")
except:
  shipped_sheet = sh.add_worksheet(title="Shipped_Out_Log", rows="1000", cols="12")
  shipped_sheet.append_row([
      "Timestamp",
      "Category",
      "Lot No.",
      "Cultivar / Product",
      "Grade",
      "Shipped Qty",
      "Consignee (ปลายทาง)",
      "Client (ลูกค้า)",
      "Shipping Reason",
      "Shipment Date",
      "Note",
  ])

# --- 4. Sidebar เมนูด้านซ้าย ---
st.sidebar.title("🌱 Thai Stick Trading")
st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "เลือกเมนูการใช้งาน",
    [
        "🚨 Delivery Issue",
        "📊 Stock Summary & Details",
        "📈 Monthly Report",
        "📝 Trading Goods Registration",
        "✏️ Edit Registration",
        "📦 Packing Module",
        "🚚 Shipped Out Module",
    ],
)
st.sidebar.markdown("---")
st.sidebar.info("ระบบจัดการสินค้า Trading และกระบวนการ Packing (Pastel Theme)")

def clean_quantity(val):
  try:
    return float(val)
  except:
    return 0.0

def format_df_quantity(df):
  """ฟังก์ชันช่วยแปลงคอลัมน์ Quantity ให้แสดงผลเป็นคอมมา (มีลูกน้ำ)"""
  if df is not None and not df.empty:
    df_copy = df.copy()
    for col in df_copy.columns:
      if "qty" in col.lower() or "quantity" in col.lower():
        df_copy[col] = df_copy[col].apply(lambda x: f"{clean_quantity(x):,.2f}" if pd.notnull(x) and x != "" else x)
    return df_copy
  return df

# ==========================================
# หน้าที่ 1: Delivery Issue
# ==========================================
if menu == "🚨 Delivery Issue":
  st.title("🚨 Delivery Issue - รายการที่ต้องจัดส่ง")
  st.markdown(
      "แสดงรายการสินค้าที่มีกำหนดส่ง (Shipment Date) โดยจัดกลุ่มตามวันจัดส่งอย่างชัดเจน"
  )

  data_main = sheet.get_all_records()
  df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()
  if not df_main.empty:
    df_main.columns = [str(c).strip() for c in df_main.columns]
    delivery_df = df_main.copy()
    delivery_df["Clean_Qty"] = delivery_df["Quantity"].apply(clean_quantity)

    if "Status" in delivery_df.columns:
      delivery_df = delivery_df[
          (~delivery_df["Status"].astype(str).str.strip().str.lower().eq("delivered"))
          & (delivery_df["Clean_Qty"] > 0)
      ]

    if not delivery_df.empty and "Shipment date" in delivery_df.columns:
      delivery_df["Shipment date str"] = delivery_df["Shipment date"].astype(str).str.strip()
      delivery_df["Is_TBD"] = delivery_df["Shipment date str"].str.lower() == "to be decided"

      tbd_df = delivery_df[delivery_df["Is_TBD"]].copy()
      dated_df = delivery_df[~delivery_df["Is_TBD"]].copy()

      dated_df = dated_df.sort_values(by="Shipment date str", ascending=True)
      unique_dates = dated_df["Shipment date str"].unique()

      for d_date in unique_dates:
        sub_df = dated_df[dated_df["Shipment date str"] == d_date]
        sub_display = sub_df.drop(columns=["Is_TBD", "Clean_Qty", "Shipment date str"], errors="ignore")
        
        st.markdown(
            f"""
            <div style="background-color: #4fd1c5; color: #1a202c; padding: 10px 15px; 
            border-radius: 8px; margin-top: 15px; margin-bottom: 10px; font-weight: bold; font-size: 16px; 
            border-left: 5px solid #319795;">
                📅 กำหนดส่งวันที่: {d_date}
            </div>
            """,
            unsafe_allow_html=True
        )
        st.dataframe(format_df_quantity(sub_display), use_container_width=True)

      if not tbd_df.empty:
        tbd_display = tbd_df.drop(columns=["Is_TBD", "Clean_Qty", "Shipment date str"], errors="ignore")
        
        st.markdown(
            f"""
            <div style="background-color: #f6ad55; color: white; padding: 10px 15px; 
            border-radius: 8px; margin-top: 25px; margin-bottom: 10px; font-weight: bold; font-size: 16px;
            border-left: 5px solid #dd6b20;">
                ⏳ กำหนดส่ง: To be decided (รอยืนยันกำหนดการ)
            </div>
            """,
            unsafe_allow_html=True
        )
        st.dataframe(format_df_quantity(tbd_display), use_container_width=True)

    else:
      st.info("ไม่พบรายการที่ต้องจัดส่งในขณะนี้")
  else:
    st.info("ยังไม่มีข้อมูลในระบบ Google Sheets")

# ==========================================
# หน้าที่ 2: Stock Summary & Details
# ==========================================
elif menu == "📊 Stock Summary & Details":
  st.title("📊 Stock Summary & Details")
  st.markdown(
      "ภาพรวมสต็อกสินค้าคงเหลือ (On-hand Stock) โดยดึงข้อมูล FG จากทั้งหน้า Goods Registration และ Packing Module"
  )

  data_main = sheet.get_all_records()
  df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()
  if not df_main.empty:
    df_main.columns = [str(c).strip() for c in df_main.columns]

  records_data = packing_records_sheet.get_all_records()
  df_packing_fg = pd.DataFrame()
  if records_data:
    df_pack_raw = pd.DataFrame(records_data)
    df_pack_raw.columns = [str(c).strip() for c in df_pack_raw.columns]
    converted_rows = []
    for _, row in df_pack_raw.iterrows():
      converted_rows.append({
          "Cultivar Name": row.get("Cultivar Name", ""),
          "Supplier Name": "",
          "Supplier Lot. No.": row.get("Source WIP Lot", ""),
          "Export Name": row.get("Export Name", ""),
          "THAI STICK Lot. No.": row.get("Generated FG Lot No.", ""),
          "Category": "FG (Finished Goods)",
          "Grade": row.get("Grade", ""),
          "Receive date": row.get("Packing Date", ""),
          "Quantity": clean_quantity(row.get("Quantity", 0)),
          "Remark Qty": "",
          "Inspection Req.": "No",
          "IR Req.": "No",
          "Shipment date": "To be decided",
          "Type": "In House Product",
          "Status": "Packing done",
          "Packing date": row.get("Packing Date", ""),
      })
    df_packing_fg = pd.DataFrame(converted_rows)

  if not df_main.empty and not df_packing_fg.empty:
    df = pd.concat([df_main, df_packing_fg], ignore_index=True)
  elif not df_main.empty:
    df = df_main
  elif not df_packing_fg.empty:
    df = df_packing_fg
  else:
    df = pd.DataFrame()

  if not df.empty:
    if "Quantity" in df.columns:
      df["Clean_Qty"] = df["Quantity"].apply(clean_quantity)
    else:
      df["Clean_Qty"] = 0.0

    filtered_df = df.copy()

    if "Status" in filtered_df.columns:
      filtered_df = filtered_df[
          ~filtered_df["Status"].astype(str).str.strip().str.lower().eq("delivered")
      ]
    filtered_df = filtered_df[filtered_df["Clean_Qty"] > 0]

    if "Category" in filtered_df.columns:
      wip_stock_df = filtered_df[
          filtered_df["Category"].astype(str).str.contains("WIP", case=False, na=False)
      ]
      fg_stock_df = filtered_df[
          filtered_df["Category"].astype(str).str.contains("FG", case=False, na=False)
      ]
    else:
      wip_stock_df = pd.DataFrame()
      fg_stock_df = pd.DataFrame()

    total_wip = wip_stock_df["Clean_Qty"].sum()
    total_fg = fg_stock_df["Clean_Qty"].sum()
    total_on_hand = total_wip + total_fg

    st.subheader("📌 สรุปยอดรวมสต็อกในคลังทั้งหมด (Total Stock Summary)")
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
      st.metric(label="✨ ยอดรวมสต็อกทั้งหมด", value=f"{total_on_hand:,.2f}")
    with col_m2:
      st.metric(label="🧪 ยอดรวมประเภท WIP", value=f"{total_wip:,.2f}")
    with col_m3:
      st.metric(label="📦 ยอดรวมประเภท FG", value=f"{total_fg:,.2f}")

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
      
      st.markdown("")
      st.dataframe(format_df_quantity(wip_stock_df), use_container_width=True)
    else:
      st.info("ไม่มีรายการสินค้า WIP ในคลังขณะนี้")

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
      
      st.markdown("")
      st.dataframe(format_df_quantity(fg_stock_df), use_container_width=True)
    else:
      st.info("ไม่มีรายการสินค้า FG ในคลังขณะนี้")

    with st.expander("🔍 ดูข้อมูลดิบทั้งหมด (Raw Data ทั้งหมดในระบบ)"):
      st.dataframe(format_df_quantity(df), use_container_width=True)
  else:
    st.warning("ยังไม่มีข้อมูลในระบบ Google Sheets สำหรับประมวลผลสต็อก")

# ==========================================
# หน้าที่ 3: Monthly Report
# ==========================================
elif menu == "📈 Monthly Report":
  st.title("📈 Monthly Report (รายงานสรุปผลการดำเนินงานรายเดือน)")
  st.markdown(
      "เลือกเดือนและปีที่ต้องการตรวจสอบรายงานสต็อกเคลื่อนไหวและประวัติการส่งออก (ค่าเริ่มต้นจะแสดงข้อมูลของเดือนปัจจุบัน)"
  )

  # --- ดึงวันที่ปัจจุบัน เพื่อตั้งค่า Default สำหรับ Year และ Month ---
  today = date.today()
  current_year = today.year
  current_month = today.month

  col_y, col_m = st.columns(2)
  with col_y:
    year_options = [2026, 2025, 2024, 2023]
    # ตรวจสอบว่าปีปัจจุบันอยู่ในลิสต์ไหม ถ้าอยู่ให้เลือกเป็น Default ถ้าไม่มีให้ใช้ช่องแรก
    default_year_idx = year_options.index(current_year) if current_year in year_options else 0
    selected_year = st.selectbox("เลือกปี (Year)", year_options, index=default_year_idx, key="rep_year")
    
  with col_m:
    months_dict = {
        "มกราคม (01)": 1, "กุมภาพันธ์ (02)": 2, "มีนาคม (03)": 3, "เมษายน (04)": 4,
        "พฤษภาคม (05)": 5, "มิถุนายน (06)": 6, "กรกฎาคม (07)": 7, "สิงหาคม (08)": 8,
        "กันยายน (09)": 9, "ตุลาคม (10)": 10, "พฤศจิกายน (11)": 11, "ธันวาคม (12)": 12,
    }
    # index ของเดือนจะเริ่มที่ 0 (มกราคม) ดังนั้นเอาเดือนปัจจุบัน - 1
    default_month_idx = current_month - 1
    selected_month_label = st.selectbox("เลือกเดือน (Month)", list(months_dict.keys()), index=default_month_idx, key="rep_month")
    selected_month_num = months_dict[selected_month_label]

  st.divider()

  rep_tab1, rep_tab2 = st.tabs([
      "📊 1. Monthly On-Hand Stock Report (WIP & FG)",
      "🚚 2. Shipped Out Report & Reasons",
  ])

  # --- ฟังก์ชันตัวช่วยสำหรับอ่านวันที่ให้แม่นยำ (แก้ปัญหา YYYY/MM/DD สลับเป็นเดือนมกราคม) ---
  def parse_robust_date(date_val):
    if not date_val or not str(date_val).strip():
        return pd.NaT
    d_str = str(date_val).strip()
    try:
        # ลองอ่านแบบมาตรฐานปกติก่อน (จะแก้ปัญหา 2026/10/01 ให้อ่าน 10 เป็นเดือนได้อย่างถูกต้อง)
        dt = pd.to_datetime(d_str, errors='coerce')
        if pd.isna(dt):
            # ถ้ายังอ่านไม่ได้ ค่อยบังคับให้ใช้อ่านแบบ วัน/เดือน/ปี (DD/MM/YYYY)
            dt = pd.to_datetime(d_str, format='%d/%m/%Y', errors='coerce')
        return dt
    except:
        return pd.NaT

  shipped_data = shipped_sheet.get_all_records()
  df_shipped = pd.DataFrame(shipped_data) if shipped_data else pd.DataFrame()
  if not df_shipped.empty:
    df_shipped.columns = [str(c).strip() for c in df_shipped.columns]
    # ใช้ฟังก์ชันอ่านวันที่ตัวใหม่
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

    main_data = sheet.get_all_records()
    df_main_rep = pd.DataFrame(main_data) if main_data else pd.DataFrame()
    pack_data = packing_records_sheet.get_all_records()
    df_pack_rep = pd.DataFrame(pack_data) if pack_data else pd.DataFrame()

    has_wip_activity = False
    has_fg_activity = False
    wip_receive_total = 0.0
    fg_receive_total = 0.0

    if not df_main_rep.empty:
      df_main_rep.columns = [str(c).strip() for c in df_main_rep.columns]
      for _, r in df_main_rep.iterrows():
        cat = str(r.get("Category", ""))
        # เผื่อไว้กรณีคุณยังไม่ได้เปลี่ยนชื่อหัวคอลัมน์ใน Sheet
        rec_date = r.get("Receive date", r.get("Date receive", r.get("Date Receive", "")))
        q = clean_quantity(r.get("Quantity", 0))
        
        # ใช้ฟังก์ชันใหม่ในการแปลงวันที่
        p_dt = parse_robust_date(rec_date)
        if pd.notna(p_dt):
            if p_dt.year == selected_year and p_dt.month == selected_month_num:
                if "WIP" in cat.upper():
                    wip_receive_total += q
                    has_wip_activity = True
                elif "FG" in cat.upper():
                    fg_receive_total += q
                    has_fg_activity = True

    if not df_pack_rep.empty:
      df_pack_rep.columns = [str(c).strip() for c in df_pack_rep.columns]
      for _, r in df_pack_rep.iterrows():
        p_date = r.get("Packing Date", "")
        q = clean_quantity(r.get("Quantity", 0))
        
        # ใช้ฟังก์ชันใหม่ในการแปลงวันที่
        p_dt = parse_robust_date(p_date)
        if pd.notna(p_dt):
            if p_dt.year == selected_year and p_dt.month == selected_month_num:
                fg_receive_total += q
                has_fg_activity = True

    wip_shipped_monthly = 0.0
    fg_shipped_monthly = 0.0
    if not monthly_shipped.empty:
      for _, r in monthly_shipped.iterrows():
        cat = str(r.get("Category", ""))
        sq = clean_quantity(r.get("Clean_Shipped_Qty", 0))
        if "WIP" in cat:
          wip_shipped_monthly += sq
          has_wip_activity = True
        else:
          fg_shipped_monthly += sq
          has_fg_activity = True

    if not has_wip_activity:
      st.warning(f"⚠️ **WIP:** ไม่มีกิจกรรมในเดือน {selected_month_label} {selected_year}")
    if not has_fg_activity:
      st.warning(f"⚠️ **FG:** ไม่มีกิจกรรมในเดือน {selected_month_label} {selected_year}")

    summary_report_data = [
        {
            "Category": "WIP (Work in Progress)",
            "Status Activity": "🟢 Active" if has_wip_activity else "⚪ No Activity (นิ่ง)",
            "Beginning Stock": max(0.0, wip_receive_total - wip_shipped_monthly),
            "Received / Inbound": wip_receive_total,
            "Shipped Out / Outbound": wip_shipped_monthly,
            "Ending On-Hand Stock": max(0.0, wip_receive_total - wip_shipped_monthly),
        },
        {
            "Category": "FG (Finished Goods)",
            "Status Activity": "🟢 Active" if has_fg_activity else "⚪ No Activity (นิ่ง)",
            "Beginning Stock": max(0.0, fg_receive_total - fg_shipped_monthly),
            "Received / Inbound": fg_receive_total,
            "Shipped Out / Outbound": fg_shipped_monthly,
            "Ending On-Hand Stock": max(0.0, fg_receive_total - fg_shipped_monthly),
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
      latest_headers = sheet.row_values(1)
      row_dict = {
          "Cultivar Name": cultivar_name, "Supplier Name": supplier_name, "Supplier Lot. No.": supplier_lot,
          "Export Name": export_name, "THAI STICK Lot. No.": thai_stick_lot, "Category": category, "Grade": grade,
          "Receive date": receive_date_str, "Quantity": str(quantity), "Remark Qty": remark_qty,
          "Inspection Req.": inspection_req, "IR Req.": ir_req, "Shipment date": shipment_date,
          "Type": prod_type, "Status": status, "Packing date": "",
      }
      sheet.append_row([str(row_dict.get(h, "")) for h in latest_headers])
      st.success("บันทึกข้อมูลสำเร็จเรียบร้อยแล้ว! 🎉")
      st.rerun()

  st.divider()
  st.subheader("📋 ประวัติข้อมูลทั้งหมดในระบบ")
  data = sheet.get_all_records()
  if data: st.dataframe(format_df_quantity(pd.DataFrame(data)), use_container_width=True)
  else: st.info("ยังไม่มีข้อมูลในระบบ Google Sheets")

# ==========================================
# หน้าที่ 5: Edit Registration
# ==========================================
elif menu == "✏️ Edit Registration":
  st.title("✏️ Edit Registration Data")
  st.markdown("ค้นหาและแก้ไขข้อมูลล็อตสินค้าที่เคยบันทึกไว้")

  data = sheet.get_all_records()
  if data:
    df_all = pd.DataFrame(data)
    df_all.columns = [str(c).strip() for c in df_all.columns]
    lot_column = "THAI STICK Lot. No." if "THAI STICK Lot. No." in df_all.columns else df_all.columns[4]

    df_all["Clean_Qty"] = df_all["Quantity"].apply(clean_quantity) if "Quantity" in df_all.columns else 0.0
    # เก็บเลขบรรทัดเดิมไว้ เพื่อให้อัปเดตได้ถูกบรรทัดเป๊ะๆ แม้ Lot ซ้ำกัน
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
        
        # สร้างข้อความสำหรับแสดงใน Dropdown
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
            # ดักรับชื่อคอลัมน์วันที่ 3 แบบตามที่เคยแก้ไว้
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
            latest_headers = sheet.row_values(1)
            updated_dict = {
                "Cultivar Name": e_cultivar, "Supplier Name": e_supplier, "Supplier Lot. No.": e_supp_lot,
                "Export Name": e_export, "THAI STICK Lot. No.": e_thai_lot, "Category": e_category, "Grade": e_grade,
                "Receive date": e_receive_date_str, 
                "Date receive": e_receive_date_str, 
                "Date Receive": e_receive_date_str, 
                "Quantity": str(e_qty), "Remark Qty": e_remark_qty,
                "Inspection Req.": e_insp, "IR Req.": e_ir, "Shipment date": e_shipment_date,
                "Type": e_type, "Status": e_status, "Packing date": "",
            }
            updated_row_data = [str(updated_dict.get(h, "")) for h in latest_headers]
            col_end_letter = chr(64 + len(latest_headers)).upper()
            
            # บันทึกลงบรรทัดเดิมเป๊ะๆ ที่ล็อคเป้าไว้
            sheet.update(f"A{matched_row_idx}:{col_end_letter}{matched_row_idx}", [updated_row_data])
            st.success(f"อัปเดตข้อมูลของ Lot: {selected_lot} (Size: {e_grade}) เรียบร้อยแล้ว! 🎉")
  else:
    st.info("ยังไม่มีข้อมูลในระบบให้แก้ไข")

# ==========================================
# หน้าที่ 6: Packing Module
# ==========================================
elif menu == "📦 Packing Module":
  st.title("📦 Packing Module (WIP Conversion & Dedicated FG Tracking)")
  
  data = sheet.get_all_records()
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
        # เก็บ Index เดิมบวก 2 เพื่อระบุบรรทัดใน Google Sheet ได้แม่นยำ (แก้ปัญหา Lot ซ้ำกันแต่คนละไซส์)
        available_wip["Original_Row_Index"] = available_wip.index + 2 

        wip_options = []
        wip_map = {}

        for _, r in available_wip.iterrows():
          lot_no = str(r.get("THAI STICK Lot. No.", "")).strip()
          if not lot_no: lot_no = str(r.get("Supplier Lot. No.", "")).strip()
          
          cultivar = str(r.get("Cultivar Name", ""))
          grade = str(r.get("Grade", ""))
          qty = float(r["Clean_Qty"])
          row_idx = r["Original_Row_Index"]
          
          # สร้างข้อความสำหรับแสดงใน Dropdown ให้เห็นทั้ง Lot และ Size (Grade)
          disp = f"Lot: {lot_no} | Size: {grade} | {cultivar} | Qty: {qty:,.2f}"
          wip_options.append(disp)
          wip_map[disp] = {
              "lot_no": lot_no,
              "cultivar": cultivar,
              "grade": grade,
              "qty": qty,
              "row_idx": row_idx
          }

        selected_wip_disp = st.selectbox("🔍 เลือก WIP Lot และ Size ที่ต้องการนำมาเบิกแพ็ก", wip_options)

        if selected_wip_disp:
          selected_info = wip_map[selected_wip_disp]
          selected_wip_lot = selected_info["lot_no"]
          max_wip_qty = selected_info["qty"]
          cultivar_val = selected_info["cultivar"]
          wip_grade = selected_info["grade"]
          matched_idx = selected_info["row_idx"] # ได้บรรทัดที่ถูกต้องของ Google Sheet มาเลย

          st.info(f"📌 **WIP Lot:** {selected_wip_lot} | **Size/Grade:** {wip_grade} | **Cultivar:** {cultivar_val} | **สต็อกคงเหลือ:** {max_wip_qty:,.2f}")

          with st.form("packing_form"):
            col_p1, col_p2 = st.columns(2)
            with col_p1:
              withdraw_qty = st.number_input("จำนวนเบิก (Withdraw Qty)", min_value=0.0, max_value=max_wip_qty, step=0.1, value=max_wip_qty)
              packing_date_str = str(st.date_input("Packing date", value=date.today()))
            with col_p2:
              packed_fg_name = st.text_input("ชื่อผลิตภัณฑ์หลังแพ็ก", value=cultivar_val)

            st.divider()
            grade_options = ["Big", "Small", "Mix", "F", "Grade A", "Premium"]
            c1, c2 = st.columns(2)
            with c1: g1_type = st.selectbox("เกรดช่องที่ 1", grade_options, index=0)
            with c2: g1_qty = st.number_input("จำนวนช่องที่ 1", min_value=0.0, step=0.1, key="g1")

            c3, c4 = st.columns(2)
            with c3: g2_type = st.selectbox("เกรดช่องที่ 2", grade_options, index=1)
            with c4: g2_qty = st.number_input("จำนวนช่องที่ 2", min_value=0.0, step=0.1, key="g2")

            c5, c6 = st.columns(2)
            with c5: g3_type = st.selectbox("เกรดช่องที่ 3", grade_options, index=2)
            with c6: g3_qty = st.number_input("จำนวนช่องที่ 3", min_value=0.0, step=0.1, key="g3")

            c7, c8 = st.columns(2)
            with c7: g4_type = st.selectbox("เกรดช่องที่ 4", grade_options, index=3)
            with c8: g4_qty = st.number_input("จำนวนช่องที่ 4", min_value=0.0, step=0.1, key="g4")

            st.divider()
            waste_qty = st.number_input("จำนวนของเสีย (Waste Qty)", min_value=0.0, step=0.1)
            note = st.text_area("หมายเหตุเพิ่มเติม (Note)")

            if st.form_submit_button("🚀 ยืนยันการแพ็กและบันทึกลงตาราง Packing Records"):
              total_packed_output = g1_qty + g2_qty + g3_qty + g4_qty
              if withdraw_qty <= 0: st.error("❌ กรุณาระบุจำนวนที่เบิกมาแพ็กมากกว่า 0")
              elif (total_packed_output + waste_qty) > withdraw_qty: st.warning("⚠️ คำเตือน: ผลผลิตรวม + ของเสีย มากกว่าจำนวนที่เบิกมา")

              remaining_wip_in_stock = max_wip_qty - withdraw_qty
              new_status = "Packing done" if remaining_wip_in_stock <= 0 else "Pending Packing"
              remaining_wip_in_stock = max(0.0, remaining_wip_in_stock)

              # เปลี่ยนจากการ Search หา Lot ไปใช้บรรทัด matched_idx ที่เราล็อคเป้าไว้เลย แม่นยำกว่า
              latest_headers = sheet.row_values(1)
              row_values = sheet.row_values(matched_idx)
              while len(row_values) < len(latest_headers): row_values.append("")

              for idx, h in enumerate(latest_headers):
                if h == "Quantity": row_values[idx] = str(remaining_wip_in_stock)
                elif h == "Status": row_values[idx] = new_status

              col_end_letter = chr(64 + len(latest_headers)).upper()
              sheet.update(f"A{matched_idx}:{col_end_letter}{matched_idx}", [row_values])

              breakdown_str = f"{g1_type}: {g1_qty} | {g2_type}: {g2_qty} | {g3_type}: {g3_qty} | {g4_type}: {g4_qty}"
              log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
              packing_sheet.append_row([log_time, selected_wip_lot, cultivar_val, packed_fg_name, breakdown_str, str(total_packed_output), str(waste_qty), packing_date_str, note])

              fg_items = [(g1_type, g1_qty), (g2_type, g2_qty), (g3_type, g3_qty), (g4_type, g4_qty)]
              created_count = 0
              for idx_fg, (g_type, g_q) in enumerate(fg_items):
                if g_q > 0:
                  created_count += 1
                  custom_fg_lot = f"{selected_wip_lot}-FG{created_count}-{g_type}"
                  packing_records_sheet.append_row([log_time, selected_wip_lot, cultivar_val, packed_fg_name, custom_fg_lot, g_type, str(g_q), str(waste_qty) if idx_fg == 0 else "0", packing_date_str, note])

              st.success(f"✅ บันทึกการแพ็กสำเร็จ! หัก WIP Size {wip_grade} เหลือ {remaining_wip_in_stock:,.2f} และสร้าง Lot FG ใหม่จำนวน {created_count} รายการ 🎉")

    st.divider()
    st.subheader("📑 ประวัติการผลิตและรหัส Lot FG แยกเกรด (Packing Records Sheet)")
    records_data = packing_records_sheet.get_all_records()
    if records_data: st.dataframe(format_df_quantity(pd.DataFrame(records_data)), use_container_width=True)

# ==========================================
# หน้าที่ 7: Shipped Out Module
# ==========================================
elif menu == "🚚 Shipped Out Module":
  st.title("🚚 Shipped Out Module (Multi-Product Shipping & Stock Deduction)")
  st.markdown("บันทึกการส่งสินค้าออกทั้ง WIP และ FG (ดึงข้อมูลจากทั้งหน้า Trading Data และ Packing Records)")

  data_main = sheet.get_all_records()
  df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()

  records_data = packing_records_sheet.get_all_records()
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
          "max_qty": max_q, "source": source, "row_idx": r["Original_Row_Index"]
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
            latest_headers = sheet.row_values(1)
            row_values = sheet.row_values(row_idx)
            while len(row_values) < len(latest_headers): row_values.append("")

            current_qty = 0.0
            for idx, h in enumerate(latest_headers):
              if h == "Quantity": current_qty = clean_quantity(row_values[idx])

            remaining_qty = max(0.0, current_qty - qty)
            new_status = "Delivered" if remaining_qty <= 0 else ("Pending Packing" if "WIP" in item["category"] else "Trading Goods Received to Stock")

            for idx, h in enumerate(latest_headers):
              if h == "Quantity": row_values[idx] = str(remaining_qty)
              elif h == "Status": row_values[idx] = new_status

            col_end_letter = chr(64 + len(latest_headers)).upper()
            sheet.update(f"A{row_idx}:{col_end_letter}{row_idx}", [row_values])

          elif source == "packing_records":
            all_pack_records = packing_records_sheet.get_all_records()
            if row_idx and (row_idx - 2) < len(all_pack_records):
              current_qty = clean_quantity(all_pack_records[row_idx - 2].get("Quantity", 0))
              packing_records_sheet.update_cell(row_idx, 7, max(0.0, current_qty - qty))

          shipped_sheet.append_row([
              log_time, item["category"], item["lot_no"], item["product"], item["grade"], str(qty),
              consignee, client_name, shipping_reason, shipment_date_str, note_ship
          ])
          success_count += 1
        st.success(f"✅ บันทึกการส่งออกสำเร็จทั้งหมด {success_count} รายการและตัดสต็อกต้นทางเรียบร้อยแล้ว! 🎉")

  st.divider()
  st.subheader("📑 ประวัติการส่งสินค้าออก (Shipped Out Log)")
  shipped_data = shipped_sheet.get_all_records()
  if shipped_data: st.dataframe(format_df_quantity(pd.DataFrame(shipped_data)), use_container_width=True)
