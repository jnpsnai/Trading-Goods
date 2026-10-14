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


# ==========================================
# หน้าที่ 1: Delivery Issue
# ==========================================
if menu == "🚨 Delivery Issue":
  st.title("🚨 Delivery Issue - รายการที่ต้องจัดส่ง")
  st.markdown(
      "แสดงรายการสินค้าที่มีกำหนดส่ง (Shipment Date) เรียงจากใกล้ไปไกล"
      " โดยรายการที่เป็น 'To be decided' จะถูกจัดไว้ท้ายสุด"
  )

  data_main = sheet.get_all_records()
  df_main = pd.DataFrame(data_main) if data_main else pd.DataFrame()
  if not df_main.empty:
    df_main.columns = [str(c).strip() for c in df_main.columns]
    delivery_df = df_main.copy()
    delivery_df["Clean_Qty"] = delivery_df["Quantity"].apply(clean_quantity)

    if "Status" in delivery_df.columns:
      delivery_df = delivery_df[
          (
              ~delivery_df["Status"]
              .astype(str)
              .str.strip()
              .str.lower()
              .eq("delivered")
          )
          & (delivery_df["Clean_Qty"] > 0)
      ]

    if not delivery_df.empty and "Shipment date" in delivery_df.columns:
      delivery_df["Shipment date str"] = (
          delivery_df["Shipment date"].astype(str).str.strip()
      )
      delivery_df["Is_TBD"] = (
          delivery_df["Shipment date str"].str.lower() == "to be decided"
      )
      delivery_df = delivery_df.sort_values(
          by=["Is_TBD", "Shipment date str"], ascending=[True, True]
      )
      display_delivery_df = delivery_df.drop(
          columns=["Is_TBD", "Clean_Qty"], errors="ignore"
      )
      st.dataframe(display_delivery_df, use_container_width=True)
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
      "ภาพรวมสต็อกสินค้าคงเหลือ (On-hand Stock) แบ่งตามประเภท WIP และ FG"
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
          "Status": "Trading Goods Received to Stock",
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
    filtered_df = filtered_df[
        filtered_df["THAI STICK Lot. No."].astype(str).str.strip() != ""
    ]

    if "Status" in filtered_df.columns and "Type" in filtered_df.columns:
      filtered_df = filtered_df[
          (
              filtered_df["Status"]
              .astype(str)
              .str.strip()
              .str.lower()
              .ne("delivered")
          )
          & (
              filtered_df["Type"]
              .astype(str)
              .str.strip()
              .str.lower()
              .ne("inspection & packing service")
          )
          & (filtered_df["Clean_Qty"] > 0)
      ]

    if "Category" in filtered_df.columns:
      wip_stock_df = filtered_df[
          filtered_df["Category"]
          .astype(str)
          .str.contains("WIP", case=False, na=False)
      ]
      fg_stock_df = filtered_df[
          filtered_df["Category"]
          .astype(str)
          .str.contains("FG", case=False, na=False)
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
      wip_big = wip_stock_df[
          wip_stock_df["Grade"].astype(str).str.strip().str.lower() == "big"
      ]["Clean_Qty"].sum()
      wip_small = wip_stock_df[
          wip_stock_df["Grade"].astype(str).str.strip().str.lower()
          == "small"
      ]["Clean_Qty"].sum()
      wip_mix = wip_stock_df[
          wip_stock_df["Grade"].astype(str).str.strip().str.lower() == "mix"
      ]["Clean_Qty"].sum()
      wip_f = wip_stock_df[
          wip_stock_df["Grade"].astype(str).str.strip().str.lower() == "f"
      ]["Clean_Qty"].sum()

      wc1, wc2, wc3, wc4, wc5 = st.columns(5)
      with wc1:
        st.metric(label="WIP: Big", value=f"{wip_big:,.2f}")
      with wc2:
        st.metric(label="WIP: Small", value=f"{wip_small:,.2f}")
      with wc3:
        st.metric(label="WIP: Mix", value=f"{wip_mix:,.2f}")
      with wc4:
        st.metric(label="WIP: F", value=f"{wip_f:,.2f}")
      with wc5:
        st.metric(label="WIP Total", value=f"{total_wip:,.2f}")

      st.markdown("")
      st.dataframe(wip_stock_df, use_container_width=True)
    else:
      st.info("ไม่มีรายการสินค้า WIP ในคลังขณะนี้")

    st.divider()
    st.subheader("📦 2. รายละเอียดสต็อกสินค้า FG (Finished Goods)")
    if not fg_stock_df.empty:
      fg_big = fg_stock_df[
          fg_stock_df["Grade"].astype(str).str.strip().str.lower() == "big"
      ]["Clean_Qty"].sum()
      fg_small = fg_stock_df[
          fg_stock_df["Grade"].astype(str).str.strip().str.lower()
          == "small"
      ]["Clean_Qty"].sum()
      fg_mix = fg_stock_df[
          fg_stock_df["Grade"].astype(str).str.strip().str.lower() == "mix"
      ]["Clean_Qty"].sum()
      fg_f = fg_stock_df[
          fg_stock_df["Grade"].astype(str).str.strip().str.lower() == "f"
      ]["Clean_Qty"].sum()

      fc1, fc2, fc3, fc4, fc5 = st.columns(5)
      with fc1:
        st.metric(label="FG: Big", value=f"{fg_big:,.2f}")
      with fc2:
        st.metric(label="FG: Small", value=f"{fg_small:,.2f}")
      with fc3:
        st.metric(label="FG: Mix", value=f"{fg_mix:,.2f}")
      with fc4:
        st.metric(label="FG: F", value=f"{fg_f:,.2f}")
      with fc5:
        st.metric(label="FG Total", value=f"{total_fg:,.2f}")

      st.markdown("")
      st.dataframe(fg_stock_df, use_container_width=True)
    else:
      st.info("ไม่มีรายการสินค้า FG ในคลังขณะนี้")

    with st.expander("🔍 ดูข้อมูลดิบทั้งหมด (Raw Data ทั้งหมดในระบบ)"):
      st.dataframe(df, use_container_width=True)
  else:
    st.warning("ยังไม่มีข้อมูลในระบบ Google Sheets สำหรับประมวลผลสต็อก")


# ==========================================
# หน้าที่ 3: Monthly Report
# ==========================================
elif menu == "📈 Monthly Report":
  st.title("📈 Monthly Report (รายงานสรุปผลการดำเนินงานรายเดือน)")
  st.markdown(
      "เลือกเดือนและปีที่ต้องการตรวจสอบรายงานสต็อกเคลื่อนไหวและประวัติการส่งออก"
  )

  col_y, col_m = st.columns(2)
  with col_y:
    selected_year = st.selectbox(
        "เลือกปี (Year)", [2026, 2025, 2024, 2023], index=0, key="rep_year"
    )
  with col_m:
    months_dict = {
        "มกราคม (01)": 1,
        "กุมภาพันธ์ (02)": 2,
        "มีนาคม (03)": 3,
        "เมษายน (04)": 4,
        "พฤษภาคม (05)": 5,
        "มิถุนายน (06)": 6,
        "กรกฎาคม (07)": 7,
        "สิงหาคม (08)": 8,
        "กันยายน (09)": 9,
        "ตุลาคม (10)": 10,
        "พฤศจิกายน (11)": 11,
        "ธันวาคม (12)": 12,
    }
    selected_month_label = st.selectbox(
        "เลือกเดือน (Month)", list(months_dict.keys()), key="rep_month"
    )
    selected_month_num = months_dict[selected_month_label]

  st.divider()

  rep_tab1, rep_tab2 = st.tabs([
      "📊 1. Monthly On-Hand Stock Report (WIP & FG)",
      "🚚 2. Shipped Out Report & Reasons",
  ])

  shipped_data = shipped_sheet.get_all_records()
  df_shipped = pd.DataFrame(shipped_data) if shipped_data else pd.DataFrame()
  if not df_shipped.empty:
    df_shipped.columns = [str(c).strip() for c in df_shipped.columns]
    df_shipped["Parsed_Date"] = pd.to_datetime(
        df_shipped["Shipment Date"], errors="coerce"
    )
    df_shipped["Clean_Shipped_Qty"] = df_shipped["Shipped Qty"].apply(
        clean_quantity
    )
    monthly_shipped = df_shipped[
        (df_shipped["Parsed_Date"].dt.year == selected_year)
        & (df_shipped["Parsed_Date"].dt.month == selected_month_num)
    ].copy()
  else:
    monthly_shipped = pd.DataFrame()

  with rep_tab1:
    st.subheader(
        f"📅 รายงานความเคลื่อนไหวสต็อกประจำเดือน {selected_month_label}"
        f" {selected_year}"
    )

    main_data = sheet.get_all_records()
    df_main_rep = pd.DataFrame(main_data) if main_data else pd.DataFrame()
    pack_data = packing_records_sheet.get_all_records()
    df_pack_rep = pd.DataFrame(pack_data) if pack_data else pd.DataFrame()

    has_wip_activity = False
    wip_receive_total = 0.0
    if not df_main_rep.empty:
      df_main_rep.columns = [str(c).strip() for c in df_main_rep.columns]
      for _, r in df_main_rep.iterrows():
        cat = str(r.get("Category", ""))
        rec_date = str(r.get("Receive date", ""))
        q = clean_quantity(r.get("Quantity", 0))
        if "WIP" in cat:
          try:
            p_dt = datetime.strptime(rec_date, "%Y-%m-%d")
            if p_dt.year == selected_year and p_dt.month == selected_month_num:
              wip_receive_total += q
              has_wip_activity = True
          except:
            pass

    has_fg_activity = False
    fg_receive_total = 0.0
    if not df_pack_rep.empty:
      df_pack_rep.columns = [str(c).strip() for c in df_pack_rep.columns]
      for _, r in df_pack_rep.iterrows():
        p_date = str(r.get("Packing Date", ""))
        q = clean_quantity(r.get("Quantity", 0))
        try:
          p_dt = datetime.strptime(p_date, "%Y-%m-%d")
          if p_dt.year == selected_year and p_dt.month == selected_month_num:
            fg_receive_total += q
            has_fg_activity = True
        except:
          pass

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
      st.warning(
          f"⚠️ **WIP (Work in Progress):** ไม่มีกิจกรรมหรือ Transaction"
          f" การเคลื่อนไหวใดๆ ในเดือน {selected_month_label} {selected_year}"
      )
    if not has_fg_activity:
      st.warning(
          f"⚠️ **FG (Finished Goods):** ไม่มีกิจกรรมหรือ Transaction"
          f" การเคลื่อนไหวใดๆ ในเดือน {selected_month_label} {selected_year}"
      )

    summary_report_data = [
        {
            "Category": "WIP (Work in Progress)",
            "Status Activity": (
                "🟢 Active" if has_wip_activity else "⚪ No Activity (นิ่ง)"
            ),
            "Beginning Stock": max(0.0, wip_receive_total - wip_shipped_monthly),
            "Received / Inbound": wip_receive_total,
            "Shipped Out / Outbound": wip_shipped_monthly,
            "Ending On-Hand Stock": max(
                0.0, wip_receive_total - wip_shipped_monthly
            ),
        },
        {
            "Category": "FG (Finished Goods)",
            "Status Activity": (
                "🟢 Active" if has_fg_activity else "⚪ No Activity (นิ่ง)"
            ),
            "Beginning Stock": max(0.0, fg_receive_total - fg_shipped_monthly),
            "Received / Inbound": fg_receive_total,
            "Shipped Out / Outbound": fg_shipped_monthly,
            "Ending On-Hand Stock": max(
                0.0, fg_receive_total - fg_shipped_monthly
            ),
        },
    ]

    df_summary_rep = pd.DataFrame(summary_report_data)
    st.markdown("")
    st.dataframe(df_summary_rep, use_container_width=True)

  with rep_tab2:
    st.subheader(
        f"🚚 Shipped Out Report ประจำเดือน {selected_month_label}"
        f" {selected_year}"
    )
    if not monthly_shipped.empty:
      reason_summary = (
          monthly_shipped.groupby("Shipping Reason")["Clean_Shipped_Qty"]
          .sum()
          .reset_index()
      )
      reason_summary.columns = [
          "Shipping Reason (เหตุผลการส่ง)",
          "Total Shipped Qty (จำนวนรวม)",
      ]
      st.markdown("##### 📌 สรุปยอดจัดส่งแยกตามเหตุผล")
      st.dataframe(reason_summary, use_container_width=True)

      st.markdown("##### 📋 รายละเอียดรายการส่งออกทั้งหมด")
      st.dataframe(monthly_shipped, use_container_width=True)
    else:
      st.info(
          f"ℹ️ ไม่พบประวัติการส่งสินค้าออกในเดือน {selected_month_label}"
          f" {selected_year} (ไม่มีกิจกรรมการส่งออก)"
      )


# ==========================================
# หน้าที่ 4: Trading Goods Registration 
# ==========================================
elif menu == "📝 Trading Goods Registration":
  st.title("📝 Trading Goods Registration")
  st.markdown("กรอกข้อมูลผลิตภัณฑ์และประเภทสินค้า (WIP หรือ FG) ลงในระบบ")

  with st.form("trading_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
      cultivar_name = st.text_input("1. Cultivar Name")
      supplier_name = st.text_input("2. Supplier Name")
      supplier_lot = st.text_input("3. Supplier Lot. No.")
      export_name = st.text_input("4. Export Name")
      thai_stick_lot = st.text_input("5. THAI STICK Lot. No.")
      category = st.selectbox(
          "6. Category", ["FG (Finished Goods)", "WIP (Work in Progress)"]
      )
      grade = st.selectbox(
          "7. Grade", ["Big", "Small", "Mix", "F", "Grade A", "Premium"]
      )
    with col2:
      receive_date_str = str(
          st.date_input(
              "8. Receive date (Column H)",
              value=date.today(),
              key="reg_rec_dt",
          )
      )
      quantity = st.number_input("9. Quantity", min_value=0.0, step=0.1)
      remark_qty = st.text_input(
          "Remark Qty (เช่น 5x1000)", placeholder="ระบุรายละเอียดจำนวนถุง/แพ็ก"
      )

      # 10. Inspection Req. (ใช้ Container เพื่อวางช่อง Checkbox ไว้ด้านล่างเสมอ)
      insp_con = st.container()
      is_tbd_insp = st.checkbox("To be decided", key="reg_tbd_insp")
      inspection_req_val = insp_con.selectbox(
          "10. Inspection Req.", ["Yes", "No"], disabled=is_tbd_insp, key="reg_insp"
      )
      inspection_req = "To be decided" if is_tbd_insp else inspection_req_val

      # 11. IR Req. (ใช้ Container)
      ir_con = st.container()
      is_tbd_ir = st.checkbox("To be decided", key="reg_tbd_ir")
      ir_req_val = ir_con.selectbox(
          "11. IR Req.", ["Yes", "No"], disabled=is_tbd_ir, key="reg_ir"
      )
      ir_req = "To be decided" if is_tbd_ir else ir_req_val

      # 12. Shipment date (ใช้ Container)
      ship_con = st.container()
      is_tbd_ship = st.checkbox("To be decided", key="reg_tbd_ship")
      shipment_date_val = ship_con.date_input(
          "12. Shipment date", value=date.today(), disabled=is_tbd_ship, key="reg_ship"
      )
      shipment_date = "To be decided" if is_tbd_ship else str(shipment_date_val)

      # 13. Type (ใช้ Container)
      type_con = st.container()
      is_tbd_type = st.checkbox("To be decided", key="reg_tbd_type")
      type_opts = [
          "Inspection & Packing Service",
          "Trading Goods (TS)",
          "Consignment",
          "In House Product"
      ]
      prod_type_val = type_con.selectbox(
          "13. Type", type_opts, disabled=is_tbd_type, key="reg_type"
      )
      prod_type = "To be decided" if is_tbd_type else prod_type_val

      status = st.selectbox(
          "14. Status",
          [
              "Trading Goods Received to Stock",
              "Inspection Completed",
              "Pending Packing",
              "Packing done",
              "Delivered",
              "Waiting for Return",
          ],
      )
      packing_date_str = str(
          st.date_input(
              "15. Packing date (Column O)",
              value=date.today(),
              key="reg_pack_dt",
          )
      )

    if st.form_submit_button("💾 บันทึกข้อมูลลง Google Sheets"):
      latest_headers = sheet.row_values(1)
      row_dict = {
          "Cultivar Name": cultivar_name,
          "Supplier Name": supplier_name,
          "Supplier Lot. No.": supplier_lot,
          "Export Name": export_name,
          "THAI STICK Lot. No.": thai_stick_lot,
          "Category": category,
          "Grade": grade,
          "Receive date": receive_date_str,
          "Quantity (g)": str(quantity),
          "Remark Qty": remark_qty,
          "Inspection Req.": inspection_req,
          "IR Req.": ir_req,
          "Shipment date": shipment_date,
          "Type": prod_type,
          "Status": status,
          "Packing date": packing_date_str,
      }
      sheet.append_row([str(row_dict.get(h, "")) for h in latest_headers])
      st.success("บันทึกข้อมูลสำเร็จเรียบร้อยแล้ว! 🎉")

  st.divider()
  st.subheader("📋 ประวัติข้อมูลทั้งหมดในระบบ")
  data = sheet.get_all_records()
  if data:
    st.dataframe(pd.DataFrame(data), use_container_width=True)
  else:
    st.info("ยังไม่มีข้อมูลในระบบ Google Sheets")


# ==========================================
# หน้าที่ 5: Edit Registration
# ==========================================
elif menu == "✏️ Edit Registration":
  st.title("✏️ Edit Registration Data")
  st.markdown(
      "ค้นหาและแก้ไขข้อมูลล็อตสินค้าที่เคยบันทึกไว้ (ไม่อนุญาตให้แก้ไขล็อตที่"
      " Stock หมด หรือมี Status เป็น Delivered แล้ว)"
  )

  data = sheet.get_all_records()
  if data:
    df_all = pd.DataFrame(data)
    df_all.columns = [str(c).strip() for c in df_all.columns]
    lot_column = (
        "THAI STICK Lot. No."
        if "THAI STICK Lot. No." in df_all.columns
        else df_all.columns[4]
    )

    df_all["Clean_Qty"] = (
        df_all["Quantity"].apply(clean_quantity)
        if "Quantity" in df_all.columns
        else 0.0
    )

    if "Status" in df_all.columns:
      editable_df = df_all[
          (df_all["Clean_Qty"] > 0)
          & (
              ~df_all["Status"]
              .astype(str)
              .str.strip()
              .str.lower()
              .eq("delivered")
          )
      ].copy()
    else:
      editable_df = df_all[df_all["Clean_Qty"] > 0].copy()

    if editable_df.empty:
      st.info(
          "ℹ️ ไม่มีล็อตสินค้าที่สามารถแก้ไขได้ในขณะนี้ (สินค้าทั้งหมดอาจจะมี Stock"
          " เป็น 0 หรือถูกส่งมอบ Delivered ไปหมดแล้ว)"
      )
    else:
      lot_list = editable_df[lot_column].astype(str).tolist()
      selected_lot = st.selectbox(
          "🔍 เลือก THAI STICK Lot. No. ที่ต้องการแก้ไข", lot_list
      )

      if selected_lot:
        matched_row_idx = (
            int(
                df_all[df_all[lot_column].astype(str) == selected_lot].index[0]
            )
            + 2
        )
        current_data = df_all[
            df_all[lot_column].astype(str) == selected_lot
        ].iloc[0]

        st.info(f"กำลังแก้ไขข้อมูลของล็อต: **{selected_lot}**")

        with st.form("edit_form"):
          col1, col2 = st.columns(2)

          with col1:
            e_cultivar = st.text_input(
                "1. Cultivar Name",
                value=str(current_data.get("Cultivar Name", "")),
            )
            e_supplier = st.text_input(
                "2. Supplier Name",
                value=str(current_data.get("Supplier Name", "")),
            )
            e_supp_lot = st.text_input(
                "3. Supplier Lot. No.",
                value=str(current_data.get("Supplier Lot. No.", "")),
            )
            e_export = st.text_input(
                "4. Export Name", value=str(current_data.get("Export Name", ""))
            )
            e_thai_lot = st.text_input(
                "5. THAI STICK Lot. No.",
                value=str(current_data.get("THAI STICK Lot. No.", "")),
            )

            cat_opt = ["FG (Finished Goods)", "WIP (Work in Progress)"]
            old_cat = str(current_data.get("Category", "FG (Finished Goods)"))
            e_category = st.selectbox(
                "6. Category",
                cat_opt,
                index=cat_opt.index(old_cat) if old_cat in cat_opt else 0,
            )

            grades_opt = ["Big", "Small", "Mix", "F", "Grade A", "Premium"]
            old_grade = str(current_data.get("Grade", "Big"))
            e_grade = st.selectbox(
                "7. Grade",
                grades_opt,
                index=grades_opt.index(old_grade)
                if old_grade in grades_opt
                else 0,
            )

          with col2:
            old_rec_date = str(current_data.get("Receive date", ""))
            try:
              default_rec_dt = datetime.strptime(
                  old_rec_date, "%Y-%m-%d"
              ).date()
            except:
              default_rec_dt = date.today()
            e_rec_date_input = st.date_input(
                "8. Receive date (Column H)",
                value=default_rec_dt,
                key="edit_receive_date",
            )
            e_receive_date_str = str(e_rec_date_input)

            try:
              old_qty = float(current_data.get("Quantity", 0.0))
            except:
              old_qty = 0.0
            e_qty = st.number_input(
                "9. Quantity", min_value=0.0, step=0.1, value=old_qty
            )
            e_remark_qty = st.text_input(
                "Remark Qty",
                value=str(current_data.get("Remark Qty", "")),
                key="edit_remark_qty",
            )

            # 10. Inspection Req. (Container & Checkbox)
            old_insp = str(current_data.get("Inspection Req.", "No"))
            insp_con = st.container()
            e_is_tbd_insp = st.checkbox("To be decided", value=(old_insp == "To be decided"), key=f"edit_tbd_insp_{selected_lot}")
            insp_opts = ["Yes", "No"]
            insp_def_idx = insp_opts.index(old_insp) if old_insp in insp_opts else 0
            e_insp_val = insp_con.selectbox(
                "10. Inspection Req.", insp_opts, index=insp_def_idx, disabled=e_is_tbd_insp, key=f"edit_insp_{selected_lot}"
            )
            e_insp = "To be decided" if e_is_tbd_insp else e_insp_val

            # 11. IR Req. (Container & Checkbox)
            old_ir = str(current_data.get("IR Req.", "No"))
            ir_con = st.container()
            e_is_tbd_ir = st.checkbox("To be decided", value=(old_ir == "To be decided"), key=f"edit_tbd_ir_{selected_lot}")
            ir_opts = ["Yes", "No"]
            ir_def_idx = ir_opts.index(old_ir) if old_ir in ir_opts else 0
            e_ir_val = ir_con.selectbox(
                "11. IR Req.", ir_opts, index=ir_def_idx, disabled=e_is_tbd_ir, key=f"edit_ir_{selected_lot}"
            )
            e_ir = "To be decided" if e_is_tbd_ir else e_ir_val

            # 12. Shipment date (Container & Checkbox)
            old_date = str(current_data.get("Shipment date", ""))
            ship_con = st.container()
            e_is_tbd_ship = st.checkbox("To be decided", value=(old_date == "To be decided"), key=f"edit_tbd_ship_{selected_lot}")
            try:
                default_dt = datetime.strptime(old_date, "%Y-%m-%d").date() if old_date != "To be decided" else date.today()
            except:
                default_dt = date.today()
            e_ship_val = ship_con.date_input(
                "12. Shipment date", value=default_dt, disabled=e_is_tbd_ship, key=f"edit_ship_{selected_lot}"
            )
            e_shipment_date = "To be decided" if e_is_tbd_ship else str(e_ship_val)

            # 13. Type (Container & Checkbox)
            old_type = str(current_data.get("Type", ""))
            type_con = st.container()
            e_is_tbd_type = st.checkbox("To be decided", value=(old_type == "To be decided"), key=f"edit_tbd_type_{selected_lot}")
            type_opts = ["Inspection & Packing Service", "Trading Goods (TS)", "Consignment", "In House Product"]
            type_def_idx = type_opts.index(old_type) if old_type in type_opts else 0
            e_type_val = type_con.selectbox(
                "13. Type", type_opts, index=type_def_idx, disabled=e_is_tbd_type, key=f"edit_type_{selected_lot}"
            )
            e_type = "To be decided" if e_is_tbd_type else e_type_val

            status_opt = [
                "Trading Goods Received to Stock",
                "Inspection Completed",
                "Pending Packing",
                "Packing done",
                "Delivered",
                "Waiting for Return",
            ]
            old_status = str(current_data.get("Status", ""))
            e_status = st.selectbox(
                "14. Status",
                status_opt,
                index=status_opt.index(old_status)
                if old_status in status_opt
                else 0,
            )

            old_pack_date = str(current_data.get("Packing date", ""))
            try:
              default_pack_dt = datetime.strptime(
                  old_pack_date, "%Y-%m-%d"
              ).date()
            except:
              default_pack_dt = date.today()
            e_pack_date_input = st.date_input(
                "15. Packing date (Column O)",
                value=default_pack_dt,
                key="edit_pack_date",
            )
            e_packing_date_str = str(e_pack_date_input)

          st.markdown("")
          update_submitted = st.form_submit_button(
              "🔄 อัปเดตข้อมูล (Update Google Sheets)"
          )

          if update_submitted:
            latest_headers = sheet.row_values(1)
            updated_dict = {
                "Cultivar Name": e_cultivar,
                "Supplier Name": e_supplier,
                "Supplier Lot. No.": e_supp_lot,
                "Export Name": e_export,
                "THAI STICK Lot. No.": e_thai_lot,
                "Category": e_category,
                "Grade": e_grade,
                "Receive date": e_receive_date_str,
                "Quantity": str(e_qty),
                "Remark Qty": e_remark_qty,
                "Inspection Req.": e_insp,
                "IR Req.": e_ir,
                "Shipment date": e_shipment_date,
                "Type": e_type,
                "Status": e_status,
                "Packing date": e_packing_date_str,
            }
            updated_row_data = [
                str(updated_dict.get(h, "")) for h in latest_headers
            ]

            col_end_letter = chr(64 + len(latest_headers)).upper()
            sheet.update(
                f"A{matched_row_idx}:{col_end_letter}{matched_row_idx}",
                [updated_row_data],
            )
            st.success(f"อัปเดตข้อมูลของล็อต {selected_lot} เรียบร้อยแล้ว! 🎉")
  else:
    st.info("ยังไม่มีข้อมูลในระบบให้แก้ไข")


# ==========================================
# หน้าที่ 6: Packing Module
# ==========================================
elif menu == "📦 Packing Module":
  st.title("📦 Packing Module (WIP Conversion & Dedicated FG Tracking)")
  st.markdown(
      "เลือก WIP Lot ที่ต้องการเบิกมาแพ็ก ระบุชื่อผลิตภัณฑ์ (Export Name)"
      " วันที่แพ็ก และบันทึกผลผลิต FG แยกเกรดลงใน Sheet เฉพาะ"
      " (`Packing_Records`) พร้อมตัดสต็อก WIP ต้นทางอัตโนมัติ"
  )

  data = sheet.get_all_records()
  if data:
    df = pd.DataFrame(data)
    df.columns = [str(c).strip() for c in df.columns]

    if "Category" in df.columns:
      wip_df = df[
          df["Category"]
          .astype(str)
          .str.contains("WIP", case=False, na=False)
      ].copy()
    else:
      wip_df = pd.DataFrame()

    if wip_df.empty:
      st.warning(
          "⚠️ ยังไม่พบข้อมูลสินค้าประเภท 'WIP (Work in Progress)' ในระบบ"
          " กรุณาไปเพิ่มข้อมูลในหน้า 'Trading Goods Registration' ก่อน"
      )
    else:
      wip_df["Clean_Qty"] = wip_df["Quantity"].apply(clean_quantity)
      available_wip = wip_df[wip_df["Clean_Qty"] > 0]

      if available_wip.empty:
        st.warning(
            "⚠️ มีข้อมูล WIP ในระบบ แต่ทุก Lot มีจำนวนคงเหลือ (Quantity) เป็น 0"
            " แล้ว"
        )
      else:
        wip_lot_list = available_wip["THAI STICK Lot. No."].astype(str).tolist()

        selected_wip_lot = st.selectbox(
            "🔍 เลือก WIP Lot ที่ต้องการนำมาเบิกแพ็ก", wip_lot_list
        )

        selected_row = available_wip[
            available_wip["THAI STICK Lot. No."].astype(str)
            == selected_wip_lot
        ].iloc[0]
        max_wip_qty = float(selected_row["Clean_Qty"])
        cultivar_val = selected_row.get("Cultivar Name", "")

        st.info(
            f"📌 **WIP Lot:** {selected_wip_lot} | **Cultivar:**"
            f" {cultivar_val} | **สต็อกคงเหลือปัจจุบัน:**"
            f" {max_wip_qty:,.2f}"
        )

        with st.form("packing_form"):
          st.subheader("1. ข้อมูลการเบิกวัตถุดิบและชื่อผลิตภัณฑ์")
          col_p1, col_p2 = st.columns(2)
          with col_p1:
            withdraw_qty = st.number_input(
                "จำนวนที่ต้องการเบิกมาแพ็กในรอบนี้ (Withdraw Qty)",
                min_value=0.0,
                max_value=max_wip_qty,
                step=0.1,
                value=max_wip_qty,
            )
            pack_date_input = st.date_input(
                "Packing date (Column O)", value=date.today()
            )
            packing_date_str = str(pack_date_input)

          with col_p2:
            packed_fg_name = st.text_input(
                "ชื่อผลิตภัณฑ์หลังแพ็ก (Packed Product / Export Name)",
                value=cultivar_val,
            )

          st.divider()
          st.subheader(
              "2. บันทึกผลผลิตหลังแพ็ก (ระบุเกรดและจำนวนแยก 4 ช่อง)"
          )
          st.markdown(
              "ระบบจะสร้างรหัส Lot เฉพาะ (เช่น `WIP-FG2-Big`)"
              " บันทึกลงตารางแยกต่างหากให้ทันที"
          )

          grade_options = ["Big", "Small", "Mix", "F", "Grade A", "Premium"]

          c1, c2 = st.columns(2)
          with c1:
            g1_type = st.selectbox("เกรดช่องที่ 1", grade_options, index=0)
          with c2:
            g1_qty = st.number_input(
                "จำนวนช่องที่ 1", min_value=0.0, step=0.1, key="g1"
            )

          c3, c4 = st.columns(2)
          with c3:
            g2_type = st.selectbox("เกรดช่องที่ 2", grade_options, index=1)
          with c4:
            g2_qty = st.number_input(
                "จำนวนช่องที่ 2", min_value=0.0, step=0.1, key="g2"
            )

          c5, c6 = st.columns(2)
          with c5:
            g3_type = st.selectbox("เกรดช่องที่ 3", grade_options, index=2)
          with c6:
            g3_qty = st.number_input(
                "จำนวนช่องที่ 3", min_value=0.0, step=0.1, key="g3"
            )

          c7, c8 = st.columns(2)
          with c7:
            g4_type = st.selectbox("เกรดช่องที่ 4", grade_options, index=3)
          with c8:
            g4_qty = st.number_input(
                "จำนวนช่องที่ 4", min_value=0.0, step=0.1, key="g4"
            )

          st.divider()
          st.subheader("3. ของเสีย (Waste) และหมายเหตุ")
          waste_qty = st.number_input(
              "จำนวนของเสีย/สูญเสีย (Waste Qty)", min_value=0.0, step=0.1
          )
          note = st.text_area("หมายเหตุเพิ่มเติม (Note)")

          pack_submitted = st.form_submit_button(
              "🚀 ยืนยันการแพ็กและบันทึกลงตาราง Packing Records"
          )

          if pack_submitted:
            total_packed_output = g1_qty + g2_qty + g3_qty + g4_qty
            if withdraw_qty <= 0:
              st.error("❌ กรุณาระบุจำนวนที่ต้องการเบิกมาแพ็กมากกว่า 0")
            elif (total_packed_output + waste_qty) > withdraw_qty:
              st.warning(
                  "⚠️ คำเตือน: ผลผลิตรวม + ของเสีย มากกว่าจำนวนที่เบิกมาในรอบนี้"
              )

            remaining_wip_in_stock = max_wip_qty - withdraw_qty
            if remaining_wip_in_stock <= 0:
              new_status = "Packing done"
              remaining_wip_in_stock = 0.0
            else:
              new_status = "Pending Packing"

            cell_wip = sheet.find(str(selected_wip_lot))
            if cell_wip:
              matched_idx = cell_wip.row
            else:
              matched_idx = (
                  int(
                      df[
                          df["THAI STICK Lot. No."].astype(str)
                          == selected_wip_lot
                      ].index[0]
                  )
                  + 2
              )

            latest_headers = sheet.row_values(1)
            row_values = sheet.row_values(matched_idx)

            while len(row_values) < len(latest_headers):
              row_values.append("")

            for idx, h in enumerate(latest_headers):
              if h == "Quantity":
                row_values[idx] = str(remaining_wip_in_stock)
              elif h == "Status":
                row_values[idx] = new_status

            col_end_letter = chr(64 + len(latest_headers)).upper()
            sheet.update(
                f"A{matched_idx}:{col_end_letter}{matched_idx}", [row_values]
            )

            breakdown_str = (
                f"{g1_type}: {g1_qty} | {g2_type}: {g2_qty} | {g3_type}:"
                f" {g3_qty} | {g4_type}: {g4_qty}"
            )
            log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            packing_sheet.append_row([
                log_time,
                selected_wip_lot,
                cultivar_val,
                packed_fg_name,
                breakdown_str,
                str(total_packed_output),
                str(waste_qty),
                packing_date_str,
                note,
            ])

            fg_items = [
                (g1_type, g1_qty),
                (g2_type, g2_qty),
                (g3_type, g3_qty),
                (g4_type, g4_qty),
            ]

            created_count = 0
            for idx_fg, (g_type, g_q) in enumerate(fg_items):
              if g_q > 0:
                created_count += 1
                custom_fg_lot = (
                    f"{selected_wip_lot}-FG{created_count}-{g_type}"
                )
                packing_records_sheet.append_row([
                    log_time,
                    selected_wip_lot,
                    cultivar_val,
                    packed_fg_name,
                    custom_fg_lot,
                    g_type,
                    str(g_q),
                    str(waste_qty) if idx_fg == 0 else "0",
                    packing_date_str,
                    note,
                ])

            st.success(
                f"✅ บันทึกการแพ็กสำเร็จ! หัก WIP เหลือ"
                f" {remaining_wip_in_stock:,.2f} และสร้างรหัส Lot FG ใหม่ลงใน"
                f" Sheet `Packing_Records` จำนวน {created_count}"
                " รายการเรียบร้อยแล้ว! 🎉"
            )

    st.divider()
    st.subheader(
        "📑 ประวัติการผลิตและรหัส Lot FG แยกเกรด (Packing Records Sheet)"
    )
    records_data = packing_records_sheet.get_all_records()
    if records_data:
      st.dataframe(pd.DataFrame(records_data), use_container_width=True)
    else:
      st.info("ยังไม่มีข้อมูลใน Sheet Packing_Records")

    with st.expander("🔍 ดูประวัติ Packing Log ทั้งหมด"):
      log_data = packing_sheet.get_all_records()
      if log_data:
        st.dataframe(pd.DataFrame(log_data), use_container_width=True)
      else:
        st.info("ยังไม่มีประวัติใน Packing_Log")
  else:
    st.warning("ยังไม่มีข้อมูลในระบบ Google Sheets")


# ==========================================
# หน้าที่ 7: Shipped Out Module
# ==========================================
elif menu == "🚚 Shipped Out Module":
  st.title("🚚 Shipped Out Module (Multi-Product Shipping & Stock Deduction)")
  st.markdown(
      "บันทึกการส่งสินค้าออก (Delivery / Shipment) ทั้ง WIP และ FG"
      " สามารถเลือกส่งออกหลายผลิตภัณฑ์พร้อมกันในครั้งเดียว"
      " ระบบจะหักสต็อกต้นทางอัตโนมัติ"
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
    for idx_row, row in df_pack_raw.iterrows():
      converted_rows.append({
          "Original_Row_Index": idx_row + 2,
          "Cultivar Name": row.get("Cultivar Name", ""),
          "Export Name": row.get("Export Name", ""),
          "THAI STICK Lot. No.": row.get("Generated FG Lot No.", ""),
          "Grade": row.get("Grade", ""),
          "Quantity": clean_quantity(row.get("Quantity", 0)),
          "Packing Date": row.get("Packing Date", ""),
      })
    df_packing_fg = pd.DataFrame(converted_rows)

  combined_options = []
  inventory_map = {}

  if not df_main.empty and "Category" in df_main.columns:
    wip_available = df_main[
        df_main["Category"].astype(str).str.contains("WIP", case=False, na=False)
    ].copy()
    wip_available["Clean_Qty"] = wip_available["Quantity"].apply(
        clean_quantity
    )
    wip_available = wip_available[
        (wip_available["Clean_Qty"] > 0)
        & (
            ~wip_available["Status"]
            .astype(str)
            .str.strip()
            .str.lower()
            .eq("delivered")
        )
    ]
    for _, r in wip_available.iterrows():
      lot_no = str(r.get("THAI STICK Lot. No.", ""))
      prod_name = str(r.get("Cultivar Name", ""))
      grade_val = str(r.get("Grade", ""))
      max_q = float(r["Clean_Qty"])
      if lot_no:
        display_str = (
            f"[WIP] {lot_no} | ชื่อ: {prod_name} | เกรด: {grade_val} |"
            f" คงเหลือ: {max_q:,.1f}"
        )
        combined_options.append(display_str)
        inventory_map[display_str] = {
            "category": "WIP (Work in Progress)",
            "lot_no": lot_no,
            "product": prod_name,
            "grade": grade_val,
            "max_qty": max_q,
            "row_idx": None,
        }

  if not df_packing_fg.empty:
    fg_available = df_packing_fg.copy()
    fg_available["Clean_Qty"] = fg_available["Quantity"].apply(clean_quantity)
    fg_available = fg_available[fg_available["Clean_Qty"] > 0]
    for _, r in fg_available.iterrows():
      lot_no = str(r.get("THAI STICK Lot. No.", ""))
      prod_name = str(r.get("Export Name", ""))
      grade_val = str(r.get("Grade", ""))
      max_q = float(r["Clean_Qty"])
      if lot_no:
        display_str = (
            f"[FG] {lot_no} | ชื่อ: {prod_name} | เกรด: {grade_val} |"
            f" คงเหลือ: {max_q:,.1f}"
        )
        combined_options.append(display_str)
        inventory_map[display_str] = {
            "category": "FG (Finished Goods)",
            "lot_no": lot_no,
            "product": prod_name,
            "grade": grade_val,
            "max_qty": max_q,
            "row_idx": r.get("Original_Row_Index"),
        }

  if "ship_row_count" not in st.session_state:
    st.session_state.ship_row_count = 3

  col_cnt1, col_cnt2, col_cnt3 = st.columns([2, 2, 6])
  with col_cnt1:
    if st.button("➕ เพิ่มช่องผลิตภัณฑ์"):
      if st.session_state.ship_row_count < 10:
        st.session_state.ship_row_count += 1
        st.rerun()
  with col_cnt2:
    if st.button("➖ ลดช่องผลิตภัณฑ์"):
      if st.session_state.ship_row_count > 1:
        st.session_state.ship_row_count -= 1
        st.rerun()

  with st.form("shipped_out_multi_form"):
    st.subheader("1. ข้อมูลการจัดส่งทั่วไป")
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
      consignee = st.text_input(
          "ชื่อปลายทาง (Consignee)", placeholder="เช่น คลังสินค้ากรุงเทพฯ"
      )
    with col_s2:
      client_name = st.text_input(
          "ลูกค้า (Client)", placeholder="ระบุชื่อบริษัทหรือลูกค้า"
      )
    with col_s3:
      shipping_reason = st.selectbox(
          "เหตุผลในการส่ง (Shipping Reason)",
          ["SAMPLE", "EXPORT", "RETURN", "REJECT", "DOMESTIC SALE"],
      )

    shipment_date_input = st.date_input(
        "วันที่จัดส่ง (Shipment Date)", value=date.today()
    )
    shipment_date_str = str(shipment_date_input)

    st.divider()
    st.subheader(
        f"2. รายการผลิตภัณฑ์ที่ต้องการส่งออก (แสดง {st.session_state.ship_row_count} จาก 10 ช่อง)"
    )

    shipment_items = []

    for i in range(st.session_state.ship_row_count):
      st.markdown(f"**รายการที่ #{i+1}**")
      row_c1, row_c2 = st.columns([4, 2])

      with row_c1:
        chosen_display = st.selectbox(
            f"เลือกรายการสินค้า #{i+1}",
            ["-- กรุณาเลือก Lot สินค้า (WIP หรือ FG) --"] + combined_options,
            key=f"ship_item_{i}",
        )

      with row_c2:
        max_q = 0.0
        if chosen_display and not chosen_display.startswith("--"):
          item_info = inventory_map.get(chosen_display)
          if item_info:
            max_q = item_info["max_qty"]

        ship_q = st.number_input(
            f"จำนวนส่งออก (Max: {max_q:,.1f})",
            min_value=0.0,
            max_value=max_q if max_q > 0 else 1000000.0,
            step=0.1,
            value=0.0,
            key=f"ship_qty_{i}",
        )

      if chosen_display and not chosen_display.startswith("--") and ship_q > 0:
        if ship_q > max_q:
          st.error(
              f"❌ รายการที่ #{i+1}: จำนวนส่งออก ({ship_q})"
              f" เกินกว่าสต็อกที่มีอยู่ ({max_q:,.1f})"
          )
        else:
          item_info = inventory_map[chosen_display]
          shipment_items.append({
              "category": item_info["category"],
              "lot_no": item_info["lot_no"],
              "product": item_info["product"],
              "grade": item_info["grade"],
              "qty": ship_q,
              "row_idx": item_info["row_idx"],
          })
      st.markdown("---")

    note_ship = st.text_area("หมายเหตุการจัดส่งรวม (Note)")
    ship_submitted = st.form_submit_button(
        "🚚 ยืนยันการส่งสินค้าออกทั้งหมด (Multi-Product Shipped Out)"
    )

    if ship_submitted:
      if not consignee.strip():
        st.error("❌ กรุณาระบุชื่อปลายทาง (Consignee)")
      elif not client_name.strip():
        st.error("❌ กรุณาระบุชื่อลูกค้า (Client)")
      elif not shipment_items:
        st.error(
            "❌ กรุณาเลือก Lot สินค้าและระบุจำนวนที่ต้องการส่งมากกว่า 0"
            " อย่างน้อย 1 รายการ"
        )
      else:
        log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        success_count = 0
        for item in shipment_items:
          cat = item["category"]
          lot = item["lot_no"]
          qty = item["qty"]
          prod = item["product"]
          grade = item["grade"]
          row_idx_fg = item["row_idx"]

          if "WIP" in cat:
            wip_cell = sheet.find(str(lot))
            matched_idx = wip_cell.row if wip_cell else 2

            latest_headers = sheet.row_values(1)
            row_values = sheet.row_values(matched_idx)
            while len(row_values) < len(latest_headers):
              row_values.append("")

            current_row_qty = 0.0
            for idx, h in enumerate(latest_headers):
              if h == "Quantity":
                current_row_qty = clean_quantity(row_values[idx])

            remaining_qty = max(0.0, current_row_qty - qty)
            new_status = "Delivered" if remaining_qty <= 0 else "Pending Packing"

            for idx, h in enumerate(latest_headers):
              if h == "Quantity":
                row_values[idx] = str(remaining_qty)
              elif h == "Status":
                row_values[idx] = new_status

            col_end_letter = chr(64 + len(latest_headers)).upper()
            sheet.update(
                f"A{matched_idx}:{col_end_letter}{matched_idx}", [row_values]
            )
          else:
            all_pack_records = packing_records_sheet.get_all_records()
            if row_idx_fg and (row_idx_fg - 2) < len(all_pack_records):
              current_fg_qty = clean_quantity(
                  all_pack_records[row_idx_fg - 2].get("Quantity", 0)
              )
              remaining_qty = max(0.0, current_fg_qty - qty)
              packing_records_sheet.update_cell(row_idx_fg, 7, remaining_qty)

          shipped_sheet.append_row([
              log_time,
              cat,
              lot,
              prod,
              grade,
              str(qty),
              consignee,
              client_name,
              shipping_reason,
              shipment_date_str,
              note_ship,
          ])
          success_count += 1

        st.success(
            f"✅ บันทึกการส่งออกสำเร็จทั้งหมด {success_count}"
            " รายการและหักสต็อกต้นทางเรียบร้อยแล้ว! 🎉"
        )

  st.divider()
  st.subheader("📑 ประวัติการส่งสินค้าออก (Shipped Out Log)")
  shipped_data = shipped_sheet.get_all_records()
  if shipped_data:
    st.dataframe(pd.DataFrame(shipped_data), use_container_width=True)
  else:
    st.info("ยังไม่มีประวัติการส่งสินค้าออกในระบบ")
