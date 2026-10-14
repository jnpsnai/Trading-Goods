from datetime import date
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import streamlit as st

# --- 1. ตั้งค่าหน้าเว็บ Streamlit ให้เป็น Wide Mode ---
st.set_page_config(
    page_title="Thai Stick Trading Portal", page_icon="🌱", layout="wide"
)

# --- 2. Custom CSS สำหรับตกแต่ง Template ให้สวยงามพรีเมียม ---
st.markdown(
    """
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border: 1px solid #e3e6f0;
    }
    .stButton>button {
        width: 100%;
        background-color: #2e7d32;
        color: white;
        border-radius: 5px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #1b5e20;
        color: white;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# --- 3. ฟังก์ชันเชื่อมต่อ Google Sheets ---
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
sheet = client.open(SHEET_NAME).sheet1

# --- 4. Sidebar เมนูด้านซ้าย ---
st.sidebar.title("🌱 Thai Stick Trading")
st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "เลือกเมนูการใช้งาน",
    ["📝 Trading Goods Registration", "📊 Overall Stock Dashboard"],
)
st.sidebar.markdown("---")
st.sidebar.info(
    "ระบบบันทึกและจัดการข้อมูลสินค้า Trading เชื่อมต่อกับ Google Sheets แบบเรียลไทม์"
)


# --- ฟังก์ชันแปลง Quantity เป็นตัวเลขเพื่อคำนวณ ---
def clean_quantity(val):
  try:
    return float(val)
  except:
    return 0.0


# ==========================================
# หน้าที่ 1: Trading Goods Registration (ลงข้อมูล)
# ==========================================
if menu == "📝 Trading Goods Registration":
  st.title("📝 Trading Goods Registration")
  st.markdown("กรอกข้อมูลผลิตภัณฑ์และรายละเอียดล็อตสินค้าใหม่ลงในระบบ")

  with st.form("trading_form", clear_on_submit=True):
    col1, col2 = st.columns(2)

    with col1:
      cultivar_name = st.text_input("1. Cultivar Name")
      supplier_name = st.text_input("2. Supplier Name")
      supplier_lot = st.text_input("3. Supplier Lot. No.")
      export_name = st.text_input("4. Export Name")
      thai_stick_lot = st.text_input("5. THAI STICK Lot. No.")
      grade = st.selectbox(
          "6. Grade", ["Big", "Small", "Mix", "Grade A", "Premium"]
      )

    with col2:
      quantity = st.number_input("7. Quantity", min_value=0.0, step=0.1)
      inspection_req = st.selectbox("8. Inspection Req.", ["Yes", "No"])
      ir_req = st.selectbox("9. IR Req.", ["Yes", "No"])

      # 10. Shipment date พร้อม checkbox 'To be decided' ไว้ด้านล่าง
      date_val = st.date_input("10. Shipment date", value=date.today())
      is_tbd = st.checkbox("To be decided")

      if is_tbd:
        shipment_date = "To be decided"
      else:
        shipment_date = str(date_val)

      # 11. Type Dropdown
      prod_type = st.selectbox(
          "11. Type",
          [
              "Inspection & Packing Service",
              "Trading Goods (TS)",
              "Consignment",
              "In House Product",
          ],
      )

      # 12. Status Dropdown
      status = st.selectbox(
          "12. Status",
          [
              "Trading Goods Received to Stock",
              "Inspection Completed",
              "Pending Packing",
              "Delivered",
              "Waiting for Return",
          ],
      )

    st.markdown("")
    submitted = st.form_submit_button("💾 บันทึกข้อมูลลง Google Sheets")

    if submitted:
      row_data = [
          cultivar_name,
          supplier_name,
          supplier_lot,
          export_name,
          thai_stick_lot,
          grade,
          str(quantity),
          inspection_req,
          ir_req,
          shipment_date,
          prod_type,
          status,
      ]
      sheet.append_row(row_data)
      st.success("บันทึกข้อมูลสำเร็จเรียบร้อยแล้ว! 🎉")

  # แสดงตารางข้อมูลล่าสุดด้านล่างฟอร์ม
  st.divider()
  st.subheader("📋 ประวัติข้อมูลทั้งหมดในระบบ")
  data = sheet.get_all_records()
  if data:
    df = pd.DataFrame(data)
    st.dataframe(df, use_container_width=True)
  else:
    st.info("ยังไม่มีข้อมูลในระบบ Google Sheets")


# ==========================================
# หน้าที่ 2: Overall Stock (สรุปและแดชบอร์ด)
# ==========================================
elif menu == "📊 Overall Stock Dashboard":
  st.title("📊 Overall Stock Dashboard")
  st.markdown("ภาพรวมสต็อกสินค้าคงเหลือ (On-hand Stock)")

  data = sheet.get_all_records()

  if data:
    df = pd.DataFrame(data)

    # แปลงคอลัมน์ Quantity ให้เป็นตัวเลข
    if "Quantity" in df.columns:
      df["Clean_Qty"] = df["Quantity"].apply(clean_quantity)
    else:
      df["Clean_Qty"] = 0.0

    # เงื่อนไขกรอง On-hand stock: Status ไม่ใช่ 'Delivered' และ Type ไม่ใช่ 'Inspection & Packing Service'
    filtered_df = df.copy()

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
      ]

    # คำนวณยอดแยกตามเกรด (Big, Small, Mix)
    if "Grade" in filtered_df.columns:
      stock_big = filtered_df[
          filtered_df["Grade"].astype(str).str.strip().str.lower() == "big"
      ]["Clean_Qty"].sum()
      stock_small = filtered_df[
          filtered_df["Grade"].astype(str).str.strip().str.lower() == "small"
      ]["Clean_Qty"].sum()
      stock_mix = filtered_df[
          filtered_df["Grade"].astype(str).str.strip().str.lower() == "mix"
      ]["Clean_Qty"].sum()
      total_on_hand = filtered_df["Clean_Qty"].sum()
    else:
      stock_big, stock_small, stock_mix, total_on_hand = 0, 0, 0, 0

    # --- แสดงผล Cardboard (Metrics) สรุป On-hand stock ---
    st.subheader("📦 On-hand Stock Summary (แยกตามเกรด)")
    st.markdown(
        "*เงื่อนไขการคำนวณ: ไม่รวมรายการที่มี Status เป็น 'Delivered' และ Type"
        " เป็น 'Inspection & Packing Service'"
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
      st.metric(label="🟢 Grade: Big", value=f"{stock_big:,.2f}")
    with col2:
      st.metric(label="🔵 Grade: Small", value=f"{stock_small:,.2f}")
    with col3:
      st.metric(label="🟠 Grade: Mix", value=f"{stock_mix:,.2f}")
    with col4:
      st.metric(label="✨ Total On-Hand Stock", value=f"{total_on_hand:,.2f}")

    st.divider()

    # แสดงตารางข้อมูล On-hand Stock ที่ผ่านการกรองแล้ว
    st.subheader("📑 รายละเอียดรายการสินค้าในคลัง (Active Stock)")
    st.dataframe(filtered_df, use_container_width=True)

    with st.expander("🔍 ดูข้อมูลดิบทั้งหมด (Raw Data ทั้งหมดในระบบ)"):
      st.dataframe(df, use_container_width=True)

  else:
    st.warning("ยังไม่มีข้อมูลใน Google Sheets สำหรับประมวลผลสต็อก")
