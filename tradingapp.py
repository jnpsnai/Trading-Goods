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

# --- 4. Sidebar เมนูด้านซ้าย (เปลี่ยนเป็นแบบคลิกเลือกด้วย radio) ---
st.sidebar.title("🌱 Thai Stick Trading")
st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "เลือกเมนูการใช้งาน",
    ["📝 Trading Registration", "📊 Overall Stock Dashboard"],
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
# หน้าที่ 1: Trading Registration (ลงข้อมูล)
# ==========================================
if menu == "📝 Trading Registration":
  st.title("📝 Trading Data Registration")
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
      shipment_date = st.date_input("10. Shipment date", value=date.today())
      prod_type = st.text_input(
          "11. Type", placeholder="เช่น Product, Service, etc."
      )
      status = st.selectbox(
          "12. Status", ["Pending", "In Progress", "Completed", "Delivery"]
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
          str(shipment_date),
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

    # เงื่อนไขกรอง On-hand stock: Status ไม่ใช่ 'Delivery' และ Type ไม่ใช่ 'Service'
    filtered_df = df.copy()

    if "Status" in filtered_df.columns and "Type" in filtered_df.columns:
      filtered_df = filtered_df[
          (filtered_df["Status"].astype(str).str.strip().str.lower() != "delivery")
          & (filtered_df["Type"].astype(str).str.strip().str.lower() != "service")
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
      stock_big, stock_
