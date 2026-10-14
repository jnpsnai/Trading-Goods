from datetime import date, datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import streamlit as st

# --- 1. ตั้งค่าหน้าเว็บ Streamlit ให้เป็น Wide Mode ---
st.set_page_config(
    page_title="Thai Stick Trading Portal", page_icon="🌱", layout="wide"
)

# --- 2. Custom CSS สำหรับตกแต่ง Template ---
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
sheet = sh.get_worksheet(0)  # แผ่นหลักสำหรับ Registration

# กำหนดโครงสร้าง Column Header ที่ต้องการ
EXPECTED_HEADERS = [
    "Cultivar Name",
    "Supplier Name",
    "Supplier Lot. No.",
    "Export Name",
    "THAI STICK Lot. No.",
    "Category",
    "Grade",
    "Quantity",
    "Inspection Req.",
    "IR Req.",
    "Shipment date",
    "Type",
    "Status",
]

# ตรวจสอบและซิงค์ Header ให้ตรงกันอัตโนมัติ
current_headers = sheet.row_values(1)
if not current_headers or len(current_headers) < len(EXPECTED_HEADERS):
  sheet.update("A1:M1", [EXPECTED_HEADERS])

# ตรวจสอบและสร้าง Sheet ย่อยสำหรับ Packing Log
try:
  packing_sheet = sh.worksheet("Packing_Log")
except:
  packing_sheet = sh.add_worksheet(title="Packing_Log", rows="1000", cols="10")
  packing_sheet.append_row([
      "Timestamp",
      "Source WIP Lot",
      "Cultivar",
      "Packed FG Name",
      "Grade",
      "Packed Qty",
      "Waste Qty",
      "Note",
  ])

# --- 4. Sidebar เมนูด้านซ้าย ---
st.sidebar.title("🌱 Thai Stick Trading")
st.sidebar.markdown("---")
menu = st.sidebar.radio(
    "เลือกเมนูการใช้งาน",
    [
        "📝 Trading Goods Registration",
        "✏️ Edit Registration",
        "📦 Packing Module",
        "📊 Overall Stock Dashboard",
    ],
)
st.sidebar.markdown("---")
st.sidebar.info("ระบบจัดการสินค้า Trading และกระบวนการ Packing เชื่อมต่อ G-Sheets")


# --- ฟังก์ชันแปลง Quantity เป็นตัวเลข ---
def clean_quantity(val):
  try:
    return float(val)
  except:
    return 0.0


# ==========================================
# หน้าที่ 1: Trading Goods Registration
# ==========================================
if menu == "📝 Trading Goods Registration":
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
          "Category", ["FG (Finished Goods)", "WIP (Work in Progress)"]
      )

      grades_opt = ["Big", "Small", "Mix", "Grade A", "Premium"]
      grade = st.selectbox("6. Grade", grades_opt)

    with col2:
      quantity = st.number_input("7. Quantity", min_value=0.0, step=0.1)
      inspection_req = st.selectbox("8. Inspection Req.", ["Yes", "No"])
      ir_req = st.selectbox("9. IR Req.", ["Yes", "No"])

      date_val = st.date_input("10. Shipment date", value=date.today())
      is_tbd = st.checkbox("To be decided")
      shipment_date = "To be decided" if is_tbd else str(date_val)

      types_opt = [
          "Inspection & Packing Service",
          "Trading Goods (TS)",
          "Consignment",
          "In House Product",
      ]
      prod_type = st.selectbox("11. Type", types_opt)

      status_opt = [
          "Trading Goods Received to Stock",
          "Inspection Completed",
          "Pending Packing",
          "Packing done",
          "Delivered",
          "Waiting for Return",
      ]
      status = st.selectbox("12. Status", status_opt)

    st.markdown("")
    submitted = st.form_submit_button("💾 บันทึกข้อมูลลง Google Sheets")

    if submitted:
      latest_headers = sheet.row_values(1)
      row_dict = {
          "Cultivar Name": cultivar_name,
          "Supplier Name": supplier_name,
          "Supplier Lot. No.": supplier_lot,
          "Export Name": export_name,
          "THAI STICK Lot. No.": thai_stick_lot,
          "Category": category,
          "Grade": grade,
          "Quantity": str(quantity),
          "Inspection Req.": inspection_req,
          "IR Req.": ir_req,
          "Shipment date": shipment_date,
          "Type": prod_type,
          "Status": status,
      }
      row_data = [str(row_dict.get(h, "")) for h in latest_headers]
      sheet.append_row(row_data)
      st.success("บันทึกข้อมูลสำเร็จเรียบร้อยแล้ว! 🎉")

  st.divider()
  st.subheader("📋 ประวัติข้อมูลทั้งหมดในระบบ")
  data = sheet.get_all_records()
  if data:
    df = pd.DataFrame(data)
    st.dataframe(df, use_container_width=True)
  else:
    st.info("ยังไม่มีข้อมูลในระบบ Google Sheets")


# ==========================================
# หน้าที่ 2: Edit Registration
# ==========================================
elif menu == "✏️ Edit Registration":
  st.title("✏️ Edit Registration Data")
  st.markdown("ค้นหาและแก้ไขข้อมูลล็อตสินค้าที่เคยบันทึกไว้")

  data = sheet.get_all_records()
  if data:
    df_all = pd.DataFrame(data)
    lot_column = (
        "THAI STICK Lot. No."
        if "THAI STICK Lot. No." in df_all.columns
        else df_all.columns[4]
    )

    lot_list = df_all[lot_column].astype(str).tolist()
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
      current_data = df_all[df_all[lot_column].astype(str) == selected_lot].iloc[
          0
      ]

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
              "Category",
              cat_opt,
              index=cat_opt.index(old_cat) if old_cat in cat_opt else 0,
          )

          grades_opt = ["Big", "Small", "Mix", "Grade A", "Premium"]
          old_grade = str(current_data.get("Grade", "Big"))
          e_grade = st.selectbox(
              "6. Grade",
              grades_opt,
              index=grades_opt.index(old_grade)
              if old_grade in grades_opt
              else 0,
          )

        with col2:
          try:
            old_qty = float(current_data.get("Quantity", 0.0))
          except:
            old_qty = 0.0
          e_qty = st.number_input(
              "7. Quantity", min_value=0.0, step=0.1, value=old_qty
          )

          insp_opt = ["Yes", "No"]
          old_insp = str(current_data.get("Inspection Req.", "No"))
          e_insp = st.selectbox(
              "8. Inspection Req.",
              insp_opt,
              index=insp_opt.index(old_insp) if old_insp in insp_opt else 0,
          )

          ir_opt = ["Yes", "No"]
          old_ir = str(current_data.get("IR Req.", "No"))
          e_ir = st.selectbox(
              "9. IR Req.",
              ir_opt,
              index=ir_opt.index(old_ir) if old_ir in ir_opt else 0,
          )

          old_date = str(current_data.get("Shipment date", ""))
          is_tbd_default = old_date == "To be decided"

          if is_tbd_default:
            default_dt = date.today()
          else:
            try:
              default_dt = datetime.strptime(old_date, "%Y-%m-%d").date()
            except:
              default_dt = date.today()

          e_is_tbd = st.checkbox(
              "To be decided", value=is_tbd_default, key="edit_tbd"
          )
          e_date_input = st.date_input(
              "10. Shipment date", value=default_dt, key="edit_date"
          )
          e_shipment_date = "To be decided" if e_is_tbd else str(e_date_input)

          types_opt = [
              "Inspection & Packing Service",
              "Trading Goods (TS)",
              "Consignment",
              "In House Product",
          ]
          old_type = str(current_data.get("Type", ""))
          e_type = st.selectbox(
              "11. Type",
              types_opt,
              index=types_opt.index(old_type) if old_type in types_opt else 0,
          )

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
              "12. Status",
              status_opt,
              index=status_opt.index(old_status)
              if old_status in status_opt
              else 0,
          )

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
              "Quantity": str(e_qty),
              "Inspection Req.": e_insp,
              "IR Req.": e_ir,
              "Shipment date": e_shipment_date,
              "Type": e_type,
              "Status": e_status,
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
# หน้าที่ 3: Packing Module
# ==========================================
elif menu == "📦 Packing Module":
  st.title("📦 Packing Module (WIP Conversion & Waste Logging)")
  st.markdown(
      "ดึงวัตถุดิบ WIP จากการ Register มาทำการแพ็ก บันทึกยอดผลิตภัณฑ์ที่ได้"
      " พร้อมตัดสต็อกอัตโนมัติ"
  )

  data = sheet.get_all_records()
  if data:
    df = pd.DataFrame(data)

    if "Category" in df.columns:
      wip_df = df[
          df["Category"]
          .astype(str)
          .str.contains("WIP", case=False, na=False)
      ].copy()
    else:
      wip_df = pd.DataFrame()

    if not wip_df.empty:
      wip_df["Clean_Qty"] = wip_df["Quantity"].apply(clean_quantity)
      available_wip = wip_df[wip_df["Clean_Qty"] > 0]

      if not available_wip.empty:
        wip_lot_list = available_wip["THAI STICK Lot. No."].astype(str).tolist()

        selected_wip_lot = st.selectbox(
            "🔍 เลือก WIP Lot ที่ต้องการนำมาแพ็ก", wip_lot_list
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
          st.subheader("รายละเอียดการแพ็กและของเสีย (Packing & Waste Input)")

          col1, col2 = st.columns(2)
          with col1:
            packed_fg_name = st.text_input(
                "Packed Product / Export Name (ชื่อผลิตภัณฑ์หลังแพ็ก)"
            )
            packed_grade = st.selectbox(
                "Grade ผลลัพธ์หลังแพ็ก",
                ["Big", "Small", "Mix", "Grade A", "Premium"],
            )
            packed_qty = st.number_input(
                "จำนวนที่แพ็กได้สำเร็จ (Packed Qty)",
                min_value=0.0,
                max_value=max_wip_qty,
                step=0.1,
            )

          with col2:
            waste_qty = st.number_input(
                "จำนวนของเสีย/สูญเสีย (Waste Qty)",
                min_value=0.0,
                max_value=max_wip_qty,
                step=0.1,
            )
            note = st.text_area("หมายเหตุเพิ่มเติม (Note)")

          pack_submitted = st.form_submit_button(
              "🚀 ยืนยันการแพ็กและตัดสต็อก WIP"
          )

          if pack_submitted:
            total_deduct = packed_qty + waste_qty
            if total_deduct > max_wip_qty:
              st.error(
                  "❌ จำนวนผลผลิตรวมกับของเสีย มากกว่ายอด WIP คงเหลือที่มีอยู่!"
              )
            else:
              new_wip_qty = max_wip_qty - total_deduct

              if new_wip_qty <= 0:
                new_status = "Packing done"
              else:
                new_status = "Pending Packing"

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
                  row_values[idx] = str(new_wip_qty)
                elif h == "Status":
                  row_values[idx] = new_status

              col_end_letter = chr(64 + len(latest_headers)).upper()
              sheet.update(
                  f"A{matched_idx}:{col_end_letter}{matched_idx}",
                  [row_values],
              )

              log_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
              packing_sheet.append_row([
                  log_time,
                  selected_wip_lot,
                  cultivar_val,
                  packed_fg_name,
                  packed_grade,
                  str(packed_qty),
                  str(waste_qty),
                  note,
              ])

              st.success(
                  f"✅ บันทึกการแพ็กสำเร็จ! หักยอด WIP คงเหลือเหลือ"
                  f" {new_wip_qty:,.2f} สถานะปรับเป็น '{new_status}' เรียบร้อยแล้ว"
              )
      else:
        st.warning(
            "⚠️ ไม่มี WIP Lot ที่มีสต็อกคงเหลือ (Quantity > 0) ให้แพ็กในขณะนี้"
        )
    else:
      st.warning(
          "⚠️ ไม่พบคอลัมน์ Category ในระบบ กรุณาตรวจสอบข้อมูลใน Google Sheets"
      )

    st.divider()
    st.subheader("📜 ประวัติการ Packing ทั้งหมด (Packing Log)")
    log_data = packing_sheet.get_all_records()
    if log_data:
      st.dataframe(pd.DataFrame(log_data), use_container_width=True)
    else:
      st.info("ยังไม่มีประวัติการแพ็กสินค้า")
  else:
    st.warning("ยังไม่มีข้อมูลในระบบ Google Sheets")


# ==========================================
# หน้าที่ 4: Overall Stock Dashboard
# ==========================================
elif menu == "📊 Overall Stock Dashboard":
  st.title("📊 Overall Stock Dashboard")
  st.markdown("ภาพรวมสต็อกสินค้าคงเหลือ (On-hand Stock)")

  data = sheet.get_all_records()

  if data:
    df = pd.DataFrame(data)

    if "Quantity" in df.columns:
      df["Clean_Qty"] = df["Quantity"].apply(clean_quantity)
    else:
      df["Clean_Qty"] = 0.0

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

    st.subheader("📦 On-hand Stock Summary (แยกตามเกรด)")
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

    st.subheader("📑 รายละเอียดรายการสินค้าในคลัง (Active Stock)")
    st.dataframe(filtered_df, use_container_width=True)

    with st.expander("🔍 ดูข้อมูลดิบทั้งหมด (Raw Data ทั้งหมดในระบบ)"):
      st.dataframe(df, use_container_width=True)
  else:
    st.warning("ยังไม่มีข้อมูลในระบบ Google Sheets สำหรับประมวลผลสต็อก")
