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
sheet = sh.get_worksheet(0)  # แผ่นหลักสำหรับ Registration (Trading Goods)

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
    "Inspection Req.",
    "IR Req.",
    "Shipment date",
    "Type",
    "Status",
    "Packing date",
]

# ตรวจสอบและซิงค์ Header แผ่นหลัก
current_headers = sheet.row_values(1)
if not current_headers or len(current_headers) < len(EXPECTED_HEADERS):
  sheet.update("A1:O1", [EXPECTED_HEADERS])

# ตรวจสอบและสร้าง Sheet ย่อยสำหรับ Packing Records
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
      "Grade & Qty Breakdown",
      "Total Packed Qty",
      "Waste Qty",
      "Packing Date",
      "Note",
  ])

# ตรวจสอบและสร้าง Sheet ย่อยสำหรับ Shipped Out Log
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
        "📝 Trading Goods Registration",
        "✏️ Edit Registration",
        "📦 Packing Module",
        "🚚 Shipped Out Module",
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
          "6. Category", ["FG (Finished Goods)", "WIP (Work in Progress)"]
      )
      grades_opt = ["Big", "Small", "Mix", "F", "Grade A", "Premium"]
      grade = st.selectbox("7. Grade", grades_opt)

    with col2:
      receive_date_val = st.date_input(
          "8. Receive date (Column H)", value=date.today(), key="reg_receive_date"
      )
      receive_date_str = str(receive_date_val)

      quantity = st.number_input("9. Quantity", min_value=0.0, step=0.1)
      inspection_req = st.selectbox("10. Inspection Req.", ["Yes", "No"])
      ir_req = st.selectbox("11. IR Req.", ["Yes", "No"])

      date_val = st.date_input("12. Shipment date", value=date.today())
      is_tbd = st.checkbox("To be decided")
      shipment_date = "To be decided" if is_tbd else str(date_val)

      types_opt = [
          "Inspection & Packing Service",
          "Trading Goods (TS)",
          "Consignment",
          "In House Product",
      ]
      prod_type = st.selectbox("13. Type", types_opt)

      status_opt = [
          "Trading Goods Received to Stock",
          "Inspection Completed",
          "Pending Packing",
          "Packing done",
          "Delivered",
          "Waiting for Return",
      ]
      status = st.selectbox("14. Status", status_opt)

      pack_date_val = st.date_input(
          "15. Packing date (Column O)", value=date.today(), key="reg_pack_date"
      )
      packing_date_str = str(pack_date_val)

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
          "Receive date": receive_date_str,
          "Quantity": str(quantity),
          "Inspection Req.": inspection_req,
          "IR Req.": ir_req,
          "Shipment date": shipment_date,
          "Type": prod_type,
          "Status": status,
          "Packing date": packing_date_str,
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

            insp_opt = ["Yes", "No"]
            old_insp = str(current_data.get("Inspection Req.", "No"))
            e_insp = st.selectbox(
                "10. Inspection Req.",
                insp_opt,
                index=insp_opt.index(old_insp) if old_insp in insp_opt else 0,
            )

            ir_opt = ["Yes", "No"]
            old_ir = str(current_data.get("IR Req.", "No"))
            e_ir = st.selectbox(
                "11. IR Req.",
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
                "12. Shipment date", value=default_dt, key="edit_date"
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
                "13. Type",
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
# หน้าที่ 3: Packing Module
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
# หน้าที่ 4: Shipped Out Module (แก้ไขการดึงข้อมูล WIP/FG แยกส่วนถูกต้อง 100%)
# ==========================================
elif menu == "🚚 Shipped Out Module":
  st.title("🚚 Shipped Out Module (Multi-Product Shipping & Stock Deduction)")
  st.markdown(
      "บันทึกการส่งสินค้าออก (Delivery / Shipment) สำหรับ WIP และ FG"
      " สามารถเลือกส่งออกหลายผลิตภัณฑ์พร้อมกันในครั้งเดียว"
      " ระบบจะหักสต็อกต้นทางอัตโนมัติ"
  )

  # ดึงข้อมูล Main Sheet และ Packing_Records มาประมวลผลสต็อกปัจจุบัน
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
          "Supplier Name": "",
          "Supplier Lot. No.": row.get("Source WIP Lot", ""),
          "Export Name": row.get("Export Name", ""),
          "THAI STICK Lot. No.": row.get("Generated FG Lot No.", ""),
          "Category": "FG (Finished Goods)",
          "Grade": row.get("Grade", ""),
          "Receive date": row.get("Packing Date", ""),
          "Quantity": clean_quantity(row.get("Quantity", 0)),
          "Inspection Req.": "No",
          "IR Req.": "No",
          "Shipment date": "To be decided",
          "Type": "In House Product",
          "Status": "Trading Goods Received to Stock",
          "Packing date": row.get("Packing Date", ""),
      })
    df_packing_fg = pd.DataFrame(converted_rows)

  # เตรียมรายการ WIP Lot ที่มีสต็อกคงเหลือ
  wip_options = []
  wip_display_map = {}
  wip_dict = {}
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
            f"{lot_no} | {prod_name} | เกรด: {grade_val} | คงเหลือ: {max_q:,.1f}"
        )
        wip_options.append(display_str)
        wip_display_map[display_str] = lot_no
        wip_dict[lot_no] = {
            "max_qty": max_q,
            "product": prod_name,
            "grade": grade_val,
        }

  # เตรียมรายการ FG Lot ที่มีสต็อกคงเหลือ
  fg_options = []
  fg_display_map = {}
  fg_dict = {}
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
            f"{lot_no} | {prod_name} | เกรด: {grade_val} | คงเหลือ: {max_q:,.1f}"
        )
        fg_options.append(display_str)
        fg_display_map[display_str] = lot_no
        fg_dict[lot_no] = {
            "max_qty": max_q,
            "product": prod_name,
            "grade": grade_val,
            "row_idx": r.get("Original_Row_Index"),
        }

  # จัดการ Session State สำหรับจำนวนแถวสินค้า (เริ่มต้น 3 แถว)
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
      row_c1, row_c2, row_c3 = st.columns([2, 4, 2])

      with row_c1:
        # Set default เป็น FG (Finished Goods) โดยให้อยู่ลำดับแรก
        item_cat = st.selectbox(
            "ประเภท",
            ["FG (Finished Goods)", "WIP (Work in Progress)"],
            key=f"cat_{i}",
            label_visibility="collapsed",
        )
      with row_c2:
        if "FG" in item_cat:
          chosen_display = st.selectbox(
              "เลือก FG Lot",
              ["-- เลือก Lot FG (Lot | ชื่อ | เกรด | คงเหลือ) --"] + fg_options,
              key=f"lot_fg_{i}",
              label_visibility="collapsed",
          )
          chosen_lot = (
              fg_display_map.get(chosen_display, "")
              if chosen_display and not chosen_display.startswith("--")
              else ""
          )
        else:
          chosen_display = st.selectbox(
              "เลือก WIP Lot",
              ["-- เลือก Lot WIP (Lot | ชื่อ | เกรด | คงเหลือ) --"] + wip_options,
              key=f"lot_wip_{i}",
              label_visibility="collapsed",
          )
          chosen_lot = (
              wip_display_map.get(chosen_display, "")
              if chosen_display and not chosen_display.startswith("--")
              else ""
          )

      with row_c3:
        max_q = 0.0
        if chosen_lot:
          if "FG" in item_cat and chosen_lot in fg_dict:
            max_q = fg_dict[chosen_lot]["max_qty"]
          elif "WIP" in item_cat and chosen_lot in wip_dict:
            max_q = wip_dict[chosen_lot]["max_qty"]

        ship_q = st.number_input(
            f"จำนวน (Max: {max_q:,.1f})",
            min_value=0.0,
            max_value=max_q if max_q > 0 else 0.0,
            step=0.1,
            key=f"qty_{i}",
        )

      if chosen_lot and ship_q > 0 and max_q > 0:
        prod_name = (
            fg_dict[chosen_lot]["product"]
            if "FG" in item_cat
            else wip_dict[chosen_lot]["product"]
        )
        grade_val = (
            fg_dict[chosen_lot]["grade"]
            if "FG" in item_cat
            else wip_dict[chosen_lot]["grade"]
        )
        row_idx = (
            fg_dict[chosen_lot]["row_idx"] if "FG" in item_cat else None
        )

        shipment_items.append({
            "category": (
                "FG (Finished Goods)" if "FG" in item_cat else "WIP (Work in Progress)"
            ),
            "lot_no": chosen_lot,
            "product": prod_name,
            "grade": grade_val,
            "qty": ship_q,
            "row_idx": row_idx,
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
            "❌ กรุณาเลือก Lot สินค้าและระบุจำนวนที่ต้องการส่งอย่างน้อย 1 รายการ"
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

          if "WIP" in cat:
            wip_info = wip_dict[lot]
            remaining_qty = wip_info["max_qty"] - qty
            new_status = "Delivered" if remaining_qty <= 0 else "Pending Packing"

            cell_wip = sheet.find(str(lot))
            matched_idx = cell_wip.row if cell_wip else 2

            latest_headers = sheet.row_values(1)
            row_values = sheet.row_values(matched_idx)
            while len(row_values) < len(latest_headers):
              row_values.append("")

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
            fg_info = fg_dict[lot]
            remaining_qty = fg_info["max_qty"] - qty
            pack_rec_row_idx = int(fg_info["row_idx"])
            packing_records_sheet.update_cell(pack_rec_row_idx, 7, remaining_qty)

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


# ==========================================
# หน้าที่ 5: Overall Stock Dashboard
# ==========================================
elif menu == "📊 Overall Stock Dashboard":
  st.title("📊 Overall Stock Dashboard")
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
      st.metric(label="✨ ยอดรวมสต็อกทั้งหมด (Total Stock)", value=f"{total_on_hand:,.2f}")
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
          wip_stock_df["Grade"].astype(str).str.strip().str.lower() == "small"
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
          fg_stock_df["Grade"].astype(str).str.strip().str.lower() == "small"
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
