            st.markdown('<div class="portal-empty">ยังไม่มีข้อมูลสต็อกในระบบ เมื่อบันทึกรับสินค้าแล้ว ยอดคงเหลือจะแสดงที่หน้านี้</div>', unsafe_allow_html=True)
    # ==========================================
    # หน้าที่ 3: Monthly Report
    # ==========================================
    elif menu == "📈 Monthly Report":
        render_monthly_page(sh)

    # ==========================================
    # หน้าที่ 4: Trading Goods Registration
    # ==========================================
    elif menu == "📝 Trading Goods Registration":
        render_registration_ui_styles()
        with st.container(key="registration_create"):
            header, shortcut = st.columns([4, 1.4], gap="large")
            with header:
                with st.container(key="registration_header"):
                    st.markdown('<p class="registration-eyebrow">NEW REGISTRATION</p>', unsafe_allow_html=True)
                    st.title("Trading Goods Registration")
                    st.caption("ลงทะเบียน WIP / FG พร้อมข้อมูลผู้จัดหา เลข Lot และเงื่อนไขการจัดส่ง")
            with shortcut:
                st.link_button("เปิด Google Sheets", sh.url if hasattr(sh, "url") else "https://docs.google.com/spreadsheets", width="stretch")
            render_registration_recovery()
            flash = st.session_state.get("registration_flash")
            if flash and flash["mode"] == "create":
                st.success(st.session_state.pop("registration_flash")["message"])
            st.session_state.setdefault("registration_receipt_transaction_id", "RECEIVE-" + uuid4().hex)
            submitted, registration = render_registration_form()
            if submitted:
                submit_registration_operation({
                    "mode": "create", "transaction_id": st.session_state["registration_receipt_transaction_id"],
                    "record": registration,
                })

    # ==========================================
    # หน้าที่ 5: Edit Registration
    # ==========================================
    elif menu == "✏️ Edit Registration":
        render_registration_ui_styles()
        with st.container(key="registration_edit"):
            with st.container(key="registration_header"):
                st.markdown('<p class="registration-eyebrow">EDIT REGISTRATION</p>', unsafe_allow_html=True)
                st.title("Edit Registration")
                st.caption("ค้นหาล็อตสินค้า ตรวจสอบข้อมูลปัจจุบัน และแก้ไขรายละเอียดในแบบฟอร์มด้านล่าง")
            render_registration_recovery()
            flash = st.session_state.get("registration_flash")
            if flash and flash["mode"] == "edit":
                st.success(st.session_state.pop("registration_flash")["message"])
            data = get_cached_data("main", sheet, include_row_index=True)
            if data:
                df_all = pd.DataFrame(data)
                df_all.columns = [str(column).strip() for column in df_all.columns]
                df_all["Clean_Qty"] = df_all["Quantity"].apply(lambda value: picking_quantity(value) or 0.0) if "Quantity" in df_all else 0.0
                active = df_all["Clean_Qty"].gt(0)
                if "Status" in df_all:
                    active &= ~df_all["Status"].astype(str).str.strip().str.casefold().eq("delivered")
                editable_df = df_all[active].copy()
                if editable_df.empty:
                    st.info("ไม่มีล็อตสินค้าที่สามารถแก้ไขได้ในขณะนี้")
                else:
                    with st.container(key="registration_filter"):
                        st.markdown("**ค้นหาล็อตที่ต้องการแก้ไข**")
                        search_col, category_col = st.columns([3, 1], gap="large")
                        with search_col:
                            query = st.text_input("ค้นหาชื่อสายพันธุ์, Supplier หรือ Lot", placeholder="พิมพ์ชื่อหรือเลข Lot ที่ต้องการค้นหา", key="registration_edit_search")
                        with category_col:
                            selected_category = st.selectbox("ประเภทสต็อก", ["ทุกประเภท", "WIP", "FG"], key="registration_edit_category")
                        filtered = editable_df
                        if query.strip():
                            fields = [field for field in ["Cultivar Name", "Export Name", "Supplier Name", "Supplier Lot. No.", "THAI STICK Lot. No."] if field in filtered]
                            search_text = filtered[fields].fillna("").astype(str).agg(" ".join, axis=1)
                            filtered = filtered[search_text.str.casefold().str.contains(query.strip().casefold(), regex=False)]
                        if selected_category != "ทุกประเภท":
                            categories = filtered["Category"].astype(str) if "Category" in filtered else pd.Series("", index=filtered.index)
                            filtered = filtered[categories.str.contains(selected_category, case=False, na=False, regex=False)]
                        st.caption(f"พบ {len(filtered):,} รายการ จากล็อตที่แก้ไขได้ทั้งหมด {len(editable_df):,} รายการ · แสดงเฉพาะสต็อกคงเหลือที่ยังไม่ Delivered")
                        records_by_row = {int(row["Original_Row_Index"]): row.to_dict() for _, row in filtered.iterrows()}
                        labels = {}
                        for index, (row_number, record) in enumerate(records_by_row.items(), start=1):
                            cultivar = picking_text(record.get("Cultivar Name")) or "ไม่ระบุสายพันธุ์"
                            lot = picking_text(record.get("THAI STICK Lot. No.")) or picking_text(record.get("Supplier Lot. No.")) or "ไม่ระบุ Lot"
                            grade = picking_text(record.get("Grade")) or "—"
                            labels[row_number] = f"{index:02d} · {cultivar} | Lot {lot} | {grade} | {record['Clean_Qty']:,.0f} g"
                        if records_by_row:
                            selected_row = st.selectbox("เลือกล็อตสินค้า", list(records_by_row), format_func=labels.get, key="registration_edit_lot")
                    if not records_by_row:
                        st.info("ไม่พบล็อตที่ตรงกับการค้นหา ลองเปลี่ยนคำค้นหรือเลือกทุกประเภท")
                    else:
                        current = records_by_row[selected_row]
                        candidate = core.inventory_from_records([current], [], include_unavailable=True)[0]
                        expected_key = "registration_expected_" + candidate["inventory_id"]
                        st.session_state.setdefault(expected_key, dict(candidate))
                        expected = st.session_state[expected_key]
                        transaction_key = "registration_edit_transaction_" + candidate["inventory_id"]
                        st.session_state.setdefault(transaction_key, "EDIT-" + uuid4().hex)
                        render_registration_selected(current)
                        st.button("โหลดข้อมูลล็อตนี้ใหม่ / ทิ้งร่าง", key="registration_reload_lot",
                                  on_click=reload_registration_edit, args=(selected_row, candidate["inventory_id"]),
                                  help="ทิ้งข้อมูลที่กรอกค้างและโหลดข้อมูลล่าสุด ใช้เมื่อมีผู้ใช้อื่นแก้ไขหรือส่งสินค้าล็อตนี้แล้ว")
                        submitted, registration = render_registration_form(editing=True, current_data=current, row_number=selected_row, inventory_id=candidate["inventory_id"])
                        if submitted:
                            submit_registration_operation({
                                "mode": "edit", "transaction_id": st.session_state[transaction_key],
                                "inventory_id": candidate["inventory_id"], "record": registration,
                                "expected": expected,
                            })
            else:
                st.info("ยังไม่มีข้อมูลในระบบให้แก้ไข")


    # ==========================================
    # หน้าที่ 6: Packing Module
    # ==========================================
    elif menu == "📦 Packing Module":
        render_packing_page(sh)

    elif menu == "🚚 Shipped Out Module":
        render_shipping_page(sh)

    loaded_at = get_display_loaded_at(sh)
    if loaded_at is not None:
        sync_stamp.caption("โหลดข้อมูลล่าสุด " + loaded_at.strftime("%d/%m/%Y %H:%M:%S") + " (เวลาไทย)")

except Exception as error:
    logging.getLogger(__name__).exception("Portal page failed")
    st.error(f"ไม่สามารถแสดงข้อมูลหน้านี้ได้: {read_error_message(error)}")
    st.info("กดรีเฟรชข้อมูลหรือลองใหม่ ระบบจะไม่แสดงยอดว่างแทนข้อมูลที่โหลดไม่สำเร็จ")
