if juan_page == "做卷有效資料":
    st.header("做卷有效資料")
    uploaded_file = st.file_uploader("請上傳 JJCustomer 報表 (xls/xlsx)", type=["xls", "xlsx"], key="main_file")
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, header=5, dtype=str)
        except Exception as e:
            st.error(f"讀取檔案時發生錯誤: {e}")
            st.stop()
        class_types = [
            "etup 測考卷 - 高小",
            "etlp 測考卷 - 初小",
            "etup 測考卷 - 高小 - 1小時",
            "etlp 測考卷 - 初小 - 1小時",
            "erlp 閱讀卷 - 初小",
            "erup 閱讀卷 - 高小",
            "ewup 寫作卷 - 高小",
            "ewlp 寫作卷 - 初小"
        ]
        class_col = [col for col in df.columns if "班別" in str(col)]
        if not class_col:
            st.error("找不到班別欄位，請檢查檔案格式。")
            st.stop()
        class_col = class_col[0]
        df_filtered = df[df[class_col].astype(str).str.contains('|'.join(class_types), na=False)]
        att_col = [col for col in df.columns if "學生出席狀況" in str(col)]
        if not att_col:
            st.error("找不到學生出席狀況欄位，請檢查檔案格式。")
            st.stop()
        att_col = att_col[0]
        df_filtered = df_filtered[df_filtered[att_col] == "出席"]
        id_col = [col for col in df.columns if "學生編號" in str(col)][0]
        name_col = [col for col in df.columns if "學栍姓名" in str(col) or "學生姓名" in str(col)][0]
        date_col = [col for col in df.columns if "上課日期" in str(col)][0]
        time_col = [col for col in df.columns if "時間" in str(col)][0]
        teacher_status_col = [col for col in df.columns if "老師出席狀況" in str(col)]
        teacher_status_col = teacher_status_col[0] if teacher_status_col else None
        group_cols = [id_col, name_col, date_col, class_col, time_col]
        if teacher_status_col:
            def pick_row(group):
                not_leave = group[group[teacher_status_col] != "請假"]
                return not_leave.iloc[0] if not_leave.shape[0] > 0 else group.iloc[0]
            df_valid = df_filtered.groupby(group_cols, as_index=False).apply(pick_row).reset_index(drop=True)
        else:
            df_valid = df_filtered.drop_duplicates(subset=group_cols, keep='first')
        grade_col = [col for col in df_valid.columns if "年級" in str(col)]
        school_col = [col for col in df_valid.columns if "學校" in str(col)]
        if not grade_col or not school_col:
            st.error("找不到年級或學校欄位，請檢查檔案格式。")
            st.stop()
        grade_col = grade_col[0]
        school_col = school_col[0]
        def extract_school_short(s):
            if pd.isna(s):
                return ""
            s = str(s)
            if s.startswith("_"):
                s = s[1:]
            result = ""
            for ch in s:
                if '\u4e00' <= ch <= '\u9fff':
                    result += ch
                elif ch == "_":
                    break
            return result
        def make_grade_juan(row):
            grade = str(row[grade_col]).strip() if not pd.isna(row[grade_col]) else ""
            school = extract_school_short(row[school_col])
            juan = f"{grade}{school}_"
            class_val = str(row[class_col]) if not pd.isna(row[class_col]) else ""
            if "1小時" in class_val:
                juan += "1小時"
            return juan
        df_valid["年級_卷"] = df_valid.apply(make_grade_juan, axis=1)
        def get_teacher(juan):
            if juan in cb_list:
                return "cb"
            elif juan in kt_list:
                return "kt"
            elif juan in mc_list:
                return "mc"
            else:
                return ""
        df_valid["出卷老師"] = df_valid["年級_卷"].apply(get_teacher)
        columns = [col for col in df_valid.columns if col not in ["年級_卷", "出卷老師"]]
        columns += ["年級_卷", "出卷老師"]
        df_valid = df_valid[columns]
        merged = df_filtered.merge(df_valid[group_cols], on=group_cols, how='left', indicator=True)
        df_duplicates = merged.loc[merged['_merge'] == 'left_only', df_filtered.columns]
        st.success(f"有效資料共 {len(df_valid)} 筆，重複資料共 {len(df_duplicates)} 筆。")
        st.subheader("有效資料")
        st.dataframe(df_valid)
        st.subheader("重複資料")
        st.dataframe(df_duplicates)
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()
        st.download_button(
            label="下載有效資料 Excel",
            data=to_excel(df_valid),
            file_name="valid_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(
            label="下載重複資料 Excel",
            data=to_excel(df_duplicates),
            file_name="duplicate_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.session_state['valid_data'] = df_valid
