elif step == "3. 新增老師出席排序與刪除補堂/測試記錄":
    st.header("步驟 3: 新增老師出席排序與刪除補堂/測試記錄")
    df_step2 = st.session_state.get('df_step2', None)
    if df_step2 is None:
        st.warning("請先執行步驟 2。")
    else:
        df_step3 = df_step2.copy()
        # 老師出席狀況排序
        teacher_att_col = find_col(df_step3.columns, ['老師出席狀況', '老師出席'])
        if teacher_att_col:
            # 前處理：去除空白、全形空白、NaN
            df_step3[teacher_att_col] = df_step3[teacher_att_col].astype(str).str.strip().str.replace('　', '').fillna('')
            # 顯示所有唯一值方便 debug
            st.write("老師出席狀況唯一值：", df_step3[teacher_att_col].unique())
            # mapping 用「包含」判斷
            def map_att(val):
                if "出席" in val:
                    return 1
                elif "請假" in val:
                    return 2
                elif "跳堂" in val:
                    return 3
                elif "病假" in val:
                    return 4
                elif "缺席" in val:
                    return 5
                elif "代課" in val:
                    return 6
                else:
                    return 99
            df_step3['老師出席排序'] = df_step3[teacher_att_col].apply(map_att)
        else:
            st.error("找不到老師出席狀況欄位，請檢查欄位名稱！")
        # 刪除補堂有紀錄的行
        makeup_col = find_col(df_step3.columns, ['補堂'])
        if makeup_col:
            df_step3 = df_step3[~df_step3[makeup_col].astype(str).str.contains('由:', na=False)]
        # 刪除學生編號TAC開頭
        studentid_col = find_col(df_step3.columns, ['學生編號'])
        if studentid_col:
            df_step3 = df_step3[~df_step3[studentid_col].astype(str).str.startswith('TAC', na=False)]
        st.session_state.df_step3 = df_step3
        st.dataframe(df_step3)
        st.write("欄位名稱檢查:", df_step3.columns.tolist())
        st.download_button(
            label="下載步驟 3 Excel",
            data=to_excel(df_step3),
            file_name="step3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
