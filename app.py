import streamlit as st
import pandas as pd
import io

st.title("Excel 出席報表處理 App - 分步執行")

uploaded_file = st.file_uploader("上傳您的 XLS 報表", type=["xls", "xlsx"])

def find_col(cols, keywords):
    """在欄位名稱中尋找包含任一關鍵字的欄位，回傳第一個找到的欄位名"""
    for kw in keywords:
        for col in cols:
            if kw in str(col):
                return col
    return None

def to_excel_bytes(df):
    output = io.BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    return output.getvalue()

if uploaded_file is not None:
    # 讀取 Excel，跳過前 6 行，用第 7 行作為標頭
    df = pd.read_excel(uploaded_file, skiprows=6, header=0, engine='xlrd')
    df.columns = [str(col).strip() for col in df.columns]

    # Session state 初始化
    for k in ['df_step1','df_step2','df_step3','df_step4','df_step5','df_step6']:
        if k not in st.session_state:
            st.session_state[k] = None

    # 步驟 1: 上傳與初始讀取
    st.subheader("步驟 1: 上傳與初始讀取 (步驟 1-3)")
    if st.button("執行步驟 1"):
        columns_to_keep = df.columns[0:15]  # A-O 欄
        df_step1 = df[columns_to_keep]
        st.session_state.df_step1 = df_step1
    if st.session_state.df_step1 is not None:
        st.dataframe(st.session_state.df_step1.head(10))
        st.write("欄位名稱檢查:", st.session_state.df_step1.columns.tolist())
        st.download_button(
            label="下載步驟 1 Excel",
            data=to_excel_bytes(st.session_state.df_step1),
            file_name="step1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 步驟 2: 新增排序欄並刪除補堂/測試記錄
    st.subheader("步驟 2: 新增排序欄並刪除補堂/測試記錄 (步驟 4-6)")
    if st.button("執行步驟 2"):
        df_step2 = st.session_state.df_step1.copy()
        attendance_map = {'出席': 1, '請假': 2, '跳堂': 3, '病假': 4, '缺席': 5, '代課': 6}
        # 自動偵測欄位
        attendance_col = find_col(df_step2.columns, ['出席'])
        makeup_col = find_col(df_step2.columns, ['補堂'])
        studentid_col = find_col(df_step2.columns, ['學生編號'])
        # 新增排序欄
        if attendance_col:
            df_step2['老師出席排序'] = df_step2[attendance_col].map(attendance_map).fillna(99)
        else:
            st.error("找不到出席狀況欄位，請檢查欄位名稱！")
        # 刪除補堂有紀錄的行
        if makeup_col:
            df_step2 = df_step2[~df_step2[makeup_col].astype(str).str.contains('由:', na=False)]
        # 刪除學生編號TAC開頭
        if studentid_col:
            df_step2 = df_step2[~df_step2[studentid_col].astype(str).str.startswith('TAC', na=False)]
        st.session_state.df_step2 = df_step2
    if st.session_state.df_step2 is not None:
        st.dataframe(st.session_state.df_step2.head(10))
        st.write("欄位名稱檢查:", st.session_state.df_step2.columns.tolist())
        st.download_button(
            label="下載步驟 2 Excel",
            data=to_excel_bytes(st.session_state.df_step2),
            file_name="step2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 步驟 3: 初始排序與刪除重複
    st.subheader("步驟 3: 初始排序與刪除重複 (步驟 7-8)")
    if st.button("執行步驟 3"):
        df_step3 = st.session_state.df_step2.copy()
        studentid_col = find_col(df_step3.columns, ['學生編號'])
        class_col = find_col(df_step3.columns, ['班別'])
        sort_col = find_col(df_step3.columns, ['老師出席排序'])
        if studentid_col and class_col and sort_col:
            df_step3 = df_step3.sort_values(by=[studentid_col, class_col, sort_col])
            df_step3 = df_step3.drop_duplicates(subset=[studentid_col, class_col])
        else:
            st.error("找不到排序或去重所需欄位，請檢查欄位名稱！")
        st.session_state.df_step3 = df_step3
    if st.session_state.df_step3 is not None:
        st.dataframe(st.session_state.df_step3.head(10))
        st.write("欄位名稱檢查:", st.session_state.df_step3.columns.tolist())
        st.download_button(
            label="下載步驟 3 Excel",
            data=to_excel_bytes(st.session_state.df_step3),
            file_name="step3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 步驟 4: 刪除 LIVE 課室與無效值
    st.subheader("步驟 4: 刪除 LIVE 課室與無效值 (步驟 9,11)")
    if st.button("執行步驟 4"):
        df_step4 = st.session_state.df_step3.copy()
        classroom_col = find_col(df_step4.columns, ['課室'])
        if classroom_col:
            df_step4 = df_step4[~df_step4[classroom_col].astype(str).str.contains('LIVE', na=False)]
        # 欠數總額欄位（如有）
        owe_col = find_col(df_step4.columns, ['欠數總額', 'AE'])
        if owe_col:
            df_step4 = df_step4[(df_step4[owe_col] != 0) & (df_step4[owe_col].notna())]
        st.session_state.df_step4 = df_step4
    if st.session_state.df_step4 is not None:
        st.dataframe(st.session_state.df_step4.head(10))
        st.write("欄位名稱檢查:", st.session_state.df_step4.columns.tolist())
        st.download_button(
            label="下載步驟 4 Excel",
            data=to_excel_bytes(st.session_state.df_step4),
            file_name="step4.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 步驟 5: 更新月份與二次排序
    st.subheader("步驟 5: 更新月份與二次排序 (步驟 10,12)")
    month_input = st.text_input("輸入月份 (e.g., 2026-02)", value="2026-02")
    if st.button("執行步驟 5"):
        df_step5 = st.session_state.df_step4.copy()
        date_col = find_col(df_step5.columns, ['上課日期'])
        studentid_col = find_col(df_step5.columns, ['學生編號'])
        class_col = find_col(df_step5.columns, ['班別'])
        sort_col = find_col(df_step5.columns, ['老師出席排序'])
        if date_col:
            df_step5[date_col] = pd.to_datetime(df_step5[date_col], errors='coerce')
            df_step5 = df_step5[df_step5[date_col].dt.strftime('%Y-%m') == month_input]
        if studentid_col and class_col and sort_col:
            df_step5 = df_step5.sort_values(by=[studentid_col, class_col, sort_col])
        st.session_state.df_step5 = df_step5
    if st.session_state.df_step5 is not None:
        st.dataframe(st.session_state.df_step5.head(10))
        st.write("欄位名稱檢查:", st.session_state.df_step5.columns.tolist())
        st.download_button(
            label="下載步驟 5 Excel",
            data=to_excel_bytes(st.session_state.df_step5),
            file_name="step5.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 步驟 6: 產生最終大數表
    st.subheader("步驟 6: 產生最終大數表 (步驟 13)")
    if st.button("執行步驟 6"):
        df_step6 = st.session_state.df_step5.copy()
        name_col = find_col(df_step6.columns, ['學栍姓名', '學生姓名'])
        class_col = find_col(df_step6.columns, ['班別'])
        if name_col and class_col:
            df_summary = df_step6.groupby(class_col)[name_col].count().reset_index(name='總人數')
            st.session_state.df_step6 = df_summary
        else:
            st.error("找不到產生大數表所需欄位，請檢查欄位名稱！")
    if st.session_state.df_step6 is not None:
        st.dataframe(st.session_state.df_step6)
        st.write("欄位名稱檢查:", st.session_state.df_step6.columns.tolist())
        st.download_button(
            label="下載步驟 6 Excel (大數表)",
            data=to_excel_bytes(st.session_state.df_step6),
            file_name="step6_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
