import streamlit as st
import pandas as pd

st.title("Excel 出席報表處理 App - 分步執行")

# 上傳檔案
uploaded_file = st.file_uploader("上傳您的 XLS 報表", type=["xls", "xlsx"])

if uploaded_file is not None:
    # 讀取 Excel，跳過前 6 行，用第 7 行作為標頭
    df = pd.read_excel(uploaded_file, skiprows=6, header=0, engine='xlrd')

    # 清理欄位名稱（去除空格）
    df.columns = [str(col).strip() for col in df.columns]

    # 使用 session_state 儲存每步結果
    if 'df_step1' not in st.session_state:
        st.session_state.df_step1 = None
    if 'df_step2' not in st.session_state:
        st.session_state.df_step2 = None
    if 'df_step3' not in st.session_state:
        st.session_state.df_step3 = None
    if 'df_step4' not in st.session_state:
        st.session_state.df_step4 = None
    if 'df_step5' not in st.session_state:
        st.session_state.df_step5 = None
    if 'df_step6' not in st.session_state:
        st.session_state.df_step6 = None

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
            data=st.session_state.df_step1.to_excel(index=False, engine='openpyxl'),
            file_name="step1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 步驟 2: 新增排序欄並刪除補堂/測試記錄
    st.subheader("步驟 2: 新增排序欄並刪除補堂/測試記錄 (步驟 4-6)")
    if st.button("執行步驟 2"):
        df_step2 = st.session_state.df_step1.copy()
        attendance_map = {'出席': 1, '請假': 2, '跳堂': 3, '病假': 4, '缺席': 5, '代課': 6}
        df_step2['老師出席排序'] = df_step2['老師出席狀況'].map(attendance_map).fillna(99)
        df_step2 = df_step2[~df_step2['補堂'].str.contains('由:', na=False)]
        df_step2 = df_step2[~df_step2['學生編號'].str.startswith('TAC', na=False)]
        st.session_state.df_step2 = df_step2
    if st.session_state.df_step2 is not None:
        st.dataframe(st.session_state.df_step2.head(10))
        st.write("欄位名稱檢查:", st.session_state.df_step2.columns.tolist())
        st.download_button(
            label="下載步驟 2 Excel",
            data=st.session_state.df_step2.to_excel(index=False, engine='openpyxl'),
            file_name="step2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 步驟 3: 初始排序與刪除重複
    st.subheader("步驟 3: 初始排序與刪除重複 (步驟 7-8)")
    if st.button("執行步驟 3"):
        df_step3 = st.session_state.df_step2.copy()
        df_step3 = df_step3.sort_values(by=['學生編號', '班別', '老師出席排序'])
        df_step3 = df_step3.drop_duplicates(subset=['學生編號', '班別'])
        st.session_state.df_step3 = df_step3
    if st.session_state.df_step3 is not None:
        st.dataframe(st.session_state.df_step3.head(10))
        st.write("欄位名稱檢查:", st.session_state.df_step3.columns.tolist())
        st.download_button(
            label="下載步驟 3 Excel",
            data=st.session_state.df_step3.to_excel(index=False, engine='openpyxl'),
            file_name="step3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 步驟 4: 刪除 LIVE 課室與無效值
    st.subheader("步驟 4: 刪除 LIVE 課室與無效值 (步驟 9,11)")
    if st.button("執行步驟 4"):
        df_step4 = st.session_state.df_step3.copy()
        df_step4 = df_step4[~df_step4['課室'].str.contains('LIVE', na=False)]
        if '欠數總額' in df_step4.columns:
            df_step4 = df_step4[(df_step4['欠數總額'] != 0) & (df_step4['欠數總額'].notna())]
        st.session_state.df_step4 = df_step4
    if st.session_state.df_step4 is not None:
        st.dataframe(st.session_state.df_step4.head(10))
        st.write("欄位名稱檢查:", st.session_state.df_step4.columns.tolist())
        st.download_button(
            label="下載步驟 4 Excel",
            data=st.session_state.df_step4.to_excel(index=False, engine='openpyxl'),
            file_name="step4.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 步驟 5: 更新月份與二次排序
    st.subheader("步驟 5: 更新月份與二次排序 (步驟 10,12)")
    month_input = st.text_input("輸入月份 (e.g., 2026-02)", value="2026-02")
    if st.button("執行步驟 5"):
        df_step5 = st.session_state.df_step4.copy()
        df_step5['上課日期'] = pd.to_datetime(df_step5['上課日期'], errors='coerce')
        df_step5 = df_step5[df_step5['上課日期'].dt.strftime('%Y-%m') == month_input]
        df_step5 = df_step5.sort_values(by=['學生編號', '班別', '老師出席排序'])
        st.session_state.df_step5 = df_step5
    if st.session_state.df_step5 is not None:
        st.dataframe(st.session_state.df_step5.head(10))
        st.write("欄位名稱檢查:", st.session_state.df_step5.columns.tolist())
        st.download_button(
            label="下載步驟 5 Excel",
            data=st.session_state.df_step5.to_excel(index=False, engine='openpyxl'),
            file_name="step5.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # 步驟 6: 產生最終大數表
    st.subheader("步驟 6: 產生最終大數表 (步驟 13)")
    if st.button("執行步驟 6"):
        df_step6 = st.session_state.df_step5.groupby('班別')['學栍姓名'].count().reset_index(name='總人數')
        st.session_state.df_step6 = df_step6
    if st.session_state.df_step6 is not None:
        st.dataframe(st.session_state.df_step6)
        st.write("欄位名稱檢查:", st.session_state.df_step6.columns.tolist())
        st.download_button(
            label="下載步驟 6 Excel (大數表)",
            data=st.session_state.df_step6.to_excel(index=False, engine='openpyxl'),
            file_name="step6_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
