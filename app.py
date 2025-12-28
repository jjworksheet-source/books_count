import streamlit as st
import pandas as pd

st.title("Excel 出席報表處理 App")

# 上傳檔案
uploaded_file = st.file_uploader("上傳您的 XLS 報表", type=["xls", "xlsx"])

if uploaded_file is not None:
    # 讀取 Excel（假設第一張 sheet）
    df = pd.read_excel(uploaded_file, sheet_name=0)  # 如果有特定 sheet 名，調整這裡

    # 顯示原始數據預覽
    st.subheader("原始數據預覽")
    st.dataframe(df.head(10))  # 只顯示前 10 行，避免過載

    # 這裡新增您的清理步驟（逐步擴充）
    # 範例：複製 A-O 欄（調整欄位名根據您的檔案）
    columns_to_keep = df.columns[0:15]  # 假設 A-O 是前 15 欄
    df_subset = df[columns_to_keep]

    # 範例：刪除包含 'LIVE' 的課室記錄
    df_cleaned = df_subset[~df_subset['課室'].str.contains('LIVE', na=False)]

    # 顯示清理後預覽
    st.subheader("清理後數據預覽")
    st.dataframe(df_cleaned.head(10))

    # 下載按鈕
    st.download_button(
        label="下載清理後的 Excel",
        data=df_cleaned.to_excel(index=False, engine='openpyxl'),
        file_name="cleaned_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
