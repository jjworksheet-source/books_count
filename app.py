import streamlit as st
import pandas as pd

st.title("Excel 出席報表處理 App")

# 上傳檔案
uploaded_file = st.file_uploader("上傳您的 XLS 報表", type=["xls", "xlsx"])

if uploaded_file is not None:
    # 讀取 Excel，跳過前 6 行元數據，讓第 7 行成為欄位頭
    df = pd.read_excel(uploaded_file, skiprows=6, engine='xlrd')  # 使用 xlrd 讀 .xls，如果是 .xlsx 可改 openpyxl

    # 清理欄位名稱（先轉 str 再去除空格，避免非 str 錯誤）
    df.columns = [str(col).strip() for col in df.columns]

    # 顯示原始數據預覽和欄位檢查
    st.subheader("原始數據預覽")
    st.dataframe(df.head(10))
    st.subheader("欄位名稱檢查")
    st.write(df.columns.tolist())

    # 步驟 2: 複製 A-O 欄（假設前 15 欄，從 '學栍姓名' 到 '課室'）
    columns_to_keep = df.columns[0:15]
    df_subset = df[columns_to_keep]

    # 步驟 4: 新增 P 欄 - 老師出席排序
    attendance_map = {
        '出席': 1,
        '請假': 2,
        '跳堂': 3,
        '病假': 4,
        '缺席': 5,
        '代課': 6
    }
    df_subset['老師出席排序'] = df_subset['老師出席狀況'].map(attendance_map).fillna(99)  # 未匹配設為 99

    # 步驟 5: 刪除補堂欄位包含 '由:' 的記錄
    df_subset = df_subset[~df_subset['補堂'].str.contains('由:', na=False)]

    # 步驟 6: 刪除學生編號以 'TAC' 開頭的記錄
    df_subset = df_subset[~df_subset['學生編號'].str.startswith('TAC', na=False)]

    # 步驟 7 & 12: 排序 - 學生編號 > 班別 > 老師出席排序
    df_subset = df_subset.sort_values(by=['學生編號', '班別', '老師出席排序'])

    # 步驟 8: 刪除重複 - 基於學生編號 + 班別
    df_subset = df_subset.drop_duplicates(subset=['學生編號', '班別'])

    # 步驟 9: 刪除課室包含 'LIVE' 的記錄
    df_cleaned = df_subset[~df_subset['課室'].str.contains('LIVE', na=False)]

    # 步驟 11: 刪除 AE 欄位（假設為 '欠數總額'）的 0 或 NA 記錄
    # 注意：根據檔案，調整欄位名，如果不是 '欠數總額'，請改成正確的
    if '欠數總額' in df_cleaned.columns:
        df_cleaned = df_cleaned[(df_cleaned['欠數總額'] != 0) & (df_cleaned['欠數總額'].notna())]

    # 步驟 10: 更新月份 - 加 UI 輸入
    month_input = st.text_input("輸入月份 (e.g., 2026-02)", value="2026-02")
    # 示例：過濾上課日期包含該月份（可擴充邏輯）
    df_cleaned = df_cleaned[df_cleaned['上課日期'].str.contains(month_input, na=False)]

    # 步驟 13: 產生大數表 - 示例：按班別計數學生
    summary = df_cleaned.groupby('班別')['學栍姓名'].count().reset_index(name='總人數')
    st.subheader("大數表彙總")
    st.dataframe(summary)

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
