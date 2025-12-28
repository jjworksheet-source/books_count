import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Jolly Jupiter IT Department", layout="wide")

st.title("中文做卷及書管理")

# Sidebar
with st.sidebar:
    st.markdown("<h2 style='font-size:2em; color:#2c3e50;'>書數</h2>", unsafe_allow_html=True)
    book_func = st.radio(
        "書數功能",
        ["書數有效範圍"]
    )
    st.markdown("---")
    st.markdown("**其他功能**")
    main_func = st.radio(
        "請選擇功能",
        [
            "做卷有效資料",
            "出卷老師資料",
            "分校做卷情況"
        ]
    )

# 書數有效範圍功能（獨立上傳，不依賴 step1）
if book_func == "書數有效範圍":
    st.header("書數有效範圍")
    uploaded_book_file = st.file_uploader("請上傳書數 Excel 檔案 (xls/xlsx)", type=["xls", "xlsx"], key="book_file")
    if uploaded_book_file:
        try:
            df_book = pd.read_excel(uploaded_book_file, dtype=str)
        except Exception as e:
            st.error(f"讀取檔案時發生錯誤: {e}")
            st.stop()
        if df_book.shape[1] < 15:
            st.error("有效資料欄位不足，請檢查資料。")
        else:
            df_range = df_book.iloc[:, :15].copy()
            teacher_status_col = None
            for col in df_range.columns:
                if "老師出席狀態" in str(col):
                    teacher_status_col = col
                    break
            if teacher_status_col is None:
                st.error("找不到老師出席狀態欄，請檢查資料。")
            else:
                status_map = {
                    "出席": "1出席",
                    "請假": "2請假",
                    "跳堂": "3跳堂",
                    "病假": "4病假",
                    "缺席": "5缺席",
                    "代課": "6代課"
                }
                def get_status_sort(val):
                    return status_map.get(str(val).strip(), "")
                df_range["老師出席狀態排序"] = df_range[teacher_status_col].apply(get_status_sort)
                st.dataframe(df_range)
                def to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    return output.getvalue()
                st.download_button(
                    label="下載書數有效範圍 Excel",
                    data=to_excel(df_range),
                    file_name="book_valid_range.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# 其餘功能（做卷有效資料、出卷老師資料、分校做卷情況）照原本 session_state 流程
# ...（其餘功能程式碼不變，直接接在這段後面即可）
