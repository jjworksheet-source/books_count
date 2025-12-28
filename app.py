import streamlit as st
import pandas as pd

st.set_page_config(page_title="Jolly Jupiter IT Department", layout="wide")

# ===== App Title =====
st.title("書數計算系統")

# ===== Sidebar with step-by-step templates (UI kept, but only first step is functional) =====
st.sidebar.title("操作步驟")
step = st.sidebar.radio(
    "請選擇步驟",
    [
        "1. 做卷有效資料",
        "2. 出卷老師資料",
        "3. 分校做卷情況",
        "4. 其他"
    ]
)

# ===== Global storage for valid_data (kept for UI compatibility with later steps) =====
if 'valid_data' not in st.session_state:
    st.session_state['valid_data'] = None

# ===== Step 1: 有效範圍 =====
if step == "1. 做卷有效資料":
    st.header("有效範圍")
    uploaded_file = st.file_uploader("請上傳 Excel 檔案 (xls/xlsx)", type=["xls", "xlsx"])

    if uploaded_file:
        try:
            # 讀取 Excel，保留原本第一列當標題
            df = pd.read_excel(uploaded_file, dtype=str)
        except Exception as e:
            st.error(f"讀取檔案時發生錯誤: {e}")
            st.stop()

        # 只顯示 A 到 O 欄
        if df.shape[1] >= 15:
            df_limited = df.iloc[:, :15]   # A~O (0~14)
        else:
            # 若欄位少於 15 欄，就顯示全部
            df_limited = df.copy()

        st.subheader("有效範圍 (只顯示 A 至 O 欄)")
        st.dataframe(df_limited)

        # ===== Step 2: 加老師出席排序 =====
        st.header("加老師出席排序")
        # 在原本 df_limited 的基礎上新增一欄 P，內容為空字串
        df_with_p = df_limited.copy()
        df_with_p["老師出席排序"] = ""  # 這會是第 16 欄 (P 欄)

        st.subheader("已新增 P 欄（老師出席排序，暫時為空）")
        st.dataframe(df_with_p)

        # 將處理後的資料存進 session_state，供未來步驟使用
        st.session_state['valid_data'] = df_with_p

else:
    # 其它步驟暫時不實作，只保留 UI
    st.header("此步驟暫未開放")
    st.info("目前只開放『1. 做卷有效資料』中的「有效範圍」與「加老師出席排序」功能，其餘功能將於稍後提供。")
