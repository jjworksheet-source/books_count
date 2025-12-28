import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="客戶課堂報表清理", layout="wide")
st.title("客戶課堂報表自動清理工具")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

uploaded_file = st.file_uploader("上傳客戶課堂報表 (xls/xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    # 1. 讀取檔案，假設第6列為標題
    try:
        df = pd.read_excel(uploaded_file, header=5, dtype=str)
    except Exception as e:
        st.error(f"讀取檔案時發生錯誤: {e}")
        st.stop()

    # 2. 只保留A-O欄
    df = df.iloc[:, :15].copy()

    # 3. 新增P欄「老師出席排序」
    teacher_att_col = [col for col in df.columns if "老師出席" in str(col)]
    if not teacher_att_col:
        st.error("找不到「老師出席」欄位，請檢查檔案格式。")
        st.stop()
    teacher_att_col = teacher_att_col[0]
    # 前處理：去除空白
    df[teacher_att_col] = df[teacher_att_col].astype(str).str.strip().str.replace('　', '').fillna('')
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
    df["老師出席排序"] = df[teacher_att_col].apply(map_att)

    # 4. 刪除不需要的紀錄
    # 刪除「補堂」欄（K欄）內容是「由」字頭
    makeup_col = [col for col in df.columns if "補堂" in str(col)]
    if makeup_col:
        makeup_col = makeup_col[0]
        df = df[~df[makeup_col].astype(str).str.startswith('由', na=False)]
    # 刪除「學生編號」欄（B欄）內容是「TAC」字頭
    studentid_col = [col for col in df.columns if "學生編號" in str(col)]
    if studentid_col:
        studentid_col = studentid_col[0]
        df = df[~df[studentid_col].astype(str).str.startswith('TAC', na=False)]
    # 刪除「課室」欄內容包含「LIVE」
    classroom_col = [col for col in df.columns if "課室" in str(col)]
    if classroom_col:
        classroom_col = classroom_col[0]
        df = df[~df[classroom_col].astype(str).str.contains('LIVE', na=False)]

    # 5. 排序
    customerid_col = [col for col in df.columns if "客戶編號" in str(col)]
    class_col = [col for col in df.columns if "班別" in str(col)]
    if not customerid_col or not class_col:
        st.error("找不到「客戶編號」或「班別」欄位，請檢查檔案格式。")
        st.stop()
    customerid_col = customerid_col[0]
    class_col = class_col[0]
    df = df.sort_values(by=[customerid_col, class_col, "老師出席排序"])

    # 6. 刪除重複（以「客戶編號」＋「班別」為 key，只保留第一筆）
    df = df.drop_duplicates(subset=[customerid_col, class_col], keep='first')

    # 顯示結果
    st.success(f"清理後剩餘 {len(df)} 筆資料。")
    st.dataframe(df)

    # 下載按鈕
    st.download_button(
        label="下載清理後 Excel",
        data=to_excel(df),
        file_name="cleaned_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
