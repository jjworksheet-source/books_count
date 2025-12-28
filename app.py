import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="出席報表處理系統", layout="wide")
st.title("出席報表處理系統")

# 側邊欄選擇步驟
st.sidebar.title("操作步驟")
step = st.sidebar.radio(
    "請選擇步驟",
    [
        "1. 上傳與初始處理 (保留A-O, 新增P欄排序)",
        "2. 刪除特定紀錄與排序",
        "3. 刪除重複並產生大數表"
    ]
)

# Session state 初始化
for k in ['df_step1', 'df_step2', 'df_step3']:
    if k not in st.session_state:
        st.session_state[k] = None

def find_col(cols, keywords):
    """在欄位名稱中尋找包含任一關鍵字的欄位，回傳第一個找到的欄位名"""
    for kw in keywords:
        for col in cols:
            if kw in str(col):
                return col
    return None

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

if step == "1. 上傳與初始處理 (保留A-O, 新增P欄排序)":
    st.header("步驟1: 上傳客戶課堂報表並初始處理")
    uploaded_file = st.file_uploader("請上傳 XLS 報表", type=["xls", "xlsx"])
    if uploaded_file:
        try:
            # 讀取 Excel，跳過前 6 行，用第 7 行作為標頭
            df = pd.read_excel(uploaded_file, skiprows=6, header=0, dtype=str)
            df.columns = [str(col).strip() for col in df.columns]
        except Exception as e:
            st.error(f"讀取檔案時發生錯誤: {e}")
            st.stop()
        
        # 保留 A-O 欄 (前15欄，假設A-O)
        columns_to_keep = df.columns[0:15]  # A-O
        df_step1 = df[columns_to_keep].copy()
        
        # 新增 P 欄：老師出席排序
        attendance_map = {'出席': 1, '請假': 2, '跳堂': 3, '病假': 4, '缺席': 5, '代課': 6}
        teacher_att_col = find_col(df_step1.columns, ['老師出席狀況'])
        if teacher_att_col:
            df_step1['老師出席排序'] = df_step1[teacher_att_col].map(attendance_map).fillna(99)
        else:
            st.error("找不到老師出席狀況欄位，請檢查檔案格式。")
            st.stop()
        
        st.session_state.df_step1 = df_step1
        st.success(f"步驟1 完成：保留A-O欄並新增P欄排序，共 {len(df_step1)} 筆資料。")
        st.subheader("步驟1 資料預覽")
        st.dataframe(df_step1.head(10))
        st.download_button(
            label="下載步驟1 Excel",
            data=to_excel(df_step1),
            file_name="step1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif step == "2. 刪除特定紀錄與排序":
    st.header("步驟2: 刪除特定紀錄並排序")
    df_step1 = st.session_state.get('df_step1', None)
    if df_step1 is None:
        st.warning("請先完成步驟1 上傳並產生初始資料。")
    else:
        df_step2 = df_step1.copy()
        
        # 刪除補堂有"由:"的紀錄 (欄位K: 學生出席狀態[補堂])
        makeup_col = find_col(df_step2.columns, ['補堂'])
        if makeup_col:
            df_step2 = df_step2[~df_step2[makeup_col].astype(str).str.contains('由:', na=False)]
        else:
            st.warning("找不到補堂欄位，跳過此刪除。")
        
        # 刪除學生編號以"TAC"開頭的紀錄 (欄位B: 學生編號)
        studentid_col = find_col(df_step2.columns, ['學生編號'])
        if studentid_col:
            df_step2 = df_step2[~df_step2[studentid_col].astype(str).str.startswith('TAC', na=False)]
        else:
            st.warning("找不到學生編號欄位，跳過此刪除。")
        
        # 刪除課室包含"LIVE"的紀錄
        classroom_col = find_col(df_step2.columns, ['課室'])
        if classroom_col:
            df_step2 = df_step2[~df_step2[classroom_col].astype(str).str.contains('LIVE', na=False)]
        else:
            st.warning("找不到課室欄位，跳過此刪除。")
        
        # 排序：學生編號 > 班別 > 老師出席排序
        class_col = find_col(df_step2.columns, ['班別'])
        sort_col = '老師出席排序'
        if studentid_col and class_col and sort_col in df_step2.columns:
            df_step2 = df_step2.sort_values(by=[studentid_col, class_col, sort_col])
        else:
            st.error("找不到排序所需欄位，請檢查檔案格式。")
            st.stop()
        
        st.session_state.df_step2 = df_step2
        st.success(f"步驟2 完成：刪除特定紀錄並排序，共 {len(df_step2)} 筆資料。")
        st.subheader("步驟2 資料預覽")
        st.dataframe(df_step2.head(10))
        st.download_button(
            label="下載步驟2 Excel",
            data=to_excel(df_step2),
            file_name="step2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif step == "3. 刪除重複並產生大數表":
    st.header("步驟3: 刪除重複並產生大數表")
    df_step2 = st.session_state.get('df_step2', None)
    if df_step2 is None:
        st.warning("請先完成步驟2 刪除特定紀錄並排序。")
    else:
        df_step3 = df_step2.copy()
        
        # 刪除重複：基於學生編號 + 班別
        studentid_col = find_col(df_step3.columns, ['學生編號'])
        class_col = find_col(df_step3.columns, ['班別'])
        if studentid_col and class_col:
            df_step3 = df_step3.drop_duplicates(subset=[studentid_col, class_col], keep='first')
        else:
            st.error("找不到去重所需欄位，請檢查檔案格式。")
            st.stop()
        
        # 產生大數表：按班別統計總人數
        name_col = find_col(df_step3.columns, ['學生姓名'])
        if class_col and name_col:
            df_summary = df_step3.groupby(class_col)[name_col].count().reset_index(name='總人數')
            st.session_state.df_step3 = df_summary
            st.success(f"步驟3 完成：刪除重複並產生大數表，共 {len(df_step3)} 筆原始資料，總結表有 {len(df_summary)} 筆。")
            st.subheader("大數表 (按班別統計總人數)")
            st.dataframe(df_summary)
            st.download_button(
                label="下載大數表 Excel",
                data=to_excel(df_summary),
                file_name="summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("找不到產生大數表所需欄位，請檢查檔案格式。")
