import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel 出席報表處理", layout="wide")
st.title("Excel 出席報表處理 App - 分步執行")

# 分步選單
step = st.sidebar.radio(
    "請選擇步驟",
    [
        "1. 上傳與初始讀取",
        "2. 新增排序欄與刪除補堂/測試記錄",
        "3. 初始排序與刪除重複",
        "4. 刪除 LIVE 課室與無效值",
        "5. 更新月份與二次排序",
        "6. 產生最終大數表"
    ]
)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def find_col(cols, keywords):
    for kw in keywords:
        for col in cols:
            if kw in str(col):
                return col
    return None

# Session state 初始化
for k in ['df_step1','df_step2','df_step3','df_step4','df_step5','df_step6']:
    if k not in st.session_state:
        st.session_state[k] = None

if step == "1. 上傳與初始讀取":
    st.header("步驟 1: 上傳與初始讀取")
    uploaded_file = st.file_uploader("上傳您的 XLS 報表", type=["xls", "xlsx"])
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, header=6-1, dtype=str)  # header=5, 第6列為標題
        except Exception as e:
            st.error(f"讀取檔案時發生錯誤: {e}")
            st.stop()
        columns_to_keep = df.columns[:15]  # A-O 欄
        df_step1 = df[columns_to_keep]
        st.session_state.df_step1 = df_step1
        st.dataframe(df_step1)
        st.write("欄位名稱檢查:", df_step1.columns.tolist())
        st.download_button(
            label="下載步驟 1 Excel",
            data=to_excel(df_step1),
            file_name="step1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif step == "2. 新增排序欄與刪除補堂/測試記錄":
    st.header("步驟 2: 新增排序欄與刪除補堂/測試記錄")
    df_step1 = st.session_state.get('df_step1', None)
    if df_step1 is None:
        st.warning("請先執行步驟 1 上傳資料。")
    else:
        df_step2 = df_step1.copy()
        attendance_map = {'出席': 1, '請假': 2, '跳堂': 3, '病假': 4, '缺席': 5, '代課': 6}
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
        st.dataframe(df_step2)
        st.write("欄位名稱檢查:", df_step2.columns.tolist())
        st.download_button(
            label="下載步驟 2 Excel",
            data=to_excel(df_step2),
            file_name="step2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif step == "3. 初始排序與刪除重複":
    st.header("步驟 3: 初始排序與刪除重複")
    df_step2 = st.session_state.get('df_step2', None)
    if df_step2 is None:
        st.warning("請先執行步驟 2。")
    else:
        df_step3 = df_step2.copy()
        studentid_col = find_col(df_step3.columns, ['學生編號'])
        class_col = find_col(df_step3.columns, ['班別'])
        sort_col = find_col(df_step3.columns, ['老師出席排序'])
        if studentid_col and class_col and sort_col:
            df_step3 = df_step3.sort_values(by=[studentid_col, class_col, sort_col])
            df_step3 = df_step3.drop_duplicates(subset=[studentid_col, class_col])
        else:
            st.error("找不到排序或去重所需欄位，請檢查欄位名稱！")
        st.session_state.df_step3 = df_step3
        st.dataframe(df_step3)
        st.write("欄位名稱檢查:", df_step3.columns.tolist())
        st.download_button(
            label="下載步驟 3 Excel",
            data=to_excel(df_step3),
            file_name="step3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif step == "4. 刪除 LIVE 課室與無效值":
    st.header("步驟 4: 刪除 LIVE 課室與無效值")
    df_step3 = st.session_state.get('df_step3', None)
    if df_step3 is None:
        st.warning("請先執行步驟 3。")
    else:
        df_step4 = df_step3.copy()
        classroom_col = find_col(df_step4.columns, ['課室'])
        if classroom_col:
            df_step4 = df_step4[~df_step4[classroom_col].astype(str).str.contains('LIVE', na=False)]
        # 欠數總額欄位（如有）
        owe_col = find_col(df_step4.columns, ['欠數總額', 'AE'])
        if owe_col:
            df_step4 = df_step4[(df_step4[owe_col] != 0) & (df_step4[owe_col].notna())]
        st.session_state.df_step4 = df_step4
        st.dataframe(df_step4)
        st.write("欄位名稱檢查:", df_step4.columns.tolist())
        st.download_button(
            label="下載步驟 4 Excel",
            data=to_excel(df_step4),
            file_name="step4.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif step == "5. 更新月份與二次排序":
    st.header("步驟 5: 更新月份與二次排序")
    df_step4 = st.session_state.get('df_step4', None)
    if df_step4 is None:
        st.warning("請先執行步驟 4。")
    else:
        month_input = st.text_input("輸入月份 (e.g., 2026-02)", value="2026-02")
        df_step5 = df_step4.copy()
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
        st.dataframe(df_step5)
        st.write("欄位名稱檢查:", df_step5.columns.tolist())
        st.download_button(
            label="下載步驟 5 Excel",
            data=to_excel(df_step5),
            file_name="step5.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif step == "6. 產生最終大數表":
    st.header("步驟 6: 產生最終大數表")
    df_step5 = st.session_state.get('df_step5', None)
    if df_step5 is None:
        st.warning("請先執行步驟 5。")
    else:
        name_col = find_col(df_step5.columns, ['學栍姓名', '學生姓名'])
        class_col = find_col(df_step5.columns, ['班別'])
        if name_col and class_col:
            df_summary = df_step5.groupby(class_col)[name_col].count().reset_index(name='總人數')
            st.session_state.df_step6 = df_summary
            st.dataframe(df_summary)
            st.write("欄位名稱檢查:", df_summary.columns.tolist())
            st.download_button(
                label="下載步驟 6 Excel (大數表)",
                data=to_excel(df_summary),
                file_name="step6_summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("找不到產生大數表所需欄位，請檢查欄位名稱！")
