import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="報表分析系統", layout="wide")
st.title("報表分析系統")

# 第一層選單
main_page = st.sidebar.radio(
    "請選擇報表類型",
    ["書數預算", "做卷資料"]
)

# 書數預算子選單
if main_page == "書數預算":
    st.sidebar.markdown("---")
    book_page = st.sidebar.radio(
        "書數預算功能",
        ["書數有效範圍", "刪除步驟", "排序和刪除重覆"]
    )

    # 書數有效範圍：上傳並暫存 處理後的 DataFrame
    if book_page == "書數有效範圍":
        st.header("書數有效範圍")
        uploaded_book_file = st.file_uploader("請上傳書數 Excel 檔案 (xls/xlsx)", type=["xls", "xlsx"], key="book_file")
        if uploaded_book_file:
            try:
                df_book = pd.read_excel(uploaded_book_file, header=5, dtype=str)
            except Exception as e:
                st.error(f"讀取檔案時發生錯誤: {e}")
                st.stop()
            if df_book.shape[1] < 15:
                st.error("有效資料欄位不足，請檢查資料。")
            else:
                df_range = df_book.iloc[:, :15].copy()
                teacher_status_col = None
                for col in df_range.columns:
                    if "老師出席狀況" in str(col):
                        teacher_status_col = col
                        break
                if teacher_status_col is None:
                    st.error("找不到老師出席狀況欄，請檢查資料。")
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
                    df_range["老師出席狀況排序"] = df_range[teacher_status_col].apply(get_status_sort)
                    st.session_state['book_range_df'] = df_range  # 暫存處理後的結果
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

    # 刪除步驟：直接用 session_state['book_range_df']
    elif book_page == "刪除步驟":
        st.header("書數預算 - 刪除步驟")
        if 'book_range_df' not in st.session_state:
            st.warning("請先在『書數有效範圍』上傳並產生有效範圍結果。")
        else:
            df_range = st.session_state['book_range_df'].copy()
            original_count = len(df_range)
            # 1. 刪除欄位K（補堂）以「由」開頭的紀錄
            col_k = [col for col in df_range.columns if col.strip() == "補堂"]
            if col_k:
                df_range = df_range[~df_range[col_k[0]].astype(str).str.startswith("由", na=False)]
            # 2. 刪除欄位B（學生編號）以「TAC」開頭的紀錄
            col_b = [col for col in df_range.columns if col.strip() == "學生編號"]
            if col_b:
                df_range = df_range[~df_range[col_b[0]].astype(str).str.startswith("TAC", na=False)]
            # 3. 刪除欄位「課室」包含「live」字樣（忽略大小寫）的紀錄
            col_room = [col for col in df_range.columns if "課室" in col]
            if col_room:
                df_range = df_range[~df_range[col_room[0]].astype(str).str.contains("live", case=False, na=False)]
            st.session_state['book_deleted_df'] = df_range  # 暫存刪除後的結果
            st.success(f"已完成刪除，剩餘 {len(df_range)} 筆（原始 {original_count} 筆）")
            st.dataframe(df_range)
            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                return output.getvalue()
            st.download_button(
                label="下載已刪除資料 Excel",
                data=to_excel(df_range),
                file_name="deleted_rows_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # 排序和刪除重覆
    elif book_page == "排序和刪除重覆":
        st.header("書數預算 - 排序和刪除重覆")
        if 'book_deleted_df' not in st.session_state:
            st.warning("請先完成『刪除步驟』。")
        else:
            df = st.session_state['book_deleted_df'].copy()
            # 先排序
            sort_cols = []
            # 找出正確欄名
            col_id = [col for col in df.columns if col.strip() == "學生編號"]
            col_class = [col for col in df.columns if "班別" in col]
            col_status = [col for col in df.columns if "老師出席狀況排序" in col]
            if not col_id or not col_class or not col_status:
                st.error("找不到排序所需欄位，請檢查資料。")
            else:
                sort_cols = [col_id[0], col_class[0], col_status[0]]
                df = df.sort_values(by=sort_cols, ascending=True, ignore_index=True)
                # 刪除重覆
                df = df.drop_duplicates(subset=[col_id[0], col_class[0]], keep='first', ignore_index=True)
                st.success(f"排序並刪除重覆後，剩餘 {len(df)} 筆。")
                st.dataframe(df)
                def to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False)
                    return output.getvalue()
                st.download_button(
                    label="下載排序及去重後 Excel",
                    data=to_excel(df),
                    file_name="sorted_deduped.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# 做卷資料子選單（原本的功能不變，以下略）
elif main_page == "做卷資料":
    st.sidebar.markdown("---")
    juan_page = st.sidebar.radio(
        "做卷資料功能",
        ["做卷有效資料", "出卷老師資料", "分校做卷情況"]
    )

    # ...（其餘做卷資料功能程式碼不變，略）
