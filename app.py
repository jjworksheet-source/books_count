# ...前略...

st.success(f"有效資料共 {len(df_valid)} 筆，重複資料共 {len(df_duplicates)} 筆。")

# 只顯示 A~O 欄（index 0~14），並新增一個空的 P 欄
valid_cols = df_valid.columns[:15]  # 取前 15 欄
df_show = df_valid.loc[:, valid_cols].copy()
df_show["P欄"] = ""  # 新增空的 P 欄

st.subheader("有效資料（A~O欄＋P欄）")
st.dataframe(df_show)
st.subheader("重複資料")
st.dataframe(df_duplicates)

# Download buttons
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

st.download_button(
    label="下載有效資料 Excel",
    data=to_excel(df_show),  # 只下載 A~O+P
    file_name="valid_data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
st.download_button(
    label="下載重複資料 Excel",
    data=to_excel(df_duplicates),
    file_name="duplicate_data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Save valid_data to session_state for step 2
st.session_state['valid_data'] = df_valid
