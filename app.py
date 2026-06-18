import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="庫存週報工具v2", layout="wide")

st.title("庫存週報工具")
st.caption("上傳上週與本週庫存表，自動計算庫存差異與推算出貨量。")

st.warning(
    "提醒：推算出貨量是用庫存差異估算。若期間有進貨、退貨、盤點調整，數字會受影響。"
)

col1, col2 = st.columns(2)

with col1:
    last_week_file = st.file_uploader("上傳上週庫存表 Excel", type=["xlsx"], key="last_week")

with col2:
    this_week_file = st.file_uploader("上傳本週庫存表 Excel", type=["xlsx"], key="this_week")


def read_excel(file):
    return pd.read_excel(file)


def to_excel_file(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="每週庫存比較")
    return output.getvalue()


if last_week_file and this_week_file:
    last_df = read_excel(last_week_file)
    this_df = read_excel(this_week_file)

    st.subheader("欄位設定")

    all_columns = list(this_df.columns)

    sku_col = st.selectbox("選擇 SKU / 商品編號欄位", all_columns)
    name_col = st.selectbox("選擇商品名稱欄位", all_columns)
    stock_col = st.selectbox("選擇庫存欄位", all_columns)

    if st.button("開始分析"):
        last = last_df[[sku_col, name_col, stock_col]].copy()
        this = this_df[[sku_col, name_col, stock_col]].copy()

        last = last.rename(columns={
            sku_col: "SKU",
            name_col: "商品名稱_上週",
            stock_col: "上週庫存"
        })

        this = this.rename(columns={
            sku_col: "SKU",
            name_col: "商品名稱_本週",
            stock_col: "本週庫存"
        })

        result = pd.merge(last, this, on="SKU", how="outer")

        result["商品名稱"] = result["商品名稱_本週"].fillna(result["商品名稱_上週"])
        result["上週庫存"] = pd.to_numeric(result["上週庫存"], errors="coerce").fillna(0)
        result["本週庫存"] = pd.to_numeric(result["本週庫存"], errors="coerce").fillna(0)

        result["庫存差異"] = result["本週庫存"] - result["上週庫存"]
        result["推算出貨量"] = result["庫存差異"].apply(lambda x: abs(x) if x < 0 else 0)

        def status(row):
            if row["上週庫存"] == 0 and row["本週庫存"] > 0:
                return "新增商品/進貨"
            elif row["上週庫存"] > 0 and row["本週庫存"] == 0:
                return "售完/需確認"
            elif row["庫存差異"] < 0:
                return "正常出貨"
            elif row["庫存差異"] > 0:
                return "可能進貨/調整"
            else:
                return "無變動"

        result["狀態"] = result.apply(status, axis=1)

        final_df = result[[
            "SKU",
            "商品名稱",
            "上週庫存",
            "本週庫存",
            "庫存差異",
            "推算出貨量",
            "狀態"
        ]].sort_values(by="推算出貨量", ascending=False)

        st.subheader("分析結果")
        st.dataframe(final_df, use_container_width=True)

        st.subheader("本週推算出貨量 Top 20")
        st.dataframe(final_df.head(20), use_container_width=True)

        excel_data = to_excel_file(final_df)

        st.download_button(
            label="下載庫存週報 Excel",
            data=excel_data,
            file_name="庫存週報.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("請先上傳上週與本週庫存表。")
