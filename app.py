import streamlit as st
import pandas as pd
import difflib
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="FORESTLOOK: Анализ прибыли", layout="wide")
st.title("📊 FORESTLOOK — Единый отчёт по прибыли за неделю")
st.markdown("Загрузите два Excel-файла: Wildberries-отчёт и юнит-экономику")

wb_file = st.file_uploader("📤 Отчёт Wildberries (.xlsx)", type="xlsx")
unit_file = st.file_uploader("📤 Юнит-экономика (.xlsx)", type="xlsx")

def classify(row):
    if row["Продаж за неделю"] < 10:
        return "Тест"
    if row["Прибыль с 1 шт"] < 0 or row["ROI"] < 20:
        return "Балласт"
    if row["ROI"] < 40:
        return "Витрина"
    if row["ROI"] >= 60 and row["Прибыль с 1 шт"] >= 300:
        return "Флагман"
    return "Обычный"

if wb_file and unit_file:
    wb_sheets = pd.read_excel(wb_file, sheet_name=None)
    wb_data = wb_sheets['Товары'].iloc[1:].copy()
    wb_data.columns = wb_sheets['Товары'].iloc[0]

    df_wb = wb_data[[
        "Артикул продавца", "Название", "Средняя цена, ₽", "Среднее количество заказов в день, шт",
        "Остатки склад ВБ, шт", "Остатки МП, шт"
    ]].copy()
    df_wb.columns = ["Артикул", "Название", "Средняя цена", "Продаж в день", "Остаток ВБ", "Остаток МП"]
    df_wb["Продаж за неделю"] = (pd.to_numeric(df_wb["Продаж в день"], errors="coerce") * 7).round()

    unit_raw = pd.read_excel(unit_file)
    unit_raw["Название"] = unit_raw.iloc[:, 0]
    unit_raw["Себестоимость"] = pd.to_numeric(unit_raw.iloc[:, 8], errors="coerce")
    unit_raw["ROI"] = pd.to_numeric(unit_raw.iloc[:, 19], errors="coerce")
    unit_raw["Прибыль с 1 шт"] = pd.to_numeric(unit_raw.iloc[:, 28], errors="coerce")
    unit_clean = unit_raw[["Название", "Себестоимость", "ROI", "Прибыль с 1 шт"]].dropna()

    def match_name(name, choices):
        match = difflib.get_close_matches(name, choices, n=1, cutoff=0.4)
        return match[0] if match else None

    df_wb["Название юнит"] = df_wb["Название"].apply(lambda x: match_name(x, unit_clean["Название"]))
    df_merged = pd.merge(df_wb, unit_clean, left_on="Название юнит", right_on="Название", how="left")

    df_merged["Чистая прибыль за неделю"] = (df_merged["Продаж за неделю"] * df_merged["Прибыль с 1 шт"]).round(2)
    df_merged["Статус"] = df_merged.apply(classify, axis=1)

    st.success("✅ Отчёт готов")
    st.dataframe(df_merged, use_container_width=True)

    def convert_df(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Отчёт FORESTLOOK")
        return output.getvalue()

    excel_data = convert_df(df_merged)
    filename = f"FORESTLOOK_отчет_{datetime.today().strftime('%Y-%m-%d')}.xlsx"

    st.download_button(
        label="📥 Скачать Excel-отчёт",
        data=excel_data,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
