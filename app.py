import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="FORESTLOOK: Анализ прибыли", layout="wide")
st.title("📊 FORESTLOOK — Прибыль за неделю")
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
    try:
        # Загрузка WB-отчёта
        wb_sheets = pd.read_excel(wb_file, sheet_name=None)
        if "Товары" not in wb_sheets:
            st.error("❌ В файле WB нет листа 'Товары'")
            st.stop()

        wb_data = wb_sheets["Товары"].iloc[1:].copy()
        wb_data.columns = wb_sheets["Товары"].iloc[0]

        required_columns = [
            "Артикул продавца", "Название", "Средняя цена, ₽",
            "Среднее количество заказов в день, шт",
            "Остатки склад ВБ, шт", "Остатки МП, шт"
        ]
        for col in required_columns:
            if col not in wb_data.columns:
                st.error(f"❌ В отчёте WB нет колонки: {col}")
                st.stop()

        df_wb = wb_data[required_columns].copy()
        df_wb.columns = ["Артикул", "Название", "Средняя цена", "Продаж в день", "Остаток ВБ", "Остаток МП"]
        df_wb["Продаж за неделю"] = (pd.to_numeric(df_wb["Продаж в день"], errors="coerce") * 7).round()

        # Загрузка юнит-экономики
        df_unit = pd.read_excel(unit_file)
        expected_cols = ["Артикул продавца", "Себестоимость", "ROI", "Прибыль с 1 шт"]
        for col in expected_cols:
            if col not in df_unit.columns:
                st.error(f"❌ В юнит-экономике нет колонки: {col}")
                st.stop()

        # Объединение только по тем артикулам, которые есть в WB-отчёте
        df_merged = pd.merge(df_wb, df_unit, how="left", left_on="Артикул", right_on="Артикул продавца")

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

    except Exception as e:
        st.error(f"❌ Ошибка: {str(e)}")
