import pandas as pd
import streamlit as st
from io import BytesIO

# --- Настройка внешнего вида ---
st.set_page_config(
    page_title="Преобразователь ABC-анализа",
    # ⚠️ Замените эмодзи на URL вашей картинки из GitHub или оставьте как есть
    page_icon="📊", 
    layout="centered"
)

# Пример ссылки на картинку из вашего репозитория
# st.image("https://raw.githubusercontent.com/ваш-аккаунт/abc-analysis/main/logo.png", width=200)

st.title("📊 Преобразователь ABC-анализа")
st.write("Загрузите ваш файл, чтобы получить матрицу «Артикул — Статья»")

uploaded_file = st.file_uploader(
    "Выберите Excel-файл", 
    type=["xlsx"],
    help="Ожидаемый формат: АВС_анализ.xlsx с листом 'Данные'"
)

if uploaded_file is not None:
    with st.spinner("Чтение данных..."):
        try:
            df_raw = pd.read_excel(uploaded_file, sheet_name="Данные", header=0)
            
            # Проверяем наличие пустой строки перед заголовками
            if len(df_raw.columns) == 1 and isinstance(df_raw.iloc[0, 0], str):  
                # Если первая строка содержит строку, а не число, считаем это шапкой
                df_raw.columns = df_raw.iloc[0]
                df_raw = df_raw.drop(index=0).reset_index(drop=True)

            cols_lower = {str(col).lower().strip(): col for col in df_raw.columns}
            article_col = next((o for k, o in cols_lower.items() if 'артикул' in k and ('поставщик' in k or 'код' in k)), None)
            stat_col = next((o for k, o in cols_lower.items() if 'стать' in k), None)
            sum_col = next((o for k, o in cols_lower.items() if ('сумм' in k or 'unnamed' in k)), None)

            if not all([article_col, stat_col, sum_col]):
                st.error(f"❌ Не найдены колонки.\nНайденные имена: {list(df_raw.columns)}")
                st.stop()

            st.success("✅ Файл прочитан!")
            st.info(f"**Распознаны колонки:**\n- Артикул: `{article_col}`\n- Статья: `{stat_col}`\n- Сумма: `{sum_col}`")
            
            # Сохраняем в сессию
            st.session_state.df_raw = df_raw
            
        except Exception as e:
            st.error(f"❌ Ошибка при чтении файла: {e}")
            st.stop()
else:
    # Если файл еще не загружен, очищаем старые результаты
    st.session_state.result_df = None


# --- Кнопка запуска расчета ---
if st.button("⚡ Сформировать матрицу"):
    
    if st.session_point.state.df_raw is None:
        st.warning("Сначала загрузите файл выше.")
    else:
        with st.spinner("Выполняется поворот таблицы..."):
            try:
                df = st.session_state.df_raw.copy()
                
                # Определяем колонки заново внутри блока кнопки (защита от сброса сессии)
                cols_lower = {str(col).lower().strip(): col for col in df.columns}
                article_col = next((o for k, o in cols_lower.items() if 'артикул' in k and ('поставщик' in k or 'код' in k)), None)
                stat_col = next((o for k, o in cols_lower.items() if 'стать' in k), None)
                sum_col = next((o for k, o in cols_lower.items() if ('сумм' in k or 'unnamed' in k)), None)

                df[stat_col] = df[stat_col].astype(str).str.strip()
                df[sum_col] = pd.to_numeric(df[sum_col], errors='coerce')
                
                before_drop = len(df)
                df_clean = df.dropna(subset=[article_col, stat_col])
                df_clean = df_clean.dropna(subset=[sum_col])
                
                pivot_table = df_clean.pivot_table(
                    index=article_col,
                    columns=stat_col,
                    values=sum_col,
                    aggfunc='sum',
                    fill_value=0
                )
                pivot_table.columns.name = None

                ### ВАШИ НОВЫЕ РАСЧЁТЫ ###
                # ✅ Пункт 1: Переименование колонки
                # Мы работаем с MultiIndex
                if 'Кол-во итого, шт.' in pivot_table.columns.get_level_values(0):
                    new_columns = [col if col != "Кол-во итого, шт." else ("Продажа итого, шт",)
                                   for col in pivot_table.columns]
                    
                    # Устанавливаем новый индекс
                    pivot_table.columns = pd.MultiIndex.from_tuples(new_columns)
                
                result_df = pivot_table.reset_index() # Только теперь сбрасываем index!

                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='Матрица')
                buffer.seek(0)
                st.session_state.result_df = result_df
                
                st.success(f"✅ Готово! Итоговых артикулов: {len(result_df)}")
                st.dataframe(result_df.head())
                
                # Передаем байты из буфера напрямую в кнопку
                st.download_button(
                    label="💾 Скачать готовый файл",
                    data=buffer.getvalue(),
                    file_name="Матрица_Артикул_Статья.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Подробная ошибка во время Pivot:\n{e}")
