import pandas as pd
import streamlit as st
from io import BytesIO
import os
# Пример ссылки на картинку из вашего же репозитория
st.image(
    "https://raw.https://github.com/Stepanova2200/abc-analysis/blob/main/logo.png",
    caption="Логотип компании", width=200
)

# --- Настройка внешнего вида ---
st.set_page_config(
    page_title="Преобразователь ABC-анализа",
    page_icon="📊",
    layout="centered"
)
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
            
            # Если у вас заголовки находятся ниже 1-й строки (например, есть пустая строка сверху),
            # раскомментируйте следующие две строчки и укажите правильный номер (обычно 0 или 1):
            # df.columns = df.iloc[0]  # берем названия из первой СТРОКИ таблицы
            # df = df.drop(index=0).reset_index(drop=True) # удаляем эту строку из данных
            
            cols_lower = {str(col).lower().strip(): col for col in df_raw.columns}
            article_col = next((o for k, o in cols_lower.items() if 'артикул' in k and 'поставщик' in k), None)
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
    
    if st.session_state.df_raw is None:
        st.warning("Сначала загрузите файл выше.")
    else:
        with st.spinner("Выполняется поворот таблицы..."):
            try:
                df = st.session_state.df_raw.copy()
                
                # Определяем колонки заново внутри блока кнопки (защита от сброса сессии)
                cols_lower = {str(col).lower().strip(): col for col in df.columns}
                article_col = next((o for k, o in cols_lower.items() if 'артикул' in k and 'поставщик' in k), None)
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
                result_df = pivot_table.reset_index()

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
