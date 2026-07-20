# Это точка входа для PyOxidizer
if __name__ == "__main__":
    main()

def main():
    import streamlit as st
    
    # ВАЖНО: Добавляем патч для запуска внутри .exe (мы уже делали это раньше)
    import sys
    if getattr(sys, 'frozen', False):
        os.environ["STREAMLIT_SERVER_ENABLE_WEBSOCKET_COMPRESSION"] = "false"
        from pathlib import Path
        __file__ = str(Path(sys.executable).parent / "__init__.py")

import streamlit as st
# Патч для работы Streamlit внутри .exe от PyInstaller
import sys
if getattr(sys, 'frozen', False):  # Проверка на запуск из .exe
    import os
    os.environ["STREAMLIT_SERVER_ENABLE_WEBSOCKET_COMPRESSION"] = "false"
    from pathlib import Path
    __file__ = str(Path(sys.executable).parent / "__init__.py")
import pandas as pd
from io import BytesIO

# --- Настройка ---
st.set_page_config(page_title="ABC -> Матрица", page_icon="📊", layout="centered")
st.title("📊 Преобразователь ABC-анализа")
st.write("Загрузите ваш файл, чтобы получить матрицу «Артикул — Статья»")

# --- Сессия: будем хранить данные здесь, чтобы они не пропадали ---
if 'df_raw' not in st.session_state:
    st.session_state.df_raw = None
    st.session_state.result_df = None

# --- Блок загрузки ---
uploaded_file = st.file_uploader(
    "Выберите Excel-файл", 
    type=["xlsx", "xls"], 
    help="Ожидаемый формат: АВС_анализ.xlsx с листом 'Данные'"
)

if uploaded_file is not None:
    with st.spinner("Чтение данных..."):
        try:
            df = pd.read_excel(uploaded_file, sheet_name="Данные", header=0)
            
            # Если у вас заголовки находятся ниже 1-й строки (например, есть пустая строка сверху),
            # раскомментируйте следующую строку и укажите правильный номер (обычно 0 или 1):
            # df.columns = df.iloc[0]  # берем названия из первой СТРОКИ таблицы
            # df = df.drop(index=0).reset_index(drop=True) # удаляем эту строку из данных
            
            cols_lower = {str(col).lower().strip(): col for col in df.columns}
            article_col = next((o for k, o in cols_lower.items() if 'артикул' in k and 'поставщик' in k), None)
            stat_col = next((o for k, o in cols_lower.items() if 'стать' in k), None)
            sum_col = next((o for k, o in cols_lower.items() if 'сумм' in k or 'unnamed' in k), None)

            if not all([article_col, stat_col, sum_col]):
                st.error(f"❌ Не найдены колонки.\nНайденные имена: {list(df.columns)}")
                st.stop()

            st.success("✅ Файл прочитан!")
            st.info(f"**Распознаны колонки:**\n- Артикул: `{article_col}`\n- Статья: `{stat_col}`\n- Сумма: `{sum_col}`")
            
            # Сохраняем в сессию
            st.session_state.df_raw = df
            
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
                sum_col = next((o for k, o in cols_lower.items() if 'сумм' in k or 'unnamed' in k), None)

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

                # *** ИСПРАВЛЕНИЕ ЗДЕСЬ ***
                # Вместо .to_excel(файл) используем буфер в памяти
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='Матрица')
                
                # Переходим в начало буфера, иначе скачается пустой файл
                buffer.seek(0)
                st.session_state.result_df = result_df
                
                st.success(f"✅ Готово! Итоговых артикулов: {len(result_df)}")
                st.dataframe(result_df)
                
                # Передаем байты из буфера напрямую в кнопку
                st.download_button(
                    label="💾 Скачать готовый файл",
                    data=buffer.getvalue(),
                    file_name="Матрица_Артикул_Статья.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Подробная ошибка во время Pivot:\n{e}")