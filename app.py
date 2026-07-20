import pandas as pd
import streamlit as st
from io import BytesIO
import numpy as np

# --- Настройка внешнего вида ---
st.set_page_config(
    page_title="Анализ ассортимента",
    layout="centered"
)
st.title("📊 Анализ ассортимента")
st.write("Загрузите файл Excel с данными о продажах.")

uploaded_file = st.file_uploader(
    "Выберите Excel-файл", 
    type=["xlsx"],
    help="Ожидаемый формат: АВС_анализ.xlsx с листом 'Данные'"
)

if uploaded_file is not None:
    with st.spinner("Чтение данных..."):
        df = pd.read_excel(uploaded_file, sheet_name="Данные", header=0)
    
    # Проверяем наличие пустой строки перед заголовками
    if len(df.columns) == 1 and isinstance(df.iloc[0, 0], str):  
        # Если первая строка содержит строку, а не число, считаем это шапкой
        df.columns = df.iloc[0]
        df = df.drop(index=0).reset_index(drop=True)

    cols_lower = {str(col).lower(): col for col in df.columns}

    #### ⚠️ НАДЕЖНЫЙ ПОИСК КОЛОНОК ###
    article_col = next((o for k, o in cols_lower.items() if ('артикул' in k or 'код товара' in k)), None)
    stat_col = cols_lower.get('статья')

    possible_sums = ["сумма", "итог"] 
    sum_col = next(  # <--- ИСПРАВЛЕННЫЙ ПОИСК СУММЫ!
        (original for key, original in cols_lower.items() 
         if any(x in key for x in possible_sums) or ("unnamed" in key)),
        None
    )

    article_col = st.selectbox(
        "🧩 Колонка с артикулами:",
        options=list(cols_lower.values()),
        placeholder="Выберите..."
    )

    stat_col = st.selectbox(
        "🧩 Колонка со списком статей:",
        options=list(cols_lower.values()),
        placeholder="Выберите..."
    )

    sum_col = st.selectbox(
        "🧩 Колонка с числовыми значениями:",
        options=list(cols_lower.values()),
        placeholder="Выберите..."
    )

    if all([article_col, stat_col, sum_col]):
        # Очистка данных
        df[stat_col] = df[stat_col].astype(str).str.strip()
        df.dropna(subset=[article_col, stat_col, sum_col], inplace=True)

        # Разворот таблицы (Pivot) - ГЛАВНОЕ ДЕЙСТВИЕ
        pivot_table = df.pivot_table(
            index=article_col,
            columns=stat_col,
            values=sum_col,
            aggfunc='sum',
            fill_value=0
        )

        result_df = pivot_table.reset_index()

        #### ОТЛАДКА И ВЫБОР СТАТЕЙ 👇👇👇 ####
        # Показываем пользователю, какие статьи попали в итоговую таблицу
        st.subheader("Найденные финансовые метрики:")
        selected_articles = [col for col in result_df.columns if col != article_col]
        
        # Теперь пользователь может выбрать статьи для расчётов
        required_articles_for_calculations = [
            st.multiselect(
                "🧩 Статьи для переменных расходов:",
                options=selected_articles,
                default=['Логистика СДЭК', 'Логистика общая, руб', 'Плановая комиссия, руб', 'Себестоимость итого, руб'],
                help="Выберите все расходы."
            ),
            
            st.selectbox(
                "🧩 Статья для базы расчётов (Продажи/Выручка):",
                options=selected_articles,
                index=next((i for i, art in enumerate(selected_articles) if 'выручк' in art.lower()), 0),
                help="Эта статья должна быть в рублях.",
                key="base_sales"
            ),
            
            st.selectbox(
                "🧩 Статья для маржинальной прибыли:",
                options=selected_articles,
                index=next((i for i, art in enumerate(selected_articles) if 'маржинал' in art.lower()), 0),
                key="profit_column"
            ),
            
            st.selectbox(
                "🧩 Статья для маржи (%):",
                options=selected_articles,
                index=next((i for i, art in enumerate(selected_articles) if '%' in art), 0),
                key="margin_column"
            )
        ]

        # Проверка наличия выбранных статей
        missing_articles = [art for art in required_articles_for_calculations[:4] if art not in result_df.columns]
        if len(missing_articles) > 0:
            st.error(f"❌ Следующие необходимые статьи отсутствуют в файле:\n- {'\n- '.join(missing_articles)}")
        elif not all(required_articles_for_calculations):
            st.warning("Пожалуйста, выберите все обязательные статьи выше.")
        else:
            #### 🎯 РАСЧЁТ МЕТРИК ДО ABC-АНАЛИЗА ####

            # Переменные расходы
            df_for_calculation = result_df[required_articles_for_calculations[0]].fillna(0)
            result_df['Итого переменные расходы,руб'] = (
                df_for_calculation.sum(axis=1)
            )

            #### 📝 ПУНКТ 3: МАРЖИНАЛЬНАЯ ПРИБЫЛЬ И МАРЖА (%) ####

            # Для расчёта используем единую базу - реальную статью пользователя.
            base_sales = required_articles_for_calculations[1]
            profit_column = required_articles_for_calculations[2]
            margin_column = required_articles_for_calculations[3]

            # Расчёт маржинальной прибыли
            # Прибыли = База Продаж (Выручка) - Переменные расходы
            result_df['Маржинальная прибыль, руб'] = (
                result_df[base_sales] -
                result_df['Итого переменные расходы,руб']
            )

            # Расчёт маржи %
            # Маржа = Прибыль / Базу Продаж (ту же самую!) * 100%
            result_df['Маржа, %'] = (
                result_df['Маржинальная прибыль, руб'] /
                result_df[margin_column]
            ).replace([np.inf, -np.inf], np.nan).fillna(0)
            result_df['Маржа, %'] = np.round(result_df['Маржа, %'].astype(float) * 100, 2)

            #### ✍️ СОЗДАНИЕ ОСНОВНОЙ ИТОГОВОЙ МАТРИЦЫ ####

            # Устанавливаем фиксированный порядок колонок
            # Используем выбор пользователя как основу для расчётов.
            columns_order = [
                article_col,
                base_sales,                     # Реальная база продаж
                'Итого переменные расходы,руб',
                'Маржинальная прибыль, руб',
                'Маржа, %'
            ]

            existing_columns = [col for col in columns_order if col in result_df.columns]

            # Сохранение основной матрицы в буфер для скачивания
            buffer_main = BytesIO()
            with pd.ExcelWriter(buffer_main, engine='openpyxl') as writer:
                result_df[existing_columns].to_excel(writer, index=False, sheet_name='Матрица')
            buffer_main.seek(0)

            st.success("✅ Основная матрица создана!")
            st.download_button(
                label="💾 Скачать основную матрицу",
                data=buffer_main.getvalue(),
                file_name=f"Матрица_{article_col}_Итоговая.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            #### ❗ ОБЩИЙ БЛОК ПОСТРОЕНИЯ ABC-АНАЛИЗА ####

            def abc_analysis(dataframe, criterion_column):
                """
                Строит ABC-анализ по заданному критерию и сохраняет результат в файл.
                Возвращает DataFrame с категориями.
                """
                abc_df = dataframe[[article_col, criterion_column]].copy()
                
                # Сортируем товары от самого большого значения к самому маленькому
                abc_df.sort_values(by=criterion_column, ascending=False, inplace=True)

                # Считаем кумулятивную сумму (нарастающий итог)
                abc_df['Cumulative_Sum'] = abc_df[criterion_column].cumsum()

                # Находим общую сумму всех значений нашего критерия
                total_sum = abc_df[criterion_column].sum()

                # Рассчитываем долю каждого товара в общем результате (%)
                abc_df['Percentage_of_Total'] = (abc_df['Cumulative_Sum'] / total_sum) * 100

                # Определяем категории ABC
                abc_df['ABC_Category'] = pd.cut(
                    abc_df['Percentage_of_Total'], 
                    bins=[0, 80, 95, float('inf')],
                    labels=['A', 'B', 'C'],
                    right=False
                )

                return abc_df

            #### ✍️ ОСНОВНОЙ ЦИКЛ АНАЛИЗА ####

            # Список критериев для анализа
            criteria = [
                ("Выручка", base_sales),
                ("Прибыль", "Маржинальная прибыль, руб"),
                ("Маржа", "Маржа, %")
            ]

            # Запускаем анализ для каждого критерия
            for name, column_name in criteria:
                if column_name in result_df.columns:
                    abc_df = abc_analysis(result_df, column_name)
                    
                    # Оставляем только нужные колонки
                    final_columns = [article_col, column_name, 'ABC_Category']

                    # Сохранение результата в буфер
                    buffer_abc = BytesIO()
                    with pd.ExcelWriter(buffer_abc, engine='openpyxl') as writer:
                        abc_df[final_columns].to_excel(writer, index=False, sheet_name=name)
                    buffer_abc.seek(0)

                    st.download_button(
                        label=f"💾 Скачать ABC-{name}",
                        data=buffer_abc.getvalue(),
                        file_name=f"Mатрица_{article_col}_ABC_{name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
