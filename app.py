import pandas as pd
import streamlit as st
from io import BytesIO
import numpy as np
import altair as alt  # Для построения диаграммы Парето

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

    article_col = next((o for k, o in cols_lower.items() if ('артикул' in k or 'код товара' in k)), None)
    stat_col = cols_lower.get('статья')
    sum_col = next((col for key, original in cols_lower.items() if any(x in key.lower() for x in ["сумма", "итог"])), None)

    # ⚠️ В ИНТЕРАКТИВНОМ РЕЖИМЕ пользователь выбирает эти колонки сам!
    # Мы даём ему список всех доступных вариантов.
    article_col = st.selectbox("Колонка с артикулами:", options=list(cols_lower.values()), index=None, format_func=lambda x: f"{x} ({cols_lower[x].title()})" if x else "")
    stat_col = st.selectbox("Колонка со списком статей:", options=list(cols_lower.values()), index=None, format_func=lambda x: f"{x} ({cols_lower[x].title()})" if x else "")
    sum_col = st.selectbox("Колонка с числовыми значениями:", options=list(cols_lower.values()), index=None, format_func=lambda x: f"{x} ({cols_lower[x].title()})" if x else "")

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

        #### ОТЛАДКА И ВЫБОР СТАТЕЙ ###
        # Показываем пользователю, какие статьи попали в итоговую таблицу
        st.subheader("Найденные финансовые метрики:")
        selected_articles = [col for col in result_df.columns if col != article_col]
        st.dataframe(selected_articles, use_container_width=True)
        
        # Теперь пользователь может выбрать статьи для расчётов
        required_articles_for_calculations = [
            st.multiselect("Статьи для переменных расходов:",
                          options=selected_articles,
                          default=['Логистика СДЭК', 'Логистика общая, руб', 'Плановая комиссия, руб', 'Себестоимость итого, руб'],
                          help="Выберите все расходы"),
            
            st.selectbox("Статья для базы расчётов (Продажи/Выручка):",
                         options=selected_articles,
                         index=next((i for i, art in enumerate(selected_articles) if 'выручк' in art.lower()), 0),
                         help="Эта статья должна быть в рублях."),
            
            st.selectbox("Статья для маржинальной прибыли:",
                         options=selected_articles,
                         index=next((i for i, art in enumerate(selected_articles) if 'маржинал' in art.lower()), 0)),
            
            st.selectbox("Статья для маржи (%):",
                         options=selected_articles,
                         index=next((i for i, art in enumerate(selected_articles) if '%' in art), 0))
        ]

        # Проверка наличия выбранных статей
        missing_articles = [art for art in required_articles_for_calculations[:4] if art not in result_df.columns]
        if len(missing_articles) > 0:
            st.error(f"❌ Следующие необходимые статьи отсутствуют в файле:\n- {'\n- '.join(missing_articles)}")
        elif not all(required_articles_for_calculations):
            st.warning("Пожалуйста, выберите все обязательные статьи выше.")
        else:
            #### 📝 РАСЧЁТ МЕТРИК ДО ABC-АНАЛИЗА ####

            # Переменные расходы
            df_for_calculation = result_df[required_articles_for_calculations[0]].fillna(0)
            result_df['Итого переменные расходы,руб'] = df_for_calculation.sum(axis=1)

            # Маржинальная прибыль
            base_sales = required_articles_for_calculations[1]
            result_df['Маржинальная прибыль, руб'] = (
                result_df[base_sales] -
                result_df['Итого переменные расходы,руб']
            )

            # Маржа %
            margin_column = required_articles_for_calculations[2]
            profit_column = required_articles_for_calculations[3]
            result_df['Маржа, %'] = (
                result_df[profit_column] / result_df[margin_column]
            ).replace([np.inf, -np.inf], np.nan).fillna(0)
            result_df['Маржа, %'] = np.round(result_df['Маржа, %'].astype(float) * 100, 2)

            #### ✍️ СОЗДАНИЕ ОСНОВНОЙ ИТОГОВОЙ МАТРИЦЫ ####

            # Устанавливаем фиксированный порядок колонок
            columns_order = [
                article_col,
                base_sales,
                'Итого переменные расходы,руб',
                'Маржинальная прибыль, руб',
                'Маржа, %'
            ]

            existing_columns = [col for col in columns_order if col in result_df.columns]

            # Сохранение основной матрицы в буфер для скачивания
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                result_df[existing_columns].to_excel(writer, index=False, sheet_name='Матрица')
            buffer.seek(0)

            st.success("✅ Основная матрица создана!")
            st.download_button(
                label="💾 Скачать основную матрицу",
                data=buffer.getvalue(),
                file_name=f"Матрица_{article_col}_Итоговая.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            #### ❗ ПОСТРОЕНИЕ ABC-АНАЛИЗА С ДИАГРАММОЙ ПАРЕТО ####

            def abc_analysis(dataframe, criterion_column):
                """
                Строит ABC-анализ по заданному критерию.
                Возвращает DataFrame с категориями и Altair-чарт.
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

                # Создаём комбинированную диаграмму (Altair вместо Excel)
                chart = (
                    alt.Chart(abc_df.reset_index())
                    .transform_window(
                        rank='rank()',
                        sort=[alt.SortField(criterion_column, order='descending')]
                    )
                    .encode(
                        x=alt.X('rank:Q', title='Артикулы (по убыванию)'),
                        y=alt.Y(criterion_column, title=criterion_column, axis=alt.Axis(format=',.0f')),
                        color=alt.Color('ABC_Category:N', legend=None)
                    )
                    .mark_bar()
                )

                line_chart = (
                    alt.Chart(abc_df.reset_index())
                    .encode(
                        x=alt.X('rank:Q'),
                        y=alt.Y('Percentage_of_Total:Q', title='% от общего результата', scale=alt.Scale(domain=(0, 100))),
                        strokeDash=alt.value([5, 5])
                    )
                    .mark_line(color='#FF7F0E')
                )

                final_chart = chart + line_chart

                return abc_df, final_chart

            # Список критериев для анализа
            criteria = [
                ("Выручка без СПП", base_sales),
                ("Прибыль", "Маржинальная прибыль, руб"),
                ("Маржа", "Маржа, %")
            ]

            # Запускаем анализ для каждого критерия
            for name, column_name in criteria:
                if column_name in result_df.columns:
                    abc_df, chart = abc_analysis(result_df, column_name)
                    
                    st.subheader(f"🔹 ABC-анализ по {name}")
                    st.altair_chart(chart, use_container_width=True)

                    # Добавляем кнопку скачивания для каждого ABC-отчёта
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        abc_df.to_excel(writer, index=False, sheet_name=name)
                    buffer.seek(0)

                    st.download_button(
                        label=f"💾 Скачать ABC-{name}",
                        data=buffer.getvalue(),
                        file_name=f"Mатрица_{article_col}_ABC_{name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
