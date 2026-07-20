import pandas as pd
import streamlit as st
from io import BytesIO
import numpy as np

# Настройка внешнего вида приложения
st.set_page_config(
    page_title="Анализ ассортимента",
    layout="centered"
)
st.title("📊 Анализ ассортимента")
st.write("Загрузите файл Excel с данными о продажах.")

uploaded_file = st.file_uploader(
    "Выберите Excel-файл", type=["xlsx"],
    help="Ожидаемый формат: АВС_анализ.xlsx с листом 'Данные'"
)

if uploaded_file is not None:
    with st.spinner("Чтение данных..."):
        df = pd.read_excel(uploaded_file, sheet_name="Данные", header=0)
    
    # Проверяем наличие пустой строки перед заголовками
    if len(df.columns) == 1 and isinstance(df.iloc[0, 0], str):  
        df.columns = df.iloc[0]
        df = df.drop(index=0).reset_index(drop=True)

    cols_lower = {str(col).lower(): col for col in df.columns}

    article_col = next((o for k, o in cols_lower.items() if ('артикул' in k or 'код товара' in k)), None)
    stat_col = cols_lower.get('статья')

    possible_sums = ["сумма", "итог"] 
    sum_col = next(  # Исправленный поиск суммы!
        (original for key, original in cols_lower.items() 
         if any(x in key for x in possible_sums) or ("unnamed" in key)),
        None
    )

    required_articles_for_pivot = [article_col, stat_col]
    missing_articles = [art for art in required_articles_for_pivot if not isinstance(art, str)]
    if len(missing_articles) > 0:
        st.error(f"❌ ОШИБКА: Не найдено всех необходимых колонок!")
        st.stop()

    # Очистка данных
    df[stat_col] = df[stat_col].astype(str).str.strip()
    df.dropna(subset=[article_col, stat_col, sum_col], inplace=True)

    #### ПОСТРОЕНИЕ СВЁДНОЙ ТАБЛИЦЫ ####
    pivot_table = df.pivot_table(
        index=article_col,
        columns=stat_col,
        values=sum_col,
        aggfunc='sum',
        fill_value=0
    )

    result_df = pivot_table.reset_index()

    #### ❗ ЗДЕСЬ ДОЛЖНА БЫТЬ ПРОВЕРКА СТАТЕЙ ###
    # Список обязательных статей для расчёта метрик.
    # Мы проверяем их наличие только ПОСЛЕ ТОГО, КАК ПОСТРОИЛИ СВЕДНУЮ ТАБЛИЦУ.
    required_articles = [
        *['Логистика СДЭК', 'Логистика общая, руб', 'Плановая комиссия, руб', 'Себестоимость итого, руб'],
        'Выручка без СПП итого, руб'
    ]

    missing_articles = [art for art in required_articles if art not in result_df.columns]
    if len(missing_articles) > 0:
        st.error(f"❌ Следующие необходимые статьи отсутствуют в файле:\n- {'\n- '.join(missing_articles)}")
        st.stop()

    #### РАСЧЁТ МЕТРИК ДО ABC-АНАЛИЗА ####
    df_for_calculation = result_df[required_articles[:-1]].fillna(0)  # Все расходы
    result_df['Итого переменные расходы,руб'] = (
        df_for_calculation.sum(axis=1)
    )

    #### ПУНКТ 3: МАРЖИНАЛЬНАЯ ПРИБЫЛЬ И МАРЖА (%) ###
    result_df['Маржинальная прибыль, руб'] = (
        result_df['Выручка без СПП итого, руб'] -
        result_df['Итого переменные расходы,руб']
    )

    df_for_margin = result_df[[
        'Маржинальная прибыль, руб',
        'Выручка без СПП итого, руб'
    ]].fillna(0)

    result_df['Маржа, %'] = (
        df_for_margin['Маржинальная прибыль, руб'] /
        df_for_margin['Выручка без СПП итого, руб']
    ).replace([np.inf, -np.inf], np.nan).fillna(0)
    result_df['Маржа, %'] = np.round(result_df['Маржа, %'].astype(float) * 100, 2)

    #### СОЗДАНИЕ ОСНОВНОЙ ИТОГОВОЙ МАТРИЦЫ ####
    columns_order = [
        article_col,
        'Выручка без СПП итого, руб',
        'Итого переменные расходы,руб',
        'Маржинальная прибыль, руб',
        'Маржа, %'
    ]

    existing_columns = [col for col in columns_order if col in result_df.columns]

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

    #### ОБЩИЙ ЦИКЛ АНАЛИЗА ####

    def abc_analysis(dataframe, criterion_column):
        """Строит ABC-анализ по заданному критерию."""
        abc_df = dataframe[[article_col, criterion_column]].copy()
        
        abc_df.sort_values(by=criterion_column, ascending=False, inplace=True)
        abc_df['Cumulative_Sum'] = abc_df[criterion_column].cumsum()
        total_sum = abc_df[criterion_column].sum()

        abc_df['Percentage_of_Total'] = (abc_df['Cumulative_Sum'] / total_sum) * 100

        abc_df['ABC_Category'] = pd.cut(
            abc_df['Percentage_of_Total'], bins=[0, 80, 95, float('inf')],
            labels=['A', 'B', 'C'],
            right=False
        )

        return abc_df

    criteria = [
        ("Выручка", "Выручка без СПП итого, руб"),
        ("Прибыль", "Маржинальная прибыль, руб"),
        ("Маржа", "Маржа, %")
    ]

    for name, column_name in criteria:
        if column_name in result_df.columns:
            abc_df = abc_analysis(result_df, column_name)
            
            final_columns = [article_col, column_name, 'ABC_Category']

            top_bottom = pd.concat([
                abc_df.head(5),
                abc_df.tail(5)[::-1]     # Перевёрнутый список последних
            ])

            st.subheader(f"🔹 ABC-анализ по {name}")
            st.dataframe(top_bottom, use_container_width=True)

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