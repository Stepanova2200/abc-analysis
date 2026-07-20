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
        df.columns = df.iloc[0]
        df = df.drop(index=0).reset_index(drop=True)

    cols_lower = {str(col).lower(): col for col in df.columns}

    article_col = next((o for k, o in cols_lower.items() if ('артикул' in k or 'код товара' in k)), None)
    stat_col = cols_lower.get('статья')

    possible_sums = ["сумма", "итог"] 
    sum_col = next(  # ИСПРАВЛЕННЫЙ ПОИСК СУММЫ!
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

    #### ❗ ОБЩИЙ БЛОК ПОСТРОЕНИЯ ABC-АНАЛИЗА ####

    def abc_analysis(dataframe, criterion_column):
        """
        Строит ABC-анализ по заданному критерию и возвращает DataFrame с категориями.
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

    #### 🎯 РАСЧЁТ МЕТРИК ДО ABC-АНАЛИЗА ####

    # Список необходимых статей для расчёта расходов
    required_articles = [
        'Логистика СДЭК',
        'Логистика общая, руб',
        'Плановая комиссия, руб',
        'Себестоимость итого, руб',
        # ВАША РЕАЛЬНАЯ СТАТЬЯ ПРОДАЖ ИЗ ОТЛАДОЧНОГО ВЫВОДА
        'Выручка без СПП итого, руб',       # <--- Используем её для базы расчётов
    ]

    # Проверяем наличие этих статей в результирующем DataFrame
    missing_articles = [art for art in required_articles if art not in result_df.columns]
    if len(missing_articles) > 0:
        st.error(f"❌ Следующие необходимые статьи отсутствуют в файле:\n- {'\n- '.join(missing_articles)}")
        st.stop()

    # Создаём временную таблицу только с нужными колонками и заменяем NaN на ноль.
    df_for_calculation = result_df[required_articles].fillna(0)

    # Добавляем новый столбец - Итого переменные расходы
    result_df['Итого переменные расходы,руб'] = (
        df_for_calculation['Логистика СДЭК'] +
        df_for_calculation['Логистика общая, руб'] +
        df_for_calculation['Плановая комиссия, руб'] +
        df_for_calculation['Себестоимость итого, руб']
    )

    #### 📝 ПУНКТ 3: МАРЖИНАЛЬНАЯ ПРИБЫЛЬ И МАРЖА (%) ###

    # Для расчёта используем единую базу - Вашу реальную статью "Выручка без СПП".
    # Все расчёты должны быть в одной валюте (в рублях).

    # Расчёт маржинальной прибыли
    # Прибыли = Выручка (без СПП) - Переменные расходы
    result_df['Маржинальная прибыль, руб'] = (
        df_for_calculation['Выручка без СПП итого, руб'] -
        result_df['Итого переменные расходы,руб']
    )

    # Расчёт маржи %
    # Маржа = Прибыль / Выручку (ту же самую!) * 100%
    # При делении на ноль или очень маленькое число может возникнуть бесконечность (+inf/-inf),
    # которую мы заменим на NaN, а затем на 0.
    df_for_margin = result_df[[
        'Маржинальная прибыль, руб',
        'Выручка без СПП итого, руб'         # Та же база, что и выше
    ]].fillna(0)

    result_df['Маржа, %'] = (
        df_for_margin['Маржинальная прибыль, руб'] /
        df_for_margin['Выручка без СПП итого, руб']
    ).replace([np.inf, -np.inf], np.nan).fillna(0)

    # Приводим проценты к удобному виду (умножаем на 100 и округляем)
    result_df['Маржа, %'] = np.round(result_df['Маржа, %'].astype(float) * 100, 2)

    #### ✍️ СОЗДАНИЕ ОСНОВНОЙ ИТОГОВОЙ МАТРИЦЫ ####

    # Устанавливаем фиксированный порядок колонок
    columns_order = [
        article_col,
        'Выручка без СПП итого, руб',           # Ваша реальная база продаж
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

    #### ❗ ОСНОВНОЙ ЦИКЛ АНАЛИЗА ####

    # Запускаем анализ для каждого критерия
    criteria = [
        ("Выручка", "Выручка без СПП итого, руб"),
        ("Прибыль", "Маржинальная прибыль, руб"),
        ("Маржа", "Маржа, %")
    ]

    for name, column_name in criteria:
        if column_name in result_df.columns:
            abc_df = abc_analysis(result_df, column_name)
            
            # Оставляем только нужные колонки
            final_columns = [article_col, column_name, 'ABC_Category']

            # ⚙️ НОВАЯ ФУНКЦИЯ: показываем первые и последние позиции
            top_bottom = pd.concat([
                abc_df.head(5),          # Первые 5 (категория A)
                abc_df.tail(5)[::-1]     # Последние 5 (категория C), перевёрнутый список
            ])

            st.subheader(f"🔹 ABC-анализ по {name}")
            st.dataframe(top_bottom, use_container_width=True)

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