import pandas as pd # министерство совпадений
from tkinter import filedialog
from tkinter import Tk
from itertools import chain, combinations
import numpy as np

def put():
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title='Выберите файл', filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")))
    return file_path

def otrezki(df, KOL_OTREZKOV, COLUMN_NAME):
    min_val = df[COLUMN_NAME].min()
    max_val = df[COLUMN_NAME].max()
    range_size = (max_val - min_val) / KOL_OTREZKOV

    def calculate_range(value):
        if pd.isna(value):
            return None
        bucket = min(int((value - min_val) / range_size), KOL_OTREZKOV - 1)
        lower_bound = min_val + bucket * range_size
        upper_bound = lower_bound + range_size
        if upper_bound > max_val or bucket == KOL_OTREZKOV - 1:
            upper_bound = max_val
        return f"{lower_bound:.2f}-{upper_bound:.2f}"

    df[f"{COLUMN_NAME}_s"] = df[COLUMN_NAME].apply(calculate_range)
    return df
def process_dataframe(df, KOL_OTREZKOV):
    numeric_columns = df.select_dtypes(include=[np.number]).columns  # Выбор всех числовых столбцов
    for column in numeric_columns:
        df = otrezki(df, KOL_OTREZKOV, column)
    return df

def transposition(df, x):
    # Создаем временный DataFrame для хранения результатов
    sorted_df = pd.DataFrame()

    # Обрабатываем каждую пару столбцов
    for i in range(0, df.shape[1], 2):
        # Получаем пару: строковый и числовой столбцы
        name_col = df.iloc[:, i]
        value_col = df.iloc[:, i+1]

        # Сортируем оба столбца по значениям числового столбца по убыванию
        pair_df = pd.concat([name_col, value_col], axis=1)
        pair_df = pair_df.sort_values(by=df.columns[i+1], ascending=False).reset_index(drop=True)

        # Добавляем отсортированные столбцы в временный DataFrame
        sorted_df = pd.concat([sorted_df, pair_df], axis=1)

    # Рассчитываем максимальные значения в первой строке для всех числовых столбцов
    max_values = [sorted_df.iloc[0, i+1] for i in range(0, sorted_df.shape[1], 2)]
    pairs = list(zip(max_values, range(0, sorted_df.shape[1], 2)))

    # Сортируем пары по максимальным значениям (убывание)
    pairs.sort(reverse=True, key=lambda x: x[0])

    # Создаем новый DataFrame с отсортированными парами столбцов
    final_df = pd.DataFrame()
    for _, idx in pairs[:x]:
        final_df = pd.concat([final_df, sorted_df.iloc[:, idx:idx+2]], axis=1)

    return final_df

def delcopy(sdf):
    # Удаляем дубликаты по всем столбцам
    return sdf.drop_duplicates()

def consist(df, max_columns):
    # Функция для генерации всех возможных непустых подмножеств до определённого размера
    def all_subsets(ss, max_len):
        return chain(*[combinations(ss, i) for i in range(1, max_len+1)])

    results = pd.DataFrame()
    all_combinations = list(all_subsets(df.columns, max_columns))  # Генерируем все комбинации столбцов до max_columns

    for cols in all_combinations:
        group_name = '+'.join(cols)  # Формируем уникальное имя группы
        group_data = df.groupby(list(cols)).size().reset_index(name='Count')
        group_data[group_name] = group_data[list(cols)].apply(lambda x: '+'.join(x.dropna().astype(str)), axis=1)
        if len(group_data)<len(df)*0.9:
            results = pd.concat([results, group_data[[group_name, 'Count']]], axis=1)
            print(cols, len(group_data), len(df))
        else:
            print("NOT", cols, len(group_data), len(df))
        #print(group_data)
        #print(results)

    return results

def main():
    file_path = put()
    if not file_path:
        return
    df = pd.read_excel(file_path, engine='openpyxl')
    print('Файл загружен:', file_path)

    df = process_dataframe(df, 5)

    # Применяем функцию составления данных
    processed_df = consist(df, 2)
    print('Группировка данных завершена')

    # Применяем функцию транспозиции
    transposed_df = transposition(processed_df, 100)
    print('Транспозиция выполнена')

    # Удаляем дубликаты
    final_df = delcopy(transposed_df)
    print('Удаление дубликатов выполнено')

    # Сохраняем результат в Excel
    output_path = file_path.rsplit('.', 1)[0] + '_processed.xlsx'
    final_df.to_excel(output_path, index=False)
    print('Файл сохранен:', output_path)

if __name__ == "__main__":
    main()
