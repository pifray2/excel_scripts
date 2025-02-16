import pandas as pd

def oam(input_path, output_path):
    # Загружаем таблицу
    df = pd.read_excel(input_path)
    
    # Преобразуем столбец "Кросс-деталь" в строки, заменяя NaN на пустые строки
    df['Кросс-деталь'] = df['Кросс-деталь'].astype(str)
    df['Кросс-деталь'] = df['Кросс-деталь'].replace('nan', '').replace('NaN', '').replace('None', '')
    
    # Группируем по "Артикул", убираем NaN, дубликаты и объединяем через запятую
    df_grouped = df.groupby("Артикул")['Кросс-деталь'].apply(lambda x: '; '.join(sorted(set(filter(None, x))))).reset_index()

    # Сохраняем в новый файл
    df_grouped.to_excel(output_path, index=False)
