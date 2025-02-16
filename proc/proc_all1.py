import pandas as pd

def annotation(input_path, output_path):
    # Загружаем таблицу
    df = pd.read_excel(input_path)
    
    # Преобразуем столбцы в строки и заменяем NaN на пустые строки
    df = df.astype(str).fillna('')
    
    # Группируем по "Запрашиваемый номер" и объединяем "Номер" и "описание"
    df_grouped = df.groupby("Запрашиваемый номер").agg({
        'Номер': lambda x: '; '.join(sorted(set(filter(None, x)))),
        'описание': lambda x: '; '.join(sorted(set(filter(None, x))))
    }).reset_index()

    # Сохраняем в новый файл
    df_grouped.to_excel(output_path, index=False)
