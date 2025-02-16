import pandas as pd
import re

def convert_weight(value):
    value = re.sub(r"Вес в инд\. упак\.\s*", "", value)
    match = re.search(r"([\d.]+)\s*кг", value)
    if match:
        kg = float(match.group(1))
        return int(kg * 1000)  # Конвертируем в граммы
    return value

def convert_length(value):
    value = re.sub(r"Длина резьбы|Длина инструмента, мм|Длина, мм|Высота, мм|Ширина, мм|Высота подхвата, мм|Длина ремня, мм|Ширина ремня, мм|Высота ремня, мм|Длина троса, |Длина лезвия, мм", "", value)
    match = re.search(r"([\d.]+)\s*см", value)
    match_m = re.search(r"м([\d.])", value)
    if match:
        cm = float(match.group(1))
        return int(cm * 10)  # Конвертируем в миллиметры
    elif match_m:
        m = float(match.group(1))
        return int(m * 1000)  # Конвертируем в миллиметры
    return value

def clean_values(value):
    if pd.isna(value) or value in ["0", "nan"]:
        return ""
    return value

def process_excel(file_path, output_path):
    df = pd.read_excel(file_path)
    
    df.iloc[:, 0] = df.iloc[:, 0].astype(str).apply(convert_weight)  # Вес
    df.iloc[:, 1:] = df.iloc[:, 1:].astype(str).applymap(convert_length)  # Остальные размеры
    df = df.applymap(clean_values)  # Очистка значений
    
    df.to_excel(output_path, index=False)
    print(f"Файл сохранен")

# Использование
file_path = "C:\\py\\spy\\excel\\303-333.xlsx"
output_path = "C:\\py\\spy\\excel\\303-scripted.xlsx"
process_excel(file_path, output_path)
