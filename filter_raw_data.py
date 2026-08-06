#!/usr/bin/env python3
"""
Фильтрует raw данные по датам и сохраняет только отфильтрованные строки.
Использование:
  python filter_raw_data.py input_file.xlsx output_file.xlsx 2026-07-01 2026-07-31
"""
import sys
import pandas as pd
from datetime import datetime
from pathlib import Path

def filter_raw_data(input_path: str, output_path: str,
                    date_from: str, date_to: str) -> None:
    """
    Читает raw данные, фильтрует по датам, сохраняет результат.
    date_from, date_to: строки в формате "YYYY-MM-DD"
    """
    # Читаем файл
    xl = pd.ExcelFile(input_path)

    # Найдем лист с данными (Raw, PowerBi-DailyGoodBadUnits, и т.д.)
    sheet_name = None
    for name in xl.sheet_names:
        if name.lower() in ("raw", "powerbidailygodbadunits", "powerbigoodbadevents"):
            sheet_name = name
            break

    if not sheet_name:
        sheet_name = xl.sheet_names[0]  # берем первый лист если не нашли

    print(f"📖 Читаю лист '{sheet_name}'...")
    df = xl.parse(sheet_name)

    # Найдем столбец с датами
    date_col = None
    for col in df.columns:
        if "date" in col.lower() or "insert" in col.lower():
            date_col = col
            break

    if not date_col:
        print("❌ Не найден столбец с датами (ожидаю 'InsertDate', 'Date' или подобное)")
        return

    print(f"📅 Столбец дат: '{date_col}'")
    print(f"Исходные данные: {len(df)} строк")

    # Конвертируем даты
    df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    start = pd.to_datetime(date_from)
    end = pd.to_datetime(date_to)

    # Фильтруем
    df_filtered = df[(df[date_col] >= start) & (df[date_col] <= end)].copy()

    print(f"✅ После фильтра: {len(df_filtered)} строк")
    print(f"   Диапазон: {df_filtered[date_col].min()} - {df_filtered[date_col].max()}")

    # Сохраняем
    df_filtered.to_excel(output_path, index=False, sheet_name="Raw")
    print(f"💾 Сохранено в {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 5:
        print(__doc__)
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    date_from = sys.argv[3]
    date_to = sys.argv[4]

    filter_raw_data(input_file, output_file, date_from, date_to)
