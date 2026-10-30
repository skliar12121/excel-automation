from pathlib import Path
from openpyxl import Workbook

INPUT_DIR = Path("input")
INPUT_DIR.mkdir(exist_ok=True)


def create_excel(filename, headers, rows):
    wb = Workbook()
    ws = wb.active

    ws.append(headers)

    for row in rows:
        ws.append(row)

    wb.save(INPUT_DIR / filename)
    print(f"Создан файл: {filename}")


create_excel(
    "supplier_1.xlsx",
    ["Артикул", "Название", "Цена"],
    [
        ["A101", "Мышь Logitech", 1500],
        ["A102", "Клавиатура", 2400],
    ],
)

create_excel(
    "supplier_2.xlsx",
    ["SKU", "Товар", "Стоимость"],
    [
        ["A101", "Мышь Logitech", 1450],
        ["A103", "Монитор", 12500],
    ],
)

create_excel(
    "supplier_3.xlsx",
    ["Код", "Наименование", "Цена, руб."],
    [
        ["A104", "Наушники", 3200],
        ["A102", "Клавиатура", None],
    ],
)

print("Все тестовые файлы готовы!")