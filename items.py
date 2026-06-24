from typing import Dict, List, Tuple

from pyopenxlsx import Worksheet

# Координаты столбцов товаров в форме
CODE_FORMS_DICT = {
    2: "КодТов",
    6: "НомСтр",
    9: "НаимТов",
    18: "КодВидТов",
    19: "ОКЕИ_Тов",
    23: "НаимЕдИзм",
    27: "КолТов",
    30: "ЦенаТов",
    36: "СтТовБезНДС",
    43: "БезАкциз",
    47: "НалСт",
    50: "СумНал",
    54: "СтТовУчНал",
    56: "КодПроисх",
    59: "КрНаимСтрПр",
    61: "НомТовПрослеж",
    66: "ЕдИзмПрослеж",
    71: "НаимЕдИзмПрослеж",
    73: "КолВЕдПрослеж",
    74: "СтТовБезНДСПрослеж",
}
# Координаты столбцов суммы в форме
CODE_SUM_DICT = {
    36: "СтТовБезНДСВсего",
    50: "СумНал",
    54: "СтТовУчНалВсего",
}


def parse_items_from_workbook(ws: Worksheet) -> Tuple[List[dict], Dict, int]:
    """
    Парсит товары с листа Excel-файла.
    Листы содержат заголовки свойств на row=16 (code row).
    Возвращает структуру: [ { code1: value1, code2: value2, ..., }, ... ].
    """
    oplata_row: int = 0
    items = []
    sum_items = {}
    max_row = ws.max_row
    fixed_row = 16  # строка заголовков свойств (1-based в openpyxl)
    # Итерируем по строкам после заголовков
    for r in range(fixed_row + 1, max_row + 1):
        # Проверка конца товара (строка начинается с "всего к оплате")
        first_col_val = ws.cell(row=r, column=6).value
        if isinstance(first_col_val, str) and first_col_val == "Всего к оплате (9)":
            oplata_row = r
            for k in CODE_SUM_DICT.keys():
                sum_items[CODE_SUM_DICT[k]] = ws.cell(row=r, column=k).value
            break
        # Сборка товаров и услуг
        item = {}
        for c, code in CODE_FORMS_DICT.items():
            val = ws.cell(row=r, column=c).value
            if isinstance(val, str) and (
                "А" in val.split(" ")
                or "УПД" in val.split(" ")
                or "Код" in val.split(" ")
                or "код" in val.split(" ")
            ):
                break
            if val is not None and val != "--":
                item[code] = val
        if item != {}:
            items.append(item)
    return items, sum_items, oplata_row
