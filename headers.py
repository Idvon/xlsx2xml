from pyopenxlsx import Worksheet
from typing import Dict, Any

# Множество слов для исключений
HEADER_WORDS = {
    "Индивидуальный",
    "предприниматель",
}
# Координаты свойств шапки формы
COORDS = {
    "1, 17": "НомерДок",
    "1, 26": "ДатаДок",
    "4, 20": ["Фамилия", "Имя", "Отчество"],
    "4, 62": "НаимОрг",
    "5, 20": "СвПродАдрТекст",
    "5, 62": "СвПокупАдрТекст",
    "6, 20": "ИННФЛ",
    "6, 62": ["ИННЮЛ", "КПП"],
    "7, 62": ["НаимОКВ", "КодОКВ"],
}


# Выделение ФИО из ячейки "Продавец"
def name_split(full_name):
    new_name = []
    for name in full_name.split():
        if name not in HEADER_WORDS:
            new_name.append(name)
    return new_name


def read_header(ws: Worksheet) -> Dict:
    """
    Сборка данных из шапки формы в словарь, где ключ аттрибут для данных
    Возвращает структуру: { key1: value1, key2: [ value2, ..., ], ... }
    """
    results = {}
    for key, val in COORDS.items():
        r, c = [int(s) for s in key.split(", ")]
        v = ws.cell(row=r, column=c).value
        if v is None:
            v = ""
        elif r == 4 and c == 20:
            v = name_split(v)
        elif r == 6 and c == 62:
            v = [s for s in v.split("/")]
        elif r == 7 and c == 62:
            v = [s for s in v.split(", ")]
        if isinstance(val, list):
            results.update(dict(zip(val, v)))
        else:
            results[val] = v
    return results
