from pyopenxlsx import Worksheet
from typing import Dict, List, Tuple

# Множество слов для исключений
HEADER_WORDS = {
    "Индивидуальный",
    "предприниматель",
}

# Координаты свойств шапки формы
COORDS = {
    "СвПрод": {
        "5, 20": "СвПродАдрТекст",
        "6, 20": "",
    },
    "СвПокуп": {
        "5, 62": "СвПокупАдрТекст",
        "6, 62": "",
        "7, 62": ["НаимОКВ", "КодОКВ"],
    }
}


# Выделение ФИО из ячейки "Продавец"
def name_split(full_name) -> List:
    new_name = []
    for name in full_name.split():
        if name not in HEADER_WORDS:
            new_name.append(name)
    return new_name


def read_header(ws: Worksheet) -> Tuple[Dict, Dict]:
    """
    Сборка данных из шапки формы в словарь, где ключ аттрибут для данных
    Возвращает структуру: { key1: value1, key2: [ value2, ..., ], ... }
    """
    sv_prod = {}
    results = {"НомерДок": ws.cell(row=1, column=17).value, "ДатаДок": ws.cell(row=1, column=26).value}
    for key, val in COORDS.items():
        result = {}
        for key_i, val_i in val.items():
            r, c = [int(s) for s in key_i.split(", ")]
            v = ws.cell(row=r, column=c).value
            if (r == 6 and c == 20) or (r == 6 and c == 62):
                v = str(v).split("/")
                if len(v) == 2:
                    val_i = ["ИННЮЛ", "КПП"]
                    result["НаимОрг"] = ws.cell(row=4, column=c).value
                    if key == "СвПрод":
                        sv_prod["НаимОрг"] = result["НаимОрг"]
                else:
                    val_i = "ИННФЛ"
                    v = v[0]
                    val_ip = ["Фамилия", "Имя", "Отчество"]
                    v_ip = ws.cell(row=4, column=c).value
                    v_ip = name_split(v_ip)
                    result.update(dict(zip(val_ip, v_ip)))
            if r == 7 and c == 62:
                v = [s for s in v.split(", ")]
            if isinstance(val_i, list):
                result.update(dict(zip(val_i, v)))
            else:
                result[val_i] = v
        results[key] = result
    return results, sv_prod
