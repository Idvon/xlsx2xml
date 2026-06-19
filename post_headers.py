import re
from pyopenxlsx import workbook

# Множество слов для исключений
HEADER_WORDS = {
    'ОГРНИП',
}
# Координаты свойств нижней шапки формы
COORDS = {
    '30, 52': ['ОГРНИП', 'ДатаОГРНИП'],
    '50, 3': 'НаимЭконСубСост',
}


# Выделение ОГРНИП и даты из ячейки реквизитов
def find_ogrnip_data(text: str) -> list[str]:
    ogrn = re.search(r'\b(\d{15})\b', text)
    date = re.search(r'\b(\d{2}\.\d{2}\.\d{4})\b', text)
    return [ogrn.group(1) if ogrn else '', date.group(1) if date else '']


def read_post_header(ws: workbook) -> dict:
    """
    Сборка данных из нижней шапки формы в словарь, где ключ аттрибут для данных
    Возвращает структуру: { key1: value1, key2: [ value2, ..., ], ... }
    """
    results = {}
    for key, val in COORDS.items():
        r, c = [int(s) for s in key.split(', ')]
        v = ws.cell(row=r, column=c).value
        if v is None:
            v = ''
        elif r == 30 and c == 52:
            v = find_ogrnip_data(v)
        if isinstance(v, list):
            results.update(dict(zip(val, v)))
        else:
            results[val] = v
    return results
