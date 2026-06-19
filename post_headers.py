import re
from pyopenxlsx import workbook


HEADER_WORDS = {
    'ОГРНИП',
}

COORDS = {
    '30, 52': ['ОГРНИП', 'ДатаОГРНИП'],
    '50, 3': 'НаимЭконСубСост',
}


def find_ogrnip_data(text: str) -> list[str]:
    ogrn = re.search(r'\b(\d{15})\b', text)
    date = re.search(r'\b(\d{2}\.\d{2}\.\d{4})\b', text)
    return [ogrn.group(1) if ogrn else '', date.group(1) if date else '']


def read_post_header(ws: workbook) -> dict:
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
