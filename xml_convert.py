import re
import uuid
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path


def _fmt(val, decimals=2):
    """
    Преобразование чисел в float и str
    """
    if val is None:
        return ""
    try:
        return f"{float(val):.{decimals}f}"
    except (ValueError, TypeError):
        return str(val)


def _make_sub_element(parent, tag, attrib=None, text=None):
    """
    Создание под элемента
    """
    el = ET.SubElement(parent, tag)
    if attrib:
        for k, v in attrib.items():
            if v is not None and str(v).strip():
                el.set(k, str(v))
    if text is not None:
        el.text = str(text)
    return el


def generate_xml_from_data(
    header: dict, post_header: dict, items: list, sum_v: dict, output_path: Path
) -> None:
    """
    Генерация xml файла по схеме
    """
    root = ET.Element("Файл")
    root.set(
        "ИдФайл",
        header.get(
            "ИдФайл", f"{datetime.now().strftime('%d%m%Y')}_{str(uuid.uuid4()).upper()}"
        ),
    )
    root.set("ВерсПрог", header.get("ВерсПрог", "xlsx2xml"))
    root.set("ВерсФорм", header.get("ВерсФорм", "5.03"))

    doc = ET.SubElement(root, "Документ")
    doc.set("КНД", "1115131")
    doc.set("ВремИнфПр", "00.00.00")
    doc.set("ДатаИнфПр", header.get("ДатаДок", ""))
    doc.set("НаимЭконСубСост", post_header.get("НаимЭконСубСост", ""))
    doc.set("Функция", "СЧФДОП")
    doc.set(
        "ПоФактХЖ",
        "Документ об отгрузке товаров (выполнении работ),"
        "передаче имущественных прав (документ об оказании услуг)",
    )
    doc.set("НаимДокОпр", "Универсальный передаточный документ")

    sv = ET.SubElement(doc, "СвСчФакт")
    sv.set("НомерДок", str(header.get("НомерДок", "")))
    sv.set("ДатаДок", header.get("ДатаДок", ""))

    sv_prod = ET.SubElement(sv, "СвПрод")
    id_sv_prod = ET.SubElement(sv_prod, "ИдСв")
    if header.get("ИННФЛ"):
        sv_ip = ET.SubElement(id_sv_prod, "СвИП")
        sv_ip.set("ИННФЛ", str(header["ИННФЛ"]))
        if post_header.get("ОГРНИП"):
            sv_ip.set("ОГРНИП", str(post_header["ОГРНИП"]))
        if post_header.get("ДатаОГРНИП"):
            sv_ip.set("ДатаОГРНИП", post_header["ДатаОГРНИП"])
        fio = ET.SubElement(sv_ip, "ФИО")
        fio.set("Фамилия", header.get("Фамилия", ""))
        fio.set("Имя", header.get("Имя", ""))
        fio.set("Отчество", header.get("Отчество", ""))
    else:
        sv_ul = ET.SubElement(id_sv_prod, "СвЮЛУч")
        sv_ul.set("НаимОрг", header.get("НаимОрг", ""))
        sv_ul.set("ИННЮЛ", str(header.get("ИННЮЛ", "")))
        sv_ul.set("КПП", str(header.get("КПП", "")))

    addr_prod = ET.SubElement(sv_prod, "Адрес")
    addr_info_prod = ET.SubElement(addr_prod, "АдрИнф")
    addr_info_prod.set("КодСтр", "643")
    addr_info_prod.set("АдрТекст", header.get("СвПродАдрТекст", ""))
    addr_info_prod.set("НаимСтран", "РОССИЯ")

    doc_podtv = ET.SubElement(sv, "ДокПодтвОтгрНом")
    doc_podtv.set(
        "РеквНаимДок", header.get("РеквНаимДок", "Универсальный передаточный документ")
    )
    doc_podtv.set("РеквНомерДок", str(header.get("НомерДок", "")))
    doc_podtv.set("РеквДатаДок", header.get("ДатаДок", ""))

    sv_pokup = ET.SubElement(sv, "СвПокуп")
    id_sv_pokup = ET.SubElement(sv_pokup, "ИдСв")
    if header.get("ИННЮЛ"):
        sv_ul_pok = ET.SubElement(id_sv_pokup, "СвЮЛУч")
        sv_ul_pok.set("НаимОрг", header.get("НаимОрг", ""))
        sv_ul_pok.set("ИННЮЛ", str(header["ИННЮЛ"]))
        sv_ul_pok.set("КПП", str(header.get("КПП", "")))
    elif header.get("ИННФЛ"):
        sv_ip_pok = ET.SubElement(id_sv_pokup, "СвИП")
        sv_ip_pok.set("ИННФЛ", str(header["ИННФЛ"]))
        fio_pok = ET.SubElement(sv_ip_pok, "ФИО")
        fio_pok.set("Фамилия", header.get("Фамилия", ""))
        fio_pok.set("Имя", header.get("Имя", ""))
        fio_pok.set("Отчество", header.get("Отчество", ""))

    addr_pokup = ET.SubElement(sv_pokup, "Адрес")
    addr_info_pokup = ET.SubElement(addr_pokup, "АдрИнф")
    addr_info_pokup.set("КодСтр", "643")
    addr_info_pokup.set("АдрТекст", header.get("СвПокупАдрТекст", ""))
    addr_info_pokup.set("НаимСтран", "РОССИЯ")

    den_izm = ET.SubElement(sv, "ДенИзм")
    den_izm.set("КодОКВ", header.get("КодОКВ", "643"))
    den_izm.set("НаимОКВ", header.get("НаимОКВ", "Российский рубль"))

    table = ET.SubElement(doc, "ТаблСчФакт")

    for item in items:
        row = ET.SubElement(table, "СведТов")
        row.set("НомСтр", str(item.get("НомСтр", "")))
        row.set("НаимТов", str(item.get("НаимТов", "")))
        row.set("ОКЕИ_Тов", str(item.get("ОКЕИ_Тов", "")))
        row.set("НаимЕдИзм", str(item.get("НаимЕдИзм", "")))

        row.set("КолТов", _fmt(item.get("КолТов", ""), 3))
        row.set("ЦенаТов", _fmt(item.get("ЦенаТов", ""), 2))
        row.set("СтТовБезНДС", _fmt(item.get("СтТовБезНДС", ""), 2))
        row.set("НалСт", item.get("НалСт", ""))
        row.set("СтТовУчНал", _fmt(item.get("СтТовУчНал", ""), 2))

        dop_sved = ET.SubElement(row, "ДопСведТов")
        _make_sub_element(dop_sved, "КодПроисх", text=str(item.get("КодПроисх", "")))
        _make_sub_element(
            dop_sved, "КрНаимСтрПр", text=str(item.get("КрНаимСтрПр", ""))
        )
        _make_sub_element(dop_sved, "КодТов", text=str(item.get("КодТов", "")))
        _make_sub_element(dop_sved, "КодВидТов", text=str(item.get("КодВидТов", "")))

        sved_pros = ET.SubElement(dop_sved, "СведПрослеж")
        _make_sub_element(
            sved_pros, "НомТовПрослеж", text=str(item.get("НомТовПрослеж", ""))
        )
        _make_sub_element(
            sved_pros, "ЕдИзмПрослеж", text=str(item.get("ЕдИзмПрослеж", ""))
        )
        _make_sub_element(
            sved_pros, "НаимЕдИзмПрослеж", text=str(item.get("НаимЕдИзмПрослеж", ""))
        )
        _make_sub_element(
            sved_pros, "КолВЕдПрослеж", text=_fmt(item.get("КолВЕдПрослеж", ""), 3)
        )
        _make_sub_element(
            sved_pros,
            "СтТовБезНДСПрослеж",
            text=_fmt(item.get("СтТовБезНДСПрослеж", ""), 2),
        )

        akciz = ET.SubElement(row, "Акциз")
        _make_sub_element(akciz, "БезАкциз", text=str(item.get("БезАкциз", "")))

        sum_nal_el = ET.SubElement(row, "СумНал")
        _make_sub_element(sum_nal_el, "СумНал", text=_fmt(item.get("СумНал", ""), 2))

    total_el = ET.SubElement(table, "ВсегоОпл")
    total_el.set("СтТовБезНДСВсего", _fmt(sum_v.get("СтТовБезНДСВсего", ""), 2))
    total_el.set("СтТовУчНалВсего", _fmt(sum_v.get("СтТовУчНалВсего", ""), 2))
    sum_nal_total_el = ET.SubElement(total_el, "СумНалВсего")
    _make_sub_element(sum_nal_total_el, "СумНал", text=_fmt(sum_v.get("СумНал", ""), 2))

    sv_per_el = ET.SubElement(doc, "СвПродПер")
    sv_per = ET.SubElement(sv_per_el, "СвПер")
    sv_per.set(
        "СодОпер", "Товары переданы, работы выполнены, услуги оказаны в полном объеме"
    )
    sv_per.set("ДатаПер", header.get("ДатаДок", ""))
    _make_sub_element(sv_per, "БезДокОснПер", text="1")

    podp = ET.SubElement(doc, "Подписант")
    podp.set("СпосПодтПолном", "2")
    podp.set("ТипПодпис", "1")
    fio_podp = ET.SubElement(podp, "ФИО")
    fio_podp.set("Фамилия", header.get("ФамилияПодписант", header.get("Фамилия", " ")))
    fio_podp.set("Имя", header.get("ИмяПодписант", header.get("Имя", " ")))
    fio_podp.set(
        "Отчество", header.get("ОтчествоПодписант", header.get("Отчество", " "))
    )

    tree = ET.ElementTree(root)
    ET.indent(tree, space="  ")
    tree.write(output_path, encoding="windows-1251", xml_declaration=True)
    with open(output_path, "rb+") as f:
        content = f.read()
        content = re.sub(rb"(\S) />", rb"\1/>", content)
        f.seek(0)
        f.write(content)
        f.truncate()
