import uuid
import zipfile
import os
import shutil
import time
import argparse
import logging
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from pyopenxlsx import load_workbook
from datetime import datetime
from xml_convert import generate_xml_from_data
from headers import read_header
from items import parse_items_from_workbook
from post_headers import read_post_header


def rename_shared_strings(zip_path: str) -> None:
    tmp_dir = zip_path + "_tmp"
    os.makedirs(tmp_dir, exist_ok=True)
    with zipfile.ZipFile(zip_path, 'r') as zin:
        zin.extractall(tmp_dir)
    old_path = os.path.join(tmp_dir, 'xl', 'SharedStrings.xml')
    new_path = os.path.join(tmp_dir, 'xl', 'sharedStrings.xml')
    if os.path.exists(old_path):
        os.rename(old_path, new_path)
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for root, _, files in os.walk(tmp_dir):
            for f in files:
                full_path = os.path.join(root, f)
                rel_path = os.path.relpath(full_path, tmp_dir)
                zout.write(full_path, rel_path)
    shutil.rmtree(tmp_dir)


def build_file_id() -> str:
    uid = str(uuid.uuid4()).upper()
    date_str = datetime.now().strftime('%Y%m%d')
    return f'{date_str}_{uid}'


def enrich_header(raw: dict) -> dict:
    raw.setdefault('ИдФайл', build_file_id())
    raw.setdefault('ВерсФорм', '5.03')
    raw.setdefault('РеквНаимДок', 'Универсальный передаточный документ')
    raw.setdefault('РеквНомерДок', raw.get('НомерДок', ''))
    raw.setdefault('РеквДатаДок', raw.get('ДатаДок', ''))
    return raw


def process_xlsx_file(xlsx_path: Path, output_dir: Path) -> str:
    str_file = str(xlsx_path.absolute())
    if not xlsx_path.is_file():
        raise FileNotFoundError(f"Файл {xlsx_path} не найден")
    rename_shared_strings(str_file)
    wb = load_workbook(str_file)
    try:
        ws = wb.active
        header_data = read_header(ws)
        header_data = enrich_header(header_data)
        items_data, sum_data = parse_items_from_workbook(ws)
        post_header_data = read_post_header(ws)
    finally:
        wb.close()
    output = output_dir / f"{xlsx_path.stem}.xml"
    generate_xml_from_data(header_data, post_header_data, items_data, sum_data, output)
    return str(output)


def process_files(xlsx_files, output_dir: Path, max_workers: int = None, logger: logging.Logger = None) -> list:
    if logger is None:
        logger = logging.getLogger(__name__)
    results = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_map = {
            executor.submit(process_xlsx_file, Path(f), output_dir): f
            for f in xlsx_files
        }
        for future in as_completed(future_map):
            f = future_map[future]
            try:
                out = future.result()
                results.append((f, out, None))
                logger.info(f"УПД XML сохранён: {out}")
            except Exception as e:
                results.append((f, None, str(e)))
                logger.error(f"Ошибка обработки {f}: {e}")
    return results


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx", nargs="*", type=str)
    parser.add_argument("--dir", type=str)
    args = parser.parse_args()
    xlsx_files = args.xlsx
    out_dir = Path(args.dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format='%(message)s')
    logger = logging.getLogger(__name__)
    process_files(xlsx_files, out_dir, max_workers=1, logger=logger)


if __name__ == '__main__':
    start_time = time.monotonic()
    main()
    print(time.monotonic() - start_time)
