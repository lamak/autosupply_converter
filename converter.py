import csv
import os
import uuid
from argparse import ArgumentParser
from datetime import datetime, timedelta
import openpyxl

parser = ArgumentParser()
parser.add_argument("-f", "--file", dest="filename", help="Импорт из файла", metavar="FILE")
parser.add_argument("-t", "--template", dest="template", help="Шаблон экспорта", metavar="TEMPLATE")
parser.add_argument("-d", "--date", dest="date", help="Дата экспорта из Oracle ГГГГММДД", metavar="DATE")
args = parser.parse_args()

date_format = '%Y%m%d'

import_file = args.filename if args.filename else 'remi_AllBuffersManagement.xlsx'
template_file = args.template if args.template else 'sku-body-template.xlsx'
export_date = args.date if args.date else datetime.strftime(datetime.now() - timedelta(1), date_format)

# export_path = '//vl20-srv15/d$/smAbmLoader/Export/CsvData/'
export_path = './'

errors = list()


def insert_or_append(d: dict, k: str, v: str):
    if d.get(k):
        d[k].append(v)
    else:
        d[k] = [v, ]


def collect_import_data(filename: str) -> dict:
    results = dict()
    try:
        wb = openpyxl.load_workbook(filename)
        ws = wb.get_active_sheet()

        # validate worksheet header
        if not (ws.cell(1, 2).value == 'SKU' and ws.cell(1, 5).value == 'Код склада'):
            errors.append('Не найдены заголовки таблицы')
        else:
            ws.delete_rows(1)
            for r in ws.rows:
                article = r[1].value
                warehouse = r[4].value
                insert_or_append(results, warehouse, article)

        wb.close()
    except openpyxl.utils.exceptions.InvalidFileException:
        errors.append('Не удалось прочитать файл импорта')

    return results


def collect_export_results(imp: dict) -> dict:
    results = dict()
    if imp:
        for warehouse in imp.keys():
            filename = f'skubody_{warehouse}_{export_date}.csv'
            filepath = os.path.join(export_path, filename)
            if os.path.isfile(filepath):
                with open(filepath, 'r', encoding='utf-8') as f:
                    csv_data = csv.reader(f, delimiter='¦')
                    for row in csv_data:
                        article = row[0]
                        insert_or_append(results, warehouse, article)
            else:
                errors.append(f'Место хранения {warehouse}: Файл {filepath} недоступен')

    return results


def make_difference(imp: dict, exp: dict):
    results = dict()
    if imp and exp:
        for warehouse, articles in imp.items():
            list_imp = articles
            list_exp = exp.get(warehouse)
            if list_exp is not None:
                results[warehouse] = set(list_imp).difference(set(list_exp))
            else:
                errors.append(f'Место хранения {warehouse} пропущено')

    return results


def write_down(fin: dict):
    if fin:
        current_row = 7  # first row after header
        wb = openpyxl.load_workbook(template_file)
        sh = wb.get_active_sheet()
        today = datetime.now().strftime(date_format)

        filename = f'autosupply_results_{today}_{uuid.uuid4()}.xlsx'

        for wh, articles in fin.items():
            idx = 0
            for idx, article in enumerate(articles):
                sh.cell(current_row + idx, 1).value = article
                sh.cell(current_row + idx, 2).value = wh
            current_row = current_row + idx

        wb.save(filename)
        print(f'Результат записан в файл: {filename}')


def __main__():
    import_results = collect_import_data(import_file)
    export_results = collect_export_results(import_results)
    finale_results = make_difference(import_results, export_results)

    print(
        f'Места хранения:\n'
        f'Импорта: {", ".join(import_results.keys())},\n'
        f'Экспорта: {", ".join(export_results.keys())},\n'
        f'Результат: {", ".join(finale_results.keys())}'
    )

    write_down(finale_results)

    if errors:
        print('\n'.join(errors))


__main__()
