import csv
import os
import uuid
from argparse import ArgumentParser
from datetime import datetime, timedelta
import openpyxl

parser = ArgumentParser()
parser.add_argument("-f", "--file", dest="filename", help="Импорт из файла", metavar="FILE")
parser.add_argument("-t", "--template", dest="template", help="Шаблон экспорта", metavar="TEMPLATE")
parser.add_argument("-p", "--prefix", dest="prefix", help="Префикс экспорта из Oracle", metavar="PREFIX")
parser.add_argument("-d", "--date", dest="date", help="Дата экспорта из Oracle ГГГГММДД", metavar="DATE")
parser.add_argument("-df", "--dateformat", dest="dateformat", help="Формат ГГГГММДД", metavar="DF")
args = parser.parse_args()

date_format = args.dateformat if args.dateformat else '%Y%m%d'
filename_prefix = args.prefix if args.prefix else 'skubody'
import_filename = args.filename if args.filename else 'remi_AllBuffersManagement.xlsx'
template_filename = args.template if args.template else 'sku-body-template.xlsx'
export_day = args.date if args.date else datetime.strftime(datetime.now() - timedelta(1), date_format)
today = datetime.now().strftime(date_format)

export_template_base_row = 7
export_path = '//vl20-srv15/d$/smAbmLoader/Export/CsvData/'

import_results = dict()
export_results = dict()
finale_results = dict()
errors = list()


def insert_or_append(d: dict, k: str, v: str):
    if d.get(k):
        d[k].append(v)
    else:
        d[k] = [v, ]


# parse incoming xlsx into map by warehouse
try:
    wb = openpyxl.load_workbook(import_filename)
    ws = wb.get_active_sheet()

    # validate worksheet header
    if not (ws.cell(1, 2).value == 'SKU' and ws.cell(1, 5).value == 'Код склада'):
        errors.append('Не найдены заголовки таблицы')
    else:
        ws.delete_rows(1)
        for i, r in enumerate(ws.rows):
            article = r[1].value
            warehouse = r[4].value
            insert_or_append(import_results, warehouse, article)

    wb.close()
except openpyxl.utils.exceptions.InvalidFileException:
    errors.append('Не удалось прочитать файл импорта')

# if import data not null, so get the results from exported data
if import_results:
    for wh in import_results.keys():
        filename_warehouse = wh
        filename = f'{filename_prefix}_{filename_warehouse}_{export_day}.csv'
        filepath = os.path.join(export_path, filename)
        # todo: add error handling on filenames, get yesterday if today not available
        if os.path.isfile(filepath):

            with open(filepath, 'r', encoding='utf-8') as f:
                reader = csv.reader(f, delimiter='¦')
                for i in reader:
                    article = i[0]
                    warehouse = i[1]  # also maybe key if we sure in exported data
                    insert_or_append(export_results, warehouse, article)
        else:
            errors.append(f'Место хранения {wh}: Файл {filepath} недоступен')

if import_results and export_results:
    for wh, articles in import_results.items():
        list_imp = articles
        list_exp = export_results.get(wh)
        if list_exp is not None:
            finale_results[wh] = set(list_imp).difference(set(list_exp))
        else:
            errors.append(f'Место хранения {wh} пропущено')
    print(
        f'Места хранения:\n'
        f'Импорта: {", ".join(import_results.keys())},\n'
        f'Экспорта: {", ".join(export_results.keys())},\n'
        f'Результат: {", ".join(finale_results.keys())}'
    )

if finale_results:
    results_wb = openpyxl.load_workbook(template_filename)
    results_sh = results_wb.get_active_sheet()

    for wh, articles in finale_results.items():
        for i, article in enumerate(articles):
            results_sh.cell(export_template_base_row + i, 1).value = article
            results_sh.cell(export_template_base_row + i, 2).value = wh
        export_template_base_row = export_template_base_row + i

    filename_new = f'autosupply_results_{today}_{uuid.uuid4()}.xlsx'
    results_wb.save(filename_new)
    print(f'Результат записан в файл: {filename_new}')

if errors:
    print('\n'.join(errors))
