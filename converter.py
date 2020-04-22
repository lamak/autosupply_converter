import uuid
from datetime import datetime, timedelta
from os import path
from typing import Tuple

import openpyxl
import pandas as pd


def process_supply(filepath: str, export_path: str, result_path: str, template: str) -> Tuple[str, list]:
    errors = list()
    date_format = '%Y%m%d'
    export_date = datetime.strftime(datetime.now() - timedelta(1), date_format)

    def collect_import_data(filename: str) -> dict:
        results = dict()

        try:
            df = pd.read_excel(
                io=filename,
                na_filter=False,  # не проверяет на попадание в список NaN
                sheet_name='Sheet1',
                usecols=['SKU', 'Код склада'],  # только 2 поля
                converters={'SKU': str, 'Код склада': int},  # типы для полей
            )
            results = df.groupby('Код склада')['SKU'].apply(list).to_dict()
        except Exception as e:
            errors.append(f'Не удалось прочитать файл импорта {e}')

        return results

    def collect_export_results(imp: dict) -> dict:
        results = dict()
        for warehouse in imp.keys():
            exp_filename = f'skubody_{warehouse}_{export_date}.csv'
            exp_filepath = path.join(export_path, exp_filename)

            if path.isfile(exp_filepath):
                warehouse_articles = pd.read_csv(
                    filepath_or_buffer=exp_filepath,
                    sep='¦',
                    header=None,  # без шапки, включая 1 строку
                    usecols=[0, ],  # первое поле с данным
                    squeeze=True,  # т.к. 1 поле, то ужимаем в лист
                    engine='python',  # для корректной обработки разделителя
                    encoding='utf-8',  # обязательно, т.к разделитель
                    converters={0: str}  # поле как строка
                ).to_list()
                results[warehouse] = warehouse_articles

            else:
                errors.append(f'Место хранения {warehouse}: Файл {exp_filepath} недоступен')

        return results

    def make_difference(imp: dict, exp: dict):
        results = dict()
        if imp and exp:
            for warehouse, articles in imp.items():
                list_imp = articles
                list_exp = exp.get(warehouse)
                if list_exp is not None:
                    difference = set(list_imp).difference(set(list_exp))
                    results[warehouse] = difference
                else:
                    errors.append(f'Место хранения {warehouse} пропущено')

        return results

    def write_down(fin: dict) -> str:
        result = ''
        if fin:
            wb = openpyxl.load_workbook(template)
            sh = wb.active
            sh._current_row = 6  # header row, to append after

            today = datetime.now().strftime(date_format)
            result = f'autosupply_results_{today}_{uuid.uuid4()}.xlsx'
            res_path = path.join(result_path, result)

            for wh, articles in fin.items():
                for article in articles:
                    sh.append((article, wh))

            wb.save(res_path)
        return result

    import_results = collect_import_data(filepath)
    export_results = collect_export_results(import_results)
    finale_results = make_difference(import_results, export_results)
    result_filename = write_down(finale_results)

    return result_filename, errors
