import os
import re
from collections import defaultdict

import openpyxl
import xlsxwriter
import xlrd
from xlsxwriter.utility import xl_rowcol_to_cell


def generate_data(from_file_path, school_name):
    print(from_file_path)
    _data = []
    _headers = []

    workbook = xlrd.open_workbook(from_file_path)
    sheets = workbook.sheet_names()

    for n, sheet_name in enumerate(sheets):
        sheet = workbook.sheet_by_name(sheet_name)
        for m, row in enumerate(sheet.get_rows()):
            row_values = [cell.value for cell in row]
            if n == 0 and m == 0:
                _headers.extend(['学校'] + row_values)
            elif row_values[3] == '平均':
                pass
            elif m > 0 and any(row_values):
                row_values.insert(0, school_name)
                _data.append(row_values)
    return _headers, _data


def save_new_excel(to_file_path, data, headers):
    if os.path.exists(to_file_path):
        os.remove(to_file_path)
    with xlsxwriter.Workbook(to_file_path) as wb:
        worksheet = wb.add_worksheet()

        data.insert(0, headers)
        cols = defaultdict(list)
        score_start_row_idx = 5
        for row_idx, row in enumerate(data):
            for col_idx, d in enumerate(row):
                worksheet.write(row_idx, col_idx, d)
                if row_idx > 0 and col_idx >= score_start_row_idx:
                    cols[col_idx].append(d)

        worksheet.write(len(data), 0, '平均')
        for col_idx in range(score_start_row_idx, len(headers)):
            if list(filter(lambda f: f is not None, cols[col_idx])):
                worksheet.write_formula(  # 使用 excel 公式设置平均值
                    len(data),
                    col_idx,
                    f'=average({xl_rowcol_to_cell(1, col_idx)}:{xl_rowcol_to_cell(len(data) - 1, col_idx)})',
                )

        # 设定分数列宽度和数字格式
        score_format = wb.add_format({'num_format': '0.0'})
        worksheet.set_column(score_start_row_idx, len(headers), width=10, cell_format=score_format)
        worksheet.freeze_panes(1, 0)  # 锁定第一行


def do(f_path):
    i = 0
    for _root, _dirs, _filenames in os.walk(f_path):
        _filenames = list(filter(lambda x: x.endswith('xlsx') and x != 'NEW.xlsx', _filenames))
        if _filenames:

            headers = []
            data = []

            for filename in _filenames:
                school_name = filename.rsplit('/', 1)[-1].rsplit('.')[0]
                file_path = os.path.join(_root, filename)
                headers, _data = generate_data(file_path, school_name)
                data.extend(_data)
                # break
            i += 1
            save_new_excel(os.path.join(path, f'{"_".join(_root.rsplit("/")[-2:])}.xlsx'), data, headers)
            # if i >= 3:
            #     break


if __name__ == '__main__':
    # path = "/Users/edz/workshop/tools/2020年7月期末"
    path = "/Users/dong/work/fltrp/workspace/tools/潍坊区县-年级成绩单"
    do(path)
