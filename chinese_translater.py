from pypinyin import Style, lazy_pinyin
from xlutils.copy import copy
import xlrd


def get_NamePY(str_data):
    rtn = ''
    for i in range(len(str_data)):
        if i == 0:
            rtn = lazy_pinyin(str_data[i], style=Style.NORMAL)
        else:
            rtn += lazy_pinyin(str_data[i], style=Style.NORMAL)
    # 首字母大写
    # rtn = [s.capitalize() for s in rtn]
    # 全大写
    rtn = [s.upper() for s in rtn]
    rtn = ' '.join(rtn)
    return rtn


rb = xlrd.open_workbook('test.xls')
r_sheet = rb.sheet_by_index(0)
wb = copy(rb)
w_sheet = wb.get_sheet(0)

for row_index in range(1, r_sheet.nrows):
    row = r_sheet.row_values(row_index)
    row[1] = get_NamePY(row[0])
    print(row)
    w_sheet.write(row_index, 1, row[1])

wb.save('test.xls')