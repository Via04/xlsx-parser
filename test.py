from collections import OrderedDict

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

wb = load_workbook('example_form_2_3.xlsx')
wb.active = 0
a = None
b = True
test = ['a', 'b', 'c']
ws = wb.active
# ws['GF7'] = '-252,7332'
# ws['H5'] = 'asgga'
# ws['I5'] = 'qwet'
# wb.save('example_form_2_3.xlsx')
# print(int(a) + float(b))
# a = True and b
print(a)
ws.cell(7, column_index_from_string('H'), 'asdg')
ws.cell(7, column_index_from_string('G'), '123fef')
test = tuple(OrderedDict.fromkeys(test))
# test = ','.join(test)
# test = '"' + test + '"'
print(test)
# print(t[0])
# print(t[1].strip())
# print(path, is_set)
wb.save('example_form_2_3.xlsx')