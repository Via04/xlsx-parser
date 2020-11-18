from collections import OrderedDict

from openpyxl import load_workbook

from xlsx_parser import parse_input

t = 'example_form_2_3.xlsx, y'
path, is_set = parse_input(t)
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
print(ws.cell(7, 871).value)
# print(int(a) + float(b))
a = True and b
print(a)
test = tuple(OrderedDict.fromkeys(test))
# test = ','.join(test)
# test = '"' + test + '"'
print(test)
print(t[0])
print(t[1].strip())
print(path, is_set)
