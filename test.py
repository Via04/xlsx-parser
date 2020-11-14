from openpyxl import load_workbook

t = '252,32'
wb = load_workbook('example_form_2_3.xlsx')
wb.active = 0
a = 4
b = 0.65
ws = wb.active
list_num = t.split(',')
ans = float(list_num[0] + '.' + list_num[1])
print(list_num)
print(ans)
ws['GF7'] = '-252,7332'
wb.save('example_form_2_3.xlsx')
print(ws.cell(7, 871).value)
print(int(a) + float(b))
