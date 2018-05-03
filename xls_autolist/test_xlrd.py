# -*- coding: cp936 -*-

import xlrd

data = xlrd.open_workbook('base.xls',formatting_info=True)
table = data.sheets()[0]

ind =table.cell_xf_index(1,2)

# xlrd.formatting.XF
print repr(data.xf_list[ind])
table.cell(1,2)

import xlwt

w = xlwt.Workbook(encoding = 'ascii')


s = w.add_sheet('a')

# style = xlwt.Style.XFStyle
s.write(1,1,label='DddD',style=data.xf_list[108])


w.save('test.xls')
