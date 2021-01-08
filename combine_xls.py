import glob
import xlrd
import xlwt
import time

"""将多个*.xls文件合并到all.xls"""
xls_list = glob.glob(r'*.xls')

# 将多个*.xls文件里的所有行读取到 all_lines里面
all_lines = []
for i in xls_list:
    data = xlrd.open_workbook(i)
    table = data.sheets()[0]
    
    # WARNING *** OLE2 inconsistency: SSCS size is 0 but SSAT size is non-zero
    # for j in range(0, table.nrows):
    for tmp in table._cell_values:
        all_lines.append(tmp)

# 将all_lines写入all_.xls
workbook = xlwt.Workbook(encoding = 'ascii')
worksheet = workbook.add_sheet(u'all')

line_index = 0
for line in all_lines:
    for ind, cell in enumerate(line):
        worksheet.write(line_index, ind , cell)
    line_index += 1

workbook.save('all_'+str(time.time())+'.xls')
