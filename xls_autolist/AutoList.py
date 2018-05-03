# -*- coding: cp936 -*-  
import xlwt
import time
import list_input

def alignment(style=None):
    """返回居中样式"""
    alignment = xlwt.Alignment() # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    style = style or xlwt.XFStyle() # Create Style
    style.alignment = alignment # Add Alignment to Style
    return style

def medium_borders(style=None):
    """返回实框线"""
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.MEDIUM
    borders.right = xlwt.Borders.MEDIUM
    borders.top = xlwt.Borders.MEDIUM
    borders.bottom = xlwt.Borders.MEDIUM
    style = style or xlwt.XFStyle() # Create Style
    style.borders = borders
    return style

def background_color(style=None):
    """122 Setting the Background Color of a Cell
123 import xlwt
124 workbook = xlwt.Workbook()
125 worksheet = workbook.add_sheet('My Sheet')
126 pattern = xlwt.Pattern() # Create the Pattern
127 pattern.pattern = xlwt.Pattern.SOLID_PATTERN # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
128 pattern.pattern_fore_colour = 5 # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
129 style = xlwt.XFStyle() # Create the Pattern
130 style.pattern = pattern # Add Pattern to Style
131 worksheet.write(0, 0, 'Cell Contents', style)
132 workbook.save('Excel_Workbook.xls')"""
    pass

def pre_with(head,pre,lst):
    """返回对应lst中的index，如果找不到符合要求的，则返回：-1"""
    for j,i in enumerate(lst):
        if i[head.index(u'产品型号')][0:len(pre)] == pre:
            return j
    return -1

def write_sheet_line(worksheet, index, row, style=None):
    style = alignment()
    for i in range(len(row)):
        # 补充：宽度，目前设计为定宽
        worksheet.write(index,i,label=row[i],style=style)    

def main():
    head,lst = list_input.main()
    ITEM_SEQUENCE = [0,2,4,5,6,7,8,9,10]
    
    project_name = u'项目'  # raw_input(u'项目名称:')  or u'项目'
    # 补充:自动扩充V1.00;空就自动添加当天日期和编号

    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet(u'宇视清单')

    # 计算列数
    items = [head[i] for i in ITEM_SEQUENCE]
    row = len(items)

    # 标题
    title = project_name + u'配置清单'  # TITLE
    style = alignment()
    style = medium_borders(style)
    # 补充：底色
    worksheet.write_merge(0,0,0,row-1, title, style)

    # 列标题
    style = alignment()
    # 补充：粗体
    for i,j in enumerate(items):
        # 补充：宽度，目前设计为定宽
        worksheet.write(1,i,label=j,style=style)

    line_index = 1
    
    # 第一部分：摄像机
    # 第二部分：网络存储设备
    # 第三部分：综合管理服务器
    # 第四部分：解码器
    # 第五部分：大屏
##    print len(lst)
    # 最后一个空字符串用来处理不在json列表里面的产品，让其最后能够显示出来
    for pre in ['IPC','NVR','VMS','DC','ADU','MW','']:
        index_tmp = pre_with(head,pre,lst)
        while(index_tmp != -1):
##            print index_tmp,pre
            row_now = lst.pop(index_tmp)
            row_now = [row_now[i] for i in ITEM_SEQUENCE]
            line_index += 1
            write_sheet_line(worksheet,line_index,row_now)
##            print row_now[0:7]
            index_tmp = pre_with(head,pre,lst)
    print len(lst)    
    
    # 总价 
    
    # 保存
    try:
        filename = u'Excel清单%s.xls'%int(time.time())
        workbook.save(filename)  # 补充：名字随项目变化，根据情况考虑版本和重名名
    except Exception,msg:
        print 'Failed!'
        print msg
        
if __name__ == '__main__':
    main()
