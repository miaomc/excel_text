# -*- coding: cp936 -*-  
import xlwt
import time
import list_input

def alignment(style=None):
    """���ؾ�����ʽ"""
    alignment = xlwt.Alignment() # Create Alignment
    alignment.horz = xlwt.Alignment.HORZ_CENTER # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
    alignment.vert = xlwt.Alignment.VERT_CENTER # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
    style = style or xlwt.XFStyle() # Create Style
    style.alignment = alignment # Add Alignment to Style
    return style

def medium_borders(style=None):
    """����ʵ����"""
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
    """���ض�Ӧlst�е�index������Ҳ�������Ҫ��ģ��򷵻أ�-1"""
    for j,i in enumerate(lst):
        if i[head.index(u'��Ʒ�ͺ�')][0:len(pre)] == pre:
            return j
    return -1

def write_sheet_line(worksheet, index, row, style=None):
    style = alignment()
    for i in range(len(row)):
        # ���䣺��ȣ�Ŀǰ���Ϊ����
        worksheet.write(index,i,label=row[i],style=style)    

def main():
    head,lst = list_input.main()
    ITEM_SEQUENCE = [0,2,4,5,6,7,8,9,10]
    
    project_name = u'��Ŀ'  # raw_input(u'��Ŀ����:')  or u'��Ŀ'
    # ����:�Զ�����V1.00;�վ��Զ���ӵ������ںͱ��

    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet(u'�����嵥')

    # ��������
    items = [head[i] for i in ITEM_SEQUENCE]
    row = len(items)

    # ����
    title = project_name + u'�����嵥'  # TITLE
    style = alignment()
    style = medium_borders(style)
    # ���䣺��ɫ
    worksheet.write_merge(0,0,0,row-1, title, style)

    # �б���
    style = alignment()
    # ���䣺����
    for i,j in enumerate(items):
        # ���䣺��ȣ�Ŀǰ���Ϊ����
        worksheet.write(1,i,label=j,style=style)

    line_index = 1
    
    # ��һ���֣������
    # �ڶ����֣�����洢�豸
    # �������֣��ۺϹ��������
    # ���Ĳ��֣�������
    # ���岿�֣�����
##    print len(lst)
    # ���һ�����ַ�������������json�б�����Ĳ�Ʒ����������ܹ���ʾ����
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
    
    # �ܼ� 
    
    # ����
    try:
        filename = u'Excel�嵥%s.xls'%int(time.time())
        workbook.save(filename)  # ���䣺��������Ŀ�仯������������ǰ汾��������
    except Exception,msg:
        print 'Failed!'
        print msg
        
if __name__ == '__main__':
    main()
