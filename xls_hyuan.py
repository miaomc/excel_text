# -*- coding:cp936 -*-
import glob
import xlrd
import xlwt
import json
    
def all2json():
    """将多个*.xls文件，读取到一个products的列表中，然后存储到all.json中"""
    xls_list = glob.glob(r'*.xls')

    products = []  # content all the products' infomation
    for i in xls_list:
        data = xlrd.open_workbook(i)
        table = data.sheets()[0]   
        for j in range(1, table.nrows):
            tmp = table.row_values(j)
            products.append(tmp)

    with open('all.json','w') as f1:
        json.dump(products,fp=f1)

def deal_data():
    """将json中的文件，全部写到all.xls表格中（分开处理，因为all2json()读取很多文件比较慢）"""
    with open('all.json','r') as f1:
        products = json.load(f1)
        
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet(u'ALL')

    for ind,i in enumerate(products):
        worksheet.write(ind,0,str(ind))
        for j,k in enumerate(i):
            worksheet.write(ind,j+1,k)
            
    workbook.save('all.xls')

def head_with(head,tmp):
    """可以用tmp.startswith(head)替换"""
    return tmp[0:len(head)] == head
        
def read_from_case():
    """根据需求，整合xls文件，返回列表lists"""
    # read from xls
    data = xlrd.open_workbook('all_origin.xls')
    table = data.sheets()[0]
    lines= []
    for j in range(1, table.nrows):
        tmp = table.row_values(j)
        lines.append(tmp)

    print len(lines)

    # deal data ---->[  [[‘包装指令：..’，...],['1-R09','纸'，‘编码’，‘实发数’],['1-R09','纸'，‘编码’，‘实发数’],...], ... ]
    lists = []
    one_case = False
    for i in lines:
        if head_with(u'包装指令',i[0]):
            baozhuang = i[0]
        elif head_with(u'装箱单号',i[0]):
            zhuangxiang = i[2]
        elif head_with(u'大箱',i[0]):
            one_case = True
            tmp = [[baozhuang,zhuangxiang]]
        elif head_with(u'备注',i[0]):
            one_case = False
            lists.append(tmp)
        elif one_case:
            tmp.append([i[0],i[1],i[3],i[7]])

    return lists
            
def read_from_sucai():
    """根据需求，整合xls文件，返回列表lists"""
    # read from xls
    data = xlrd.open_workbook('sucai.xls')
    table = data.sheets()[0]
    lines= []
    for j in range(1, table.nrows):
        tmp = table.row_values(j)
        lines.append(tmp)

    # create data ---> {'BOM1':{'xinghao':'','lists':[[u'主机',1,u'台'],[],[]]},'BOM2':{},...}
    dirs = {}
    for i in lines:
        #print repr(i[0])
        if len(i[0])>3:  # 排除空格和标题
            dirs[i[0]] = {'xinghao':i[1],'lists':[[i[2],int(i[3]),i[4]]]}
            key = i[0]
        elif i[0] != 'BOM':
            dirs[key]['lists'].append([i[2],int(i[3]),i[4]])

    return dirs
            
def write2xls(l,d):
    """最终写入"""
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet(u'监控系统')

    # l = [  [[‘包装指令：..’，...],['1-R09','纸'，‘BOM编码’，‘实发数’],['1-R09','纸'，‘编码’，‘实发数’],...], ... ]
    # d = {'BOM1':{'xinghao':'','lists':[[u'主机',1,u'台'],[],[]]},'BOM2':{},...}

    index_st = -1
    index_ed = -1
    for case in l:
        # case =  [[‘包装指令：..’，...],['1-R09','纸'，‘BOM编码’，‘实发数’],['1-R09','纸'，‘编码’，‘实发数’],...]
        for i in case[1:]:
            # i = ['1-R09','纸'，‘BOM编码’，‘实发数’]
            # from index_start to index_end
            index_st = index_ed + 1  # 此处保证可空一行
            if i[2] in d.keys():
                index_ed = index_st + len(d[i[2]]['lists'])  # [index_st, index_ed )
            else:
                print 'No Such Code:',i[2]
                continue
            worksheet.write_merge(index_st,index_ed-1,0,0,i[0])  # 箱号
            worksheet.write_merge(index_st,index_ed-1,1,1,i[1]+u'箱')  # 纸箱
            
            for num,products in enumerate(d[i[2]]['lists']):
                # products = [u'主机',1,u'台']
                worksheet.write(index_st+num,2,products[0])  # 名称
                worksheet.write(index_st+num,3,d[i[2]]['xinghao'])  # 型号
                worksheet.write(index_st+num,4,products[2])  # 单位
                worksheet.write(index_st+num,5,products[1]*int(i[3]))  # 数量
                worksheet.write(index_st+num,6,case[0][0])  # 备注
            
    workbook.save('final.xls')    
    
if __name__ == '__main__':
    # all2json()
    # deal_data()
    l = read_from_case()
    d = read_from_sucai()
    write2xls(l,d)
