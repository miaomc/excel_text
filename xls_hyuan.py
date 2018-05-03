# -*- coding:cp936 -*-
import glob
import xlrd
import xlwt
import json
    
def all2json():
    """�����*.xls�ļ�����ȡ��һ��products���б��У�Ȼ��洢��all.json��"""
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
    """��json�е��ļ���ȫ��д��all.xls����У��ֿ�������Ϊall2json()��ȡ�ܶ��ļ��Ƚ�����"""
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
    """������tmp.startswith(head)�滻"""
    return tmp[0:len(head)] == head
        
def read_from_case():
    """������������xls�ļ��������б�lists"""
    # read from xls
    data = xlrd.open_workbook('all_origin.xls')
    table = data.sheets()[0]
    lines= []
    for j in range(1, table.nrows):
        tmp = table.row_values(j)
        lines.append(tmp)

    print len(lines)

    # deal data ---->[  [[����װָ�..����...],['1-R09','ֽ'�������롯����ʵ������],['1-R09','ֽ'�������롯����ʵ������],...], ... ]
    lists = []
    one_case = False
    for i in lines:
        if head_with(u'��װָ��',i[0]):
            baozhuang = i[0]
        elif head_with(u'װ�䵥��',i[0]):
            zhuangxiang = i[2]
        elif head_with(u'����',i[0]):
            one_case = True
            tmp = [[baozhuang,zhuangxiang]]
        elif head_with(u'��ע',i[0]):
            one_case = False
            lists.append(tmp)
        elif one_case:
            tmp.append([i[0],i[1],i[3],i[7]])

    return lists
            
def read_from_sucai():
    """������������xls�ļ��������б�lists"""
    # read from xls
    data = xlrd.open_workbook('sucai.xls')
    table = data.sheets()[0]
    lines= []
    for j in range(1, table.nrows):
        tmp = table.row_values(j)
        lines.append(tmp)

    # create data ---> {'BOM1':{'xinghao':'','lists':[[u'����',1,u'̨'],[],[]]},'BOM2':{},...}
    dirs = {}
    for i in lines:
        #print repr(i[0])
        if len(i[0])>3:  # �ų��ո�ͱ���
            dirs[i[0]] = {'xinghao':i[1],'lists':[[i[2],int(i[3]),i[4]]]}
            key = i[0]
        elif i[0] != 'BOM':
            dirs[key]['lists'].append([i[2],int(i[3]),i[4]])

    return dirs
            
def write2xls(l,d):
    """����д��"""
    workbook = xlwt.Workbook(encoding = 'ascii')
    worksheet = workbook.add_sheet(u'���ϵͳ')

    # l = [  [[����װָ�..����...],['1-R09','ֽ'����BOM���롯����ʵ������],['1-R09','ֽ'�������롯����ʵ������],...], ... ]
    # d = {'BOM1':{'xinghao':'','lists':[[u'����',1,u'̨'],[],[]]},'BOM2':{},...}

    index_st = -1
    index_ed = -1
    for case in l:
        # case =  [[����װָ�..����...],['1-R09','ֽ'����BOM���롯����ʵ������],['1-R09','ֽ'�������롯����ʵ������],...]
        for i in case[1:]:
            # i = ['1-R09','ֽ'����BOM���롯����ʵ������]
            # from index_start to index_end
            index_st = index_ed + 1  # �˴���֤�ɿ�һ��
            if i[2] in d.keys():
                index_ed = index_st + len(d[i[2]]['lists'])  # [index_st, index_ed )
            else:
                print 'No Such Code:',i[2]
                continue
            worksheet.write_merge(index_st,index_ed-1,0,0,i[0])  # ���
            worksheet.write_merge(index_st,index_ed-1,1,1,i[1]+u'��')  # ֽ��
            
            for num,products in enumerate(d[i[2]]['lists']):
                # products = [u'����',1,u'̨']
                worksheet.write(index_st+num,2,products[0])  # ����
                worksheet.write(index_st+num,3,d[i[2]]['xinghao'])  # �ͺ�
                worksheet.write(index_st+num,4,products[2])  # ��λ
                worksheet.write(index_st+num,5,products[1]*int(i[3]))  # ����
                worksheet.write(index_st+num,6,case[0][0])  # ��ע
            
    workbook.save('final.xls')    
    
if __name__ == '__main__':
    # all2json()
    # deal_data()
    l = read_from_case()
    d = read_from_sucai()
    write2xls(l,d)
