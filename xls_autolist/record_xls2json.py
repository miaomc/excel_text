# -*- coding: cp936 -*-
import json
import xlrd

def read_from_xls():
    
    data = xlrd.open_workbook('base.xls')
    table = data.sheets()[0]

    head = table.row_values(0)
    if len(set(head)) != len(head):
        print(u'在参数读入的excel中，表头有重复的值！' )

    products = []  # content all the products' infomation
    for i in range(1, table.nrows):
        tmp = table.row_values(i)
        products.append(tmp)

    return head, products

def main():
    head, products = read_from_xls()

    with open('head_products.json','w') as f1:
        json.dump([head, products],fp=f1)

##    with open('head_products.json','r') as f1:
##        h,p = json.load(f1)
##
##    for i in p:
##        for j in  i:
##            print j,
##        print

if __name__ == '__main__':
    main()
