# -*- encoding:cp936 -*-
import json

def select(tmp):
    if tmp:
        return int(tmp)-1
    else:
        return 0
    
def main():
##    project_name = raw_input(u'项目名称:')  or u'项目'
##    # 补充:自动扩充V1.00;空就自动添加当天日期和编号

    with open('head_products.json','r') as f1:
        head, products = json.load(f1)

    chanpin_index = head.index(u'产品型号')

    lst = []
    tmp = raw_input('Add Device?:')
    while tmp:
        key = tmp.split()[0]

        # 进行产品型号确定，根据产品型号后面的一个数据来确认（一般是@配置）
        choice_list = []
        for i in products:
            if key in i[chanpin_index]:
                choice_list.append(i)
        for n,i in enumerate(choice_list):
            print '%s: %s @ %s'%(n+1,i[chanpin_index],i[chanpin_index+1])

        # 进行产品个数
        if tmp.split()[1:2]:
            num = int(tmp.split()[1])
        else:
            num = 1

        # 统一把所有信息录入到lst列表中去
        if choice_list:  # 假如输入的型号不存在就记录一个输入型号的空列表
            lst.append([choice_list[select(raw_input("Which One?:"))],num])
        else:
            tmp_list = ['']*len(head)
            tmp_list[head.index(u'产品型号')] = key
            lst.append([tmp_list,num])
            
        tmp = raw_input('Add Device?:')
        
    # 将 i[1]这个数量，整到列表中去
    shuliang_index = head.index(u'数量')
    num2lst = [i[0][0:shuliang_index]+[i[1]]+i[0][shuliang_index+1:] for i in lst]

    for i in num2lst:
        print('%s @ %s, %s'%(i[2],i[3],i[6]))
    return head,num2lst

if __name__ == '__main__':
    main()
