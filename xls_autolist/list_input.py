# -*- encoding:cp936 -*-
import json

def select(tmp):
    if tmp:
        return int(tmp)-1
    else:
        return 0
    
def main():
##    project_name = raw_input(u'��Ŀ����:')  or u'��Ŀ'
##    # ����:�Զ�����V1.00;�վ��Զ���ӵ������ںͱ��

    with open('head_products.json','r') as f1:
        head, products = json.load(f1)

    chanpin_index = head.index(u'��Ʒ�ͺ�')

    lst = []
    tmp = raw_input('Add Device?:')
    while tmp:
        key = tmp.split()[0]

        # ���в�Ʒ�ͺ�ȷ�������ݲ�Ʒ�ͺź����һ��������ȷ�ϣ�һ����@���ã�
        choice_list = []
        for i in products:
            if key in i[chanpin_index]:
                choice_list.append(i)
        for n,i in enumerate(choice_list):
            print '%s: %s @ %s'%(n+1,i[chanpin_index],i[chanpin_index+1])

        # ���в�Ʒ����
        if tmp.split()[1:2]:
            num = int(tmp.split()[1])
        else:
            num = 1

        # ͳһ��������Ϣ¼�뵽lst�б���ȥ
        if choice_list:  # ����������ͺŲ����ھͼ�¼һ�������ͺŵĿ��б�
            lst.append([choice_list[select(raw_input("Which One?:"))],num])
        else:
            tmp_list = ['']*len(head)
            tmp_list[head.index(u'��Ʒ�ͺ�')] = key
            lst.append([tmp_list,num])
            
        tmp = raw_input('Add Device?:')
        
    # �� i[1]��������������б���ȥ
    shuliang_index = head.index(u'����')
    num2lst = [i[0][0:shuliang_index]+[i[1]]+i[0][shuliang_index+1:] for i in lst]

    for i in num2lst:
        print('%s @ %s, %s'%(i[2],i[3],i[6]))
    return head,num2lst

if __name__ == '__main__':
    main()
