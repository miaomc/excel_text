# -*- coding:cp936 -*-
"""
��û�б�����У����뵽�����б��������ȥ��
1.����ÿһ�У������һ�л��ߵڶ��������ݾͷŵ�lst_res��ȥ
2.���򣬾Ͱ����ӵ���һ����ȥ
"""

# ��csv�ļ�
with open('11.csv','r') as f1:
    all = f1.read()

#for i in all.split('\n'):
#    print i

# ���˳�csv�ļ��е�ÿһ��
l = all.split('\n')

lst_res = []

# ����ÿһ�У������һ�л��ߵڶ��������ݾͷŵ�lst_res��ȥ
# ���򣬾Ͱ����ӵ���һ����ȥ
for line in l:
    tmp_line = line.split(',')
    
    if len(tmp_line) > 1 and (tmp_line[0] or tmp_line[1]):
        lst_res.append(line)
        
    elif len(tmp_line) > 3 and tmp_line[3]:
        new_lst_last_line = lst_res[-1].split(',')
        new_lst_last_line[3] += '\t' + 
        lst_res[-1] = ','.join(new_lst_last_line) + tmp_line[3]

    else:
        lst_res.append(line)

    #print '==',lst_res[-1]

# ��lst_res�洢����һ��csv�ļ���
with open('222.csv','w') as f1:
    for i in lst_res:
        f1.write(i+'\n')
