# -*- coding:cp936 -*-
"""
把没有标题的行，合入到上面有标题的行中去：
1.处理每一行，如果第一列或者第二列有内容就放到lst_res中去
2.否则，就把他加到上一行中去
"""

# 打开csv文件
with open('11.csv','r') as f1:
    all = f1.read()

#for i in all.split('\n'):
#    print i

# 过滤出csv文件中的每一行
l = all.split('\n')

lst_res = []

# 处理每一行，如果第一列或者第二列有内容就放到lst_res中去
# 否则，就把他加到上一行中去
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

# 把lst_res存储到另一个csv文件中
with open('222.csv','w') as f1:
    for i in lst_res:
        f1.write(i+'\n')
