# -*- coding: utf-8 -*-
"""
repeat assistant for PaiBanBiao

@author: cxumol
"""
import re
from xlrd import open_workbook

# init
pbb_l = []
d=dict()
#show info
print u'''请在 这个文件夹的 待查列表.txt 里，写上
待查重复的排班表的文件名。每个文件名占一行\n
如果用windows自带记事本修改过，也许造成第一行作废
如果遇到这种情况，第一行不写或者随便写点什么，从第二行开始写得了
---------------------'''.encode('gbk')


# Get pbb file names
to_name = u'待查列表.txt'
try:
    to_file = open(to_name,'r')
    print u'已找到 待查列表.txt 获取待查重复的排班\n'.encode('gbk')
except:
    print u'未找到 待查列表.txt 将要退出'.encode('gbk')
    wait =input("Prease <enter>")
    quit()

to_f = open(to_name,'r')

for line in to_f : 
    pbb_l.append(line.decode('utf8').strip('\n'))

# form the dict
for f in pbb_l :
    try:
        wb = open_workbook(f)
    except Exception as e:
        print u'文件 %s 打开错误!已忽略/n'.encode('gbk') %f.encode('gbk'),e
        continue
    print u'从 %s 中找重复的人'.encode('gbk') %f.encode('gbk')
    for sh in wb.sheets():
#        print sh.name
        h = sh.nrows #行数
        l = sh.ncols #列数
        for i in range(h):
            for j in range(l):
                #rr = sh.row_values(i)  #这一行读出来是个列表
                ptn = re.compile(ur'(.+?)[\(（]')  #正则匹配表达式
                #            t= open('./test.txt', 'a')
                s = sh.cell(i,j).value
                if type(s) == unicode:  #排除无意义的单元格
                    m = ptn.search(s)
                    if m: 
                        try:
                            d[m.group(1)].append((f,sh.name,chr(j+65),i+1))
#                            print '2nd'
                        except:
                            d[m.group(1)] = [(f,sh.name,chr(j+65),i+1)]
#                            print '1st',m
wait =raw_input(u"敲回车以继续".encode('gbk'))
print '---------------'
for name in d:
    if len(d[name]) > 1 :
        for loca in d[name]:
            print name.encode('gbk'),
            for lo in loca:
                if type(lo) == unicode : lo=lo.encode('gbk')
                print lo,
                
            print
            
print '---------------'
wait = raw_input(u'以上。'.encode('gbk'))
