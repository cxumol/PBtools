# -*- coding: utf-8 -*-
"""
@author: cxumol
"""
# import xlrd
import re, os, time
from xlrd import open_workbook
#import traceback
#from encodings import gbk
#old_pbb=u'温暖衣冬排班表12月31日更新.xlsx'
#kkb=u'2014-2015上学期青协空课表（最终版）.xls'
maxweek = 20
errlist = []

#try:
#    sh = bk.sheet_by_name("Sheet1")
#except:
#    print "no sheet in %s named Sheet1" % old_pbbText
def getFormerkuli(f):
    #开启文件，初始化
    namelist = []
    try:
        wb = open_workbook(f)
    except:
        #tkMessageBox.showinfo('呵呵', u'旧排班表读取失败，不考虑上一次摆过摊的人的情况下进行筛选')
        return []
    for sh in wb.sheets():
        #获取行数
        h = sh.nrows
        for i in range(h):
            rr = sh.row_values(i)  #这一行读出来是个列表
            ptn = re.compile(ur'(.+?)[\(（]')  #正则匹配表达式
            #            t= open('./test.txt', 'a')
            for s in rr:
                if type(s) == unicode:  #排除无意义的单元格
                    #                    s=s.decode('utf8')
                    m = ptn.search(s)
                    #                    print m
                    #                    print 'ok'
                    if m: namelist.append(m.group(1).replace(' ', ''))  #去掉空行，添加进入名单
                #                t.write(rr)
                #                t.write('\n')
    return namelist
    #row_list.append(row_data)


def expt(f):

    namelist = []
    try:
        wb = open_workbook(f)
    except:
        return []
    for sh in wb.sheets():
        h = sh.ncols  #获取列数
        for i in range(h):
            rr = sh.col_values(i)
            for s in rr:
                if s:
                    #try:
                    namelist.append(s.replace(' ', ''))  #去掉空行，添加进入名单
                    #except:print 'f=',f
    return namelist


def getkkb(f):
    try:
        wb = open_workbook(f)
    except:
        tkMessageBox.showinfo(u'呵呵', u'空课表打开失败')
    #初始化多维数组，其中0-17表示周数，0-4表示星期，0-4表示课时,注意输入时的时间转换
    table = []
    for i in range(maxweek):
        table.append([])
        for j in range(5):
            table[i].append([])
            for k in range(5):
                table[i][j].append([])
            #    print '调试 '


    for sh in wb.sheets():
        department = sh.name
        for rowx in range(2, 7):
            for coly in range(1, 6):
                rawVal = sh.cell_value(rowx, coly)
                xq = coly - 1;
                ks = rowx - 1 - 1
                #正则提取每人的周数
                ptn = re.compile(ur'[^,，].+?[\(（].+?[\)）]')
                rawLi = ptn.findall(rawVal)
                #分析每人
                for p in rawLi:
                    p = p.replace(u'，', ',').replace('.', ',').replace(' ', '').replace(u'（', '(').replace(u'）', ')')
                    pname = re.search(ur'(.+?)[\(（]', p).group(1) + u'(%s)' % department
                    pweeks = re.search(ur'[\(（](.+?)[\)）]', p).group(1).replace(u'、', ',').split(',')
                    zsb = []#V2，改为把周数添加到周数列，便于调试,添加位置不当曾导致严重错误，这是个教训
                    for w in pweeks: #调查括号里有哪些周
                        g = w.find('-')  #g是横杠的位置
                        if g != -1:
                            if g == len(w) - 1:  #减一是因为切片特性
                                try:
                                    start = int(w[:g])
                                except UnicodeEncodeError:
                                    errlist.append(u'%s 在星期%i第%i时辰 没写数字 :“%s”' % (pname, xq + 1, ks + 1, w))
                                    continue
                                except ValueError:
                                    errlist.append(u'%s在星期%i第%i时辰的格式真心写错了，没准是多了个逗号 %s' % (pname, xq + 1, ks + 1, w))
                                    continue
                                for zs in range(start - 1, maxweek):
                                    zsb.append(zs)
                                    #print 'atzs',zs,'xq',xq,'ks',ks,pname
                            else:
                                try:
                                    start = int(w[:g])
                                    end = int(w[g + 1:])
                                except UnicodeEncodeError:
                                    errlist.append(u'%s 在星期%i第%i时辰 没写数字 :“%s”' % (pname, xq + 1, ks + 1, w))
                                    continue
                                except ValueError:
                                    errlist.append(u'%s在星期%i第%i时辰的格式真心写错了，没准是多了个逗号 %s' % (pname, xq + 1, ks + 1, w))
                                    continue
                                for zs in range(start - 1, end):
                                    #try:
                                    #table[zs][xq][ks].append(pname)
                                    zsb.append(zs)
                                    #except:
                                        #print 'zs', zs, 'xq', xq, 'ks', ks, pname
                                    #print 'zs',zs,'xq',xq,'ks',ks,pname
                        else:
                            #print pname,xq,ks,w
                            try:
                                zs = int(w)-1
                            except UnicodeEncodeError:
                                errlist.append(u'%s 在星期%i第%i时辰 没写数字 :“%s”' % (pname, xq + 1, ks + 1, w))
                                continue
                            except ValueError:
                                errlist.append(u'%s在星期%i第%i时辰的格式真心写错了，没准是多了个逗号 %s' % (pname, xq + 1, ks + 1, w))
                                continue
                            else:
                                if zs < maxweek: zsb.append(zs)
                    for zhsh in zsb:
                        table[zhsh][xq][ks].append(pname)

    return table

def getTXLdict(f):
    TXLdict = {}
    try:
        wb = open_workbook(f)
    except:
        tkMessageBox.showinfo('呵呵', u'通讯录打开失败，没法获取人物资料')
        return {}
    for sh in wb.sheets():
        #获取行数
        h = sh.nrows
        #range范围：从1（第二行）到最大行数-1（最大行数-1）
        for rowx in range(h-1):
            pname= unicode(sh.cell_value(rowx, 3)).replace(' ','')
            TXLdict[pname]=[unicode(sh.cell_value(rowx, 4)),unicode(sh.cell_value(rowx, 7))]
    return TXLdict
#=========================================================
# MAIN
#=========================================================
def main(old_pbb, kkb, zs, xq, sk, exf, txlf):
    #tiaoshi
    #tkMessageBox.showinfo(u'呵呵',u'main')
    #文件名排错
    kbok = False

    if not os.path.isfile(kkb):
        tkMessageBox.showinfo('呵呵', u'空课表文件名写错了，再检查一下。【再见】')
        return None  #退出主程序
    elif not os.path.isfile(txlf):
        kbok = True;txlok = False
        tkMessageBox.showinfo('呵呵', u'没有通讯录，没关系，按“确定”继续筛选')
    elif not os.path.isfile(old_pbb):
        kbok = True;txlok = True
        tkMessageBox.showinfo('呵呵', u'没有旧排班表，没关系，按“确定”继续筛选')
    elif not os.path.isfile(exf):
        kbok = True;txlok = True
        tkMessageBox.showinfo('呵呵', u'没有排除人员表，没关系，按“确定”继续筛选')
    else:
        kbok = True
        opbok = True
        exok = True
        txlok = True
    #有空课表
    if kbok:
        kuli = getFormerkuli(old_pbb)
        kkTable = getkkb(kkb)
        ex = expt(exf)
        for kl in kuli:
            ex.append(kl)
        txldict = getTXLdict(txlf)

        eorF = open(u'./错误记录.txt', 'a')
        eorF.write(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())) + '\n')
        eorF.write(u'空课表文件:'.encode('utf8') + kkb.encode('utf8') + "\n")
        for err in errlist:
            #print
            eorF.write(err.encode('utf8') + '\n')

        #输出的前期准备
        opt = [] #初始化输出表
        if ex: #排除的名单
            for p in kkTable[zs][xq][sk]:
                pisex = False
                for exman in ex:
                    if exman in p:
                        pisex = True
                if not pisex:
                    opt.append(p)
        else:
            for p in kkTable[zs][xq][sk]:
                opt.append(p)

        #通讯录信息加进来
        opt2=[]
        if txlok and txldict:
            for i in range(len(opt)):
                name = re.search(ur'(.+?)[\(（]', opt[i]).group(1)
                name=unicode(name)
                try:
                    opt[i]=opt[i]+','.join(txldict[name])
                except:
                    txlerr=u'通讯录中，%s找不到，或其他错误' % name
                    eorF.write(txlerr.encode('utf8') + '\n')
                    continue


        #开始导出到文件
        optF = open(u'./候选人.txt', 'a')
        optF.write(time.strftime(u'输出时间%Y-%m-%d %H:%M:%S'.encode('utf8'), time.localtime(time.time())) + '\n')
        for p in opt:
            optF.write(p.encode('utf8') + '\n')
        #optF.write('\n'.join(opt))
        optF.write(u'以上是适合在第%s周，星期%s,第%s时辰摆摊的人'.encode('utf8') % (zs + 1, xq + 1, sk + 1) + '\n' + '\n')
        optF.close()
        eorF.close()
        tkMessageBox.showinfo('哈哈', u'恭喜！筛选出来的人，已经放到“候选人txt”了。去看看呗（注意首行的输出时间）。\n P.S.别急着关闭程序。改一下时间后再按“嗯嗯”，会继续在后头增加新筛选的结果哦')

#=========================================================
# GUI
#=========================================================

import tkMessageBox

import Tkinter

root = Tkinter.Tk()
root.title(u"排班表神器")
Tkinter.Label(root, text=ur"让以下两个文件处于程序的同一目录，或者输入完整路径，如:C:\abc\xxx.xls").pack(side='top')

#空课表
e = Tkinter.Frame(root)
kkbInput = Tkinter.Entry(e, width=40)
kkbText = Tkinter.Label(e, text=u"空课表文件名（如“2014-2015上学期青协空课表（最终版）.xls”）：")
kkbText.pack(side="left", fill="x")
kkbInput.pack(side="right")
e.pack(side="top", fill="x")

#根据通讯录，找到性别和电话
txl = Tkinter.Frame(root)
txlInput = Tkinter.Entry(txl, width=40)
txlText = Tkinter.Label(txl, text=u"通讯录,（如“xxx.xls(x)”）:    可留空")
txlText.pack(side="left", fill="x")
txlInput.pack(side="right")
txl.pack(side="top", fill="x")

#排除纯名单
b = Tkinter.Frame(root)
exInput = Tkinter.Entry(b, width=40)
exText = Tkinter.Label(b, text=u"需要排除的人,只能有人名,（如“xxx.xls(x)”）:    可留空")
exText.pack(side="left", fill="x")
exInput.pack(side="right")
b.pack(side="top", fill="x")

#上次的排班表
d = Tkinter.Frame(root)
old_pbbInput = Tkinter.Entry(d, width=40)
old_pbbText = Tkinter.Label(d, text=u"根据排班表排除人员(如,温暖衣冬排班表12月31日更新.xlsx):  可留空")
old_pbbText.pack(side="left", fill="x")
old_pbbInput.pack(side="right")
d.pack(side="top", fill="x")

#本次摆摊信息
bc = Tkinter.Frame(root)
Tkinter.Label(bc, text=u"星期几写12345。注意！时刻为1表示上午12节课,2表示3-4节课,5表示晚上9-10节课,以此类推").pack(side='top')

zsInput = Tkinter.Entry(bc, width=5)
Tkinter.Label(bc, text=u"周数").pack(side="left", padx=12)
zsInput.pack(side="left", padx=12)

xqInput = Tkinter.Entry(bc, width=5)
Tkinter.Label(bc, text=u"星期").pack(side="left", padx=12)
xqInput.pack(side="left", padx=12)

skInput = Tkinter.Entry(bc, width=5)
Tkinter.Label(bc, text=u"时刻").pack(side="left", padx=12)
skInput.pack(side="left", padx=12)

bc.pack(side="top")
#确认按钮
def submit():
    old_pbb = old_pbbInput.get() or u'空白'
    kkb = kkbInput.get() or u'空白'
    exf = exInput.get() or u'空白'
    txlf = txlInput.get() or u'空白'

    bcok = False
    zs, xq, sk = zsInput.get(), xqInput.get(), skInput.get()
    try:
        zs, xq, sk = map(int, (zs, xq, sk))
    except:
        tkMessageBox.showinfo(u'呵呵', u'别乱写周数星期时刻，看清上面的提示')
    if (zs not in range(1, maxweek + 1)) or (xq not in range(1, 6)) or (sk not in range(1, 6)):
        tkMessageBox.showinfo(u'呵呵', u'数字不在有效范围')
    else:
        bcok = True

    #tkMessageBox.showinfo(u'呵呵',u'工作目录%s'%os.getcwd())    

    if bcok:
        #tkMessageBox.showinfo(u'呵呵',u'bcok')#tiaoshi

        zs, xq, sk = zs - 1, xq - 1, sk - 1
        main(old_pbb, kkb, zs, xq, sk, exf, txlf)


btn = Tkinter.Button(root, text=u'嗯嗯', command=submit)
btn.pack(side='bottom')

root.mainloop()


#print kuli
#!!!f = open("out.html","w",encoding='utf-8')!!!