import xlrd
from pyExcelerator import *
import random
fname = "vocabulary.xls"
bk = xlrd.open_workbook(fname)
shxrange = range(bk.nsheets)
try:
 sh = bk.sheet_by_name("Sheet1")
except:
 print ("no sheet in %s named Sheet1" % fname)
nrows = sh.nrows
ncols = sh.ncols
print ("nrows %d, ncols %d" % (nrows,ncols))
#获取第一行第一列数据 
cell_value = sh.cell_value(1,1)
#print cell_value
  
row_list = []
#获取各行数据
for i in range(0,nrows):
 row_data = sh.cell_value(i,0)
 row_list.append(row_data)

test_list=[] 
w = Workbook()  #创建一个工作簿
ws = w.add_sheet('Hey, Hades')  #创建一个工作表
for f in range(0,50):
    x=random.randint(0,99)
    while row_list[x] in test_list:
        x=random.randint(0,99)
    test_list.append(row_list[x])
    ws.write(f,0,row_list[x]) #在1行1列写入bit
w.save('test.xls')  #保存
