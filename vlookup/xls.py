#-*- coding:utf-8 -*-
import xlrd
import xdrlib,sys

data = xlrd.open_workbook('name.xlsx')
table = data.sheets()[0]
nrows = table.nrows
ncols = table.ncols

#map_nan 为得到的对比数据
map_man= [[0 for i in range(20)] for i in range(20)]

#map_women得到的女性对比数据
map_women=[[0 for i in range(20)]for i in range(20)]


#得到第6列的性别
for i in range(1,nrows):
    xingbie = table.row_values(i)[5]

#对第二体重的数据进行取整（例如：5,10,15...）
for i in range(1,nrows):
    tizhong=0
    tizhong=table.row_values(i)[1]
    if tizhong=='':
        tizhong=60 # 没有体重，默认 60kg
    if(tizhong %5==0):
        tizhong=float(tizhong)
    else:
        tizhong=float((int(tizhong)/int(5) +1)*5)
    table.put_cell(i, 1, 2, tizhong, 0)

#对第四列身高的数据进行取整（例如：5,10,15...）
for i in range(1,nrows):
    shengao=table.row_values(i)[3]

    if  shengao == '':
        shengao = 170.0  # 没有身高，默认 170
    if(not isinstance(shengao, float) ):
        shengao=170.0
    if (  shengao % 5 == 0):
        shengao = float(shengao)
    else:
        shengao = float((int(shengao /5 + 1) * 5))
    table.put_cell(i, 3, 2,  shengao, 0)

#对比对的男性别数据进行存储
for i in range(9):

    for j in range(12):
        shuju=table.row_values(i+2)[j+9]
        map_man[i][j]=shuju

#对比对的女性别的数据进行存储
for i in range(6):

    for j in range(8):
        shuju = table.row_values(i + 2)[j + 23]
        map_women[i][j]=shuju

chima=[]
#得到第6列的性别
for i in range(1,nrows):
    xingbie = table.row_values(i)[5]
    if xingbie==u'\u7537': #xingbie=="男"
        shengao=table.row_values(i)[3]
        tizhong=table.row_values(i)[1]
        if tizhong>105.0:
            tizhong=105.0
        chima.append(map_man[int((shengao-160)/5)][int((tizhong-50)/5)])

    else:  #xingbie=="女"
        shengao = table.row_values(i)[3]
        tizhong = table.row_values(i)[1]
        if tizhong > 75.0:
            tizhong = 75.0
        chima.append(map_women[int((shengao - 150) / 5)][int((tizhong - 40) / 5)])
#得到最终的数据
for i in range(len(chima)):
    name=table.row_values(i+1)[0]
    print name,chima[i]
