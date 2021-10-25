#import numpy as np
import xlrd
import xlwt
import itertools
from itertools import combinations
from itertools import permutations
from xlutils.copy import copy
from xlwt.Workbook import Workbook

excel_name = 'C:/excel/导出表格.xls'
sheet_name = '测试数据'
title = ['测试步骤']

def array_int():
    n = int(input("请输入变量数："))
    arr=["tmp"]*n
    for num in range(0,n):
        print("请输入变量",num+1,"--记得输入分隔符(空格，逗号等)--")
        arr[num]=input()
    print("----------------------")

    #o1=list(combinations(arr,2))
    num2 = int(input("输出个数："))
    o2 = list(permutations(arr,num2))
    #print(o1)
    return o2

#新建表格
def excel_int(path, sheet_name):
    workbook = xlwt.Workbook() #新建一个工作簿
    workbook.add_sheet(sheet_name) #在工作簿中新建一个表格
    workbook.save(path) #保存工作簿
    print("新建表格成功，表格名称：",path)

""" #写入表头
def excel_write_title(path, titles):
    workbook = xlrd.open_workbook(path) #打开工作簿
    new_workbook = copy(workbook) #将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0) #获取转化后工作簿中的第一个表格
    for j in range(0, len(titles)):
        new_workbook.write(0, j, str(titles[j])) #表格中写入对应的数据（对应的行）
    new_workbook.save(path)
    print("表头写入成功!") """

#向表格按列写入一维数组
def excel_write_array(path, value, column):
    workbook = xlrd.open_workbook(path) #打开工作簿
    new_workbook = copy(workbook) #将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0) ## 获取转化后工作簿中的第一个表格
    for i in range(0, len(value)):
        # 向表格中写入数据（对应的列），初始位置加1（因为有表头）
        new_worksheet.write(i, column, value[i])
    new_workbook.save(path)
    print("数组写入成功!")

print("---请确认已在C盘根目录创建名为excel的文件夹---")

arr = array_int()
excel_int(excel_name, sheet_name)
#excel_write_title(excel_name, title)
excel_write_array(excel_name, arr, 0)

end=input("---输入任意字符结束程序---")