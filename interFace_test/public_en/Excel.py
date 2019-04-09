# -*- coding: utf-8 -*-
# @Time    : 2019/1/6  12:27
# @Author  : MrLu！！
# @FileName: Excel.py
# @Software: PyCharm

# from common.Log import run_log as logger
import xlrd

class OperationExcel():

    def __init__(self,fileName,sheetName="Sheet1"):
        self.data = xlrd.open_workbook(fileName)
        self.table = self.data.sheet_by_name(sheetName)
    # 获取第一行内容作为key值
        self.keys = self.table.row_values(0)

    #获取总行数
        self.rowNum = self.table.nrows

    #获取总列数
        self.colNum = self.table.ncols
