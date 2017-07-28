#-*- coding: UTF-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import xlwt
import xlrd
from xlutils.copy import copy

PY2 = sys.version_info[0] == 2
if PY2:
    import codecs
    from functools import partial
    open = partial(codecs.open, encoding='utf-8')

#对应excel内的各种单例
class Singleton(object):
	__instance=None
	def __init__(self):
		pass
	def __new__(cls,*args,**kwd):
		if Singleton.__instance is None:
			Singleton.__instance=object.__new__(cls,*args,**kwd)
		return Singleton.__instance

#加载excel文件
def load_excel_workbook(filepath):

    excelInstance = Singleton()
    excelInstance.workbook = xlrd.open_workbook(filepath)
    excelInstance.copyworkbook = copy(excelInstance.workbook)

#加载对应sheet
def load_sheet(sheettag):
    excelInstance = Singleton()
    if isinstance(sheettag, int):
        excelInstance.sheet = excelInstance.workbook.sheet_by_index(sheettag)
    elif isinstance(sheettag, str):
        excelInstance.sheet = excelInstance.workbook.sheet_by_name(sheettag)
    else:
        return

#设置写入sheet
def set_write_sheet(index):
    excelInstance = Singleton()
    excelInstance.write_sheet = excelInstance.copyworkbook.get_sheet(index)

#读取行数据
def read_row(index):
    excelInstance = Singleton()
    result_value = excelInstance.sheet.row_values(index)
    if result_value:
        return result_value
    else:
        return

#读取列数据
def read_col(index):
    excelInstance = Singleton()
    result_value = excelInstance.sheet.col_values(index)
    if result_value:
        return result_value
    else:
        return

#读取单元格数据
def read_cell(row, col):
    excelInstance = Singleton()
    result_value = excelInstance.sheet.cell(row, col).value
    if result_value:
        return result_value
    else:
        return

# load_excel_workbook("/Users/ggod/Desktop/wordxls/WORDTOEXL.xls")
# load_sheet(1)
# result = read_row(3)
# print result