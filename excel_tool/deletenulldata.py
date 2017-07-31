import BasePackage
base = BasePackage.myexcelbase

from Tkinter import *
import tkFileDialog
import tkMessageBox
import xlwt
import xlrd
from xlutils.copy import copy
import json

filepath = ""

f = xlwt.Workbook()
sheet0 = f.add_sheet(u'Word_0', cell_overwrite_ok=True)

row0 = [u'WorldIndex', u'SubWorldIndex', u'LevelIndex', u'Word', u'Answers', u'Valid Words', u'AloneWord', u'Layout']
for i in range(0, len(row0)):
    sheet0.write(0, i, row0[i])

def split_words(word_list=[]):
    cross_word_list = []
    new_list = word_list.split('|')
    for word in new_list:
        if word != "":
            cross_word_list.append([word])
    return cross_word_list

def get_file():
    global filepath
    count = 0
    index = 0
    subwordindex = 0
    levelindex = 0
    word0tag = 0
    filepath = tkFileDialog.askopenfilename(filetypes=[("text file", "*.xls")])
    print filepath
    base.load_excel_workbook(filepath)
    base.load_sheet(0)
    base.set_write_sheet(0)
    result = base.read_cell(2, 7)
    result = json.loads(result)

    excelInstance = base.Singleton()
    sheet = excelInstance.sheet
    writesheet = excelInstance.write_sheet
    copyworkbook = excelInstance.copyworkbook
    nums = sheet.nrows

    for i in range(nums):
        if i < 1:
            continue
        data = base.read_cell(i, 7)
        if data:
            data1 = base.read_cell(i, 0)
            data2 = base.read_cell(i, 4)
            data3 = base.read_cell(i, 6)
            data4 = base.read_cell(i, 7)
            data5 = base.read_cell(i, 9)
            count += 1
            sheet0.write(count, 3, data1)
            sheet0.write(count, 4, data2)
            sheet0.write(count, 5, data5)
            sheet0.write(count, 6, data3)
            sheet0.write(count, 7, data4)

    copyworkbook.save(filepath)

    f.save("/Users/ggod/Desktop/wordxls/finishopt.xls")
    print "over"


def ShowMessageBox(message):
	tkMessageBox.showinfo("notice", message)

root = Tk()
root.title("word calculate gui")
root.geometry("1024x768")
root.resizable(width=False, height=False)
Button(root,text="OpenFile",command=get_file,height=2,width=8).grid(row = 1, column = 2, columnspan = 2, sticky=N)
root.mainloop()