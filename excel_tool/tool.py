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
sheet1 = f.add_sheet(u'Word_1', cell_overwrite_ok=True)
sheet2 = f.add_sheet(u'Word_2', cell_overwrite_ok=True)
sheet3 = f.add_sheet(u'Word_3', cell_overwrite_ok=True)
sheet4 = f.add_sheet(u'Word_4', cell_overwrite_ok=True)
sheet5 = f.add_sheet(u'Word_5', cell_overwrite_ok=True)


row0 = [u'WorldIndex', u'SubWorldIndex', u'LevelIndex', u'Word', u'Answers', u'Valid Words', u'AloneWord', u'Layout']
for i in range(0, len(row0)):
    sheet0.write(0, i, row0[i])
    sheet1.write(0, i, row0[i])
    sheet2.write(0, i, row0[i])
    sheet3.write(0, i, row0[i])
    sheet4.write(0, i, row0[i])
    sheet5.write(0, i, row0[i])

def split_words(word_list=[]):
    cross_word_list = []
    new_list = word_list.split('|')
    for word in new_list:
        if word != "":
            cross_word_list.append([word])
    return cross_word_list

def writedata(ptrsheet, count, data = []):
    for i in range(len(data)):
        ptrsheet.write(count, i+3, data[i])

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
            data1 = base.read_cell(i, 3)
            data2 = base.read_cell(i, 4)
            data3 = base.read_cell(i, 5)
            data4 = base.read_cell(i, 6)
            data5 = base.read_cell(i, 7)
            if data3:
                wordlist = split_words(data3)
                data3 = wordlist[0][0] + "|" + wordlist[1][0] + "|"
            count += 1
            if index == 0:
                sheet0.write(count, 0, "World_0")
                sheet0.write(count, 1, "SubWorld_" + str(subwordindex))
                if not word0tag:
                    sheet0.write(count, 2, "Level_" + str(levelindex) + ".asset")
                    levelindex += 1
                    if levelindex > 11:
                        subwordindex += 1
                        levelindex = 0
                        word0tag = 1
                else:
                    sheet0.write(count, 2, "Level_" + str(levelindex) + ".asset")
                    levelindex += 1
                    if levelindex > 17:
                        subwordindex += 1
                        levelindex = 0
                writedata(sheet0, count, [data1, data2, data3, data4, data5])
                if count ==84:
                    index += 1
                    count = 0
                    subwordindex = 0
                    levelindex = 0
            elif index == 1:
                sheet1.write(count, 0, "World_1")
                sheet1.write(count, 1, "SubWorld_" + str(subwordindex))
                sheet1.write(count, 2, "Level_" + str(levelindex) + ".asset")
                levelindex += 1
                if levelindex > 17:
                    subwordindex += 1
                    levelindex = 0
                writedata(sheet1, count, [data1, data2, data3, data4, data5])
                if count == 90:
                    index += 1
                    count = 0
                    subwordindex = 0
                    levelindex = 0
            elif index == 2:
                sheet2.write(count, 0, "World_2")
                sheet2.write(count, 1, "SubWorld_" + str(subwordindex))
                sheet2.write(count, 2, "Level_" + str(levelindex) + ".asset")
                levelindex += 1
                if levelindex > 17:
                    subwordindex += 1
                    levelindex = 0
                writedata(sheet2, count, [data1, data2, data3, data4, data5])
                if count == 90:
                    index += 1
                    count = 0
                    subwordindex = 0
                    levelindex = 0
            elif index == 3:
                sheet3.write(count, 0, "World_3")
                sheet3.write(count, 1, "SubWorld_" + str(subwordindex))
                sheet3.write(count, 2, "Level_" + str(levelindex) + ".asset")
                levelindex += 1
                if levelindex > 17:
                    subwordindex += 1
                    levelindex = 0
                writedata(sheet3, count, [data1, data2, data3, data4, data5])
                if count == 90:
                    index += 1
                    count = 0
                    subwordindex = 0
                    levelindex = 0
            elif index == 4:
                sheet4.write(count, 0, "World_4")
                sheet4.write(count, 1, "SubWorld_" + str(subwordindex))
                sheet4.write(count, 2, "Level_" + str(levelindex) + ".asset")
                levelindex += 1
                if levelindex > 17:
                    subwordindex += 1
                    levelindex = 0
                writedata(sheet4, count, [data1, data2, data3, data4, data5])
                if count == 90:
                    index += 1
                    count = 0
                    subwordindex = 0
                    levelindex = 0
            elif index == 5:
                sheet5.write(count, 0, "World_5")
                sheet5.write(count, 1, "SubWorld_" + str(subwordindex))
                sheet5.write(count, 2, "Level_" + str(levelindex) + ".asset")
                levelindex += 1
                if levelindex > 17:
                    subwordindex += 1
                    levelindex = 0
                writedata(sheet5, count, [data1, data2, data3, data4, data5])


    copyworkbook.save(filepath)

    print result
    f.save("/Users/ggod/Desktop/wordxls/demo1.xls")
    print "over"


def ShowMessageBox(message):
	tkMessageBox.showinfo("notice", message)

root = Tk()
root.title("word calculate gui")
root.geometry("1024x768")
root.resizable(width=False, height=False)
Button(root,text="OpenFile",command=get_file,height=2,width=8).grid(row = 1, column = 2, columnspan = 2, sticky=N)
root.mainloop()