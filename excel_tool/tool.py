import BasePackage
base = BasePackage.myexcelbase

from Tkinter import *
import tkFileDialog
import tkMessageBox

filepath = ""

def get_file():
    global filepath
    filepath = tkFileDialog.askopenfilename(filetypes=[("text file", "*.xls")])
    print filepath
    base.load_excel_workbook(filepath)
    base.load_sheet(1)
    result = base.read_row(3)
    print result


def ShowMessageBox(message):
	tkMessageBox.showinfo("notice", message)

root = Tk()
root.title("word calculate gui")
root.geometry("1024x768")
root.resizable(width=False, height=False)
Button(root,text="OpenFile",command=get_file,height=2,width=8).grid(row = 1, column = 2, columnspan = 2, sticky=N)
root.mainloop()