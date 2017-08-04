#-*- coding: UTF-8 -*-
from PIL import Image
import os
import re
import sys
import glob
import os
from Tkinter import *
import Tkinter
import tkFileDialog
import tkMessageBox

filetype2='.png'

def open_allfile(path,filetype):
    data=[]
    read_files=glob.glob(path+'*'+filetype)
    print read_files
    for i in read_files:
        with open(i,'rb') as infile:
            data.append(infile.read())
    return data

def get_filename(path,filetype):
    import os
    name=[]
    for root,dirs,files in os.walk(path):
        for i in files:
            if filetype in i:
                name.append(i.replace(filetype,''))
    return name

def output_result():
    base_img_width = 1024
    base_img_height = 512

    str1 = "info face=\"Arial\" size=32 bold=0 italic=0 charset="" unicode=1 stretchH=100 smooth=1 aa=1 padding=0,0,0,0 spacing=1,1 outline=0\n"
    str2 = "common lineHeight=32 base=26 scaleW=1024 scaleH=1024 pages=1 packed=0 alphaChnl=1 redChnl=0 greenChnl=0 blueChnl=0\n"
    str3 = "page id=0 file=\"fontname_0.png\"\n"
    str4 = "chars count=26\n"

    # 创建底图
    base_img = Image.new('RGBA', (1024, 1024), (255, 255, 255, 0))
    fp = open("/Users/ggod/Desktop/pic/26/fontname.fnt", 'w')
    fp.write(str1)
    fp.write(str2)
    fp.write(str3)
    fp.write(str4)

    charcount = 0

    print chr(97)

    for i in range(5):
        for j in range(6):
            charcount += 1
            if charcount > 26:
                break
            tmp_img = Image.open(ur'/Users/ggod/Desktop/pic/26/' + chr(96 + charcount) + '.png')
            img_size = tmp_img.size[0]
            box = (img_size * j, img_size * i, img_size * (j + 1), img_size * (i + 1))
            region = tmp_img
            base_img.paste(region, box)
            str_data = "char " + "id=" + str(64 + charcount) + " x=" + str(img_size * j) + " y=" + str(
                img_size * i) + " width=" + str(img_size) + " height=" + str(
                img_size) + " xoffset=0 yoffset=0 xadvance=" + str(img_size) + " page=0 chnl=15\n"
            fp.write(str_data)

    base_img.save('/Users/ggod/Desktop/pic/26/fontname_0.png')
    fp.close()

#output_result()

#UI_Ctrl
def get_directory():
    dirname = tkFileDialog.askdirectory(parent=root, initialdir="/Users/ggod/Desktop/pic", title='Pick a directory')
    data1 = open_allfile(dirname, filetype2)
    print len(data1)
    name2 = get_filename(dirname, filetype2)
    print name2
    print dirname

#UI
root = Tk()
root.title("word calculate gui")
root.geometry("1024x768")
root.resizable(width=False, height=False)
Button(root,text="ChooseDirectory",command=get_directory,height=4,width=16).grid(row = 1, column = 2, columnspan = 2, sticky=N)
root.mainloop()