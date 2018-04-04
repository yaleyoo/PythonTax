# -*- coding: utf-8 -*- #
from openpyxl import load_workbook, Workbook
import Tkinter
import tkFileDialog 
from tkFileDialog import *
from   Tkinter import *
import main
from main import *
import child
from child import *

def checking():
	file = text.get()
	wb = load_workbook(file,data_only=True)
	main_sheet = wb.get_sheet_by_name(u'A100000主表')
	detailed_sheet = wb.get_sheet_by_name(u'A107040减免所得税优惠')
	report = open("report.txt",'w')
	checking_main(wb,report)
	checking_detailed(wb,report)
	report.close()

def get_file():
    global filename
    #创建文件对话框,只打开txt类型文件
    filename = tkFileDialog.askopenfilename(filetypes=[("text file", "*.xlsx")])
    text.delete(0,END)
    text.insert(10,filename)

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
 
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
 
    return False


if __name__ == '__main__':
	window = Tkinter.Tk()
	window.title("纳税申报表自动填写工具")
	window.geometry('290x160')

	text = Entry(window)
	text.pack()
	button = Button(window,text="打开文件",command=get_file,height=2,width=8)
	button.pack(side=LEFT)
	button1 = Button(window, text="确认",command=checking,height=2,width=8)
	button1.pack(side=RIGHT)

	# 进入消息循环
	window.mainloop()
