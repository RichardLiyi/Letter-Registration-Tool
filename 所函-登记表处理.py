from tkinter import *
from tkinter import messagebox
from datetime import datetime
import tkinter as tk
from mailmerge import MailMerge
# import winreg
import os
import sys
window = tk.Tk()

#以下代码是用于将窗口居中
# 获取屏幕的宽度和高度
screen_width = window.winfo_screenwidth()
screen_height = window.winfo_screenheight()
# 计算窗口的左上角坐标
x = (screen_width - 800) // 2
y = (screen_height - 600) // 2
# 将窗口移到该坐标
window.geometry("800x600+{}+{}".format(x, y))


#设置标题及窗口大小
window.title("所函-登记表处理工具")
window.geometry("500x250")

#全局变量定义
#定义当前py文件的路径
dir_path = os.path.dirname(os.path.abspath(sys.argv[0]))
letterNumber = ''
zoneName = ''
counterpart = ''
clientName = ''
briefCase = ''
month = ''
day = ''
letter_template_Path = ''
registrationForm_template_Path = ''
resultFolder_Path = ''
caseMoney = ''
folderPath = ''

# key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
# desktop_path = winreg.QueryValueEx(key, "Desktop")[0]

# letter表示律师函。letterNumber表示函号
label_letterNumber = Label(window, text="请输入函号：")
label_letterNumber.grid(column=0, row=0)
txt_letterNumber = Entry(window, width=40)
txt_letterNumber.grid(column=1, row=0)

#zoneName即处理单位
label_zoneName = Label(window, text="请输入处理机关：")
label_zoneName.grid(column=0, row=1)
txt_zoneName = Entry(window, width=40)
txt_zoneName.grid(column=1, row=1)

#counterpart表示被委托人
label_counterpart = Label(window, text="对方当事人姓名：")
label_counterpart.grid(column=0, row=2)
txt_counterpart = Entry(window, width=40)
txt_counterpart.grid(column=1, row=2)


#client表示委托人
label_clientName = Label(window, text="委托人姓名：")
label_clientName.grid(column=0, row=3)
txt_clientName = Entry(window, width=40)
txt_clientName.grid(column=1, row=3)

#briefCase 案件概况
label_briefCase = Label(window, text="案件概况：")
label_briefCase.grid(column=0, row=4)
txt_briefCase = Entry(window, width=40)
txt_briefCase.grid(column=1, row=4)

#caseMoney表示涉案标的
label_caseMoney = Label(window, text="涉案标的：")
label_caseMoney.grid(column=0, row=5)
txt_caseMoney = Entry(window, width=40)
txt_caseMoney.grid(column=1, row=5)

#提交按钮的click事件
def submitClicked():
    global letterNumber, zoneName, counterpart, clientName, briefCase, month, day,folderPath,caseMoney,dir_path
    letterNumber = txt_letterNumber.get()
    zoneName = txt_zoneName.get()
    counterpart = txt_counterpart.get()
    clientName = txt_clientName.get()
    briefCase = txt_briefCase.get()
    caseMoney = txt_caseMoney.get()
    now = datetime.now()
    month = str(now.month) + "月"
    day = str(now.day)+"日"
    folderPath = dir_path+"\\"+clientName+"所函-登记表"
    os.makedirs(folderPath, exist_ok=True)
    messagebox.showinfo("提示", "您提交的数据为：\n"
                              "函号：%s\n"
                              "处理机关：%s\n"
                              "对方当事人%s\n"
                              "委托人：%s\n"
                              "案件概况：%s\n"
                              "涉案标的：%s\n"
                              "系统将创建相应委托人文件夹！%s" %(letterNumber,zoneName,counterpart,clientName,briefCase,caseMoney,folderPath))
#提交按钮
btn_submit = Button(window, text="提交表单", command=submitClicked)
btn_submit.grid(column=1, row=6)

# #letter_template_Path表示律师函模板路径
# label_letter_template_Path = Label(window, text="律师函模板路径:")
# label_letter_template_Path.grid(column=0, row=7)
# txt_letter_template_Path= Entry(window, width=40)
# txt_letter_template_Path.grid(column=1, row=7)
# txt_letter_template_Path.insert(0, "D:\onedrive\桌面\立案文件\所函模板.docx")
#
# #registrationForm_template_Path表示登记表模板路径
# label_registrationForm_template_Path = Label(window, text="收案登记表模板路径:")
# label_registrationForm_template_Path.grid(column=0, row=8)
# txt_registrationForm_template_Path= Entry(window, width=40)
# txt_registrationForm_template_Path.grid(column=1, row=8)
# txt_registrationForm_template_Path.insert(0, "D:\onedrive\桌面\立案文件\收案登记表模板.docx")

#文件存储文件夹路径 resultFolder_Path
# label_resultFolder_Path = Label(window, text="生成文件存储文件夹路径:")
# label_resultFolder_Path.grid(column=0, row=9)
# txt_resultFolder_Path= Entry(window, width=40)
# txt_resultFolder_Path.grid(column=1, row=9)




def writeClicked():
    global letter_template_Path,registrationForm_template_Path,resultFolder_Path
    letter_template_Path = os.path.join(dir_path, "所函模板.docx")
    registrationForm_template_Path = os.path.join(dir_path, "收案登记表模板.docx")
    #导入所函模板
    letter_template = MailMerge(letter_template_Path)
    letter_template.merge(
        案件情况=briefCase,
        对方当事人姓名=counterpart,
        委托人姓名=clientName,
        某月=month,
        某日=day,
        处理机关=zoneName,
        函号=letterNumber,
    )

    letter_template.write(os.path.join(folderPath,clientName+"所函.docx"))
    #导入收案登记表模板
    registrationForm_template = MailMerge(registrationForm_template_Path)
    registrationForm_template.merge(
        案件情况= briefCase,
        对方当事人姓名=counterpart,
        委托人姓名=clientName,
        某月=month,
        某日=day,
        处理机关=zoneName,
        涉案标的=caseMoney
    )
    registrationForm_template.write(os.path.join(folderPath,clientName+"收案登记表.docx"))
    messagebox.showinfo("提示", "写入成功！请前往对应文件夹查看")

btn_write = Button(window, text="写入文件", command=writeClicked)
btn_write.grid(column=1, row=10)

# def createFolder():
#
# btn_write = Button(window, text="创建文件夹", command=createFolder)
# btn_write.grid(column=1, row=11)

#窗口进入运行
window.mainloop()