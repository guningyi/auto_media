#!/usr/bin/python
# -*- coding: UTF-8 -*-

from tkinter import *
from tkinter.filedialog import askopenfilename
import xlrd


class UnionRobot(object):
    def __init__(self, proxy_list=[], web_site_list=[]):
        self.proxy_list = proxy_list
        self.web_site_list = web_site_list
        self.click_count = StringVar()


    def read_xls_file(self):
        openfilename = askopenfilename(filetypes=[('xls', '*.xls')])
        wb = xlrd.open_workbook(openfilename)
        sh = wb.sheet_by_index(0)  # 第一个表
        nrows = sh.nrows  # 获取一共有多少行
        for i in range(nrows):  # 遍历输出
            temp = sh.cell(i, 0).value.split('@')
            self.proxy_list.append(temp[0])
        print(self.proxy_list)

    def read_xls(self, proxyfile_xls):
        wb = xlrd.open_workbook(proxyfile_xls)
        sh = wb.sheet_by_index(0)  # 第一个表
        proxy_list = []
        nrows = sh.nrows  # 获取一共有多少行
        for i in range(nrows):  # 遍历输出
            temp = sh.cell(i, 0).value.split('@')
            proxy_list.append(temp[0])
        print(proxy_list)


    def read_web_site_list_file(self):
        openfilename = askopenfilename(filetypes=[('xls', '*.xls')])
        wb = xlrd.open_workbook(openfilename)
        sh = wb.sheet_by_index(0)  # 第一个表
        nrows = sh.nrows  # 获取一共有多少行
        for i in range(nrows):  # 遍历输出
            web_site = sh.cell(i, 0).value
            count = sh.cell(i, 1).value
            tuple = (web_site, count)
            self.web_site_list.append(tuple)


    def global_test(self):
        print(self.proxy_list)
        print(self.web_site_list)

    def resize(self, ev=None):
        label.config(font='Helvetica -%d bold' % scale.get())

    def start_click(self):
        self.click_count.set(10)
        print("start!")

if __name__=='__main__':
    root = Tk()
    root.title("涌涌奥特曼")
    root.geometry('600x450')  #是x 不是*

    root.resizable(width=True, height=True) #宽不可变, 高可变,默认为True
    union_root = UnionRobot()






    #frm = Frame(root)

    # left
    frm_L = Frame(width=40,height =20,relief="ridge",borderwidth=1)

    label = Label(frm_L, borderwidth=1, padx=1, pady=1, background="white", height=20, width=35, relief="ridge",
          textvariable=union_root.click_count, font=('Arial', 10)).pack(side=TOP)
    # 使用lambda传递参数
    #Button(frm_L, height=2,width=10,text="导入代理", font=('宋体', 10),command=lambda: union_root.read_xls('f:\\proxy.xls')).pack(side=LEFT)
    # 使用askopenfilename去读取文件
    Button(frm_L, height=2,width=10,text="导入代理列表", font=('宋体', 10),command=union_root.read_xls_file).pack(side=LEFT)
    Button(frm_L, height=2,width=10,text="导入目标网站列表", font=('宋体', 10),command=union_root.read_web_site_list_file).pack(side=LEFT)
    Button(frm_L, height=2,width=10,text="显示全局变量", font=('宋体', 10),command=union_root.global_test).pack(side=LEFT)

    # 进度条控件
    # scale = Scale(frm_L, from_=10, to=40, orient=HORIZONTAL, command=union_root.resize)  # 10-40
    # scale.set(12)  # 初始位置
    # scale.pack(fill=X, expand=1)

    frm_L.pack(side=LEFT)

    # middle

    #frm_M = Frame(frm)
    # frm_M = Frame(width=30,height =20,relief="ridge")
    # Label(frm_M, borderwidth=1, padx=1, pady=1, background="white", height=15, width=20, relief="ridge",
    #       textvariable=union_root.click_count, font=('Arial', 10)).pack(side=TOP)
    #
    # # 使用lambda传递参数
    # Button(frm_M, height=2,width=10,text="导入代理", font=('宋体', 10),command=lambda: union_root.read_xls('f:\\proxy.xls')).pack()
    # # 使用askopenfilename去读取文件
    # Button(frm_M, height=2,width=10,text="导入代理列表", font=('宋体', 10),command=union_root.read_xls_file).pack()
    # Button(frm_M, height=2,width=10,text="导入目标网站列表", font=('宋体', 10),command=union_root.read_web_site_list_file).pack()
    # Button(frm_M, height=2,width=10,text="显示全局变量", font=('宋体', 10),command=union_root.global_test).pack()
    #
    # frm_M.pack(side=LEFT)


    #right

    #frm_R = Frame(frm)
    frm_R = Frame(width=40, height=20, relief="ridge",borderwidth=1)
    Label(frm_R, borderwidth=1, padx=1 , pady=1, background = "white", height=20, width=35, relief="ridge",
          textvariable=union_root.click_count, font=('Arial', 10)).pack(side=TOP)
    Button(frm_R,height=2,width=10,text="开始点击", font=('宋体', 10),command=union_root.start_click).pack(side=BOTTOM)

    frm_R.pack(side=RIGHT)


    #frm.pack()



    root.mainloop()