from tkinter import *
from tkinter.filedialog import askopenfilename
import win32gui
import win32api
import win32con
import time
import win32clipboard as w
from win32api import GetSystemMetrics
import urllib.request
import sys
import xlrd
import xlwt
import datetime
import random
from tkinter.scrolledtext import ScrolledText
import tkinter.font as tkFont
import threading


def proxy_test(ip, port, url):
    result = ""
    # 设置代理
    print("使用代理:" + ip + ':' + port)
    # str = "使用代理:"+ip+':'+port
    # self.information.set('aa')
    proxy_handler = urllib.request.ProxyHandler({'http': 'http://' + ip + ':' + str(port) + '/'})
    opener = urllib.request.build_opener(proxy_handler)
    urllib.request.install_opener(opener)
    global max_num
    max_num = 6
    for i in range(max_num):
        try:
            # 访问网页,带10秒超时
            req = urllib.request.urlopen(url, None, 3)
            print(req.geturl())
            result = req.geturl()
            break
        except Exception as e:
            if i < max_num - 1:
                continue
            else:
                print("time out" + str(e))
                text.insert('insert', ip + ':' + port + '  ' + "time out" + str(e) + '\n')
                text.update()
    return result

def wrap_test_local_proxy_list():
    t = threading.Thread(target=test_local_proxy_list)
    t.setDaemon(True)
    t.start()
    t.join()

def test_local_proxy_list():
    #从本地代理文件中读入代理
    proxy_list = []
    available_proxy_list = []
    openfilename = askopenfilename(filetypes=[('xls', '*.xls')])
    wb = xlrd.open_workbook(openfilename)
    sh = wb.sheet_by_index(0)  # 第一个表
    nrows = sh.nrows  # 获取一共有多少行
    for i in range(nrows):  # 遍历输出
        temp = sh.cell(i, 0).value.split('@')
        proxy_list.append(temp[0])
    #print(proxy_list)

    #测试代理有效性
    for proxy in proxy_list:
        temp = proxy.split(':')
        ip = temp[0]
        port = temp[1]
        result = proxy_test(ip, port, 'http://www.gvpld.cn/')
        if result == 'http://www.gvpld.cn/':
            available_proxy_list.append(proxy)
    print(available_proxy_list)

    #将筛选出的代理写入本地的文件，起名为时间_数量_pass_proxy.xls
    #例如 20160626_165_pass_proxy.xls

    file = xlwt.Workbook(encoding='utf-8')
    now = datetime.datetime.now()
    time_str = (now.strftime('%Y-%m-%d %H:%M:%S')).split(' ')
    date = time_str[0]
    num = len(available_proxy_list)
    name = date + '_' + str(num) + '_pass_proxy.xls'
    table = file.add_sheet('proxy')
    if len(available_proxy_list) > 0:
        text.tag_config('blue', foreground='blue')
        # 使用TAG 'red'来指定文本属性,设为红色字体
        text.insert('insert', '恭喜！可用的代理一共有'+ str(len(available_proxy_list)) + '个'+ '\n','blue')
        text.insert('insert', '代理列表将保存在当前目录下xls文件中'+ '\n','blue')
        text.update()
    else:
        text.tag_config('red', foreground='red')
        # 使用TAG 'red'来指定文本属性,设为红色字体
        text.insert('insert', '日了狗了！没有可用的代理' + '\n','red')
        text.update()
    for index in range(len(available_proxy_list)):
        table.write(index, 0, available_proxy_list[index])
        text.insert('insert', available_proxy_list[index] + '\n')
        text.update()
    file.save(name)


def test_internet_proxy():
    print("待开发")

def spider_internet_proxy():
    print("待开发")




if __name__ == "__main__":
    root = Tk()
    root.title("代理测试器")
    root.geometry('600x450')  # 是x 不是*

    root.resizable(width=False, height=False)  # 宽不可变, 高不可变,默认为True

    local_proxy_test_info = StringVar()
    internet_proxy_test_info = StringVar()
    spider_internet_proxy_info = StringVar()


    # left
    frm_L = Frame(root, width=40, height=20, relief="ridge", borderwidth=1)

    text = ScrolledText(master=frm_L, borderwidth=1, padx=1, pady=1, background="white", height=20, width=35,
                        relief="ridge",
                        font=('Arial', 10))
    text.pack(side=TOP)
    Button(frm_L, height=2, width=14, text="测试本地代理列表", font=('宋体', 10),
           command=test_local_proxy_list).pack(side=LEFT)
    # Button(frm_L, height=2, width=10, text="测试代理", font=('宋体', 10),
    #        command=union_root.read_web_site_list_file).pack(side=LEFT)
    Button(frm_L, height=2, width=14, text="测试网络代理列表", font=('宋体', 10),
           command=test_internet_proxy).pack(side=LEFT)

    frm_L.pack(side=LEFT)

    # right

    # frm_R = Frame(frm)
    frm_R = Frame(root, width=40, height=20, relief="ridge", borderwidth=1)
    Label(frm_R, borderwidth=1, padx=1, pady=1, background="white", height=20, width=35, relief="ridge",
          textvariable=spider_internet_proxy_info, font=('Arial', 10)).pack(side=TOP)
    Button(frm_R, height=2, width=10, text="开始点击", font=('宋体', 10), command=spider_internet_proxy).pack(
        side=BOTTOM)

    frm_R.pack(side=RIGHT)

    # frm.pack()

    root.mainloop()