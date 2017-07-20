#!/usr/bin/python
# -*- coding: UTF-8 -*-

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
import random
from tkinter.scrolledtext import ScrolledText
from selenium import webdriver
import os
from selenium.webdriver.common.proxy import Proxy
from selenium.webdriver.common.proxy import ProxyType
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

class UnionRobot(object):
    def __init__(self, proxy_list=[], web_site_list=[]):
        self.proxy_list = proxy_list
        self.web_site_list = web_site_list
        self.click_count = StringVar()
        self.available_proxy_list = []
        self.information = StringVar()
        self.counter = 0

    def read_xls_file(self):
        openfilename = askopenfilename(filetypes=[('xls', '*.xls')])
        wb = xlrd.open_workbook(openfilename)
        sh = wb.sheet_by_index(0)  # 第一个表
        nrows = sh.nrows  # 获取一共有多少行
        for i in range(nrows):  # 遍历输出
            temp = sh.cell(i, 0).value.split('@')
            self.proxy_list.append(temp[0])
        print(self.proxy_list)

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
        print(self.web_site_list)

    def proxy_test(self, ip, port, url):
        result = ""
        # 设置代理
        print("使用代理:"+ip+':'+port)
        #str = "使用代理:"+ip+':'+port
        #self.information.set('aa')
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
        return result

    def mouse_click(self,pos):
        win32api.SetCursorPos(pos)
        time.sleep(2)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN | win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)

    def mouse_click_with_sleep(self,pos):
        win32api.SetCursorPos(pos)
        time.sleep(2)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN| win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)

    def right_click(self,pos):
        win32api.SetCursorPos(pos)
        time.sleep(2)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN | win32con.MOUSEEVENTF_RIGHTUP ,0,0,0,0)

    #写入剪切板，使用SetClipboardText()
    def setText(self,aString):
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardText(aString)
        w.CloseClipboard()

    def start_click(self):
        print("start!")
        # 过滤代理列表，通过连接baidu来测试，并选出可用的代理
        # for proxy in self.proxy_list:
        #     temp = proxy.split(':')
        #     ip = temp[0]
        #     port = temp[1]
        #     # print(ip)
        #     # print(port)
        #     result = self.proxy_test(ip, port, 'http://www.gvpld.cn/')
        #     if result == 'http://www.gvpld.cn/':
        #         self.available_proxy_list.append(proxy)
        # print(self.available_proxy_list)
        #self.information.set('测试结果可用的代理列表\n')
        #self.information.set(self.available_proxy_list)

        self.available_proxy_list = self.proxy_list
        counter = 0

        # 获取屏幕分辨率
        width = GetSystemMetrics(0)
        height = GetSystemMetrics(1)


        if width == 1920 and height == 1080:
            print('你的屏幕分辨率是:'+str(width)+'X'+str(height))
        elif width == 1680 and height == 1050:
            print('你的屏幕分辨率是:' + str(width) + 'X' + str(height))
        else:
            print("屏幕尺寸未定义-5秒后将退出程序！")
            time.sleep(5)
            sys.exit()

        #设置信息输出框中字体的颜色
        r_text.tag_config('blue', foreground='blue')

        for proxy in self.available_proxy_list:
            if len(self.web_site_list) == 0:
                print('全部网站已经点击完成，退出！')
                break;
            temp = proxy.split(':')
            proxy_address = temp[0]
            proxy_port = temp[1]
            #设置Firefox浏览器代理参数
            # 设置代理 111.155.116.195 8123
            # proxy = Proxy(
            #     {
            #         # 'proxyType': ProxyType.MANUAL,  # 用不用都行
            #         # '111.155.116.195': 8123
            #         proxy_address:proxy_port
            #     }
            # )

            print('设置Firefox代理:'+proxy_address+':'+proxy_port)
            profile = webdriver.FirefoxProfile()
            # 手动设置代理
            profile.set_preference('network.proxy.type', 1)
            profile.set_preference('network.proxy.http', proxy_address)
            profile.set_preference('network.proxy.http_port', int(proxy_port))  # int
            profile.update_preferences()

            #######################################以上已经完成了代理的设置###################################


            for target_url in self.web_site_list:
                print("开始遍历网站列表....")
                url = target_url[0]
                max_counter = target_url[1]

                #console log 输出
                print("网址:"+url+' 目标点击次数:'+str(max_counter))
                print('当前已点击次数:'+str(counter))

                #信息框输出
                r_text.insert('insert', "网址:"+url+' 目标点击次数:'+str(max_counter)+ '\n', 'blue')
                r_text.insert('insert', '当前已点击次数:'+str(counter) + '\n', 'blue')
                r_text.update()

                if counter < max_counter : #假如该站点的点击预设点击次数还没有达到，那么继续。
                    # 打开浏览器
                    browser = webdriver.Firefox(firefox_profile=profile)
                    try:
                        browser.get(url)
                    except Exception as e:
                        print("使用代理:"+ proxy_address+':'+proxy_port +" 访问网站发生异常" + str(e))
                        r_text.tag_config('red', foreground='red')
                        r_text.insert('insert', '使用代理:' + proxy_address+':'+proxy_port + " 访问网站发生异常" +'\n', 'red')
                        r_text.update()
                    browser.maximize_window()
                    time.sleep(5)

                    adv_level_1_x = 0
                    adv_level_1_y = 0
                    # 在指定的广告框范围内，随机生成一组坐标
                    if url == 'http://www.gvpld.cn/': #site
                        adv_level_1_x = random.randint(2, 946)
                        adv_level_1_y = random.randint(360, 629)
                    else: #site 2 http://www.china-holeo.com/
                        adv_level_1_x = random.randint(359, 1297)
                        adv_level_1_y = random.randint(359, 529)

                    # 点击第一层广告
                    self.mouse_click_with_sleep((adv_level_1_x, adv_level_1_y))

                    # 等待广告页面展示完成
                    time.sleep(20)

                    # 在指定的广告框范围内，随机生成一组坐标
                    #adv_level_2_x = random.randint(15, 900)
                    #adv_level_2_y = random.randint(100, 250)

                    # 点击第二层广告
                    #self.mouse_click_with_sleep((adv_level_2_x, adv_level_2_y))

                    inputs = browser.find_elements_by_css_selector("#e_idea_pp li h3 a")
                    print('找到二级广告链接数：'+str(len(inputs)))
                    for input in inputs:
                        print(input.get_attribute('href'))
                        time.sleep(2)
                        input.click()
                    time.sleep(3)

                    counter = counter + 1

                    #关闭浏览器
                    browser.close()
                else:
                    #将这个站点的信息从缓存中删除
                    self.web_site_list.remove(target_url)
                    #self.information.set(target_url+'：点击已经完成，一共点击'+counter+'次')



if __name__ == "__main__":
    root = Tk()
    root.title("大智手 v1.0")
    root.geometry('600x450')  # 是x 不是*

    root.resizable(width=False, height=False)  # 宽不可变, 高可变,默认为True
    union_root = UnionRobot()

    # left
    frm_L = Frame(width=40, height=20, relief="ridge", borderwidth=0)

    l_text = ScrolledText(frm_L, borderwidth=3, padx=1, pady=1, background="white", height=20, width=35, relief="ridge",
                  font=('Arial', 10)).pack(padx = 10,side=TOP)

    Button(frm_L, height=2, width=10, text="导入代理列表", font=('宋体', 10),bg='green',fg='white',
           command=union_root.read_xls_file).pack(ipadx =4,ipady = 3, padx =23, pady = 10,
        side=LEFT)
    # Button(frm_L, height=2, width=10, text="测试代理", font=('宋体', 10),
    #        command=union_root.read_web_site_list_file).pack(side=LEFT)
    Button(frm_L, height=2, width=10, text="导入网站列表", font=('宋体', 10),bg='green',fg='white',
           command=union_root.read_web_site_list_file).pack(ipadx =4,ipady = 3, padx =23, pady = 10,side=LEFT)


    frm_L.pack(side=LEFT)

    # right

    # frm_R = Frame(frm)
    frm_R = Frame(width=40, height=20, relief="ridge", borderwidth=0)
    r_text = ScrolledText(frm_R, borderwidth=3, padx=1, pady=1, background="white", height=20, width=35, relief="ridge",
           font=('Arial', 10))
    r_text.pack(padx = 10,side=TOP)
    Button(frm_R, height=2, width=10, text="开始点击", font=('宋体', 10), bg='green',fg='white',command=union_root.start_click).pack(ipadx =4,ipady = 3,pady = 10,
        side=BOTTOM)

    frm_R.pack(side=RIGHT)

    # frm.pack()



    root.mainloop()
