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

class UnionRobot(object):
    def __init__(self, proxy_list=[], web_site_list=[]):
        self.proxy_list = proxy_list
        self.web_site_list = web_site_list
        self.click_count = StringVar()
        self.available_proxy_list = []
        self.information = StringVar()

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
        for proxy in self.proxy_list:
            temp = proxy.split(':')
            ip = temp[0]
            port = temp[1]
            # print(ip)
            # print(port)
            result = self.proxy_test(ip, port, 'http://www.gvpld.cn/')
            if result == 'http://www.gvpld.cn/':
                self.available_proxy_list.append(proxy)
        print(self.available_proxy_list)
        #self.information.set('测试结果可用的代理列表\n')
        #self.information.set(self.available_proxy_list)

        counter = 0

        # 获取屏幕分辨率
        width = GetSystemMetrics(0)
        height = GetSystemMetrics(1)

        # 定义浏览器各控件位置

        # IE浏览器工具按钮
        ie_toolButton_x = 0
        ie_toolButton_y = 0

        # Intenet 选                                                        项
        internet_option_x = 0
        internet_option_y = 0

        # connection
        connection_x = 0
        connection_y = 0

        # LAN设置
        lan_setting_x = 0
        lan_setting_y = 0

        # ip address文本框
        ip_address_x = 0
        ip_address_y = 0

        # port 文本框
        port_x = 0
        port_y = 0

        # 右键后"全选"的位置
        all_select_x = 0
        all_select_y = 0

        # 右键后"删除"的位置
        delete_x = 0
        delete_y = 0

        # ip address粘贴的位置
        ip_address_paste_x = 0
        ip_address_paste_y = 0

        # lan设置确定
        lan_setting_ok_x = 0
        lan_setting_ok_y = 0

        # internet设置确定
        internet_setting_ok_x = 0
        internet_setting_ok_y = 0

        # ie浏览器地址输入栏
        ie_address_edit_x = 0
        ie_address_edit_y = 0

        # ie浏览器地址输入栏右键后的全选坐标
        ie_address_all_select_x = 0
        ie_address_all_select_y = 0

        # ie浏览器地址输入栏右键后的delete的坐标
        ie_address_delete_x = 0
        ie_address_delete_y = 0

        # ie浏览器地址输入栏右键后的paste的坐标
        ie_address_paste_x = 0
        ie_address_paste_y = 0

        # ie浏览器访问箭头
        ie_address_access_x = 0
        ie_address_access_y = 0

        if width == 1920 and height == 1080:

            ie_toolButton_x = 1900
            ie_toolButton_y = 37

            # Intenet 选项
            internet_option_x = 1700
            internet_option_y = 320

            # connection
            connection_x = 227
            connection_y = 65

            # LAN设置
            lan_setting_x = 400
            lan_setting_y = 499

            # ip address文本框
            ip_address_x = 173
            ip_address_y = 347

            # ip右键后"全选"的位置
            all_select_x = 240
            all_select_y = 485

            # 右键后"删除"的位置
            delete_x = 260
            delete_y = 455

            # ip address粘贴的位置
            ip_address_paste_x = 260
            ip_address_paste_y = 430

            # port 文本框
            port_x = 315
            port_y = 345

            # port右键后出现的菜单的all select的坐标
            port_all_select_x = 350
            port_all_select_y = 475

            # port右键后出现的菜单的delete的坐标
            port_delete_x = 400
            port_delete_y = 450

            # port右键后出现的菜单的粘贴的菜单
            port_paste_x = 380
            port_paste_y = 430

            # lan设置确定
            lan_setting_ok_x = 300
            lan_setting_ok_y = 440

            # internet设置确定
            internet_setting_ok_x = 250
            internet_setting_ok_y = 670

            # ie浏览器地址输入栏
            ie_address_edit_x = 500
            ie_address_edit_y = 41

            # ie浏览器地址输入栏右键后的全选坐标
            ie_address_all_select_x = 530
            ie_address_all_select_y = 170

            # ie浏览器地址输入栏右键后的delete的坐标
            ie_address_delete_x = 530
            ie_address_delete_y = 150

            # ie浏览器地址输入栏右键后的paste的坐标
            ie_address_paste_x = 530
            ie_address_paste_y = 120

            # ie浏览器访问箭头
            ie_address_access_x = 942
            ie_address_access_y = 39


        elif width == 1680 and height == 1050:
            ie_toolButton_x = 1664
            ie_toolButton_y = 38

            # Intenet 选项
            internet_option_x = 1530
            internet_option_y = 317

            # connection
            connection_x = 240
            connection_y = 50

            # LAN设置
            lan_setting_x = 330
            lan_setting_y = 415

            # ip address文本框
            ip_address_x = 185
            ip_address_y = 280

            # 右键后"全选"的位置
            all_select_x = 195
            all_select_y = 417

            # 右键后"删除"的位置
            delete_x = 195
            delete_y = 395

            # ip address粘贴的位置
            ip_address_paste_x = 195
            ip_address_paste_y = 370

            # port 文本框
            port_x = 278
            port_y = 280

            # port右键后出现的菜单的all select的坐标
            port_all_select_x = 361
            port_all_select_y = 416

            # port右键后出现的菜单的delete的坐标
            port_delete_x = 354
            port_delete_y = 385

            # port右键后出现的菜单的粘贴的菜单
            port_paste_x = 354
            port_paste_y = 370

            # lan设置确定
            lan_setting_ok_x = 250
            lan_setting_ok_y = 350

            # internet设置确定
            internet_setting_ok_x = 217
            internet_setting_ok_y = 522

            # ie浏览器地址输入栏
            ie_address_edit_x = 300
            ie_address_edit_y = 38

            # ie浏览器地址输入栏右键后的全选坐标
            ie_address_all_select_x = 309
            ie_address_all_select_y = 179

            # ie浏览器地址输入栏右键后的delete的坐标
            ie_address_delete_x = 310
            ie_address_delete_y = 150

            # ie浏览器地址输入栏右键后的paste的坐标
            ie_address_paste_x = 310
            ie_address_paste_y = 130

            # ie浏览器访问箭头
            ie_address_access_x = 526
            ie_address_access_y = 38

        else:
            print("屏幕尺寸未定义-5秒后将退出程序！")
            time.sleep(5)
            sys.exit()

        for proxy in self.available_proxy_list:
            temp = proxy.split(':')
            proxy_address = temp[0]
            proxy_port = temp[1]

            # dlg = win32gui.FindWindow('Internet Explorer_Server', None)

            # print(dlg)

            # dlg2 = win32gui.FindWindow('ToolbarWindow32', None)

            # print(dlg2)

            # dlg3 = win32gui.FindWindowEx(dlg2, 0, "ComboBox", None)

            # win32gui.ShowWindow(dlg2, win32con.SW_RESTORE)

            # IE浏览器 工具按钮
            self.mouse_click((ie_toolButton_x, ie_toolButton_y))

            # win32gui.ShowWindow(dlg2, win32con.SW_RESTORE)
            # Intenet 选项
            self.mouse_click_with_sleep((internet_option_x, internet_option_y))

            # 点击connection
            self.mouse_click_with_sleep((connection_x, connection_y))

            # dlg3 = win32gui.FindWindow('SysTabControl32', None)
            # win32gui.ShowWindow(dlg3, win32con.SW_RESTORE)

            # mouse_click_with_sleep(pos)

            # 点击LAN 设置
            self.mouse_click_with_sleep((lan_setting_x, lan_setting_y))

            # 获取Local Area Network (LAN) Settings的句柄
            # 注意，如果你要处理某个子控件，一定要找到其准确的父窗口。
            # 不能仅仅凭借图形界面上的结构来确定窗口与控件之间的父子关系
            # 不能根据 '#32770 (Dialog)' 来查找Local Area Network的句柄
            # 因为在Desktop下所有的Dialog都是'#32770 (Dialog)'
            # 假如使用 win32gui.FindWindow('#32770', None) 得到的PyHandler值是Ctrix的值而不是IE的值
            # 这是因为Ctrix出现的位置那棵树中靠前
            # 使用className, windowName结合，win32gui.FindWindow('#32770', 'Local Area Network (LAN) Settings')
            # 找到的值就是IE中的Local Area Network(LAN) Settings的值
            # 以上语句在pythonWin IDLE中执行正确(指将语句copy进IDLE的console中执行)
            # 但如果在IDLE中打开.py文件执行，则得到的结果不正确

            # dlg3 = win32gui.FindWindow('#32770', 'Local Area Network (LAN) Settings')
            # #if (dlg3 > 0):
            # print(dlg3)
            #
            # editIp = win32gui.FindWindowEx(dlg3, None, 'Edit', None)
            # print(editIp)
            #
            # len = win32gui.SendMessage(3152454, win32con.WM_GETTEXTLENGTH) + 1  # 获取edit控件文本长度
            # print(len)
            #
            # buffer = '0' * 50
            # win32gui.SendMessage(3152454, win32con.WM_GETTEXT, len, buffer)  # 读取文本
            # print(buffer)
            #
            # win32gui.SendMessage(3152454, win32con.WM_SETTEXT, None, '11.11.11.11')


            # 移动到ip address文本框右击
            self.right_click((ip_address_x, ip_address_y))

            time.sleep(2)

            # 全选
            self.mouse_click_with_sleep((all_select_x, all_select_y))

            # 再次移动到文本框右击
            self.right_click((ip_address_x, ip_address_y))

            # 点击delete清空内容
            self.mouse_click_with_sleep((delete_x, delete_y))

            time.sleep(2)

            # ip = '139.224.237.33'
            # 将'11.11.11.11'写入剪切板
            self.setText(proxy_address)

            # 第三次移动到Ip地址框，右击
            self.right_click((ip_address_x, ip_address_y))

            # 将IP地址Paste进去
            self.mouse_click_with_sleep((ip_address_paste_x, ip_address_paste_y))

            # 移动到port文本框右击
            self.right_click((port_x, port_y))

            time.sleep(2)

            # 全选
            self.mouse_click_with_sleep((port_all_select_x, port_all_select_y))

            # 再次移动到port文本框右击
            self.right_click((port_x, port_y))

            # 点击delete清空port内容
            self.mouse_click_with_sleep((port_delete_x, port_delete_y))

            time.sleep(2)

            # port = '8888'
            # 将'80'写入剪切板
            self.setText(proxy_port)

            # 第三次移动到port文本框，右击
            self.right_click((port_x, port_y))

            # 将port Paste进去
            self.mouse_click_with_sleep((port_paste_x, port_paste_y))

            # 点击确定
            self.mouse_click_with_sleep((lan_setting_ok_x, lan_setting_ok_y))

            time.sleep(3)

            # 点击外层的确定
            self.mouse_click_with_sleep((internet_setting_ok_x, internet_setting_ok_y))

            time.sleep(3)

            #######################################以上已经完成了代理的设置###################################

            for target_url in self.web_site_list:
                url = target_url[0]
                max_counter = target_url[1]
                print(max_counter)
                print(counter)
                if max_counter < counter: #假如该站点的点击预设点击次数还没有达到，那么继续。
                    # 修改IE地址，
                    # 移动到IE地址文本框右击
                    self.right_click((ie_address_edit_x, ie_address_edit_y))

                    time.sleep(2)

                    # 全选
                    self.mouse_click_with_sleep((ie_address_all_select_x, ie_address_all_select_y))

                    # 再次移动到IE地址文本框右击
                    self.right_click((ie_address_edit_x, ie_address_edit_y))

                    # 点击delete清空IE地址内容
                    self.mouse_click_with_sleep((ie_address_delete_x, ie_address_delete_y))

                    time.sleep(2)

                    self.setText(url)

                    # 第三次移动到IE地址框，右击
                    self.right_click((ie_address_edit_x, ie_address_edit_y))

                    # 将IE地址Paste进去
                    self.mouse_click_with_sleep((ie_address_paste_x, ie_address_paste_y))

                    # 第四次移动到IE地址框，
                    win32api.SetCursorPos((ie_address_edit_x, ie_address_edit_y))

                    # 点击访问箭头
                    self.mouse_click_with_sleep((ie_address_access_x, ie_address_access_y))

                    # 回车，访问此网站
                    # enter()
                    time.sleep(10)

                    # 在指定的广告框范围内，随机生成一组坐标
                    adv_level_1_x = random.randint(15, 900)
                    adv_level_1_y = random.randint(100, 250)

                    # 点击第一层广告
                    self.mouse_click_with_sleep((adv_level_1_x, adv_level_1_y))

                    # 等待广告页面展示完成
                    time.sleep(20)

                    # 在指定的广告框范围内，随机生成一组坐标
                    adv_level_2_x = random.randint(15, 900)
                    adv_level_2_y = random.randint(100, 250)

                    # 点击第二层广告

                    counter = counter + 1
                else:
                    #将这个站点的信息从缓存中删除
                    self.web_site_list.remove(target_url)
                    #self.information.set(target_url+'：点击已经完成，一共点击'+counter+'次')


if __name__ == "__main__":
    root = Tk()
    root.title("涌涌奥特曼")
    root.geometry('600x450')  # 是x 不是*

    root.resizable(width=True, height=True)  # 宽不可变, 高可变,默认为True
    union_root = UnionRobot()

    # left
    frm_L = Frame(width=40, height=20, relief="ridge", borderwidth=1)

    label = Label(frm_L, borderwidth=1, padx=1, pady=1, background="white", height=20, width=35, relief="ridge",
                  textvariable=union_root.click_count, font=('Arial', 10)).pack(side=TOP)

    Button(frm_L, height=2, width=10, text="导入代理列表", font=('宋体', 10),
           command=union_root.read_xls_file).pack(
        side=LEFT)
    # Button(frm_L, height=2, width=10, text="测试代理", font=('宋体', 10),
    #        command=union_root.read_web_site_list_file).pack(side=LEFT)
    Button(frm_L, height=2, width=10, text="导入目标网站列表", font=('宋体', 10),
           command=union_root.read_web_site_list_file).pack(side=LEFT)


    frm_L.pack(side=LEFT)

    # right

    # frm_R = Frame(frm)
    frm_R = Frame(width=40, height=20, relief="ridge", borderwidth=1)
    Label(frm_R, borderwidth=1, padx=1, pady=1, background="white", height=20, width=35, relief="ridge",
          textvariable=union_root.click_count, font=('Arial', 10)).pack(side=TOP)
    Button(frm_R, height=2, width=10, text="开始点击", font=('宋体', 10), command=union_root.start_click).pack(
        side=BOTTOM)

    frm_R.pack(side=RIGHT)

    # frm.pack()



    root.mainloop()
