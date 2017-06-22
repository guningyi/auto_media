# -*- coding:utf-8 -*-
import win32gui
import win32api
import win32con
import time
import win32clipboard as w
from win32api import GetSystemMetrics
import urllib.request
import sys

#http://timgolden.me.uk/pywin32-docs/contents.html
#you should use the spyxx or spyxx_amd64, and read the chm files.
#you can find the description of the item which was showed on the window.
#https://zhidao.baidu.com/question/1929386142632439947.html
#http://blog.csdn.net/lukaishilong/article/details/51900460


# def proxy_test(ip, port, url):
#     # 设置代理
#     proxy_handler = urllib2.ProxyHandler({'http': 'http://' + ip + ':' + str(port) + '/'})
#     opener = urllib2.build_opener(proxy_handler)
#     urllib2.install_opener(opener)
#     # 访问网页
#     req = urllib2.Request(url)

def proxy_test(ip, port, url):
    # 设置代理
    proxy_handler = urllib.request.ProxyHandler({'http': 'http://' + ip + ':' + str(port) + '/'})
    opener = urllib.request.build_opener(proxy_handler)
    urllib.request.install_opener(opener)
    # 访问网页
    req = urllib.request.Request(url)


def mouse_click(pos):
    win32api.SetCursorPos(pos)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN | win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)

def mouse_click_with_sleep(pos):
    win32api.SetCursorPos(pos)
    time.sleep(2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN| win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)

def right_click(pos):
    win32api.SetCursorPos(pos)
    time.sleep(2)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN | win32con.MOUSEEVENTF_RIGHTUP ,0,0,0,0)

#读取剪切板
# def getText():
#     w.OpenClipboard()
#     d = w.GetClipboardData(win32con.CF_TEXT)
#     w.CloseClipboard()
#     return d

#写入剪切板
# def setText(aString):
#     w.OpenClipboard()
#     w.EmptyClipboard()
#     w.SetClipboardData(win32con.CF_TEXT, aString)
#     w.CloseClipboard()

#写入剪切板，使用SetClipboardText()
def setText(aString):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardText(aString)
    w.CloseClipboard()


# def find_window_in_same_class(className, windowName):
#     dlg = win32gui.FindWindow(className, None)
#     if(dlg > 0):


if __name__ == "__main__":

    #获取屏幕分辨率
    width = GetSystemMetrics(0)
    height = GetSystemMetrics(1)

    #定义浏览器各控件位置

    #IE浏览器工具按钮
    ie_toolButton_x = 0
    ie_toolButton_y = 0

    #Intenet 选项
    internet_option_x = 0
    internet_option_y = 0

    #connection
    connection_x = 0
    connection_y = 0

    #LAN设置
    lan_setting_x = 0
    lan_setting_y = 0

    #ip address文本框
    ip_address_x = 0
    ip_address_y = 0

    #port 文本框
    port_x = 0
    port_y = 0

    #右键后"全选"的位置
    all_select_x = 0
    all_select_y = 0

    #右键后"删除"的位置
    delete_x = 0
    delete_y = 0

    #ip address粘贴的位置
    ip_address_paste_x = 0
    ip_address_paste_y = 0

    #lan设置确定
    lan_setting_ok_x = 0
    lan_setting_ok_y = 0

    #internet设置确定
    internet_setting_ok_x = 0
    internet_setting_ok_y = 0


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

        #port右键后出现的菜单的all select的坐标
        port_all_select_x = 350
        port_all_select_y = 475

        #port右键后出现的菜单的delete的坐标
        port_delete_x = 400
        port_delete_y = 450

        #port右键后出现的菜单的粘贴的菜单
        port_paste_x = 380
        port_paste_y = 430

        #lan设置确定
        lan_setting_ok_x = 300
        lan_setting_ok_y = 440

        #internet设置确定
        internet_setting_ok_x = 250
        internet_setting_ok_y = 670


    elif width == 1680 and height == 1050 :
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
        internet_setting_ok_x = 195
        internet_setting_ok_y = 370
    else:
        print("屏幕尺寸未定义-5秒后将退出程序！")
        time.sleep(5)
        sys.exit()


    dlg = win32gui.FindWindow('Internet Explorer_Server', None)

    #print(dlg)

    dlg2 = win32gui.FindWindow('ToolbarWindow32', None)

    #print(dlg2)

    #dlg3 = win32gui.FindWindowEx(dlg2, 0, "ComboBox", None)

    win32gui.ShowWindow(dlg2, win32con.SW_RESTORE)
    #IE浏览器 工具按钮
    mouse_click((ie_toolButton_x, ie_toolButton_y))


    #win32gui.ShowWindow(dlg2, win32con.SW_RESTORE)
    #Intenet 选项
    mouse_click_with_sleep((internet_option_x,internet_option_y))


    #点击connection
    mouse_click_with_sleep((connection_x, connection_y))

    #dlg3 = win32gui.FindWindow('SysTabControl32', None)
    #win32gui.ShowWindow(dlg3, win32con.SW_RESTORE)

    #mouse_click_with_sleep(pos)

    #点击LAN 设置
    mouse_click_with_sleep((lan_setting_x, lan_setting_y))

    #获取Local Area Network (LAN) Settings的句柄
    #注意，如果你要处理某个子控件，一定要找到其准确的父窗口。
    #不能仅仅凭借图形界面上的结构来确定窗口与控件之间的父子关系
    #不能根据 '#32770 (Dialog)' 来查找Local Area Network的句柄
    #因为在Desktop下所有的Dialog都是'#32770 (Dialog)'
    #假如使用 win32gui.FindWindow('#32770', None) 得到的PyHandler值是Ctrix的值而不是IE的值
    #这是因为Ctrix出现的位置那棵树中靠前
    #使用className, windowName结合，win32gui.FindWindow('#32770', 'Local Area Network (LAN) Settings')
    #找到的值就是IE中的Local Area Network(LAN) Settings的值
    #以上语句在pythonWin IDLE中执行正确(指将语句copy进IDLE的console中执行)
    #但如果在IDLE中打开.py文件执行，则得到的结果不正确

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


    #移动到ip address文本框右击
    right_click((ip_address_x,ip_address_y))

    time.sleep(2)

    #全选
    mouse_click_with_sleep((all_select_x, all_select_y))

    # 再次移动到文本框右击
    right_click((ip_address_x, ip_address_y))


    #点击delete清空内容
    mouse_click_with_sleep((delete_x, delete_y))

    time.sleep(2)

    ip = '139.224.237.33'
    # 将'11.11.11.11'写入剪切板
    setText(ip)

    #第三次移动到Ip地址框，右击
    right_click((ip_address_x, ip_address_y))

    #将IP地址Paste进去
    mouse_click_with_sleep((ip_address_paste_x, ip_address_paste_y))




    #移动到port文本框右击
    right_click((port_x,port_y))

    time.sleep(2)

    # 全选
    mouse_click_with_sleep((port_all_select_x, port_all_select_y))

    # 再次移动到port文本框右击
    right_click((port_x, port_y))

    # 点击delete清空port内容
    mouse_click_with_sleep((port_delete_x, port_delete_y))

    time.sleep(2)

    port = '8888'
    # 将'80'写入剪切板
    setText(port)

    # 第三次移动到port文本框，右击
    right_click((port_x, port_y))

    # 将port Paste进去
    mouse_click_with_sleep((port_paste_x, port_paste_y))

    #点击确定
    mouse_click_with_sleep((lan_setting_ok_x, lan_setting_ok_y))

    time.sleep(3)

    #点击外层的确定
    mouse_click_with_sleep((internet_setting_ok_x, internet_setting_ok_y))

    time.sleep(3)