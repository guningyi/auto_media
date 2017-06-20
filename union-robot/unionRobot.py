# -*- coding:utf-8 -*-
import win32gui
import win32api
import win32con
import time
import win32clipboard as w
from win32api import GetSystemMetrics

#http://timgolden.me.uk/pywin32-docs/contents.html
#you should use the spyxx or spyxx_amd64, and read the chm files.
#you can find the description of the item which was showed on the window.

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
def getText():
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_TEXT)
    w.CloseClipboard()
    return d

#写入剪切板
def setText(aString):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_TEXT, aString)
    w.CloseClipboard()


# def find_window_in_same_class(className, windowName):
#     dlg = win32gui.FindWindow(className, None)
#     if(dlg > 0):


if __name__ == "__main__":

    #获取屏幕分辨率
    width = GetSystemMetrics(0)
    height = GetSystemMetrics(1)


    dlg = win32gui.FindWindow('Internet Explorer_Server', None)

    #print(dlg)

    dlg2 = win32gui.FindWindow('ToolbarWindow32', None)

    #print(dlg2)

    #dlg3 = win32gui.FindWindowEx(dlg2, 0, "ComboBox", None)

    win32gui.ShowWindow(dlg2, win32con.SW_RESTORE)
    mouse_click((1664, 38))


    #win32gui.ShowWindow(dlg2, win32con.SW_RESTORE)
    mouse_click_with_sleep((1530,317))


    #点击connection
    mouse_click_with_sleep((240, 50))

    #dlg3 = win32gui.FindWindow('SysTabControl32', None)
    #win32gui.ShowWindow(dlg3, win32con.SW_RESTORE)

    #mouse_click_with_sleep(pos)

    #点击LAN 设置
    mouse_click_with_sleep((330, 415))

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


    #移动到文本框右击
    right_click((185,280))

    time.sleep(2)

    #全选
    mouse_click_with_sleep((195, 417))

    # 再次移动到文本框右击
    right_click((185, 280))


    #点击delete清空内容
    mouse_click_with_sleep((195, 395))

    time.sleep(2)



    ip = '11.11.11.11'
    # 将'11.11.11.11'写入剪切板
    setText(ip)

    #第三次移动到Ip地址框，右击
    right_click((185, 280))

    #将IP地址Paste进去
    mouse_click_with_sleep((195, 370))


    #点击确定
    #mouse_click_with_sleep((250, 350))

    #点击外层的确定
    #mouse_click_with_sleep((190, 520))