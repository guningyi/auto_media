import win32gui
import win32api
import win32con
import time


def mouse_click(pos):
    win32api.SetCursorPos(pos)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN | win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)

def mouse_click_with_sleep(pos):
    win32api.SetCursorPos(pos)
    time.sleep(2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN| win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)

if __name__ == "__main__":
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

    #获取proxy and server的句柄
    dlg3 = win32gui.FindWindow('Proxy server', None)
    print('start to get Ip Address')
    #if (dlg3 > 0):
    print(dlg3)
    editIp = win32gui.FindWindowEx(dlg3, 0, 'Edit', None)
    print(editIp)
    buffer = '0' * 50
    len = win32gui.SendMessage(editIp, win32con.WM_GETTEXTLENGTH) + 1  # 获取edit控件文本长度
    print(len)
    win32gui.SendMessage(editIp, win32con.WM_GETTEXT, len, buffer)  # 读取文本
    print(buffer[:len - 1])

