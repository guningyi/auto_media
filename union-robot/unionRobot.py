import win32gui
import win32api
import win32con


def mouse_click(pos):
    win32api.SetCursorPos(pos)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN | win32con.MOUSEEVENTF_RIGHTUP,0,0,0,0)

if __name__ == "__main__":
    dlg = win32gui.FindWindow(None, 'Internet Explorer')
    win32gui.ShowWindow(dlg, win32con.SW_RESTORE)
    mouse_click((1500, 500))



