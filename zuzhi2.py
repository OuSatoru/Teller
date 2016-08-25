# -*- coding:utf-8 -*-

from typeclip import *
import win32con
import time
import pythoncom
import pyHook
import win32gui


def onKeyboardEvent(event):
    HWND = win32gui.GetForegroundWindow()
    tt = win32gui.GetWindowText(HWND)

    if tt.find(FORE_TITLE) > -1 and event.Key == 'Rcontrol':
        with open('C:\\zz.csv','w+') as csvs:
            for line in lines:
                t = line.strip()
                typer('205040')
                typer(line.strip())
                press('enter','enter')
                time.sleep(1)
                mouse_absolute(687,303,987,303)
                time.sleep(0.02)
                mouse_rclick(700,303)
                time.sleep(0.08)
                mouse_click(740,321)
                time.sleep(0.02)
                if get_text().strip() == '':
                    press('enter','enter')
                time.sleep(0.8)
                mouse_absolute(687,303,987,303)
                time.sleep(0.02)
                mouse_rclick(700,303)
                time.sleep(0.08)
                mouse_click(740,321)
                time.sleep(0.02)
                tp = get_text().strip()
                t += ',' + tp
                time.sleep(0.01)
                csvs.write(t + '\n')
                press('num_lock')
                time.sleep(0.5)

    return True


print u'''
#改关系
按右CTRL开始'''

with open('C:\\khh.txt') as text:
    lines = text.readlines()



        






'''global a
a = -1'''

# 创建一个“钩子”管理对象
hm = pyHook.HookManager()

# 监听所有键盘事件
hm.KeyDown = onKeyboardEvent

hm.HookKeyboard()
# 一直监听，直到手动退出程序
pythoncom.PumpMessages(1000)






### aa.split('\r\n') \t
### resource@ https://gist.github.com/chriskiehl/2906125
