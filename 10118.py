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
        with open('C:\\zz.txt','w+') as csvs:
            for line in lines:
                t = line.strip()
                time.sleep(0.8)
                typer('10118')
                time.sleep(0.05)
                press('enter','enter')
                time.sleep(0.05)
                typer(line.strip())
                time.sleep(0.05)
                press('enter','enter')
                time.sleep(1)
                mouse_absolute(668,250,963,250)
                time.sleep(0.02)
                mouse_rclick(700,250)
                time.sleep(0.08)
                mouse_click(740,268)
                time.sleep(0.02)
                
                tp = get_text().strip()
                t += ',' + tp
                time.sleep(0.01)
                csvs.write(t + '\n')
                press('num_lock','num_lock')
                time.sleep(0.5)

    return True


print u'''
#改关系
按右CTRL开始'''

with open('F:\\WORK\\re.txt') as text:
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
