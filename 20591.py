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
        global a
        a += 1
        print a
        if a < len(csvs):
            r = csvs[a].strip().split(',')[0]
            i = csvs[a].strip().split(',')[1]
            time.sleep(0.8)
            typer('2059110')
            time.sleep(0.02)
            typer(r)
            time.sleep(0.02)
            press('enter','enter','enter')
            time.sleep(0.02)
            typer('a')
            time.sleep(0.02)
            typer(i)
            time.sleep(0.02)
            press('enter')
    return True


print u'''
改关系
按右CTRL开始'''

with open('C:\\retail.txt') as csvfile:
    csvs = csvfile.readlines()

global a
a = -1

# 创建一个“钩子”管理对象
hm = pyHook.HookManager()

# 监听所有键盘事件
hm.KeyDown = onKeyboardEvent

hm.HookKeyboard()
# 一直监听，直到手动退出程序
pythoncom.PumpMessages(1000)






### aa.split('\r\n') \t
### resource@ https://gist.github.com/chriskiehl/2906125
