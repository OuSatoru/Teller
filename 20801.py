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
            card = csvs[a].strip()
            time.sleep(0.8)
            typer('208010')
            time.sleep(0.02)
            typer(card)
            time.sleep(0.02)
            press('enter')
            time.sleep(5)
            typer('01020036011000')
            time.sleep(0.02)
            press('enter')
            time.sleep(0.02)
            typer('2000')
            time.sleep(0.02)
            press_over_time('enter', 3)
    return True


print u'''
改关系
按右CTRL开始'''

with open('C:\\card.txt') as csvfile:
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
