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
        with open('C:\\idcard.csv','w+') as csvs:
            for line in lines:
                t = line.strip()
                typer('30121001')
                typer('a')
                typer(line.strip())
                press('enter')
                time.sleep(1)
                mouse_absolute(788,278,1100,278)
                time.sleep(0.02)
                mouse_rclick(1000,278)
                time.sleep(0.08)
                mouse_click(1040,296)
                time.sleep(0.04)
                if get_text().strip() == '':
                    press('enter','num_lock')
                '''time.sleep(0.8)
                mouse_absolute(687,303,987,303)
                time.sleep(0.02)
                mouse_rclick(700,303)
                time.sleep(0.08)
                mouse_click(740,321)
                time.sleep(0.02)'''
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

with open('C:\\id.txt') as text:
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
