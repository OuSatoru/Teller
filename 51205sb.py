# -*- coding:utf-8 -*-

from typeclip import *
import win32con
import time
import pythoncom
import pyHook
import win32gui
import re


def onKeyboardEvent(event):
    HWND = win32gui.GetForegroundWindow()
    tt = win32gui.GetWindowText(HWND)

    if tt.find(FORE_TITLE) > -1 and event.Key == 'Rcontrol':
        with open('C:\\idsb.txt','w+') as csvs:
            for line in lines:
                t = line.strip()
                time.sleep(0.8)
                typer('51205a0')
                time.sleep(0.05)
                typer(t)
                time.sleep(0.05)
                typer('0')
                time.sleep(0.05)
                press('enter','enter')
                time.sleep(0.9)
                mouse_absolute(170,250,740,500)
                time.sleep(0.02)
                mouse_rclick(470,350)
                time.sleep(0.08)
                mouse_click(510,368)
                time.sleep(0.02)
                ho = get_text()
                tp = re.findall(r"6230666135\d+", ho, re.MULTILINE)
                if len(tp) != 0:
                    t += ',' + tp[0]
                else:
                    t += ','
                time.sleep(0.01)
                csvs.write(t + '\n')
                press('num_lock','num_lock')
                time.sleep(0.5)

    return True


print u'''
身份证id.txt->身份证,社保卡号idsb.txt'''

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
