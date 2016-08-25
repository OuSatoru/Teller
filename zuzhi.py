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
                t = line
                typer('101042')
                typer(line.strip())
                press('enter','enter')
                time.sleep(1)
                mouse_absolute(280,437,605,437)
                time.sleep(0.02)
                mouse_rclick(500,437)
                time.sleep(0.08)
                mouse_click(540,455)
                time.sleep(0.02)
                if get_text().strip() == '':
                    press('enter','enter')
                time.sleep(0.8)
                mouse_absolute(280,437,605,437)
                time.sleep(0.02)
                mouse_rclick(500,437)
                time.sleep(0.08)
                mouse_click(540,455)
                time.sleep(0.02)
                tax = get_text().strip()
                t += ',' + tax
                time.sleep(0.01)
                mouse_absolute(943,194,1236,194)
                time.sleep(0.02)
                mouse_rclick(1000,194)
                time.sleep(0.08)
                mouse_click(1040,212)
                time.sleep(0.01)
                idd = get_text().strip()
                if tax[8:17] == idd[0:8]+idd[-1]:
                    t += ',' + idd
                else:
                    t += ','
                mouse_absolute(149,276,1150,276)
                time.sleep(0.02)
                mouse_rclick(1000,276)
                time.sleep(0.08)
                mouse_click(1040,294)
                time.sleep(0.02)
                location = get_text().strip()
                #print location
                t += ',' + location
                time.sleep(0.01)
                mouse_absolute(362,248,677,248)
                time.sleep(0.02)
                mouse_rclick(600,248)
                time.sleep(0.08)
                mouse_click(640,266)
                time.sleep(0.01)
                phone = get_text().strip()
                t += ',' + phone
                mouse_absolute(278,331,456,331)
                time.sleep(0.02)
                mouse_rclick(400,331)
                time.sleep(0.08)
                mouse_click(440,349)
                time.sleep(0.01)
                name = get_text().strip()
                t += ',' + name
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
