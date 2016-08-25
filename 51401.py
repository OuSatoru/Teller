# -*- coding:utf-8 -*-

from typeclip import *
import win32con
import time
import pythoncom
import pyHook
import win32gui
import sqlite3
import win32api


def onKeyboardEvent(event):
    HWND = win32gui.GetForegroundWindow()
    tt = win32gui.GetWindowText(HWND)

    if tt.find(FORE_TITLE) > -1 and event.Key == 'Rcontrol':
        global a
        a += 1
        print a
        time.sleep(0.8)

        typer('5140111')
        time.sleep(2.6)
        mouse_absolute(767,167,1085,167)
        time.sleep(0.02)
        mouse_rclick(900,167)
        time.sleep(0.08)
        mouse_click(940,185)
        time.sleep(0.02)
        cardnum = get_text()
        with open('C:\\edcard.txt', 'a') as txt:
            txt.write(cardnum + '\n')
        conn = sqlite3.connect(DB_PATH)
        cur = conn.execute("SELECT idnum, addr from card WHERE cardnum = " + cardnum)
        for i in cur:
            idnum = i[0].encode()
            addr = i[1]
        conn.close()
        press('enter','enter')
        time.sleep(0.05)
        typer('a')
        time.sleep(0.05)
        typer(idnum)
        time.sleep(0.05)
        press('enter')
        win32api.MessageBox(0, addr, u'地址', win32con.MB_OK)
        
        time.sleep(0.5)

    return True


print u'''
数据库 C:\sbk.db, 已激活的卡C:\edcard.txt'''

DB_PATH = 'C:\\sbk.db'
TABLE = 'card'
global a
a = -1



        






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
