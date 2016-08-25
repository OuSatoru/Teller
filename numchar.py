# -*- coding:utf-8 -*-

import typeclip
import time
import pyHook
from bs4 import BeautifulSoup
import pythoncom
import win32gui


def onKeyboardEvent(event):
    HWND = win32gui.GetForegroundWindow()
    tt = win32gui.GetWindowText(HWND)
    if '考试试卷' in tt and event.Key == 'Rcontrol':





flag = True

hm = pyHook.HookManager()
hm.KeyDown = onKeyboardEvent
hm.HookKeyboard()
pythoncom.PumpMessages(1000)