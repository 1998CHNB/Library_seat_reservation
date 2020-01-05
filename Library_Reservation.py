import os
import win32gui
import win32con
import win32api
import time
import random
import pyautogui
import autopy
import tkinter as tk
from threading import Timer
from datetime import datetime
yuyueday = 0
yuyuetime = '25:00'
sleeptime = 1111
def get():
    yuyueday = e1.get()
    yuyuetime = e2.get()
    yuyuehour = int(yuyuetime[0])*10+int(yuyuetime[1])
    yuyueminute = int(yuyuetime[3])*10+int(yuyuetime[4])
    dt=datetime.now()
    sleeptime = int(yuyueday)*24*3600+(yuyuehour-dt.hour)*3600+(yuyueminute-dt.minute)*60
    t = Timer(sleeptime, reservation)
    t.start()
window = tk.Tk()
window.title('图书馆座位预约')
window.geometry('1000x700')
l = tk.Label(window, text='你好，为保证预约成功，请提前登陆微信，并把公众号：“我去图书馆” 置顶。', bg='blue', font=('Arial', 12), width=60, height=2)
l.pack()
tishi1 = tk.Label(window, text='请先输入你想要几天后预约:', font=('Arial', 12), width=30, height=1)
tishi1.pack()
e1 = tk.Entry(window, show = None)
e1.pack()
tishi2 = tk.Label(window, text='请输入你想要预约的具体时间:(格式-例如：6:01请输入：06:01)', font=('Arial', 12), width=50, height=1)
tishi2.pack()
e2 = tk.Entry(window, show = None)
e2.pack()
b1 = tk.Button(window, text='输入确定', font=('Arial', 12),width=10,height=1, command=lambda : get())
b1.pack()
l1 = tk.Label(window, text='', font=('Arial', 12), width=1, height=14)
l1.pack()
l2 = tk.Label(window, text='初次使用时请将文件夹下的png图片换成自己电脑上对应的截图',font=('Arial', 12), width=50, height=1)
l2.pack()
l3 = tk.Label(window, text='这样能使图像识别的准确率提升',font=('Arial', 12), width=40, height=1)
l3.pack()
def reservation():
    T = 0
    while T < 2:
        pingmu_X = win32api.GetSystemMetrics(0)
        pingmu_Y = win32api.GetSystemMetrics(1)
        def move_click(x, y, t=0): 
            win32api.SetCursorPos((x, y)) 
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN |
                                win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)
            if t == 0:
                time.sleep(random.random()*2+1) 
            else:
                time.sleep(t)
            return 0
        AppName = '微信'
        hwnd1 = win32gui.FindWindow("WeChatMainWndForPC",AppName)
        time.sleep(0.8)
        win32gui.ShowWindow(hwnd1, win32con.SW_MAXIMIZE)
        time.sleep(1)
        try:
            cposition = pyautogui.locateOnScreen('gongzhonghao.png')  
            cc = pyautogui.center(cposition)
            pyautogui.moveTo(cc[0],cc[1])  
            pyautogui.click(clicks=1)
        except:
            move_click(int(pingmu_X/7),int(pingmu_Y/7),1)
            time.sleep(0.1)
            move_click(int(pingmu_X/7),int(pingmu_Y/7),1)
        try:
            cposition = pyautogui.locateOnScreen('zuowei.png')  
            cc = pyautogui.center(cposition)
            pyautogui.moveTo(cc[0],cc[1])  
            pyautogui.click(clicks=1)
        except:
            move_click(int(pingmu_X/2.3),int(pingmu_Y/1.1),2)
            time.sleep(1)
        hwnd2 = win32gui.FindWindow(None,AppName)
        win32gui.SetWindowPos(hwnd2,win32con.HWND_TOPMOST,0,0,int(pingmu_X/3),int(pingmu_Y/1.1),win32con.SWP_SHOWWINDOW)
        time.sleep(3)
        try:
            cposition = pyautogui.locateOnScreen('qiandao.png')
            cc = pyautogui.center(cposition)[0]
            win32gui.PostMessage(hwnd1, win32con.WM_CLOSE, 0, 0)
            win32gui.PostMessage(hwnd2, win32con.WM_CLOSE, 0, 0)
        except:
            pass
        try:
            cposition = pyautogui.locateOnScreen('zuoweibiaozhi.png')  
            cc = pyautogui.center(cposition)
            pyautogui.moveTo(cc[0]-int(pingmu_X/4.5),cc[1])  
            pyautogui.click(clicks=1)
            time.sleep(3)
            pyautogui.moveTo(cc[0]-int(pingmu_X/11),cc[1])  
            pyautogui.click(clicks=1)
            time.sleep(3)
        except:
            try:
                cposition = pyautogui.locateOnScreen('changyongzuowei.png')  
                cc = pyautogui.center(cposition)
                pyautogui.moveTo(cc[0]-int(pingmu_X/11),cc[1]+int(pingmu_Y/1.9))  
                pyautogui.click(clicks=1)
                time.sleep(3)
                pyautogui.moveTo(cc[0]+int(pingmu_X/11),cc[1]+int(pingmu_Y/1.9))  
                pyautogui.click(clicks=1)
                time.sleep(3)
            except:
                move_click(int(pingmu_X/11),int(pingmu_Y/1.8),3)
                move_click(int(pingmu_X/4.9),int(pingmu_Y/1.8),3)
        win32gui.PostMessage(hwnd1, win32con.WM_CLOSE, 0, 0)
        win32gui.PostMessage(hwnd2, win32con.WM_CLOSE, 0, 0)
        T+=1
window.mainloop()

