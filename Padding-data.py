import pandas
import win32com.client
from pynput import mouse
import pyautogui
import time
import threading
import pythoncom
import os
from mss import mss
def to_go():
    # 获取当前活动的 Excel 应用程序
    excel_app = win32com.client.GetActiveObject("Excel.Application")
    # 检查是否有活动的工作簿和工作表
    if excel_app.ActiveWorkbook is not None:
        # 获取活动的工作表（虽然这一步不是必需的，因为后面我们会直接从 Application 获取 ActiveCell）
        active_sheet = excel_app.ActiveSheet
        # 获取活动单元格（注意：这是从 Application 对象获取的，而不是 Worksheet）
        active_cell = excel_app.ActiveCell
        cell_adress = active_cell.Address
        a = list(cell_adress)
        if a[2] == '$':
            cell_adress = cell_adress[1] + cell_adress[3:]
        elif a[3] == '$':
            cell_adress = cell_adress[1:3] + cell_adress[4:]
        elif a[4] == '$':
            cell_adress = cell_adress[1:4] + cell_adress[5:]
        print(cell_adress)
    return  str(cell_adress)
def will_go():
    # 读取数据
    df = pandas.read_csv('C:/Users\hp\Desktop/g.txt', header=None, encoding='GB2312',sep='\s+', skiprows=5)
    a = to_go()
    b = list(a)
    if a[2] in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']:
        d = b[3:]
        d = ''.join(d)
    elif a[1] in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']:
        d = b[2:]
        d = ''.join(d)
    elif a[0] in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']:
        d = b[1:]
        d = ''.join(d)
    f = int(d)
    e = []
    e.append(a)
    for i in range(len(df[1])):
        f += 1
        if a[2] in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']:
            g = a[0:3] + str(f)
        elif a[1] in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']:
            g = a[0:2] + str(f)
        elif a[0] in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']:
            g = a[0:1] + str(f)
        e.append(g)
    return e
#将数据和所在位置封装成字典以便调用
def position():
    a = will_go()
    print(a)
    # 读取数据
    df = pandas.read_csv('C:/Users\hp\Desktop/g.txt', header=None, encoding='GB2312', sep='\s+', skiprows=5)
    b = list(df[1])
    dictionary = {}
    for i in range(len(b)):
        dictionary[a[i]] = b[i]
    return dictionary
def putindata():
    try:
        print('调用了这个方法')
        pythoncom.CoInitialize()
        # 插入数据
        excel_app = win32com.client.GetActiveObject("Excel.Application")
        # 检查是否有活动的工作簿和工作表
        if excel_app.ActiveWorkbook is not None:
            # 获取活动的工作表（虽然这一步不是必需的，因为后面我们会直接从 Application 获取 ActiveCell）
            active_sheet = excel_app.ActiveSheet
            # 获取活动单元格（注意：这是从 Application 对象获取的，而不是 Worksheet）
            for key, value in position().items():
                print(key)
                print(value)
                active_sheet.Range(key).Value = value
                print(position())
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # 清理 COM 库
        pythoncom.CoUninitialize()
# 定义一个函数，用来执行耗时的截图和打开软件操作
def take_screenshot_and_open_software():
    # 获取鼠标点击的起始坐标
    print("请使用鼠标点击屏幕的起始位置...")
    time.sleep(3)
    start_x, start_y = pyautogui.position()
    # 获取鼠标点击的结束坐标
    print("请使用鼠标点击屏幕的结束位置...")
    time.sleep(3)
    end_x, end_y = pyautogui.position()
    # 计算截图区域
    width = abs(end_x - start_x)
    height = abs(end_y - start_y)
    # 截取指定区域
    screenshot = pyautogui.screenshot(region=(start_x, start_y, width, height))
    # 保存截图到文件
    screenshot.save("C:/Users\hp\Desktop/g.jpg")
    # 要打开的软件的完整路径
    software_path = "D:\浏览器下载\GetData (1)\破解补丁/GetData.exe"
    # 使用 os.startfile 打开软件
    os.startfile(software_path)
    time.sleep(2)
    # 模拟按下 Alt+F 组合键
    pyautogui.hotkey('alt', 'f')
    # 模拟按下 o 组合键
    pyautogui.press('o')
    # 把中文输入法改为英文
    pyautogui.hotkey('ctrl', 'space')
    # 模拟按下 g.jpg 组合键
    pyautogui.typewrite('g.jpg')  # 使用 pyautogui.typewrite 输入文件名
    pyautogui.press('enter')
    pyautogui.hotkey('alt','o')
    pyautogui.press('s')
def on_click( x, y,button,pressed):
    if button == mouse.Button.middle and pressed:
        print("Mouse middle button pressed at ({}, {})".format(x, y))
        pyautogui.hotkey('ctrl', 'alt','e')
        time.sleep(0.1)
        pyautogui.hotkey('alt','s')
        time.sleep(0.1)
        pyautogui.hotkey('alt', 'y')
        time.sleep(0.1)
        pyautogui.hotkey('alt', 'o','r')
        time.sleep(0.1)
        pyautogui.hotkey('alt', 'y')
        time.sleep(1)
        thread = threading.Thread(target=putindata)
        thread.start()
    if button == mouse.Button.right and pressed:
        thread = threading.Thread(target=take_screenshot_and_open_software)
        thread.start()
# 创建鼠标监听器，监听点击事件
with mouse.Listener(on_click=on_click) as listion:
    listion.join()