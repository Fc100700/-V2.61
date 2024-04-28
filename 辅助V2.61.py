from tkinter import messagebox
from colorama import Fore, init
from playsound import playsound
from win32com.client import Dispatch
from pynput import keyboard
import datetime
import os
import platform
import threading
import tkinter as tk
import tkinter.messagebox
import webbrowser
import random
import keyboard as keys
import pyautogui
import sys
import time
import ttkbootstrap
import win32com.client
import psutil
import json


init()
starttime = datetime.datetime.now()
Auto_Number = 1
l1 = 0.01
Click_Times_ = 1000
Auto_Pause = 0.1
QQ_2 = 0
QQ_1 = 0
Random_list = [1, 2, 3]
Click_Pause = 0.01
Version = 2.61
target_process_name = "QQ.exe"  #进程名称
print(Fore.GREEN)
print('■■■■■■■')
print('■                                                     ■')
print('■                                                     ■ ')
print('■                 ■        ■       ■■■■■       ■                 ■■■■■        ■  ■■■')
print('■■■■■■       ■        ■     ■                 ■■■■■       ■          ■      ■■      ■')
print('■                 ■        ■     ■                 ■        ■     ■          ■      ■          ■')
print('■                 ■        ■     ■                 ■        ■     ■■■■■■■      ■          ■')
print('■                 ■        ■     ■                 ■        ■     ■                  ■          ■')
print('■                 ■        ■     ■                 ■        ■     ■                  ■          ■')
print('■                 ■        ■     ■                 ■        ■     ■                  ■          ■')
print('■                 ■■■■■■       ■■■■■       ■        ■       ■■■■■■      ■          ■')
print(Fore.RESET + '此窗口为控制台请勿关闭')
print('此软件为浮沉制作\n\n')
print(Fore.BLUE)
print('[' + time.strftime("%H:%M:%S") + ']' + '初始化成功')


try:
    with open(f"position.json", "r") as file:
        U_data = json.load(file)
        lineEdit = U_data["lineEdit"]
        send_button = U_data["send_button"]
        infor1 = U_data["infor1"]
        infor2 = U_data["infor2"]
        infor3 = U_data["infor3"]
    print('[' + time.strftime("%H:%M:%S") + ']' + '配置信息获取成功')
except:
    print(Fore.RED +'[' + time.strftime("%H:%M:%S") + ']' + '配置信息获取失败 将使用默认位置' + Fore.BLUE)
    lineEdit = [125, 977]
    send_button = [1644, 1025]
    infor1 = [1649,866]
    infor2 = [1649,894]
    infor3 = [1649,921]


print('[' + time.strftime("%H:%M:%S") + ']' + f'聊天框位置: {lineEdit}')
print('[' + time.strftime("%H:%M:%S") + ']' + f'发送按钮位置: {send_button}')
print('[' + time.strftime("%H:%M:%S") + ']' + f'快捷消息发送位置: {infor1}，{infor2}，{infor3}')


class MyThread(threading.Thread):  # 多线程封装（我也看不懂反正就是这么用的）
    def __init__(self, func, *args):
        super().__init__()

        self.func = func
        self.args = args

        self.setDaemon(True)
        self.start()  # 在这里开始

    def run(self):
        self.func(*self.args)


def play_prompt_sound(file_path):
    MyThread(playsound, file_path)


def gx():  # 更新
    webbrowser.open('http://fcyang.top/')


def Auto__Ckick():
    print("XFBSOMDFLS")


def liand1():  # 连点器时间
    global l1
    parameter = pyautogui.prompt('请输入连点器执行时间\n默认为0.01s每次')
    if parameter:
        l1 = float(parameter)  # 获取自动脚本每次点击的时间l1


def liand2():  # 连点器次数
    global Click_Times_
    parameter = pyautogui.prompt('请输入连点器执行次数\n默认为1000次')
    if parameter:
        Click_Times_ = int(parameter)  # 获取自动脚本执行的次数l2


def liand5():  # 自动脚本点击时间
    global Auto_Pause
    parameter = pyautogui.prompt('请输入每次自动点击时间\n默认为0.1s每次')
    if parameter:
        Auto_Pause = float(parameter)  # 获取连点器每次点击的时间ld


def liand6():  # 自动脚本执行次数
    global Auto_Number
    parameter = pyautogui.prompt('请输入自动脚本次数\n默认为一次')
    Auto_Number = int(parameter)  # 获取连点器点击的次数lo


def Auto_Record():  # 记录自动脚本
    play_prompt_sound("C:\\Windows\\Media\\Windows Notify Messaging.wav")
    global figure
    figure = 0
    try:
        os.remove('1.xml')
        print(sys_time + '删除上次配置')
    except:
        print(sys_time + '创建配置')
    print(sys_time + '开始记录脚本位置')
    spam = []

    def on_press(key):
        if key == keyboard.Key.home:  # 按下home键开始记录点击位置
            global figure
            figure = figure + 1
            a1 = pyautogui.position()
            a2 = [a1.x, a1.y]
            spam.append(a2)
            print(sys_time + f'第{figure}处位置设置成功')
            play_prompt_sound("C:\\Windows\\Media\\Windows Notify Email.wav")

    def on_release(key):
        if key == keyboard.Key.end:  # 按下end键结束记录
            if not spam:
                print(Fore.RED + sys_time + '本次未记录点击位置' + Fore.BLUE)
                pyautogui.confirm('未记录点击位置')
                sys.exit()
            so = open('1.xml', 'w')
            for c in range(0, len(spam)):
                so.write(f'{spam[c]}\n')
            so.close()
            print(sys_time + '结束设置')
            play_prompt_sound("C:\\Windows\\Media\\Windows Message Nudge.wav")
            sys.exit()

    with keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()  # 热键监听


def Auto_Run():  # 执行自动脚本
    play_prompt_sound("C:\\Windows\\Media\\Windows Notify Messaging.wav")
    shu = os.path.exists('1.xml')
    spam1 = []
    if not shu:
        print(Fore.RED + sys_time + '未找到数据' + Fore.BLUE)  # 未找到xml文件 暂停函数
        tk.messagebox.showwarning('', '未找到数据')
        sys.exit()
    with open("1.xml") as lines:
        for line in lines:
            spam1.append(eval(line))
    if not spam1:
        print(Fore.RED + sys_time + '未找到数据' + Fore.BLUE)  # xml中无数据 暂停函数
        pyautogui.confirm('未找到数据')
        sys.exit()
    pyautogui.PAUSE = 0.01
    i1 = 0
    for x in range(Auto_Number):  # 脚本次数
        for i in range(0, len(spam1)):  # 获取ldq列表的长度 决定点击次数
            i1 = i1 + 1
            pyautogui.click(spam1[i])  # 依次点击
            time.sleep(Auto_Pause)
            print(sys_time + f'第{i + 1}处位置''点击成功')  # 点击状态提示


def Open_Web():
    play_prompt_sound("C:\\Windows\\Media\\Windows Notify Messaging.wav")
    webbrowser.open('https://www.baidu.com')  # 打开浏览器
    print(sys_time + '浏览器打开成功')


def qq1():
    global QQ_1
    QQ_1 = pyautogui.prompt('请输入QQ号')  # QQ号1


def qq2():
    global QQ_2
    QQ_2 = pyautogui.prompt('请输入QQ号')  # QQ号2


def check_process_exists(process_name):
    for process in psutil.process_iter(attrs=['pid', 'name']):
        if process.info['name'] == process_name:
            return True
    return False


def Click():  # 连点器
    play_prompt_sound("C:\\Windows\\Media\\Windows Notify Messaging.wav")
    print(sys_time + '连点器打开成功  按下f9开始连点')

    def on_press(key):
        if key == keyboard.Key.f9:  # 按下f9开始连点器
            pyautogui.PAUSE = l1  # 每次点击的时间为l1
            for i in range(Click_Times_):  # 点击次数为l2 次数到了结束点击
                pyautogui.click()

    with keyboard.Listener(on_press=on_press) as listener:  # 热键监听（钩子）
        listener.join()


def Send_Copy():  # 发送复制消息
    if not check_process_exists(target_process_name):
        pyautogui.confirm("QQ未启动!")
        return 0
    play_prompt_sound("C:\\Windows\\Media\\Windows Notify Messaging.wav")
    time.sleep(3)
    pyautogui.PAUSE = Click_Pause
    while True:
        if keys.is_pressed("F10"):  # Check if F10 key is pressed
            print(Fore.GREEN + sys_time + '结束发送' + Fore.BLUE)
            break
        pyautogui.click(lineEdit)  # 点击第一处位置
        pyautogui.hotkey('ctrl', 'v')  # 粘贴
        randfigure = random.choice(Random_list)  # 随机字符输入
        if randfigure == 1:
            pyautogui.press('.')
        if randfigure == 2:
            pyautogui.press('。')
        if randfigure == 3:
            pyautogui.press(',')
        pyautogui.click(send_button)  # 点击第二处位置


def QQ1():  # QQ号1
    if not check_process_exists(target_process_name):
        pyautogui.confirm("QQ未启动!")
        return 0
    play_prompt_sound("C:\\Windows\\Media\\Windows Notify Messaging.wav")
    if QQ_1 == 0:
        pyautogui.confirm('请输入QQ号')
        sys.exit()
    if len(str(QQ_1)) <= 5:
        pyautogui.confirm('请输入正确的QQ号')
        sys.exit()
    if len(str(QQ_1)) >= 11:
        pyautogui.confirm('请输入正确的QQ号')
        sys.exit()
    time.sleep(3)
    pyautogui.PAUSE = 0.015
    while True:
        if keys.is_pressed("F10"):  # Check if F10 key is pressed
            print(Fore.GREEN + sys_time + '结束发送' + Fore.BLUE)
            break
        pyautogui.click(lineEdit)
        pyautogui.write(f'@{QQ_1}')
        time.sleep(0.03)
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl', 'v')
        randfigure = random.choice(Random_list)  # 随机
        if randfigure == 1:
            pyautogui.press('.')
        if randfigure == 2:
            pyautogui.press('。')
        if randfigure == 3:
            pyautogui.press(',')
        pyautogui.click(send_button)


def QQ2():  # QQ号2
    if not check_process_exists(target_process_name):
        pyautogui.confirm("QQ未启动!")
        return 0
    play_prompt_sound("C:\\Windows\\Media\\Windows Notify Messaging.wav")
    if QQ_2 == 0:
        pyautogui.confirm('请输入QQ号')
        sys.exit()
    if len(str(QQ_2)) <= 5:
        pyautogui.confirm('请输入正确的QQ号')
        sys.exit()
    if len(str(QQ_2)) >= 11:
        pyautogui.confirm('请输入正确的QQ号')
        sys.exit()
    time.sleep(3)
    pyautogui.PAUSE = 0.015
    while True:
        if keys.is_pressed("F10"):  # Check if F10 key is pressed
            print(Fore.GREEN + sys_time + '结束发送' + Fore.BLUE)
            break
        pyautogui.click(lineEdit)
        pyautogui.write(f'@{QQ_2}')
        time.sleep(0.03)
        pyautogui.press('enter')
        pyautogui.hotkey('ctrl', 'v')
        randfigure = random.choice(Random_list)
        if randfigure == 1:
            pyautogui.press('.')  # 随机符号1
        if randfigure == 2:
            pyautogui.press('。')  # 随机符号2
        if randfigure == 3:
            pyautogui.press(',')  # 随机符号3
        pyautogui.click(send_button)


def Send_Quick():  # 发送快捷消息
    if not check_process_exists(target_process_name):
        pyautogui.confirm("QQ未启动!")
        return 0
    play_prompt_sound("C:\\Windows\\Media\\Windows Notify Messaging.wav")
    time.sleep(3)
    random_list = [1, 2, 3]
    while True:
        if keys.is_pressed("F10"):  # Check if F10 key is pressed
            print(Fore.GREEN + sys_time + '结束发送' + Fore.BLUE)
            break
        randfigure = random.choice(random_list)  # 随机在三个当中运行
        pyautogui.PAUSE = 0.015  # 每次点击时间为0.015
        if randfigure == 1:
            #pyautogui.click(1649, 1009)
            pyautogui.click(send_button)
            pyautogui.click(infor1)  # 发送第一条快捷消息消息
        if randfigure == 2:
            pyautogui.click(send_button)
            pyautogui.click(infor2)  # 发送第二条快捷消息
        if randfigure == 3:
            pyautogui.click(send_button)
            pyautogui.click(infor3)  # 发送第三条快捷消息

def get_mouse():
    pyautogui.mouseInfo()

def get_now_time():  # 获取时间
    global var
    var.set(time.strftime("%H:%M:%S"))  # 获取当前时间
    global sys_time
    sys_time = ('[' + time.strftime("%H:%M:%S") + ']')
    win.after(1000, get_now_time)  # 每隔1s调用函数 gettime 自身获取时间


def get_process():  # 获取进程
    def process():
        def proc_exist(process_name):
            is_exist = False
            wmi = win32com.client.GetObject('winmgmts:')
            processCodeCov = wmi.ExecQuery('select * from Win32_Process where name=\"%s\"' % process_name)
            if len(processCodeCov) > 0:
                is_exist = True
            return is_exist

        if proc_exist('QQ.exe'):
            bar.set('QQ状态:运行中')
        else:
            bar.set('QQ状态:未启动')

    process()
    win.after(1000, get_process)


def get_time():  # 运行时间
    def run_time():
        endtime = datetime.datetime.now()
        times = endtime - starttime
        times = str(times)
        times = times[0:10]
        car.set('运行时间:' + str(times))

    run_time()
    win.after(50, get_time)


def Top():  # 窗口置顶
    if win.wm_attributes('-topmost') == 0:
        win.wm_attributes('-topmost', 1)
    else:
        win.wm_attributes('-topmost', 0)


def Information_1():  # 系统信息
    Infor_1 = '操作系统名称:' + platform.system() + '\n'
    Infor_2 = '操作系统名称及版本号:' + platform.platform() + '\n'
    Infor_3 = '操作系统版本号:' + platform.version() + '\n'
    Infor_4 = '系统位数:' + str(platform.architecture()) + '\n'
    Infor_5 = '计算机类型:' + platform.machine() + '\n'
    Infor_6 = '计算机的网络名称:' + platform.node() + '\n'
    spam = (Infor_1 + Infor_2 + Infor_3 + Infor_4 + Infor_5 + Infor_6)
    pyautogui.confirm(spam)


win = ttkbootstrap.Window()  # 利用ttk库创建tk窗口
win.geometry('600x400+660+340')  # 窗口大小（600x400）
win.iconbitmap("./image/window.ico")
win.resizable(False, False)  # 窗口大小不可改变
win.title("Fuchen")
# win.attributes("-alpha", 1)  # 窗口透明度
var = tkinter.StringVar()  # 时间变量
bar = tkinter.StringVar()  # 状态变量
car = tkinter.StringVar()  # 运行时间变量
menubar = tk.Menu(win)  # 创建菜单
filemenu1 = tk.Menu(menubar, tearoff=0)
filemenu2 = tk.Menu(menubar)
filemenu1.add_command(label='QQ号1', command=qq1)
filemenu1.add_command(label='QQ号2', command=qq2)
filemenu1.add_cascade(label='连点器执行时间', command=liand1)
filemenu1.add_cascade(label='连点器执行次数', command=liand2)
filemenu1.add_cascade(label='自动脚本每次点击间隔时间', command=liand5)
filemenu1.add_cascade(label='自动脚本执行次数', command=liand6)
filemenu1.add_cascade(label='置顶/取消置顶', command=Top)
filemenu2.add_cascade(label='获取电脑信息', command=Information_1)
filemenu2.add_cascade(label='鼠标位置信息', command=get_mouse)
filemenu1.add_cascade(label='版本更新', command=gx)
menubar.add_cascade(label='设置', menu=filemenu1)
menubar.add_cascade(label='工具', menu=filemenu2)
win.config(menu=menubar)  # 创建菜单
image = tk.PhotoImage(file='./image/Background/bj.gif')  # 左上角
theLabel = tk.Label(win, textvariable=var, justify=tk.LEFT, font=('黑体', 15))  # 显示时间
theLabel.pack()  # 绘制标签 时间标签
theLabel.place(x=110, y=280)
Label2 = tk.Label(win, textvariable=bar, font=('黑体', 15))
Label2.pack()
Label2.place(x=220, y=280)  # 状态标签
Label3 = tk.Label(win, text='时间:', font=('黑体', 15))
Label3.pack()
Label3.place(x=55, y=280)
Label5 = tk.Label(win, text=f'辅助版本:V{Version}', font=('黑体', 15))  # 版本标签
Label5.pack()
Label5.place(x=380, y=280)
Label6 = tk.Label(win, textvariable=car)  # 运行时间标签
Label6.pack()
Label6.place(x=480, y=380)
Label7 = tk.Label(win, image=image)  # 左上角图标显示
Label7.pack()
Label7.place(x=-10, y=0)
Button1 = tk.Button(win, text='发送复制内容', command=lambda: MyThread(Send_Copy))
Button1.pack()  # 绘制按钮 下同
Button1.place(x=137, y=170)  # 用place布局 下同
Button2 = tk.Button(win, text='发送快捷消息', command=lambda: MyThread(Send_Quick))
Button2.pack()
Button2.place(x=137, y=200)
Button3 = tk.Button(win, text='记录自动脚本', command=lambda: MyThread(Auto_Record))
Button3.pack()
Button3.place(x=290, y=170)
Button7 = tk.Button(win, text='执行自动脚本', command=lambda: MyThread(Auto_Run))
Button7.pack()
Button7.place(x=290, y=200)
Button4 = tk.Button(win, text='开始(QQ1)', command=lambda: MyThread(QQ1))
Button4.pack()
Button4.place(x=220, y=170)
Button5 = tk.Button(win, text='开始(QQ2)', command=lambda: MyThread(QQ2))
Button5.pack()
Button5.place(x=220, y=200)
Button6 = tk.Button(win, text='打开连点器', command=lambda: MyThread(Click))
Button6.pack()
Button6.place(x=372, y=200)
Button8 = tk.Button(win, text='打开浏览器', command=lambda: MyThread(Open_Web))
Button8.pack()
Button8.place(x=372, y=170)
get_now_time()  # 获取时间
get_process()  # 判断进程是否存在
get_time()  # 运行时间记录
print(sys_time + '窗口创建成功')

win.mainloop()  # 创建窗口