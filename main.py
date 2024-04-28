import tkinter as tk
import tkinter.messagebox as messagebox
import time
import os
import json
import requests
import sys
import subprocess
import ctypes
import win32com.client
import threading

def on_closing():
    with open(file_path, 'a') as f:
        f.write(f"执行退出程序！\n\n\n\n")
    root.quit()
    sys.exit()
def login():
    # 获取用户名和密码
    username = user_entry.get()
    password = password_entry.get()
    # 获取复选框的状态
    save_password = bool(save_password_var.get())
    auto_login = bool(auto_login_var.get())

    # 如果复选框被选中，保存用户名和密码
    if save_password or auto_login:
        save_info()

    # 显示消息
    login_info_label.config(text="正在登录，请稍后！\n速度取决于您的网络...", font=("微软雅黑", 16))

    # 创建一个新的线程来进行网络请求
    t = threading.Thread(target=do_request, args=(username, password))
    t.start()

def get_url(username, password):
    timestamp = int(time.time() * 1000)
    url = f"http://172.17.100.200:801/eportal/?c=GetMsg&a=loadToken&callback=jQuery_{timestamp}&account={username}&password={password}&mac=000000000000&_={timestamp}"
    return url

def send_request(url):
    try:
        with open(file_path, 'a') as f:
            f.write(f"发送请求：\n{url}\n\n")
        response = requests.get(url)
        with open(file_path, 'a') as f:
            f.write(f"发送请求：成功！等待返回结果...\n\n")
        root.after(0, lambda: login_info_label.config(text=f"发送请求：成功！等待返回结果..."))
        return response.content,None
    except Exception as error:
        return None, error

def handle_response(response, error):
    if error:
        root.after(0, lambda: login_info_label.config(text=f"网络请求失败(登录失败): {str(error)}"))
        with open(file_path, 'a') as f:
            f.write(f"{str(error)}\n\n")
        messagebox.showinfo("登陆失败", "请检查网络连接！\n日志在log.txt")
        return

    with open(file_path, 'a') as f:
        f.write(f"返回值：\n{response}\n\n") # 将返回值写入文件

    # 解析返回值
    timestamp = int(time.time() * 1000)
    # 检查返回值中是否包含"result":"ok"
    if b'"result":"ok"' in response:
        if opt_out_var.get():
            countdown(4)
        else:
            root.after(0, lambda: login_info_label.config(text=f"登录成功{timestamp}", font=("微软雅黑", 22)))
    else:
        root.after(0, lambda: login_info_label.config(text=f"登录失败{response}", font=("微软雅黑", 16)))
        
def countdown(seconds):
    if seconds > 0:
        login_info_label.config(text=f"登录成功！\n即将在{seconds-1}秒后退出程序！", font=("微软雅黑", 17))
        root.after(1000, countdown, seconds-1)  # 在1秒后再次调用countdown函数
        if seconds < 2:
            on_closing()

def do_request(username, password):
    url = get_url(username, password)
    response, error = send_request(url)
    handle_response(response, error)

def save_info():
    # 获取用户名和密码
    username = user_entry.get()
    password = password_entry.get()
    timestamp = int(time.time() * 1000)
    # 获取复选框的状态
    auto_login = bool(auto_login_var.get())
    global opt_out_var  # 声明 opt_out 是一个全局变量
    opt_out = bool(opt_out_var.get())
    global save_password_var  # 声明 save_password_var 是一个全局变量
    save_password = bool(save_password_var.get())
    

    # 保存用户名和密码
    with open(login_info_path, 'w') as f:
        json.dump({'username': username, 'password': password,  'opt_out': opt_out, 'save_password': save_password,'auto_login': auto_login}, f)
    # 显示消息框
    login_info_label.config(text=f"保存账号密码成功{timestamp}", font=("微软雅黑", 16))

def set_auto_start():
    # 获取当前exe程序的绝对路径
    exe_path = os.path.abspath(sys.argv[0])

    # 创建任务计划
    result=subprocess.call(f'schtasks /Create /SC ONLOGON /TN "开机登陆校园网" /TR "{exe_path}" /F', shell=True)

    # 创建任务计划服务对象
    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()

    # 获取任务计划的根文件夹
    root_folder = scheduler.GetFolder('\\')

    # 获取任务计划
    task_definition = root_folder.GetTask('开机登陆校园网').Definition

    # 获取任务计划的设置
    settings = task_definition.Settings

    # 修改任务计划的属性
    settings.DisallowStartIfOnBatteries = False
    settings.StopIfGoingOnBatteries = False

    # 保存修改后的任务计划
    root_folder.RegisterTaskDefinition('开机登陆校园网', task_definition, 6, '', '', 3)
    # 检查是否成功创建
    if result == 0:
        login_info_label.config(text=f"创建开机自启动成功")
    else:
        login_info_label.config(text=f"创建开机自启动失败")
    # 显示消息框
    messagebox.showinfo("提示！","移动位置后要重新设置哦！！")
    # 在设置任务后，检查任务是否存在，然后更新按钮的文本
    if check_task_exists("开机登陆校园网"):
        auto_start_button.config(text="关开机自启", command=remove_auto_start)
    else:
        auto_start_button.config(text="开机自启", command=set_auto_start)
    auto_start_button.update_idletasks()

def Get_administrator_privileges():
        save_info()
        messagebox.showinfo("提示！","将以管理员权限重新运行！！\n已自动保存账号密码")
        # 如果不是管理员，那么就提升权限然后重新运行
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{__file__}" set_title', None, 0)
        # 关闭原窗口
        os._exit(0)

def remove_auto_start():
    # 弹出询问框，询问用户是否要删除任务
    if messagebox.askyesno("确定？", "你确定要删除开机自启动任务吗？"):
        # 用户点击 "Yes"，删除任务
        delete_result = subprocess.call('schtasks /Delete /TN "开机登陆校园网" /F', shell=True)
    
    # 检查是否成功删除
    if delete_result == 0:
        login_info_label.config(text=f"关闭开机自启动成功")
    else:
        login_info_label.config(text=f"关闭开机自启动失败")
    # 在删除任务后，检查任务是否存在，然后更新按钮的文本
    if check_task_exists("开机登陆校园网"):
        auto_start_button.config(text="关开机自启", command=remove_auto_start)
    else:
        auto_start_button.config(text="开机自启", command=set_auto_start)
    auto_start_button.update_idletasks()

def set_window_title(title):
    ctypes.windll.kernel32.SetConsoleTitleW(title)

# 判断是否拥有管理员权限
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def check_and_set_auto_start():
    if is_admin():
        # 如果已经是管理员，那么就运行
        set_auto_start()
    else:
        save_info()
        messagebox.showinfo("提示！","将以管理员权限重新运行！！\n已自动保存账号密码")
        # 如果不是管理员，那么就提升权限然后重新运行
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{__file__}" set_title', None, 0)
        # 关闭原窗口
        os._exit(0)
        
def open_log_folder():
    os.startfile(data_dir)
#检查是否创建自启动计划
def check_task_exists(task_name):
    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()
    root_folder = scheduler.GetFolder('\\')
    for task in root_folder.GetTasks(0):
        if task.Name == task_name:
            return True
    return False

def truncate_large_file(file_path, max_size=1024*1024*20):
    # 检查文件大小
    if os.path.getsize(file_path) > max_size:
        # 文件大小超过限制
        with open(file_path, 'w') as f:
            f.truncate(0)  # 将文件截断到0字节

# 创建主窗口
root = tk.Tk()
root.protocol("WM_DELETE_WINDOW", on_closing)
#设置标题
if is_admin():
    root.title("登录校园网 - 管理员")
else:
    root.title("登录校园网")

# 设置窗口大小
root.geometry("400x290")

# 获取用户的应用数据目录
appdata_dir = os.getenv('APPDATA')
# 创建一个新的目录来存储你的程序的数据
data_dir = os.path.join(appdata_dir, 'Campus Network')
# 如果目录不存在，创建它
if not os.path.exists(data_dir):
    os.makedirs(data_dir)
    # 设置目录属性为隐藏
    # subprocess.run(['attrib', '+h', data_dir])

# 修改你的代码，将所有的文件路径都改为新的目录下的路径
file_path = os.path.join(data_dir, 'log.txt')
login_info_path = os.path.join(data_dir, 'login_info.json')

truncate_large_file(file_path)

# 创建一个标签控件显示操作提示，并设置文本大小为15
label_label = tk.Label(root, text="不重启和断网线\n就不会掉线(除了服务器原因)", font=("微软雅黑", 14))
label_label.place(x=55, y=5)

open_log_button = tk.Button(root, text="打开日志", command=open_log_folder, font=("微软雅黑", 14))
open_log_button.place(x=10, y=175)

# 创建一个标签控件作为用户名的提示，并设置文本大小为15
user_label = tk.Label(root, text="用户名:", font=("微软雅黑", 14))
user_label.place(x=25, y=65)

# 创建一个文本框控件用于输入用户名
user_entry = tk.Entry(root, font=("微软雅黑", 14))
user_entry.place(x=125, y=65)

# 创建一个标签控件作为密码的提示，并设置文本大小为15
password_label = tk.Label(root, text="密码:", font=("微软雅黑", 14))
password_label.place(x=25, y=115)

# 创建一个文本框控件用于输入密码
password_entry = tk.Entry(root, font=("微软雅黑", 14), show="*")
password_entry.place(x=125, y=115)

# 创建一个复选框控件作为自动退出的选项
opt_out_var = tk.IntVar()
opt_out_checkbutton = tk.Checkbutton(root, text="自动退出", variable=opt_out_var)
opt_out_checkbutton.place(x=15, y=145)
# 创建一个复选框控件作为保存密码的选项
save_password_var = tk.IntVar()
save_password_checkbutton = tk.Checkbutton(root, text="保存密码", variable=save_password_var)
save_password_checkbutton.place(x=125, y=145)

# 创建一个复选框控件作为自动登录的选项
auto_login_var = tk.IntVar()
auto_login_checkbutton = tk.Checkbutton(root, text="自动登录", variable=auto_login_var)
auto_login_checkbutton.place(x=225, y=145)

# 创建一个按钮控件作为登录按钮，并设置文本大小为15
login_button = tk.Button(root, text="登录", font=("微软雅黑", 14), command=login)
login_button.place(x=125, y=175)

# 创建一个按钮控件作为保存按钮，并设置文本大小为15
save_button = tk.Button(root, text="保存", font=("微软雅黑", 14), command=save_info)
save_button.place(x=200, y=175)

# 创建一个标签控件显示登录提示，并设置文本大小为15
login_info_label = tk.Label(root, font=("微软雅黑", 14),wraplength=370)
login_info_label.place(x=15, y=225)

if is_admin():
    login_info_label.config(text=f"当前拥有“管理员权限”谨慎操作\n点击“开机自启”或“关闭开机自启”设置")
else:
    login_info_label.config(text=f"需要“管理员权限”才能设置开机自启\n点击“获取权限”按钮或手动获取权限")
    
# 创建一个按钮，并设置文本大小为15
auto_start_button = tk.Button(root, font=("微软雅黑", 14))
auto_start_button.place(x=275, y=175)


if is_admin():
    if check_task_exists("开机登陆校园网"):
        auto_start_button.config(text="关开机自启",command=remove_auto_start)
    else:
        auto_start_button.config(text="开机自启",command=set_auto_start)
else:
    auto_start_button.config(text="获取权限",command=Get_administrator_privileges)

# 如果存在保存的用户名和密码，自动填入
# 如果存在'login_info.json'文件，则读取文件内容并更新用户输入和自动登录选项
if os.path.exists(login_info_path):
    with open(login_info_path, 'r') as f:
        login_info = json.load(f)
        user_entry.insert(0, login_info['username'])  # 将文件中的用户名更新到用户输入框
        password_entry.insert(0, login_info['password'])  # 将文件中的密码更新到密码输入框
        auto_login_var.set(bool(login_info.get('auto_login', 0)))  # 设置自动登录选项为文件中的值（如果不存在则默认为0）
        opt_out_var.set(bool(login_info.get('opt_out', 0))) # 设置自动退出选项为文件中的值（如果不存在则默认为0）
        save_password_var.set(bool(login_info.get('save_password', 0))) # 设置自动退出选项为文件中的值（如果不存在则默认为0）
# 如果自动登录选项被选中，则执行登录函数
if bool(auto_login_var.get()):
    login()

# 运行主事件循环
root.mainloop()