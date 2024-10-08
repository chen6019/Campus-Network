import shlex
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
import logging

# 创建一个命名的互斥体
mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "dlxyw_mutex")

# 检查互斥体是否已经存在
if ctypes.windll.kernel32.GetLastError() == 183:
    messagebox.showerror("错误", "应用程序已在运行。")
    sys.exit()

def on_closing():
    logger.info("执行退出程序！\n")
    root.quit()
    # 释放互斥体
    ctypes.windll.kernel32.ReleaseMutex(mutex)
    sys.exit()
def login():
    # 获取用户名和密码
    username = user_entry.get()
    password = password_entry.get()
    Link = Link_mode_entry.get()
    # 获取复选框的状态
    save_password = bool(save_password_var.get())
    auto_login = bool(auto_login_var.get())
    Link_mode = bool(Link_mode_var.get())

    if save_password or auto_login:
        save_info()
        
    login_button.config(state=tk.DISABLED)
    
    if Link_mode:
        if Link == "":
            messagebox.showinfo("错误", "请输入链接！")
            login_info_label.config(text="请输入链接！", font=("微软雅黑", 26))
            login_button.config(state=tk.NORMAL)
        else:
            messagebox.showinfo("当前为链接模式",Link)
            login_info_label.config(text="正在登录，请稍后！\n速度取决于您的网络...", font=("微软雅黑", 16))
            url = f"{Link}"
            t = threading.Thread(target=send_request, args=(url,))
            t.daemon = True
            t.start()
    else:
        # 创建一个新的线程来进行网络请求
        login_info_label.config(text="正在登录，请稍后！\n速度取决于您的网络...", font=("微软雅黑", 16))
        t = threading.Thread(target=do_request, args=(username, password))
        t.daemon = True
        t.start()

def get_url(username, password):
    timestamp = int(time.time() * 1000)
    base_url = "http://172.17.100.200:801/eportal/"
    params = {
        "c": "GetMsg",
        "a": "loadToken",
        "callback": f"jQuery_{timestamp}",
        "account": username,
        "password": password,
        "mac": "000000000000",
        "_": timestamp
    }
    url = f"{base_url}?{'&'.join(f'{key}={value}' for key, value in params.items())}"
    return url


def send_request(url):
    Link_mode = bool(Link_mode_var.get())
    try:
        logger.info(f"发送请求：{url}\n")
        logger.info("发送请求：成功！等待返回结果...\n")
        root.after(0, lambda: login_info_label.config(text=f"发送请求：成功！等待返回结果..."))
        response = requests.get(url)
        if Link_mode:
            handle_response(response.content, None)
        else:
            return response.content,None
    except Exception as error:
        if Link_mode:
            handle_response(None,error)
        else:
            return None, error

# #链接模式处理
# def Link_send_request(response):
#     # 解析响应内容
#     data = response.json()
#     messagebox.showinfo("返回值",data["message"])
#     # 检查HTTP状态码
#     if response.status_code == 200:
#         # 处理成功的响应
#         root.after(0, lambda: login_info_label.config(text=f"登录成功: {data['message']}"))
#         logger.info(f"登录成功:{data['message']}\n\n")
#     else:
#         # 处理错误的响应
#         root.after(0, lambda: login_info_label.config(text=f"登录失败: {data['error']}"))
#         logger.error(f"登录失败:{data['error']}\n\n")
#     login_button.config(state=tk.NORMAL)    

def handle_response(response, error):
    if error:
        root.after(0, lambda: login_info_label.config(text=f"网络请求失败(登录失败): {str(error)}"))
        logger.error(f"登录失败:{str(error)}\n\n")
        messagebox.showinfo("登陆失败", "请检查网络连接！\n日志：log.txt")
        login_button.config(state=tk.NORMAL)
        return

    logger.info(f"返回值：\n{response}\n")

    # 检查返回值中是否包含"result":"ok"
    if b'"result":"ok"' in response:
        if opt_out_var.get():
            countdown(4)
        else:
            root.after(0, lambda: login_info_label.config(text=f"登录成功！", font=("微软雅黑", 30)))
    elif b'"result":"fail"' in response:
        root.after(0, lambda: login_info_label.config(text=f"登录失败,服务器拒绝登陆！详情：{response}", font=("微软雅黑", 16)))
    elif b'"result":"no"' in response:
        root.after(0, lambda: login_info_label.config(text=f"登录失败,账号或密码错误！详情：{response}", font=("微软雅黑", 16)))
    else:
        root.after(0, lambda: login_info_label.config(text=f"登录失败,未知错误！详情：{response}", font=("微软雅黑", 16)))
    login_button.config(state=tk.NORMAL)
    
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
    Link = Link_mode_entry.get()
    timestamp = int(time.time() * 1000)
    # 获取复选框的状态
    auto_login = bool(auto_login_var.get())
    opt_out = bool(opt_out_var.get())
    save_password = bool(save_password_var.get())
    Link_mode = bool(Link_mode_var.get())
    
    

    with open(login_info_path, 'w') as f:
        json.dump({'username': username, 'password': password, 'Link':Link, 'Link_mode':Link_mode,'opt_out': opt_out, 'save_password': save_password,'auto_login': auto_login}, f)
    login_info_label.config(text=f"保存账号密码成功{timestamp}", font=("微软雅黑", 16))


def check_task():
    if check_task_exists("开机登陆校园网"):
        auto_start_button.config(text="关开机自启", command=remove_auto_start)
    else:
        auto_start_button.config(text="开机自启", command=set_auto_start)
    auto_start_button.update_idletasks()

def set_auto_start():
    # 获取当前exe程序的绝对路径
    exe_path = os.path.abspath(sys.argv[0])
    quoted_exe_path = shlex.quote(exe_path)
    # 创建任务计划
    result=subprocess.call(f'schtasks /Create /SC ONLOGON /TN "开机登陆校园网" /TR "{quoted_exe_path}" /F', shell=True)

    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()

    root_folder = scheduler.GetFolder('\\')

    task_definition = root_folder.GetTask('开机登陆校园网').Definition

    settings = task_definition.Settings

    settings.DisallowStartIfOnBatteries = False
    settings.StopIfGoingOnBatteries = False

    # 保存修改后的任务计划
    root_folder.RegisterTaskDefinition('开机登陆校园网', task_definition, 6, '', '', 3)
    if result == 0:
        login_info_label.config(text=f"创建开机自启动成功")
    else:
        login_info_label.config(text=f"创建开机自启动失败")
    messagebox.showinfo("提示！","移动文件位置后要重新设置哦！！")
    check_task()

def Get_administrator_privileges():
        save_info()
        messagebox.showinfo("提示！","将以管理员权限重新运行！！\n已自动保存账号密码")
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{__file__}" set_title', None, 0)
        os._exit(0)

def remove_auto_start():
    # 弹出询问框，询问是否要删除任务
    if messagebox.askyesno("确定？", "你确定要删除开机自启动任务吗？"):
        delete_result = subprocess.call('schtasks /Delete /TN "开机登陆校园网" /F', shell=True)
    if delete_result == 0:
        login_info_label.config(text=f"关闭开机自启动成功")
    else:
        login_info_label.config(text=f"关闭开机自启动失败")
    check_task()


# 判断是否拥有管理员权限
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

        
def open_log_folder():
    os.startfile(data_dir)
#检查是否有计划
def check_task_exists(task_name):
    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()
    root_folder = scheduler.GetFolder('\\')
    for task in root_folder.GetTasks(0):
        if task.Name == task_name:
            return True
    return False

# 文件过大时截断文件
def truncate_large_file(file_path, max_size=1024*1024*50):
    if os.path.getsize(file_path) > max_size:
        with open(file_path, 'w') as f:
            pass

# 创建主窗口
root = tk.Tk()
root.protocol("WM_DELETE_WINDOW", on_closing)
if is_admin():
    root.title("登录校园网 - 管理员")
else:
    root.title("登录校园网")

root.geometry("400x360")

# 获取用户的应用数据目录
appdata_dir = os.getenv('APPDATA')
data_dir = os.path.join(appdata_dir, 'Campus Network')
if not os.path.exists(data_dir):
    os.makedirs(data_dir)

# 创建一个logger
logger = logging.getLogger('my_logger')
logger.setLevel(logging.DEBUG)  # 设置日志级别

# 创建一个handler，用于写入日志文件
file_path = os.path.join(data_dir, 'log.txt')
file_handler = logging.FileHandler(file_path)
file_handler.setLevel(logging.DEBUG)
# 创建一个handler，用于输出到控制台
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)

# 定义handler的输出格式
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# 给logger添加handler
logger.addHandler(file_handler)
logger.addHandler(console_handler)

login_info_path = os.path.join(data_dir, 'login_info.json')

truncate_large_file(file_path)

label_label = tk.Label(root, text="不重启和断网线\n就不会掉线(除了服务器原因)", font=("微软雅黑", 14))
label_label.place(x=55, y=5)

open_log_button = tk.Button(root, text="打开日志", command=open_log_folder, font=("微软雅黑", 14))
open_log_button.place(x=10, y=225)

user_label = tk.Label(root, text="用户名:", font=("微软雅黑", 14))
user_label.place(x=25, y=65)

user_entry = tk.Entry(root, font=("微软雅黑", 14))
user_entry.place(x=125, y=65)

password_label = tk.Label(root, text="密码:", font=("微软雅黑", 14))
password_label.place(x=25, y=115)

password_entry = tk.Entry(root, font=("微软雅黑", 14), show="*")
password_entry.place(x=125, y=115)

Link_mode_var = tk.IntVar()
Link_mode_checkbutton = tk.Checkbutton(root, text="链接模式：", variable=Link_mode_var)
Link_mode_checkbutton.place(x=10, y=150)

Link_mode_entry = tk.Entry(root, font=("微软雅黑", 14))
Link_mode_entry.place(x=125, y=155)

opt_out_var = tk.IntVar()
opt_out_checkbutton = tk.Checkbutton(root, text="自动退出", variable=opt_out_var)
opt_out_checkbutton.place(x=15, y=185)
save_password_var = tk.IntVar()
save_password_checkbutton = tk.Checkbutton(root, text="保存密码", variable=save_password_var)
save_password_checkbutton.place(x=125, y=185)

auto_login_var = tk.IntVar()
auto_login_checkbutton = tk.Checkbutton(root, text="自动登录", variable=auto_login_var)
auto_login_checkbutton.place(x=225, y=185)

login_button = tk.Button(root, text="登录", font=("微软雅黑", 14), command=login)
login_button.place(x=125, y=225)

save_button = tk.Button(root, text="保存", font=("微软雅黑", 14), command=save_info)
save_button.place(x=200, y=225)

login_info_label = tk.Label(root, font=("微软雅黑", 14),wraplength=370)
login_info_label.place(x=15, y=275)

if is_admin():
    login_info_label.config(text=f"当前拥有“管理员权限”谨慎操作\n点击“开机自启”或“关闭开机自启”设置")
else:
    login_info_label.config(text=f"需要“管理员权限”才能设置开机自启\n点击“获取权限”按钮或手动获取权限")
    
auto_start_button = tk.Button(root, font=("微软雅黑", 14))
auto_start_button.place(x=275, y=225)


if is_admin():
    check_task()
else:
    auto_start_button.config(text="获取权限",command=Get_administrator_privileges)

# 如果存在'login_info.json'文件，则读取文件内容并更新用户输入和自动登录选项
if os.path.exists(login_info_path):
    with open(login_info_path, 'r') as f:
        login_info = json.load(f)
        user_entry.insert(0, login_info.get('username',''))
        password_entry.insert(0, login_info.get('password',''))
        Link_mode_entry.insert(0, login_info.get('Link',''))
        Link_mode_var.set(bool(login_info.get('Link_mode', 0)))
        auto_login_var.set(bool(login_info.get('auto_login', 0)))
        opt_out_var.set(bool(login_info.get('opt_out', 0)))
        save_password_var.set(bool(login_info.get('save_password', 0)))
# 如果自动登录选项被选中，则执行登录函数
if bool(auto_login_var.get()):
    login()

root.mainloop()
# 释放互斥体
ctypes.windll.kernel32.ReleaseMutex(mutex)