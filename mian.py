import tkinter as tk
import tkinter.messagebox as messagebox
import time
import os
import json
import requests
import os
import winshell

def login():

    # 获取用户名和密码
    timestamp = time.time() * 1000
    username = user_entry.get()
    password = password_entry.get()

    # 获取复选框的状态
    save_password = bool(save_password_var.get())
    auto_login = bool(auto_login_var.get())

    # 如果复选框被选中，保存用户名和密码
    # 如果 save_password 或 auto_login 为 True，则调用保存信息函数
    if save_password or auto_login:
        save_info()
    
    # 访问链接
    url = f"http://172.17.100.200:801/eportal/?c=GetMsg&a=loadToken&callback=jQuery_{timestamp}&account={username}&password={password}&mac=000000000000&_={timestamp}"
    response = requests.get(url)
    
    # 显示消息框
    messagebox.showinfo("登录信息：", f"用户名: {username}, 密码: {password}, 保存密码: {save_password}, 返回值: {response.text}")
    
    # 检查返回值中是否包含"result":"ok"
    if '"result":"ok"' in response.text:
        login_info_label.config(text=f"登录成功{timestamp}")
    else:
        login_info_label.config(text=f"登录失败{timestamp}")
    
def update_time():
    # 更新时间戳
    timestamp = int(time.time() * 1000)
    timestamp_label.config(text=f"当前时间戳(ms): {timestamp}")

    # 每秒更新一次
    root.after(1000, update_time)

def save_info():
    # 获取用户名和密码
    username = user_entry.get()
    password = password_entry.get()
    timestamp = int(time.time() * 1000)
    # 获取复选框的状态
    auto_login = bool(auto_login_var.get())

    # 保存用户名和密码
    with open('login_info.json', 'w') as f:
        json.dump({'username': username, 'password': password, 'auto_login': auto_login}, f)

    # 显示消息框
    login_info_label.config(text=f"保存成功{timestamp}")

def set_auto_start():
    # 获取当前exe程序的绝对路径
    exe_path = os.path.realpath("登陆校园网.exe")

    # 获取"启动"文件夹的路径
    startup_folder = winshell.startup()

    # 创建快捷方式的路径
    shortcut_path = os.path.join(startup_folder, '登陆校园网.lnk')

    # 创建快捷方式
    winshell.CreateShortcut(
        Path=shortcut_path,
        Target=exe_path,
        Icon=(exe_path, 0),
        Description="开机自动登陆校园网"
    )

    # 显示消息框
    messagebox.showinfo("提示！","移动位置后要重新设置哦！！")
    login_info_label.config(text=f"设置开机自启动成功")

def remove_auto_start():
    # 获取"启动"文件夹的路径
    startup_folder = winshell.startup()

    # 创建快捷方式的路径
    shortcut_path = os.path.join(startup_folder, '登陆校园网.lnk')

    # 如果快捷方式存在，则删除它
    if os.path.exists(shortcut_path):
        os.remove(shortcut_path)
        login_info_label.config(text=f"删除开机自启动成功")
    else:
        login_info_label.config(text=f"开机自启动未设置")

# 创建主窗口
root = tk.Tk()

# 设置窗口标题
root.title("登陆校园网")

# 设置窗口大小
root.geometry("430x270")

# 创建一个标签控件显示时间戳，并设置文本大小为15
timestamp_label = tk.Label(root,  font=("微软雅黑", 15))
timestamp_label.place(x=25, y=25)

# 创建一个标签控件作为用户名的提示，并设置文本大小为15
user_label = tk.Label(root, text="用户名:", font=("微软雅黑", 15))
user_label.place(x=25, y=75)

# 创建一个文本框控件用于输入用户名
user_entry = tk.Entry(root, font=("微软雅黑", 15))
user_entry.place(x=125, y=75)

# 创建一个标签控件作为密码的提示，并设置文本大小为15
password_label = tk.Label(root, text="密码:", font=("微软雅黑", 15))
password_label.place(x=25, y=125)

# 创建一个文本框控件用于输入密码
password_entry = tk.Entry(root, font=("微软雅黑", 15))
password_entry.place(x=125, y=125)

# 创建一个复选框控件作为保存密码的选项
save_password_var = tk.IntVar()
save_password_checkbutton = tk.Checkbutton(root, text="保存密码", variable=save_password_var)
save_password_checkbutton.place(x=125, y=155)

# 创建一个复选框控件作为自动登录的选项
auto_login_var = tk.IntVar()
auto_login_checkbutton = tk.Checkbutton(root, text="自动登录", variable=auto_login_var)
auto_login_checkbutton.place(x=225, y=155)

# 创建一个按钮控件作为登录按钮，并设置文本大小为15
login_button = tk.Button(root, text="登录", font=("微软雅黑", 15), command=login)
login_button.place(x=125, y=185)

# 创建一个按钮控件作为保存按钮，并设置文本大小为15
save_button = tk.Button(root, text="保存", font=("微软雅黑", 15), command=save_info)
save_button.place(x=200, y=185)

# 创建一个标签控件显示登录提示，并设置文本大小为15
login_info_label = tk.Label(root, font=("微软雅黑", 15))
login_info_label.place(x=30, y=235)

# 创建一个按钮控件作为设置开机自启动的按钮，并设置文本大小为15
auto_start_button = tk.Button(root, text="开机自启", font=("微软雅黑", 15), command=set_auto_start)
auto_start_button.place(x=15, y=185)

# 创建一个按钮控件作为删除开机自启动的按钮，并设置文本大小为15
remove_auto_start_button = tk.Button(root, text="删除开机自启", font=("微软雅黑", 15), command=remove_auto_start)
remove_auto_start_button.place(x=275, y=185)

# 如果存在保存的用户名和密码，自动填入
# 如果存在'login_info.json'文件，则读取文件内容并更新用户输入和自动登录选项
if os.path.exists('login_info.json'):
    with open('login_info.json', 'r') as f:
        login_info = json.load(f)
        user_entry.insert(0, login_info['username'])  # 将文件中的用户名更新到用户输入框
        password_entry.insert(0, login_info['password'])  # 将文件中的密码更新到密码输入框
        auto_login_var.set(bool(login_info.get('auto_login', 0)))  # 设置自动登录选项为文件中的值（如果不存在则默认为0）
# 如果自动登录选项被选中，则执行登录函数
if bool(auto_login_var.get()):
    login()
# 开始更新时间
update_time()

# 运行主事件循环
root.mainloop()