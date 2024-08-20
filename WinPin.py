import tkinter as tk
from tkinter import font
from tkinter import ttk
import pygetwindow as gw
import win32gui
import win32con
import winreg
import sys

# pyinstaller -F -w -i C:\Users\ASUS\Desktop\mycode\cs_time\WinPin\image\favicon.ico C:\Users\ASUS\Desktop\mycode\cs_time\WinPin\WinPin.py

# pyinstaller --onefile --windowed --icon=.\image\favicon.ico .\WinPin.py
# pyinstaller --onefile --windowed --icon=.\image\android-chrome-192x192.png .\WinPin.py
# pyinstaller --onefile --windowed --add-data=".\image\favicon-32x32.png" --icon=.\image\favicon-32x32.png .\WinPin.py

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows Pin Manager")
        self.root.attributes('-topmost', 1)
        self.is_hidden = False  # 用于跟踪窗口是否隐藏
        # 设置字体样式
        self.font_style = font.Font(family='Arial', size=12, weight='bold')

        # 设置窗口的最小尺寸
        self.root.minsize(400, 300)

        # 创建一个框架用于包含两个列表框，以实现上下布局
        self.frame_listboxes = ttk.Frame(root, padding="10")
        self.frame_listboxes.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=15, pady=10)
        self.listbox_items = {}
        # 创建两个列表框，一个用于普通窗口，一个用于已置顶的窗口
        self.listbox_normal = tk.Listbox(self.frame_listboxes, height=9, width=50, font=self.font_style)
        self.listbox_normal.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=20)

        self.listbox_topmost = tk.Listbox(self.frame_listboxes, height=10, width=50, font=self.font_style)
        self.listbox_topmost.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        # 为两个列表框添加滚动条
        self.scroll_normal = ttk.Scrollbar(self.frame_listboxes, orient=tk.VERTICAL, command=self.listbox_normal.yview)
        self.scroll_topmost = ttk.Scrollbar(self.frame_listboxes, orient=tk.VERTICAL,
                                            command=self.listbox_topmost.yview)

        self.listbox_normal.bind('<<ListboxSelect>>', self.on_select)
        self.listbox_normal.bind('<Double-1>', self.on_double_click)
        self.listbox_topmost.bind('<<ListboxSelect>>', self.on_select_topmost)
        self.listbox_topmost.bind('<Double-1>', self.on_double_click)

        # 信息说明栏目和功能按钮
        self.info_frame = ttk.Frame(root, padding="10")
        self.info_frame.pack(side=tk.LEFT, fill=tk.BOTH)

        # 添加“刷新当前列表”按钮
        self.refresh_button = ttk.Button(self.info_frame, text="\n  刷  新\n当前列表\n", command=self.refresh_listbox)
        self.refresh_button.pack(side=tk.LEFT, fill=tk.X, padx=10)
        # 添加“显示钉选窗口”按钮
        self.show_pinned_button = ttk.Button(self.info_frame, text="\n  显  示\n钉选窗口\n", command=self.show_pinned_windows)
        self.show_pinned_button.pack(side=tk.LEFT, fill=tk.X, padx=10)
        # 将开机自启动按钮替换为“最小化所有窗口”按钮
        self.minimize_button = ttk.Button(self.info_frame, text="\n 最小化\n所有窗口\n", command=self.minimize_all_windows)
        self.minimize_button.pack(side=tk.RIGHT, fill=tk.X, padx=10)

        # 使用 Text 组件显示多行信息，并允许自动换行
        self.info_text = tk.Text(self.info_frame, height=4, width=40, font=self.font_style, wrap=tk.WORD)
        self.info_text.insert(tk.END, "\n      使用Alt+B可以快速最小化该客户端，\n      若多个窗口钉选成功，则会在同一层面上。")
        self.info_text.config(state=tk.DISABLED)  # 使文本框不可编辑
        self.info_text.pack(side=tk.TOP, fill=tk.X, padx=15)

        # 绑定 Alt+b
        self.root.bind('<Alt-b>', self.toggle_show_hide)

        self.populate_listbox()

    def populate_listbox(self):
        self.listbox_normal.delete(0, tk.END)
        for window in gw.getAllWindows():
            if window.title:
                self.listbox_normal.insert(tk.END, window.title)
                # 存储窗口标题和对应的句柄
                self.listbox_items[window.title] = window._hWnd

    def on_select(self, event):
        selection = self.listbox_normal.curselection()
        if selection:
            index = selection[0]
            title = self.listbox_normal.get(index)
            hwnd = self.listbox_items.get(title)
            if hwnd:
                self.set_window_always_on_top(hwnd)
                # 移动到置顶列表
                self.listbox_normal.delete(index)
                self.listbox_topmost.insert(tk.END, f"[ 已置顶钉选 ] {title}")

    def on_select_topmost(self, event):
        # 从置顶列表移除并设置为非置顶
        selection = self.listbox_topmost.curselection()
        if selection:
            index = selection[0]
            title = self.listbox_topmost.get(index)[len("[ 已置顶钉选 ] "):]
            hwnd = self.listbox_items.get(title)
            if hwnd:
                self.unset_window_always_on_top(hwnd)
                self.listbox_topmost.delete(index)
                self.listbox_normal.insert(tk.END, title)

    def on_double_click(self, event):
        listbox = event.widget
        index = listbox.nearest(event.y)
        title = listbox.get(index)
        if title.startswith("[ 已置顶钉选 ] "):
            original_title = title[len("[ 已置顶钉选 ] "):]
            hwnd = self.listbox_items.get(original_title)
            if hwnd:
                self.unset_window_always_on_top(hwnd)
                # 移动到普通列表
                listbox.delete(index)
                self.listbox_normal.insert(tk.END, original_title)
        else:
            self.on_select(event)  # 处理普通列表的双击事件

    def set_window_always_on_top(self, hwnd):
        # 确保这里实现的是设置窗口为总是置顶的逻辑
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                               win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) | win32con.WS_EX_TOPMOST)
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    def unset_window_always_on_top(self, hwnd):
        # 确保这里实现的是取消窗口总是置顶的逻辑
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                               win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE) & ~win32con.WS_EX_TOPMOST)
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    def refresh_listbox(self):
        # 清空现有列表
        self.listbox_normal.delete(0, tk.END)
        # 重新获取窗口信息并填充列表
        for window in gw.getAllWindows():
            if window.title:
                self.listbox_normal.insert(tk.END, window.title)
                # 更新存储的窗口标题和对应的句柄
                self.listbox_items[window.title] = window._hWnd

    def show_pinned_windows(self):
        # 最小化所有顶级窗口
        def minimize_all(hwnd, top_windows):
            if win32gui.IsWindowVisible(hwnd) and hwnd != self.root.winfo_id():
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            return True

        # 遍历所有顶级窗口并最小化
        win32gui.EnumWindows(minimize_all, None)

        # 遍历置顶列表中的窗口，并将它们设置为可见
        for i in range(self.listbox_topmost.size()):
            title = self.listbox_topmost.get(i)
            if title.startswith("[ 已置顶钉选 ] "):
                # 去除前缀以获取原始窗口标题
                original_title = title[len("[ 已置顶钉选 ] "):]
                hwnd = self.listbox_items.get(original_title)
                if hwnd:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        # 最后，确保Tkinter窗口是可见的
        self.root.deiconify()

    def toggle_show_hide(self, event):
        # 切换窗口的显示状态
        if self.root.state() == 'normal':
            # 如果窗口是正常状态，我们将其最小化
            self.root.iconify()

    def minimize_all_windows(self):
        def is_not_console(hwnd):
            # 获取窗口类名
            class_name = win32gui.GetClassName(hwnd)
            # 检查窗口是否不是控制台窗口
            return class_name != "ConsoleWindowClass"

        def minimize_window(hwnd, top_windows):
            if is_not_console(hwnd) and win32gui.IsWindowVisible(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            return True

        win32gui.EnumWindows(minimize_window, None)


def toggle_show_hide(self, event):
    if self.root.state() == 'normal':
        self.root.withdraw()  # 隐藏窗口
    else:
        self.root.deiconify()  # 显示窗口


def toggle_autostart(self):
    # 切换开机自启动设置
    app_name = "WindowManager"
    with winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER) as reg:
        startup_key = winreg.OpenKey(reg, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_ALL_ACCESS)
        try:
            winreg.SetValueEx(startup_key, app_name, 0, winreg.REG_SZ, sys.executable)
            self.autostart_button.config(text="已设置开机自启动")
        except FileNotFoundError:
            winreg.DeleteValue(startup_key, app_name)
            self.autostart_button.config(text="取消开机自启动")
        winreg.CloseKey(startup_key)


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
