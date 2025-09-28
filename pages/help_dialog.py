import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *

class HelpDialog:
    def __init__(self, parent):
        self.parent = parent
        self.dialog = ttkb.Toplevel(parent)
        self.dialog.title("帮助")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)  # 设置为主窗口的子窗口
        self.dialog.grab_set()  # 模态窗口
        
        # 设置窗口大小和位置
        dialog_width = 400
        dialog_height = 300
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # 创建标签页控件
        notebook = ttk.Notebook(self.dialog)
        notebook.pack(fill=BOTH, expand=YES, padx=10, pady=10)
        
        # 使用说明标签页
        usage_frame = ttk.Frame(notebook, padding=10)
        notebook.add(usage_frame, text="使用说明")
        
        usage_text = "1. 使用管理员权限启动\n" + \
                     "2. 选择MIDI文件和音轨\n" + \
                     "3. 点击播放按钮开始演奏\n" + \
                     "4. 支持36键模式"
        
        usage_label = ttk.Label(usage_frame, text=usage_text, justify=LEFT)
        usage_label.pack(fill=BOTH, expand=YES)
        
        # 快捷键说明标签页 - 已隐藏，使用设置对话框配置快捷键
        
        # 确定按钮
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=X, padx=10, pady=10)
        
        ok_button = ttk.Button(button_frame, text="确定", command=self.dialog.destroy)
        ok_button.pack(side=RIGHT, padx=5)