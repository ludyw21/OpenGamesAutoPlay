import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as ttkb
import ctypes
import platform

# 设置DPI感知，确保在高DPI显示器上显示正常
ctypes.windll.shcore.SetProcessDpiAwareness(1)
from ttkbootstrap.constants import *
import json
import os
import keyboard

class SettingsDialog:
    def __init__(self, parent, config_manager):
        self.parent = parent
        self.config_manager = config_manager
        self.dialog = ttkb.Toplevel(parent)
        self.dialog.title("设置")
        self.dialog.transient(parent)  # 设置为主窗口的子窗口
        self.dialog.grab_set()  # 模态窗口
        
        # 设置窗口大小和位置
        dialog_width = 600
        dialog_height = 350
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # 从配置中加载快捷键设置
        self.shortcuts = self.config_manager.data.get('shortcuts', {
            'START_PAUSE': 'alt+-',
            'STOP': 'alt+=',
            'PREV_SONG': 'alt+up',
            'NEXT_SONG': 'alt+down'
        })
        
        # 保存原始快捷键设置，用于取消操作
        self.original_shortcuts = self.shortcuts.copy()
        
        # 创建标签页控件
        self.notebook = ttk.Notebook(self.dialog)
        self.notebook.pack(fill=BOTH, expand=YES, padx=10, pady=10)
        
        # 创建按键设置标签页
        self.create_key_settings_tab()
        
        # 创建快捷键设置标签页
        self.create_shortcut_settings_tab()
        
        # 按钮区域
        self.create_button_area()
    
    def create_key_settings_tab(self):
        """创建按键设置标签页"""
        key_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(key_frame, text="按键设置")
        
        # 显示提示信息
        tip_label = ttk.Label(key_frame, text="按键设置功能暂未实现")
        tip_label.pack(pady=20)
    
    def get_modifier_options(self):
        """根据操作系统获取修饰键选项"""
        is_mac = platform.system() == 'Darwin'
        if is_mac:
            return ['Cmd', 'Option', 'Shift']
        else:
            return ['Ctrl', 'Alt', 'Shift']
    
    def parse_shortcut(self, shortcut_str):
        """解析快捷键字符串，返回修饰键1、修饰键2和主按键"""
        modifiers = self.get_modifier_options()
        parts = shortcut_str.split('+')
        mod1 = ''
        mod2 = ''
        key = ''
        
        # 找到主按键（不在修饰键列表中的部分）
        for part in parts:
            if part.lower() not in [m.lower() for m in modifiers]:
                key = part
                break
        
        # 提取修饰键
        mod_parts = [p for p in parts if p != key]
        if mod_parts:
            # 将修饰键映射回显示名称
            mod_map = {m.lower(): m for m in modifiers}
            mod_parts = [mod_map.get(p.lower(), p) for p in mod_parts]
            mod_parts.sort()  # 保持一致性
            mod1 = mod_parts[0]
            mod2 = mod_parts[1] if len(mod_parts) > 1 else '不使用'
        
        return mod1, mod2, key
    
    def create_shortcut_settings_tab(self):
        """创建快捷键设置标签页"""
        shortcut_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(shortcut_frame, text="快捷键设置")
        
        # 获取修饰键选项
        self.modifier_options = self.get_modifier_options()
        self.modifier2_options = ['不使用'] + self.modifier_options.copy()
        
        # 存储各个快捷键的变量
        self.shortcut_vars = {
            'START_PAUSE': {
                'mod1': tk.StringVar(),
                'mod2': tk.StringVar(value='不使用'),
                'key': tk.StringVar(),
                'display': tk.StringVar()
            },
            'STOP': {
                'mod1': tk.StringVar(),
                'mod2': tk.StringVar(value='不使用'),
                'key': tk.StringVar(),
                'display': tk.StringVar()
            },
            'PREV_SONG': {
                'mod1': tk.StringVar(),
                'mod2': tk.StringVar(value='不使用'),
                'key': tk.StringVar(),
                'display': tk.StringVar()
            },
            'NEXT_SONG': {
                'mod1': tk.StringVar(),
                'mod2': tk.StringVar(value='不使用'),
                'key': tk.StringVar(),
                'display': tk.StringVar()
            }
        }
        
        # 创建快捷键设置区域
        self.create_shortcut_setting(shortcut_frame, '播放/暂停', 'START_PAUSE')
        self.create_shortcut_setting(shortcut_frame, '停止播放', 'STOP')
        self.create_shortcut_setting(shortcut_frame, '上一首', 'PREV_SONG')
        self.create_shortcut_setting(shortcut_frame, '下一首', 'NEXT_SONG')
        
        # 提示信息
        tip_label = ttk.Label(
            shortcut_frame, 
            text="提示：支持F1-F12、A-Z、0-9及特殊字符等按键",
            foreground="#666"
        )
        tip_label.pack(pady=20)
    
    def create_shortcut_setting(self, parent, label_text, action_type):
        """创建单个快捷键设置行"""
        # 创建主框架
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=X, pady=10)
        
        # 左侧标签 - 固定宽度80像素
        ttk.Label(main_frame, text=f"{label_text}:", width=8).pack(side=LEFT, padx=5, pady=5)
        
        # 解析现有快捷键
        current_shortcut = self.shortcuts.get(action_type, '')
        mod1, mod2, key = self.parse_shortcut(current_shortcut)
        
        # 设置变量
        vars = self.shortcut_vars[action_type]
        if mod1:
            vars['mod1'].set(mod1)
        if mod2:
            vars['mod2'].set(mod2)
        if key:
            vars['key'].set(key)
        vars['display'].set(current_shortcut)
        
        # 修饰键1下拉框（必选）- 缩小宽度为6
        mod1_combo = ttk.Combobox(main_frame, textvariable=vars['mod1'], values=self.modifier_options, width=6, state='readonly')
        mod1_combo.pack(side=LEFT, padx=5, pady=5)
        mod1_combo.bind('<<ComboboxSelected>>', lambda e, act=action_type: self.handle_modifier_change(e, act, 'mod1'))
        
        # 加号标签
        ttk.Label(main_frame, text="+").pack(side=LEFT, padx=2)
        
        # 修饰键2下拉框（非必选）- 缩小宽度为6
        mod2_combo = ttk.Combobox(main_frame, textvariable=vars['mod2'], values=self.modifier2_options, width=6, state='readonly')
        mod2_combo.pack(side=LEFT, padx=5, pady=5)
        mod2_combo.bind('<<ComboboxSelected>>', lambda e, act=action_type: self.handle_modifier_change(e, act, 'mod2'))
        
        # 加号标签
        ttk.Label(main_frame, text="+").pack(side=LEFT, padx=2)
        
        # 主按键输入框
        key_entry = ttk.Entry(main_frame, textvariable=vars['key'], width=5)
        key_entry.pack(side=LEFT, padx=5, pady=5)
        key_entry.bind('<KeyRelease>', lambda e: self.limit_key_entry(e, action_type))
        
        # 显示最终快捷键的标签
        ttk.Label(main_frame, textvariable=vars['display'], foreground='#0066cc', width=15).pack(side=LEFT, padx=10)
    
    def handle_modifier_change(self, event, action_type, mod_type):
        """处理修饰键变更，确保两个修饰键互斥"""
        vars = self.shortcut_vars[action_type]
        current_value = event.widget.get()
        
        # 实现互斥逻辑
        if mod_type == 'mod1' and current_value in self.modifier_options:
            # 如果第一个修饰键不是空且第二个修饰键与第一个相同，设置第二个为不使用
            if vars['mod2'].get() == current_value:
                vars['mod2'].set('不使用')
        elif mod_type == 'mod2' and current_value in self.modifier_options:
            # 如果第二个修饰键不是不使用且与第一个相同，设置第一个为与第二个不同的值
            if vars['mod1'].get() == current_value:
                # 选择一个不同的修饰键
                for opt in self.modifier_options:
                    if opt != current_value:
                        vars['mod1'].set(opt)
                        break
        
        self.update_shortcut_display(action_type)
    
    def limit_key_entry(self, event, action_type):
        """限制按键输入框，支持单个字符和功能键"""
        entry = event.widget
        key = event.keysym
        
        # 处理功能键F1-F12
        if key.startswith('F') and len(key) <= 3:
            try:
                f_num = int(key[1:])
                if 1 <= f_num <= 12:
                    entry.delete(0, tk.END)
                    entry.insert(0, key)
                    self.update_shortcut_display(action_type)
                    return
            except ValueError:
                pass
        
        # 处理其他按键，只保留最后一个字符
        text = entry.get()
        if text:
            # 只保留最后一个可见字符
            entry.delete(0, tk.END)
            entry.insert(0, text[-1])
        self.update_shortcut_display(action_type)
    
    def update_shortcut_display(self, action_type):
        """更新快捷键显示并保存到配置"""
        vars = self.shortcut_vars[action_type]
        mod1 = vars['mod1'].get()
        mod2 = vars['mod2'].get()
        key = vars['key'].get()
        
        # 构建快捷键字符串
        parts = []
        if mod1:
            parts.append(mod1)
        if mod2 and mod2 != '不使用':
            parts.append(mod2)
        if key:
            parts.append(key)
        
        shortcut_str = '+'.join(parts)
        vars['display'].set(shortcut_str)
        
        # 更新快捷键配置
        self.shortcuts[action_type] = shortcut_str
    
    def _process_shortcut(self, action_type, shortcut):
        """处理快捷键（兼容旧代码）"""
        if shortcut:
            self.shortcuts[action_type] = shortcut
            # 更新显示
            if action_type in self.shortcut_vars:
                self.shortcut_vars[action_type]['display'].set(shortcut)
        
        # 重新启用主窗口的键盘钩子
        if hasattr(self.parent, 'update_keyboard_hooks'):
            self.parent.update_keyboard_hooks()
    
    def create_button_area(self):
        """创建按钮区域"""
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=X, padx=10, pady=10)
        
        # 取消按钮
        cancel_button = ttk.Button(button_frame, text="取消", command=self.cancel)
        cancel_button.pack(side=RIGHT, padx=5)
        
        # 保存设置按钮 - 更明确的按钮文本
        save_button = ttk.Button(button_frame, text="保存设置", command=self.save_settings, style="Primary.TButton")
        save_button.pack(side=RIGHT, padx=5)
    
    def save_settings(self):
        """保存设置"""
        # 获取所有快捷键
        action_types = ['START_PAUSE', 'STOP', 'PREV_SONG', 'NEXT_SONG']
        
        # 验证所有快捷键
        for action_type in action_types:
            shortcut = self.shortcuts.get(action_type, '')
            
            # 检查是否为空
            if not shortcut:
                messagebox.showerror("错误", f"请完整设置所有快捷键")
                return
            
            # 检查是否包含修饰键
            if '+' not in shortcut:
                messagebox.showerror("错误", f"快捷键必须是组合键（至少包含一个修饰键和一个主按键）")
                return
        
        # 验证快捷键是否与keyboard库兼容
        try:
            import keyboard
            # 标准化修饰键名称以适配keyboard库
            def normalize_shortcut(shortcut):
                # 将显示名称转换为keyboard库识别的名称
                replacements = {
                    'Ctrl': 'ctrl',
                    'Alt': 'alt',
                    'Shift': 'shift',
                    'Cmd': 'cmd' if platform.system() == 'Darwin' else 'win'
                }
                parts = shortcut.split('+')
                normalized_parts = [replacements.get(p, p) for p in parts]
                return '+'.join(normalized_parts)
            
            # 测试每个快捷键
            for action_type in action_types:
                shortcut = self.shortcuts[action_type]
                normalized = normalize_shortcut(shortcut)
                keyboard.add_hotkey(normalized, lambda: None)
                keyboard.remove_hotkey(normalized)
        except Exception as e:
            messagebox.showerror("快捷键错误", f"快捷键格式无效或不兼容: {str(e)}")
            return
        
        # 保存到配置文件
        config_data = self.config_manager.data
        config_data['shortcuts'] = self.shortcuts
        self.config_manager.save(config_data)
        
        # 通知主窗口更新键盘钩子，确保快捷键立即生效
        if hasattr(self.parent, 'update_keyboard_hooks'):
            self.parent.update_keyboard_hooks()
        
        # 关闭窗口
        self.dialog.destroy()
    
    def cancel(self):
        """取消设置"""
        self.shortcuts = self.original_shortcuts.copy()
        self.dialog.destroy()
