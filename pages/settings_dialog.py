import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as ttkb
import ctypes
import platform
from ttkbootstrap.constants import *
import json
import os
import keyboard

# 设置DPI感知，确保在高DPI显示器上显示正常
ctypes.windll.shcore.SetProcessDpiAwareness(1)

class SettingsDialog:
    def __init__(self, parent, config_manager):
        self.parent = parent
        self.config_manager = config_manager
        self.dialog = ttkb.Toplevel(parent.root)
        self.dialog.title("设置")
        self.dialog.transient(parent.root)  # 设置为主窗口的子窗口
        self.dialog.grab_set()  # 模态窗口
        
        # 获取DPI缩放比例
        dpi_scale = 1.0
        if hasattr(os, 'name') and os.name == 'nt':
            try:
                import ctypes
                dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
                print(f"设置窗口DPI缩放比例: {dpi_scale:.2f}")
            except Exception as e:
                print(f"设置窗口获取DPI缩放比例失败: {str(e)}")
        
        # 根据DPI缩放比例计算窗口大小，位置
        base_width, base_height = 518, 548  # 增加高度20
        dialog_width = int(base_width * dpi_scale)
        dialog_height = int(base_height * dpi_scale)
        
        parent_x = parent.root.winfo_x()
        parent_y = parent.root.winfo_y()
        parent_width = parent.root.winfo_width()
        parent_height = parent.root.winfo_height()
        
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
        
        # 创建按钮区域
        self.create_button_area()
    
    def create_key_settings_tab(self):
        """创建按键设置标签页"""
        from groups import GROUPS, NOTE_NAMES, get_note_name
        from keyboard_mapping import NOTE_TO_KEY
        
        key_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(key_frame, text="按键设置")
        
        # 从配置加载按键设置
        self.key_settings = self.config_manager.data.get('key_settings', {})
        self.note_to_key = NOTE_TO_KEY.copy()
        
        # 更新用户自定义按键
        if 'note_to_key' in self.key_settings:
            self.note_to_key.update(self.key_settings['note_to_key'])
        
        # 定义预设配置
        self.preset_configs = {
            "燕云十六声(36键)": {
                "min_note": 48,
                "max_note": 83,
                "black_key_mode": "support_black_key",
                "note_to_key": NOTE_TO_KEY  # 使用完整的NOTE_TO_KEY映射
            },
            "燕云十六声(21键)": {
                "min_note": 48,
                "max_note": 83,
                "black_key_mode": "auto_sharp",
                "note_to_key": {note: key for note, key in NOTE_TO_KEY.items() if '+' not in key}  # 只包含单个按键，不包含组合键
            }
        }
        
        # 添加快捷设置区域在基础配置上方
        preset_frame = ttk.Frame(key_frame)
        preset_frame.pack(fill=X, padx=5, pady=10, anchor=W)
        
        # 快捷设置标签
        ttk.Label(preset_frame, text="快捷设置：").pack(side=LEFT, padx=5, pady=5)
        
        # 预设下拉框
        self.preset_var = tk.StringVar()
        preset_combo = ttk.Combobox(preset_frame, textvariable=self.preset_var, 
                                   values=list(self.preset_configs.keys()), width=20, state='readonly')
        # 默认选择第一个选项
        if self.preset_configs:
            preset_combo.current(0)
        preset_combo.pack(side=LEFT, padx=5, pady=5)
        
        # 确认按钮
        confirm_button = ttk.Button(preset_frame, text="确定", command=self.apply_preset_config)
        confirm_button.pack(side=LEFT, padx=5, pady=5)
        
        # 基础配置组
        base_frame = ttk.LabelFrame(key_frame, text="基础配置")
        base_frame.pack(fill=X, padx=5, pady=5)
        
        # 创建最低音、最高音和黑键设置 - 使用grid布局
        config_container = ttk.Frame(base_frame)
        config_container.pack(fill=X, expand=YES, padx=10, pady=10)
        
        # 左侧 - 音域设置
        notes_frame = ttk.Frame(config_container)
        # 设置相同的内边距和对齐方式，确保顶部对齐
        notes_frame.grid(row=0, column=0, padx=(10, 30), pady=5, sticky=N)
        
        # 最低音级联选择
        min_note_frame = ttk.Frame(notes_frame)
        # 移除额外的内边距，确保顶部对齐
        min_note_frame.pack(fill=X, pady=0)
        ttk.Label(min_note_frame, text="最低音:").grid(row=0, column=0, padx=5, pady=5, sticky=W)
        # 最低音分组下拉框
        self.min_group_var = tk.StringVar()
        self.min_group_combo = ttk.Combobox(min_note_frame, textvariable=self.min_group_var, width=15, state='readonly')
        self.min_group_combo.grid(row=0, column=1, padx=2, pady=5)
        # 最低音音符下拉框
        self.min_note_var = tk.StringVar()
        self.min_note_combo = ttk.Combobox(min_note_frame, textvariable=self.min_note_var, width=10, state='readonly')
        self.min_note_combo.grid(row=0, column=2, padx=2, pady=5)
        
        # 最高音级联选择
        max_note_frame = ttk.Frame(notes_frame)
        max_note_frame.pack(fill=X, pady=5)
        ttk.Label(max_note_frame, text="最高音:").grid(row=0, column=0, padx=5, pady=5, sticky=W)
        # 最高音分组下拉框
        self.max_group_var = tk.StringVar()
        self.max_group_combo = ttk.Combobox(max_note_frame, textvariable=self.max_group_var, width=15, state='readonly')
        self.max_group_combo.grid(row=0, column=1, padx=2, pady=5)
        # 最高音音符下拉框
        self.max_note_var = tk.StringVar()
        self.max_note_combo = ttk.Combobox(max_note_frame, textvariable=self.max_note_var, width=10, state='readonly')
        self.max_note_combo.grid(row=0, column=2, padx=2, pady=5)
        
        # 右侧 - 黑键设置单选组
        black_key_frame = ttk.Frame(config_container)
        # 增加左侧边距以增加间隔，确保与左侧顶部对齐
        black_key_frame.grid(row=0, column=1, padx=10, pady=5, sticky=N)
        self.black_key_mode = tk.StringVar(value=self.key_settings.get('black_key_mode', 'support_black_key'))
        # 调整第一个单选按钮的padding，确保与左侧顶部对齐
        ttk.Radiobutton(black_key_frame, text="支持黑键", variable=self.black_key_mode, 
                       value="support_black_key", command=self.update_keyboard_settings).pack(side=TOP, pady=5, anchor=W)
        ttk.Radiobutton(black_key_frame, text="黑键自动降音", variable=self.black_key_mode, 
                       value="auto_sharp", command=self.update_keyboard_settings).pack(side=TOP, pady=5, anchor=W)
        
        # 创建分组下拉框数据结构
        self.create_group_combobox_data()
        
        # 绑定事件
        self.min_group_combo.bind('<<ComboboxSelected>>', lambda e: self.on_group_selected('min'))
        self.max_group_combo.bind('<<ComboboxSelected>>', lambda e: self.on_group_selected('max'))
        self.min_note_combo.bind('<<ComboboxSelected>>', self.on_note_range_change)
        self.max_note_combo.bind('<<ComboboxSelected>>', self.on_note_range_change)
        
        # 按键设置区域
        self.keys_frame = ttk.LabelFrame(key_frame, text="按键配置")
        self.keys_frame.pack(fill=BOTH, expand=YES, padx=5, pady=5)
        
        # 初始更新按键设置
        self.update_keyboard_settings()
        

    
    def create_group_combobox_data(self):
        """创建分组级联下拉框的数据结构"""
        from groups import GROUPS, get_note_name
        
        # 存储分组数据
        self.group_note_data = {}
        
        # 获取所有分组名称
        group_names = list(GROUPS.keys())
        
        # 设置分组下拉框值
        self.min_group_combo['values'] = group_names
        self.max_group_combo['values'] = group_names
        
        # 为每个分组准备音符数据
        for group_name, (start, end) in GROUPS.items():
            group_notes = []
            note_data = {}
            
            for note_num in range(start, end + 1):
                note_name = get_note_name(note_num)
                option_text = f"{note_name}({note_num})"
                group_notes.append(option_text)
                note_data[option_text] = note_num
            
            self.group_note_data[group_name] = (group_notes, note_data)
        
        # 根据默认音符值设置初始选择
        default_min_note = self.key_settings.get('min_note', 48)
        default_max_note = self.key_settings.get('max_note', 83)
        
        # 找到默认音符所在的分组
        for group_name, (start, end) in GROUPS.items():
            if start <= default_min_note <= end:
                self.min_group_var.set(group_name)
                # 触发分组选择事件以填充音符下拉框
                self.on_group_selected('min')
                # 设置默认音符
                for note_text, note_num in self.group_note_data[group_name][1].items():
                    if note_num == default_min_note:
                        self.min_note_var.set(note_text)
                        break
            
            if start <= default_max_note <= end:
                self.max_group_var.set(group_name)
                # 触发分组选择事件以填充音符下拉框
                self.on_group_selected('max')
                # 设置默认音符
                for note_text, note_num in self.group_note_data[group_name][1].items():
                    if note_num == default_max_note:
                        self.max_note_var.set(note_text)
                        break
    
    def on_group_selected(self, prefix):
        """处理分组选择事件"""
        if prefix == 'min':
            group_var = self.min_group_var
            note_combo = self.min_note_combo
        else:
            group_var = self.max_group_var
            note_combo = self.max_note_combo
        
        selected_group = group_var.get()
        if selected_group and selected_group in self.group_note_data:
            # 填充音符下拉框
            group_notes = self.group_note_data[selected_group][0]
            note_combo['values'] = group_notes
            # 如果之前没有选择音符，默认选择第一个
            if not note_combo.get() and group_notes:
                note_combo.set(group_notes[0])
    
    def get_selected_note_number(self, prefix):
        """获取选中的音符编号"""
        if prefix == 'min':
            group = self.min_group_var.get()
            note = self.min_note_var.get()
        else:
            group = self.max_group_var.get()
            note = self.max_note_var.get()
        
        if group and note and group in self.group_note_data:
            note_data = self.group_note_data[group][1]
            return note_data.get(note)
        return None
    
    def on_note_range_change(self, event=None):
        """处理音域范围变更"""
        # 获取选择的音符值
        min_note = self.get_selected_note_number('min')
        max_note = self.get_selected_note_number('max')
        
        # 验证并修正范围
        if min_note is not None and max_note is not None:
            if min_note > max_note:
                # 如果最小值大于最大值，交换它们的值在显示上
                min_group = self.min_group_var.get()
                max_group = self.max_group_var.get()
                min_note_text = self.min_note_var.get()
                max_note_text = self.max_note_var.get()
                
                # 交换显示
                self.min_group_var.set(max_group)
                self.max_group_var.set(min_group)
                self.min_note_var.set(max_note_text)
                self.max_note_var.set(min_note_text)
                
                # 重新加载音符下拉框
                self.on_group_selected('min')
                self.on_group_selected('max')
        
        # 更新按键设置
        self.update_keyboard_settings()
    
    def update_keyboard_settings(self):
        """更新按键设置显示"""
        from groups import GROUPS, get_note_name
        
        # 清空现有内容
        for widget in self.keys_frame.winfo_children():
            widget.destroy()
        
        # 解析选择的音符值
        min_note = self.get_selected_note_number('min')
        max_note = self.get_selected_note_number('max')
        
        # 如果没有选择，使用默认值
        if min_note is None:
            min_note = 48
        if max_note is None:
            max_note = 83
        
        # 确定当前选择的分组
        visible_groups = []
        for group_name, (group_start, group_end) in GROUPS.items():
            # 检查分组是否与选择范围有重叠
            if not (group_end < min_note or group_start > max_note):
                visible_groups.append((group_name, group_start, group_end))
        
        # 按音高排序分组
        visible_groups.sort(key=lambda x: x[1])
        
        # 创建水平和垂直滚动条，设置为按需显示
        h_scrollbar = ttk.Scrollbar(self.keys_frame, orient="horizontal")
        v_scrollbar = ttk.Scrollbar(self.keys_frame, orient="vertical")
        
        # 创建画布，确保滚动条命令正确配置
        canvas = tk.Canvas(self.keys_frame, 
                          xscrollcommand=h_scrollbar.set, 
                          yscrollcommand=v_scrollbar.set,
                          highlightthickness=0)
        h_scrollbar.config(command=canvas.xview)
        v_scrollbar.config(command=canvas.yview)
        
        scrollable_frame = ttk.Frame(canvas)
        
        # 配置滚动区域并处理滚动条显示逻辑
        def update_scrollbars(event=None):
            # 更新画布的滚动区域
            canvas.configure(scrollregion=canvas.bbox("all"))
            
            # 获取画布的实际大小和内容大小
            canvas_width = canvas.winfo_width()
            canvas_height = canvas.winfo_height()
            content_width = scrollable_frame.winfo_width()
            content_height = scrollable_frame.winfo_height()
            
            # 垂直滚动条按需显示
            if content_height > canvas_height:
                v_scrollbar.pack(side="right", fill="y")
                # 调整画布宽度，为垂直滚动条留出空间
                canvas.pack_configure(fill="both", expand=True, side="left")
            else:
                v_scrollbar.pack_forget()
                # 画布占据所有可用空间
                canvas.pack_configure(fill="both", expand=True, side="left")
            
            # 水平滚动条按需显示
            if content_width > canvas_width:
                h_scrollbar.pack(side="bottom", fill="x", before=canvas)
            else:
                h_scrollbar.pack_forget()
        
        # 绑定事件以更新滚动条
        scrollable_frame.bind("<Configure>", update_scrollbars)
        canvas.bind("<Configure>", update_scrollbars)
        
        # 使用正确的锚点创建窗口
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # 确保画布先于滚动条布局
        canvas.pack(side="left", fill="both", expand=True)
        
        # 初始调用一次更新函数确保正确显示
        self.keys_frame.after(100, update_scrollbars)
        
        # 创建按键配置区域
        self.note_entries = {}
        black_key_mode = self.black_key_mode.get()
        
        # 创建一个容器框架用于水平排列LabelFrame - 确保从左侧显示
        groups_container = ttk.Frame(scrollable_frame)
        groups_container.pack(fill=BOTH, expand=YES, padx=10, pady=10, anchor=W)
        
        # 使用grid布局来水平排列分组
        for i, (group_name, group_start, group_end) in enumerate(visible_groups):
            # 确保分组在选择范围内
            actual_start = max(group_start, min_note)
            actual_end = min(group_end, max_note)
            
            # 使用LabelFrame替代普通Label作为分组容器
            group_frame = ttk.LabelFrame(groups_container, text=group_name, relief=RAISED)
            # 使用grid布局，设置sticky=N确保顶部对齐
            group_frame.grid(row=0, column=i, padx=10, pady=5, sticky=N)
            
            # 根据黑键模式决定显示的音符
            for note_num in range(actual_start, actual_end + 1):
                # 检查是否跳过黑键
                if black_key_mode == 'auto_sharp':
                    # 获取原始音符名（不包含八度符号）进行检查
                    # 使用NOTE_NAMES直接从音符索引获取原始名称
                    from groups import NOTE_NAMES
                    note_index = note_num % 12
                    raw_note_name = NOTE_NAMES[note_index]
                    if '#' in raw_note_name:
                        # 只跳过带#的黑键，保留所有自然音符（包括b音）
                        continue
                
                # 创建音符行
                row_frame = ttk.Frame(group_frame)
                row_frame.pack(fill=X, pady=2)
                
                # 音符标签 - 减少宽度
                note_name = get_note_name(note_num)
                note_label = ttk.Label(row_frame, text=f"{note_name}({note_num}):", width=8, anchor=W)
                note_label.pack(side=LEFT, padx=5)
                
                # 按键输入框 - 减少宽度
                key_var = tk.StringVar(value=self.note_to_key.get(note_num, ''))
                key_entry = ttk.Entry(row_frame, textvariable=key_var, width=10)
                key_entry.pack(side=LEFT, padx=5, fill=X, expand=True)
                
                # 存储输入框引用
                self.note_entries[note_num] = key_var
    
    def create_button_area(self):
        """创建按钮区域"""
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=X, padx=10, pady=10)
        
        # 取消按钮
        cancel_button = ttk.Button(button_frame, text="取消", command=self.cancel)
        cancel_button.pack(side=RIGHT, padx=5)
        
        # 保存设置按钮
        save_button = ttk.Button(button_frame, text="保存设置", command=self.save_settings, style="Primary.TButton")
        save_button.pack(side=RIGHT, padx=5)
    
    def cancel(self):
        """取消操作"""
        self.dialog.destroy()
    
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
        
        # 恢复默认按钮
        restore_frame = ttk.Frame(shortcut_frame)
        restore_frame.pack(fill=X, pady=10)
        restore_button = ttk.Button(restore_frame, text="恢复默认", command=self.restore_default_shortcuts)
        restore_button.pack(side=RIGHT, padx=10, pady=5)
    
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
        """保存设置（整合按键设置和快捷键设置）"""
        # 先验证快捷键设置
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
        
        # 保存按键设置
        from groups import GROUPS, get_note_name
        
        # 解析选择的音符值
        min_note = self.get_selected_note_number('min')
        max_note = self.get_selected_note_number('max')
        
        # 如果没有选择，使用默认值
        if min_note is None:
            min_note = 48
        if max_note is None:
            max_note = 83
        
        # 收集音符按键设置
        note_to_key = {}
        for note_num, var in self.note_entries.items():
            key_value = var.get().strip()
            if key_value:
                note_to_key[note_num] = key_value
        
        # 构建按键设置
        key_settings = {
            'min_note': min_note,
            'max_note': max_note,
            'black_key_mode': self.black_key_mode.get(),
            'note_to_key': note_to_key
        }
        
        # 更新配置数据
        config_data = self.config_manager.data
        config_data['shortcuts'] = self.shortcuts
        config_data['key_settings'] = key_settings
        
        # 保存配置
        self.config_manager.save(config_data)
        
        # 通知主窗口更新键盘钩子，确保快捷键立即生效
        if hasattr(self.parent, 'update_keyboard_hooks'):
            print("[DEBUG] 调用update_keyboard_hooks方法")
            self.parent.update_keyboard_hooks()
        
        # 通知主窗口更新音轨分析标题
        if hasattr(self.parent, 'update_analysis_frame_title'):
            print("[DEBUG] 调用update_analysis_frame_title方法")
            self.parent.update_analysis_frame_title()
        else:
            print("[DEBUG] update_analysis_frame_title方法不存在")
        
        # 显示成功消息
        # messagebox.showinfo("成功", "设置已保存")
        
        # 关闭窗口
        self.dialog.destroy()
    
    def apply_preset_config(self):
        """应用预设配置"""
        try:
            # 获取选择的预设配置
            selected_preset = self.preset_var.get()
            if not selected_preset or selected_preset not in self.preset_configs:
                return
            
            preset = self.preset_configs[selected_preset]
            
            # 更新配置数据
            config_data = self.config_manager.data
            key_settings = {
                'min_note': preset['min_note'],
                'max_note': preset['max_note'],
                'black_key_mode': preset['black_key_mode'],
                'note_to_key': {str(note): key for note, key in preset['note_to_key'].items()}  # 转换为字符串键
            }
            config_data['key_settings'] = key_settings
            
            # 保存配置
            self.config_manager.save(config_data)
            
            # 更新当前设置
            self.key_settings = key_settings
            self.note_to_key = preset['note_to_key']
            
            # 更新UI
            self.black_key_mode.set(preset['black_key_mode'])
            
            # 更新最低音和最高音选择
            from groups import GROUPS, group_for_note, get_note_name
            
            # 为最低音找到对应的分组
            min_group = group_for_note(preset['min_note'])
            
            # 直接设置分组值并触发事件
            self.min_group_var.set(min_group)
            self.on_group_selected('min')
            
            # 设置默认音符
            for note_text, note_num in self.group_note_data[min_group][1].items():
                if note_num == preset['min_note']:
                    self.min_note_var.set(note_text)
                    break
            
            # 为最高音找到对应的分组
            max_group = group_for_note(preset['max_note'])
            
            # 直接设置分组值并触发事件
            self.max_group_var.set(max_group)
            self.on_group_selected('max')
            
            # 设置默认音符
            for note_text, note_num in self.group_note_data[max_group][1].items():
                if note_num == preset['max_note']:
                    self.max_note_var.set(note_text)
                    break
            
            # 更新按键设置显示
            self.update_keyboard_settings()
            
            # 通知主窗口更新音轨分析标题
            if hasattr(self.parent, 'update_analysis_frame_title'):
                self.parent.update_analysis_frame_title()
            
            # 移除成功消息弹窗，直接刷新页面
            
        except Exception as e:
            print(f"应用预设配置时出错: {str(e)}")
            messagebox.showerror("错误", f"应用预设配置时出错: {str(e)}")
    
    def cancel(self):
        """取消设置"""
        self.shortcuts = self.original_shortcuts.copy()
        self.dialog.destroy()
    
    def restore_default_shortcuts(self):
        """恢复默认快捷键设置"""
        from keyboard_mapping import CONTROL_KEYS
        
        try:
            # 应用默认快捷键
            for action_type, default_shortcut in CONTROL_KEYS.items():
                if action_type in self.shortcuts and action_type in self.shortcut_vars:
                    # 更新配置
                    self.shortcuts[action_type] = default_shortcut
                    
                    # 更新UI变量
                    vars = self.shortcut_vars[action_type]
                    mod1, mod2, key = self.parse_shortcut(default_shortcut)
                    vars['mod1'].set(mod1)
                    vars['mod2'].set(mod2)
                    vars['key'].set(key)
                    vars['display'].set(default_shortcut)
            
            # 更新配置文件
            config_data = self.config_manager.data
            config_data['shortcuts'] = self.shortcuts
            self.config_manager.save(config_data)
            
            # 通知主窗口更新键盘钩子
            if hasattr(self.parent, 'update_keyboard_hooks'):
                self.parent.update_keyboard_hooks()
            
            # 通知主窗口更新音轨分析标题
            if hasattr(self.parent, 'update_analysis_frame_title'):
                self.parent.update_analysis_frame_title()
            
            # messagebox.showinfo("成功", "已恢复默认快捷键设置")
        
        except Exception as e:
            print(f"恢复默认快捷键时出错: {str(e)}")
            messagebox.showerror("错误", f"恢复默认快捷键时出错: {str(e)}")
