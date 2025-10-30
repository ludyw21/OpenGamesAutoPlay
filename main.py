"""
MIDI自动演奏程序 - 一个基于ttkbootstrap的MIDI文件播放器，支持选择音轨和键盘控制。
提供直观的界面来加载、选择和播放MIDI文件，并支持全局快捷键控制。
"""

import os
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pages'))

# 设置DPI感知，确保在高DPI显示器上正常显示
if hasattr(os, 'name') and os.name == 'nt':
    try:
        # 导入ctypes以调用Windows API
        import ctypes
        # 设置进程DPI感知为系统DPI感知
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        # 获取当前DPI缩放比例
        dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
        print(f"检测到DPI缩放比例: {dpi_scale:.2f}")
    except Exception as e:
        print(f"设置DPI感知失败: {str(e)}")

import json
import keyboard
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import threading
import time
import pygame.mixer
from midi_player import MidiPlayer
from keyboard_mapping import CONTROL_KEYS
import mido
from pages.help_dialog import HelpDialog
from pages.settings_dialog import SettingsDialog
from midi_analyzer import MidiAnalyzer
from midi_preview_wrapper import get_preview_wrapper

# 忽略废弃警告
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

# 在主程序中添加或更新版本号
VERSION = "1.1.0"

class Config:
    def __init__(self, filename="config.json"):
        self.filename = filename
        self.data = self.load()
    
    def load(self):
        try:
            if os.path.exists(self.filename):
                with open(self.filename, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"加载配置文件失败: {str(e)}")
        return self.get_default_config()
    
    def save(self, data):
        try:
            with open(self.filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置文件失败: {str(e)}")
    
    @staticmethod
    def get_default_config():
        return {
            'last_directory': '',
            'stay_on_top': False,
            'theme': 'pink',  # 默认主题
            'shortcuts': {
                'START_PAUSE': 'alt+-',  # Alt + 减号键
                'STOP': 'alt+=',         # Alt + 等号键
                'PREV_SONG': 'alt+up',   # Alt + 上箭头键
                'NEXT_SONG': 'alt+down'  # Alt + 下箭头键
            }
        }

def handle_error(func_name):
    def decorator(func):
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                print(f"{func_name}时出错: {str(e)}")
                # 可以添加通用的错误恢复逻辑
        return wrapper
    return decorator

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title(f"开放世界-自动演奏by深瞳潜入梦-{VERSION}")
        
        # 创建配置管理器实例
        self.config_manager = Config()
        # 从配置管理器获取配置
        self.config = self.config_manager.data
        
        # 获取DPI缩放比例
        dpi_scale = 1.0
        if hasattr(os, 'name') and os.name == 'nt':
            try:
                import ctypes
                dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
            except:
                pass
        
        # 根据DPI缩放比例计算窗口大小，减少高度以消除底部多余空间
        base_width, base_height = 750, 520  # 增加200宽度，减少高度
        scaled_width = int(base_width * dpi_scale)
        scaled_height = int(base_height * dpi_scale)
        
        # 检查是否有保存的窗口大小和位置
        saved_width = self.config.get('window_width')
        saved_height = self.config.get('window_height')
        saved_x = self.config.get('window_x')
        saved_y = self.config.get('window_y')
        
        # 先获取屏幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 优先使用保存的位置（即使没有保存大小）
        if saved_x is not None and saved_y is not None:
            # 确定要使用的宽度和高度
            target_width = saved_width if saved_width else scaled_width
            target_height = saved_height if saved_height else scaled_height
            
            # 确保窗口位置不会超出屏幕（添加边界检查）
            x = max(0, min(saved_x, screen_width - target_width))
            y = max(0, min(saved_y, screen_height - target_height))
            
            # 一次性设置大小和位置
            self.root.geometry(f"{target_width}x{target_height}+{x}+{y}")
            print(f"应用保存的窗口位置: {target_width}x{target_height}+{x}+{y}")
        elif saved_width and saved_height:
            # 只有保存的大小，没有保存的位置，居中显示
            x = (screen_width - saved_width) // 2
            y = (screen_height - saved_height) // 2
            self.root.geometry(f"{saved_width}x{saved_height}+{x}+{y}")
            print(f"应用保存的窗口大小并居中: {saved_width}x{saved_height}+{x}+{y}")
        else:
            # 使用默认的缩放后尺寸并居中
            x = (screen_width - scaled_width) // 2
            y = (screen_height - scaled_height) // 2
            self.root.geometry(f"{scaled_width}x{scaled_height}+{x}+{y}")
            print(f"应用默认窗口大小并居中: {scaled_width}x{scaled_height}+{x}+{y}")
        
        self.root.minsize(scaled_width, scaled_height)  # 保持最小尺寸限制
        
        self.root.resizable(True, True)  # 允许调整窗口大小
        
        # 设置主题
        theme = self.config.get('theme', 'pink')
        self.style = ttkb.Style(theme=theme)
        self.last_directory = self.config.get('last_directory', '')
        
        # 初始化stay_on_top_var（提前初始化以避免属性错误）
        self.stay_on_top_var = tk.BooleanVar()
        
        # 初始化current_song_label为None，确保在setup_ui完成前不会报错
        self.current_song_label = None
        
        # 添加键盘事件防抖动
        self.last_key_time = 0
        self.key_cooldown = 0.2  # 200ms冷却时间
        
        # 初始化预览状态
        self.is_previewing = False
        # 试听MIDI相关状态
        self.is_playing_midi = False
        
        # 初始化pygame mixer（如果还没有初始化）
        if not pygame.mixer.get_init():
            try:
                pygame.mixer.init(frequency=44100)
            except Exception as e:
                print(f"初始化音频系统失败: {str(e)}")
        
        # 初始化其他属性
        self.current_index = -1
        self.midi_files = []
        self.tracks_info = []  # 初始化音轨信息列表
        self.selected_tracks = set()  # 存储选中的音轨索引
        self.current_events = []  # 存储事件表数据
        # 为每个音轨单独存储转音和转位设置
        self.track_transpose_vars = {}  # 存储每个音轨的移调设置
        self.track_octave_vars = {}    # 存储每个音轨的转位设置
        self.track_vars = {}           # 存储每个音轨的复选框变量
        self.track_ui_elements = {}    # 存储每个音轨的UI元素引用
        # 存储每个音轨的分析结果
        self.track_analysis_results = {}
        
        # 全局移调和转位变量
        self.transpose_var = tk.IntVar(value=0)
        self.octave_var = tk.IntVar(value=0)
        
        try:
            self.midi_player = MidiPlayer()
            # 为MidiPlayer添加窗口切换失败的回调
            self.midi_player.window_switch_failed_callback = self.handle_window_switch_failed
        except PermissionError as e:
            messagebox.showerror("权限错误", str(e))
            self.root.destroy()
            return
        
        # 用于更新进度条的计时器
        self.progress_timer = None
        
        # 预览按钮状态
        self.preview_button = None
        
        # 添加窗口状态检查定时器
        self.window_check_timer = None
        
        self.setup_ui()
        self.setup_keyboard_hooks()
        
        # 应用保存的置顶状态
        stay_on_top = self.config.get('stay_on_top', True)  # 默认为True
        if stay_on_top:
            self.root.attributes('-topmost', True)
        self.stay_on_top_var.set(stay_on_top)
        
        # 如果有上次的目录，自动加载
        if self.last_directory and os.path.exists(self.last_directory):
            self.load_directory(self.last_directory)
        
        # 启动定时器
        self.start_timers()
        
        # 设置关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.close_event)
    
    def center_window(self):
        """将窗口居中显示"""
        # 更新窗口信息
        self.root.update_idletasks()
        
        # 获取窗口宽度和高度
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        
        # 获取屏幕宽度和高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 计算居中位置
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        # 设置窗口位置
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        print(f"窗口居中: {width}x{height}+{x}+{y}")
        
    def start_timers(self):
        """启动所有定时器"""
        # 进度更新定时器
        self.progress_timer = self.root.after(100, self.update_progress)
        
        # 窗口状态检查定时器
        self.window_check_timer = self.root.after(200, self.check_window_state)
    
    def stop_timers(self):
        """停止所有定时器"""
        if self.progress_timer:
            self.root.after_cancel(self.progress_timer)
            self.progress_timer = None
        if self.window_check_timer:
            self.root.after_cancel(self.window_check_timer)
            self.window_check_timer = None
    
    def close_event(self):
        """关闭窗口时的处理"""
        try:
            # 停止预览
            if hasattr(self, 'is_previewing') and self.is_previewing:
                self.stop_preview()
            
            # 停止播放
            if hasattr(self, 'midi_player'):
                self.midi_player.stop()
            
            # 保存配置
            if hasattr(self, 'config'):
                self.config['last_directory'] = self.last_directory
                # 保存窗口大小和位置
                current_width = self.root.winfo_width()
                current_height = self.root.winfo_height()
                current_x = self.root.winfo_x()
                current_y = self.root.winfo_y()
                
                # 只有当窗口大小与初始大小有明显差异时才保存
                base_width, base_height = 750, 520
                dpi_scale = 1.0
                if hasattr(os, 'name') and os.name == 'nt':
                    try:
                        import ctypes
                        dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
                    except:
                        pass
                scaled_width = int(base_width * dpi_scale)
                scaled_height = int(base_height * dpi_scale)
                
                # 保存窗口大小和位置
                if abs(current_width - scaled_width) > 10 or abs(current_height - scaled_height) > 10:
                    self.config['window_width'] = current_width
                    self.config['window_height'] = current_height
                    print(f"保存窗口大小: {current_width}x{current_height}")
                
                # 始终保存窗口位置
                self.config['window_x'] = current_x
                self.config['window_y'] = current_y
                print(f"保存窗口位置: {current_x}, {current_y}")
                
                self.config_manager.save(self.config)
            
            # 停止所有定时器
            self.stop_timers()
            
            # 移除所有键盘钩子
            keyboard.unhook_all()
            
            self.root.destroy()
            
        except Exception as e:
            print(f"关闭窗口时出错: {str(e)}")
            self.root.destroy()
    
    def handle_window_switch_failed(self):
        """处理窗口切换失败"""
        messagebox.showwarning("窗口切换失败", "无法切换到目标窗口，请确保游戏窗口已打开！")
    
    def setup_ui(self):
        """设置UI界面"""
        try:
            # 创建主框架
            main_frame = ttk.Frame(self.root, padding=10)
            main_frame.pack(fill=BOTH, expand=YES)
            
            # 设置行权重使内容占满高度
            main_frame.grid_rowconfigure(0, weight=1)  # 行占满高度
            # 右侧列设置权重使其自适应
            main_frame.grid_columnconfigure(1, weight=1)  # 右侧自适应
            
            # 创建左侧框架 - 固定宽度180px
            left_frame = ttk.LabelFrame(main_frame, text="文件管理", padding=10)
            left_frame.grid(row=0, column=0, sticky='nsw', padx=5, pady=5)  # 只设置nsw，不包括e，防止拉伸
            left_frame.config(width=180)  # 使用config方法设置宽度
            left_frame.grid_propagate(False)  # 防止内部组件改变框架大小
            
            # 置顶复选框
            top_frame = ttk.Frame(left_frame)
            top_frame.pack(fill=X, pady=5)
            
            self.stay_on_top_var = tk.BooleanVar()
            stay_on_top_checkbox = ttk.Checkbutton(
                top_frame, 
                text="窗口置顶", 
                variable=self.stay_on_top_var,
                command=self.toggle_stay_on_top
            )
            stay_on_top_checkbox.pack(side=LEFT)
            
            # 文件选择按钮
            self.file_button = ttk.Button(
                left_frame, 
                text="选择MIDI文件夹", 
                command=self.select_directory,
                style="Primary.TButton"
            )
            self.file_button.pack(fill=X, pady=5)
            
            # 搜索框
            search_frame = ttk.Frame(left_frame)
            search_frame.pack(fill=X, pady=5)
            
            self.search_input = ttk.Entry(search_frame)
            self.search_input.pack(fill=X)
            self.search_input.insert(0, "搜索歌曲...")
            self.search_input.bind("<FocusIn>", lambda e: self.search_input.delete(0, END))
            self.search_input.bind("<KeyRelease>", lambda e: self.filter_songs())
            
            # 修改滚动条样式使其更宽
            style = ttk.Style()
            # 对于Windows主题，需要设置箭头按钮的宽度和滑块的宽度
            style.configure("Wide.TScrollbar", arrowsize=20, width=20)
            # 为滚动条轨道和滑块设置明确的宽度
            style.layout("Wide.TScrollbar", [
                ('Vertical.Scrollbar.trough', {'sticky': 'nswe', 'children': [
                    ('Vertical.Scrollbar.thumb', {'expand': '1', 'sticky': 'nswe'})
                ]})
            ])
            
            # 创建歌曲列表的垂直滚动条 - 将使用绝对定位叠加在列表上
            self.song_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, style="Wide.TScrollbar")
            
            # 创建歌曲列表 - 适应左侧框架宽度
            self.song_list = ttk.Treeview(left_frame, columns=["song"], show="headings", yscrollcommand=self.song_scrollbar.set)
            
            # 配置滚动条与Treeview的关联
            self.song_scrollbar.config(command=self.song_list.yview)
            
            # 设置列表属性
            self.song_list.heading("song", text="歌曲列表", anchor='center')
            self.song_list.column("song", width=160)  # 固定宽度以适应左侧面板
            self.song_list.bind("<<TreeviewSelect>>", lambda e: self.song_selected())
            
            # 绑定鼠标滚轮事件 - 只在歌曲列表上滚动时响应
            self.song_list.bind("<MouseWheel>", self._on_song_list_mousewheel)
            
            # 先放置歌曲列表，占据完整空间
            self.song_list.pack(fill=BOTH, expand=YES, pady=5)
            
            # 添加歌曲列表更新后的回调，用于控制滚动条显示
            def on_song_list_updated(event=None):
                # 强制更新布局
                left_frame.update_idletasks()
                
                # 获取Treeview的实际可见高度和位置
                visible_height = self.song_list.winfo_height()
                x = self.song_list.winfo_x()
                y = self.song_list.winfo_y()
                width = self.song_list.winfo_width()
                
                # 获取项目数量并计算总高度
                item_count = len(self.song_list.get_children())
                item_height = 20  # 估算的每行高度
                total_height = item_count * item_height
                
                # 按需显示滚动条
                if item_count > 0 and total_height > visible_height:
                    # 使用place方法将滚动条绝对定位在歌曲列表的右侧
                    # 当使用in_参数时，坐标是相对于父窗口的
                    scroll_width = 20  # 显式设置滚动条宽度
                    # 不使用in_参数，直接使用相对于left_frame的坐标
                    self.song_scrollbar.place(x=width-scroll_width, y=y, width=scroll_width, height=visible_height)
                    self.song_scrollbar.lift()  # 确保滚动条在顶层
                else:
                    # 不需要滚动条时隐藏
                    self.song_scrollbar.place_forget()
            
            # 保存回调函数引用，以便后续调用
            self._on_song_list_updated = on_song_list_updated
            
            # 初始调用一次
            on_song_list_updated()
            
            # 右侧框架 - 自适应宽度
            right_frame = ttk.LabelFrame(main_frame, text="播放控制", padding=10)
            right_frame.grid(row=0, column=1, sticky=NSEW, padx=5, pady=5)  # 全方向拉伸，使其自适应
            
            # 获取配置的最低音和最高音
            config_min_note = 48  # 默认值
            config_max_note = 83  # 默认值
            try:
                from midi_analyzer import MidiAnalyzer
                config_min_note, config_max_note, _ = MidiAnalyzer._get_key_settings()
            except:
                pass
            
            # 导入get_note_name函数
            try:
                from groups import get_note_name
            except:
                # 如果导入失败，创建一个简单的替代函数
                def get_note_name(note):
                    return str(note)
            
            # 创建音轨详情LabelFrame - 包含最低音和最高音信息以及详细标题
            self.tracks_frame = ttk.LabelFrame(right_frame, 
                                         text=f"音轨详情【 当前播放范围：{get_note_name(config_min_note)}({config_min_note}) - {get_note_name(config_max_note)}({config_max_note}) 】",
                                         padding=10)
            self.tracks_frame.pack(fill=X, pady=5)
            
            # 当前歌曲名称标签 - 设置为不换行，超出部分不显示，占满整行
            song_label_frame = ttk.Frame(self.tracks_frame)
            song_label_frame.pack(fill=X, pady=2)
            
            # 设置文本左对齐，wraplength=0确保不换行，超出部分会自动截断
            self.current_song_label = ttk.Label(song_label_frame, text="当前歌曲：未选择", anchor=W, wraplength=0, justify=LEFT)
            self.current_song_label.pack(fill=X, expand=True)
            
            # 创建音轨列表表格框架 - 支持垂直滚动
            track_table_frame = ttk.Frame(self.tracks_frame)
            track_table_frame.pack(fill=BOTH, expand=True, pady=5)
            
            # 创建Canvas作为滚动区域
            self.track_canvas = tk.Canvas(track_table_frame)
            
            # 添加垂直滚动条
            self.track_scrollbar = ttk.Scrollbar(track_table_frame, orient=tk.VERTICAL, command=self.track_canvas.yview)
            
            # 配置Canvas的滚动
            self.track_canvas.configure(yscrollcommand=self.track_scrollbar.set)
            
            # 创建内部框架来容纳所有音轨行
            self.track_rows_frame = ttk.Frame(self.track_canvas)
            self.track_canvas_window = self.track_canvas.create_window((0, 0), window=self.track_rows_frame, anchor="nw")
            
            # 绑定鼠标滚轮事件以支持滚动
            self.track_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
            
            # 当调整窗口大小时，调整Canvas中内容的宽度和处理滚动条按需显示
            def on_canvas_configure(event):
                # 当Canvas宽度改变时，更新内部框架的宽度
                width = event.width
                self.track_canvas.itemconfig(self.track_canvas_window, width=width)
                update_scrollbars()
            
            # 当内部框架大小改变时，更新滚动区域
            def on_track_rows_configure(event):
                update_scrollbars()
            
            # 处理滚动条按需显示的函数
            def update_scrollbars(event=None):
                # 确保在更新前强制更新所有窗口大小
                track_table_frame.update_idletasks()
                self.track_rows_frame.update_idletasks()
                
                # 更新Canvas的滚动区域
                bbox = self.track_canvas.bbox("all")
                if bbox:
                    # 确保内容从顶部开始，y0设为0
                    self.track_canvas.configure(scrollregion=(0, 0, bbox[2], bbox[3]))
                
                # 获取Canvas的实际大小和内容大小
                canvas_height = self.track_canvas.winfo_height()
                content_height = self.track_rows_frame.winfo_height()
                
                # 垂直滚动条按需显示
                if content_height > canvas_height + 10:  # 增加一点余量避免闪烁
                    self.track_scrollbar.pack(side="right", fill="y")
                    self.track_canvas.pack(side="left", fill=BOTH, expand=True)
                else:
                    self.track_scrollbar.pack_forget()
                    self.track_canvas.pack(side="left", fill=BOTH, expand=True)
                    # 确保内容始终在顶部
                    self.track_canvas.yview_moveto(0)
            
            # 绑定事件
            self.track_canvas.bind('<Configure>', on_canvas_configure)
            self.track_rows_frame.bind('<Configure>', on_track_rows_configure)
            
            # 初始显示Canvas
            self.track_canvas.pack(side="left", fill=BOTH, expand=True)
            
            # 初始更新滚动条状态
            update_scrollbars()
            
            # 操作区域LabelFrame
            operation_frame = ttk.LabelFrame(right_frame, text="操作", padding=10)
            operation_frame.pack(fill=X, pady=5)
            
            # 时间显示
            self.time_label = ttk.Label(operation_frame, text="剩余时间: 00:00", anchor=CENTER)
            self.time_label.pack(fill=X, pady=5)
            
            # 控制按钮布局
            control_frame = ttk.Frame(operation_frame)
            control_frame.pack(fill=X, pady=10)
            
            # 播放/暂停按钮
            self.play_button = ttk.Button(
                control_frame, 
                text="播放", 
                command=self.toggle_play,
                style="Success.TButton",
                state=DISABLED
            )
            self.play_button.pack(side=LEFT, padx=5, expand=YES, fill=X)
            
            # 停止按钮
            self.stop_button = ttk.Button(
                control_frame, 
                text="停止", 
                command=self.stop_playback,
                style="Danger.TButton",
                state=DISABLED
            )
            self.stop_button.pack(side=LEFT, padx=5, expand=YES, fill=X)
            
            # 预览按钮
            self.preview_button = ttk.Button(
                control_frame, 
                text="预览", 
                command=self.toggle_preview,
                style="Info.TButton",
                state=DISABLED
            )
            self.preview_button.pack(side=LEFT, padx=5, expand=YES, fill=X)
            
            # 试听MIDI按钮 - 直接播放原始MIDI文件
            self.midi_play_button = ttk.Button(
                control_frame, 
                text="试听MIDI", 
                command=self.toggle_midi_playback,
                style="Info.TButton",
                state=DISABLED
            )
            self.midi_play_button.pack(side=LEFT, padx=5, expand=YES, fill=X)
        
            # 其他LabelFrame
            other_frame = ttk.LabelFrame(right_frame, text="其他", padding=10)
            other_frame.pack(fill=X, pady=5)
            
            # 其他功能按钮布局
            other_buttons_frame = ttk.Frame(other_frame)
            other_buttons_frame.pack(fill=X, pady=5)
            
            # 事件表按钮
            self.event_button = ttk.Button(
                other_buttons_frame, 
                text="事件表", 
                command=self.show_event_table
            )
            self.event_button.pack(side=LEFT, padx=5, fill=X)
            
            # 设置按钮
            self.settings_button = ttk.Button(
                other_buttons_frame, 
                text="设置", 
                command=self.show_settings
            )
            self.settings_button.pack(side=LEFT, padx=5, fill=X)
            
            # 帮助按钮
            self.help_button = ttk.Button(
                other_buttons_frame, 
                text="帮助", 
                command=self.show_help
            )
            self.help_button.pack(side=LEFT, padx=5, fill=X)
            
            # 底部预留空间
            bottom_space = ttk.Frame(right_frame)
            bottom_space.pack(fill=X, pady=5)
            
        except Exception as e:
            print(f"设置UI界面时出错: {str(e)}")
            messagebox.showerror("UI错误", f"设置界面时出错: {str(e)}")
    
    def toggle_stay_on_top(self):
        """切换窗口置顶状态"""
        stay_on_top = self.stay_on_top_var.get()
        self.root.attributes('-topmost', stay_on_top)
        
    def update_event_data(self):
        """根据选中音轨预生成事件数据（唯一负责生成数据的方法）"""
        # 如果有当前文件和选中的音轨，生成事件数据
        if hasattr(self, 'current_file_path') and self.current_file_path and self.selected_tracks:
            try:
                print("[DEBUG] 开始更新事件数据")
                # 获取当前的全局移调和转位设置
                global_transpose = self.transpose_var.get()
                global_octave_shift = self.octave_var.get()
                print(f"[DEBUG] 全局移调: {global_transpose}, 全局转位: {global_octave_shift}")
                
                # 使用MidiAnalyzer来生成事件数据，传递全局移调和转位参数
                events, analysis_result, track_names, track_note_counts = MidiAnalyzer.analyze_midi_file(
                    self.current_file_path, 
                    self.selected_tracks,
                    transpose=global_transpose,
                    octave_shift=global_octave_shift
                )
                # 保存音轨名称映射和音符计数
                self.track_names = track_names
                self.track_note_counts = track_note_counts
                print(f"[DEBUG] 音轨音符计数: {track_note_counts}")
                
                # 应用单个音轨的移调和转位设置
                if hasattr(self, 'track_transpose_vars') and hasattr(self, 'track_octave_vars'):
                    print("[DEBUG] 应用单个音轨的转音设置")
                    for track_index in self.track_transpose_vars:
                        if track_index in self.selected_tracks:
                            track_transpose = self.track_transpose_vars[track_index].get()
                            track_octave = self.track_octave_vars[track_index].get()
                            print(f"[DEBUG] 音轨{track_index} 移调: {track_transpose}, 转位: {track_octave}")
                    
                    modified_count = 0
                    for event in events:
                        track_index = event.get('track')
                        if track_index is not None and track_index in self.track_transpose_vars:
                            # 获取该音轨的移调和转位设置
                            track_transpose = self.track_transpose_vars[track_index].get()
                            track_octave = self.track_octave_vars[track_index].get()
                            
                            # 应用音轨特定的移调和转位
                            total_offset = track_transpose + track_octave * 12
                            if total_offset != 0:
                                old_note = event['note']
                                event['note'] += total_offset
                                modified_count += 1
                                
                                # 更新分组信息
                                from groups import group_for_note
                                event['group'] = group_for_note(event['note'])
                                
                                # 调试信息：只打印少量修改的事件
                                if modified_count <= 5:
                                    print(f"[DEBUG] 修改音轨{track_index} 音符: {old_note} -> {event['note']}")
                    
                    print(f"[DEBUG] 共修改 {modified_count} 个事件的音符值")
                
                # 更新当前事件数据和分析结果
                self.current_events = events
                self.current_analysis_result = analysis_result
                
                print(f"[DEBUG] 已生成事件数据：{len(self.current_events)}个事件，{len(self.current_events)/2}个音符。")
                
                # 如果有打开的事件表对话框，通知它刷新显示
                if hasattr(self, 'current_event_table_dialog') and self.current_event_table_dialog:
                    try:
                        print("[DEBUG] 通知事件表对话框刷新")
                        self.current_event_table_dialog.populate_event_table()
                        print("[DEBUG] 事件表对话框刷新完成")
                    except Exception as e:
                        print(f"[DEBUG] 通知事件表对话框刷新时出错: {str(e)}")
                else:
                    print("[DEBUG] 当前没有打开的事件表对话框")
                
            except Exception as e:
                print(f"生成事件数据时出错: {str(e)}")
                # 如果出错，使用空列表
                self.current_events = []
                self.current_analysis_result = None
    
    def show_event_table(self):
        """显示事件表，直接使用预生成的事件数据"""
        from pages.event_table_dialog import EventTableDialog
        
        # 确保有事件数据（如果没有则临时生成）
        if not hasattr(self, 'current_events') or not self.current_events:
            self.update_event_data()
        
        # 如果已经有打开的事件表对话框，先关闭它
        if hasattr(self, 'current_event_table_dialog') and self.current_event_table_dialog:
            try:
                self.current_event_table_dialog.dialog.destroy()
            except:
                pass
        
        # 创建新的事件表对话框并保存引用
        self.current_event_table_dialog = EventTableDialog(self)
        
        # 绑定窗口关闭事件，以便在对话框关闭时清除引用
        self.current_event_table_dialog.dialog.protocol("WM_DELETE_WINDOW", self._on_event_table_close)

    def _on_event_table_close(self):
        """处理事件表对话框关闭事件"""
        if hasattr(self, 'current_event_table_dialog') and self.current_event_table_dialog:
            try:
                # 先销毁对话框
                self.current_event_table_dialog.dialog.destroy()
            except:
                pass
            # 然后清除引用
            self.current_event_table_dialog = None

    def show_settings(self):
        """显示设置对话框"""
        SettingsDialog(self, self.config_manager)
    
    def show_help(self):
        """显示帮助对话框"""
        HelpDialog(self.root)
    
    def update_theme(self, theme_name):
        """立即更新应用程序主题"""
        try:
            # 更新样式
            self.style.theme_use(theme_name)
            
            # 更新配置
            self.config['theme'] = theme_name
            self.config_manager.save(self.config)
            
            print(f"[DEBUG] 主题已更新为: {theme_name}")
        except Exception as e:
            print(f"[DEBUG] 更新主题时出错: {str(e)}")
    
    def setup_keyboard_hooks(self):
        """设置键盘快捷键"""
        try:
            # 从keyboard_mapping.py导入默认的控制键配置
            from keyboard_mapping import CONTROL_KEYS
            
            # 默认快捷键设置，使用keyboard_mapping中的配置
            default_shortcuts = CONTROL_KEYS.copy()
            
            # 从配置中获取快捷键设置，如果没有则使用默认值
            shortcuts = self.config.get('shortcuts', default_shortcuts.copy())
            
            # 标记是否需要更新配置
            need_update_config = False
            
            # 预先验证所有快捷键的有效性
            for action, shortcut in shortcuts.items():
                if action in default_shortcuts:
                    try:
                        # 临时添加并立即移除来验证快捷键格式
                        keyboard.add_hotkey(shortcut, lambda: None)
                        keyboard.remove_hotkey(shortcut)
                    except Exception as e:
                        print(f"验证快捷键 {action}: {shortcut} 时出错: {str(e)}")
                        need_update_config = True
                        break
            
            # 如果发现任何无效的快捷键，将所有快捷键更新为默认值
            if need_update_config:
                print("发现无效的快捷键配置，将所有快捷键更新为默认值")
                shortcuts = default_shortcuts.copy()
            
            # 逐个添加快捷键，允许单个快捷键失败而不影响其他快捷键
            shortcuts_added = 0
            
            # 播放/暂停
            try:
                keyboard.add_hotkey(shortcuts.get('START_PAUSE', default_shortcuts['START_PAUSE']), 
                                  lambda: self.safe_key_handler(self.toggle_play),
                                  suppress=True, trigger_on_release=True)
                shortcuts_added += 1
            except Exception as e:
                print(f"设置播放/暂停快捷键时出错: {str(e)}")
                print(f"将使用keyboard_mapping中的默认播放/暂停快捷键: {default_shortcuts['START_PAUSE']}")
                # 更新配置中的错误值
                shortcuts['START_PAUSE'] = default_shortcuts['START_PAUSE']
                need_update_config = True
                try:
                    keyboard.add_hotkey(default_shortcuts['START_PAUSE'], 
                                      lambda: self.safe_key_handler(self.toggle_play),
                                      suppress=True, trigger_on_release=True)
                    shortcuts_added += 1
                except:
                    pass
            
            # 停止
            try:
                keyboard.add_hotkey(shortcuts.get('STOP', default_shortcuts['STOP']), 
                                  lambda: self.safe_key_handler(self.stop_playback),
                                  suppress=True, trigger_on_release=True)
                shortcuts_added += 1
            except Exception as e:
                print(f"设置停止快捷键时出错: {str(e)}")
                print(f"将使用keyboard_mapping中的默认停止快捷键: {default_shortcuts['STOP']}")
                # 更新配置中的错误值
                shortcuts['STOP'] = default_shortcuts['STOP']
                need_update_config = True
                try:
                    keyboard.add_hotkey(default_shortcuts['STOP'], 
                                      lambda: self.safe_key_handler(self.stop_playback),
                                      suppress=True, trigger_on_release=True)
                    shortcuts_added += 1
                except:
                    pass
            
            # 上一首
            try:
                keyboard.add_hotkey(shortcuts.get('PREV_SONG', default_shortcuts['PREV_SONG']), 
                                  lambda: self.safe_key_handler(self.play_previous_song),
                                  suppress=True, trigger_on_release=True)
                shortcuts_added += 1
            except Exception as e:
                print(f"设置上一首快捷键时出错: {str(e)}")
                print(f"将使用keyboard_mapping中的默认上一首快捷键: {default_shortcuts['PREV_SONG']}")
                # 更新配置中的错误值
                shortcuts['PREV_SONG'] = default_shortcuts['PREV_SONG']
                need_update_config = True
                try:
                    keyboard.add_hotkey(default_shortcuts['PREV_SONG'], 
                                      lambda: self.safe_key_handler(self.play_previous_song),
                                      suppress=True, trigger_on_release=True)
                    shortcuts_added += 1
                except:
                    pass
            
            # 下一首
            try:
                keyboard.add_hotkey(shortcuts.get('NEXT_SONG', default_shortcuts['NEXT_SONG']), 
                                  lambda: self.safe_key_handler(self.play_next_song),
                                  suppress=True, trigger_on_release=True)
                shortcuts_added += 1
            except Exception as e:
                print(f"设置下一首快捷键时出错: {str(e)}")
                print(f"将使用keyboard_mapping中的默认下一首快捷键: {default_shortcuts['NEXT_SONG']}")
                # 更新配置中的错误值
                shortcuts['NEXT_SONG'] = default_shortcuts['NEXT_SONG']
                need_update_config = True
                try:
                    keyboard.add_hotkey(default_shortcuts['NEXT_SONG'], 
                                      lambda: self.safe_key_handler(self.play_next_song),
                                      suppress=True, trigger_on_release=True)
                    shortcuts_added += 1
                except:
                    pass
            
            # 如果检测到错误并更新了快捷键配置，则保存到config.json
            if need_update_config:
                print("检测到快捷键配置错误，已使用默认值更新config.json")
                self.config['shortcuts'] = shortcuts
                try:
                    self.config_manager.save(self.config)
                except Exception as e:
                    print(f"保存配置时出错: {str(e)}")
            
            if shortcuts_added > 0:
                print(f"键盘快捷键设置完成，成功添加了 {shortcuts_added} 个快捷键")
            else:
                print("警告：无法设置任何键盘快捷键，请检查系统权限和键盘库安装")
                
        except Exception as e:
            print(f"设置键盘快捷键时出错: {str(e)}")
            messagebox.showerror("快捷键错误", f"设置快捷键时出错: {str(e)}")
            # 尝试使用最基本的默认快捷键
            try:
                keyboard.add_hotkey('alt+1', lambda: self.safe_key_handler(self.toggle_play))
                keyboard.add_hotkey('alt+2', lambda: self.safe_key_handler(self.stop_playback))
                print("已设置基本默认快捷键")
            except:
                pass
    
    def safe_key_handler(self, func):
        """安全地处理键盘事件，添加防抖动和状态检查"""
        try:
            current_time = time.time()
            if current_time - self.last_key_time < self.key_cooldown:
                return
            
            self.last_key_time = current_time
            
            # 确保窗口可见
            if self.root.state() != 'iconic':
                func()
                
        except Exception as e:
            print(f"处理键盘事件时出错: {str(e)}")
    
    def select_directory(self):
        """选择MIDI文件夹"""
        directory = filedialog.askdirectory(initialdir=self.last_directory)
        if directory:
            self.last_directory = directory
            self.load_directory(directory)
    
    def load_directory(self, dir_path):
        """加载目录中的MIDI文件"""
        try:
            self.midi_files = self._load_midi_files(dir_path)
            self.update_song_list()
            
            # 保存配置
            self.config['last_directory'] = dir_path
            self.config_manager.save(self.config)
            
        except Exception as e:
            print(f"加载目录时出错: {str(e)}")
            messagebox.showerror("加载错误", f"加载目录时出错: {str(e)}")
    
    def _load_midi_files(self, dir_path):
        """加载指定目录下的所有MIDI文件"""
        midi_files = []
        for root, _, files in os.walk(dir_path):
            for file in files:
                if file.lower().endswith(('.mid', '.midi')):
                    midi_files.append(os.path.join(root, file))
        return midi_files
    
    def update_song_list(self):
        """更新歌曲列表"""
        # 清空当前列表
        for item in self.song_list.get_children():
            self.song_list.delete(item)
        
        # 添加文件
        for file_path in self.midi_files:
            file_name = os.path.basename(file_path)
            self.song_list.insert('', END, values=[file_name], tags=(file_path,))
        
        # 更新滚动条显示状态
        if hasattr(self, '_on_song_list_updated'):
            self._on_song_list_updated()
    
    def disable_keyboard_hooks(self):
        """暂时禁用所有键盘钩子，用于快捷键设置时避免冲突"""
        try:
            # 移除所有现有的键盘钩子
            keyboard.unhook_all()
            print("键盘钩子已暂时禁用")
        except Exception as e:
            print(f"禁用键盘钩子时出错: {str(e)}")
            
    def update_keyboard_hooks(self):
        """更新键盘钩子"""
        try:
            # 移除所有现有的键盘钩子
            keyboard.unhook_all()
            # 重新设置键盘钩子
            self.setup_keyboard_hooks()
            print("键盘快捷键已更新")
        except Exception as e:
            print(f"更新键盘快捷键时出错: {str(e)}")
            messagebox.showerror("快捷键错误", f"更新快捷键时出错: {str(e)}")
    
    def update_analysis_frame_title(self):
        """更新音轨分析LabelFrame的标题"""
        try:
            # 获取配置的最低音和最高音
            config_min_note = 48  # 默认值
            config_max_note = 83  # 默认值
            try:
                from midi_analyzer import MidiAnalyzer
                config_min_note, config_max_note, _ = MidiAnalyzer._get_key_settings()
                print(f"[DEBUG] 从配置获取的音域范围: {config_min_note} - {config_max_note}")
            except Exception as e:
                print(f"[DEBUG] 获取配置音域时出错: {str(e)}, 使用默认值: {config_min_note} - {config_max_note}")
            
            # 导入get_note_name函数
            try:
                from groups import get_note_name
                print(f"[DEBUG] 成功导入get_note_name函数")
            except:
                # 如果导入失败，创建一个简单的替代函数
                def get_note_name(note):
                    return str(note)
                print(f"[DEBUG] 使用替代的get_note_name函数")
            
            # 更新tracks_frame的标题
            if hasattr(self, 'tracks_frame'):
                new_title = f"音轨详情【 当前播放范围：{get_note_name(config_min_note)}({config_min_note}) - {get_note_name(config_max_note)}({config_max_note}) 】"
                print(f"[DEBUG] 准备更新标题为: {new_title}")
                self.tracks_frame.config(text=new_title)
                print(f"[DEBUG] 音轨详情标题已更新: {new_title}")
            else:
                print("[DEBUG] tracks_frame属性不存在，无法更新标题")
        except Exception as e:
            print(f"[DEBUG] 更新音轨分析标题时出错: {str(e)}")
            
    def filter_songs(self):
        """根据搜索文本过滤歌曲列表"""
        search_text = self.search_input.get().lower()
        
        # 清空当前列表
        for item in self.song_list.get_children():
            self.song_list.delete(item)
        
        # 添加匹配的文件
        for file_path in self.midi_files:
            file_name = os.path.basename(file_path)
            if search_text in file_name.lower():
                self.song_list.insert('', END, values=[file_name], tags=(file_path,))
        
        # 更新滚动条显示状态
        if hasattr(self, '_on_song_list_updated'):
            self._on_song_list_updated()
    
    def song_selected(self):
        """选择歌曲时的处理"""
        selected_items = self.song_list.selection()
        if not selected_items:
            return
        
        # 获取选中的文件路径
        item = selected_items[0]
        file_path = self.song_list.item(item, "tags")[0]
        self.current_file = file_path
        
        # 更新当前歌曲标签（确保标签已初始化）
        if hasattr(self, 'current_song_label') and self.current_song_label is not None:
            file_name = os.path.basename(file_path)
            self.current_song_label.config(text=f"当前歌曲：{file_name}")
        
        # 重置移调和转位值为0
        if hasattr(self, 'transpose_var'):
            self.transpose_var.set(0)
        if hasattr(self, 'octave_var'):
            self.octave_var.set(0)
        
        # 使用单次扫描方法加载和分析MIDI文件
        self._load_and_analyze_midi(file_path)
        
        # 启用预览按钮（确保按钮已初始化）
        if hasattr(self, 'preview_button') and self.preview_button is not None:
            self.preview_button.config(state=NORMAL)
        
        # 如果当前正在播放MIDI，自动切换到新选择的MIDI文件
        if hasattr(self, 'is_playing_midi') and self.is_playing_midi:
            # 停止当前播放的MIDI
            self.stop_midi_playback()
            # 自动开始播放新选择的MIDI
            self.start_midi_playback()
    
    def _load_and_analyze_midi(self, file_path):
        """单次扫描MIDI文件，同时获取音轨信息和生成事件数据"""
        try:
            # 清空当前音轨列表
            for widget in self.track_rows_frame.winfo_children():
                widget.destroy()
            
            # 清空选中的音轨集合
            self.selected_tracks.clear()
            self.track_transpose_vars = {}
            self.track_octave_vars = {}
            self.track_analysis_results = {}
            
            # 显示加载状态
            loading_label = ttk.Label(self.track_rows_frame, text="正在加载和分析MIDI文件...", padding=10)
            loading_label.pack(fill=X)
            self.root.update()
            
            # 导入必要的库
            import mido
            import threading
            import time
            from midi_analyzer import MidiAnalyzer
            
            def parse_and_analyze_in_thread():
                """在线程中解析和分析MIDI文件，避免UI冻结"""
                start_time = time.time()
                
                try:
                    # 初始化数据结构
                    all_tracks = set()  # 稍后会从track_names中获取
                    selected_tracks = set()
                    
                    # 一次性扫描MIDI文件，同时获取事件数据、分析结果、音轨名称和音轨音符数量统计
                    # 这是唯一一次扫描整个MIDI文件的地方
                    events, analysis_result, track_names, track_note_counts = MidiAnalyzer.analyze_midi_file(
                        file_path,
                        all_tracks,  # 空集合表示分析所有音轨
                        transpose=0,  # 初始移调为0
                        octave_shift=0  # 初始转位为0
                    )
                    
                    # 更新all_tracks集合
                    all_tracks = set(track_names.keys())
                    
                    # 构建最终的音轨信息，过滤掉没有音符的音轨
                    tracks_info = []
                    for track_idx in track_names:
                        # 只有有音符的音轨才显示
                        if track_idx in track_note_counts and track_note_counts[track_idx] > 0:
                            note_count = track_note_counts[track_idx]
                            track_name = track_names[track_idx]
                            
                            # 使用专门的乱码修复方法处理音轨名称
                            fixed_name = self._fix_mojibake(track_name)
                            
                            # 构建显示名称，添加音轨标号
                            display_name = f"音轨{track_idx+1}：{fixed_name} ({note_count}个音符)"
                            
                            tracks_info.append({
                                "track_index": track_idx, 
                                "note_count": note_count, 
                                "display_name": display_name,
                                "original_name": track_name,
                                "fixed_name": fixed_name
                            })
                            selected_tracks.add(track_idx)
                    
                    # 在主线程中更新UI
                    def update_ui():
                        # 移除加载状态
                        loading_label.destroy()
                        
                        # 添加全选/取消全选控制行
                        all_tracks_frame = ttk.Frame(self.track_rows_frame)
                        all_tracks_frame.pack(fill=X, pady=2)
                        
                        # 全选复选框
                        self.all_tracks_var = tk.BooleanVar(value=True)
                        all_checkbox = ttk.Checkbutton(all_tracks_frame, variable=self.all_tracks_var, 
                                                     command=self.toggle_select_all)
                        all_checkbox.pack(side=LEFT, padx=5)
                        
                        # 全选标签
                        ttk.Label(all_tracks_frame, text="全部音轨", font=('微软雅黑', 9, 'bold')).pack(side=LEFT, padx=5)
                        
                        # 添加音轨信息
                        self.tracks_info = []
                        
                        # 为每个音轨创建独立的行
                        for info in tracks_info:
                            track_index = info["track_index"]
                            
                            # 创建音轨行框架
                            track_frame = ttk.Frame(self.track_rows_frame)
                            track_frame.pack(fill=X, pady=2, ipady=2)
                            
                            # 设置grid布局，让中间列占据剩余空间
                            track_frame.grid_columnconfigure(1, weight=1)  # 中间列（分析区域）占据剩余空间
                            
                            # 第一列：复选框（固定宽度）
                            track_var = tk.BooleanVar(value=True)
                            track_checkbox = ttk.Checkbutton(track_frame, variable=track_var,
                                                           command=lambda idx=track_index: self.toggle_track_selection(idx))
                            track_checkbox.grid(row=0, column=0, sticky='nsw', padx=5)  # 左侧固定位置
                            
                            # 第二列：音轨及分析信息（自适应宽度）
                            analysis_frame = ttk.LabelFrame(track_frame, text="音轨及分析")
                            analysis_frame.grid(row=0, column=1, sticky='nsew', padx=5)  # 填充整行高度和宽度
                            analysis_frame.grid_columnconfigure(0, weight=1)  # 内部列也设置权重
                            
                            # 音轨名称标签 - 添加更详细的信息
                            track_name_label = ttk.Label(analysis_frame, text=info["display_name"], font=('微软雅黑', 9, 'bold'))
                            track_name_label.pack(fill=X, padx=5, pady=2)  # 填充整个宽度
                            
                            # 分析结果文本框 - 初始就创建Text组件而非Label
                            analysis_label = tk.Text(analysis_frame, height=3, font=('微软雅黑', 9), wrap=tk.WORD)
                            analysis_label.insert(tk.END, "正在分析...")
                            analysis_label.pack(fill=X, padx=5, pady=1)  # 填充整个宽度
                            analysis_label.config(state=tk.DISABLED)  # 设置为只读
                            
                            # 第三列：转音设置（固定宽度）
                            transpose_frame = ttk.LabelFrame(track_frame, text="转音设置")
                            transpose_frame.grid(row=0, column=2, sticky='nse', padx=5)  # 右侧固定位置
                            
                            # 创建移调和转位控制
                            track_transpose_var = tk.IntVar(value=0)
                            track_octave_var = tk.IntVar(value=0)
                            
                            # 移调控制
                            transpose_control_frame = ttk.Frame(transpose_frame)
                            transpose_control_frame.pack(fill=X, pady=2)
                            
                            ttk.Label(transpose_control_frame, text="移调(半音):", font=('微软雅黑', 9)).pack(side=LEFT, padx=2)
                            ttk.Button(transpose_control_frame, text="-", width=2,
                                      command=lambda var=track_transpose_var, idx=track_index:
                                          self.adjust_track_transpose(idx, var, -1)).pack(side=LEFT)
                            transpose_entry = ttk.Entry(transpose_control_frame, textvariable=track_transpose_var, 
                                                      width=3, justify='center', font=('微软雅黑', 9))
                            transpose_entry.pack(side=LEFT, padx=1)
                            # 添加事件监听器
                            transpose_entry.bind('<Return>', lambda event, idx=track_index, var=track_transpose_var:
                                                self.on_track_transpose_change(idx, var))
                            transpose_entry.bind('<FocusOut>', lambda event, idx=track_index, var=track_transpose_var:
                                                self.on_track_transpose_change(idx, var))
                            ttk.Button(transpose_control_frame, text="+", width=2,
                                      command=lambda var=track_transpose_var, idx=track_index:
                                          self.adjust_track_transpose(idx, var, 1)).pack(side=LEFT)
                            
                            # 转位控制
                            octave_control_frame = ttk.Frame(transpose_frame)
                            octave_control_frame.pack(fill=X, pady=2)
                            
                            ttk.Label(octave_control_frame, text="转位(八度):", font=('微软雅黑', 9)).pack(side=LEFT, padx=2)
                            ttk.Button(octave_control_frame, text="-", width=2,
                                      command=lambda var=track_octave_var, idx=track_index:
                                          self.adjust_track_octave(idx, var, -1)).pack(side=LEFT)
                            octave_entry = ttk.Entry(octave_control_frame, textvariable=track_octave_var, 
                                                    width=3, justify='center', font=('微软雅黑', 9))
                            octave_entry.pack(side=LEFT, padx=1)
                            # 添加事件监听器
                            octave_entry.bind('<Return>', lambda event, idx=track_index, var=track_octave_var:
                                            self.on_track_octave_change(idx, var))
                            octave_entry.bind('<FocusOut>', lambda event, idx=track_index, var=track_octave_var:
                                            self.on_track_octave_change(idx, var))
                            ttk.Button(octave_control_frame, text="+", width=2,
                                      command=lambda var=track_octave_var, idx=track_index:
                                          self.adjust_track_octave(idx, var, 1)).pack(side=LEFT)
                            
                            # 重置转音超链接文本
                            reset_frame = ttk.Frame(transpose_frame)
                            reset_frame.pack(fill=X, pady=2)
                            
                            # 创建超链接样式的标签
                            reset_label = ttk.Label(reset_frame, text="<重置转音>", 
                                                  foreground="blue", cursor="hand2",
                                                  font=('微软雅黑', 9, 'underline'))
                            reset_label.pack(side=RIGHT, padx=5)
                            
                            # 绑定点击事件
                            reset_label.bind("<Button-1>", lambda event, idx=track_index, 
                                            t_var=track_transpose_var, o_var=track_octave_var:
                                            self.reset_track_transpose(idx, t_var, o_var))
                            
                            # 存储音轨信息
                            self.tracks_info.append({
                                "track_index": track_index,
                                "note_count": info["note_count"],
                                "frame": track_frame,
                                "checkbox_var": track_var,
                                "analysis_label": analysis_label
                            })
                            
                            # 存储转音和转位变量
                            self.track_transpose_vars[track_index] = track_transpose_var
                            self.track_octave_vars[track_index] = track_octave_var
                            
                            # 存储track_vars和分析结果
                            self.track_vars[track_index] = track_var
                            self.track_analysis_results[track_index] = {}
                            
                            # 保存UI元素引用
                            self.track_ui_elements[track_index] = {
                                "frame": track_frame,
                                "checkbox": track_checkbox,
                                "analysis_label": analysis_label
                            }
                        
                        self.selected_tracks = selected_tracks
                        
                        # 保存MIDI文件路径
                        self.current_file_path = file_path
                        
                        # 直接使用之前生成的事件数据和分析结果
                        # 但需要根据当前的选中音轨过滤事件
                        filtered_events = [e for e in events if e['track'] in selected_tracks]
                        self.current_events = filtered_events
                        self.current_analysis_result = analysis_result
                        
                        elapsed_time = time.time() - start_time
                        print(f"成功加载和分析MIDI文件：{file_path}，共找到{len(self.tracks_info)}个有效音轨，生成{len(self.current_events)}个事件，耗时{elapsed_time:.2f}秒")
                        
                        # 更新分析信息显示（异步进行，避免UI冻结）
                        self.root.after(100, self.update_analysis_info)
                        
                        # 启用试听MIDI按钮
                        if hasattr(self, 'midi_play_button'):
                            self.midi_play_button.config(state=NORMAL)
                        
                        # 初始化按钮状态：播放按钮亮，暂停按钮灰
                        if hasattr(self, 'play_button'):
                            self.play_button.config(state=NORMAL)
                        if hasattr(self, 'stop_button'):
                            self.stop_button.config(state=DISABLED)
                        
                        # 更新Canvas的滚动区域
                        self.track_canvas.configure(scrollregion=self.track_canvas.bbox("all"))
                            
                        # 启用试听MIDI按钮
                        if hasattr(self, 'midi_play_button'):
                            self.midi_play_button.config(state=NORMAL)
                    
                    # 在主线程中执行UI更新
                    self.root.after(0, update_ui)
                    
                except Exception as e:
                    error_msg = f"解析MIDI文件时出错: {str(e)}"
                    print(error_msg)
                    # 在主线程中显示错误
                    self.root.after(0, lambda: self._show_error_and_cleanup(error_msg, loading_label))
            
            # 启动解析线程
            thread = threading.Thread(target=parse_and_analyze_in_thread)
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            error_msg = f"启动MIDI解析线程时出错: {str(e)}"
            print(error_msg)
            messagebox.showerror("MIDI错误", error_msg)
            # 清理UI
            for widget in self.track_rows_frame.winfo_children():
                widget.destroy()
    
    def _fix_mojibake(self, text):
        """修复已被错误解码的字符串（ mojibake ）- 优化版本"""
        if not isinstance(text, str) or not text:
            return text
        
        # 快速检查：如果文本不包含可能乱码的字符，直接返回
        if not any(192 <= ord(c) <= 255 for c in text):
            return text
        
        # 记录原始文本用于调试
        original_text = text
        
        # 优化：优先尝试最常见的几种编码修复，减少不必要的尝试
        try:
            # 1. 首先尝试最常见的UTF-8被错误解码为Latin-1的情况
            # 这种模式最常见，成功率最高
            try:
                utf8_fixed = text.encode('latin-1', errors='replace').decode('utf-8', errors='replace')
                # 检查是否有中文字符且结果不同（表示修复成功）
                if any('\u4e00' <= c <= '\u9fff' for c in utf8_fixed) and utf8_fixed != text:
                    return utf8_fixed
            except Exception:
                pass
            
            # 2. 尝试GBK编码修复（中文环境常见）
            try:
                gbk_fixed = text.encode('latin-1', errors='replace').decode('gbk', errors='replace')
                if any('\u4e00' <= c <= '\u9fff' for c in gbk_fixed) and gbk_fixed != text:
                    return gbk_fixed
            except Exception:
                pass
            
            # 3. 尝试cp1252 -> UTF-8（Windows环境常见）
            try:
                cp1252_fixed = text.encode('cp1252', errors='replace').decode('utf-8', errors='replace')
                if any('\u4e00' <= c <= '\u9fff' for c in cp1252_fixed) and cp1252_fixed != text:
                    return cp1252_fixed
            except Exception:
                pass
            
            # 4. 如果以上都失败，尝试GB18030（支持更多字符）
            try:
                gb18030_fixed = text.encode('latin-1', errors='replace').decode('gb18030', errors='replace')
                if any('\u4e00' <= c <= '\u9fff' for c in gb18030_fixed) and gb18030_fixed != text:
                    return gb18030_fixed
            except Exception:
                pass
            
            # 5. 最后尝试直接UTF-8解码（针对特殊情况）
            try:
                # 检查是否包含UTF-8特征字节模式
                if any(0xC0 <= ord(c) <= 0xDF for c in text) or any(0xE0 <= ord(c) <= 0xEF for c in text):
                    raw_bytes = ''.join(chr(ord(c)) for c in text).encode('latin-1', errors='replace')
                    utf8_direct = raw_bytes.decode('utf-8', errors='replace')
                    if any('\u4e00' <= c <= '\u9fff' for c in utf8_direct) and utf8_direct != text:
                        return utf8_direct
            except Exception:
                pass
                
        except Exception:
            pass  # 静默处理异常，避免影响性能
        
        # 如果所有修复都失败，返回原始文本
        return text
    
    def track_selected(self, event=None):
        """选择音轨时的处理"""
        # 启用或禁用播放按钮 - 有选中音轨时亮
        if hasattr(self, 'play_button'):
            self.play_button.config(state=NORMAL if self.selected_tracks else DISABLED)
        
        # 启用或禁用停止按钮 - 默认为灰（只有播放时才亮）
        if hasattr(self, 'stop_button'):
            # 检查是否正在播放
            is_playing = hasattr(self, 'is_playing') and self.is_playing
            self.stop_button.config(state=NORMAL if is_playing else DISABLED)
        
        # 启用或禁用预览按钮
        if hasattr(self, 'preview_button'):
            self.preview_button.config(state=NORMAL if self.selected_tracks else DISABLED)
    
    def update_tracks_ui(self, all_tracks_selected, track_states):
        """更新音轨UI状态：复选框和高亮"""
        # 先清除所有选择，然后重新设置
        self.tracks_list.selection_set([])
        
        # 更新"全部音轨"项
        all_values = list(self.tracks_list.item(self.all_tracks_item, "values"))
        all_values[0] = "✓" if all_tracks_selected else "□"
        self.tracks_list.item(self.all_tracks_item, values=all_values)
        
        # 设置"全部音轨"项的高亮状态
        if all_tracks_selected:
            self.tracks_list.selection_add(self.all_tracks_item)
        
        # 更新所有子音轨
        for info in self.tracks_info:
            track_index = info['track_index']
            item_id = info['item_id']
            is_selected = track_states.get(track_index, False)
            
            # 更新复选框状态
            values = list(self.tracks_list.item(item_id, "values"))
            values[0] = "✓" if is_selected else "□"
            self.tracks_list.item(item_id, values=values)
            
            # 更新高亮状态
            if is_selected:
                self.tracks_list.selection_add(item_id)
    
    def on_track_click(self, event):
        """处理音轨列表的点击事件，实现复选框和整行选中功能"""
        # # 打印点击前的状态
        # print("===== 点击前状态 =====")
        # self._print_track_states()
        
        region = self.tracks_list.identify_region(event.x, event.y)
        item = self.tracks_list.identify_row(event.y)
        
        # 如果点击了无效项，直接返回
        if not item:
            return
            
        # 获取标签
        tags = self.tracks_list.item(item, "tags")
        
        # 处理任何行点击，不仅仅是复选框列
        if region == "cell" or region == "row":
            # 阻止事件传播，防止默认的选择行为
            event.widget.focus_set()
            
            # 检查是否点击的是"全部音轨"项
            if tags == "all_tracks" or (isinstance(tags, tuple) and "all_tracks" in tags):
                # 获取当前全部音轨状态
                all_values = list(self.tracks_list.item(self.all_tracks_item, "values"))
                current_all_selected = all_values[0] == "✓"
                
                # 切换全部音轨状态
                new_all_selected = not current_all_selected
                
                # 更新所有子音轨状态
                if new_all_selected:
                    self.selected_tracks = set(info['track_index'] for info in self.tracks_info)
                else:
                    self.selected_tracks.clear()
                
                # 创建音轨状态字典
                track_states = {}
                for info in self.tracks_info:
                    track_states[info['track_index']] = new_all_selected
                
                # 更新UI
                self.update_tracks_ui(new_all_selected, track_states)
            else:
                # 获取音轨索引
                if isinstance(tags, tuple) and tags:
                    track_index = int(tags[0])
                    
                    # 获取当前音轨状态
                    current_selected = track_index in self.selected_tracks
                    
                    # 切换音轨状态
                    new_selected = not current_selected
                    
                    # 更新selected_tracks集合
                    if new_selected:
                        self.selected_tracks.add(track_index)
                    else:
                        self.selected_tracks.remove(track_index)
                    
                    # 检查是否所有子音轨都被选中
                    all_selected = len(self.selected_tracks) == len(self.tracks_info)
                    
                    # 创建音轨状态字典
                    track_states = {}
                    for info in self.tracks_info:
                        track_states[info['track_index']] = info['track_index'] in self.selected_tracks
                    
                    # 更新UI
                    self.update_tracks_ui(all_selected, track_states)
        
        # # 打印点击后的状态
        # print("===== 点击后状态 =====")
        # self._print_track_states()
        
        # 更新按钮状态
        self.track_selected()
        
        # 先更新事件数据（音轨选择变动时预生成事件）
        self.update_event_data()
        
        # 然后更新分析信息显示
        self.update_analysis_info()
        
        # 确保返回'break'以阻止默认行为
        return 'break'
        
    def _print_track_states(self):
        # 打印所有音轨的状态
        for item in self.tracks_list.get_children():
            # 获取项目值（包括复选框状态）
            values = self.tracks_list.item(item, "values")
            checked = values[0] if values else "-"
            # 获取项目文本
            text = values[1] if len(values) > 1 else ""
            # 检查是否选中
            is_selected = item in self.tracks_list.selection()
            # 获取音轨索引
            tags = self.tracks_list.item(item, "tags")
            track_index = tags[0] if isinstance(tags, tuple) and tags else "N/A"
            print(f"项目: {item}, 索引: {track_index}, 文本: {text}, 复选框: {checked}, 选中状态: {is_selected}")
        # 打印选中的音轨集合
        print(f"选中的音轨集合: {self.selected_tracks}")
        print(f"全部音轨项ID: {self.all_tracks_item}")
    
    def toggle_select_all(self):
        """切换全选状态"""
        # 获取当前状态
        is_selected = self.all_tracks_var.get()
        
        # 更新所有音轨的选中状态
        self.selected_tracks.clear()
        for track_info in self.tracks_info:
            track_index = track_info["track_index"]
            track_info["checkbox_var"].set(is_selected)
            
            if is_selected:
                self.selected_tracks.add(track_index)
        
        # 更新分析信息
        self.update_analysis_info()
        # 更新事件数据
        self.update_event_data()
        
        # 打印状态
        print(f"全选状态更新: {'全选' if is_selected else '取消全选'}，当前选中{len(self.selected_tracks)}个音轨")
    
    def toggle_track_selection(self, track_index):
        """切换单个音轨的选中状态"""
        # 查找对应的音轨信息
        for track_info in self.tracks_info:
            if track_info["track_index"] == track_index:
                is_selected = track_info["checkbox_var"].get()
                
                # 更新选中集合
                if is_selected:
                    self.selected_tracks.add(track_index)
                else:
                    self.selected_tracks.discard(track_index)
                
                # 更新全选复选框状态
                self._update_all_tracks_checkbox()
                
                # 更新分析信息和事件数据
                self.update_analysis_info()
                self.update_event_data()
                
                print(f"音轨{track_index}选择状态: {'选中' if is_selected else '取消选中'}")
                break
    
    def _update_all_tracks_checkbox(self):
        """更新全选复选框的状态"""
        if not self.tracks_info:
            self.all_tracks_var.set(False)
            return
        
        all_selected = all(track_info["checkbox_var"].get() for track_info in self.tracks_info)
        self.all_tracks_var.set(all_selected)
    
    def adjust_track_transpose(self, track_index, var, delta):
        """调整单个音轨的移调值"""
        current_value = var.get()
        var.set(current_value + delta)
        self.on_track_transpose_change(track_index, var)
    
    def adjust_track_octave(self, track_index, var, delta):
        """调整单个音轨的转位值"""
        current_value = var.get()
        var.set(current_value + delta)
        self.on_track_octave_change(track_index, var)
    
    def on_track_transpose_change(self, track_index, var):
        """当单个音轨的移调值改变时"""
        # 调用统一的更新函数
        self.update_track_transpose(track_index, transpose_var=var)
    
    def update_track_transpose(self, track_index, transpose_value=None, octave_value=None, transpose_var=None, octave_var=None):
        """统一处理音轨转音设置更新（移调和转位）"""
        # 如果提供了变量引用，确保值是整数
        if transpose_var:
            try:
                value = transpose_var.get()
                transpose_var.set(int(value))
                print(f"[DEBUG] 音轨{track_index}移调值: {transpose_var.get()}")
            except ValueError:
                transpose_var.set(0)
                print(f"[DEBUG] 音轨{track_index}移调值无效，已重置为0")
        
        if octave_var:
            try:
                value = octave_var.get()
                octave_var.set(int(value))
                print(f"[DEBUG] 音轨{track_index}转位值: {octave_var.get()}")
            except ValueError:
                octave_var.set(0)
                print(f"[DEBUG] 音轨{track_index}转位值无效，已重置为0")
        
        # 如果提供了具体值，直接设置
        if transpose_value is not None and transpose_var:
            transpose_var.set(transpose_value)
            print(f"[DEBUG] 音轨{track_index}移调值设置为: {transpose_value}")
            
        if octave_value is not None and octave_var:
            octave_var.set(octave_value)
            print(f"[DEBUG] 音轨{track_index}转位值设置为: {octave_value}")
        
        # 重新分析该音轨
        if hasattr(self, 'selected_tracks') and track_index in self.selected_tracks:
            self._analyze_single_track(track_index)
        
        # 更新事件数据
        self.update_event_data()
        print(f"[DEBUG] 音轨{track_index}转音设置已更新，事件数据已刷新")
    
    def reset_track_transpose(self, track_index, transpose_var, octave_var):
        """重置音轨的移调和转位设置"""
        # 调用统一的更新函数，设置移调和转位为0
        self.update_track_transpose(track_index, transpose_value=0, octave_value=0, 
                                  transpose_var=transpose_var, octave_var=octave_var)
        print(f"音轨{track_index}已重置转音设置")
    
    def on_track_octave_change(self, track_index, var):
        """当单个音轨的转位值改变时"""
        # 调用统一的更新函数
        self.update_track_transpose(track_index, octave_var=var)
    
    def _analyze_single_track(self, track_index):
        """分析单个音轨"""
        if not self.current_file_path:
            return
        
        try:
            # 导入groups.py中的标准函数
            from groups import get_note_name, group_for_note
            
            # 获取配置的最低音和最高音
            try:
                from midi_analyzer import MidiAnalyzer
                config_min_note, config_max_note, _ = MidiAnalyzer._get_key_settings()
            except:
                config_min_note, config_max_note = 48, 83  # 默认值
            
            import mido
            mid = mido.MidiFile(self.current_file_path)
            track = mid.tracks[track_index]
            
            # 获取转音和转位值
            transpose = self.track_transpose_vars.get(track_index, tk.IntVar(value=0)).get()
            octave = self.track_octave_vars.get(track_index, tk.IntVar(value=0)).get()
            
            # 分析音轨并应用偏移
            notes = []
            shifted_notes = []  # 保存应用偏移后的音符
            for msg in track:
                if msg.type == 'note_on' and msg.velocity > 0:
                    notes.append(msg.note)
                    # 计算应用偏移后的音符值
                    total_offset = transpose + (octave * 12)
                    shifted_notes.append(msg.note + total_offset)
            
            if not notes:
                # 如果没有音符，设置空分析结果并更新UI
                analysis_text = "此音轨没有音符数据"
                # 更新分析结果
                self.track_analysis_results[track_index] = {
                    "state": "completed",
                    "analysis_text": analysis_text
                }
                
                # 在主线程中更新UI显示
                def update_ui_empty():
                    for track_info in self.tracks_info:
                        if track_info["track_index"] == track_index:
                            try:
                                label = track_info["analysis_label"]
                                label.configure(text=analysis_text)
                            except Exception as e:
                                print(f"更新空音轨UI时出错: {str(e)}")
                            break
                
                self.root.after(0, update_ui_empty)
                return
            
            # 计算最高音和最低音
            max_note = max(notes)  # 原始最高音
            min_note = min(notes)  # 原始最低音
            
            # 计算应用偏移后的最高音和最低音
            shifted_max_note = max(shifted_notes)
            shifted_min_note = min(shifted_notes)
            
            # 计算超限数量（使用配置的范围）
            upper_over_limit = sum(1 for n in shifted_notes if n > config_max_note)
            lower_over_limit = sum(1 for n in shifted_notes if n < config_min_note)
            
            # 全面检查是否超限：
            # 1. 检查最高音是否超过上限 或 低于下限
            # 2. 检查最低音是否低于下限 或 超过上限
            is_max_over_limit = shifted_max_note > config_max_note or shifted_max_note < config_min_note
            is_min_over_limit = shifted_min_note < config_min_note or shifted_min_note > config_max_note
            
            # 额外检查整个音域是否在配置范围内
            is_range_valid = config_min_note <= shifted_min_note and shifted_max_note <= config_max_note
            
            # 输出详细的超限检查信息
            print(f"超限检查详细信息：")
            print(f"  应用偏移后最高音: {shifted_max_note}, 最低音: {shifted_min_note}")
            print(f"  配置范围: {config_min_note}-{config_max_note}")
            print(f"  最高音是否超限: {is_max_over_limit} (超过上限: {shifted_max_note > config_max_note}, 低于下限: {shifted_max_note < config_min_note})")
            print(f"  最低音是否超限: {is_min_over_limit} (低于下限: {shifted_min_note < config_min_note}, 超过上限: {shifted_min_note > config_max_note})")
            
            # 获取音符名称和分组（使用应用偏移后的值）
            shifted_max_note_name = get_note_name(shifted_max_note)
            shifted_min_note_name = get_note_name(shifted_min_note)
            max_octave_group = group_for_note(shifted_max_note)
            min_octave_group = group_for_note(shifted_min_note)
            
            # 构建分析结果字典 - 保存应用偏移后的值用于建议计算
            analysis_result = {
                'max_note': shifted_max_note,  # 保存应用偏移后的最高音
                'min_note': shifted_min_note,  # 保存应用偏移后的最低音
                'original_max_note': max_note,  # 保存原始最高音
                'original_min_note': min_note,  # 保存原始最低音
                'is_max_over_limit': is_max_over_limit,
                'is_min_over_limit': is_min_over_limit,
                'upper_over_limit': upper_over_limit,
                'lower_over_limit': lower_over_limit
            }
            
            # 输出日志用于调试
            print(f"音轨{track_index}分析调试:")
            print(f"  原始最高音: {max_note}, 原始最低音: {min_note}")
            print(f"  当前移调: {transpose}, 当前转位: {octave}")
            print(f"  应用偏移后最高音: {shifted_max_note}, 最低音: {shifted_min_note}")
            print(f"  配置范围: {config_min_note}-{config_max_note}")
            print(f"  是否超限: 最高音{is_max_over_limit}({shifted_max_note} > {config_max_note}), 最低音{is_min_over_limit}({shifted_min_note} < {config_min_note})")
            # 额外检查最高音是否低于配置最小值
            is_max_below_min = shifted_max_note < config_min_note
            print(f"  最高音低于配置最小值: {is_max_below_min}({shifted_max_note} < {config_min_note})")
            
            # 计算建议 - 基于应用偏移后的值
            suggestion_text = self._calculate_transpose_suggestion(
                analysis_result, config_min_note, config_max_note, transpose, octave
            )
            
            # 构建分析文本 - 显示应用偏移后的音符值
            analysis_text = (
                f"最高音: {shifted_max_note_name}({shifted_max_note})  {max_octave_group}  {'超限' if upper_over_limit > 0 else '未超限'}  超限数量: {upper_over_limit}\n"
                f"最低音: {shifted_min_note_name}({shifted_min_note})  {min_octave_group}  {'超限' if lower_over_limit > 0 else '未超限'}  超限数量: {lower_over_limit}\n"
            )
            
            # 添加建议文本
            if suggestion_text:
                analysis_text += suggestion_text
            
            # 更新分析结果
            self.track_analysis_results[track_index] = {
                "max_note": max_note,
                "min_note": min_note,
                "is_max_over_limit": is_max_over_limit,
                "is_min_over_limit": is_min_over_limit,
                "upper_over_limit": upper_over_limit,
                "lower_over_limit": lower_over_limit,
                "analysis_text": analysis_text,
                "suggestion_text": suggestion_text,  # 保存建议文本用于超链接处理
                "state": "completed"  # 标记分析已完成
            }
            
            # 在主线程中更新UI显示，避免渲染问题
            def update_ui():
                for track_info in self.tracks_info:
                    if track_info["track_index"] == track_index:
                        try:
                            # 获取当前Text组件
                            text_widget = track_info["analysis_label"]
                            
                            # 启用编辑并更新内容
                            text_widget.config(state=tk.NORMAL)
                            text_widget.delete(1.0, tk.END)
                            text_widget.insert(tk.END, analysis_text)
                            
                            # 清除所有现有标签，避免重复绑定
                            for tag in text_widget.tag_names():
                                text_widget.tag_remove(tag, "1.0", tk.END)
                            
                            # 查找并添加超链接（如果需要）
                            if "<最高音>" in analysis_text:
                                start_pos = text_widget.search("<最高音>", "1.0", tk.END)
                                if start_pos:
                                    end_pos = f"{start_pos}+4c"
                                    text_widget.tag_add("max_note_link", start_pos, end_pos)
                                    text_widget.tag_config("max_note_link", foreground="blue", underline=True)
                                    text_widget.tag_bind("max_note_link", "<Button-1>", 
                                                        lambda e, idx=track_index: self._apply_max_note_suggestion(idx))
                            
                            if "<最低音>" in analysis_text:
                                start_pos = text_widget.search("<最低音>", "1.0", tk.END)
                                if start_pos:
                                    end_pos = f"{start_pos}+4c"
                                    text_widget.tag_add("min_note_link", start_pos, end_pos)
                                    text_widget.tag_config("min_note_link", foreground="blue", underline=True)
                                    text_widget.tag_bind("min_note_link", "<Button-1>", 
                                                        lambda e, idx=track_index: self._apply_min_note_suggestion(idx))
                            
                            # 禁用编辑
                            text_widget.config(state=tk.DISABLED)
                            print(f"成功更新音轨{track_index}的Text组件内容")
                            
                        except Exception as text_error:
                            # 如果更新失败，记录错误信息
                            print(f"更新Text组件失败: {str(text_error)}")
                            # 尝试备选方案：创建一个新的Text组件
                            try:
                                parent = text_widget.master
                                text_widget.pack_forget()
                                
                                # 创建新的Text组件，确保自适应宽度
                                new_text_widget = tk.Text(parent, height=4, font=('微软雅黑', 9), wrap=tk.WORD)
                                new_text_widget.insert(tk.END, analysis_text)
                                new_text_widget.pack(fill=X, padx=5, pady=1)  # 填充整个宽度
                                new_text_widget.config(state=tk.DISABLED)
                                
                                # 更新引用
                                track_info["analysis_label"] = new_text_widget
                                print(f"成功为音轨{track_index}创建新的Text组件")
                            except Exception as fallback_error:
                                print(f"备选方案也失败: {str(fallback_error)}")
                        except Exception as e:
                            print(f"更新音轨{track_index}UI时发生未预期错误: {str(e)}")
                        break
            
            # 使用after方法在主线程中执行UI更新
            self.root.after(0, update_ui)
            
        except Exception as e:
            print(f"分析音轨{track_index}时出错: {str(e)}")
            import traceback
            traceback.print_exc()
            
            # 出错时也尝试更新UI，显示错误信息
            def update_ui_error():
                for track_info in self.tracks_info:
                    if track_info["track_index"] == track_index:
                        try:
                            text_widget = track_info["analysis_label"]
                            # 对于Text组件，使用insert方法而不是configure
                            text_widget.config(state=tk.NORMAL)
                            text_widget.delete(1.0, tk.END)
                            text_widget.insert(tk.END, f"分析出错: {str(e)}")
                            text_widget.config(state=tk.DISABLED)
                        except:
                            pass
                        break
            
            self.root.after(0, update_ui_error)
    
    # 修改update_analysis_info方法以支持单音轨分析

    
    def adjust_value(self, var, delta):
        """调整数值变量"""
        var.set(var.get() + delta)
        # 重新生成事件数据（应用新的移调/转位设置）
        self.update_event_data()
        # 更新分析信息
        self.update_analysis_info()
    
    def on_transpose_octave_change(self):
        """当移调或转位值直接在输入框中修改时触发"""
        # 确保输入值是整数
        try:
            transpose_value = int(self.transpose_var.get())
            self.transpose_var.set(transpose_value)
        except ValueError:
            self.transpose_var.set(0)
        
        try:
            octave_value = int(self.octave_var.get())
            self.octave_var.set(octave_value)
        except ValueError:
            self.octave_var.set(0)
        
        # 重新生成事件数据（应用新的移调/转位设置）
        self.update_event_data()
        # 更新分析信息
        self.update_analysis_info()
    
    def _on_mousewheel(self, event):
        """处理鼠标滚轮事件，只用于音轨列表Canvas滚动"""
        # 歌曲列表的滚动已通过单独的_on_song_list_mousewheel方法处理
        
        # 首先检查是否有track_canvas和track_rows_frame
        if hasattr(self, 'track_canvas') and hasattr(self, 'track_rows_frame'):
            # 更新窗口大小
            self.track_canvas.update_idletasks()
            self.track_rows_frame.update_idletasks()
            
            # 获取Canvas的实际大小和内容大小
            canvas_height = self.track_canvas.winfo_height()
            content_height = self.track_rows_frame.winfo_height()
            
            # 只有当内容高度大于Canvas高度时才允许滚动
            if content_height > canvas_height + 10:
                # 根据操作系统不同，event.delta的单位可能不同
                if hasattr(event, 'delta'):
                    # Windows上的处理方式
                    delta = event.delta
                else:
                    # 其他系统的处理方式
                    delta = -event.delta * 120
                
                # 垂直滚动Canvas
                self.track_canvas.yview_scroll(int(-1 * (delta / 120)), "units")
            else:
                # 当内容不需要滚动时，确保内容始终在顶部
                self.track_canvas.yview_moveto(0)
        
    def _on_song_list_mousewheel(self, event):
        """处理歌曲列表的鼠标滚轮事件"""
        # 计算滚动方向
        delta = event.delta if hasattr(event, 'delta') else -event.num
        # 滚动歌曲列表
        self.song_list.yview_scroll(int(-1 * (delta / 120)), "units")
        # 在Tkinter中，return "break"可以阻止事件冒泡
        return "break"
    
    def _update_canvas_scrollregion(self):
        """更新Canvas的滚动区域，确保内容从顶部开始"""
        bbox = self.track_canvas.bbox("all")
        if bbox:
            # 确保y0是0，防止向上滚动过多
            self.track_canvas.configure(scrollregion=(0, 0, bbox[2], bbox[3]))
        
    def update_analysis_info(self):
        """更新分析信息 - 支持单独音轨分析"""
        if not self.current_file_path or not self.selected_tracks:
            return
        
        # 创建一个线程来处理分析，避免UI冻结
        def analyze_in_thread():
            try:
                # 为每个选中的音轨分析
                for track_index in list(self.selected_tracks):
                    self._analyze_single_track(track_index)
                
                # 更新Canvas的滚动区域，确保内容从顶部开始
                self.root.after(0, lambda: self._update_canvas_scrollregion())
                
            except Exception as e:
                print(f"分析音轨时出错: {str(e)}")
        
        # 启动分析线程
        thread = threading.Thread(target=analyze_in_thread)
        thread.daemon = True
        thread.start()
    
    def _update_analysis_text_widget(self):
        """更新分析文本组件，支持超链接"""
        # 启用编辑
        self.analysis_text_widget.config(state=tk.NORMAL)
        # 清空内容
        self.analysis_text_widget.delete(1.0, tk.END)
        
        # 插入文本
        self.analysis_text_widget.insert(tk.END, self.analysis_text)
        
        # 查找并添加超链接
        # 处理最高音超链接
        if "<最高音>" in self.analysis_text:
            start_index = self.analysis_text.find("<最高音>")
            end_index = start_index + len("<最高音>")
            
            # 计算在Text组件中的位置
            lines_before = self.analysis_text[:start_index].count('\n')
            if lines_before > 0:
                last_newline = self.analysis_text[:start_index].rfind('\n')
                char_in_line = start_index - last_newline - 1
                line_start = lines_before + 1
            else:
                char_in_line = start_index
                line_start = 1
            
            # 先移除旧的标签（如果存在）
            self.analysis_text_widget.tag_remove("max_note_link", "1.0", tk.END)
            # 添加超链接样式
            self.analysis_text_widget.tag_add("max_note_link", f"{line_start}.{char_in_line}", f"{line_start}.{char_in_line + 4}")
            self.analysis_text_widget.tag_config("max_note_link", foreground="blue", underline=True, cursor="hand2")
            
            # 绑定点击事件
            self.analysis_text_widget.tag_bind("max_note_link", "<Button-1>", lambda e: self._apply_max_note_suggestion())
        
        # 处理最低音超链接
        if "<最低音>" in self.analysis_text:
            start_index = self.analysis_text.find("<最低音>")
            end_index = start_index + len("<最低音>")
            
            # 计算在Text组件中的位置
            lines_before = self.analysis_text[:start_index].count('\n')
            if lines_before > 0:
                last_newline = self.analysis_text[:start_index].rfind('\n')
                char_in_line = start_index - last_newline - 1
                line_start = lines_before + 1
            else:
                char_in_line = start_index
                line_start = 1
            
            # 先移除旧的标签（如果存在）
            self.analysis_text_widget.tag_remove("min_note_link", "1.0", tk.END)
            # 添加超链接样式
            self.analysis_text_widget.tag_add("min_note_link", f"{line_start}.{char_in_line}", f"{line_start}.{char_in_line + 4}")
            self.analysis_text_widget.tag_config("min_note_link", foreground="blue", underline=True, cursor="hand2")
            
            # 绑定点击事件
            self.analysis_text_widget.tag_bind("min_note_link", "<Button-1>", lambda e: self._apply_min_note_suggestion())
        
        # 禁用编辑
        self.analysis_text_widget.config(state=tk.DISABLED)
    
    def _calculate_transpose_suggestion(self, analysis_result, config_min_note, config_max_note, current_transpose, current_octave):
        """计算建议的移调和转位值
        
        Args:
            analysis_result: 分析结果字典
            config_min_note: 配置的最低音
            config_max_note: 配置的最高音
            current_transpose: 当前移调值
            current_octave: 当前转位值
            
        Returns:
            str: 建议文本，包含超链接
        """
        # 初始化存储建议的字典（如果不存在）
        if not hasattr(self, 'suggestion_cache'):
            self.suggestion_cache = {}
        
        # 检查是否有有效的最高音和最低音数据
        if analysis_result['max_note'] is None or analysis_result['min_note'] is None:
            return ""
        
        # 检查是否超限，只有超限时才显示建议
        max_over_limit = analysis_result.get('is_max_over_limit', False)
        min_over_limit = analysis_result.get('is_min_over_limit', False)
        
        # 如果都没有超限，不显示建议
        if not max_over_limit and not min_over_limit:
            return ""
        
        # 计算最高音的建议移调和转位（只在超限时计算）
        max_suggestion_text = ""
        if max_over_limit:
            max_diff = config_max_note - analysis_result['max_note']
            # 优化建议逻辑：以移调+转位的绝对值最小为准，优先选择5、6、7
            max_suggestions = self._optimize_transpose_suggestion(max_diff, current_transpose, current_octave)
            if max_suggestions:
                best_suggestion = max_suggestions[0]  # 取最优解
                # 存储建议结果，供后续应用时使用
                self.suggestion_cache['max_note'] = {
                    'transpose': best_suggestion['transpose'],
                    'octave': best_suggestion['octave']
                }
                # 输出日志用于调试
                print(f"最高音建议计算调试:")
                print(f"  当前移调: {current_transpose}, 当前转位: {current_octave}")
                print(f"  音高差: {max_diff}")
                print(f"  显示的建议移调: {best_suggestion['transpose']}, 显示的建议转位: {best_suggestion['octave']}")
                max_suggestion_text = f"<最高音>移调{best_suggestion['transpose']}，转位{best_suggestion['octave']}"
        
        # 计算最低音的建议移调和转位（只在超限时计算）
        min_suggestion_text = ""
        if min_over_limit:
            min_diff = config_min_note - analysis_result['min_note']
            # 优化建议逻辑：以移调+转位的绝对值最小为准，优先选择5、6、7
            min_suggestions = self._optimize_transpose_suggestion(min_diff, current_transpose, current_octave)
            if min_suggestions:
                best_suggestion = min_suggestions[0]  # 取最优解
                # 存储建议结果，供后续应用时使用
                self.suggestion_cache['min_note'] = {
                    'transpose': best_suggestion['transpose'],
                    'octave': best_suggestion['octave']
                }
                # 输出日志用于调试
                print(f"最低音建议计算调试:")
                print(f"  当前移调: {current_transpose}, 当前转位: {current_octave}")
                print(f"  音高差: {min_diff}")
                print(f"  显示的建议移调: {best_suggestion['transpose']}, 显示的建议转位: {best_suggestion['octave']}")
                min_suggestion_text = f"<最低音>移调{best_suggestion['transpose']}，转位{best_suggestion['octave']}"
        
        # 构建建议文本
        suggestion_text = "建议"
        if max_suggestion_text and min_suggestion_text:
            suggestion_text += f"{max_suggestion_text}  {min_suggestion_text}"
        elif max_suggestion_text:
            suggestion_text += f"{max_suggestion_text}"
        elif min_suggestion_text:
            suggestion_text += f"{min_suggestion_text}"
        
        return suggestion_text
    
    def _optimize_transpose_suggestion(self, diff, current_transpose, current_octave):
        """优化移调建议逻辑，以移调+转位的绝对值最小为准，优先选择5、6、7
        
        Args:
            diff: 需要调整的音符差值
            current_transpose: 当前移调值
            current_octave: 当前转位值
            
        Returns:
            list: 排序后的建议列表，每个元素为{'transpose': 移调值, 'octave': 转位值, 'score': 评分}
        """
        suggestions = []
        
        # 生成所有可能的移调+转位组合（-2到+2个八度）
        for octave_shift in range(-2, 3):
            # 计算需要的总移调量
            total_transpose_needed = diff - (octave_shift * 12)
            
            # 计算最终的移调和转位值（叠加到当前设置上）
            final_transpose = current_transpose + total_transpose_needed
            final_octave = current_octave + octave_shift
            
            # 计算评分：移调+转位的绝对值（越小越好）
            # 使用绝对值进行评分，确保正数和负数有相同的权重
            score = abs(final_transpose) + abs(final_octave)
            
            # 如果移调值的绝对值在5、6、7范围内，给予额外加分（优先级更高）
            if 5 <= abs(final_transpose) <= 7:
                score -= 0.5  # 给予加分，使这些值优先级更高
            
            suggestions.append({
                'transpose': final_transpose,
                'octave': final_octave,
                'score': score
            })
        
        # 按评分排序（评分越小越优先）
        suggestions.sort(key=lambda x: x['score'])
        
        return suggestions
    
    def _apply_max_note_suggestion(self, track_index=None):
        """应用最高音的建议移调和转位设置到指定音轨"""
        # 如果没有指定音轨，使用当前选中的第一个音轨
        if track_index is None:
            if not self.selected_tracks:
                messagebox.showinfo("提示", "请先选择一个音轨")
                return
            track_index = next(iter(self.selected_tracks))
        
        # 检查是否有缓存的建议结果
        if not hasattr(self, 'suggestion_cache') or 'max_note' not in self.suggestion_cache:
            messagebox.showinfo("提示", "没有可用的最高音建议")
            return
        
        # 获取缓存的建议结果
        suggestion = self.suggestion_cache['max_note']
        suggested_transpose = suggestion['transpose']
        suggested_octave = suggestion['octave']
        
        # 输出日志用于调试
        print(f"最高音建议应用调试:")
        print(f"  直接应用缓存的建议值: 移调{suggested_transpose}, 转位{suggested_octave}")
        
        # 直接应用缓存的建议结果
        self.track_transpose_vars[track_index].set(suggested_transpose)
        self.track_octave_vars[track_index].set(suggested_octave)
        
        # 更新事件数据
        self.update_event_data()
        # 重新分析音轨
        self._analyze_single_track(track_index)
    
    def _apply_min_note_suggestion(self, track_index=None):
        """应用最低音的建议移调和转位设置到指定音轨"""
        # 如果没有指定音轨，使用当前选中的第一个音轨
        if track_index is None:
            if not self.selected_tracks:
                messagebox.showinfo("提示", "请先选择一个音轨")
                return
            track_index = next(iter(self.selected_tracks))
        
        # 检查是否有缓存的建议结果
        if not hasattr(self, 'suggestion_cache') or 'min_note' not in self.suggestion_cache:
            messagebox.showinfo("提示", "没有可用的最低音建议")
            return
        
        # 获取缓存的建议结果
        suggestion = self.suggestion_cache['min_note']
        suggested_transpose = suggestion['transpose']
        suggested_octave = suggestion['octave']
        
        # 输出日志用于调试
        print(f"最低音建议应用调试:")
        print(f"  直接应用缓存的建议值: 移调{suggested_transpose}, 转位{suggested_octave}")
        
        # 直接应用缓存的建议结果
        self.track_transpose_vars[track_index].set(suggested_transpose)
        self.track_octave_vars[track_index].set(suggested_octave)
        
        # 更新事件数据
        self.update_event_data()
        # 重新分析音轨
        self._analyze_single_track(track_index)
    
    def _apply_transpose_suggestion(self):
        """应用建议的移调和转位设置"""
        # 获取当前的分析结果
        if not hasattr(self, 'current_analysis_result') or not self.current_analysis_result:
            messagebox.showinfo("提示", "没有可用的分析结果")
            return
        
        # 获取配置信息
        config_min_note = 48  # 默认值
        config_max_note = 83  # 默认值
        try:
            from midi_analyzer import MidiAnalyzer
            config_min_note, config_max_note, _ = MidiAnalyzer._get_key_settings()
        except:
            pass
        
        # 获取当前设置
        current_transpose = self.transpose_var.get()
        current_octave = self.octave_var.get()
        
        # 计算建议值（使用新的优化逻辑）
        analysis_result = self.current_analysis_result
        
        # 检查是否超限
        max_over_limit = analysis_result.get('is_max_over_limit', False)
        min_over_limit = analysis_result.get('is_min_over_limit', False)
        
        # 优先应用最高音的建议（避免超出上限）
        if max_over_limit:
            max_diff = config_max_note - analysis_result['max_note']
            suggestions = self._optimize_transpose_suggestion(max_diff, current_transpose, current_octave)
            if suggestions:
                best_suggestion = suggestions[0]
                self.transpose_var.set(best_suggestion['transpose'])
                self.octave_var.set(best_suggestion['octave'])
        elif min_over_limit:
            # 如果没有最高音超限，但最低音超限，应用最低音建议
            min_diff = config_min_note - analysis_result['min_note']
            suggestions = self._optimize_transpose_suggestion(min_diff, current_transpose, current_octave)
            if suggestions:
                best_suggestion = suggestions[0]
                self.transpose_var.set(best_suggestion['transpose'])
                self.octave_var.set(best_suggestion['octave'])
        else:
            messagebox.showinfo("提示", "没有需要调整的超限音符")
            return
        
        # 更新显示
        self.update_analysis_info()
        
        # 重新解析事件表
        self.update_event_data()
        
        messagebox.showinfo("提示", f"已应用建议设置：移调{self.transpose_var.get()}，转位{self.octave_var.get()}")
    
    def _apply_max_note_suggestion_global(self):
        """应用最高音的建议移调和转位设置（全局）"""
        # 获取当前的分析结果
        if not hasattr(self, 'current_analysis_result') or not self.current_analysis_result:
            messagebox.showinfo("提示", "没有可用的分析结果")
            return
        
        # 获取配置信息
        config_min_note = 48  # 默认值
        config_max_note = 83  # 默认值
        try:
            from midi_analyzer import MidiAnalyzer
            config_min_note, config_max_note, _ = MidiAnalyzer._get_key_settings()
        except:
            pass
        
        # 获取当前设置
        current_transpose = self.transpose_var.get()
        current_octave = self.octave_var.get()
        
        # 计算最高音的建议移调和转位（使用新的优化逻辑）
        analysis_result = self.current_analysis_result
        max_diff = config_max_note - analysis_result['max_note']
        
        # 检查最高音是否超限
        if not analysis_result.get('is_max_over_limit', False):
            messagebox.showinfo("提示", "最高音没有超限，无需调整")
            return
        
        # 使用优化逻辑计算建议
        suggestions = self._optimize_transpose_suggestion(max_diff, current_transpose, current_octave)
        if suggestions:
            best_suggestion = suggestions[0]
            self.transpose_var.set(best_suggestion['transpose'])
            self.octave_var.set(best_suggestion['octave'])
        
        # 重新解析事件表（这会重新分析MIDI文件并更新分析结果）
        self.update_event_data()
        
        # 更新显示（使用新的分析结果）
        self.update_analysis_info()
    
    # 删除重复的_apply_min_note_suggestion方法，保留第1924行有track_index参数的版本
    
    def _update_play_button_during_countdown(self, remaining_seconds):
        """倒计时期间更新播放按钮文本
        
        Args:
            remaining_seconds: 剩余秒数
        """
        try:
            print(f"更新按钮倒计时文本: 播放 ({remaining_seconds}s)")
            # 确保在UI线程中更新
            def update_button():
                try:
                    self.play_button.config(text=f"播放 ({remaining_seconds}s)")
                    print(f"按钮文本已更新为: 播放 ({remaining_seconds}s)")
                except Exception as inner_e:
                    print(f"直接更新按钮文本时出错: {str(inner_e)}")
            
            # 使用after方法在UI线程中更新按钮文本
            self.play_button.after(0, update_button)
        except Exception as e:
            print(f"更新按钮倒计时文本时出错: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def _update_play_button_to_pause(self):
        """倒计时结束后将按钮文本更新为暂停"""
        try:
            print("更新按钮文本为: 暂停")
            # 确保在UI线程中更新
            def update_button():
                try:
                    self.play_button.config(text="暂停")
                    print("按钮文本已更新为: 暂停")
                except Exception as inner_e:
                    print(f"直接更新按钮文本时出错: {str(inner_e)}")
            
            # 使用after方法在UI线程中更新按钮文本
            self.play_button.after(0, update_button)
        except Exception as e:
            print(f"更新按钮文本为暂停时出错: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def toggle_play(self):
        """切换播放/暂停状态"""
        if not self.midi_player.playing:
            self.start_playback()
        else:
            # 检查是否处于暂停状态
            if self.midi_player.paused:
                # 如果是暂停状态，恢复播放
                try:
                    # 立即更新按钮文本为倒计时状态，避免延迟
                    self._update_play_button_during_countdown(3)
                    
                    # 在恢复播放倒计时开始时，显示正确的剩余时间
                    if self.midi_player and hasattr(self.midi_player, 'total_time') and hasattr(self.midi_player, 'get_current_time'):
                        total_time = self.midi_player.get_total_time()
                        current_time = self.midi_player.get_current_time()  # 应该是0，因为在倒计时中
                        remaining_time = max(0, total_time - current_time)
                        minutes, seconds = divmod(int(remaining_time), 60)
                        time_str = f"剩余时间: {minutes:02d}:{seconds:02d}"
                        self.time_label.config(text=time_str)
                    
                    # 创建一个函数来封装恢复播放的逻辑
                    def resume_playback_thread():
                        # 定义恢复播放时的倒计时回调
                        def resume_countdown_callback(remaining_seconds):
                            self._update_play_button_during_countdown(remaining_seconds)
                        
                        # 定义倒计时完成回调
                        def resume_completion_callback():
                            self._update_play_button_to_pause()
                        
                        # 执行恢复播放，包含倒计时和完成回调
                        self.midi_player.resume(
                            countdown_callback=resume_countdown_callback,
                            completion_callback=resume_completion_callback
                        )
                    
                    # 在单独的线程中执行恢复播放，避免阻塞UI
                    threading.Thread(target=resume_playback_thread, daemon=True).start()
                except Exception as e:
                    print(f"恢复播放时出错: {str(e)}")
                    messagebox.showerror("播放错误", f"恢复播放时出错: {str(e)}")
            else:
                # 如果不是暂停状态，暂停播放
                self.pause_playback()
    
    def start_playback(self):
        """开始播放，使用预处理的事件表"""
        try:
            if hasattr(self, 'current_file_path') and self.current_file_path:
                # 如果正在播放MIDI，先停止
                if hasattr(self, 'is_playing_midi') and self.is_playing_midi:
                    self.stop_midi_playback()
                
                # 确保有事件数据，如果没有则更新
                if not hasattr(self, 'current_events') or not self.current_events:
                    self.update_event_data()
                
                # 检查是否有有效的事件数据
                if not self.current_events:
                    messagebox.showerror("错误", "没有有效的事件数据，请重新加载MIDI文件")
                    return
                
                # 立即更新按钮文本为倒计时状态，避免延迟
                self._update_play_button_during_countdown(3)
                
                # 使用新的play_from_events方法播放预处理的事件表
                # 计算总时长
                total_time = None
                if self.current_events:
                    max_time = max(event['time'] for event in self.current_events)
                    total_time = max_time
                    # 在倒计时开始时，显示总时长作为剩余时间
                    minutes, seconds = divmod(int(total_time), 60)
                    time_str = f"剩余时间: {minutes:02d}:{seconds:02d}"
                    self.time_label.config(text=time_str)
                
                # 定义开始播放时的倒计时回调
                def play_countdown_callback(remaining_seconds):
                    self._update_play_button_during_countdown(remaining_seconds)
                
                # 定义完整的播放线程函数
                def play_thread_function():
                    try:
                        # 定义倒计时完成回调
                        def play_completion_callback():
                            self._update_play_button_to_pause()
                        
                        # 播放开始，传入倒计时回调和完成回调
                        self.midi_player.play_from_events(
                            self.current_events, 
                            total_time, 
                            countdown_callback=play_countdown_callback,
                            completion_callback=play_completion_callback
                        )
                    except Exception as e:
                        print(f"播放线程错误: {str(e)}")
                        import traceback
                        traceback.print_exc()
                
                threading.Thread(target=play_thread_function, 
                               daemon=True).start()
                
                self.stop_button.config(state=NORMAL)
        except Exception as e:
            print(f"开始播放时出错: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("播放错误", f"开始播放时出错: {str(e)}")
    
    def pause_playback(self):
        """暂停播放"""
        try:
            # 调用pause方法并检查是否成功暂停
            if self.midi_player.pause():
                # 只有成功暂停时才更新按钮文本
                self.play_button.config(text="播放")
        except Exception as e:
            print(f"暂停播放时出错: {str(e)}")
            messagebox.showerror("播放错误", f"暂停播放时出错: {str(e)}")
    
    def stop_playback(self):
        """停止播放"""
        try:
            # 停止播放，包括停止任何正在进行的倒计时
            self.midi_player.stop()
            
            # 使用UI线程安全的方式更新按钮状态
            def update_buttons():
                try:
                    # 立即将播放按钮文本设置为"播放"
                    self.play_button.config(text="播放", state=NORMAL)
                    self.stop_button.config(state=DISABLED)
                    print("按钮状态已更新为播放状态")
                except Exception as inner_e:
                    print(f"更新按钮状态时出错: {str(inner_e)}")
            
            # 使用after方法在UI线程中更新按钮
            self.root.after(0, update_buttons)
            
        except Exception as e:
            print(f"停止播放时出错: {str(e)}")
            messagebox.showerror("播放错误", f"停止播放时出错: {str(e)}")
    
    def toggle_midi_playback(self):
        """切换MIDI直接播放状态"""
        if not self.is_playing_midi:
            self.start_midi_playback()
        else:
            self.stop_midi_playback()
    
    def start_midi_playback(self):
        """直接播放原始MIDI文件"""
        try:
            if hasattr(self, 'current_file_path') and self.current_file_path:
                # 停止可能正在进行的预览
                if hasattr(self, 'is_previewing') and self.is_previewing:
                    self.stop_preview()
                
                # 停止可能正在进行的MIDI播放
                pygame.mixer.music.stop()
                
                # 加载MIDI文件
                pygame.mixer.music.load(self.current_file_path)
                
                # 获取MIDI文件时长
                try:
                    # 使用mido库获取MIDI文件时长
                    import mido
                    mid = mido.MidiFile(self.current_file_path)
                    self.midi_total_duration = mid.length
                    print(f"[试听] 使用mido获取MIDI文件时长: {int(self.midi_total_duration // 60)}分{int(self.midi_total_duration % 60)}秒 ({self.midi_total_duration:.2f}秒)")
                except Exception as e:
                    print(f"[试听] 获取MIDI文件时长失败: {str(e)}")
                    self.midi_total_duration = 0
                
                # 记录开始播放的时间
                self.midi_start_time = time.time()
                
                # 开始播放MIDI文件
                pygame.mixer.music.play()
                
                # 更新状态
                self.is_playing_midi = True
                # 使用ttkbootstrap的内置Danger样式
                self.midi_play_button.config(text="停止MIDI")
                
                # 禁用其他播放按钮
                self.play_button.config(state=DISABLED)
                self.preview_button.config(state=DISABLED)
                
                # 定期检查播放是否结束并更新时间
                self.check_midi_playback()
                self.update_midi_playback_time()
                
        except Exception as e:
            print(f"播放MIDI文件时出错: {str(e)}")
            messagebox.showerror("播放错误", f"无法播放MIDI文件: {str(e)}")
            self.is_playing_midi = False
    
    def stop_midi_playback(self):
        """停止直接播放MIDI文件"""
        pygame.mixer.music.stop()
        self.is_playing_midi = False
        # 使用ttkbootstrap的内置Primary样式
        self.midi_play_button.config(text="试听MIDI")
        
        # 清理MIDI播放相关的变量
        if hasattr(self, 'midi_total_duration'):
            delattr(self, 'midi_total_duration')
        if hasattr(self, 'midi_start_time'):
            delattr(self, 'midi_start_time')
            
        # 重置剩余时间显示
        self.update_remaining_time_label("00:00")
        
        # 重新启用其他播放按钮
        if hasattr(self, 'current_file_path') and self.current_file_path:
            self.play_button.config(state=NORMAL)
            self.preview_button.config(state=NORMAL)
    
    def check_midi_playback(self):
        """检查MIDI播放是否结束"""
        if self.is_playing_midi:
            if not pygame.mixer.music.get_busy():
                # 播放已结束
                self.stop_midi_playback()
            else:
                # 继续检查
                self.root.after(1000, self.check_midi_playback)
    
    def update_midi_playback_time(self):
        """更新MIDI播放的剩余时间"""
        if self.is_playing_midi and hasattr(self, 'midi_total_duration') and hasattr(self, 'midi_start_time'):
            # 计算已播放时间
            elapsed_time = time.time() - self.midi_start_time
            # 计算剩余时间
            remaining_time = max(0, self.midi_total_duration - elapsed_time)
            remaining_minutes = int(remaining_time // 60)
            remaining_seconds = int(remaining_time % 60)
            
            # 更新剩余时间标签
            remaining_time_str = f"{str(remaining_minutes).zfill(2)}:{str(remaining_seconds).zfill(2)}"
            self.update_remaining_time_label(remaining_time_str)
            
            # 继续更新
            self.root.after(500, self.update_midi_playback_time)
    
    def toggle_preview(self):
        """切换预览状态"""
        if not self.is_previewing:
            # 如果正在播放MIDI，先停止
            if hasattr(self, 'is_playing_midi') and self.is_playing_midi:
                self.stop_midi_playback()
            
            # 确保有事件数据
            if not self.current_events:
                self.update_event_data()
            
            if self.current_events:
                self.start_preview()
            else:
                messagebox.showinfo("提示", "请先选择MIDI文件并确保有有效事件数据")
        else:
            self.stop_preview()
    
    def start_preview(self):
        """开始预览"""
        try:
            # 禁用试听MIDI按钮
            if hasattr(self, 'midi_play_button'):
                self.midi_play_button.config(state=DISABLED)
            
            # 在单独的线程中播放预览
            self.is_previewing = True
            self.preview_button.config(text="停止预览")
            
            # 启动预览线程
            threading.Thread(target=self._preview_thread, daemon=True).start()
            
        except Exception as e:
            print(f"开始预览时出错: {str(e)}")
            self.is_previewing = False
            self.preview_button.config(text="预览")
            # 重新启用试听MIDI按钮
            if hasattr(self, 'midi_play_button'):
                self.midi_play_button.config(state=NORMAL)
    
    def _preview_thread(self):
        """预览线程：通过midi_preview模块生成临时MIDI并播放"""
        temp_midi_path = None
        try:
            # 确保有事件数据
            print(f"调试：检查current_events - 类型: {type(self.current_events)}, 长度: {len(self.current_events)}")
            
            if not self.current_events:
                print("调试：current_events为空")
                self.root.after(0, lambda: messagebox.showinfo("提示", "请先选择MIDI文件并确保有有效事件数据"))
                return
            
            # 导入并使用midi_preview_wrapper来生成临时MIDI文件
            from midi_preview_wrapper import get_preview_wrapper
            preview_wrapper = get_preview_wrapper()
            
            # 尝试获取BPM信息
            bpm = 120  # 默认值
            if hasattr(self, 'current_analysis_result') and self.current_analysis_result:
                bpm = self.current_analysis_result.get('bpm', 120)
            
            # 调用预览包装器生成临时MIDI文件
            temp_midi_path = preview_wrapper.generate_preview_midi(self.current_events, bpm)
            if not temp_midi_path:
                print("调试：生成预览MIDI文件失败")
                self.root.after(0, lambda: messagebox.showerror("错误", "生成预览MIDI文件失败"))
                return
            
            # 获取MIDI文件时长
            duration_seconds, duration_minutes, duration_seconds_remainder = preview_wrapper.preview_generator.get_midi_duration(temp_midi_path)
            print(f"[预览] MIDI文件时长: {duration_minutes}分{duration_seconds_remainder}秒 ({duration_seconds:.2f}秒)")
            
            # 播放临时MIDI文件
            preview_wrapper.play_preview(temp_midi_path)
            
            # 记录开始播放的时间
            start_time = time.time()
            
            # 等待播放完成，但要定期检查是否需要停止并更新剩余时间
            while preview_wrapper.is_playing() and self.is_previewing:
                # 计算已播放时间
                elapsed_time = time.time() - start_time
                # 计算剩余时间
                remaining_time = max(0, duration_seconds - elapsed_time)
                remaining_minutes = int(remaining_time // 60)
                remaining_seconds = int(remaining_time % 60)
                
                # 更新按钮上方的剩余时间标签
                remaining_time_str = f"{str(remaining_minutes).zfill(2)}:{str(remaining_seconds).zfill(2)}"
                self.root.after(0, lambda time_str=remaining_time_str: 
                               self.update_remaining_time_label(time_str))
                
                time.sleep(0.5)  # 每0.5秒更新一次剩余时间
            
            # 如果仍在播放但预览已取消，停止播放
            if preview_wrapper.is_playing():
                preview_wrapper.stop_playback()
                
        except Exception as e:
            error_msg = f"预览线程出错: {str(e)}"
            print(error_msg)
            self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
        finally:
            # 删除临时文件
            if temp_midi_path and os.path.exists(temp_midi_path):
                try:
                    os.remove(temp_midi_path)
                    print(f"已删除临时预览MIDI文件: {temp_midi_path}")
                except Exception as e:
                    print(f"删除临时文件时出错: {str(e)}")
            
            # 清理剩余时间显示
            self.root.after(0, lambda: self.update_remaining_time_label("00:00"))
                
            # 确保恢复状态
            if self.is_previewing:
                self.root.after(0, lambda: self.preview_button.config(text="预览"))
                self.is_previewing = False
    
    def stop_preview(self):
        """停止预览"""
        self.is_previewing = False
        # 确保在主线程中更新UI
        self.root.after(0, lambda: self.preview_button.config(text="预览"))
        self.root.after(0, lambda: self.update_remaining_time_label("00:00"))
        
        # 使用预览包装器清理资源和停止播放
        try:
            from midi_preview_wrapper import get_preview_wrapper
            preview_wrapper = get_preview_wrapper()
            # 首先停止播放
            preview_wrapper.stop_playback()
            # 然后清理所有资源
            preview_wrapper.cleanup()
            print("预览播放已停止并清理资源")
        except Exception as e:
            print(f"停止预览和清理资源时出错: {str(e)}")
        
        # 重新启用试听MIDI按钮
        if hasattr(self, 'current_file_path') and self.current_file_path and hasattr(self, 'midi_play_button'):
            self.root.after(0, lambda: self.midi_play_button.config(state=NORMAL))
            
    def update_remaining_time_label(self, time_str):
        """更新剩余时间标签"""
        # 使用操作区域顶部预留的时间标签来显示剩余时间
        self.time_label.config(text=f"剩余时间: {time_str}")
        
        # 清理可能存在的旧标签
        if hasattr(self, 'remaining_time_label'):
            try:
                self.remaining_time_label.destroy()
                delattr(self, 'remaining_time_label')
            except Exception as e:
                print(f"清理旧标签时出错: {str(e)}")
            
    def play_previous_song(self):
        """播放上一首歌曲"""
        if not self.midi_files:
            return
            
        try:
            # 找到当前文件在列表中的索引
            current_index = -1
            if hasattr(self, 'current_file_path') and self.current_file_path:
                for i, file_path in enumerate(self.midi_files):
                    if file_path == self.current_file_path:
                        current_index = i
                        break
            
            # 计算上一首的索引（循环播放）
            if current_index > 0:
                prev_index = current_index - 1
            else:
                prev_index = len(self.midi_files) - 1  # 最后一首
            
            # 加载上一首文件
            next_file = self.midi_files[prev_index]
            self._load_and_analyze_midi(next_file)
            
            # 如果正在播放，自动开始播放新文件
            if hasattr(self.midi_player, 'playing') and self.midi_player.playing:
                self.toggle_play()  # 先暂停当前播放
                self.toggle_play()  # 再开始播放新文件
                
        except Exception as e:
            print(f"播放上一首歌曲时出错: {str(e)}")
            messagebox.showerror("播放错误", f"播放上一首歌曲时出错: {str(e)}")
            
    def play_next_song(self):
        """播放下一首歌曲"""
        if not self.midi_files:
            return
            
        try:
            # 找到当前文件在列表中的索引
            current_index = -1
            if hasattr(self, 'current_file_path') and self.current_file_path:
                for i, file_path in enumerate(self.midi_files):
                    if file_path == self.current_file_path:
                        current_index = i
                        break
            
            # 计算下一首的索引（循环播放）
            next_index = (current_index + 1) % len(self.midi_files)
            
            # 加载下一首文件
            next_file = self.midi_files[next_index]
            self._load_and_analyze_midi(next_file)
            
            # 如果正在播放，自动开始播放新文件
            if hasattr(self.midi_player, 'playing') and self.midi_player.playing:
                self.toggle_play()  # 先暂停当前播放
                self.toggle_play()  # 再开始播放新文件
                
        except Exception as e:
            print(f"播放下一首歌曲时出错: {str(e)}")
            messagebox.showerror("播放错误", f"播放下一首歌曲时出错: {str(e)}")
    
    def update_progress(self):
        """更新进度显示"""
        try:
            # 检查midi_player是否存在并且正在播放，同时确保不在倒计时状态
            if self.midi_player and self.midi_player.playing:
                # 检查是否有counting_down属性并且不在倒计时状态
                if hasattr(self.midi_player, 'counting_down') and not self.midi_player.counting_down:
                    current_time = self.midi_player.get_current_time()
                    total_time = self.midi_player.get_total_time()
                    
                    if total_time > 0:
                        remaining_time = total_time - current_time
                        # 格式化为分:秒
                        minutes, seconds = divmod(int(remaining_time), 60)
                        time_str = f"剩余时间: {minutes:02d}:{seconds:02d}"
                        self.time_label.config(text=time_str)
            
            # 继续定时更新
            self.progress_timer = self.root.after(100, self.update_progress)
            
        except Exception as e:
            print(f"更新进度时出错: {str(e)}")
            # 尝试继续定时更新
            self.progress_timer = self.root.after(100, self.update_progress)
    
    def check_window_state(self):
        """检查窗口状态"""
        try:
            # 这里可以实现窗口状态检查逻辑
            # 例如检测目标窗口是否打开
            
            # 继续定时检查
            self.window_check_timer = self.root.after(200, self.check_window_state)
            
        except Exception as e:
            print(f"检查窗口状态时出错: {str(e)}")
            # 尝试继续定时检查
            self.window_check_timer = self.root.after(200, self.check_window_state)

def main():
    """主函数"""
    # 创建ttkbootstrap应用
    root = ttkb.Window(themename="cosmo")
    
    # 创建并运行主窗口
    app = MainWindow(root)
    
    # 运行主循环
    root.mainloop()

if __name__ == "__main__":
    main()
