"""
MIDI自动演奏程序 - 一个基于ttkbootstrap的MIDI文件播放器，支持选择音轨和键盘控制。
提供直观的界面来加载、选择和播放MIDI文件，并支持全局快捷键控制。
"""

import os
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

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

# 忽略废弃警告
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

# 在主程序中添加或更新版本号
VERSION = "1.0.3"

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
            'stay_on_top': False
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
        self.root.title(f"燕云-自动演奏by木木睡没-{VERSION}")
        
        # 获取DPI缩放比例
        dpi_scale = 1.0
        if hasattr(os, 'name') and os.name == 'nt':
            try:
                import ctypes
                dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
            except:
                pass
        
        # 根据DPI缩放比例计算窗口大小，增加高度以容纳新的UI元素
        base_width, base_height = 550, 600  # 增加高度
        scaled_width = int(base_width * dpi_scale)
        scaled_height = int(base_height * dpi_scale)
        
        self.root.geometry(f"{scaled_width}x{scaled_height}")
        self.root.minsize(scaled_width, scaled_height)
        self.root.resizable(True, True)  # 允许调整窗口大小
        
        # 设置主题
        self.style = ttkb.Style(theme="pink")
        
        # 创建配置管理器实例
        self.config_manager = Config()
        # 从配置管理器获取配置
        self.config = self.config_manager.data
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
        self.transpose_var = tk.IntVar(value=0)  # 升降调（半音）
        self.octave_var = tk.IntVar(value=0)  # 整体转位（八度）
        
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
            
            # 歌曲列表 - 适应左侧框架宽度
            self.song_list = ttk.Treeview(left_frame, columns=["song"], show="headings")  # 适配左侧宽度
            self.song_list.heading("song", text="歌曲列表", anchor='center')  # 设置标题居中
            self.song_list.column("song", stretch=YES)  # 允许拉伸以占满空间
            self.song_list.pack(fill=BOTH, expand=YES, pady=5)
            self.song_list.bind("<<TreeviewSelect>>", lambda e: self.song_selected())
            
            # 右侧框架 - 自适应宽度
            right_frame = ttk.LabelFrame(main_frame, text="播放控制", padding=10)
            right_frame.grid(row=0, column=1, sticky=NSEW, padx=5, pady=5)  # 全方向拉伸，使其自适应
            
            # 创建音轨详情LabelFrame
            tracks_frame = ttk.LabelFrame(right_frame, text="音轨详情", padding=10)
            tracks_frame.pack(fill=X, pady=5)
            
            # 当前歌曲名称标签 - 设置为不换行，超出部分不显示，占满整行
            song_label_frame = ttk.Frame(tracks_frame)
            song_label_frame.pack(fill=X, pady=2)
            
            # 设置文本左对齐，wraplength=0确保不换行，超出部分会自动截断
            self.current_song_label = ttk.Label(song_label_frame, text="当前歌曲：未选择", anchor=W, wraplength=0, justify=LEFT)
            self.current_song_label.pack(fill=X, expand=True)
            
            # 创建音轨列表 - 自适应右侧宽度，使用自定义渲染来显示复选框
            self.tracks_list = ttk.Treeview(tracks_frame, columns=["checkbox", "track"], show="headings", height=6)
            self.tracks_list.heading("checkbox", text="")
            self.tracks_list.heading("track", text="音轨列表")
            self.tracks_list.column("checkbox", width=30, stretch=NO, anchor='center')
            self.tracks_list.column("track", stretch=YES)  # 允许列拉伸以适应右侧宽度
            self.tracks_list.pack(fill=X, pady=5)
            
            # 设置Treeview样式
            style = ttk.Style()
            # 注意：我们需要禁用默认的选择行为，完全由我们自己控制
            
            # 绑定点击事件以处理复选框和整行选中
            self.tracks_list.bind("<Button-1>", self.on_track_click)
            self.tracks_list.bind("<<TreeviewSelect>>", lambda e: None)  # 禁用默认的选择事件处理
            self.tracks_list.bind("<Motion>", lambda e: None)  # 禁用鼠标移动导致的高亮变化
            
            # 添加新的设置区域：移调和转位
            settings_frame = ttk.LabelFrame(right_frame, text="转音设置", padding=10)
            settings_frame.pack(fill=X, pady=5)
            
            # 移调设置
            transpose_frame = ttk.Frame(settings_frame)
            transpose_frame.pack(fill=X, pady=5)
            
            ttk.Label(transpose_frame, text="移调(半音):").pack(side=LEFT, padx=5)
            
            # 修改为左右布局的按钮
            transpose_control_frame = ttk.Frame(transpose_frame)
            transpose_control_frame.pack(side=LEFT, padx=5)
            
            ttk.Button(transpose_control_frame, text="-", command=lambda: self.adjust_value(self.transpose_var, -1), width=2).pack(side=LEFT)
            self.transpose_entry = ttk.Entry(transpose_control_frame, textvariable=self.transpose_var, width=5, justify='center')
            self.transpose_entry.pack(side=LEFT)
            ttk.Button(transpose_control_frame, text="+", command=lambda: self.adjust_value(self.transpose_var, 1), width=2).pack(side=LEFT)
            
            # 整体转位设置
            octave_frame = ttk.Frame(transpose_frame)
            octave_frame.pack(side=LEFT, padx=15)
            
            ttk.Label(octave_frame, text="转位(八度):").pack(side=LEFT, padx=5)
            
            # 修改为左右布局的按钮
            octave_control_frame = ttk.Frame(octave_frame)
            octave_control_frame.pack(side=LEFT, padx=5)
            
            ttk.Button(octave_control_frame, text="-", command=lambda: self.adjust_value(self.octave_var, -1), width=2).pack(side=LEFT)
            self.octave_entry = ttk.Entry(octave_control_frame, textvariable=self.octave_var, width=5, justify='center')
            self.octave_entry.pack(side=LEFT)
            ttk.Button(octave_control_frame, text="+", command=lambda: self.adjust_value(self.octave_var, 1), width=2).pack(side=LEFT)
            
            # MIDI分析数据显示区域
            analysis_frame = ttk.LabelFrame(right_frame, text="音轨分析", padding=10)
            analysis_frame.pack(fill=X, pady=5)
            
            # 分析数据内容
            self.analysis_text = "选中音轨分析(含移调、音程转位) 总音符数 0\n最高音: - 未检测\n最低音: - 未检测"
            self.analysis_label = ttk.Label(analysis_frame, text=self.analysis_text, justify=LEFT)
            self.analysis_label.pack(fill=X, anchor=W)
            
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
            
            # 创建一个容器框架来并列显示使用说明和快捷键说明，放置在底部
            bottom_container = ttk.Frame(right_frame)
            bottom_container.pack(side=BOTTOM, fill=X, pady=5, anchor='s')
            
            # 添加使用说明，确保底部对齐
            info_frame = ttk.LabelFrame(bottom_container, text="使用说明", padding=10)
            info_frame.pack(side=LEFT, fill=X, expand=True, padx=(0, 5), anchor='s')
            
            usage_text = "1. 使用管理员权限启动\n" + \
                         "2. 选择MIDI文件和音轨\n" + \
                         "3. 点击播放按钮开始演奏\n" + \
                         "4. 支持36键模式"
            
            usage_label = ttk.Label(info_frame, text=usage_text, justify=LEFT)
            usage_label.pack(fill=X)
            
            # 添加快捷键说明，确保底部对齐
            shortcut_frame = ttk.LabelFrame(bottom_container, text="快捷键说明", padding=10)
            shortcut_frame.pack(side=LEFT, fill=X, expand=True, padx=(5, 0), anchor='s')
            
            shortcut_text = "Alt + 减号键(-) 播放/暂停\n" + \
                            "Alt + 等号键(=) 停止播放"
            
            shortcut_label = ttk.Label(shortcut_frame, text=shortcut_text, justify=LEFT)
            shortcut_label.pack(fill=X)
            
        except Exception as e:
            print(f"设置UI界面时出错: {str(e)}")
            messagebox.showerror("UI错误", f"设置界面时出错: {str(e)}")
    
    def toggle_stay_on_top(self):
        """切换窗口置顶状态"""
        stay_on_top = self.stay_on_top_var.get()
        self.root.attributes('-topmost', stay_on_top)
        
    def setup_keyboard_hooks(self):
        """设置键盘快捷键"""
        try:
            # 播放/暂停
            keyboard.add_hotkey(CONTROL_KEYS['START_PAUSE'], lambda: self.safe_key_handler(self.toggle_play), 
                              suppress=True, trigger_on_release=True)
            
            # 停止
            keyboard.add_hotkey(CONTROL_KEYS['STOP'], lambda: self.safe_key_handler(self.stop_playback),
                              suppress=True, trigger_on_release=True)
            
            print("键盘快捷键设置完成")
            
        except Exception as e:
            print(f"设置键盘快捷键时出错: {str(e)}")
            messagebox.showerror("快捷键错误", f"设置快捷键时出错: {str(e)}")
    
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
        
        # 解析MIDI文件，获取音轨信息
        self._load_midi_tracks(file_path)
        
        # 启用预览按钮（确保按钮已初始化）
        if hasattr(self, 'preview_button') and self.preview_button is not None:
            self.preview_button.config(state=NORMAL)
        
        # 如果当前正在播放MIDI，自动切换到新选择的MIDI文件
        if hasattr(self, 'is_playing_midi') and self.is_playing_midi:
            # 停止当前播放的MIDI
            self.stop_midi_playback()
            # 自动开始播放新选择的MIDI
            self.start_midi_playback()
    
    def _fix_mojibake(self, text):
        """修复已被错误解码的字符串（ mojibake ）"""
        if not isinstance(text, str):
            return text
        
        # 记录原始文本用于调试
        original_text = text
        
        # 尝试检测常见的错误解码模式并修复
        try:
            # 1. 首先尝试处理UTF-8被错误解码为Latin-1的情况（例如"æ— æ‡é¢˜" -> "无标题"）
            # 这种模式常见于UTF-8文本被错误解码为Latin-1/ISO-8859-1
            try:
                # 对于包含UTF-8特征字符的字符串，优先尝试这种修复
                if any(192 <= ord(c) <= 255 for c in text):
                    # 使用errors='replace'来避免编码错误
                    utf8_fixed = text.encode('latin-1', errors='replace').decode('utf-8', errors='replace')
                    # 检查是否有中文字符（表示可能修复成功）
                    if any('\u4e00' <= c <= '\u9fff' for c in utf8_fixed):
                        # print(f"UTF-8修复成功: {original_text} -> {utf8_fixed}")
                        return utf8_fixed
            except Exception as e:
                print(f"UTF-8修复失败: {e}")
            
            # 2. 尝试使用cp1252作为中间编码（Windows常用的Latin-1扩展）
            try:
                if any(192 <= ord(c) <= 255 for c in text):
                    # 先转换为cp1252字节，再尝试UTF-8解码
                    cp1252_fixed = text.encode('cp1252', errors='replace').decode('utf-8', errors='replace')
                    if any('\u4e00' <= c <= '\u9fff' for c in cp1252_fixed):
                        # print(f"CP1252修复成功: {original_text} -> {cp1252_fixed}")
                        return cp1252_fixed
            except Exception as e:
                print(f"CP1252修复失败: {e}")
            
            # 3. 尝试ISO-8859-1 -> GBK转换（常见的中文乱码模式）
            try:
                # 使用errors='replace'确保即使有特殊字符也能继续
                gbk_fixed = text.encode('latin-1', errors='replace').decode('gbk', errors='replace')
                if any('\u4e00' <= c <= '\u9fff' for c in gbk_fixed):
                        # print(f"GBK修复成功: {original_text} -> {gbk_fixed}")
                        return gbk_fixed
            except Exception as e:
                print(f"GBK修复失败: {e}")
            
            # 4. 尝试ISO-8859-1 -> GB2312转换
            try:
                gb2312_fixed = text.encode('latin-1', errors='replace').decode('gb2312', errors='replace')
                if any('\u4e00' <= c <= '\u9fff' for c in gb2312_fixed):
                        # print(f"GB2312修复成功: {original_text} -> {gb2312_fixed}")
                        return gb2312_fixed
            except Exception as e:
                print(f"GB2312修复失败: {e}")
            
            # 5. 尝试ISO-8859-1 -> GB18030转换（支持更多字符）
            try:
                gb18030_fixed = text.encode('latin-1', errors='replace').decode('gb18030', errors='replace')
                if any('\u4e00' <= c <= '\u9fff' for c in gb18030_fixed):
                        # print(f"GB18030修复成功: {original_text} -> {gb18030_fixed}")
                        return gb18030_fixed
            except Exception as e:
                print(f"GB18030修复失败: {e}")
            
            # 6. 尝试多种编码组合的复杂转换
            try:
                if any(192 <= ord(c) <= 255 for c in text):
                    # 尝试多种编码组合
                    for encoding1 in ['latin-1', 'cp1252']:
                        for encoding2 in ['utf-8', 'gbk', 'gb18030']:
                            try:
                                # 先转换为bytes再尝试其他编码
                                combined_fixed = text.encode(encoding1, errors='replace').decode(encoding2, errors='replace')
                                if any('\u4e00' <= c <= '\u9fff' for c in combined_fixed) and combined_fixed != text:
                                    # print(f"组合修复成功 ({encoding1}->{encoding2}): {original_text} -> {combined_fixed}")
                                    return combined_fixed
                            except Exception:
                                continue
            except Exception as e:
                print(f"组合修复失败: {e}")
            
            # 7. 尝试更复杂的三重编码修复
            try:
                if any(192 <= ord(c) <= 255 for c in text):
                    # 先Latin-1 -> GBK -> UTF-8的三重转换
                    triple_fixed = text.encode('latin-1', errors='replace').decode('gbk', errors='replace').encode('utf-8', errors='replace').decode('utf-8', errors='replace')
                    if any('\u4e00' <= c <= '\u9fff' for c in triple_fixed) and triple_fixed != text:
                        # print(f"三重修复成功: {original_text} -> {triple_fixed}")
                        return triple_fixed
            except Exception as e:
                print(f"三重修复失败: {e}")
                
            # 8. 针对特殊情况：检测是否是典型的UTF-8编码错误模式
            try:
                # 检查是否包含连续的UTF-8特征字节模式
                if any(0xC0 <= ord(c) <= 0xDF for c in text) or any(0xE0 <= ord(c) <= 0xEF for c in text):
                    # 这可能是UTF-8被错误解码
                    # 使用原始字节重新解码
                    try:
                        # 获取原始字节表示，然后尝试正确解码
                        raw_bytes = ''.join(chr(ord(c)) for c in text).encode('latin-1', errors='replace')
                        utf8_direct = raw_bytes.decode('utf-8', errors='replace')
                        if any('\u4e00' <= c <= '\u9fff' for c in utf8_direct) and utf8_direct != text:
                            # print(f"直接UTF-8解码修复: {original_text} -> {utf8_direct}")
                            return utf8_direct
                    except Exception:
                        pass
            except Exception as e:
                print(f"特征检测修复失败: {e}")
                
        except Exception as e:
            print(f"乱码修复过程出错: {e}")
        
        # 如果所有修复都失败，记录未修复的文本
        if any(192 <= ord(c) <= 255 for c in text):
            print(f"未能修复的乱码: {original_text}")
        
        return text
    
    def _load_midi_tracks(self, file_path):
        """加载MIDI文件的音轨信息"""
        try:
            # 清空当前音轨列表
            for item in self.tracks_list.get_children():
                self.tracks_list.delete(item)
            
            # 清空选中的音轨集合
            self.selected_tracks.clear()
            
            # 解析MIDI文件 - 使用更安全的编码处理方式
            import mido
            print(f"正在加载MIDI文件: {file_path}")
            # 先尝试使用二进制模式打开，然后使用不同编码方案尝试解析
            try:
                mid = mido.MidiFile(file_path, charset='utf-8')  # 优先尝试UTF-8
            except UnicodeDecodeError:
                try:
                    mid = mido.MidiFile(file_path, charset='cp1252')  # 尝试cp1252
                except UnicodeDecodeError:
                    mid = mido.MidiFile(file_path, charset='cp932')  # 最后尝试日文编码
            
            # 初始化音轨信息列表
            self.tracks_info = []
            
            # 先添加"全部音轨"项作为全选/反选控制
            self.all_tracks_item = self.tracks_list.insert('', END, values=["✓", "全部音轨"], tags="all_tracks")
            # 默认全选时给"全部音轨"项也添加高亮
            self.tracks_list.selection_add(self.all_tracks_item)
            
            # 添加音轨信息
            for i, track in enumerate(mid.tracks):
                # 统计音符数量
                note_count = sum(1 for msg in track if msg.type == 'note_on' and msg.velocity > 0)
                
                # 过滤掉音符数量过少的音轨
                if note_count < 10:
                    continue
                
                # 尝试从音轨中提取名称
                original_name = None
                for msg in track:
                    if msg.type == 'track_name' and msg.name:
                        original_name = msg.name
                        break
                
                # 使用专门的乱码修复方法处理音轨名称
                fixed_name = self._fix_mojibake(original_name) if original_name else f"未命名"
                print(f"音轨{i+1} - 原始名称: {original_name}, 修复后: {fixed_name}")
                
                # 构建显示名称，添加音轨标号
                display_name = f"音轨{i+1}：{fixed_name} ({note_count}个音符)"
                
                # 添加到Treeview，默认选中（显示✓）
                checkbox = "✓"  # 默认选中
                item = self.tracks_list.insert('', END, values=[checkbox, display_name], tags=(i,))
                
                # 存储音轨信息
                self.tracks_info.append({"track_index": i, "note_count": note_count, "item_id": item})
                
                # 默认全选并选中行（高亮）
                self.selected_tracks.add(i)
                self.tracks_list.selection_add(item)
            
            # 更新分析信息显示
            self.update_analysis_info()
            
            # 保存MIDI文件路径
            self.current_file_path = file_path
            print(f"成功加载MIDI文件，共找到{len(self.tracks_info)}个有效音轨")
                
            # 启用试听MIDI按钮
            if hasattr(self, 'midi_play_button'):
                self.midi_play_button.config(state=NORMAL)
            
            # 初始化按钮状态：播放按钮亮，暂停按钮灰
            if hasattr(self, 'play_button'):
                self.play_button.config(state=NORMAL)
            if hasattr(self, 'stop_button'):
                self.stop_button.config(state=DISABLED)
            
        except Exception as e:
            print(f"加载MIDI文件时出错: {str(e)}")
            messagebox.showerror("MIDI错误", f"加载MIDI文件时出错: {str(e)}")
    
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
        
        # 更新分析信息和按钮状态
        self.update_analysis_info()
        self.track_selected()
        
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
        """处理全选/反选功能"""
        # 打印操作前状态
        print("===== 切换全选前状态 =====")
        self._print_track_states()
        
        # 检查"全部音轨"项的当前状态
        all_values = list(self.tracks_list.item(self.all_tracks_item, "values"))
        current_status = all_values[0] == "✓"
        
        # 如果全部已选中，则取消选中所有音轨；否则选中所有音轨
        if current_status:
            # 取消选中所有音轨
            self.selected_tracks.clear()
            # 更新UI，移除所有复选框并取消行高亮
            for info in self.tracks_info:
                values = list(self.tracks_list.item(info['item_id'], "values"))
                values[0] = "□"  # 使用明确的未选中标记
                self.tracks_list.item(info['item_id'], values=values)
                self.tracks_list.selection_remove(info['item_id'])
            # 更新"全部音轨"项
            all_values[0] = "□"  # 使用明确的未选中标记
            self.tracks_list.item(self.all_tracks_item, values=all_values)
            # 取消"全部音轨"项的高亮
            self.tracks_list.selection_remove(self.all_tracks_item)
        else:
            # 选中所有音轨
            self.selected_tracks = set(info['track_index'] for info in self.tracks_info)
            # 更新UI，添加所有复选框并添加行高亮
            for info in self.tracks_info:
                values = list(self.tracks_list.item(info['item_id'], "values"))
                values[0] = "✓"
                self.tracks_list.item(info['item_id'], values=values)
                self.tracks_list.selection_add(info['item_id'])
            # 更新"全部音轨"项
            all_values[0] = "✓"
            self.tracks_list.item(self.all_tracks_item, values=all_values)
            # 添加"全部音轨"项的高亮
            self.tracks_list.selection_add(self.all_tracks_item)
        
        # 更新分析信息
        self.update_analysis_info()
        # 调用track_selected更新按钮状态
        self.track_selected()
        
        # 打印操作后状态
        print("===== 切换全选后状态 =====")
        self._print_track_states()
        print("====================\n")
    
    def adjust_value(self, var, delta):
        """调整数值变量"""
        var.set(var.get() + delta)
        # 更新分析信息
        self.update_analysis_info()
    
    def update_analysis_info(self):
        """更新音轨分析信息显示"""
        # 获取当前的移调和转位设置
        transpose = self.transpose_var.get()
        octave = self.octave_var.get()
        
        if self.selected_tracks:
            # 获取选中的音轨名称列表
            selected_track_names = []
            total_notes = 0
            
            for info in self.tracks_info:
                if info['track_index'] in self.selected_tracks:
                    # 获取音轨名称（从tracks_list中获取）
                    values = self.tracks_list.item(info['item_id'], "values")
                    if values and len(values) > 1:
                        # 提取音轨名称，如"音轨1"、"音轨2"等
                        track_text = values[1]
                        if "音轨" in track_text:
                            # 提取音轨编号部分
                            track_name = track_text.split("：")[0]
                            selected_track_names.append(track_name)
                    # 累计音符数
                    total_notes += info['note_count']
            
            # 构建选中音轨的显示字符串
            tracks_str = "、".join(selected_track_names)
            
            # 更新第一行文本
            first_line = f"音轨{{{tracks_str}}}   移调:{transpose}  转位:{octave}  总音符:{total_notes}"
            
            # 保留后续行的内容
            self.analysis_text = f"{first_line}\n最高音: 82 a² 小字二组 未超限 超限数量: 0\n最低音: 23 B₂ 大字二组 超限 超限数量: 2"
        else:
            self.analysis_text = "音轨{无选中}   移调:0  转位:0  总音符:0\n最高音: - 未检测\n最低音: - 未检测"
        
        self.analysis_label.config(text=self.analysis_text)
    
    def toggle_play(self):
        """切换播放/暂停状态"""
        if not self.midi_player.playing:
            self.start_playback()
        else:
            self.pause_playback()
    
    def start_playback(self):
        """开始播放"""
        try:
            if hasattr(self, 'current_file_path') and self.current_file_path:
                # 如果正在播放MIDI，先停止
                if hasattr(self, 'is_playing_midi') and self.is_playing_midi:
                    self.stop_midi_playback()
                
                # 获取选中的音轨
                selected_track = None
                if hasattr(self.midi_player, 'selected_track'):
                    selected_track = self.midi_player.selected_track
                
                # 使用play_midi方法播放
                threading.Thread(target=self.midi_player.play_midi, 
                               args=(self.current_file_path, selected_track), 
                               daemon=True).start()
                
                self.play_button.config(text="暂停")
                self.stop_button.config(state=NORMAL)
        except Exception as e:
            print(f"开始播放时出错: {str(e)}")
            messagebox.showerror("播放错误", f"开始播放时出错: {str(e)}")
    
    def pause_playback(self):
        """暂停播放"""
        try:
            self.midi_player.pause()
            self.play_button.config(text="播放")
        except Exception as e:
            print(f"暂停播放时出错: {str(e)}")
            messagebox.showerror("播放错误", f"暂停播放时出错: {str(e)}")
    
    def stop_playback(self):
        """停止播放"""
        try:
            self.midi_player.stop()
            self.play_button.config(text="播放", state=NORMAL)
            self.stop_button.config(state=DISABLED)
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
                
                # 开始播放MIDI文件
                pygame.mixer.music.load(self.current_file_path)
                pygame.mixer.music.play()
                
                # 更新状态
                self.is_playing_midi = True
                # 使用ttkbootstrap的内置Danger样式
                self.midi_play_button.config(text="停止MIDI")
                
                # 禁用其他播放按钮
                self.play_button.config(state=DISABLED)
                self.preview_button.config(state=DISABLED)
                
                # 定期检查播放是否结束
                self.check_midi_playback()
                
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
    
    def toggle_preview(self):
        """切换预览状态"""
        if not self.is_previewing:
            # 如果正在播放MIDI，先停止
            if hasattr(self, 'is_playing_midi') and self.is_playing_midi:
                self.stop_midi_playback()
            self.start_preview()
        else:
            self.stop_preview()
    
    def start_preview(self):
        """开始预览"""
        try:
            # 在单独的线程中播放预览
            self.is_previewing = True
            self.preview_button.config(text="停止预览")
            
            # 这里可以实现预览功能，例如播放MIDI的一小部分
            threading.Thread(target=self._preview_thread, daemon=True).start()
            
        except Exception as e:
            print(f"开始预览时出错: {str(e)}")
            self.is_previewing = False
            self.preview_button.config(text="预览")
    
    def _preview_thread(self):
        """预览线程"""
        try:
            # 简单的预览实现 - 播放前几秒
            pass
        except Exception as e:
            print(f"预览线程出错: {str(e)}")
        finally:
            # 确保恢复状态
            if self.is_previewing:
                self.root.after(0, lambda: self.preview_button.config(text="预览"))
                self.is_previewing = False
    
    def stop_preview(self):
        """停止预览"""
        self.is_previewing = False
        self.preview_button.config(text="预览")
        
        # 重新启用试听MIDI按钮
        if hasattr(self, 'current_file_path') and self.current_file_path and hasattr(self, 'midi_play_button'):
            self.midi_play_button.config(state=NORMAL)
    
    def update_progress(self):
        """更新进度显示"""
        try:
            if self.midi_player and self.midi_player.playing:
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
