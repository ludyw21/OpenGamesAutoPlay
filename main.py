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
        
        # 根据DPI缩放比例计算窗口大小
        base_width, base_height = 550, 480
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
        
        # 添加键盘事件防抖动
        self.last_key_time = 0
        self.key_cooldown = 0.2  # 200ms冷却时间
        
        # 初始化预览状态
        self.is_previewing = False
        
        # 初始化pygame mixer（如果还没有初始化）
        if not pygame.mixer.get_init():
            try:
                pygame.mixer.init(frequency=44100)
            except Exception as e:
                print(f"初始化音频系统失败: {str(e)}")
        
        # 初始化其他属性
        self.current_index = -1
        self.midi_files = []
        
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
            
            # 使用grid布局确保左右两栏比例固定
            main_frame.grid_columnconfigure(0, weight=1)  # 左侧占1份
            main_frame.grid_columnconfigure(1, weight=2)  # 右侧占2份
            
            # 创建左侧框架 - 固定占1/3宽度
            left_frame = ttk.LabelFrame(main_frame, text="文件管理", padding=10)
            left_frame.grid(row=0, column=0, sticky=NSEW, padx=5, pady=5)
            left_frame.grid_propagate(False)  # 防止内部组件改变框架大小
            left_frame.configure(width=260, height=400)  # 设置固定高度和宽度
            
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
            
            # 歌曲列表
            self.song_list = ttk.Treeview(left_frame, columns=["song"], show="headings", height=15)
            self.song_list.heading("song", text="歌曲列表")
            self.song_list.column("song", width=200)
            self.song_list.pack(fill=BOTH, expand=YES, pady=5)
            self.song_list.bind("<<TreeviewSelect>>", lambda e: self.song_selected())
            
            # 右侧框架 - 固定占2/3宽度
            right_frame = ttk.LabelFrame(main_frame, text="播放控制", padding=10)
            right_frame.grid(row=0, column=1, sticky=NSEW, padx=5, pady=5)
            right_frame.grid_propagate(False)  # 防止内部组件改变框架大小
            right_frame.configure(width=520, height=400)  # 设置固定宽度和高度（左侧260，右侧520，保持1:2比例）
            
            # 创建音轨详情区域
            tracks_label = ttk.Label(right_frame, text="音轨详情", font=('Arial', 12, 'bold'))
            tracks_label.pack(anchor=W, pady=5)
            
            # 当前歌曲名称标签
            self.current_song_label = ttk.Label(right_frame, text="当前歌曲：未选择")
            self.current_song_label.pack(anchor=W, pady=2)
            
            # 创建音轨列表 - 固定宽度和高度
            self.tracks_list = ttk.Treeview(right_frame, columns=["track"], show="headings", height=8)
            self.tracks_list.heading("track", text="音轨列表")
            self.tracks_list.column("track", width=480, stretch=False)  # 增加宽度，确保完整显示
            self.tracks_list.pack(fill=X, pady=5)
            self.tracks_list.bind("<<TreeviewSelect>>", lambda e: self.track_selected())
            
            # 时间显示
            self.time_label = ttk.Label(right_frame, text="剩余时间: 00:00", anchor=CENTER)
            self.time_label.pack(fill=X, pady=5)
            
            # 控制按钮布局
            control_frame = ttk.Frame(right_frame)
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
            
            # 添加使用说明
            info_frame = ttk.LabelFrame(right_frame, text="使用说明", padding=10)
            info_frame.pack(fill=X, pady=10)
            
            usage_text = "注意：工具支持36键模式!\n" + \
                         "使用说明：\n" + \
                         "1. 使用管理员权限启动\n" + \
                         "2. 选择MIDI文件\n" + \
                         "3. 选择要播放的音轨\n" + \
                         "4. 点击播放按钮开始演奏"
            
            usage_label = ttk.Label(info_frame, text=usage_text, justify=LEFT)
            usage_label.pack(fill=X)
            
            # 添加快捷键说明
            shortcut_frame = ttk.LabelFrame(right_frame, text="快捷键", padding=10)
            shortcut_frame.pack(fill=X, pady=5)
            
            shortcut_text = "快捷键说明：\n" + \
                            "Alt + 减号键(-) 播放/暂停\n" + \
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
        
        # 更新当前歌曲标签
        file_name = os.path.basename(file_path)
        self.current_song_label.config(text=f"当前歌曲：{file_name}")
        
        # 解析MIDI文件，获取音轨信息
        self._load_midi_tracks(file_path)
        
        # 启用预览按钮
        self.preview_button.config(state=NORMAL)
    
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
                
                # 添加到Treeview
                self.tracks_list.insert('', END, values=[display_name], tags=(i,))
                
                # 存储音轨信息
                self.tracks_info.append({"track_index": i, "note_count": note_count})
            
            # 绑定选择事件
            self.tracks_list.bind('<<TreeviewSelect>>', self.track_selected)
            
            # 保存MIDI文件路径
            self.current_file_path = file_path
            print(f"成功加载MIDI文件，共找到{len(self.tracks_info)}个有效音轨")
            
        except Exception as e:
            print(f"加载MIDI文件时出错: {str(e)}")
            messagebox.showerror("MIDI错误", f"加载MIDI文件时出错: {str(e)}")
    
    def track_selected(self, event=None):
        """选择音轨时的处理"""
        selected_items = self.tracks_list.selection()
        if not selected_items:
            return
        
        # 获取选中的音轨索引
        item = selected_items[0]
        track_index = int(self.tracks_list.item(item, "tags")[0])
        
        # 设置选中的音轨
        self.midi_player.selected_track = track_index
        
        # 启用播放和停止按钮
        self.play_button.config(state=NORMAL)
        self.stop_button.config(state=NORMAL)
    
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
    
    def toggle_preview(self):
        """切换预览状态"""
        if not self.is_previewing:
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
