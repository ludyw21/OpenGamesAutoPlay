import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import os
import ctypes
import csv
from datetime import datetime
# 导入音符分组信息
from groups import group_for_note

class EventTableDialog:
    def __init__(self, main_window):
        self.main_window = main_window
        self.parent = main_window.root  # 获取Tkinter root窗口
        self.dialog = ttkb.Toplevel(self.parent)
        self.dialog.title("事件表")
        self.dialog.transient(self.parent)  # 设置为主窗口的子窗口
        self.dialog.grab_set()  # 模态窗口
        
        # 获取DPI缩放比例
        dpi_scale = 1.0
        if hasattr(os, 'name') and os.name == 'nt':
            try:
                dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
                print(f"事件表窗口DPI缩放比例: {dpi_scale:.2f}")
            except Exception as e:
                print(f"事件表窗口获取DPI缩放比例失败: {str(e)}")
        
        # 根据DPI缩放比例计算窗口大小，位置
        base_width, base_height = 500, 400
        dialog_width = int(base_width * dpi_scale)
        dialog_height = int(base_height * dpi_scale)
        
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # 初始化变量
        self.out_of_range_count_var = tk.StringVar(value="超限音符数量：0")
        self.show_only_out_of_range_var = tk.BooleanVar(value=False)
        
        # 创建工具栏
        self.create_toolbar()
        
        # 创建事件表
        self.create_event_table()
        
        # 初始化事件数据
        self.populate_event_table()
    
    def create_toolbar(self):
        """创建工具栏"""
        evt_toolbar = ttk.Frame(self.dialog)
        evt_toolbar.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 5))
        
        # 导出按钮
        ttk.Button(evt_toolbar, text="导出事件CSV", command=self.export_event_csv).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(evt_toolbar, text="导出按键谱", command=self.export_key_notation).pack(side=tk.LEFT, padx=(0, 8))
        
        # 超限音符显示开关和数量提示
        switch_frame = ttk.Frame(evt_toolbar)
        switch_frame.pack(side=tk.LEFT, padx=(20, 8))
        
        # 创建标签
        ttk.Label(switch_frame, textvariable=self.out_of_range_count_var).pack(side=tk.LEFT, padx=(0, 8))
        
        # 创建开关按钮
        switch_btn = ttk.Checkbutton(
            switch_frame, 
            text="仅显示超限音符", 
            variable=self.show_only_out_of_range_var, 
            command=self.toggle_display
        )
        switch_btn.pack(side=tk.LEFT)
    
    def create_event_table(self):
        """创建事件表格"""
        evt_top = ttk.Frame(self.dialog)
        evt_top.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # 定义列名和宽度
        columns = ("#", "time", "type", "note", "channel", "group", "end", "dur")
        self.tree = ttk.Treeview(evt_top, columns=columns, show='headings', height=25)
        
        # 设置列标题和宽度
        headers = ["序号", "时间", "事件", "音符", "通道", "分组", "结束", "时长"]
        widths = [100, 110, 110, 100, 100, 200, 100, 100]
        
        for i, col in enumerate(columns):
            self.tree.heading(col, text=headers[i])
            self.tree.column(col, width=widths[i], minwidth=0, stretch=False, anchor=tk.CENTER)
        
        # 添加垂直滚动条
        vbar = ttk.Scrollbar(evt_top, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vbar.set)
        
        # 布局
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定双击事件
        self.tree.bind('<Double-1>', self.on_event_double_click)
    
    def populate_event_table(self):
        """填充事件表数据"""
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 获取当前事件数据
        events = self.get_current_events()
        
        # 计算超限音符数量
        out_of_range_count = len([e for e in events if self.is_out_of_range(e)])
        self.out_of_range_count_var.set(f"超限音符数量：{out_of_range_count}")
        
        # 根据显示设置筛选事件
        if self.show_only_out_of_range_var.get():
            events = [e for e in events if self.is_out_of_range(e)]
        
        # 显示所有事件数据，不再限制数量
        # 如果事件数量非常大，可以考虑添加配置选项来控制显示数量
        
        # 填充表格
        for i, event in enumerate(events):
            values = [
                i + 1,
                f"{event['time']:.2f}",
                event['type'],
                self._format_note_display(event['note']),  # 格式化音符显示
                event['channel'],
                self._get_note_group(event['note']),  # 获取正确的分组名称
                f"{event['end']:.2f}" if 'end' in event else "-",
                f"{event['duration']:.2f}" if 'duration' in event else "-"
            ]
            self.tree.insert('', 'end', values=values)
    
    def get_current_events(self):
        """获取当前事件数据 - 从MainWindow实例获取已解析好的事件数据"""
        # 从MainWindow实例获取事件数据
        if hasattr(self.main_window, 'current_events') and self.main_window.current_events:
            return self.main_window.current_events
        
        # 如果没有实际数据，返回示例数据
        sample_events = []
        # 生成示例事件，使用正确的分组
        notes = [60, 64, 67, 62, 65, 69]
        for note in notes:
            # 添加note_on事件
            sample_events.append({
                'time': 0.0 if note in [60, 64, 67] else 0.5,
                'type': 'note_on',
                'note': note,
                'channel': 0,
                'group': self._get_note_group(note),
                'duration': 0.5,
                'end': 0.5 if note in [60, 64, 67] else 1.0
            })
            # 添加note_off事件
            sample_events.append({
                'time': 0.5,
                'type': 'note_off',
                'note': note,
                'channel': 0,
                'group': self._get_note_group(note)
            })
        
        return sample_events
    
    def _get_note_group(self, note_number):
        """获取音符所属的分组名称"""
        # 使用 groups.py 中的 group_for_note 函数获取分组名称
        try:
            from groups import group_for_note
            return group_for_note(note_number)
        except (ImportError, AttributeError):
            # 如果导入失败，返回未知分组
            return "未知"
    
    def _format_note_display(self, note_number):
        """将音符数字格式化为 音符名称(音符数字) 格式"""
        # 使用 groups.py 中的 get_note_name 函数获取音符名称
        try:
            from groups import get_note_name
            note_name = get_note_name(note_number)
            # 移除之前可能添加的数字，确保只显示音符名称部分
            # 提取音符名称部分（移除括号和数字）
            if '(' in note_name and ')' in note_name:
                note_name = note_name.split('(')[0].strip()
            return f"{note_name}({note_number})"
        except (ImportError, AttributeError):
            # 如果导入失败，只返回音符数字
            return f"{note_number}"
    
    def is_out_of_range(self, event):
        """判断音符是否超限"""
        # 这里应该根据应用程序的规则判断
        # 暂时返回False
        return False
    
    def toggle_display(self):
        """切换显示模式"""
        self.populate_event_table()
    
    def on_event_double_click(self, event):
        """双击事件处理"""
        # 这里可以实现双击事件的处理逻辑
        pass
    
    def export_event_csv(self):
        """导出事件CSV"""
        try:
            # 获取当前时间作为文件名一部分
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"event_table_{timestamp}.csv"
            
            # 写入CSV文件
            with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # 写入表头
                writer.writerow(["序号", "时间", "事件", "音符", "通道", "分组", "结束", "时长"])
                
                # 获取所有事件
                events = self.get_current_events()
                if self.show_only_out_of_range_var.get():
                    events = [e for e in events if self.is_out_of_range(e)]
                
                # 写入数据
                for i, event in enumerate(events):
                    writer.writerow([
                        i + 1,
                        f"{event['time']:.2f}",
                        event['type'],
                        event['note'],
                        event['channel'],
                        event['group'],
                        f"{event['end']:.2f}" if 'end' in event else "-",
                        f"{event['duration']:.2f}" if 'duration' in event else "-"
                    ])
            
            # 显示成功消息
            tk.messagebox.showinfo("导出成功", f"事件表已成功导出到\n{filename}")
        except Exception as e:
            tk.messagebox.showerror("导出失败", f"导出事件表时出错：\n{str(e)}")
    
    def export_key_notation(self):
        """导出按键谱"""
        # 这个功能可以根据需要实现
        tk.messagebox.showinfo("提示", "导出按键谱功能暂未实现")
