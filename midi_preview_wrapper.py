"""
MIDI预览功能包装器，用于将main.py中的预览功能委托给midi_preview模块
"""
import os
import time
from midi_preview import get_midi_preview

class MidiPreviewWrapper:
    def __init__(self):
        self.preview_generator = get_midi_preview()
        self.current_temp_path = None
        # 初始化pygame mixer状态
        self._pygame_initialized = False
    
    def generate_preview_midi(self, events, bpm=120):
        """生成预览MIDI文件
        
        Args:
            events: MIDI事件列表
            bpm: 节拍速度
            
        Returns:
            str: 生成的临时MIDI文件路径，如果失败则返回None
        """
        try:
            temp_midi_path = self.preview_generator.generate_preview_midi(events, bpm)
            self.current_temp_path = temp_midi_path
            return temp_midi_path
        except Exception as e:
            print(f"生成预览MIDI文件时出错: {str(e)}")
            return None
    
    def play_preview(self, midi_path):
        """播放预览MIDI文件
        
        Args:
            midi_path: MIDI文件路径
        """
        try:
            import pygame.mixer
            # 确保pygame mixer已初始化
            if not pygame.mixer.get_init():
                pygame.mixer.init(frequency=44100)
                self._pygame_initialized = True
            
            # 停止当前可能正在播放的音乐
            if pygame.mixer.music.get_busy():
                pygame.mixer.music.stop()
            
            # 加载并播放MIDI文件
            pygame.mixer.music.load(midi_path)
            pygame.mixer.music.play()
            print(f"开始播放预览MIDI文件: {midi_path}")
        except Exception as e:
            print(f"播放预览时出错: {str(e)}")
    
    def is_playing(self):
        """检查是否正在播放
        
        Returns:
            bool: 如果正在播放则返回True，否则返回False
        """
        try:
            import pygame.mixer
            if pygame.mixer.get_init():
                return pygame.mixer.music.get_busy()
            return False
        except Exception as e:
            print(f"检查播放状态时出错: {str(e)}")
            return False
    
    def stop_playback(self):
        """停止播放
        """
        try:
            import pygame.mixer
            if pygame.mixer.get_init():
                pygame.mixer.music.stop()
                print("已停止播放预览")
        except Exception as e:
            print(f"停止播放时出错: {str(e)}")
    
    def generate_and_play_preview(self, current_events, root, preview_button, update_remaining_time_func):
        """生成并播放预览MIDI文件，返回临时文件路径和时长信息"""
        if not current_events:
            print("[预览包装器] current_events为空")
            root.after(0, lambda: messagebox.showinfo("提示", "请先选择MIDI文件并确保有有效事件数据"))
            return None, (0, 0, 0)
        
        print(f"[预览包装器] 开始处理，事件总数: {len(current_events)}")
        
        try:
            # 使用midi_preview模块生成临时MIDI文件
            temp_midi_path = self.generate_preview_midi(current_events)
            
            if not temp_midi_path:
                root.after(0, lambda: messagebox.showerror("错误", "生成临时MIDI文件失败！"))
                return None, (0, 0, 0)
            
            # 获取MIDI文件时长
            duration_seconds, duration_minutes, duration_seconds_remainder = self.preview_generator.get_midi_duration(temp_midi_path)
            print(f"[预览包装器] MIDI文件时长: {duration_minutes}分{duration_seconds_remainder}秒 ({duration_seconds:.2f}秒)")
            
            # 播放预览
            self.play_preview(temp_midi_path)
            
            return temp_midi_path, (duration_seconds, duration_minutes, duration_seconds_remainder)
        except Exception as e:
            print(f"[预览包装器] 生成预览失败: {str(e)}")
            root.after(0, lambda: messagebox.showerror("错误", f"生成预览失败: {str(e)}"))
            return None, (0, 0, 0)
    
    def cleanup(self):
        """清理资源"""
        if self.current_temp_path:
            try:
                self.preview_generator.cleanup_temp_file(self.current_temp_path)
                self.current_temp_path = None
            except Exception as e:
                print(f"[预览包装器] 清理临时文件时出错: {str(e)}")
        
        # 停止播放
        self.stop_playback()
        
        # 清理pygame资源
        try:
            import pygame.mixer
            if self._pygame_initialized and pygame.mixer.get_init():
                pygame.mixer.quit()
                self._pygame_initialized = False
            print("已清理预览资源")
        except Exception as e:
            print(f"清理pygame资源时出错: {str(e)}")
        
        # 调用midi_preview的stop_preview方法
        try:
            self.preview_generator.stop_preview()
        except Exception as e:
            print(f"[预览包装器] 停止预览时出错: {str(e)}")

# 创建全局实例
_preview_wrapper = None
def get_preview_wrapper():
    global _preview_wrapper
    if _preview_wrapper is None:
        _preview_wrapper = MidiPreviewWrapper()
    return _preview_wrapper
