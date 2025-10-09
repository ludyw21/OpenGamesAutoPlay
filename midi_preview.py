# -*- coding: utf-8 -*-
"""
MIDI预览模块
- 专门用于处理事件簿数据并生成临时MIDI文件
- 为预览功能提供支持
"""
from typing import Optional, Dict, List
import os
import time
from datetime import datetime
import mido

class MidiPreviewGenerator:
    """MIDI预览生成器"""
    
    def __init__(self):
        """初始化预览生成器"""
        self.is_previewing = False
        self.current_midi_path = None
        self.temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp')
        # 确保临时目录存在
        os.makedirs(self.temp_dir, exist_ok=True)
    
    def generate_preview_midi(self, events: List[Dict], bpm: Optional[float] = None) -> Optional[str]:
        """
        从事件簿生成预览MIDI文件
        
        参数:
            events: 事件列表，包含note_on和note_off事件
            bpm: 可选的BPM值，用于设置MIDI文件的速度
            
        返回:
            str: 生成的临时MIDI文件路径，如果失败返回None
        """
        try:
            # 过滤有效事件（is_over_limit=False）
            valid_events = [event for event in events if event.get('is_over_limit', False) is False]
            print(f"[预览] 过滤有效事件: 原始={len(events)}, 有效={len(valid_events)}")
            
            # 检查原始事件数量
            if not events:
                print("[预览] 没有事件数据可处理")
                return None
            
            # 分离所有note_on和note_off事件
            note_events = []
            for event in valid_events:
                event_type = event.get('type')
                # 修改：不依赖velocity字段，直接处理所有note_on事件
                if event_type == 'note_on':
                    # 为note_on事件查找对应的结束时间
                    note = event.get('note')
                    time_on = event.get('time', 0)
                    # 尝试获取duration或end字段
                    if 'duration' in event:
                        duration = event.get('duration', 0.5)
                        time_off = time_on + duration
                    elif 'end' in event:
                        time_off = event.get('end', time_on + 0.5)
                    else:
                        # 如果没有duration或end，查找对应的note_off事件
                        time_off = self._find_note_off_time(valid_events, note, time_on)
                    
                    # 修改：设置默认velocity值，因为midi_analyzer生成的事件没有velocity字段
                    velocity = min(127, max(0, event.get('velocity', 64)))
                    channel = min(15, max(0, event.get('channel', 0)))
                    
                    # 添加note_on事件
                    note_events.append({
                        'time': time_on,
                        'type': 'note_on',
                        'note': note,
                        'velocity': velocity,
                        'channel': channel
                    })
                    
                    # 添加对应的note_off事件
                    note_events.append({
                        'time': time_off,
                        'type': 'note_off',
                        'note': note,
                        'velocity': velocity,  # 使用与note_on相同的velocity
                        'channel': channel
                    })
            
            # 即使没有有效事件，也尝试生成MIDI文件，只是没有音符
            if not note_events:
                print("[预览] 没有有效音符事件，将生成空MIDI文件")
            
            # 检查note_events是否为空，如果为空，创建一个空的MIDI文件，时长为0
            
            # 按时间排序所有事件，如果时间相同，note_off事件优先于note_on事件
            note_events.sort(key=lambda x: (x['time'], 1 if x['type'] == 'note_on' else 0))
            
            # 创建临时文件路径
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            temp_midi_path = os.path.join(self.temp_dir, f"preview_temp_{timestamp}.mid")
            
            # 设置MIDI参数
            ticks_per_beat = 480  # 默认值
            
            # 优先使用传入的bpm参数，如果没有则尝试从事件数据中获取
            if bpm is not None:
                global_tempo = mido.bpm2tempo(bpm)
                print(f"[预览] 使用传入的BPM: {bpm}")
            else:
                # 默认120 BPM
                global_tempo = mido.bpm2tempo(120)
                
                # 尝试从事件数据中获取tempo信息
                first_event = events[0] if events else None
                if first_event:
                    if 'initial_tempo' in first_event:
                        event_bpm = first_event['initial_tempo']
                        if event_bpm:
                            global_tempo = mido.bpm2tempo(event_bpm)
                            print(f"[预览] 从事件数据获取BPM: {event_bpm}")
                    elif 'tempo' in first_event:
                        tempo = first_event['tempo']
                        if tempo:
                            global_tempo = tempo
                            print(f"[预览] 从事件数据获取tempo: {tempo}")
            
            print(f"[预览] 使用ticks_per_beat={ticks_per_beat}, tempo={mido.tempo2bpm(global_tempo):.2f} BPM")
            
            # 创建MIDI文件
            midi_file = mido.MidiFile(ticks_per_beat=ticks_per_beat)
            track = mido.MidiTrack()
            midi_file.tracks.append(track)
            
            # 添加tempo消息
            track.append(mido.MetaMessage('set_tempo', tempo=global_tempo, time=0))
            
            # 生成MIDI消息
            last_tick = 0
            for event in note_events:
                # 转换时间为ticks
                current_tick = int(mido.second2tick(event['time'], ticks_per_beat, global_tempo))
                delta_time = max(0, current_tick - last_tick)
                
                if event['type'] == 'note_on':
                    track.append(mido.Message('note_on', 
                                            note=event['note'], 
                                            velocity=event['velocity'], 
                                            channel=event['channel'], 
                                            time=delta_time))
                else:  # note_off
                    track.append(mido.Message('note_off', 
                                            note=event['note'], 
                                            velocity=event['velocity'], 
                                            channel=event['channel'], 
                                            time=delta_time))
                last_tick = current_tick
            
            # 添加end_of_track消息
            track.append(mido.MetaMessage('end_of_track', time=0))
            
            # 保存MIDI文件
            midi_file.save(temp_midi_path)
            
            self.current_midi_path = temp_midi_path
            self.is_previewing = True
            
            # 计算并打印文件时长
            total_seconds = midi_file.length
            minutes = int(total_seconds // 60)
            seconds = int(total_seconds % 60)
            print(f"[预览] 成功生成临时MIDI文件: {os.path.basename(temp_midi_path)}, "
                  f"包含 {len(note_events)} 个MIDI事件, 时长: {minutes}分{seconds}秒 ({total_seconds:.2f}秒)")
            
            return temp_midi_path
            
        except Exception as e:
            print(f"[预览] 生成临时MIDI文件失败: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
    
    def _find_note_off_time(self, events: List[Dict], note: int, start_time: float) -> float:
        """
        查找对应的note_off事件时间
        
        参数:
            events: 事件列表
            note: 音符值
            start_time: 开始时间
            
        返回:
            float: 结束时间
        """
        for event in events:
            event_type = event.get('type')
            event_note = event.get('note')
            event_time = event.get('time', 0)
            
            # 找到相同音符且时间在start_time之后的note_off事件
            if event_time > start_time and event_note == note:
                if (event_type == 'note_off') or (event_type == 'note_on' and event.get('velocity', 0) == 0):
                    return event_time
        
        # 如果没找到，默认持续0.5秒
        return start_time + 0.5
    
    def cleanup_temp_file(self, file_path: Optional[str] = None) -> None:
        """
        清理临时文件
        
        参数:
            file_path: 要删除的文件路径，如果为None则删除当前文件
        """
        try:
            path_to_delete = file_path or self.current_midi_path
            if path_to_delete and os.path.exists(path_to_delete):
                os.remove(path_to_delete)
                print(f"[预览] 已删除临时MIDI文件: {os.path.basename(path_to_delete)}")
                if path_to_delete == self.current_midi_path:
                    self.current_midi_path = None
                    self.is_previewing = False
        except Exception as e:
            print(f"[预览] 删除临时文件时出错: {str(e)}")
    
    def stop_preview(self) -> None:
        """停止预览并清理资源"""
        self.is_previewing = False
        self.cleanup_temp_file()
    
    def get_midi_duration(self, midi_path: str) -> tuple[float, int, int]:
        """
        获取MIDI文件时长
        
        参数:
            midi_path: MIDI文件路径
            
        返回:
            tuple: (总秒数, 分钟, 秒)
        """
        try:
            mid = mido.MidiFile(midi_path)
            total_seconds = mid.length
            minutes = int(total_seconds // 60)
            seconds = int(total_seconds % 60)
            return total_seconds, minutes, seconds
        except Exception as e:
            print(f"[预览] 获取MIDI文件时长失败: {str(e)}")
            return 0.0, 0, 0

# 全局实例
g_midi_preview = None

def get_midi_preview() -> MidiPreviewGenerator:
    """
    获取MIDI预览生成器实例（单例模式）
    
    返回:
        MidiPreviewGenerator: 预览生成器实例
    """
    global g_midi_preview
    if g_midi_preview is None:
        g_midi_preview = MidiPreviewGenerator()
    return g_midi_preview
