import mido
import json
import os
from groups import group_for_note, get_note_name

class MidiAnalyzer:
    """MIDI文件分析器，负责解析MIDI文件并生成事件数据"""
    
    # 默认音域范围
    DEFAULT_MIN_NOTE = 48
    DEFAULT_MAX_NOTE = 83
    
    @staticmethod
    def _get_key_settings():
        """
        从config.json获取key_settings中的设置
        
        Returns:
            tuple: (min_note, max_note, black_key_mode)
        """
        try:
            config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if 'key_settings' in config:
                        min_note = config['key_settings'].get('min_note', MidiAnalyzer.DEFAULT_MIN_NOTE)
                        max_note = config['key_settings'].get('max_note', MidiAnalyzer.DEFAULT_MAX_NOTE)
                        black_key_mode = config['key_settings'].get('black_key_mode', 'auto_sharp')
                        return min_note, max_note, black_key_mode
        except Exception as e:
            print(f"获取配置时出错: {str(e)}")
        
        # 如果获取失败，返回默认值
        return MidiAnalyzer.DEFAULT_MIN_NOTE, MidiAnalyzer.DEFAULT_MAX_NOTE, 'auto_sharp'
    
    @staticmethod
    def get_over_limit_info(analysis_result):
        """
        从分析结果中提取超限信息
        
        Args:
            analysis_result: 分析结果字典
            
        Returns:
            dict: 包含超限信息的字典，字段为：
                - min_note: 最低音
                - under_min_count: 低于最低音的数量
                - max_note: 最高音
                - over_max_count: 高于最高音的数量
        """
        return {
            'min_note': analysis_result.get('min_note'),
            'under_min_count': analysis_result.get('under_min_count', 0),
            'max_note': analysis_result.get('max_note'),
            'over_max_count': analysis_result.get('over_max_count', 0)
        }
    
    @staticmethod
    def analyze_midi_file(file_path, selected_tracks, transpose=0, octave_shift=0):
        """
        分析MIDI文件，从选中的音轨生成事件数据，并标记超限音符
        
        Args:
            file_path: MIDI文件路径
            selected_tracks: 选中的音轨索引集合
            transpose: 移调值（半音），默认为0
            octave_shift: 转位值（八度），默认为0
            
        Returns:
            tuple: (events, analysis_result)
                events: 生成的事件列表，按时间排序
                analysis_result: 分析结果字典，包含超限信息
        """
        try:
            # 获取音域范围和黑键模式
            min_note, max_note, black_key_mode = MidiAnalyzer._get_key_settings()
            
            # 初始化事件列表和音符开始记录
            events = []
            # 使用字典存储note_on事件的索引，避免O(n²)的查找
            note_on_events = {}
            
            # 初始化统计数据
            min_note_value = float('inf')
            max_note_value = -float('inf')
            over_max_count = 0
            under_min_count = 0
            
            # 加载MIDI文件
            mid = mido.MidiFile(file_path)
            ppqn = mid.ticks_per_beat
            
            # 遍历选中的音轨
            for track_idx in selected_tracks:
                if 0 <= track_idx < len(mid.tracks):
                    track = mid.tracks[track_idx]
                    track_time = 0.0
                    
                    # 遍历音轨中的所有消息
                    for msg in track:
                        # 累加时间（转换为秒）
                        track_time += msg.time / ppqn * 4
                        
                        # 处理音符开始事件
                        if msg.type == 'note_on' and msg.velocity > 0:
                            note_key = (msg.channel, msg.note)
                            
                            # 更新统计数据
                            min_note_value = min(min_note_value, msg.note)
                            max_note_value = max(max_note_value, msg.note)
                            
                            # 判断是否超限
                            is_over_limit = msg.note < min_note or msg.note > max_note
                            
                            # 统计超限数量
                            if msg.note < min_note:
                                under_min_count += 1
                            elif msg.note > max_note:
                                over_max_count += 1
                            
                            # 添加note_on事件，包含超限标记
                            event = {
                                'time': track_time,
                                'type': 'note_on',
                                'note': msg.note,
                                'channel': msg.channel,
                                'group': group_for_note(msg.note),
                                'track': track_idx,
                                'is_over_limit': is_over_limit  # 添加超限标记
                            }
                            events.append(event)
                            
                            # 存储事件引用，用于后续快速查找
                            if note_key not in note_on_events:
                                note_on_events[note_key] = []
                            note_on_events[note_key].append({
                                'event': event,
                                'time': track_time,
                                'track': track_idx
                            })
                        
                        # 处理音符结束事件
                        elif (msg.type == 'note_off') or (msg.type == 'note_on' and msg.velocity == 0):
                            note_key = (msg.channel, msg.note)
                            if note_key in note_on_events and note_on_events[note_key]:
                                # 找到最近的未配对的note_on事件
                                # 倒序查找，找到最近的匹配事件
                                for i in range(len(note_on_events[note_key]) - 1, -1, -1):
                                    note_info = note_on_events[note_key][i]
                                    if note_info['track'] == track_idx and 'duration' not in note_info['event']:
                                        # 计算音符持续时间
                                        start_time = note_info['time']
                                        duration = track_time - start_time
                                        
                                        # 直接更新事件，避免线性查找
                                        note_info['event']['duration'] = duration
                                        note_info['event']['end'] = track_time
                                        break
                            
                            # 判断是否超限
                            is_over_limit = msg.note < min_note or msg.note > max_note
                            
                            # 添加note_off事件，包含超限标记
                            events.append({
                                'time': track_time,
                                'type': 'note_off',
                                'note': msg.note,
                                'channel': msg.channel,
                                'group': group_for_note(msg.note),
                                'track': track_idx,
                                'is_over_limit': is_over_limit  # 添加超限标记
                            })
            
            # 释放临时字典内存
            note_on_events = None
            
            # 应用移调、转位和黑键自动降音处理
            # 计算总偏移量：移调值 + 转位值*12
            total_offset = transpose + (octave_shift * 12)
            
            # 需要进行黑键自动降音的音符值（X#）
            black_notes = {61, 63, 66, 68, 70, 73, 75, 78, 80, 82, 85, 87}
            
            # 更新事件列表中的音符值
            for event in events:
                # 应用移调和转位
                if 'note' in event:
                    # 先应用移调和转位
                    event['note'] += total_offset
                    
                    # 如果是黑键自动降音模式，且音符是黑键，则降1个半音
                    if black_key_mode == 'auto_sharp' and event['note'] in black_notes:
                        event['note'] -= 1
                    
                    # 更新音组
                    event['group'] = group_for_note(event['note'])
                    
                    # 重新判断是否超限
                    event['is_over_limit'] = event['note'] < min_note or event['note'] > max_note
            
            # 按时间排序所有事件
            events.sort(key=lambda x: x['time'])
            
            # 重新计算统计数据（基于处理后的音符值）
            processed_min_note = float('inf')
            processed_max_note = -float('inf')
            processed_under_min_count = 0
            processed_over_max_count = 0
            
            # 只考虑note_on事件的音符值进行统计
            for event in events:
                if event['type'] == 'note_on' and 'note' in event:
                    note_value = event['note']
                    processed_min_note = min(processed_min_note, note_value)
                    processed_max_note = max(processed_max_note, note_value)
                    
                    if note_value < min_note:
                        processed_under_min_count += 1
                    elif note_value > max_note:
                        processed_over_max_count += 1
            
            # 构建分析结果（基于处理后的音符值）
            analysis_result = {
                'min_note': processed_min_note if processed_min_note != float('inf') else None,
                'max_note': processed_max_note if processed_max_note != -float('inf') else None,
                'under_min_count': processed_under_min_count,
                'over_max_count': processed_over_max_count,
                'min_note_name': get_note_name(processed_min_note) if processed_min_note != float('inf') else '',
                'max_note_name': get_note_name(processed_max_note) if processed_max_note != -float('inf') else '',
                'min_note_group': group_for_note(processed_min_note) if processed_min_note != float('inf') else '',
                'max_note_group': group_for_note(processed_max_note) if processed_max_note != -float('inf') else '',
                'is_min_over_limit': processed_min_note < min_note if processed_min_note != float('inf') else False,
                'is_max_over_limit': processed_max_note > max_note if processed_max_note != -float('inf') else False,
                'total_over_limit_count': processed_under_min_count + processed_over_max_count,
                'config_min_note': min_note,
                'config_max_note': max_note,
                'transpose': transpose,
                'octave_shift': octave_shift,
                'black_key_mode': black_key_mode
            }
            
            print(f"MIDI分析器: 已生成事件数据：{len(events)}个事件，{len(events)/2}个音符。")
            print(f"超限分析: 低于最低音数量={processed_under_min_count}, 高于最高音数量={processed_over_max_count}")
            print(f"处理设置: 移调={transpose}, 转位={octave_shift}, 黑键模式={black_key_mode}")
            
            return events, analysis_result
            
        except Exception as e:
            print(f"MIDI分析器: 生成事件数据时出错: {str(e)}")
            # 如果出错，返回空列表和空分析结果
            return [], {
                'min_note': None,
                'max_note': None,
                'under_min_count': 0,
                'over_max_count': 0,
                'min_note_name': '',
                'max_note_group': '',
                'max_note_name': '',
                'min_note_group': '',
                'is_min_over_limit': False,
                'is_max_over_limit': False,
                'total_over_limit_count': 0
            }
