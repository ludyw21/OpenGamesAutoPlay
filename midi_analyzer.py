import json
import os
try:
    import pretty_midi
except ImportError:
    print("Warning: pretty_midi not found, please install it with 'pip install pretty-midi'")
    pretty_midi = None

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
        使用pretty_midi库获取准确的秒级时间信息
        
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
            # 检查文件是否存在
            if not os.path.exists(file_path):
                print(f"[MidiAnalyzer] 文件不存在: {file_path}")
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
                    'total_over_limit_count': 0,
                    'config_min_note': None,
                    'config_max_note': None,
                    'transpose': transpose,
                    'octave_shift': octave_shift,
                    'black_key_mode': None,
                    'error': f"文件不存在: {file_path}"
                }
            
            # 获取音域范围和黑键模式
            min_note, max_note, black_key_mode = MidiAnalyzer._get_key_settings()
            
            # 初始化事件列表
            events = []
            
            # 初始化统计数据
            min_note_value = float('inf')
            max_note_value = -float('inf')
            
            # 使用pretty_midi库解析MIDI文件
            if pretty_midi is None:
                print("Error: pretty_midi is not available. Please install it.")
                # 如果没有pretty_midi，尝试使用mido作为备选方案
                try:
                    import mido
                    print("尝试使用mido作为备选方案解析MIDI文件")
                    
                    # mido解析逻辑
                    mid = mido.MidiFile(file_path)
                    events = []
                    active_notes = {}
                    ticks_per_beat = mid.ticks_per_beat
                    tempo = 500000  # 默认速度 (120 BPM)
                    
                    # 扫描所有轨道查找 tempo 信息
                    for track in mid.tracks:
                        current_tick = 0
                        for msg in track:
                            if msg.type == 'set_tempo':
                                tempo = msg.tempo
                                break
                        if tempo != 500000:
                            break
                    
                    # 计算tick到秒的转换因子
                    tick_to_second = mido.tempo2bpm(tempo) / 60.0 / ticks_per_beat
                    
                    # 遍历选中的音轨
                    for track_idx in selected_tracks:
                        if 0 <= track_idx < len(mid.tracks):
                            track = mid.tracks[track_idx]
                            current_tick = 0
                            for msg in track:
                                current_tick += msg.time
                                current_time = current_tick * tick_to_second
                                
                                if msg.type == 'note_on' and msg.velocity > 0:
                                    # 记录活动音符
                                    note_key = (msg.channel, msg.note)
                                    active_notes[note_key] = {
                                        'start_time': current_time,
                                        'velocity': msg.velocity,
                                        'track': track_idx
                                    }
                                elif (msg.type == 'note_off') or (msg.type == 'note_on' and msg.velocity == 0):
                                    # 处理note_off事件
                                    note_key = (msg.channel, msg.note)
                                    if note_key in active_notes:
                                        start_time = active_notes[note_key]['start_time']
                                        velocity = active_notes[note_key]['velocity']
                                        duration = current_time - start_time
                                        
                                        # 创建note_on和note_off事件
                                        note_on_event = {
                                            'time': start_time,
                                            'type': 'note_on',
                                            'note': msg.note,
                                            'channel': msg.channel,
                                            'group': group_for_note(msg.note),
                                            'track': track_idx,
                                            'duration': duration,
                                            'end': current_time,
                                            'velocity': velocity,
                                            'is_over_limit': False
                                        }
                                        events.append(note_on_event)
                                        
                                        note_off_event = {
                                            'time': current_time,
                                            'type': 'note_off',
                                            'note': msg.note,
                                            'channel': msg.channel,
                                            'group': group_for_note(msg.note),
                                            'track': track_idx,
                                            'velocity': velocity,
                                            'is_over_limit': False
                                        }
                                        events.append(note_off_event)
                                        
                                        del active_notes[note_key]
                    
                    # 应用移调并计算统计信息
                    transpose_total = transpose + octave_shift * 12
                    under_min_count = 0
                    over_max_count = 0
                    
                    for event in events:
                        event['note'] += transpose_total
                        event['group'] = group_for_note(event['note'])
                        
                        if event['note'] < min_note:
                            event['is_over_limit'] = True
                            under_min_count += 1
                        elif event['note'] > max_note:
                            event['is_over_limit'] = True
                            over_max_count += 1
                    
                    # 计算统计数据
                    if events:
                        min_note_value = min(events, key=lambda x: x['note'])['note']
                        max_note_value = max(events, key=lambda x: x['note'])['note']
                    else:
                        min_note_value = None
                        max_note_value = None
                    
                    analysis_result = {
                        'min_note': min_note_value,
                        'max_note': max_note_value,
                        'under_min_count': under_min_count,
                        'over_max_count': over_max_count,
                        'min_note_name': get_note_name(min_note_value) if min_note_value is not None else '',
                        'max_note_name': get_note_name(max_note_value) if max_note_value is not None else '',
                        'min_note_group': group_for_note(min_note_value) if min_note_value is not None else '',
                        'max_note_group': group_for_note(max_note_value) if max_note_value is not None else '',
                        'is_min_over_limit': min_note_value < min_note if min_note_value is not None else False,
                        'is_max_over_limit': max_note_value > max_note if max_note_value is not None else False,
                        'total_over_limit_count': under_min_count + over_max_count,
                        'config_min_note': min_note,
                        'config_max_note': max_note,
                        'transpose': transpose,
                        'octave_shift': octave_shift,
                        'black_key_mode': black_key_mode,
                        'warning': "使用mido备选方案解析，可能与pretty_midi结果有所不同"
                    }
                    
                    print(f"[MidiAnalyzer] 使用mido备选方案成功解析文件，事件数量: {len(events)}")
                    return sorted(events, key=lambda x: x['time']), analysis_result
                except Exception as fallback_error:
                    print(f"[MidiAnalyzer] mido备选方案也失败: {str(fallback_error)}")
                    # 返回空结果
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
                        'total_over_limit_count': 0,
                        'config_min_note': min_note,
                        'config_max_note': max_note,
                        'transpose': transpose,
                        'octave_shift': octave_shift,
                        'black_key_mode': black_key_mode,
                        'error': "pretty_midi和mido都不可用"
                    }
            
            # 加载MIDI文件，增加更详细的异常处理
            try:
                pm = pretty_midi.PrettyMIDI(file_path)
                # 初始化selected_tracks为所有音轨
                selected_tracks = list(range(len(pm.instruments)))
                print("[MidiAnalyzer] 使用pretty_midi处理音符")

                # 遍历选中的音轨（乐器）
                try:
                    for track_idx in selected_tracks:
                        if 0 <= track_idx < len(pm.instruments):
                            instrument = pm.instruments[track_idx]
                            channel = 9 if instrument.is_drum else track_idx
                            
                            # 遍历乐器中的所有音符
                            for note in instrument.notes:
                                try:
                                    # 直接获取秒级时间
                                    start_time = note.start
                                    end_time = note.end
                                    duration = end_time - start_time
                                    
                                    # 更新统计数据
                                    min_note_value = min(min_note_value, note.pitch)
                                    max_note_value = max(max_note_value, note.pitch)
                                    
                                    # 添加note_on事件
                                    note_on_event = {
                                        'time': start_time,
                                        'type': 'note_on',
                                        'note': note.pitch,
                                        'channel': channel,
                                        'group': group_for_note(note.pitch),
                                        'track': track_idx,
                                        'duration': duration,
                                        'end': end_time,
                                        'velocity': note.velocity,
                                        'is_over_limit': False  # 将在后面应用移调后更新
                                    }
                                    events.append(note_on_event)
                                    
                                    # 添加note_off事件
                                    note_off_event = {
                                        'time': end_time,
                                        'type': 'note_off',
                                        'note': note.pitch,
                                        'channel': channel,
                                        'group': group_for_note(note.pitch),
                                        'track': track_idx,
                                        'velocity': note.velocity,
                                        'is_over_limit': False  # 将在后面应用移调后更新
                                    }
                                    events.append(note_off_event)
                                except Exception as note_error:
                                    print(f"[MidiAnalyzer] 处理音符时出错: {str(note_error)}")
                                    # 跳过出错的音符，继续处理其他音符
                                    continue
                except Exception as instrument_error:
                    print(f"[MidiAnalyzer] 处理乐器时出错: {str(instrument_error)}")
                    # 如果pretty_midi的乐器处理失败，切换到mido备选方案
                    try:
                        import mido
                        print("[MidiAnalyzer] 切换到mido备选方案处理音符")
                        
                        mid = mido.MidiFile(file_path)
                        active_notes = {}
                        ticks_per_beat = mid.ticks_per_beat
                        tempo = 500000  # 默认速度 (120 BPM)
                        
                        # 扫描所有轨道查找 tempo 信息
                        for track in mid.tracks:
                            for msg in track:
                                if msg.type == 'set_tempo':
                                    tempo = msg.tempo
                                    break
                            if tempo != 500000:
                                break
                        
                        # 计算tick到秒的转换因子
                        tick_to_second = mido.tempo2bpm(tempo) / 60.0 / ticks_per_beat
                        
                        # 遍历选中的音轨
                        for track_idx in selected_tracks:
                            if 0 <= track_idx < len(mid.tracks):
                                track = mid.tracks[track_idx]
                                current_tick = 0
                                for msg in track:
                                    current_tick += msg.time
                                    current_time = current_tick * tick_to_second
                                    
                                    try:
                                        if msg.type == 'note_on' and msg.velocity > 0:
                                            # 记录活动音符
                                            note_key = (msg.channel, msg.note)
                                            active_notes[note_key] = {
                                                'start_time': current_time,
                                                'velocity': msg.velocity,
                                                'track': track_idx
                                            }
                                        elif (msg.type == 'note_off') or (msg.type == 'note_on' and msg.velocity == 0):
                                            # 处理note_off事件
                                            note_key = (msg.channel, msg.note)
                                            if note_key in active_notes:
                                                start_time = active_notes[note_key]['start_time']
                                                velocity = active_notes[note_key]['velocity']
                                                duration = current_time - start_time
                                                
                                                # 创建note_on和note_off事件
                                                note_on_event = {
                                                    'time': start_time,
                                                    'type': 'note_on',
                                                    'note': msg.note,
                                                    'channel': msg.channel,
                                                    'group': group_for_note(msg.note),
                                                    'track': track_idx,
                                                    'duration': duration,
                                                    'end': current_time,
                                                    'velocity': velocity,
                                                    'is_over_limit': False
                                                }
                                                events.append(note_on_event)
                                                
                                                note_off_event = {
                                                    'time': current_time,
                                                    'type': 'note_off',
                                                    'note': msg.note,
                                                    'channel': msg.channel,
                                                    'group': group_for_note(msg.note),
                                                    'track': track_idx,
                                                    'velocity': velocity,
                                                    'is_over_limit': False
                                                }
                                                events.append(note_off_event)
                                                
                                                del active_notes[note_key]
                                    except Exception as msg_error:
                                        print(f"[MidiAnalyzer] 处理MIDI消息时出错: {str(msg_error)}")
                                        continue
                        print(f"[MidiAnalyzer] mido备选方案成功处理事件，当前事件数量: {len(events)}")
                    except Exception as mido_error:
                        print(f"[MidiAnalyzer] mido备选方案也失败: {str(mido_error)}")
                        # 继续使用已经收集到的事件（如果有）
            
            except Exception as load_error:
                print(f"[MidiAnalyzer] 加载MIDI文件时出错: {str(load_error)}")
                # 尝试使用mido作为最后的备选方案
                try:
                    import mido
                    print("[MidiAnalyzer] 最后尝试mido备选方案")
                    
                    mid = mido.MidiFile(file_path)
                    events = []
                    active_notes = {}
                    ticks_per_beat = mid.ticks_per_beat
                    tempo = 500000
                    
                    for track in mid.tracks:
                        for msg in track:
                            if msg.type == 'set_tempo':
                                tempo = msg.tempo
                                break
                        if tempo != 500000:
                            break
                    
                    tick_to_second = mido.tempo2bpm(tempo) / 60.0 / ticks_per_beat
                    
                    for track_idx, track in enumerate(mid.tracks):
                        current_tick = 0
                        for msg in track:
                            current_tick += msg.time
                            current_time = current_tick * tick_to_second
                            
                            try:
                                if msg.type == 'note_on' and msg.velocity > 0:
                                    note_key = (msg.channel, msg.note)
                                    active_notes[note_key] = {
                                        'start_time': current_time,
                                        'velocity': msg.velocity,
                                        'track': track_idx
                                    }
                                elif (msg.type == 'note_off') or (msg.type == 'note_on' and msg.velocity == 0):
                                    note_key = (msg.channel, msg.note)
                                    if note_key in active_notes:
                                        start_time = active_notes[note_key]['start_time']
                                        velocity = active_notes[note_key]['velocity']
                                        duration = current_time - start_time
                                        
                                        note_on_event = {
                                            'time': start_time,
                                            'type': 'note_on',
                                            'note': msg.note,
                                            'channel': msg.channel,
                                            'group': group_for_note(msg.note),
                                            'track': track_idx,
                                            'duration': duration,
                                            'end': current_time,
                                            'velocity': velocity,
                                            'is_over_limit': False
                                        }
                                        events.append(note_on_event)
                                        
                                        note_off_event = {
                                            'time': current_time,
                                            'type': 'note_off',
                                            'note': msg.note,
                                            'channel': msg.channel,
                                            'group': group_for_note(msg.note),
                                            'track': track_idx,
                                            'velocity': velocity,
                                            'is_over_limit': False
                                        }
                                        events.append(note_off_event)
                                        
                                        del active_notes[note_key]
                            except Exception:
                                continue
                except Exception:
                    print("[MidiAnalyzer] 所有解析方法都失败")
            
            # 应用移调、转位和黑键自动降音处理
            transpose_total = transpose + octave_shift * 12
            over_max_count = 0
            under_min_count = 0
            
            # 先对事件进行排序，确保按照时间顺序处理
            events.sort(key=lambda x: x['time'])
            
            # 应用移调和计算超限情况
            for event in events:
                # 应用移调
                original_note = event['note']
                event['note'] = original_note + transpose_total
                
                # 更新组信息
                event['group'] = group_for_note(event['note'])
                
                # 判断是否超限（基于移调后的值）
                if event['note'] < min_note:
                    event['is_over_limit'] = True
                    under_min_count += 1
                elif event['note'] > max_note:
                    event['is_over_limit'] = True
                    over_max_count += 1
                else:
                    event['is_over_limit'] = False
                
                # 应用黑键自动降音处理（仅对非超限音符有效）
                if not event['is_over_limit'] and black_key_mode == "auto_sharp":
                    # 检查是否是需要降调的黑键（只对升号#进行处理，忽略降号b）
                    note_name = get_note_name(event['note'])
                    if '#' in note_name:
                        # 降半音到白键
                        event['note'] -= 1
                        # 更新组信息
                        event['group'] = group_for_note(event['note'])
            
            # 计算最终的最小和最大音符值（移调后的）
            if events:
                min_note_value = min(events, key=lambda x: x['note'])['note']
                max_note_value = max(events, key=lambda x: x['note'])['note']
            else:
                min_note_value = None
                max_note_value = None
            
            # 构建分析结果
            analysis_result = {
                'min_note': min_note_value,
                'max_note': max_note_value,
                'under_min_count': under_min_count,
                'over_max_count': over_max_count,
                'min_note_name': get_note_name(min_note_value) if min_note_value is not None else '',
                'max_note_name': get_note_name(max_note_value) if max_note_value is not None else '',
                'min_note_group': group_for_note(min_note_value) if min_note_value is not None else '',
                'max_note_group': group_for_note(max_note_value) if max_note_value is not None else '',
                'is_min_over_limit': min_note_value < min_note if min_note_value is not None else False,
                'is_max_over_limit': max_note_value > max_note if max_note_value is not None else False,
                'total_over_limit_count': under_min_count + over_max_count,
                'config_min_note': min_note,
                'config_max_note': max_note,
                'transpose': transpose,
                'octave_shift': octave_shift,
                'black_key_mode': black_key_mode
            }
            
            # 输出详细的调试信息，包括时间精度
            print(f"[MidiAnalyzer] 事件数量: {len(events)}, 最小音符: {min_note_value}, 最大音符: {max_note_value}")
            print(f"[MidiAnalyzer] 超限统计: 低于最小 {under_min_count}, 高于最大 {over_max_count}")
            
            # 如果有事件，打印第一个和最后一个事件的时间信息
            if events:
                first_event = events[0]
                last_event = events[-1]
                print(f"[MidiAnalyzer] 第一个事件时间: {first_event['time']:.3f}s, 最后一个事件时间: {last_event['time']:.3f}s")
            
            return events, analysis_result
        except Exception as e:
            print(f"[MidiAnalyzer] 分析MIDI文件出错: {e}")
            import traceback
            traceback.print_exc()  # 输出详细的错误堆栈
            # 返回空结果，避免应用崩溃
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
                'total_over_limit_count': 0,
                'config_min_note': None,
                'config_max_note': None,
                'transpose': transpose,
                'octave_shift': octave_shift,
                'black_key_mode': None
            }
