import json
import os

try:
    import mido
except ImportError:
    print("Warning: mido not found, please install it with 'pip install mido'")
    mido = None

from groups import group_for_note, get_note_name


def _gather_notes_from_mido(mid, selected_tracks):
    """
    使用mido库精确解析MIDI文件，严格遵循MeowField_AutoPiano的解析逻辑
    """
    ticks_per_beat = getattr(mid, 'ticks_per_beat', 480)
    tempo = 500000  # default 120 BPM
    # 以绝对tick记录的tempo变化 (tick, tempo_us_per_beat)
    tempo_changes = [(0, tempo)]
    # accumulate absolute time per track to the merged view
    events = []

    for i, track in enumerate(mid.tracks):
        t = 0
        on_stack = {}
        cur_tempo = tempo
        for msg in track:
            t += msg.time
            if msg.type == 'set_tempo':
                cur_tempo = msg.tempo
                tempo_changes.append((int(t), int(cur_tempo)))
            if msg.type == 'note_on' and msg.velocity > 0:
                on_stack.setdefault((msg.channel, msg.note), []).append((t, msg.velocity))
            elif msg.type in ('note_off', 'note_on'):
                if msg.type == 'note_on' and msg.velocity > 0:
                    continue
                key = (msg.channel, msg.note)
                if key in on_stack and on_stack[key]:
                    start_tick, vel = on_stack[key].pop(0)
                    events.append({
                        'start_tick': start_tick,
                        'end_tick': t,
                        'channel': msg.channel,
                        'note': msg.note,
                        'velocity': vel,
                        'track': i
                    })
    # 转换为精确时间：基于 tempo_changes 分段积分（PPQ），SMPTE 简化为常量换算
    if not events:
        return []

    # 统一排序并去除同tick重复，仅保留最后一次tempo（同tick后者覆盖前者）
    tempo_changes_sorted = sorted(tempo_changes, key=lambda x: int(x[0]))
    dedup = []
    for tk, tp in tempo_changes_sorted:
        if not dedup or int(tk) != int(dedup[-1][0]):
            dedup.append((int(tk), int(tp)))
        else:
            dedup[-1] = (int(tk), int(tp))
    tempo_changes = dedup if dedup else [(0, tempo)]

    def tick_to_seconds_ppq(target_tick):
        if target_tick <= 0:
            return 0.0
        acc = 0.0
        prev_tick = 0
        prev_tempo = tempo_changes[0][1]
        for i in range(1, len(tempo_changes)):
            cur_tick, cur_tempo = tempo_changes[i]
            if cur_tick > target_tick:
                break
            dt = max(0, cur_tick - prev_tick)
            acc += (dt * prev_tempo) / (ticks_per_beat * 1_000_000.0)  # 修正为1_000_000.0
            prev_tick = cur_tick
            prev_tempo = cur_tempo
        # tail
        dt_tail = max(0, int(target_tick) - int(prev_tick))
        acc += (dt_tail * prev_tempo) / (ticks_per_beat * 1_000_000.0)  # 修正为1_000_000.0
        return acc

    is_smpte = bool(ticks_per_beat < 0)
    # SMPTE: 这里暂不从 division 拆解fps和ticks_per_frame，后续如需可与 auto_player 统一实现
    # 先按近似：若为SMPTE，尝试使用 mido.MidiFile.length 比例映射（保持相对位置），否则退化为PPQ路径
    mf_len = 0.0
    try:
        mf_len = float(getattr(mid, 'length', 0.0) or 0.0)
    except Exception:
        mf_len = 0.0

    if not is_smpte:
        for e in events:
            st = tick_to_seconds_ppq(int(e['start_tick']))
            et = tick_to_seconds_ppq(int(e['end_tick']))
            e['start_time'] = st
            e['end_time'] = max(et, st)
            e['duration'] = max(0.0, e['end_time'] - e['start_time'])
            e['group'] = group_for_note(e['note'])
    else:
        # 近似：按ticks线性映射到mido.length（若可用），保持事件相对位置
        max_tick = 0
        try:
            max_tick = max(int(e['end_tick']) for e in events)
        except Exception:
            max_tick = 0
        scale = (mf_len / max_tick) if (mf_len > 0.0 and max_tick > 0) else 0.0
        if scale <= 0.0:
            # 无法近似时退化为默认120BPM（尽量避免抖动）
            for e in events:
                st = (int(e['start_tick']) * tempo) / (ticks_per_beat * 1_000_000.0)  # 修正为1_000_000.0
                et = (int(e['end_tick']) * tempo) / (ticks_per_beat * 1_000_000.0)  # 修正为1_000_000.0
                e['start_time'] = st
                e['end_time'] = max(et, st)
                e['duration'] = max(0.0, e['end_time'] - e['start_time'])
                e['group'] = group_for_note(e['note'])
        else:
            for e in events:
                st = int(e['start_tick']) * scale
                et = int(e['end_tick']) * scale
                e['start_time'] = st
                e['end_time'] = max(et, st)
                e['duration'] = max(0.0, e['end_time'] - e['start_time'])
                e['group'] = group_for_note(e['note'])
    
    # 过滤选中的音轨
    if selected_tracks:
        events = [e for e in events if e['track'] in selected_tracks]
    
    return events


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
        使用mido库获取准确的秒级时间信息
        
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
            
            # 只使用mido库进行精确解析，遵循MeowField_AutoPiano的解析逻辑
            if mido is None:
                print("Error: mido is not available. Please install it with 'pip install mido'")
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
                    'error': "mido库不可用"
                }
            
            # 使用mido进行精确解析
            print("[MidiAnalyzer] 使用mido精确解析MIDI文件")
            try:
                mid = mido.MidiFile(file_path)
                # 使用新的精确解析函数
                events = _gather_notes_from_mido(mid, selected_tracks)
                
                # 将事件转换为note_on/note_off格式以保持兼容性
                formatted_events = []
                
                for event in events:
                    # 添加note_on事件
                    note_on_event = {
                        'time': event['start_time'],
                        'type': 'note_on',
                        'note': event['note'],
                        'channel': event['channel'],
                        'group': event['group'],
                        'track': event['track'],
                        'duration': event['duration'],
                        'end': event['end_time'],
                        'velocity': event['velocity'],
                        'is_over_limit': False
                    }
                    formatted_events.append(note_on_event)
                    
                    # 添加note_off事件
                    note_off_event = {
                        'time': event['end_time'],
                        'type': 'note_off',
                        'note': event['note'],
                        'channel': event['channel'],
                        'group': event['group'],
                        'track': event['track'],
                        'velocity': event['velocity'],
                        'is_over_limit': False
                    }
                    formatted_events.append(note_off_event)
                
                events = formatted_events
                
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
                    'black_key_mode': black_key_mode
                }
                
                print(f"[MidiAnalyzer] 使用mido精确解析成功，事件数量: {len(events)}")
                return sorted(events, key=lambda x: x['time']), analysis_result
                
            except Exception as mido_error:
                print(f"[MidiAnalyzer] mido解析失败: {str(mido_error)}")
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
                    'error': "mido库解析失败"
                }
            
            # 使用mido库加载MIDI文件
            try:
                mid = mido.MidiFile(file_path)
                events = []
                active_notes = {}
                ticks_per_beat = mid.ticks_per_beat
                tempo = 500000
                
                # 扫描所有轨道查找tempo信息
                for track in mid.tracks:
                    for msg in track:
                        if msg.type == 'set_tempo':
                            tempo = msg.tempo
                            break
                    if tempo != 500000:
                        break
                
                # 计算tick到秒的转换因子
                tick_to_second = mido.tempo2bpm(tempo) / 60.0 / ticks_per_beat
                
                # 遍历所有音轨
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
                                    
                                    # 更新统计数据
                                    min_note_value = min(min_note_value, msg.note)
                                    max_note_value = max(max_note_value, msg.note)
                                    
                                    # 创建note_on事件
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
                                    
                                    # 创建note_off事件
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
            except Exception as load_error:
                print(f"[MidiAnalyzer] 加载MIDI文件时出错: {str(load_error)}")
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
                    'error': "mido库加载MIDI文件失败"
                }
            
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
