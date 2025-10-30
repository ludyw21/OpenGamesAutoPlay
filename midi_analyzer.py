import json
import os
import sys

try:
    import mido
except ImportError:
    print("Warning: mido not found, please install it with 'pip install mido'")
    mido = None

from groups import group_for_note, get_note_name

# 定义黑键和白键的音级常量（参考MeowField_AutoPiano）
BLACK_PCS = {1, 3, 6, 8, 10}  # 黑键音级：C#, D#, F#, G#, A#
WHITE_PCS = [0, 2, 4, 5, 7, 9, 11]  # 白键音级：C, D, E, F, G, A, B

def _nearest_white_pc(pc: int, mode: str = "nearest") -> int:
    """将黑键音级映射到最近的白键音级
    mode: 'down' (优先向下), 'nearest' (最小绝对距离), 'scale' (到C大调的最接近音级)
    """
    pc = pc % 12
    if pc in WHITE_PCS:
        return pc
    if mode == "down":
        # 向下步进直到找到白键
        for d in range(1, 7):
            cand = (pc - d) % 12
            if cand in WHITE_PCS:
                return cand
    # 按绝对距离找到最近的白键（平局时优先选择较低的）
    best = None
    best_dist = 99
    for w in WHITE_PCS:
        dist = min((pc - w) % 12, (w - pc) % 12)
        if dist < best_dist or (dist == best_dist and ((w - pc) % 12) > ((pc - (best or 0)) % 12)):
            best = w
            best_dist = dist
    return best if best is not None else pc

def transpose_black_keys(events: list, strategy: str = "nearest") -> list:
    """将黑键音符转调到白键，保持note_on/note_off对的一致性
    strategy: 'down' | 'nearest' | 'scale'
    返回浅拷贝的事件列表
    """
    if not events:
        return []
    # 构建note on/off对
    result = []
    for ev in events:
        ev = dict(ev)
        note = ev.get('note')
        if note is not None:
            pc = note % 12
            if pc in BLACK_PCS:
                new_pc = _nearest_white_pc(pc, 'down' if strategy == 'down' else 'nearest')
                # 保持八度不变
                ev['note'] = (note - pc) + new_pc
        result.append(ev)
    return result


def _gather_notes_from_mido(mid, selected_tracks):
    """
    使用mido库精确解析MIDI文件，严格遵循MeowField_AutoPiano的解析逻辑
    激进优化版本：针对超大型MIDI文件进行性能优化
    """
    # 初始化全局缓存变量
    global _track_names_cache
    _track_names_cache = {}
    
    # 先获取所有音轨名称
    for track_idx, track in enumerate(mid.tracks):
        # 默认音轨名称
        track_name = f"Track {track_idx}"
        # 检查是否有音轨名称消息
        for msg in track:
            if msg.type == 'track_name':
                track_name = msg.name
                break
        # 存储音轨名称
        _track_names_cache[track_idx] = track_name
    
    print(f"[统计] 音轨信息: {_track_names_cache}")
    
    ticks_per_beat = getattr(mid, 'ticks_per_beat', 480)
    tempo = 500000  # default 120 BPM
    # 以绝对tick记录的tempo变化 (tick, tempo_us_per_beat)
    tempo_changes = [(0, tempo)]
    
    # 激进优化：预分配事件列表容量，大幅减少内存重分配
    total_track_events = sum(len(track) for track in mid.tracks)
    # 更精确的事件数量估计：假设1/3的事件是音符事件
    total_events_estimate = total_track_events // 3
    
    # 针对超大型文件进行特殊处理
    if total_track_events > 50000:  # 超过5万个总事件
        print(f"[性能警告] 检测到超大型MIDI文件，总事件数: {total_track_events}")
        # 限制最大事件数量，避免内存爆炸
        total_events_estimate = min(total_events_estimate, 100000)
    
    events = [None] * total_events_estimate if total_events_estimate > 1000 else []
    event_count = 0

    # 激进优化：快速跳过非音符密集的音轨
    for i, track in enumerate(mid.tracks):
        # 快速统计音轨中的音符事件比例
        total_msgs = len(track)
        if total_msgs == 0:
            continue
            
        note_events = 0
        for msg in track:
            if msg.type in ('note_on', 'note_off'):
                note_events += 1
        
        # 如果音符事件比例过低（<10%），快速跳过
        if note_events / total_msgs < 0.1 and total_msgs > 100:
            continue
            
        t = 0
        on_stack = {}
        cur_tempo = tempo
            
        for msg in track:
            t += msg.time
            
            # 激进优化：快速跳过非关键事件类型
            if msg.type not in ('note_on', 'note_off', 'set_tempo'):
                continue
                
            if msg.type == 'set_tempo':
                cur_tempo = msg.tempo
                tempo_changes.append((int(t), int(cur_tempo)))
            elif msg.type == 'note_on' and msg.velocity > 0:
                key = (msg.channel, msg.note)
                if key not in on_stack:
                    on_stack[key] = []
                on_stack[key].append((t, msg.velocity))
            elif msg.type in ('note_off', 'note_on'):
                if msg.type == 'note_on' and msg.velocity > 0:
                    continue
                key = (msg.channel, msg.note)
                if key in on_stack and on_stack[key]:
                    start_tick, vel = on_stack[key].pop(0)
                    
                    # 激进优化：直接赋值，避免条件判断
                    if event_count >= len(events):
                        # 动态扩容，但限制最大容量
                        if len(events) < 200000:  # 最大20万个事件
                            events.append({
                                'start_tick': start_tick,
                                'end_tick': t,
                                'channel': msg.channel,
                                'note': msg.note,
                                'velocity': vel,
                                'track': i
                            })
                        else:
                            # 超过容量限制，跳过后续事件
                            continue
                    else:
                        events[event_count] = {
                            'start_tick': start_tick,
                            'end_tick': t,
                            'channel': msg.channel,
                            'note': msg.note,
                            'velocity': vel,
                            'track': i
                        }
                    event_count += 1
    
    # 转换为精确时间：基于 tempo_changes 分段积分（PPQ），SMPTE 简化为常量换算
    # 过滤掉None值，确保events列表中只包含有效的事件对象
    events = [e for e in events if e is not None]
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

    # 激进优化：针对超大型文件优化tick转换缓存
    tick_cache = {}
    
    # 预计算tempo段的边界，减少循环次数
    tempo_segments = []
    for i in range(len(tempo_changes) - 1):
        tempo_segments.append((tempo_changes[i][0], tempo_changes[i+1][0], tempo_changes[i][1]))
    if tempo_changes:
        last_start, last_tempo = tempo_changes[-1]
        tempo_segments.append((last_start, float('inf'), last_tempo))
    
    def tick_to_seconds_ppq(target_tick):
        target_tick_int = int(target_tick)
        if target_tick_int in tick_cache:
            return tick_cache[target_tick_int]
            
        if target_tick_int <= 0:
            return 0.0
            
        # 使用预计算的tempo段进行快速查找
        for segment_start, segment_end, segment_tempo in tempo_segments:
            if segment_start <= target_tick_int < segment_end:
                # 计算时间
                acc = 0.0
                prev_tick = 0
                
                # 累加之前所有段的时间
                for prev_start, prev_end, prev_tempo in tempo_segments:
                    if prev_end <= segment_start:
                        dt = max(0, prev_end - prev_start)
                        acc += (dt * prev_tempo) / (ticks_per_beat * 1_000_000.0)
                        prev_tick = prev_end
                    else:
                        break
                
                # 当前段的时间
                dt_tail = max(0, target_tick_int - segment_start)
                acc += (dt_tail * segment_tempo) / (ticks_per_beat * 1_000_000.0)
                
                tick_cache[target_tick_int] = acc
                return acc
        
        # 默认计算（理论上不会执行到这里）
        acc = target_tick_int * (tempo / 1_000_000.0) / ticks_per_beat
        tick_cache[target_tick_int] = acc
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
        # 性能优化：批量处理事件时间计算
        for e in events:
            if e is not None:
                try:
                    st = tick_to_seconds_ppq(int(e['start_tick']))
                    et = tick_to_seconds_ppq(int(e['end_tick']))
                    e['start_time'] = st
                    e['end_time'] = max(et, st)
                    e['duration'] = max(0.0, e['end_time'] - e['start_time'])
                    e['group'] = group_for_note(e['note'])
                except (KeyError, TypeError, ValueError) as ex:
                    print(f"[警告] 处理事件时出错: {ex}")
    else:
        # 近似：按ticks线性映射到mido.length（若可用），保持事件相对位置
        max_tick = 0
        try:
            max_tick = max(int(e['end_tick']) for e in events if e is not None)
        except Exception:
            max_tick = 0
        scale = (mf_len / max_tick) if (mf_len > 0.0 and max_tick > 0) else 0.0
        if scale <= 0.0:
            # 无法近似时退化为默认120BPM（尽量避免抖动）
            for e in events:
                if e is not None:
                    try:
                        st = (int(e['start_tick']) * tempo) / (ticks_per_beat * 1_000_000.0)
                        et = (int(e['end_tick']) * tempo) / (ticks_per_beat * 1_000_000.0)
                        e['start_time'] = st
                        e['end_time'] = max(et, st)
                        e['duration'] = max(0.0, e['end_time'] - e['start_time'])
                        e['group'] = group_for_note(e['note'])
                    except (KeyError, TypeError, ValueError) as ex:
                        print(f"[警告] 处理SMPTE事件时出错: {ex}")
        else:
            for e in events:
                if e is not None:
                    try:
                        st = int(e['start_tick']) * scale
                        et = int(e['end_tick']) * scale
                        e['start_time'] = st
                        e['end_time'] = max(et, st)
                        e['duration'] = max(0.0, e['end_time'] - e['start_time'])
                        e['group'] = group_for_note(e['note'])
                    except (KeyError, TypeError, ValueError) as ex:
                        print(f"[警告] 处理SMPTE事件时出错: {ex}")
    
    # 过滤选中的音轨
    if selected_tracks:
        events = [e for e in events if e is not None and 'track' in e and e['track'] in selected_tracks]
    
    # 统计每个音轨的音符数量
    track_note_counts = {}
    for event in events:
        if event is not None:
            track_idx = event.get('track')
            if track_idx is not None:
                if track_idx not in track_note_counts:
                    track_note_counts[track_idx] = 0
                track_note_counts[track_idx] += 1
    print(f"[统计] 各音轨音符数量: {track_note_counts}")
    
    return events, track_note_counts


# 全局变量，用于存储最近一次解析的音轨名称信息
_track_names_cache = {}


class MidiAnalyzer:
    """MIDI文件分析器，负责解析MIDI文件并生成事件数据"""
    
    # 默认音域范围
    DEFAULT_MIN_NOTE = 48
    DEFAULT_MAX_NOTE = 83
    
    @staticmethod
    def get_track_names():
        """
        获取最近一次解析的MIDI文件的音轨名称信息
        
        Returns:
            dict: 音轨索引到音轨名称的映射
        """
        global _track_names_cache
        return _track_names_cache.copy()
    
    @staticmethod
    def _get_key_settings():
        """
        从config.json获取key_settings中的设置
        
        Returns:
            tuple: (min_note, max_note, black_key_mode)
        """
        try:
            # 首先尝试当前目录下的config.json
            config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if 'key_settings' in config:
                        min_note = config['key_settings'].get('min_note', MidiAnalyzer.DEFAULT_MIN_NOTE)
                        max_note = config['key_settings'].get('max_note', MidiAnalyzer.DEFAULT_MAX_NOTE)
                        black_key_mode = config['key_settings'].get('black_key_mode', 'support_black_key')
                        print(f"[MidiAnalyzer] 从{config_path}读取配置: black_key_mode={black_key_mode}")
                        return min_note, max_note, black_key_mode
            
            # 如果当前目录没有，尝试exe文件所在目录（适用于编译后的exe）
            if getattr(sys, 'frozen', False):
                # 如果是编译后的exe，使用exe所在目录
                exe_dir = os.path.dirname(sys.executable)
                config_path = os.path.join(exe_dir, 'config.json')
                if os.path.exists(config_path):
                    with open(config_path, 'r', encoding='utf-8') as f:
                        config = json.load(f)
                        if 'key_settings' in config:
                            min_note = config['key_settings'].get('min_note', MidiAnalyzer.DEFAULT_MIN_NOTE)
                            max_note = config['key_settings'].get('max_note', MidiAnalyzer.DEFAULT_MAX_NOTE)
                            black_key_mode = config['key_settings'].get('black_key_mode', 'support_black_key')
                            print(f"[MidiAnalyzer] 从exe目录{config_path}读取配置: black_key_mode={black_key_mode}")
                            return min_note, max_note, black_key_mode
        except Exception as e:
            print(f"获取配置时出错: {str(e)}")
        
        # 如果获取失败，返回默认值
        print("[MidiAnalyzer] 使用默认配置: black_key_mode=support_black_key")
        return MidiAnalyzer.DEFAULT_MIN_NOTE, MidiAnalyzer.DEFAULT_MAX_NOTE, 'support_black_key'
    
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
            tuple: (events, analysis_result, track_names)
                events: 生成的事件列表，按时间排序
                analysis_result: 分析结果字典，包含超限信息
                track_names: 音轨索引到音轨名称的映射
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
                events, track_note_counts = _gather_notes_from_mido(mid, selected_tracks)
                
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
                    else:
                        event['is_over_limit'] = False
                    
                    # 应用黑键自动降音处理（仅对非超限音符有效）
                    if not event['is_over_limit'] and black_key_mode == "auto_sharp":
                        # 使用更精准的黑键处理逻辑（参考MeowField_AutoPiano）
                        note = event['note']
                        pc = note % 12
                        if pc in BLACK_PCS:
                            # 找到最近的白键音级
                            new_pc = _nearest_white_pc(pc, 'nearest')
                            # 保持八度不变，只改变音级
                            event['note'] = (note - pc) + new_pc
                            # 更新组信息
                            event['group'] = group_for_note(event['note'])
                            # print(f"[MidiAnalyzer] 黑键处理: {get_note_name(note)} -> {get_note_name(event['note'])}")
                
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
                # 获取音轨名称信息
                track_names = MidiAnalyzer.get_track_names()
                return sorted(events, key=lambda x: x['time']), analysis_result, track_names, track_note_counts
                
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
                }, {}, {}
            

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
            }, {}, {}
