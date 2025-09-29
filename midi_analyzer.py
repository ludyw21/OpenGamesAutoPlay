import mido
from groups import group_for_note

class MidiAnalyzer:
    """MIDI文件分析器，负责解析MIDI文件并生成事件数据"""
    
    @staticmethod
    def analyze_midi_file(file_path, selected_tracks):
        """
        分析MIDI文件，从选中的音轨生成事件数据
        
        Args:
            file_path: MIDI文件路径
            selected_tracks: 选中的音轨索引集合
            
        Returns:
            list: 生成的事件列表，按时间排序
        """
        try:
            # 初始化事件列表和音符开始记录
            events = []
            note_starts = {}
            
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
                            note_starts[note_key] = {
                                'time': track_time,
                                'track': track_idx,
                                'velocity': msg.velocity
                            }
                            
                            # 添加note_on事件
                            events.append({
                                'time': track_time,
                                'type': 'note_on',
                                'note': msg.note,
                                'channel': msg.channel,
                                'group': group_for_note(msg.note),
                                'track': track_idx
                            })
                        
                        # 处理音符结束事件
                        elif (msg.type == 'note_off') or (msg.type == 'note_on' and msg.velocity == 0):
                            note_key = (msg.channel, msg.note)
                            if note_key in note_starts:
                                # 计算音符持续时间
                                start_time = note_starts[note_key]['time']
                                duration = track_time - start_time
                                
                                # 更新对应的note_on事件，添加持续时间和结束时间
                                for event in events:
                                    if (event['type'] == 'note_on' and 
                                        event['channel'] == msg.channel and 
                                        event['note'] == msg.note and 
                                        event['track'] == track_idx and
                                        abs(event['time'] - start_time) < 0.01):
                                        event['duration'] = duration
                                        event['end'] = track_time
                                        break
                            
                            # 添加note_off事件
                            events.append({
                                'time': track_time,
                                'type': 'note_off',
                                'note': msg.note,
                                'channel': msg.channel,
                                'group': group_for_note(msg.note),
                                'track': track_idx
                            })
            
            # 按时间排序所有事件
            events.sort(key=lambda x: x['time'])
            print(f"MIDI分析器: 已生成事件数据：{len(events)}个事件，{len(events)/2}个音符。")
            return events
            
        except Exception as e:
            print(f"MIDI分析器: 生成事件数据时出错: {str(e)}")
            # 如果出错，返回空列表
            return []
