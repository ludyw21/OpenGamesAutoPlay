"""
Piano pitch groups and helper utilities.
Groups are defined in MIDI note numbers.
"""
from typing import Dict, Tuple, List

# 定义音符名称
NOTE_NAMES = ['c', 'c#', 'd', 'd#', 'e', 'f', 'f#', 'g', 'g#', 'a', 'a#', 'b']

# MIDI note numbers: A₂=21, c¹=60 (中央C), c⁵=108
# 按照标准钢琴音域定义音组 (21..108)
GROUPS: Dict[str, Tuple[int, int]] = {
    "大字二组 (A₂-B₂)": (21, 23),
    "大字一组 (C₁-B₁)": (24, 35),
    "大字组 (C-B)": (36, 47),
    "小字组 (c-b)": (48, 59),
    "小字一组 (c¹-b¹)": (60, 71),
    "小字二组 (c²-b²)": (72, 83),
    "小字三组 (c³-b³)": (84, 95),
    "小字四组 (c⁴-b⁴)": (96, 107),
    "小字五组 (c⁵)": (108, 108),
}

ORDERED_GROUP_NAMES: List[str] = list(GROUPS.keys())

# 预生成88个钢琴音符的名称映射
NOTE_NAME_MAP: Dict[int, str] = {}

# 初始化音符名称映射
for note_number in range(21, 109):
    # 获取音符名（不包含八度）
    note_index = note_number % 12
    note_name = NOTE_NAMES[note_index]
    
    # 确定八度符号和音符大小写
    if note_number >= 60:
        # 小字一组及以上（使用小写字母）
        # 小字一组 (60-71)
        if 60 <= note_number <= 71:
            octave_symbol = '¹'
        # 小字二组 (72-83)
        elif 72 <= note_number <= 83:
            octave_symbol = '²'
        # 小字三组 (84-95)
        elif 84 <= note_number <= 95:
            octave_symbol = '³'
        # 小字四组 (96-107)
        elif 96 <= note_number <= 107:
            octave_symbol = '⁴'
        # 小字五组 (108)
        else:
            octave_symbol = '⁵'
    elif note_number >= 48:
        # 小字组 (48-59)
        # 小字组使用小写字母，无八度标记
        octave_symbol = ''
    elif note_number >= 36:
        # 大字组 (36-47)
        # 大字组使用大写字母，无八度标记
        note_name = note_name.upper()
        octave_symbol = ''
    elif note_number >= 24:
        # 大字一组 (24-35)
        note_name = note_name.upper()
        octave_symbol = '₁'
    else:
        # 大字二组 (21-23)
        note_name = note_name.upper()
        octave_symbol = '₂'
    
    # 组合音符名称和八度
    full_note_name = f"{note_name}{octave_symbol}"
    NOTE_NAME_MAP[note_number] = full_note_name


def group_for_note(note: int) -> str:
    """根据音符数字获取所属分组"""
    for name, (lo, hi) in GROUPS.items():
        if lo <= note <= hi:
            return name
    return "未知"

def get_note_name(note: int) -> str:
    """根据音符数字获取标准钢琴音符名称
    例如：60 为 c¹，61 为 c¹#，62 为 d¹
    """
    # 检查音符是否在有效范围内
    if 21 <= note <= 108:
        return NOTE_NAME_MAP.get(note, str(note))
    return str(note)


def filter_notes_by_groups(notes: List[dict], selected_groups: List[str]) -> List[dict]:
    if not selected_groups:
        return notes
    ranges = [GROUPS[name] for name in selected_groups if name in GROUPS]
    if not ranges:
        return notes
    out = []
    for ev in notes:
        n = ev.get('note')
        if n is None:
            out.append(ev)
            continue
        for lo, hi in ranges:
            if lo <= n <= hi:
                out.append(ev)
                break
    return out
