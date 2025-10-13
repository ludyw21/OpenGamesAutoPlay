import os
import sys

def setup_environment():
    # 确保临时文件目录存在
    temp_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'OpenGamesAutoPlay')
    os.makedirs(temp_dir, exist_ok=True)
    
    # 设置工作目录
    if getattr(sys, 'frozen', False):
        os.chdir(os.path.dirname(sys.executable))

# 强制导入ttkbootstrap相关模块，确保PyInstaller正确打包
import ttkbootstrap
import ttkbootstrap.style
import ttkbootstrap.constants
import ttkbootstrap.dialogs
import ttkbootstrap.tooltip
import ttkbootstrap.validation

# 强制导入PIL模块，ttkbootstrap需要PIL来处理图像
import PIL
import PIL.Image
import PIL.ImageTk

setup_environment()