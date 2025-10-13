import os
import shutil
import re
import sys
import subprocess

def get_version():
    """从main.py中获取版本号"""
    try:
        with open('main.py', 'r', encoding='utf-8') as f:
            content = f.read()
            match = re.search(r'VERSION\s*=\s*["\']([^"\']+)["\']', content)
            if match:
                return match.group(1)
    except Exception as e:
        print(f"读取版本号时出错: {str(e)}")
    return "1.0.0"  # 默认版本号

def should_clean_dist():
    """判断是否需要清理dist目录"""
    dist_dir = 'dist'
    
    # 如果dist目录不存在，不需要清理
    if not os.path.exists(dist_dir):
        return False
    
    # 检查dist目录是否为空
    if not os.listdir(dist_dir):
        return False
    
    # 检查dist目录中是否有OpenGamesAutoPlay子目录
    target_subdir = os.path.join(dist_dir, 'OpenGamesAutoPlay')
    if not os.path.exists(target_subdir):
        return True  # dist目录存在但目标子目录不存在，需要清理
    
    # 检查目标子目录中是否有exe文件
    version = get_version()
    exe_name = f"开放世界自动演奏_v{version}.exe"
    exe_path = os.path.join(target_subdir, exe_name)
    
    if not os.path.exists(exe_path):
        return True  # 目标exe文件不存在，需要清理
    
    # 检查exe文件是否可访问（没有权限问题）
    try:
        with open(exe_path, 'rb') as f:
            f.read(100)  # 尝试读取前100字节
        return False  # 文件可访问，不需要清理
    except (PermissionError, OSError):
        return True  # 文件访问权限有问题，需要清理

def clean_build():
    """清理构建文件夹"""
    dirs_to_clean = ['build']
    files_to_clean = ['*.spec']
    
    # 智能判断是否需要清理dist目录
    if should_clean_dist():
        dirs_to_clean.append('dist')
        print("检测到dist目录需要清理...")
    else:
        print("dist目录状态正常，跳过清理")
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                print(f"已删除 {dir_name} 目录")
            except PermissionError as e:
                print(f"警告: 无法删除 {dir_name} 目录，权限问题: {str(e)}")
                print("请手动关闭可能锁定该目录的程序后重试")
    
    for pattern in files_to_clean:
        for file in os.listdir('.'):
            if file.endswith('.spec'):
                try:
                    os.remove(file)
                    print(f"已删除 {file}")
                except PermissionError as e:
                    print(f"警告: 无法删除 {file}，权限问题: {str(e)}")

def ensure_pyinstaller():
    """确保 PyInstaller 正确安装"""
    try:
        # 尝试导入 PyInstaller
        import PyInstaller
        return True
    except ImportError:
        print("正在重新安装 PyInstaller...")
        try:
            # 使用 subprocess 运行 pip 命令
            subprocess.check_call([sys.executable, '-m', 'pip', 'uninstall', 'pyinstaller', '-y'])
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
            return True
        except Exception as e:
            print(f"安装 PyInstaller 失败: {str(e)}")
            return False

def build_exe():
    """构建exe文件"""
    # 确保 PyInstaller 已正确安装
    if not ensure_pyinstaller():
        return False
        
    # 获取版本号和设置输出目录
    version = get_version()
    output_dir = "dist/OpenGamesAutoPlay"
    exe_name = f"开放世界自动演奏_v{version}"
    
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(current_dir, 'icon.ico')
    runtime_hook = os.path.join(current_dir, 'runtime_hook.py')
    
    # 确保runtime_hook.py存在
    if not os.path.exists(runtime_hook):
        print("错误: 未找到 runtime_hook.py")
        return False
    
    try:
        # 清理旧的构建文件
        clean_build()
        
        print(f"开始构建OpenGamesAutoPlay版本 {version}...")
        
        # 使用 subprocess 运行 PyInstaller
        cmd = [
            sys.executable,
            '-m',
            'PyInstaller',
            'main.py',
            f'--name={exe_name}',
            '--onefile',
            '--windowed',
            f'--icon={icon_path}',
            f'--runtime-hook={runtime_hook}',
            '--add-data=icon.ico;.',
            '--uac-admin',
            # 隐藏导入 - 系统模块
            '--hidden-import=ctypes',
            '--hidden-import=ctypes.wintypes',
            '--hidden-import=json',
            '--hidden-import=threading',
            '--hidden-import=time',
            '--hidden-import=warnings',
            # 隐藏导入 - GUI相关
            '--hidden-import=tkinter',
            '--hidden-import=tkinter.ttk',
            '--hidden-import=ttkbootstrap',
            '--hidden-import=ttkbootstrap.constants',
            '--hidden-import=ttkbootstrap.style',
            '--hidden-import=ttkbootstrap.dialogs',
            '--hidden-import=ttkbootstrap.tooltip',
            '--hidden-import=ttkbootstrap.validation',
            # PIL相关模块
            '--hidden-import=PIL',
            '--hidden-import=PIL.Image',
            '--hidden-import=PIL.ImageTk',
            # 隐藏导入 - 音频和MIDI相关
            '--hidden-import=pygame',
            '--hidden-import=pygame.mixer',
            '--hidden-import=mido',
            '--hidden-import=rtmidi',
            # 隐藏导入 - 键盘控制
            '--hidden-import=keyboard',
            # 隐藏导入 - 自定义模块
            '--hidden-import=midi_player',
            '--hidden-import=keyboard_mapping',
            '--hidden-import=midi_analyzer',
            '--hidden-import=midi_preview_wrapper',
            # 隐藏导入 - 页面模块
            '--hidden-import=pages.help_dialog',
            '--hidden-import=pages.settings_dialog',
            '--hidden-import=pages.event_table_dialog',
            # 排除不必要的模块以减小体积
            '--exclude-module=matplotlib',
            '--exclude-module=numpy',
            '--exclude-module=pandas',
            '--exclude-module=scipy',
            # PIL模块是ttkbootstrap必需的，不能排除
            # '--exclude-module=PIL',
            '--exclude-module=PyQt5',
            '--exclude-module=PyQt5.QtCore',
            '--exclude-module=PyQt5.QtGui',
            '--exclude-module=PyQt5.QtWidgets',
            f'--distpath={output_dir}'
        ]
        
        subprocess.check_call(cmd)
        
        # 检查构建结果
        exe_path = os.path.join(output_dir, f"{exe_name}.exe")
        if os.path.exists(exe_path):
            # 复制必要文件到输出目录
            if os.path.exists('README.md'):
                shutil.copy2('README.md', output_dir)
            if os.path.exists('LICENSE'):
                shutil.copy2('LICENSE', output_dir)
                
            # 创建zip文件
            zip_name = f"OpenGamesAutoPlay_{version}.zip"
            shutil.make_archive(
                os.path.join('dist', f"OpenGamesAutoPlay_{version}"),
                'zip',
                output_dir
            )
            
            print("\n构建成功！")
            print(f"exe文件位置: {exe_path}")
            print(f"zip文件位置: dist/{zip_name}")
            
            # 显示文件大小
            exe_size = os.path.getsize(exe_path) / (1024 * 1024)  # MB
            print(f"生成文件大小: {exe_size:.2f} MB")
            
            return True
        else:
            print("\n构建失败：未找到输出文件")
            return False
            
    except Exception as e:
        print(f"\n构建过程中出错: {str(e)}")
        return False

def check_dependencies():
    """检查项目依赖是否已安装"""
    print("检查项目依赖...")
    
    dependencies = [
        'ttkbootstrap',
        'keyboard', 
        'mido',
        'pygame',
        'PyInstaller'
    ]
    
    missing_deps = []
    for dep in dependencies:
        try:
            __import__(dep)
            print(f"✓ {dep}")
        except ImportError:
            missing_deps.append(dep)
            print(f"✗ {dep}")
    
    if missing_deps:
        print(f"\n缺少依赖: {', '.join(missing_deps)}")
        print("请运行: pip install -r requirements.txt")
        return False
    
    print("所有依赖已安装！")
    return True

def main():
    """主函数"""
    print("=" * 50)
    print("OpenGamesAutoPlay 构建工具")
    print("=" * 50)
    
    # 检查依赖
    if not check_dependencies():
        return False
    
    # 构建exe文件
    if build_exe():
        print("\n" + "=" * 50)
        print("构建完成！")
        print("=" * 50)
        print("\n使用说明:")
        print("1. 生成的exe文件在 dist/OpenGamesAutoPlay 目录中")
        print("2. 请以管理员身份运行程序")
        print("3. 首次运行可能需要等待几秒钟")
        print("4. 确保游戏窗口在前台以获得最佳效果")
        return True
    else:
        print("\n构建失败，请检查错误信息")
        return False

if __name__ == '__main__':
    if main():
        sys.exit(0)
    else:
        sys.exit(1)
