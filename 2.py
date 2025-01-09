import os
import sys
import zipfile
import winreg
import shutil
from pathlib import Path

def get_script_cmd():
    # 如果是 exe，使用 exe 路径
    if getattr(sys, 'frozen', False):
        exe_path = sys.executable
        return f'"{exe_path}"'
    # 如果是 py 脚本，使用 python.exe 来运行
    else:
        return f'"{sys.executable}" "{os.path.realpath(__file__)}"'

def add_context_menu():
    # 获取命令路径
    base_cmd = get_script_cmd()
    
    # 为文件夹添加压缩菜单
    folder_cmd = f'{base_cmd} compress "%1"'
    key_path = r'Directory\shell\CompressToZip'
    try:
        key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)
        winreg.SetValue(key, '', winreg.REG_SZ, '添加到&ZIP (Z)')
        command_key = winreg.CreateKey(key, 'command')
        winreg.SetValue(command_key, '', winreg.REG_SZ, folder_cmd)
    except Exception as e:
        print(f"添加文件夹右键菜单失败: {e}")

    # 为 .zip 文件添加右键菜单
    zip_cmd = f'{base_cmd} extract "%1"'
    menu_text = '解压到同名文件夹 (&E)'
    key_paths = [
        r'SystemFileAssociations\.zip\shell\ExtractToFolder',
        r'.zip\shell\ExtractToFolder'
    ]
    
    # 先尝试删除旧的注册表项
    for key_path in key_paths:
        try:
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path + r'\command')
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)
        except WindowsError:
            pass

    # 重新创建注册表项
    for key_path in key_paths:
        try:
            key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, key_path)
            winreg.SetValue(key, '', winreg.REG_SZ, menu_text)
            command_key = winreg.CreateKey(key, 'command')
            winreg.SetValue(command_key, '', winreg.REG_SZ, zip_cmd)
        except Exception as e:
            print(f"添加ZIP右键菜单失败 {key_path}: {e}")

def compress_folder(folder_path):
    folder_path = Path(folder_path)
    zip_path = folder_path.with_suffix('.zip')
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname)

def extract_zip(zip_path):
    zip_path = Path(zip_path)
    extract_path = zip_path.with_suffix('')
    
    with zipfile.ZipFile(zip_path, 'r') as zipf:
        zipf.extractall(extract_path)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        # 如果没有参数，添加右键菜单
        add_context_menu()
    elif len(sys.argv) == 3:
        command, path = sys.argv[1:3]
        if command == 'compress':
            compress_folder(path)
        elif command == 'extract':
            extract_zip(path) 