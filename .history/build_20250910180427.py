#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地打包脚本
"""


import os
import sys
import subprocess
import platform

def build_app():
    """打包应用程序"""
    print("开始打包应用程序...")
    
    # 检查 PyInstaller
    try:
        subprocess.run(["pyinstaller", "--version"], check=True, capture_output=True)
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("错误: 未找到 PyInstaller，请先安装：pip install pyinstaller")
        return False
    
    # 打包命令
    system = platform.system()
    if system == "Windows":
        cmd = [
            "pyinstaller",
            "--onefile",
            "--windowed",
            "--name=Excel图片下载器",
            "--icon=icon.ico",
            "main.py"
        ]
    else:  # macOS/Linux
        cmd = [
            "pyinstaller",
            "--onefile",
            "--windowed",
            "--name=Excel图片下载器",
            "main.py"
        ]
    
    try:
        print(f"执行命令: {' '.join(cmd)}")
        result = subprocess.run(cmd, check=True)
        print("✅ 打包成功！")
        print(f"可执行文件位置: {os.path.join('dist', 'Excel图片下载器')}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 打包失败: {e}")
        return False

def clean_build():
    """清理构建文件"""
    import shutil
    
    dirs_to_remove = ["build", "dist", "__pycache__"]
    files_to_remove = ["*.spec"]
    
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"已删除目录: {dir_name}")
    
    # 删除 .spec 文件
    for file in os.listdir("."):
        if file.endswith(".spec") and file != "main.spec":
            os.remove(file)
            print(f"已删除文件: {file}")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "clean":
        clean_build()
    else:
        build_app()
