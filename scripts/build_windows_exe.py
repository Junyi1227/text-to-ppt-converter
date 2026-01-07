#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包腳本 - 將 Python 程式打包成 Windows .exe 執行檔
Build Script - Package Python program to Windows .exe

在 Mac 上執行此腳本需要使用交叉編譯，建議在 Windows 環境執行
或使用 GitHub Actions / CI/CD 自動打包
"""

import os
import sys
import subprocess

def check_pyinstaller():
    """檢查 PyInstaller 是否已安裝"""
    try:
        import PyInstaller
        print("✅ PyInstaller 已安裝")
        return True
    except ImportError:
        print("❌ PyInstaller 未安裝")
        print("正在安裝 PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        return True

def build_exe():
    """打包成 .exe 執行檔"""
    
    print("=" * 60)
    print("文字轉 PowerPoint 工具 - Windows 執行檔打包程式")
    print("=" * 60)
    print()
    
    # 檢查必要套件
    if not check_pyinstaller():
        print("❌ 無法安裝 PyInstaller")
        return False
    
    # PyInstaller 指令
    # --onefile: 打包成單一執行檔
    # --windowed: 不顯示命令列視窗（GUI 程式）
    # --name: 執行檔名稱
    # --icon: 圖示檔案（選用）
    
    cmd = [
        "pyinstaller",
        "--onefile",                    # 單一執行檔
        "--windowed",                   # GUI 模式（無命令列視窗）
        "--name=文字轉PPT工具",          # 執行檔名稱
        "--add-data=範例輸入文字.txt;.", # 包含範例檔案
        # "--icon=icon.ico",            # 如果有圖示檔案
        "text_to_ppt_gui.py"
    ]
    
    # Windows 上的語法稍有不同
    if sys.platform == 'win32':
        cmd[4] = "--add-data=範例輸入文字.txt;."
    else:
        cmd[4] = "--add-data=範例輸入文字.txt:."
    
    print("執行打包指令：")
    print(" ".join(cmd))
    print()
    
    try:
        result = subprocess.run(cmd, check=True)
        print()
        print("=" * 60)
        print("✅ 打包完成！")
        print("=" * 60)
        print()
        print("執行檔位置：dist/文字轉PPT工具.exe")
        print()
        print("請將以下檔案一起發布：")
        print("  • dist/文字轉PPT工具.exe")
        print("  • 範例輸入文字.txt（選用）")
        print("  • Windows使用說明.txt（選用）")
        print()
        return True
        
    except subprocess.CalledProcessError as e:
        print()
        print("❌ 打包失敗")
        print(f"錯誤：{e}")
        return False

def build_exe_simple():
    """簡易版本 - 僅打包命令列版本"""
    
    print("=" * 60)
    print("打包命令列版本")
    print("=" * 60)
    print()
    
    cmd = [
        "pyinstaller",
        "--onefile",
        "--name=text_to_ppt",
        "text_to_ppt.py"
    ]
    
    try:
        subprocess.run(cmd, check=True)
        print()
        print("✅ 命令列版本打包完成：dist/text_to_ppt.exe")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ 打包失敗：{e}")
        return False

if __name__ == "__main__":
    print()
    print("請選擇打包模式：")
    print("1. GUI 版本（圖形介面，推薦給一般使用者）")
    print("2. 命令列版本（適合進階使用者）")
    print("3. 兩者都打包")
    print()
    
    choice = input("請選擇 (1/2/3) [預設: 1]: ").strip() or "1"
    print()
    
    if choice == "1":
        build_exe()
    elif choice == "2":
        build_exe_simple()
    elif choice == "3":
        build_exe()
        print()
        build_exe_simple()
    else:
        print("❌ 無效的選擇")
