#!/usr/bin/env python3
"""
ダンジョンと不思議の箱 - Python Edition
Dungeon and Mysterious Box - Python Refactored Version

オリジナルのVisual Basic 6版をPythonでリファクタリング

使い方:
    python main.py

操作方法:
    ← → ↑ ↓    : 移動
    Z           : 全体攻撃
    X           : HP全快
    C           : 全消去（箱、壁、モンスター）
    D           : モンスター除去
    Enter       : 次の階へ
    A           : 緑箱を全て開ける

ゲームの目的:
    1000階到達を目指してダンジョンを探索
"""
import sys
import os

# 必要なモジュールのインポートチェック
try:
    import tkinter as tk
except ImportError:
    print("Error: tkinter が見つかりません。")
    print("Pythonのtkinterモジュールをインストールしてください。")
    sys.exit(1)

# GUIを起動
from gui import main as run_gui


if __name__ == "__main__":
    print("=" * 60)
    print("  ダンジョンと不思議の箱 - Python Edition")
    print("  Dungeon and Mysterious Box - Refactored in Python")
    print("=" * 60)
    print()
    print("ゲームを起動中...")
    print()

    run_gui()
