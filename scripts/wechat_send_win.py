#!/usr/bin/env python3
"""
微信批量发送助手 — Windows 版自动化核心
依赖: pip install uiautomation pyperclip pywin32
"""
from __future__ import annotations

import time
from pathlib import Path

import pyperclip
import uiautomation as auto

WECHAT_WIN_TITLE = "微信"


def check_wechat_running() -> bool:
    """检查微信窗口是否已打开"""
    win = auto.WindowControl(Name=WECHAT_WIN_TITLE, searchDepth=1)
    return win.Exists(0)


def activate_wechat() -> auto.WindowControl:
    """激活微信窗口并返回控制对象"""
    win = auto.WindowControl(Name=WECHAT_WIN_TITLE, searchDepth=1)
    if not win.Exists(3):
        raise RuntimeError("未找到微信窗口，请先打开并登录 PC 微信")
    win.SetActive()
    time.sleep(0.4)
    return win


def search_contact(contact_name: str):
    """搜索并打开联系人/群聊"""
    auto.SendKeys("{Ctrl}f")
    time.sleep(0.25)
    pyperclip.copy(contact_name)
    auto.SendKeys("{Ctrl}v")
    time.sleep(0.2)
    auto.SendKeys("{Enter}")
    time.sleep(0.5)


def send_text(text: str):
    """发送文字消息"""
    pyperclip.copy(text)
    auto.SendKeys("{Ctrl}v")
    time.sleep(0.15)
    auto.SendKeys("{Enter}")


def send_image(image_path: str):
    """发送图片（通过 Ctrl+V 粘贴）"""
    import subprocess
    import sys

    # 确保剪贴板里有图片
    import win32clipboard
    import win32con
    from PIL import Image

    img = Image.open(image_path)
    img.convert("RGB")

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_BITMAP, img)
    finally:
        win32clipboard.CloseClipboard()

    time.sleep(0.1)
    auto.SendKeys("{Ctrl}v")
    time.sleep(0.2)
    auto.SendKeys("{Enter}")


def send_text_with_image(text: str, image_path: str):
    """发送文字+图片"""
    pyperclip.copy(text)
    auto.SendKeys("{Ctrl}v")
    time.sleep(0.15)
    send_image(image_path)


def call_send(target: str, msg_type: str, text: str, image_path: str):
    """主调用入口（供外部调用）"""
    activate_wechat()
    search_contact(target)

    if msg_type == "文字":
        send_text(text)
    elif msg_type == "图片":
        send_image(image_path)
    elif msg_type == "文字+图片":
        send_text_with_image(text, image_path)
    else:
        raise ValueError(f"不支持的消息类型: {msg_type}")
