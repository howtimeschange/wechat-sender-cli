#!/usr/bin/env python3
"""
微信监控 + 自动回复 — Windows 版（带风控）
功能：监听指定联系人，自动回复（可配合 Excel 批量发送）
依赖: pip install uiautomation pyperclip pywin32 Pillow
"""
from __future__ import annotations

import hashlib
import time
from collections import defaultdict
from datetime import datetime
from threading import Lock

import pyperclip
import uiautomation as auto

WECHAT_WIN_TITLE = "微信"

# ─── 风控配置 ────────────────────────────────────────────
KEYWORD_BLOCK = {"转账", "借钱", "合同", "发票", "付款", "汇款", "账户", "银行卡", "密码"}
MAX_PER_MINUTE = 3          # 同联系人每分钟最多 3 条
CIRCUIT_BREAK_THRESHOLD = 3  # 连续失败 3 次则熔断
WHITELIST = set()            # 白名单联系人（留空表示不限制）

# ─── 运行时状态 ──────────────────────────────────────────
last_sig = ""
last_sig_lock = Lock()

# {contact: [(timestamp, count)]}
rate_counter: dict[str, list[float]] = defaultdict(list)
rate_lock = Lock()

# 连续失败计数
fail_counts: dict[str, int] = defaultdict(int)
circuit_lock = Lock()

# 全局熔断
global_circuit_broken = False


def reset_sig():
    global last_sig
    with last_sig_lock:
        last_sig = ""


def changed(text: str) -> bool:
    global last_sig
    if not text.strip():
        return False
    sig = hashlib.md5(text.encode("utf-8")).hexdigest()
    with last_sig_lock:
        if sig == last_sig:
            return False
        last_sig = sig
        return True


def check_keyword(text: str) -> bool:
    """检测敏感关键词"""
    t = text.lower()
    return any(kw in t for kw in KEYWORD_BLOCK)


def check_rate(contact: str) -> bool:
    """检查并更新频率限制，返回是否允许发送"""
    now = time.time()
    with rate_lock:
        # 清理超过1分钟的记录
        rate_counter[contact] = [t for t in rate_counter[contact] if now - t < 60]
        if len(rate_counter[contact]) >= MAX_PER_MINUTE:
            return False
        rate_counter[contact].append(now)
        return True


def check_circuit(contact: str) -> bool:
    """检查熔断状态"""
    with circuit_lock:
        return fail_counts.get(contact, 0) < CIRCUIT_BREAK_THRESHOLD


def record_fail(contact: str):
    with circuit_lock:
        fail_counts[contact] = fail_counts.get(contact, 0) + 1


def record_success(contact: str):
    with circuit_lock:
        fail_counts[contact] = 0


def activate_wechat() -> auto.WindowControl:
    win = auto.WindowControl(Name=WECHAT_WIN_TITLE, searchDepth=1)
    if not win.Exists(3):
        raise RuntimeError("未找到微信窗口")
    win.SetActive()
    time.sleep(0.3)
    return win


def open_contact(contact_name: str):
    auto.SendKeys("{Ctrl}f")
    time.sleep(0.2)
    pyperclip.copy(contact_name)
    auto.SendKeys("{Ctrl}v")
    time.sleep(0.2)
    auto.SendKeys("{Enter}")
    time.sleep(0.4)


def send_text(text: str):
    pyperclip.copy(text)
    auto.SendKeys("{Ctrl}v")
    time.sleep(0.1)
    auto.SendKeys("{Enter}")


def read_last_message() -> str:
    """读取当前聊天区域最新消息（简化版：Ctrl+A Ctrl+C）"""
    try:
        auto.SendKeys("{Ctrl}a")
        time.sleep(0.05)
        auto.SendKeys("{Ctrl}c")
        time.sleep(0.1)
        txt = pyperclip.paste() or ""
        lines = [x.strip() for x in txt.splitlines() if x.strip()]
        return "\n".join(lines[-8:])
    except Exception:
        return ""


# ─── 自动回复主循环 ───────────────────────────────────────

def watch_loop(
    target_contact: str,
    reply_fn=None,
    poll_seconds: int = 2,
    dry_run: bool = False,
):
    """
    监控指定联系人，有新消息时调用 reply_fn 生成并发送回复

    reply_fn(user_text: str) -> str
    """
    global global_circuit_broken

    print(f"开始监听: {target_contact} | 轮询间隔: {poll_seconds}s")
    if dry_run:
        print("⚠️  模拟运行模式，不会真实发送")
    if KEYWORD_BLOCK:
        print(f"✅ 关键词拦截: {KEYWORD_BLOCK}")
    if WHITELIST:
        print(f"✅ 白名单: {WHITELIST}")

    activate_wechat()
    open_contact(target_contact)
    reset_sig()

    consecutive_errors = 0

    while True:
        try:
            tail = read_last_message()
            if tail and changed(tail):
                if check_keyword(tail):
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] ⛔ 关键词拦截，跳过")
                    continue

                # 白名单检查
                if WHITELIST and target_contact not in WHITELIST:
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] ⛔ 不在白名单，跳过")
                    continue

                # 熔断检查
                if not check_circuit(target_contact):
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] ⛔ 熔断触发，暂停监听")
                    time.sleep(30)
                    if check_circuit(target_contact):
                        print("✅ 熔断恢复")
                    continue

                # 频率检查
                if not check_rate(target_contact):
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] ⏳ 频率超限，等待...")
                    time.sleep(5)
                    continue

                reply = (reply_fn or default_reply)(tail)

                if not dry_run:
                    send_text(reply)

                record_success(target_contact)
                consecutive_errors = 0
                print(f"[{datetime.now().strftime('%H:%M:%S')}] ✅ 已回复: {reply[:50]}")

            consecutive_errors = 0
        except Exception as e:
            consecutive_errors += 1
            print(f"[{datetime.now().strftime('%H:%M:%S')}] ❌ 错误: {e}")
            if consecutive_errors >= CIRCUIT_BREAK_THRESHOLD:
                global_circuit_broken = True
                print("⚠️  连续错误，休息 60 秒...")
                time.sleep(60)
                global_circuit_broken = False
                consecutive_errors = 0

        time.sleep(poll_seconds)


def default_reply(user_text: str) -> str:
    return f"收到：{user_text[-30:]}（自动回复 {datetime.now().strftime('%H:%M:%S')}）"


# ─── CLI 入口 ─────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="微信监控自动回复 — Windows 版")
    parser.add_argument("--contact", default="文件传输助手", help="监控的联系人")
    parser.add_argument("--poll", type=int, default=2, help="轮询间隔（秒）")
    parser.add_argument("--dry", action="store_true", help="模拟运行")
    args = parser.parse_args()

    watch_loop(args.contact, poll_seconds=args.poll, dry_run=args.dry)
