#!/usr/bin/env python3
"""
飞书群聊消息实时监听器
- 使用 lark-cli event +subscribe WebSocket 长连接接收消息
- 将消息追加写入本地文件（txt + json）
"""

import subprocess
import json
import sys
from datetime import datetime
from pathlib import Path

OUTPUT_DIR = Path.home() / "feishu_messages"
OUTPUT_DIR.mkdir(exist_ok=True)

LOG_TXT = OUTPUT_DIR / "messages.txt"
LOG_JSON = OUTPUT_DIR / "messages.jsonl"


def log(msg: str):
    """打印带时间戳的日志"""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def append_to_file(text: str, filepath: Path):
    """追加文本到文件"""
    with open(filepath, "a", encoding="utf-8") as f:
        f.write(text + "\n")


def parse_and_save(line: str):
    """解析 NDJSON 行并写入文件"""
    try:
        event = json.loads(line)
    except json.JSONDecodeError:
        log(f"JSON 解析失败: {line[:100]}")
        return

    event_type = event.get("type", "unknown")
    log(f"收到事件: {event_type}")

    # 提取关键字段
    chat_id = event.get("chat_id", "")
    sender_id = event.get("sender_id", "")
    content = event.get("content", "")
    message_id = event.get("message_id", "")
    create_time = event.get("create_time", "")

    # 格式化时间
    try:
        ts = datetime.fromtimestamp(int(create_time) / 1000).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        ts = create_time

    # 写入 txt（人类可读格式）
    txt_line = f"[{ts}] [{chat_id}] {sender_id}: {content}"
    append_to_file(txt_line, LOG_TXT)

    # 写入 jsonl（完整数据）
    append_to_file(line, LOG_JSON)


def main():
    log("飞书消息监听器启动")
    log(f"输出目录: {OUTPUT_DIR}")
    log("按 Ctrl+C 停止")

    # 构建 lark-cli 命令
    cmd = [
        "lark-cli", "event", "+subscribe",
        "--event-types", "im.message.receive_v1",
        "--compact",
        "--quiet",
    ]

    log(f"执行: {' '.join(cmd)}")

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1,
        )
    except FileNotFoundError:
        log("错误: 未找到 lark-cli，请先安装: npm install -g @larksuite/cli")
        sys.exit(1)

    try:
        import os
        import select

        # 设置 stdout 为非阻塞模式
        import fcntl
        fd = proc.stdout.fileno()
        fl = fcntl.fcntl(fd, fcntl.F_GETFL)
        fcntl.fcntl(fd, fcntl.F_SETFL, fl | os.O_NONBLOCK)

        while True:
            # 检查进程是否退出
            if proc.poll() is not None:
                # 读取所有剩余输出
                remaining = proc.stdout.read()
                if remaining:
                    for line in remaining.splitlines():
                        if line.strip():
                            parse_and_save(line)
                break

            # 非阻塞读取
            try:
                line = proc.stdout.readline()
                if line:
                    line = line.strip()
                    if line:
                        parse_and_save(line)
            except (IOError, BlockingIOError):
                pass

            # 检查 stderr
            try:
                err = proc.stderr.readline()
                if err:
                    log(f"lark-cli: {err.strip()}")
            except (IOError, BlockingIOError):
                pass

            import time
            time.sleep(0.1)
    except KeyboardInterrupt:
        log("收到停止信号，退出...")
    finally:
        proc.terminate()
        proc.wait()
        log("监听器已停止")


if __name__ == "__main__":
    main()
