#!/usr/bin/env python3
"""
飞书群消息定时拉取器
- 每分钟调用 lark-cli im +chat-messages-list 拉取新消息
- 支持富文本消息（图片、文件、post等）下载存储
"""

import os
import subprocess
import json
import sys
import time
import re
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict

# 配置
CHAT_ID = "oc_28cf04fd87a5694667a7d807b70a3257"
APP_ID = os.getenv("LARKBOT_APP_ID", "YOUR_APP_ID")
APP_SECRET = os.getenv("LARKBOT_APP_SECRET", "YOUR_APP_SECRET")
OUTPUT_DIR = Path.home() / "feishu_messages"
FILES_DIR = OUTPUT_DIR / "files"
STATE_FILE = OUTPUT_DIR / ".last_message_id"
LOG_FILE = OUTPUT_DIR / "fetch.log"

OUTPUT_DIR.mkdir(exist_ok=True)
FILES_DIR.mkdir(exist_ok=True)


def log(msg: str):
    """打印并记录日志"""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def get_last_message_id() -> Optional[str]:
    """读取上次处理的最新消息ID"""
    if STATE_FILE.exists():
        return STATE_FILE.read_text().strip()
    return None


def save_last_message_id(msg_id: str):
    """保存最新消息ID"""
    STATE_FILE.write_text(msg_id)


def fetch_messages(since_id: Optional[str] = None) -> list:
    """调用 lark-cli 拉取消息"""
    cmd = [
        "/opt/homebrew/bin/lark-cli", "im", "+chat-messages-list",
        "--chat-id", CHAT_ID,
        "--sort", "desc",
        "--page-size", "50",
        "--as", "bot",
    ]

    # 传递凭证环境变量
    env = {
        **os.environ,
        "LARKBOT_APP_ID": os.getenv("LARKBOT_APP_ID", ""),
        "LARKBOT_APP_SECRET": os.getenv("LARKBOT_APP_SECRET", ""),
    }

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True,
            env=env,
        )
        data = json.loads(result.stdout)
        if not data.get("ok"):
            log(f"API 返回错误: {data}")
            return []
        return data.get("data", {}).get("messages", [])
    except subprocess.CalledProcessError as e:
        log(f"执行失败: {e.stderr}")
        return []
    except json.JSONDecodeError as e:
        log(f"JSON 解析失败: {e}")
        return []


def get_tenant_access_token() -> Optional[str]:
    """获取 tenant_access_token"""
    import requests
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    data = {"app_id": APP_ID, "app_secret": APP_SECRET}
    try:
        resp = requests.post(url, json=data, timeout=10)
        result = resp.json()
        if result.get("code") == 0:
            return result.get("tenant_access_token")
        log(f"获取 token 失败: {result}")
    except Exception as e:
        log(f"获取 token 异常: {e}")
    return None


def download_file(file_key: str, msg_id: str, msg_type: str, token: str, ext: str = None) -> Optional[str]:
    """下载图片或文件到本地，返回本地路径"""
    # file_key 格式: file_v3_0011a_xxx.png 或 img_v3_xxx
    # 优先使用传入的 ext，否则从 file_key 或 file_name 提取
    if ext is None:
        parts = file_key.split(".")
        ext = parts[-1] if len(parts) > 1 else ("png" if msg_type == "image" else "bin")

    local_name = f"{msg_id}.{ext}"
    local_path = FILES_DIR / local_name

    if local_path.exists():
        return str(local_path)

    # 使用 lark-cli 下载（output 必须是相对路径）
    cmd = [
        "/opt/homebrew/bin/lark-cli", "im", "+messages-resources-download",
        "--message-id", msg_id,
        "--file-key", file_key,
        "--type", msg_type,
        "--output", local_name,  # 相对路径
        "--as", "bot",
    ]
    env = {**os.environ, "LARKBOT_APP_ID": APP_ID, "LARKBOT_APP_SECRET": APP_SECRET}

    try:
        # 在正确目录执行
        result = subprocess.run(cmd, capture_output=True, text=True, check=True, env=env, cwd=str(FILES_DIR), timeout=30)
        data = json.loads(result.stdout)
        if data.get("ok"):
            saved = data.get("data", {}).get("saved_path")
            if saved:
                saved_path = Path(saved)
                if not local_path.exists():
                    if saved_path.exists() and saved_path.name == local_name:
                        # 已经在正确位置
                        pass
                    else:
                        import shutil
                        shutil.move(str(saved_path), local_path)
                return str(local_path)
    except Exception as e:
        log(f"下载文件失败: {e}")

    return None


def parse_rich_content(msg: Dict) -> Dict:
    """解析富文本消息内容，返回可读文本和附件信息"""
    msg_type = msg.get("msg_type", "text")
    content = msg.get("content", "")
    msg_id = msg.get("message_id", "")
    token = get_tenant_access_token()

    result = {"text": "", "attachments": []}

    try:
        if msg_type == "text":
            # 文本消息的 content 可能是纯文本或 JSON
            try:
                obj = json.loads(content)
                result["text"] = obj.get("text", content) if isinstance(obj, dict) else content
            except json.JSONDecodeError:
                result["text"] = content  # 纯文本直接使用

        elif msg_type == "image":
            obj = json.loads(content)
            file_key = obj.get("file_key", "")
            # 从 file_key 提取扩展名
            img_ext = file_key.split(".")[-1] if "." in file_key else "png"
            if file_key and token:
                local_path = download_file(file_key, msg_id, "image", token, img_ext)
                if local_path:
                    result["attachments"].append({
                        "type": "image",
                        "local_path": local_path,
                        "file_key": file_key
                    })
            result["text"] = f"[图片: {file_key}]"

        elif msg_type == "file":
            # 文件内容是 XML 格式: <file key="xxx" name="xxx.png"/>
            import re
            key_match = re.search(r'key="([^"]+)"', content)
            name_match = re.search(r'name="([^"]+)"', content)
            file_key = key_match.group(1) if key_match else ""
            file_name = name_match.group(1) if name_match else "unknown"

            # 从 file_name 提取扩展名
            ext = file_name.split(".")[-1] if "." in file_name else "bin"

            if file_key and token:
                local_path = download_file(file_key, msg_id, "file", token, ext)
                if local_path:
                    result["attachments"].append({
                        "type": "file",
                        "local_path": local_path,
                        "file_key": file_key,
                        "file_name": file_name
                    })
            result["text"] = f"[文件: {file_name}]"

        elif msg_type == "post":
            obj = json.loads(content)
            text_parts = []
            for section in obj.get("zh_cn", {}).get("content", []):
                for item in section:
                    if item.get("tag") == "text":
                        text_parts.append(item.get("text", ""))
                    elif item.get("tag") == "at":
                        text_parts.append(f"@{item.get('user_name', '')}")
            result["text"] = "".join(text_parts)

        elif msg_type == "card":
            result["text"] = f"[卡片消息]"

        else:
            result["text"] = f"[{msg_type}消息]"

    except Exception as e:
        result["text"] = f"[解析失败: {content[:50]}]"

    return result


def save_messages(messages: list):
    """保存消息到文件（含富文本附件信息）"""
    txt_path = OUTPUT_DIR / "messages.txt"
    jsonl_path = OUTPUT_DIR / "messages.jsonl"

    with open(jsonl_path, "a", encoding="utf-8") as f:
        for msg in messages:
            # 解析富文本
            parsed = parse_rich_content(msg)
            msg["parsed_text"] = parsed["text"]
            msg["attachments"] = parsed["attachments"]
            f.write(json.dumps(msg, ensure_ascii=False) + "\n")

    # 格式化写入 txt
    lines = []
    for msg in messages:
        sender = msg.get("sender", {})
        sender_name = sender.get("id", "unknown")
        content = msg.get("parsed_text", msg.get("content", ""))
        ct = msg.get("create_time", "")

        # 格式化时间
        try:
            dt = datetime.fromisoformat(ct.replace("Z", "+00:00"))
            ts = dt.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            ts = ct

        # 附件信息
        attachments = msg.get("attachments", [])
        attach_info = ""
        if attachments:
            attach_info = " " + " ".join([f"[{a['type']}]" for a in attachments])

        lines.append(f"[{ts}] [{CHAT_ID}] {sender_name}: {content}{attach_info}")

    if lines:
        with open(txt_path, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")


def main():
    log("飞书消息拉取器启动")
    log(f"监听群组: {CHAT_ID}")

    # 读取上次位置
    last_id = get_last_message_id()
    if last_id:
        log(f"从上次消息之后开始: {last_id}")

    # 拉取消息
    messages = fetch_messages()
    log(f"本次拉取到 {len(messages)} 条消息")

    if not messages:
        return

    # 找出新消息（比 last_id 更新的）
    new_messages = []
    latest_id = None

    for msg in messages:
        msg_id = msg.get("message_id")
        if msg_id == last_id:
            break
        new_messages.append(msg)
        if latest_id is None:
            latest_id = msg_id

    if new_messages:
        # 倒序（按时间从旧到新）
        new_messages.reverse()
        log(f"新增 {len(new_messages)} 条消息")
        save_messages(new_messages)

        # 更新位置
        if latest_id:
            save_last_message_id(latest_id)
            log(f"已更新锚点: {latest_id}")
    else:
        log("没有新消息")

    log("本次拉取完成")


if __name__ == "__main__":
    main()
