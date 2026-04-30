"""
飞书机器人 - 接入 Agent（IM → 任务规划 → 执行 → Doc）
"""

import json
import os
import re
import time
from threading import Lock

from dotenv import load_dotenv
import lark_oapi as lark
from lark_oapi.api.im.v1 import CreateMessageRequest, CreateMessageRequestBody
from lark_oapi.event.dispatcher_handler import EventDispatcherHandler

from agent import Agent

# 加载 .env 环境变量
load_dotenv()

# ============== 配置 ==============
APP_ID = os.getenv("FEISHU_APP_ID", "YOUR_APP_ID")
APP_SECRET = os.getenv("FEISHU_APP_SECRET", "YOUR_APP_SECRET")
# ==================================


# 已处理消息 ID 集合（防止重复处理）
processed_msg_ids: set = set()
processed_lock = Lock()

# 机器人启动时间（Unix 毫秒时间戳）
bot_start_time: int = 0


def send_reply(chat_id: str, content: str) -> None:
    """发送回复到飞书群聊"""
    try:
        request = (
            CreateMessageRequest.builder()
            .receive_id_type("chat_id")
            .request_body(
                CreateMessageRequestBody.builder()
                .receive_id(chat_id)
                .msg_type("text")
                .content(json.dumps({"text": content}))
                .build()
            )
            .build()
        )
        feishu_client.request(request)
    except Exception as e:
        print(f"[发送失败] {e}")


def send_card(card_msg_id: str, content: str) -> None:
    """发送/更新状态卡片（目前用普通消息代替卡片）"""
    # TODO: 后续替换为飞书交互卡片 API
    lines = content.split("\n")
    for line in lines:
        if line.strip():
            print(f"   {line}")
    print()


# 飞书客户端（全局）
feishu_client: lark.Client = None


def handle_message(event: lark.im.v1.P2ImMessageReceiveV1) -> None:
    """处理接收到的消息事件"""
    global bot_start_time

    message = event.event.message
    if not message or not message.message_id:
        return

    # 过滤机器人启动前积压的离线消息
    if message.create_time and int(message.create_time) < bot_start_time:
        print(f"[跳过历史消息] create_time={message.create_time} < bot_start_time={bot_start_time}")
        return

    # 幂等处理
    msg_id = message.message_id
    with processed_lock:
        if msg_id in processed_msg_ids:
            print(f"[跳过重复消息] message_id={msg_id}")
            return
        processed_msg_ids.add(msg_id)
        if len(processed_msg_ids) > 1000:
            processed_msg_ids.clear()

    # 解析消息内容，统一提取 user_text
    user_text = ""
    chat_id = message.chat_id

    # 检测是否来自 OpenCLAW：私聊群 + sender_id 为空 = OpenCLAW 发来的消息
    sender_id = ""
    sender_obj = getattr(event.event, "sender", None)
    if sender_obj:
        sender_id_obj = getattr(sender_obj, "sender_id", None)
        if sender_id_obj is not None:
            # sender_id 可能是 UserId 对象或 dict
            if hasattr(sender_id_obj, "open_id"):
                sender_id = sender_id_obj.open_id or ""
            elif isinstance(sender_id_obj, dict):
                sender_id = sender_id_obj.get("open_id", "") or ""

    print(f"[DEBUG] chat_id={chat_id}, sender_id={repr(sender_id)}")

    # 私聊群 + 包含 PPT 相关内容 = 可能是 OpenCLAW 或用户转发的 PPT 消息
    if chat_id == "oc_ea67a09ec7edc0143ce7140b549635db":
        # 先提取消息内容
        msg_content = ""
        raw_content = message.content if message.content else ""
        if message.message_type == "text" and message.content:
            data = json.loads(message.content)
            msg_content = data.get("text", "").strip()
        elif message.message_type == "post" and message.content:
            data = json.loads(message.content)
            content_list = data.get("content", [])
            texts = []
            urls = []
            for row in content_list:
                for item in row:
                    if item.get("tag") == "text":
                        texts.append(item.get("text", ""))
                    elif item.get("tag") == "a":
                        texts.append(item.get("text", ""))
                        if item.get("href"):
                            urls.append(item.get("href"))
            msg_content = "".join(texts).strip()
            if urls:
                msg_content += " " + " ".join(urls)

        # 检查是否包含 PPT 完成消息
        print(f"[DEBUG] msg_type={message.message_type}, msg_content={repr(msg_content[:100])}")
        if ("PPT 已生成" in msg_content or "slides" in msg_content) and "feishu.cn/slides" in msg_content:
            urls = re.findall(r'https?://\S+', msg_content)
            ppt_url = urls[0] if urls else ""
            if ppt_url and "slides" in ppt_url:
                forward_content = f"PPT 已生成，请查收：{ppt_url}"
                print(f"[PPT 转发] 提取到 PPT 链接，转发到用户群: {forward_content}")
                try:
                    req = (
                        CreateMessageRequest.builder()
                        .receive_id_type("chat_id")
                        .request_body(
                            CreateMessageRequestBody.builder()
                            .receive_id("oc_28cf04fd87a5694667a7d807b70a3257")
                            .msg_type("text")
                            .content(json.dumps({"text": forward_content}))
                            .build()
                        )
                        .build()
                    )
                    feishu_client.request(req)
                    print("[PPT 转发] 成功")
                except Exception as e:
                    print(f"[PPT 转发] 失败: {e}")
            return

    if message.message_type == "text" and message.content:
        data = json.loads(message.content)
        user_text = data.get("text", "").strip()
        print(f"[DEBUG] 普通消息 sender_id={sender_id}, text={user_text[:50]}")
        if user_text.startswith("@_user_1"):
            user_text = user_text.replace("@_user_1", "", 1).strip()
        elif user_text.startswith("@雍和宫"):
            user_text = user_text.replace("@雍和宫", "", 1).strip()
        else:
            print(f"[忽略] 非艾特消息: {user_text}")
            return
    elif message.message_type == "post" and message.content:
        # 提取 post 消息中的纯文本和 URL
        data = json.loads(message.content)
        content_list = data.get("content", [])
        texts = []
        urls = []
        for row in content_list:
            for item in row:
                if item.get("tag") == "text":
                    texts.append(item.get("text", ""))
                elif item.get("tag") == "at":
                    texts.append(f"@{item.get('user_name', '')}")
                elif item.get("tag") == "a":
                    texts.append(item.get("text", ""))
                    if item.get("href"):
                        urls.append(item.get("href"))
        user_text = "".join(texts).strip()
        if urls:
            user_text += " " + " ".join(urls)
    else:
        return

    if not user_text:
        return

    print(f"[收到消息] {user_text}")

    # 如果是私聊群消息包含 PPT 链接，直接转发到用户群，不走 Agent
    if chat_id == "oc_ea67a09ec7edc0143ce7140b549635db" and "slides" in user_text:
        urls = re.findall(r'https?://\S+', user_text)
        ppt_url = urls[0] if urls else ""
        if ppt_url and "slides" in ppt_url:
            forward_content = f"PPT 已生成，请查收：{ppt_url}"
            print(f"[PPT 转发] 私聊群消息含 PPT 链接，转发到用户群: {forward_content}")
            try:
                req = (
                    CreateMessageRequest.builder()
                    .receive_id_type("chat_id")
                    .request_body(
                        CreateMessageRequestBody.builder()
                        .receive_id("oc_28cf04fd87a5694667a7d807b70a3257")
                        .msg_type("text")
                        .content(json.dumps({"text": forward_content}))
                        .build()
                    )
                    .build()
                )
                feishu_client.request(req)
                print("[PPT 转发] 成功")
            except Exception as e:
                print(f"[PPT 转发] 失败: {e}")
        return

    # 初始化 Agent（每次消息创建一个新的 Agent 实例，共享飞书客户端）
    agent = Agent(
        feishu_client=feishu_client,
        chat_id=chat_id,
        send_card_func=send_card,
    )

    # 运行 Agent：规划 + 执行
    try:
        results = agent.run(user_text)
        done = sum(1 for t in results if t.status == "done")
        failed = sum(1 for t in results if t.status == "failed")
        print(f"[Agent 执行完成] {done} 成功, {failed} 失败")
    except Exception as e:
        import traceback
        print(f"[Agent 执行异常] {type(e).__name__}: {e}")
        traceback.print_exc()
        send_reply(chat_id, f"Agent 执行出错：{e}")


def main():
    global feishu_client

    # 记录启动时间
    import time
    global bot_start_time
    bot_start_time = int(time.time() * 1000)

    # 初始化飞书客户端
    feishu_client = (
        lark.Client.builder()
        .app_id(APP_ID)
        .app_secret(APP_SECRET)
        .build()
    )

    # 告诉用户机器人已启动
    print(f"[机器人启动] {time.strftime('%Y-%m-%d %H:%M:%S')}")

    # 创建事件调度器
    dispatcher = (
        EventDispatcherHandler.builder(encrypt_key="", verification_token="")
        .register_p2_im_message_receive_v1(handle_message)
        .build()
    )

    # 建立 WebSocket 长连接
    ws_client = lark.ws.Client(APP_ID, APP_SECRET, event_handler=dispatcher)
    ws_client.start()


if __name__ == "__main__":
    main()
