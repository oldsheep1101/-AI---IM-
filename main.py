"""
飞书机器人 - 接入 Agent（IM → 任务规划 → 执行 → Doc）
"""

import json
import os
import re
import time
import httpx
from threading import Lock

from dotenv import load_dotenv
import lark_oapi as lark
from lark_oapi.api.im.v1 import CreateMessageRequest, CreateMessageRequestBody
from lark_oapi.event.callback.model.p2_card_action_trigger import P2CardActionTrigger
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

# 卡片状态缓存（chat_id -> card_json），用于 PPT 完成时更新卡片
card_states: dict = {}
card_lock = Lock()

# Agent 状态（chat_id -> {"msg_id": ..., "card_msg_id": ..., "context": {}}）
agent_states: dict = {}
agent_lock = Lock()

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


def send_card(chat_id: str, card_json: dict) -> str:
    """发送飞书交互卡片，返回 message_id"""
    try:
        # 获取 tenant_access_token
        token_url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
        token_resp = httpx.post(token_url, json={
            "app_id": APP_ID,
            "app_secret": APP_SECRET
        }, verify=True, timeout=30.0)
        token_data = token_resp.json()
        token = token_data.get("tenant_access_token", "")
        if not token:
            print(f"[卡片] 获取 token 失败")
            return ""

        url = "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id"
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        payload = {
            "receive_id": chat_id,
            "msg_type": "interactive",
            "content": json.dumps(card_json)
        }
        resp = httpx.post(url, headers=headers, json=payload, verify=True, timeout=30.0)
        data = resp.json()
        if data.get("code") == 0:
            msg_id = data.get("data", {}).get("message_id", "")
            print(f"[卡片] 已发送, message_id={msg_id}")
            with card_lock:
                card_states[chat_id] = {"msg_id": msg_id, "card_json": card_json}
            return msg_id
        else:
            print(f"[卡片] 发送失败: {data}")
            return ""
    except Exception as e:
        print(f"[卡片] 异常: {e}")
        return ""


def update_card(message_id: str, card_json: dict) -> bool:
    """PATCH 更新飞书卡片"""
    if not message_id:
        return False
    try:
        # 先获取 tenant_access_token
        token_url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
        token_resp = httpx.post(token_url, json={
            "app_id": APP_ID,
            "app_secret": APP_SECRET
        }, verify=True, timeout=30.0)
        token_data = token_resp.json()
        token = token_data.get("tenant_access_token", "")
        if not token:
            print(f"[卡片] 获取 token 失败: {token_data}")
            return False

        url = f"https://open.feishu.cn/open-apis/im/v1/messages/{message_id}"
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        resp = httpx.patch(url, headers=headers, json={"content": json.dumps(card_json)}, verify=True, timeout=30.0)
        data = resp.json()
        code = data.get("code", data.get("StatusCode", -1))
        if code == 0:
            print(f"[卡片] 更新成功, message_id={message_id}")
            with card_lock:
                if message_id:
                    card_states[message_id] = {"msg_id": message_id, "card_json": card_json}
            return True
        else:
            print(f"[卡片] 更新失败: {data}")
            return False
    except Exception as e:
        print(f"[卡片] 更新异常: {e}")
        return False


def _finalize_ppt_card() -> None:
    """PPT 转发成功后，将主群卡片的步骤3更新为完成状态"""
    main_chat_id = "oc_28cf04fd87a5694667a7d807b70a3257"
    with card_lock:
        state = card_states.get(main_chat_id)
    if not state or not state.get("msg_id"):
        return
    card_json = {
        "schema": "2.0",
        "config": {"update_multi": True},
        "header": {
            "title": {"tag": "plain_text", "content": "✅ 任务完成"},
            "template": "green"
        },
        "body": {
            "elements": [
                {"tag": "markdown", "content": "**步骤1：抓取群聊聊天记录**\n<font color='green'>✅ 完成</font>", "margin": "8px 0px 4px 0px"},
                {"tag": "markdown", "content": "**步骤2：生成讨论总结文档**\n<font color='green'>✅ 完成</font>", "margin": "4px 0px 4px 0px"},
                {"tag": "markdown", "content": "**步骤3：PPT 制作与转发**\n<font color='green'>✅ 完成</font>", "margin": "4px 0px 8px 0px"},
                {"tag": "hr"},
                {"tag": "markdown", "content": "🤖 所有任务已完成", "margin": "8px 0px 0px 0px"}
            ]
        }
    }
    update_card(state["msg_id"], card_json)


# 飞书客户端（全局）
feishu_client: lark.Client = None


from lark_oapi.event.callback.model.p2_card_action_trigger import P2CardActionTrigger


def handle_card_action(event: P2CardActionTrigger) -> None:
    """处理卡片按钮点击事件"""
    import traceback
    try:
        trigger_data = event.event
        action = trigger_data.action
        context = trigger_data.context
        chat_id = context.open_chat_id if context else ""
        message_id = context.open_message_id if context else ""
        print(f"[CardAction] chat_id={chat_id}, message_id={message_id}")
        print(f"[CardAction] action object: {dir(action)}")
        print(f"[CardAction] action attrs: {[(k, getattr(action, k, None)) for k in dir(action) if not k.startswith('_')]}")

        # 解析按钮 value
        action_data = {}
        if hasattr(action, "value") and action.value:
            action_data = action.value
            if isinstance(action_data, str):
                action_data = json.loads(action_data)
        elif hasattr(action, "action_value") and action.action_value:
            action_data = action.action_value or {}
            if isinstance(action_data, str):
                action_data = json.loads(action_data)

        action_type = action_data.get("action", "")
        print(f"[CardAction] action_type={action_type}, action_data={action_data}")

        # 查找对应 chat_id 的 Agent 状态
        with agent_lock:
            state = agent_states.get(chat_id)
        if not state:
            print(f"[CardAction] 未找到 chat_id={chat_id} 的 Agent 状态，跳过")
            return

        tasks = state.get("tasks", [])
        if not tasks:
            print(f"[CardAction] tasks 为空，跳过")
            return

        # 构建 Agent 实例（复用 card 函数以便更新卡片）
        agent = Agent(
            feishu_client=feishu_client,
            chat_id=chat_id,
            send_card_func=send_card,
            update_card_func=update_card,
        )
        agent.card_msg_id = state.get("card_msg_id") or state.get("msg_id")
        agent.context = state.get("context", {})

        if action_type == "confirm_ppt":
            print("[CardAction] 用户确认大纲，继续生成 PPT")
            # 更新卡片为 PPT 制作中状态
            ppt_card = agent._build_ppt_card()
            if agent.card_msg_id:
                update_card(agent.card_msg_id, ppt_card)
            # 恢复执行
            results = agent.resume(tasks, confirmed=True)
            # 重新保存状态
            with agent_lock:
                agent_states[chat_id] = {
                    "msg_id": agent.card_msg_id,
                    "card_msg_id": agent.card_msg_id,
                    "context": agent.context,
                    "tasks": results,
                }

        elif action_type == "retry":
            print("[CardAction] 用户打回，等待手动修改文档")
            # 提示用户手动修改文档，修改完再点 confirm
            wait_card = {
                "schema": "2.0",
                "config": {"update_multi": True},
                "header": {
                    "title": {"tag": "plain_text", "content": "📝 请手动修改文档"},
                    "template": "orange"
                },
                "body": {
                    "elements": [
                        {"tag": "markdown", "content": "**请前往云文档手动修改内容**", "margin": "8px 0px 4px 0px"},
                        {"tag": "markdown", "content": f"文档链接：{agent.context.get('doc_link', '未知')}", "margin": "4px 0px 4px 0px"},
                        {"tag": "hr"},
                        {"tag": "markdown", "content": "修改完成后，点下方按钮继续生成 PPT", "margin": "4px 0px 8px 0px"},
                        {
                            "tag": "button",
                            "text": {"tag": "plain_text", "content": "✅ 确认大纲，生成 PPT"},
                            "type": "callback",
                            "value": {"action": "confirm_ppt", "doc_link": agent.context.get("doc_link", "")}
                        }
                    ]
                }
            }
            if agent.card_msg_id:
                update_card(agent.card_msg_id, wait_card)
            # 不再重置 context，保留 doc_link 等信息
            # 等待用户修改完点 confirm_ppt
            with agent_lock:
                agent_states[chat_id] = {
                    "msg_id": agent.card_msg_id,
                    "card_msg_id": agent.card_msg_id,
                    "context": agent.context,
                    "tasks": tasks,
                }
        else:
            print(f"[CardAction] 未知 action_type={action_type}")

    except Exception as e:
        print(f"[CardAction] 异常: {type(e).__name__}: {e}")
        traceback.print_exc()


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
                    # PPT 转发成功后，更新主群里那张 Agent 进度卡片的步骤3为完成
                    _finalize_ppt_card()
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
        update_card_func=update_card,
    )

    # === Human-in-the-loop 文字指令处理 ===
    # 当 Agent 处于等待确认状态时，接收 confirm/retry 文字命令
    user_text_lower = user_text.lower().strip()
    if user_text_lower in ("confirm", "retry"):
        print(f"[文字指令]收到 {user_text_lower}，查找待确认状态")
        with agent_lock:
            state = agent_states.get(chat_id)
        if state and state.get("context", {}).get("awaiting_confirm"):
            tasks = state.get("tasks", [])
            agent.card_msg_id = state.get("card_msg_id") or state.get("msg_id")
            agent.context = state.get("context", {})
            confirmed = (user_text_lower == "confirm")
            print(f"[文字指令] 执行 resume, confirmed={confirmed}")
            if confirmed:
                # 更新卡片为 PPT 制作中
                ppt_card = agent._build_ppt_card()
                if agent.card_msg_id:
                    update_card(agent.card_msg_id, ppt_card)
            else:
                # 打回：更新卡片为等待修改状态
                wait_card = {
                    "schema": "2.0",
                    "config": {"update_multi": True},
                    "header": {
                        "title": {"tag": "plain_text", "content": "📝 请手动修改文档"},
                        "template": "orange"
                    },
                    "body": {
                        "elements": [
                            {"tag": "markdown", "content": "**请前往云文档手动修改内容**", "margin": "8px 0px 4px 0px"},
                            {"tag": "markdown", "content": f"文档链接：{agent.context.get('doc_link', '未知')}", "margin": "4px 0px 4px 0px"},
                            {"tag": "hr"},
                            {"tag": "markdown", "content": "修改完成后，**直接回复 `confirm`** 继续生成 PPT", "margin": "4px 0px 8px 0px"},
                        ]
                    }
                }
                if agent.card_msg_id:
                    update_card(agent.card_msg_id, wait_card)
            results = agent.resume(tasks, confirmed=confirmed)
            with agent_lock:
                agent_states[chat_id] = {
                    "msg_id": agent.card_msg_id,
                    "card_msg_id": agent.card_msg_id,
                    "context": agent.context,
                    "tasks": results,
                }
            return

    # 运行 Agent：规划 + 执行
    try:
        results = agent.run(user_text)
        done = sum(1 for t in results if t.status == "done")
        failed = sum(1 for t in results if t.status == "failed")
        print(f"[Agent 执行完成] {done} 成功, {failed} 失败")

        # 保存 Agent 状态（供 card_action 回调查找）
        if results:
            with agent_lock:
                agent_states[chat_id] = {
                    "msg_id": agent.card_msg_id,
                    "card_msg_id": agent.card_msg_id,
                    "context": agent.context,
                    "tasks": results,
                }
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
        .register_p2_card_action_trigger(handle_card_action)
        .build()
    )

    # 建立 WebSocket 长连接
    ws_client = lark.ws.Client(APP_ID, APP_SECRET, event_handler=dispatcher)
    ws_client.start()


if __name__ == "__main__":
    main()
