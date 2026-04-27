"""
Agent 模拟测试：跑通规划 + 执行全流程
"""

import os
from dotenv import load_dotenv
load_dotenv()  # 确保在 import agent 之前加载 .env

import json
from agent import Agent, TaskType


# ==================== 模拟飞书客户端 & 卡片发送 ====================

class MockFeishuClient:
    """模拟飞书客户端"""
    def im(self):
        return self
    def v1(self):
        return self
    def message(self):
        return self

    def create(self, *args, **kwargs):
        print(f"   [模拟飞书 API] create message: args={args}, kwargs={kwargs}")
        return MockResponse(code=0, data={"message_id": "mock_msg_123"})


class MockSendCard:
    """模拟发送状态卡片"""
    def __init__(self):
        self.card_msg_id = "mock_card_msg_id"

    def __call__(self, card_msg_id, content):
        print(f"\n   【状态卡片更新】")
        print(f"   {content}")
        print()


class MockResponse:
    """模拟 HTTP 响应"""
    def __init__(self, code=0, data=None):
        self.code = code
        self.data = data


# ==================== 创建 Agent 实例（Mock 模式） ====================

feishu_client = MockFeishuClient()
send_card = MockSendCard()

agent = Agent(feishu_client=feishu_client, chat_id="mock_chat_123", send_card_func=send_card)


# ==================== 模拟执行场景 ====================

if __name__ == "__main__":
    print("=" * 60)
    print("Agent 模拟测试：极光项目文档整理")
    print("=" * 60)

    user_input = "把刚才讨论的极光项目整理成文档发给我"

    print(f"\n[用户输入] {user_input}")
    print("\n[Phase 1] Planner 规划中...\n")

    tasks = agent.plan(user_input)

    print(f"[Planner 输出] 生成了 {len(tasks)} 个任务步骤:")
    for t in tasks:
        print(f"  步骤{t.step} [{t.type}] {t.desc}")
        print(f"           params: {t.params}")

    print(f"\n[Phase 2] Executor 执行中...\n")

    results = agent.execute(tasks)

    print("\n[执行完成] 最终结果:")
    for t in results:
        status_icon = {"done": "✅", "failed": "❌", "running": "⏳"}.get(t.status, "⚪")
        print(f"  {status_icon} 步骤{t.step} {t.desc}: {t.status}")
        if t.result:
            print(f"           结果: {t.result}")

    print("\n" + "=" * 60)
    print("模拟测试完成")
    print("=" * 60)
