#!/usr/bin/env python3
"""
飞书群聊模拟器 - 根据项目总结生成测试用群聊记录
"""

import json
import os
import sys
from datetime import datetime, timedelta

# 默认群聊成员
DEFAULT_MEMBERS = [
    {"name": "张明(PM)", "id": "ou_pm_001", "id_type": "open_id"},
    {"name": "李强(DEV)", "id": "ou_dev_001", "id_type": "open_id"},
    {"name": "王芳(TEST)", "id": "ou_test_001", "id_type": "open_id"},
    {"name": "陈总(PMO)", "id": "ou_pmo_001", "id_type": "open_id"},
    {"name": "刘洋(UI)", "id": "ou_ui_001", "id_type": "open_id"},
]

BASE_TIME = 1777219200000  # 基准时间（毫秒时间戳，2026-04-27）


def parse_project_summary(summary: str) -> dict:
    """
    解析项目阶段总结，提取关键信息用于生成对话
    """
    result = {
        "project_name": "软件开发项目",
        "stage": "第一阶段",
        "goal": "需求分析与系统设计",
        "completed": ["用户需求收集", "用例分析", "系统架构设计"],
        "problems": ["部分用户需求不明确", "设计讨论延迟"],
        "solutions": ["两轮用户访谈", "明确需求优先级"],
        "next_steps": ["进入开发阶段", "重点模块优先实现"],
    }

    lines = summary.replace("；", ";").replace("：", ":").split("\n")
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # 按 ; 分割成多个字段
        segments = line.split(";")
        for seg in segments:
            seg = seg.strip()
            if not seg:
                continue
            if "遇到的问题" in seg or ("问题" in seg and "完成" not in seg and "阶段" not in seg):
                # 取最后一个 : 后的内容
                val = seg.split(":")[-1].strip()
                result["problems"] = [s.strip() for s in val.split("，") if s.strip()]
            elif "解决方案" in seg or "解决" in seg:
                val = seg.split(":")[-1].strip()
                result["solutions"] = [s.strip() for s in val.split("，") if s.strip()]
            elif "下一步计划" in seg or "下一步" in seg:
                val = seg.split(":")[-1].strip()
                result["next_steps"] = [s.strip() for s in val.split("，") if s.strip()]
            elif "完成情况" in seg or ("完成" in seg and "目标" not in seg):
                val = seg.split(":")[-1].strip()
                result["completed"] = [s.strip() for s in val.split("，") if s.strip()]
            elif "阶段目标" in seg or ("阶段" in seg and "目标" in seg):
                # 格式: "第一阶段:阶段目标:xxx"
                parts = seg.split(":")
                if len(parts) >= 3:
                    result["stage"] = parts[0].strip()
                    result["goal"] = parts[-1].strip()
                elif len(parts) == 2:
                    result["goal"] = parts[-1].strip()

    return result


def generate_dialogue(project_info: dict, n_rounds: int = 26) -> list[dict]:
    """生成多轮协商对话（默认26条，覆盖完整项目周期）"""
    messages = []
    msg_id = 1

    stage = project_info.get("stage", "第一阶段")
    goal = project_info.get("goal", "需求分析与系统设计")
    project_name = project_info.get("project_name", "软件开发项目")
    completed = project_info.get("completed", ["用户需求收集", "用例分析", "系统架构设计"])
    problems = project_info.get("problems", ["部分用户需求不明确", "设计讨论延迟"])
    solutions = project_info.get("solutions", ["两轮用户访谈", "明确需求优先级"])
    next_steps = project_info.get("next_steps", ["进入开发阶段", "重点模块优先实现"])

    templates = [
        # === 阶段一：需求评审启动 ===
        {"sender": "张明(PM)", "content": "各位早上好！{stage}{goal}，正式启动，请各方注意。"},
        {"sender": "李强(DEV)", "content": "收到，DEV这边准备好了。"},
        {"sender": "王芳(TEST)", "content": "TEST收到，了解需求后可以开始编写测试用例。"},
        {"sender": "刘洋(UI)", "content": "UI这边随时待命。"},
        # === 阶段二：需求收集讨论 ===
        {"sender": "张明(PM)", "content": "目前已完成：{completed}。大家看一下有没有遗漏。"},
        {"sender": "李强(DEV)", "content": "系统架构设计这块我们内部评审过了，方案可行。"},
        {"sender": "王芳(TEST)", "content": "用例分析覆盖了主要业务流程，测试范围基本确定。"},
        {"sender": "刘洋(UI)", "content": "但是有些交互细节还需要和PM确认。"},
        # === 阶段三：遇到问题 ===
        {"sender": "张明(PM)", "content": "不过过程中确实遇到了问题：{problems}。"},
        {"sender": "李强(DEV)", "content": "需求不明确确实影响进度，有些设计方案被迫延迟。"},
        {"sender": "陈总(PMO)", "content": "延迟了多少？Q3能上线吗？"},
        {"sender": "张明(PM)", "content": "目前还在可控范围，我们已经采取了措施。"},
        # === 阶段四：解决方案讨论 ===
        {"sender": "张明(PM)", "content": "解决方案：{solutions}。效果还不错。"},
        {"sender": "李强(DEV)", "content": "两轮访谈确实把优先级理清了，DEV这边可以放手干了。"},
        {"sender": "王芳(TEST)", "content": "需求明确后，测试用例编写效率会提高不少。"},
        {"sender": "刘洋(UI)", "content": "对，交互细节确认后，设计稿也能加速。"},
        # === 阶段五：进度同步 ===
        {"sender": "陈总(PMO)", "content": "目前进度正常吗？进入下阶段有什么风险？"},
        {"sender": "张明(PM)", "content": "整体可控，{next_steps}。"},
        {"sender": "李强(DEV)", "content": "DEV评估开发阶段，重点模块优先实现，方案可行。"},
        {"sender": "王芳(TEST)", "content": "TEST这边可以提前介入设计评审，尽早发现质量问题。"},
        # === 阶段六：资源协调 ===
        {"sender": "陈总(PMO)", "content": "开发阶段人手够吗？要不要加资源？"},
        {"sender": "李强(DEV)", "content": "目前5人团队基本够用，但如果需求继续增加可能需要调整。"},
        {"sender": "张明(PM)", "content": "需求范围已经冻结，不会再扩。"},
        {"sender": "刘洋(UI)", "content": "UI设计这边排期没问题，可以配合开发进度。"},
        # === 阶段七：阶段总结 ===
        {"sender": "张明(PM)", "content": "第一阶段总结：需求分析和设计完成，遇到的问题已解决。"},
        {"sender": "陈总(PMO)", "content": "好，下一阶段进入开发，重点模块优先。请各方保持沟通。"},
        {"sender": "李强(DEV)", "content": "DEV收到，下周一前出详细排期。"},
        {"sender": "王芳(TEST)", "content": "TEST这边同步开始准备测试环境。"},
        {"sender": "刘洋(UI)", "content": "UI这边设计稿周一可以全部交付。"},
        {"sender": "张明(PM)", "content": "收到，各方配合顺畅，感谢大家！"},
    ]

    # 截取需要的轮次
    used_templates = templates[:min(n_rounds, len(templates))]

    for i, tpl in enumerate(used_templates):
        content = tpl["content"].format(
            stage=stage,
            goal=goal,
            completed="、".join(completed) if completed else "用户需求收集、用例分析、系统架构设计",
            problems="、".join(problems) if problems else "部分需求不明确、设计讨论延迟",
            solutions="、".join(solutions) if solutions else "两轮用户访谈、明确优先级",
            next_steps="、".join(next_steps) if next_steps else "进入开发阶段、重点模块优先实现",
        )

        # 找到 sender 对应的成员信息
        sender_info = None
        for m in DEFAULT_MEMBERS:
            if m["name"].startswith(tpl["sender"].split("(")[0]):
                sender_info = m
                break

        if sender_info is None:
            sender_info = DEFAULT_MEMBERS[0]

        messages.append({
            "message_id": f"om_sim_{msg_id:03d}",
            "msg_type": "text",
            "sender": sender_info,
            "content": content,
            "create_time": str(BASE_TIME + i * 60000),  # 每条间隔1分钟
            "chat_id": f"oc_sim_{project_name}",
            "message_position": str(i + 1),
            "update_time": str(BASE_TIME + i * 60000),
            "deleted": False,
            "updated": False,
        })
        msg_id += 1

    return messages


def generate_chat_record(project_summary: str, output_path: str = None) -> dict:
    """
    主入口：根据项目总结生成完整群聊记录

    Args:
        project_summary: 项目总结文本
        output_path: 输出文件路径，默认 ~/Desktop/feishu-bot/test_data/simulated_chat.json
    """
    project_info = parse_project_summary(project_summary)

    # 固定生成 26 条消息，覆盖完整项目周期
    n_rounds = 26

    messages = generate_dialogue(project_info, n_rounds=n_rounds)

    result = {
        "group_name": f"{project_info['project_name']}需求评审群",
        "group_id": f"oc_sim_{project_info['project_name']}",
        "messages": messages,
    }

    if output_path is None:
        output_path = os.path.join(
            os.path.expanduser("~"),
            "Desktop", "feishu-bot", "test_data",
            "simulated_chat.json"
        )

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    return result


if __name__ == "__main__":
    # 示例用法
    if len(sys.argv) > 1:
        summary = sys.argv[1]
    else:
        # 默认示例：用户提供的项目阶段总结
        summary = (
            "第一阶段：阶段目标：完成需求分析与系统设计；"
            "完成情况：完成了用户需求收集、用例分析和系统架构设计；"
            "遇到的问题：部分用户需求不明确，导致设计讨论延迟；"
            "解决方案：安排了两轮用户访谈，明确需求优先级；"
            "下一步计划：进入开发阶段，重点模块优先实现"
        )

    result = generate_chat_record(summary)

    print(f"生成完成，共 {len(result['messages'])} 条消息")
    print(f"保存至：~/Desktop/feishu-bot/test_data/simulated_chat.json")
    print()
    print("预览（前10条）：")
    for msg in result["messages"][:10]:
        print(f"  [{msg['sender']['name']}] {msg['content']}")
    print()
    print("各角色发言统计：")
    stats = {}
    for msg in result["messages"]:
        name = msg["sender"]["name"]
        stats[name] = stats.get(name, 0) + 1
    for name, count in sorted(stats.items(), key=lambda x: -x[1]):
        print(f"  {name}: {count} 条")
