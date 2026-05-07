"""
Agent 核心模块：任务规划 + 执行调度
"""

import datetime
import json
import os
import re
import time
import httpx
from enum import Enum
from typing import Optional

from openai import OpenAI
from lark_oapi.api.im.v1 import (
    CreateMessageRequest, CreateMessageRequestBody,
    ListMessageRequest,
)
from lark_oapi.api.docx.v1 import (
    CreateDocumentRequest, CreateDocumentRequestBody,
    CreateDocumentBlockChildrenRequest, CreateDocumentBlockChildrenRequestBody,
)


# ============== MiniMax 客户端（懒加载） ==============
_minimax_client = None

def _get_minimax_client() -> OpenAI:
    global _minimax_client
    if _minimax_client is None:
        _minimax_client = OpenAI(
            api_key=os.getenv("MINIMAX_API_KEY", ""),
            base_url="https://api.minimax.chat/v1",
        )
    return _minimax_client


class TaskType(str, Enum):
    """支持的原子任务类型"""
    RESEARCH = "RESEARCH"      # 读取聊天记录
    DOC = "DOC"                # 创建飞书文档
    BITABLE = "BITABLE"        # 创建多维表格/看板
    PPT = "PPT"                # 创建演示稿
    REPORT = "REPORT"          # 向用户汇报


class Task:
    """单个任务步骤"""
    def __init__(
        self,
        step: int,
        task_type: str,
        action: str,
        desc: str,
        params: Optional[dict] = None,
    ):
        self.step = step
        self.type = task_type
        self.action = action
        self.desc = desc
        self.params = params or {}
        self.status = "pending"  # pending / running / done / failed
        self.result = None

    def to_dict(self) -> dict:
        return {
            "step": self.step,
            "type": self.type,
            "action": self.action,
            "desc": self.desc,
            "status": self.status,
            **self.params,
        }


class Agent:
    """
    AI Agent：Planner（规划器）+ Executor（执行器）
    """

    # Planner 的 System Prompt
    PLANNER_PROMPT = """你是一个智能办公任务调度中心（AI Agent）。

你的核心能力是：将用户的自然语言需求，拆解成【有序的任务清单】，并驱动飞书套件完成办公自动化。

你拥有以下原子技能（每个 skill 对应一个可执行的函数）：

| skill       | 描述                     | 参数                    |
|-------------|--------------------------|-------------------------|
| RESEARCH    | 读取群聊上下文/历史记录   | query: 搜索关键词, date_range: 时间范围（如"今天"、"近3天"、"本周"、"上周"） |
| DOC         | 创建并编辑飞书云文档      | title: 文档标题, content: 初始内容 |
| BITABLE     | 创建飞书多维表格（看板）  | title: 表格标题, fields: 字段列表 |
| REPORT      | 向用户发送文字汇报        | content: 汇报文本, need_ppt: 是否需要制作PPT |

【意图识别规则】
- 用户说"总结"、"整理"、"汇总"群聊记录 → 必须 RESEARCH + DOC + REPORT，need_ppt=false
- 用户只说"整理成文档"、"生成文档"、"写成文档" → REPORT 的 need_ppt=false，不联系 OpenCLAW
- 用户说"做成 PPT"、"生成 PPT"、"做汇报"、"汇报 PPT" → REPORT 的 need_ppt=true，需要联系 OpenCLAW 制作
- 如果不确定，就默认 need_ppt=false（不制作 PPT）

【输出格式强制要求】
当你收到用户需求时，你必须：
1. 先在大脑里分析：用户需要什么？需要哪些步骤？
2. 只输出一个合法的 JSON 数组（不要任何其他文字），格式如下：

[
  {"step": 1, "type": "RESEARCH", "action": "research", "desc": "抓取群聊中关于xxx的讨论", "params": {"query": "xxx"}},
  {"step": 2, "type": "DOC", "action": "create_doc", "desc": "生成总结文档v1", "params": {"title": "xxx", "content": "..."}},
  ...
]

【规则】
- type 必须是 TaskType 中的值（RESEARCH/DOC/BITABLE/REPORT）
- step 从 1 开始，按顺序执行
- desc 用简短中文描述这个步骤在干啥
- 如果用户需求只需要 1 个步骤，也必须输出数组（如只需汇报）
- 不要输出任何解释性文字，只输出纯 JSON

【示例】
用户说："把刚才讨论的极光项目整理成文档发给我"
输出（不需要 PPT）：
[
  {"step": 1, "type": "RESEARCH", "action": "research", "desc": "抓取群聊中关于极光项目的讨论", "params": {"query": "极光项目", "date_range": ""}},
  {"step": 2, "type": "DOC", "action": "create_doc", "desc": "生成极光项目方案文档", "params": {"title": "极光项目方案V1", "content": ""}},
  {"step": 3, "type": "REPORT", "action": "report", "desc": "向用户发送完成汇报", "params": {"content": "文档已生成：xxx", "need_ppt": false}}
]

用户说："把这一周的项目讨论整理成文档发给我"
输出（本周记录，不需要 PPT）：
[
  {"step": 1, "type": "RESEARCH", "action": "research", "desc": "抓取本周群聊中关于项目的讨论", "params": {"query": "项目", "date_range": "本周"}},
  {"step": 2, "type": "DOC", "action": "create_doc", "desc": "生成项目周报文档", "params": {"title": "项目周报", "content": ""}},
  {"step": 3, "type": "REPORT", "action": "report", "desc": "向用户发送完成汇报", "params": {"content": "文档已生成：xxx", "need_ppt": false}}
]

用户说："把这一天的聊天记录做成 PPT 发给我"
输出（今天记录，需要 PPT）：
[
  {"step": 1, "type": "RESEARCH", "action": "research", "desc": "抓取今天群聊讨论", "params": {"query": "", "date_range": "今天"}},
  {"step": 2, "type": "DOC", "action": "create_doc", "desc": "生成今日讨论文档", "params": {"title": "今日讨论总结", "content": ""}},
  {"step": 3, "type": "REPORT", "action": "report", "desc": "向用户发送完成汇报", "params": {"content": "文档已生成：xxx", "need_ppt": true}}
]
"""

    def __init__(self, feishu_client, chat_id: str, send_card_func=None, update_card_func=None):
        """
        feishu_client: 飞书客户端实例
        chat_id: 当前群聊 ID（用于发消息和搜消息）
        send_card_func: 发送卡片的函数(chat_id, card_json) -> message_id
        update_card_func: 更新卡片的函数(message_id, card_json) -> bool
        """
        self.feishu_client = feishu_client
        self.chat_id = chat_id
        self.send_card_func = send_card_func
        self.update_card_func = update_card_func
        self.card_msg_id: Optional[str] = None
        # 跨步骤共享上下文（DOC 创建后把文档链接写这里，REPORT 读这里）
        self.context: dict = {}

    # ==================== Planner ====================

    def _resolve_sender_name(self, sender_obj: dict, feishu_client=None) -> str:
        """从 sender 字典直接取 name，机器人或无 name 时返回 id 尾部"""
        if not sender_obj or not isinstance(sender_obj, dict):
            return "未知用户"
        name = sender_obj.get("name", "")
        if name:
            return name
        sender_type = sender_obj.get("sender_type", "")
        sender_id_type = sender_obj.get("id_type", "")
        sender_id = sender_obj.get("id", "")
        if sender_type == "app" or sender_id_type == "app_id":
            return f"[BOT] {sender_id.split('_')[-1]}" if sender_id else "[BOT]"
        return sender_id.split("_")[-1] if sender_id else "未知用户"

    def plan(self, user_input: str) -> list[Task]:
        """调用 MiniMax 生成任务清单"""
        response = _get_minimax_client().chat.completions.create(
            model="MiniMax-M2.7",
            messages=[
                {"role": "system", "content": self.PLANNER_PROMPT},
                {"role": "user", "content": user_input},
            ],
            max_tokens=1024,
            temperature=0.3,
        )
        raw = response.choices[0].message.content.strip()
        # 过滤 thinking 标签
        raw = re.sub(r"<think>[\s\S]*?</think>", "", raw).strip()

        # 尝试提取 JSON
        tasks = self._parse_json(raw)
        if not tasks:
            # 容错：直接尝试解析整段
            tasks = self._parse_json(raw)
        normalized = []
        for t in tasks:
            # JSON 用 type，Task 用 task_type
            if "type" in t:
                t["task_type"] = t.pop("type")
            normalized.append(Task(**t))
        return normalized

    def _parse_json(self, raw: str) -> list:
        """从字符串中提取 JSON 数组"""
        # 优先尝试用 ```json 包裹的内容
        match = re.search(r"```(?:json)?\s*(\[[\s\S]*?\])\s*```", raw)
        if match:
            try:
                return json.loads(match.group(1))
            except json.JSONDecodeError:
                pass
        # 其次尝试直接解析
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            # 兜底：找第一个 [ 到最后一个 ] 之间的内容
            start = raw.find("[")
            end = raw.rfind("]")
            if start != -1 and end != -1:
                try:
                    return json.loads(raw[start:end + 1])
                except json.JSONDecodeError:
                    pass
        return []

    # ==================== Executor ====================

    def execute(self, tasks: list[Task]) -> list[Task]:
        """顺序执行任务清单"""
        # 发送初始状态卡片
        self._update_card(tasks)

        for task in tasks:
            task.status = "running"
            self._update_card(tasks)

            try:
                result = self._execute_task(task)
                # 如果 need_ppt=true，REPORT 步骤保持 running，等 PPT 转发回来才标记完成
                if task.type == TaskType.REPORT.value and self.context.get("ppt_pending"):
                    task.status = "running"
                    task.result = result
                else:
                    task.status = "done"
                    task.result = result
            except Exception as e:
                task.status = "failed"
                task.result = str(e)
                # 后续步骤标记为取消
                for t in tasks[tasks.index(task) + 1:]:
                    t.status = "cancelled"

            self._update_card(tasks)

            # === Human-in-the-loop：DOC 步骤完成后、REPORT 执行前 ===
            # 发确认卡片，等用户点按钮才继续（通过 WebSocket card_action 事件触发 resume）
            # 判断条件：DOC 是当前步骤 且 下一个步骤是 REPORT 且 need_ppt=true
            if task.type == TaskType.DOC.value:
                idx = tasks.index(task)
                next_task = tasks[idx + 1] if idx + 1 < len(tasks) else None
                need_ppt = next_task and next_task.type == TaskType.REPORT.value and next_task.params.get("need_ppt")
                if need_ppt:
                    self._send_confirm_card(tasks)
                    self.context["awaiting_confirm"] = True
                    self.context["pending_tasks"] = tasks
                    print("[Agent] 等待用户确认大纲...")
                    return tasks

        return tasks

    def resume(self, tasks: list[Task], confirmed: bool) -> list[Task]:
        """从确认卡片处恢复执行（用户点击按钮后由 main.py 调用）"""
        print(f"[Agent] resume called, confirmed={confirmed}")
        self.context["awaiting_confirm"] = False

        if not confirmed:
            # 打回场景已改为用户手动修改文档，此分支不再需要
            # 保持原有任务状态不变，等待用户修改完成后点 confirm
            print("[Agent] 用户打回，等待手动修改文档")
            return tasks

        # 确认后执行 REPORT（已在列表中，跳过已完成的步骤）
        for task in tasks:
            if task.status in ("done", "failed", "cancelled"):
                continue
            task.status = "running"
            self._update_card(tasks)
            try:
                result = self._execute_task(task)
                if task.type == TaskType.REPORT.value and self.context.get("ppt_pending"):
                    task.status = "running"
                    task.result = result
                else:
                    task.status = "done"
                    task.result = result
            except Exception as e:
                task.status = "failed"
                task.result = str(e)
            self._update_card(tasks)

        return tasks

    def _execute_task(self, task: Task) -> str:
        """根据 task 类型分发执行"""
        if task.type == TaskType.RESEARCH.value:
            return self._do_research(task)
        elif task.type == TaskType.DOC.value:
            return self._do_create_doc(task)
        elif task.type == TaskType.BITABLE.value:
            return self._do_create_bitable(task)
        elif task.type == TaskType.PPT.value:
            return self._do_create_ppt(task)
        elif task.type == TaskType.REPORT.value:
            return self._do_report(task)
        else:
            return f"未知任务类型: {task.type}"

    def _parse_relative_date(self, text: str) -> tuple:
        """
        解析中文相对日期，返回 (start_ts_ms, end_ts_ms)。
        支持：今天、昨天、近N天、最近N天、本周、上周
        """
        import datetime
        now = datetime.datetime.now()
        today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        start_ts = None
        end_ts = int(now.timestamp() * 1000)

        text = text.strip()

        if text in ("今天", "今日"):
            start_ts = int(today_start.timestamp() * 1000)
        elif text in ("昨天", "昨日"):
            yesterday = today_start - datetime.timedelta(days=1)
            start_ts = int(yesterday.timestamp() * 1000)
        elif text in ("本周", "这周"):
            weekday = today_start.weekday()
            start_ts = int((today_start - datetime.timedelta(days=weekday)).timestamp() * 1000)
        elif text in ("上周", "上个星期"):
            weekday = today_start.weekday()
            this_week_start = today_start - datetime.timedelta(days=weekday)
            start_ts = int((this_week_start - datetime.timedelta(weeks=1)).timestamp() * 1000)
        elif text in ("本月", "这个月"):
            start_ts = int(today_start.replace(day=1).timestamp() * 1000)
        elif text in ("上月", "上个月"):
            first_day_this_month = today_start.replace(day=1)
            start_ts = int((first_day_this_month - datetime.timedelta(days=1)).replace(day=1).timestamp() * 1000)
        else:
            m = re.search(r"近([0-9]+)天", text)
            if m:
                days = int(m.group(1))
                start_ts = int((today_start - datetime.timedelta(days=days - 1)).timestamp() * 1000)
            else:
                m = re.search(r"最近([0-9]+)天", text)
                if m:
                    days = int(m.group(1))
                    start_ts = int((today_start - datetime.timedelta(days=days - 1)).timestamp() * 1000)

        if start_ts is None:
            return None, None
        return start_ts, end_ts

    def _do_research(self, task: Task) -> str:
        """读取本地聊天记录文件（messages.jsonl），支持图片/文件等多媒体，按时间筛选后用 LLM 总结"""
        msg_file = "/Users/phoenix_oldsheep/feishu_messages/messages.jsonl"
        files_dir = "/Users/phoenix_oldsheep/feishu_messages/files"
        query = task.params.get("query", "")
        date_range = task.params.get("date_range", "")  # 如"今天"、"近3天"、"本周"

        # 解析时间范围
        start_ts, end_ts = self._parse_relative_date(date_range)
        print(f"[RESEARCH] 读取文件: {msg_file}, date_range={date_range}, start={start_ts}, end={end_ts}")

        try:
            with open(msg_file, "r", encoding="utf-8") as f:
                lines = f.readlines()
            print(f"[RESEARCH] 文件读取成功，共 {len(lines)} 条消息")
        except FileNotFoundError:
            return f"未找到聊天记录文件：{msg_file}"
        except Exception as e:
            return f"读取文件失败: {e}"

        # 按时间筛选并组装内容
        messages_for_summary = []
        image_files: dict = {}  # filename -> local_path，用于后续文档嵌入
        for line in lines:
            line = line.strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except json.JSONDecodeError:
                continue

            # 时间筛选
            create_time_str = msg.get("create_time", "")
            # 格式可能是 "2026-05-02 14:49" 或时间戳字符串
            line_ts = None
            try:
                if create_time_str.isdigit():
                    line_ts = int(create_time_str)
                else:
                    dt = datetime.datetime.strptime(create_time_str, "%Y-%m-%d %H:%M")
                    line_ts = int(dt.timestamp() * 1000)
            except Exception:
                pass

            if start_ts and line_ts and line_ts < start_ts:
                continue
            if end_ts and line_ts and line_ts > end_ts:
                continue

            msg_type = msg.get("msg_type", msg.get("message_type", "text"))
            parsed_text = msg.get("parsed_text", "") or msg.get("content", "")
            sender_obj = msg.get("sender", {})
            sender = self._resolve_sender_name(sender_obj, self.feishu_client)

            attachments = msg.get("attachments", [])
            extra_content = ""
            for att in attachments:
                att_type = att.get("type", "")
                local_path = att.get("local_path", "")
                file_name = att.get("file_name", "")

                if not local_path or not os.path.exists(local_path):
                    if file_name:
                        extra_content += f" [附件: {file_name}]"
                    continue

                ext = local_path.rsplit(".", 1)[-1].lower()

                if ext in ("png", "jpg", "jpeg", "gif", "webp", "bmp"):
                    # 图片：OCR 识别文字，同时收集文件路径供后续嵌入文档用
                    ocr_text = ""
                    try:
                        import subprocess
                        result = subprocess.run(
                            ["tesseract", local_path, "stdout", "--psm", "6", "-l", "chi_sim+eng"],
                            capture_output=True, timeout=30
                        )
                        ocr_text = result.stdout.decode("utf-8").strip()
                    except Exception as e:
                        pass
                    print(f"[DEBUG] OCR完成 file_name={file_name}, ocr_text长度={len(ocr_text)}, local_path={local_path}")
                    image_files[file_name] = local_path  # key 用 file_name，方便后续匹配 LLM 输出
                    if ocr_text:
                        extra_content += f"[图片文件: {file_name}] [图片内容: {ocr_text}]"
                    else:
                        extra_content += f"[图片文件: {file_name}]"

                elif ext == "pdf":
                    # PDF：提取文本
                    try:
                        from pypdf import PdfReader
                        reader = PdfReader(local_path)
                        pdf_texts = []
                        for page in reader.pages:
                            text = page.extract_text()
                            if text:
                                pdf_texts.append(text.strip())
                        if pdf_texts:
                            extra_content += f" [PDF内容: {' '.join(pdf_texts)}]"
                        else:
                            extra_content += f" [文件: {file_name}]"
                    except Exception as e:
                        extra_content += f" [文件: {file_name}]"

                elif ext in ("docx", "doc"):
                    # Word 文档：提取段落文本
                    try:
                        from docx import Document
                        doc = Document(local_path)
                        texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
                        if texts:
                            extra_content += f" [Word内容: {' '.join(texts)}]"
                        else:
                            extra_content += f" [文件: {file_name}]"
                    except Exception as e:
                        extra_content += f" [文件: {file_name}]"

                elif ext in ("xlsx", "xls"):
                    # Excel：读取单元格内容
                    try:
                        import openpyxl
                        wb = openpyxl.load_workbook(local_path, data_only=True)
                        sheet_texts = []
                        for sheet in wb.worksheets:
                            rows = list(sheet.iter_rows(values_only=True))
                            for row in rows:
                                row_text = " | ".join(str(c) if c is not None else "" for c in row)
                                if row_text.strip():
                                    sheet_texts.append(row_text)
                        print(f"[DEBUG] Excel read: {file_name}, rows={len(sheet_texts)}")
                        if sheet_texts:
                            extra_content += f" [Excel内容: {' ; '.join(sheet_texts)}]"
                        else:
                            extra_content += f" [文件: {file_name}]"
                    except Exception as e:
                        print(f"[DEBUG] Excel read error: {e}")
                        extra_content += f" [文件: {file_name}]"

                else:
                    extra_content += f" [文件: {file_name}]"

            if msg_type == "text" or msg_type == "post":
                line_text = f"{sender}：{parsed_text}{extra_content}"
            elif msg_type in ("image", "file"):
                line_text = f"{sender}：{extra_content}"
            else:
                line_text = f"{sender}：{parsed_text or '[无法解析的消息]'}{extra_content}"

            messages_for_summary.append(line_text)

        raw_content = "\n".join(messages_for_summary)
        if not raw_content:
            return "在指定时间范围内未找到聊天记录"

        print(f"[RESEARCH] 筛选后 {len(messages_for_summary)} 条消息，字符数: {len(raw_content)}")

        # 用 MiniMax LLM 总结（纯文本），让 LLM 决定哪些图片要保留
        SYSTEM_PROMPT = """你是一个专业的飞书文档助手。请将下面的群聊记录整理成一份适合写入飞书云文档的纯文本总结。

【重要：图片处理规则】
群聊中出现的所有图片（包括海报、Banner、广告图、设计图），一律必须使用 [保留图片: 文件名] 写入文档，绝对禁止使用 [图片总结]。

具体要求：
1. 任何 .png / .jpg / .gif / .webp / .bmp 后缀的附件文件，消息中会显示为 [图片文件: xxx.png]，必须从中提取原始文件名，输出 [保留图片: xxx.png]
2. 绝对禁止自己起文件名，必须使用消息中 [图片文件:] 标签后的实际文件名
3. [图片内容: ...] 是 OCR 识别结果，仅供文字参考，不影响图片嵌入判断
4. 海报/广告图/设计图等有视觉价值的图片，无论 OCR 结果如何，都必须用 [保留图片:] 嵌入原图

注意以下几类文件的内容已经是完整文字，不需要额外标记保留：
- Excel/Word/PDF 等文件的内容，已通过 [Excel内容:] / [Word内容:] / [PDF内容:] 的形式完整嵌入文档，
  不要再对这些文件使用 [保留图片:] 或 [保留文件:] 等任何标记。
- 如果群聊消息中只是"提到了某个文件"而没有实际附件，也不要标记。
- 同一文件如果已有 [Excel内容:] / [Word内容:] 等标记，不要重复标记。

【文件名规则 - 禁止自行命名】
所有 [保留图片:] 输出的文件名，必须100%匹配消息中 attachments 数组里的 file_name 字段，不允许更改、翻译、概括、添加后缀。

例如：消息里附件的 file_name 是"大促大屏广告.png"，你就必须写 [保留图片: 大促大屏广告.png]，绝对不能写成 [保留图片: 爆发期.jpg]、[保留图片: 广告图1.png] 等。

文件扩展名也必须与原始文件名一致，不得擅自改变（如 .png 不能变成 .jpg）。
如果不确定原始文件名，参考 [图片内容:] 前的文件名部分。

以下图片必须保留（这些无法用文字替代）：
- 海报、Banner、设计图、视觉稿
- 截图（含文字、界面、数据）
- 流程图、架构图、PPT截图
- 图表、表格截图（尤其是设计稿、数据大屏等）
- 代码截图

【来源引用规则】
文档正文中引用任何来自附件的内容时，必须注明来源：
- 引用图片内容：格式为"根据群聊中上传的【文件名】显示/指出/表明……"
- 引用Excel数据：格式为"根据附件【文件名】显示/表明……"
- 引用PDF内容：格式为"参考附件【文件名】，……"
- 引用Word内容：格式为"根据【文件名】，……"

示例：
- "根据群聊中上传的 Q3营收折线图.png 显示，第三季度增长达到15%。"
- "参考附件 用户需求.pdf，核心痛点已提炼如下……"
- "根据 20241103_双11项目预热结案_库存与投放复盘_V2.xlsx，明星精华液套装缺口7000件。"

禁止：直接描述附件内容而不注明来源，或用"根据图表/表格/文件显示"而不写具体文件名。

【输出格式要求】
1. 用 [H1]标题 表示一级标题（对应飞书文档的一级标题）
2. 用 [H2]标题 表示二级标题
3. 用 [H3]标题 表示三级标题
4. 普通内容直接写，每段话一行，不要用 Markdown 语法（不要写 #、-、* 等符号）
5. 不要输出任何 <think>、<think>、</think>、</think> 等标签内容
6. 结构清晰，分段合理，方便阅读

格式示例：
[H1]会议总结
[H2]一、主要讨论话题
[保留图片: 架构图.png]
这里是内容...
[H2]二、已达成的结论
[图片总结: 会议现场拍摄的照片，显示白板上写着关键结论]
结论1
结论2"""

        print(f"[RESEARCH] raw_content 前200字:\n{raw_content[:200]}")
        print(f"[RESEARCH] raw_content 中是否有'[图片内容:': {'[图片内容:' in raw_content}")
        print(f"[RESEARCH] raw_content 中是否有'[Excel内容:': {'[Excel内容:' in raw_content}")
        print(f"[RESEARCH] 准备调用 MiniMax API...")

        try:
            response = _get_minimax_client().chat.completions.create(
                model="MiniMax-M2.7",
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": raw_content}
                ],
                max_tokens=4000,
            )
            print(f"[RESEARCH] API 调用成功")
            summary = response.choices[0].message.content
        except Exception as e:
            import traceback
            import sys
            print(f"[RESEARCH] 调用异常: {type(e).__name__}: {e}")
            traceback.print_exc()
            sys.stdout.flush()
            with open("/tmp/debug_research.txt", "w") as f:
                f.write(raw_content)
            print(f"[RESEARCH] raw_content 已写入 /tmp/debug_research.txt ({len(raw_content)} 字)")
            sys.stdout.flush()
            return f"LLM 总结失败: {e}"

        print(f"[RESEARCH] 总结完成，长度: {len(summary)} 字符")
        print(f"[DEBUG] LLM输出中是否有'[保留图片:': {'[保留图片:' in summary}")
        print(f"[DEBUG] LLM输出中是否有'[图片总结:': {'[图片总结:' in summary}")
        print(f"[DEBUG] LLM输出前300字:\n{summary[:300]}")
        print(f"[DEBUG] image_files 内容: {image_files}")
        self.context["research_result"] = summary
        self.context["summary"] = summary
        self.context["image_files"] = image_files
        return summary

    def _get_bot_id(self) -> str:
        """获取机器人自身的 open_id"""
        try:
            url = "https://open.feishu.cn/open-apis/bot/v3/info"
            resp = httpx.get(url, headers={"Authorization": f"Bearer {self._get_token()}"}, verify=True, timeout=15.0)
            data = resp.json()
            if data.get("code") == 0:
                return data.get("bot", {}).get("open_id", "")
        except Exception as e:
            print(f"[DOC] 获取 bot id 失败: {e}")
        return ""

    def _get_token(self) -> str:
        """获取 tenant_access_token（带缓存）"""
        if hasattr(self, '_cached_token') and self._cached_token:
            return self._cached_token
        url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
        resp = httpx.post(url, json={
            "app_id": os.getenv("FEISHU_APP_ID"),
            "app_secret": os.getenv("FEISHU_APP_SECRET")
        }, verify=True, timeout=30.0)
        data = resp.json()
        self._cached_token = data.get("tenant_access_token", "")
        print(f"[DEBUG] token 获取: {self._cached_token[:20] if self._cached_token else 'EMPTY'}...")
        return self._cached_token

    def _upload_image_to_doc(self, local_path: str, document_id: str) -> str:
        """上传图片到飞书文档，返回 image_key"""
        import os
        filename = os.path.basename(local_path)
        with open(local_path, "rb") as f:
            img_data = f.read()
        self._cached_token = None  # 强制刷新 token
        token = self._get_token()
        headers = {"Authorization": f"Bearer {token}"}
        # 用 drive API 上传文档图片
        resp = httpx.post(
            "https://open.feishu.cn/open-apis/drive/v1/medias/upload_all",
            headers=headers,
            data={
                "parent_type": "docx_image",
                "parent_id": document_id,
                "extra": json.dumps({"drive_route_token": document_id}),
            },
            files={"file": (filename, img_data, "image/png")},
            timeout=30.0
        )
        result = resp.json()
        print(f"[DEBUG] drive上传 result={result}")
        if result.get("code") == 0:
            return result.get("data", {}).get("file_token", "")
        print(f"[DOC] 图片上传失败: {result}")
        return ""

    def _add_doc_image_block(self, doc_id: str, image_key: str) -> dict:
        """向文档插入图片块"""
        url = f"https://open.feishu.cn/open-apis/docx/v1/documents/{doc_id}/blocks/{doc_id}/children"
        headers = {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json"
        }
        payload = {
            "children": [{
                "block_type": 27,
                "image": {
                    "token": image_key,
                    "width": 600,
                    "height": 400
                }
            }]
        }
        print(f"[DEBUG] 插入图片块 doc_id={doc_id}, image_key={image_key}, url={url}")
        try:
            resp = httpx.post(url, headers=headers, json=payload, verify=True, timeout=30.0)
            if resp.status_code != 200:
                return {"code": -1, "msg": f"HTTP {resp.status_code}: {resp.text[:100]}"}
            return resp.json()
        except Exception as e:
            return {"code": -1, "msg": str(e)}

    def _add_doc_block(self, doc_id: str, text: str, block_type: int = 2) -> dict:
        """写入文档段落块"""
        url = f"https://open.feishu.cn/open-apis/docx/v1/documents/{doc_id}/blocks/{doc_id}/children"
        headers = {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json"
        }
        # heading 块结构与 paragraph 不同，需要单独处理
        if block_type == 3:
            payload = {"children": [{"block_type": 3, "heading1": {"elements": [{"text_run": {"content": text}}]}}]}
        elif block_type == 4:
            payload = {"children": [{"block_type": 4, "heading2": {"elements": [{"text_run": {"content": text}}]}}]}
        elif block_type == 5:
            payload = {"children": [{"block_type": 5, "heading3": {"elements": [{"text_run": {"content": text}}]}}]}
        else:
            payload = {"children": [{"block_type": 2, "text": {"elements": [{"text_run": {"content": text}}], "style": {"align": 1}}}]}
        try:
            resp = httpx.post(url, headers=headers, json=payload, verify=True, timeout=30.0)
            if resp.status_code != 200:
                return {"code": -1, "msg": f"HTTP {resp.status_code}: {resp.text[:100]}"}
            return resp.json()
        except Exception as e:
            return {"code": -1, "msg": str(e)}

    def _do_create_doc(self, task: Task) -> str:
        """创建飞书文档"""
        import time
        title = task.params.get("title", f"文档_{int(time.time())}")
        content = task.params.get("content", "") or self.context.get("research_result", "")
        print(f"[DOC] 开始创建文档: title={title}, content={content[:80] if content else 'empty'}")
        try:
            # 1. 创建空文档（通过 SDK）
            doc_req = (
                CreateDocumentRequest.builder()
                .request_body(
                    CreateDocumentRequestBody.builder()
                    .title(title)
                    .folder_token("")
                    .build()
                )
                .build()
            )
            doc_resp = self.feishu_client.request(doc_req)
            if doc_resp.code != 0:
                return f"创建文档失败: {doc_resp.msg}"

            raw = json.loads(doc_resp.raw.content.decode("utf-8"))
            doc_id = raw.get("data", {}).get("document", {}).get("document_id")
            if not doc_id:
                return f"创建文档失败: 找不到document_id"
            doc_link = f"https://feishu.cn/docx/{doc_id}"
            print(f"[DOC] 文档已创建: doc_id={doc_id}")

            # 2. 开放文档编辑权限给所有参与讨论的人（包括 bot 自身）
            try:
                perm_url = f"https://open.feishu.cn/open-apis/drive/v1/permissions/{doc_id}/members?type=docx"
                headers = {"Authorization": f"Bearer {self._get_token()}", "Content-Type": "application/json"}
                # 获取 bot 自身的 open_id
                bot_id = self._get_bot_id()
                member_ids = [
                    bot_id,  # bot 自身
                    "ou_9f07561e326f93aa980c784bfdda6e29",  # PM
                ]
                for mid in member_ids:
                    resp = httpx.post(perm_url, headers=headers, json={
                        "member_type": "openid",
                        "member_id": mid,
                        "perm": "edit"
                    }, verify=True, timeout=15.0)
                    print(f"[DOC] 权限开放 resp={resp.json()}")
            except Exception as e:
                print(f"[DOC] 权限开放失败: {e}")

            # 2. 写入内容（通过 httpx，绕过 SDK 的 bug）
            # 过滤掉 <think> 等标签内容
            import re
            content = re.sub(r"<think>[\s\S]*?</think>", "", content)
            content = re.sub(r"<think>[\s\S]*?</think>", "", content)

            if content:
                lines = content.split("\n")
                image_files = self.context.get("image_files", {})
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue

                    # 处理图片标记
                    import re
                    keep_match = re.match(r"\[保留图片:\s*(.+?)\]", line)
                    summary_match = re.match(r"\[图片总结:\s*(.+?)\]", line)
                    file_match = re.match(r"\[保留文件:\s*(.+?)\]", line)
                    if file_match:
                        # [保留文件:] 是 LLM 误用的标记，内容已通过 [Excel内容:] 等嵌入，跳过
                        continue
                    if keep_match:
                        fname = keep_match.group(1).strip()
                        print(f"[DEBUG] 发现 [保留图片:] fname={fname}, image_files keys={list(image_files.keys())}")
                        if fname in image_files:
                            local_path = image_files[fname]
                            self._cached_token = None  # 强制刷新 token
                            image_key = self._upload_image_to_doc(local_path, doc_id)
                            if image_key:
                                self._cached_token = None  # 强制刷新 token
                                result = self._add_doc_image_block(doc_id, image_key)
                                if result.get("code") == 0:
                                    print(f"[DOC] 插入图片成功: {fname}")
                                else:
                                    print(f"[DOC] 插入图片失败（OCR已有文字描述，跳过）: {result}")
                            else:
                                print(f"[DOC] 上传图片失败，跳过（OCR已有文字描述）")
                        continue
                    if summary_match:
                        # 图片总结写成一行文字
                        summary_text = summary_match.group(1).strip()
                        result = self._add_doc_block(doc_id, f"（图片：{summary_text}）", block_type=2)
                        if result.get("code") == 0:
                            print(f"[DOC] 写入图片总结: {summary_text[:30]}")
                        continue

                    # 识别标题标记 [H1]/[H2]/[H3]
                    block_type = 2  # 默认普通段落
                    if line.startswith("[H1]"):
                        block_type = 3
                        line = line[4:].strip()
                    elif line.startswith("[H2]"):
                        block_type = 4
                        line = line[4:].strip()
                    elif line.startswith("[H3]"):
                        block_type = 5
                        line = line[4:].strip()
                    # 去掉残留的列表标记
                    if line.startswith("- "):
                        line = line[2:]
                    elif len(line) > 1 and line[0].isdigit() and line[1:].startswith(". "):
                        line = line[line.find(". ")+2:]
                    result = self._add_doc_block(doc_id, line, block_type=block_type)
                    if result.get("code") != 0:
                        print(f"[DOC] 写入失败: block_type={block_type}, line={line[:30]}, result={result}")
                    else:
                        print(f"[DOC] 写入成功 [{block_type}]: {line[:30]}")

            # 存入共享上下文
            print(f"[DOC] 完成: link={doc_link}")
            self.context["doc_title"] = title
            self.context["doc_link"] = doc_link
            return f"文档已创建：{title}\n链接：{doc_link}"
        except Exception as e:
            import traceback
            traceback.print_exc()
            return f"创建文档失败: {e}"

    def _do_create_bitable(self, task: Task) -> str:
        """创建飞书多维表格"""
        title = task.params.get("title", "未命名表格")
        # TODO: 调用飞书多维表格 API
        return f"多维表格已创建：{title}"

    def _do_create_ppt(self, task: Task) -> str:
        """创建飞书演示文稿"""
        title = task.params.get("title", "未命名演示")
        # TODO: 调用飞书幻灯片 API
        return f"演示稿已创建：{title}"

    def _do_report(self, task: Task) -> str:
        """向群聊发送文字汇报"""
        import re
        # 优先用 context 里的文档链接（DOC 步骤写入的）
        doc_link = self.context.get("doc_link", "")
        if doc_link:
            content = f"{self.context.get('doc_title', '文档')}已生成，请查收：{doc_link}"
        else:
            # 没有文档时，用 LLM 生成的研究总结（research_result）
            summary = self.context.get("research_result", "")
            if summary:
                # 去掉<think>...</think>标签
                summary = re.sub(r"<think>[\s\S]*?</think>", "", summary).strip()
            content = summary or "任务已完成"
        print(f"[REPORT] 准备发送汇报到 chat_id={self.chat_id}, content={content[:100] if content else 'empty'}")
        try:
            request = (
                CreateMessageRequest.builder()
                .receive_id_type("chat_id")
                .request_body(
                    CreateMessageRequestBody.builder()
                    .receive_id(self.chat_id)
                    .msg_type("text")
                    .content(json.dumps({"text": content}))
                    .build()
                )
                .build()
            )
            resp = self.feishu_client.request(request)
            print(f"[REPORT] resp.code={resp.code}, resp.msg={getattr(resp, 'msg', 'N/A')}, resp.data={getattr(resp, 'data', 'N/A')}")
            if resp.code != 0:
                return f"发送失败: {resp.msg}"

            # 如果 need_ppt=true，直接把 doc_link 发给 PPT agent 即可
            need_ppt = task.params.get("need_ppt", False)
            if need_ppt:
                self.context["need_ppt"] = True
            if need_ppt and self.context.get("doc_link"):
                doc_url = self.context['doc_link']
                # post 类型消息，用 at 标签正确 @ OpenCLAW
                post_content = json.dumps({
                    "zh_cn": {
                        "title": "",
                        "content": [[
                            {"tag": "at", "user_id": "ou_27ff916fd7f7fdc7455ee3806e7ad23d", "user_name": "OpenClaw"},
                            {"tag": "text", "text": f" 文档已生成，请根据此链接制作PPT：{doc_url}"}
                        ]]
                    }
                })
                ppt_request = (
                    CreateMessageRequest.builder()
                    .receive_id_type("chat_id")
                    .request_body(
                        CreateMessageRequestBody.builder()
                        .receive_id("oc_ea67a09ec7edc0143ce7140b549635db")
                        .msg_type("post")
                        .content(post_content)
                        .build()
                    )
                    .build()
                )
                ppt_resp = self.feishu_client.request(ppt_request)
                print(f"[REPORT->OpenCLAW] resp.code={ppt_resp.code}, resp.msg={getattr(ppt_resp, 'msg', 'N/A')}")
                # PPT 还在制作中，暂不标记完成，保持 running 状态
                self.context["ppt_pending"] = True

            return f"已发送汇报：{content[:50]}..."
        except Exception as e:
            return f"发送失败: {e}"

    def _build_ppt_card(self) -> dict:
        """构建 PPT 制作中的卡片"""
        steps = [
            ("🔍 读取文档", "waiting"),
            ("✏️ 生成幻灯片", "waiting"),
            ("💾 导出文件", "waiting"),
        ]
        elements = []
        for i, (name, status) in enumerate(steps):
            color = {"waiting": "grey", "running": "orange", "done": "green"}.get(status, "grey")
            icon = {"waiting": "🔘 等待中", "running": "🔄 进行中...", "done": "✅ 已完成"}.get(status, "🔘 等待中")
            bar = {"waiting": "", "running": "[██░░░░░░░░] 20%", "done": ""}.get(status, "")
            content = f"**{name}**\n<font color='{color}'>{icon}</font>"
            if bar:
                content += f"\n{bar}"
            elements.append({"tag": "markdown", "content": content, "margin": "8px 0px 4px 0px" if i == 0 else "4px 0px 4px 0px"})
        elements.append({"tag": "hr"})
        elements.append({"tag": "markdown", "content": "<font color='grey'>预计剩余时间：约 1 分钟</font>", "text_size": "small", "margin": "4px 0px 0px 0px"})
        return {
            "schema": "2.0",
            "config": {"update_multi": True},
            "header": {
                "title": {"tag": "plain_text", "content": "📊 PPT 制作中"},
                "template": "blue"
            },
            "body": {"elements": elements}
        }

    def _send_ppt_card_to_main(self) -> str:
        """发一张 PPT 制作中的卡片到主群，返回 message_id"""
        if not self.send_card_func:
            return ""
        card_json = self._build_ppt_card()
        msg_id = self.send_card_func(self.chat_id, card_json)
        print(f"[PPT卡片] 已发送到主群, message_id={msg_id}")
        # 清空 card_msg_id，防止主 agent 后续覆盖这张卡
        self.card_msg_id = None
        return msg_id

    # ==================== 状态卡片 ====================

    def _build_card(self, tasks: list[Task]) -> dict:
        """构建卡片 JSON"""
        status_colors = {
            "pending": "<font color='grey'>🔘 等待中</font>",
            "running": "<font color='orange'>🔄 进行中</font>",
            "done": "<font color='green'>✅ 完成</font>",
            "failed": "<font color='red'>❌ 失败</font>",
            "cancelled": "<font color='grey'>⚠️ 已取消</font>",
        }
        elements = []
        for t in tasks:
            color_tag = status_colors.get(t.status, status_colors["pending"])
            elements.append({
                "tag": "markdown",
                "content": f"**步骤{t.step}：{t.desc}**\n{color_tag}",
                "margin": "8px 0px 4px 0px"
            })
        elements.append({"tag": "hr"})
        elements.append({
            "tag": "markdown",
            "content": "🤖 Agent 执行中...",
            "margin": "8px 0px 0px 0px"
        })
        return {
            "schema": "2.0",
            "config": {"update_multi": True},
            "header": {
                "title": {"tag": "plain_text", "content": "🤖 Agent 执行进度"},
                "template": "blue"
            },
            "body": {"elements": elements}
        }

    def _build_confirm_card(self, tasks: list[Task]) -> dict:
        """构建待确认的卡片（带按钮）"""
        elements = []
        for t in tasks:
            color_tag = {
                "pending": "<font color='grey'>🔘 等待中</font>",
                "running": "<font color='orange'>🔄 进行中</font>",
                "done": "<font color='green'>✅ 完成</font>",
                "failed": "<font color='red'>❌ 失败</font>",
            }.get(t.status, "<font color='grey'>🔘 等待中</font>")
            elements.append({
                "tag": "markdown",
                "content": f"**步骤{t.step}：{t.desc}**\n{color_tag}",
                "margin": "8px 0px 4px 0px"
            })

        elements.append({"tag": "hr"})
        elements.append({
            "tag": "markdown",
            "content": f"📄 **文档已生成**：{self.context.get('doc_link', '')}\n请确认内容无误后，点击下方按钮生成 PPT",
            "margin": "8px 0px 4px 0px"
        })
        elements.append({
            "tag": "button",
            "text": {"tag": "plain_text", "content": "✅ 确认大纲，生成 PPT"},
            "type": "callback",
            "value": {"action": "confirm_ppt", "doc_link": self.context.get("doc_link", "")}
        })
        elements.append({
            "tag": "button",
            "text": {"tag": "plain_text", "content": "🔄 大纲有误，去云文档修改"},
            "type": "callback",
            "value": {"action": "retry"}
        })
        return {
            "schema": "2.0",
            "config": {"update_multi": True},
            "header": {
                "title": {"tag": "plain_text", "content": "⏳ 待确认大纲"},
                "template": "yellow"
            },
            "body": {"elements": elements}
        }

    def _send_confirm_card(self, tasks: list[Task]) -> str:
        """发送待确认卡片，返回 message_id"""
        if not self.send_card_func:
            return ""
        card_json = self._build_confirm_card(tasks)
        msg_id = self.send_card_func(self.chat_id, card_json)
        self.card_msg_id = msg_id
        print(f"[确认卡片] 已发送, message_id={msg_id}")
        return msg_id

    def _update_card(self, tasks: list[Task]) -> None:
        """发送或更新飞书卡片（首次发，后续更新）"""
        if not self.send_card_func:
            return
        card_json = self._build_card(tasks)
        if not self.card_msg_id:
            msg_id = self.send_card_func(self.chat_id, card_json)
            self.card_msg_id = msg_id
            print(f"[卡片] 已发送, message_id={msg_id}")
        else:
            if self.update_card_func:
                ok = self.update_card_func(self.card_msg_id, card_json)
                print(f"[卡片] 更新{'成功' if ok else '失败'}, message_id={self.card_msg_id}")

    # ==================== 入口 ====================

    def run(self, user_input: str) -> list[Task]:
        """
        统一入口：规划 + 执行
        """
        # 1. 规划
        print(f"[Agent] 开始规划任务: {user_input}")
        tasks = self.plan(user_input)
        print(f"[Agent] 规划结果: {len(tasks)} 个步骤")

        if not tasks:
            print("[Agent] 规划失败，未生成任务清单")
            return []

        # 2. 执行
        print(f"[Agent] 开始执行 {len(tasks)} 个步骤")
        results = self.execute(tasks)

        # 3. 汇总
        done = sum(1 for t in results if t.status == "done")
        failed = sum(1 for t in results if t.status == "failed")
        print(f"[Agent] 执行完成: {done} 成功, {failed} 失败")

        return results
