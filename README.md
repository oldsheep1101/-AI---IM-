# 飞书 AI Agent 机器人

通过自然语言指令，在飞书群聊中自动完成"读取聊天记录 → 生成云文档 → 制作 PPT"的完整工作流。

## 架构图

```graph TB
    subgraph 用户层["👤 用户层"]
        Mobile["📱 移动端 iOS/Android"]
        Desktop["💻 桌面端 macOS/Windows"]
    end

    subgraph 飞书平台["☁️ 飞书开放平台"]
        IM["💬 IM 群聊/单聊"]
        Card["🧩 交互卡片"]
        DocAPI["📄 文档 API"]
        PPTAPI["📊 PPT 联动"]
    end

    subgraph 接入层["📡 接入层"]
        WS["WebSocket 长连接<br/>lark.ws.Client"]
        Dispatcher["EventDispatcherHandler<br/>• p2_im_message_receive_v1<br/>• p2_card_action_trigger"]
    end

    subgraph 核心引擎["🧠 AI Agent 核心引擎 main.py / agent.py"]
        Router["事件路由<br/>handle_message() / handle_card_action()"]
        Planner["LLM Planner<br/>MiniMax M2.7<br/>意图理解 → 任务拆解"]
        Executor["任务执行器 Executor"]
        Research["🔍 RESEARCH<br/>读消息/OCR/Excel/PDF"]
        Doc["📝 DOC<br/>创建文档/写块/传图"]
        Report["📣 REPORT<br/>发总结/推链接/触发PPT"]
        StateMachine["状态机<br/>planning→researching→writing→confirming→done"]
        HIL["Human-in-the-loop<br/>卡片确认中断/恢复"]
    end

    subgraph 数据层["💾 数据层"]
        MsgStore["messages.jsonl<br/>群聊消息持久化"]
        Attachment["附件处理<br/>Tesseract OCR<br/>openpyxl / pypdf"]
    end

    subgraph 外部服务["🔌 外部服务"]
        MiniMax["MiniMax API"]
        OpenCLAW["OpenCLAW PPT 生成"]
    end

    Mobile --> IM
    Desktop --> IM
    IM --> WS
    Card --> WS
    WS --> Dispatcher
    Dispatcher --> Router
    Router --> Planner
    Planner --> MiniMax
    Planner --> Executor
    Executor --> Research
    Executor --> Doc
    Executor --> Report
    Research --> MsgStore
    Research --> Attachment
    Doc --> DocAPI
    Report --> PPTAPI
    PPTAPI --> OpenCLAW
    Report --> Card
    StateMachine --> HIL
    HIL --> Card
    Router --> StateMachine

    style 核心引擎 fill:#e1f5fe,stroke:#0288d1
    style Planner fill:#fff3e0,stroke:#f57c00
    style Executor fill:#e8f5e9,stroke:#388e3c
    style HIL fill:#fce4ec,stroke:#c62828
```

## 目录结构

```
feishu-bot/
├── main.py                      # 入口：WebSocket 长连接 + 事件分发
├── agent.py                     # 核心：LLM 规划 + 任务执行
├── lark_message_listener.py     # 消息收集：WebSocket 实时监听 → 写入 messages.jsonl
├── lark_message_fetcher.py     # 消息收集：定时拉取 + 富文本/附件下载
├── test_data/                   # 测试数据
└── README.md                    # 本文档
```

## 工作流程

```
用户发送指令（如"把昨天的聊天记录总结并生成PPT"）
    ↓
WebSocket 接收消息 → handle_message()
    ↓
Agent.plan() → MiniMax LLM 解读意图，生成任务清单
    ↓
[示例任务清单]
Step 1: RESEARCH → 读取 messages.jsonl（支持时间范围筛选）
Step 2: DOC      → 创建飞书云文档（标题/段落写入）
Step 3: REPORT   → 发链接到群里，触发 PPT agent（如需）
    ↓
Executor 顺序执行，实时更新卡片进度
    ↓
[Human-in-loop] DOC 完成后弹出确认卡片
    ↓
用户点击"确认大纲，生成PPT" → 触发 OpenCLAW
    ↓
OpenCLAW 完成 → 机器人收到消息 → 转发 PPT 链接到群
```

## 使用方法

### 1. 环境配置

```bash
# 安装依赖
pip install -r requirements.txt
# requirements.txt 至少包含：
# lark-oapi, openai, httpx, openpyxl, pypdf, python-docx

# 安装 Tesseract（用于图片 OCR）
brew install tesseract tesseract-lang  # macOS
# Ubuntu: sudo apt install tesseract-ocr tesseract-ocr-chi-sim
```

### 2. 环境变量

在项目根目录创建 `.env` 文件：

```env
FEISHU_APP_ID=your_app_id
FEISHU_APP_SECRET=your_app_secret
MINIMAX_API_KEY=your_minimax_api_key
LARKBOT_APP_ID=your_larkbot_app_id
LARKBOT_APP_SECRET=your_larkbot_app_secret
```

### 3. 启动机器人

```bash
SSL_CERT_FILE=/etc/ssl/cert.pem SSL_CERT_DIR=/etc/ssl/certs python main.py
```

### 4. 消息收集（可选）

**方式一：实时监听（推荐）**
```bash
python lark_message_listener.py
```
需要先配置 `lark-cli` 并登录：
```bash
lark-cli auth login --domain feishu
```

**方式二：定时拉取**
```bash
python lark_message_fetcher.py
# 可配合 crontab 每分钟执行
# */1 * * * * /path/to/venv/bin/python /path/to/lark_message_fetcher.py
```

消息默认保存到 `~/feishu_messages/messages.jsonl`。

### 5. 群聊指令示例

| 指令 | 行为 |
|------|------|
| `把昨天的聊天记录总结给我` | RESEARCH → DOC → 发云文档链接到群里 |
| `把今天的聊天内容总结并生成PPT` | RESEARCH → DOC → 发文档链接 → 触发 OpenCLAW 制作 PPT |
| `总结本周的讨论` | RESEARCH → DOC → 发云文档链接到群里 |

### 6. 卡片与确认流程

- 机器人会实时发送带进度的交互卡片（步骤完成状态）
- 生成云文档后会弹出**大纲确认卡片**，点击按钮后才触发 PPT 生成
- 卡片消息支持"确认大纲"和"大纲有误，去云文档修改"两个操作

## 核心模块说明

### agent.py

- **Planner**：用 MiniMax LLM 解读用户意图，输出 JSON 任务清单
- **Executor**：顺序执行任务，支持 Human-in-loop 中断恢复
- **RESEARCH**：读取 `messages.jsonl`，对图片做 OCR（tesseract）、Excel 全量读取（openpyxl）、PDF 文本提取（pypdf）
- **DOC**：调用飞书文档 API 创建文档、写入段落/标题/图片 block
- **REPORT**：发送消息到群聊，可触发 OpenCLAW 制作 PPT

### main.py

- 使用 `lark_oapi` 的 WebSocket 长连接接收飞书事件
- `handle_message`：处理新消息，调用 Agent 规划和执行
- `handle_card_action`：处理卡片按钮点击事件，触发 resume 恢复执行

### lark_message_listener.py / lark_message_fetcher.py

- 将飞书群消息持久化到本地 `messages.jsonl`
- 支持图片、文件等富文本附件的下载存储
- `lark_message_listener.py` 通过 WebSocket 实时接收（需 lark-cli）
- `lark_message_fetcher.py` 通过轮询拉取（无需 lark-cli）

## 注意事项

- OCR 依赖本地 Tesseract 安装，中文识别需 `tesseract-lang`
- Excel 完整读取无字符限制，数据量大时处理时间较长
- 图片嵌入飞书文档需使用 `drive/v1/medias/upload_all` API
- PPT 制作由 OpenCLAW 独立完成，机器人仅负责转发链接
