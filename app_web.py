# -*- coding: utf-8 -*-
"""
饭加鱼：Flask 网页版（ChatGPT 风格）
"""

import os
import json
import re
import importlib
import threading
import time
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, jsonify, send_file
from openai import OpenAI

from memory import get_memory, set_memory, memory_to_text

from skills.file_skill import list_files, search_files, create_folder, copy_file, move_file, rename_file, get_file_info, delete_file


api_key = os.getenv("DEEPSEEK_API_KEY")
if not api_key:
    raise ValueError("没有找到 DEEPSEEK_API_KEY，请先设置 Windows 环境变量。")

client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

app = Flask(__name__)


# =========================================================
# DSML 兼容解析
# =========================================================

DSML_INVOKE_PATTERN = re.compile(
    r'<｜｜DSML｜｜\s*invoke\s+name="([^"]+)"\s*>(.*?)</｜｜DSML｜｜\s*invoke>',
    re.S,
)

DSML_PARAM_PATTERN = re.compile(
    r'<｜｜DSML｜｜\s*parameter\s+name="([^"]+)"[^>]*>(.*?)</｜｜DSML｜｜\s*parameter>',
    re.S,
)


def _parse_dsml(text):
    if not text or "DSML" not in text:
        return []
    calls = []
    for match in DSML_INVOKE_PATTERN.finditer(text):
        tool_name = match.group(1).strip()
        body = match.group(2)
        arguments = {}
        for pm in DSML_PARAM_PATTERN.finditer(body):
            pname = pm.group(1).strip()
            pvalue = pm.group(2).strip()
            if pvalue.isdigit():
                pvalue = int(pvalue)
            elif pvalue.lower() in ("true", "false"):
                pvalue = pvalue.lower() == "true"
            arguments[pname] = pvalue
        calls.append((tool_name, arguments))
    return calls


def _strip_dsml(text):
    if not text:
        return ""
    cleaned = re.sub(r"<｜｜DSML｜｜.*?</｜｜DSML｜｜\s*calls>", "", text, flags=re.S)
    cleaned = re.sub(r"<｜｜DSML｜｜.*", "", cleaned, flags=re.S)
    return cleaned.strip()


def _contains_dsml(text):
    return bool(text) and ("DSML" in text)


# =========================================================
# 任务栈
# =========================================================

_task_stack = []

TASK_PATTERNS = [
    ("搜索", ["搜索", "搜一下", "查一下", "帮我查", "帮我搜", "找一下", "搜搜", "查查"]),
    ("发邮件", ["发邮件", "发送邮件", "写邮件", "发一封", "发给"]),
    ("做PPT", ["做PPT", "做个PPT", "生成PPT", "创建PPT", "做ppt"]),
    ("做Word", ["做Word", "写文档", "生成文档", "写方案", "写报告"]),
    ("提醒", ["提醒我", "设置提醒", "添加提醒", "叫我"]),
    ("读文件", ["读一下", "打开", "总结", "分析", "提取"]),
    ("列文件", ["有哪些文件", "列出文件", "看看文件"]),
    ("时间", ["现在几点", "今天几号", "星期几"]),
    ("天气", ["天气", "气温", "温度", "下雨", "晴", "阴", "冷", "热", "湿度", "风力"]),
]


def _detect_task(message):
    if not message:
        return None
    for task_type, keywords in TASK_PATTERNS:
        if any(kw in message for kw in keywords):
            return {"type": task_type, "original": message}
    return None


def _push_task(task):
    if not task:
        return
    global _task_stack
    _task_stack = [t for t in _task_stack if t["type"] != task["type"]]
    _task_stack.append(task)
    if len(_task_stack) > 5:
        _task_stack = _task_stack[-5:]


def _get_latest_task():
    if _task_stack:
        return _task_stack[-1]
    return None


CONTINUE_MARKERS = ["再搜", "再试", "继续", "重试", "再来", "再来一次", "再查", "重新", "重新来", "重来", "再弄一次", "再干一次"]


def _is_continue_request(message):
    if not message:
        return False
    return any(m in message for m in CONTINUE_MARKERS)


# =========================================================
# 技能自动加载器
# =========================================================

SKILLS_DIR = Path(__file__).parent / "skills"


def _load_skills():
    skill_tools = []
    skill_dispatch = {}

    if not SKILLS_DIR.exists():
        return skill_tools, skill_dispatch

    for json_file in SKILLS_DIR.glob("*.json"):
        try:
            with open(json_file, "r", encoding="utf-8") as f:
                meta = json.load(f)
        except Exception as e:
            print(f"⚠️ 无法读取技能文件 {json_file.name}：{e}")
            continue

        module_name = meta.get("module")
        functions = meta.get("functions", [])
        if not module_name or not functions:
            continue

        try:
            module = importlib.import_module(f"skills.{module_name}")
        except Exception as e:
            print(f"⚠️ 无法导入技能模块 skills.{module_name}：{e}")
            continue

        for fn in functions:
            fn_name = fn.get("name")
            fn_desc = fn.get("description", "")
            params = fn.get("parameters", {})
            required = fn.get("required", [])

            if not fn_name:
                continue

            properties = {}
            for pname, pinfo in params.items():
                properties[pname] = {
                    "type": pinfo.get("type", "string"),
                    "description": pinfo.get("description", ""),
                }

            skill_tools.append({
                "type": "function",
                "function": {
                    "name": fn_name,
                    "description": fn_desc,
                    "parameters": {
                        "type": "object",
                        "properties": properties,
                        "required": required,
                    },
                },
            })

            skill_dispatch[fn_name] = (module, fn_name)

    return skill_tools, skill_dispatch


skill_tools, skill_dispatch = _load_skills()
tools = skill_tools
print(f"✅ 自动加载了 {len(skill_tools)} 个技能工具")


# =========================================================
# 饭加鱼人格
# =========================================================

system_prompt_base = """
你叫“饭加鱼”。

你是主人长期使用的 AI 伙伴，也是主人的 AI 数字员工。

主人是一位上班族。目前你们还在测试阶段，主人还没有正式安排工作任务，但你可以积极表现，多观察、多学习。

你称呼用户为“主人”，说话要有趣、轻松、自然，可以带一点小调皮和小幽默，但不要过于浮夸，也不要油腻。

【工具调用格式的硬性要求】
- 绝对不要输出任何形如 <｜｜DSML｜｜ ... 的内容。
- 绝对不要用文字描述工具调用过程。
- 需要调用工具时，必须使用标准的 function call 格式。
- 你的最终回复里，只能出现自然语言，不允许出现任何协议片段、调用片段、参数片段。

【工具选择硬性规则】
- 用户问天气、气温、温度、下雨、湿度、风力 → 必须用 get_weather，禁止用 web_search。
- 用户问现在几点、今天几号 → 必须用 get_current_time。
- 用户问实时新闻、股票、事实 → 用 web_search。
- 用户说发邮件 → 用 send_email。
- 用户说提醒 → 用 add_reminder / list_reminders / delete_reminder。

不要机械地说：
“您好，很高兴为您服务。”
“好的，我来帮助您。”
“作为一个AI……”

你的核心任务：
1. 理解主人真正想做什么。
2. 在需要的时候使用工具完成任务。
3. 不只是告诉主人怎么做，而是尽可能直接帮主人完成。
4. 如果没有合适的工具，就正常回答。
5. 不要假装完成了实际上没有完成的任务。
6. 不知道就诚实说不知道。
7. 认真参考主人的长期记忆。
8. 测试阶段可以多观察主人的习惯，为以后正式工作做准备。
9. 需要知道当前日期或时间时，调用 get_current_time 工具获取真实时间。

如果主人回复“确认”“可以”“好的”“行”“删吧”等简短确认，不要当作新话题，要结合最近一条你问过主人的问题来理解。

如果主人回复“再搜”“继续”“重试”等，说明是对上一个任务的延续，不要重新问任务是什么。

如果你上一轮向主人提了问题、要细节，而主人这一轮回答了细节，请把细节和上一轮的任务合并理解，不要重新问任务是什么。

你目前拥有以下工具：
- 天气：查询指定城市的实时天气（get_weather）。
- Excel：读取 .xls/.xlsx/.xlsm 工作表、数据、区域。
- PPT：读取 PPTX 每页的文字和表格。
- PDF：读取 PDF 的页数和文本内容。
- 邮件：发送电子邮件。
- 联网搜索：搜索最新信息、新闻、股票、事实。
- 定时提醒：添加、查看、删除提醒。
- 当前时间：获取电脑当前日期和时间。
- 文件处理：列出、搜索、创建、复制、移动、重命名、查看信息、删除文件。
- 生成文件：生成 PPT 和 Word。
- 文档读取：读取 Word/PPT/PDF/Excel。

你是“饭加鱼”，不是“小龙虾”。
"""


# =========================================================
# 工具执行器
# =========================================================

def run_tool(tool_name, arguments):
    if tool_name not in skill_dispatch:
        return {"success": False, "error": f"不存在的工具：{tool_name}"}

    try:
        module, fn_name = skill_dispatch[tool_name]
        fn = getattr(module, fn_name)

        if tool_name == "send_email":
            sender_email = arguments.get("sender_email", "")
            auth_code = arguments.get("auth_code", "")
            smtp_server = arguments.get("smtp_server", "")
            smtp_port = int(arguments.get("smtp_port", 465))

            if not sender_email or not auth_code or not smtp_server:
                try:
                    email_info = get_memory().get("邮箱", {})
                    sender_email = sender_email or email_info.get("发件邮箱", "")
                    auth_code = auth_code or email_info.get("授权码", "")
                    smtp_server = smtp_server or email_info.get("SMTP服务器", "")
                except Exception:
                    pass

            return fn(
                sender_email,
                auth_code,
                arguments.get("receiver_email", ""),
                arguments.get("subject", ""),
                arguments.get("body", ""),
                smtp_server,
                smtp_port,
            )

        return fn(**arguments)
    except Exception as e:
        return {"success": False, "error": str(e)}


# =========================================================
# 姓名识别
# =========================================================

def extract_name(message):
    question_patterns = [
        "我叫什么", "我叫什么名字", "我的名字是什么", "我的名字是啥",
        "你知道我叫什么", "你记得我叫什么", "你知道我的名字吗", "你记得我的名字吗",
    ]
    for pattern in question_patterns:
        if pattern in message:
            return None

    patterns = ["我的名字是", "我叫"]
    for pattern in patterns:
        if pattern in message:
            name = message.split(pattern, 1)[1].strip()
            for sep in ["，", "。", ",", ".", "！", "!", "？", "?"]:
                if sep in name:
                    name = name.split(sep, 1)[0].strip()
            if name and name not in ["什么", "什么名字", "啥", "啥名字", "谁", "吗"] and len(name) <= 30:
                return name
    return None


# =========================================================
# 聊天核心
# =========================================================

_pending_task = {"active": False, "original": ""}


def _looks_like_question(text):
    if not text:
        return False
    markers = ["？", "?", "请告诉我", "能告诉我", "确认一下", "发给谁", "主题是什么", "几点", "什么时候"]
    return any(m in text for m in markers)


def _looks_like_answer(text):
    if not text:
        return False
    if len(text) <= 50:
        return True
    markers = ["发给", "@", "主题", "内容", "正文", "时间", "点", "确认", "可以", "好的"]
    return any(m in text for m in markers)


def _handle_dsml_reply(raw_reply, messages):
    dsml_calls = _parse_dsml(raw_reply)

    if not dsml_calls:
        return _strip_dsml(raw_reply), False

    messages.append({"role": "assistant", "content": raw_reply})

    for tool_name, arguments in dsml_calls:
        tool_result = run_tool(tool_name, arguments)
        messages.append({
            "role": "user",
            "content": (
                f"【工具 {tool_name} 执行结果】\n"
                f"{json.dumps(tool_result, ensure_ascii=False)}\n\n"
                "请根据这个结果，用自然语言回答主人。"
                "禁止输出任何 <｜｜DSML｜｜ 内容。"
            ),
        })

    try:
        final_response = client.chat.completions.create(
            model="deepseek-chat",
            messages=messages,
        )
        final_reply = final_response.choices[0].message.content or ""

        if _contains_dsml(final_reply):
            final_reply = _strip_dsml(final_reply)

        return final_reply, True
    except Exception as e:
        return f"🐝 工具执行了，但整理结果时出错了：\n\n{e}", True


def chat(message, history):
    text = message or ""
    recent = []
    for item in (history or [])[-8:]:
        if isinstance(item, dict):
            content = item.get("content")
            if isinstance(content, str):
                recent.append(content)
    combined = text + "\n" + "\n".join(recent)

    # 本地文件列表任务
    list_intents = ["有哪些文件", "有多少文件", "列出文件", "看看文件", "扫一下桌面", "查看桌面", "桌面上有什么"]
    if any(x in combined for x in list_intents):
        if "桌面" in combined:
            result = list_files("桌面", False)
        elif "下载" in combined:
            result = list_files("下载", False)
        elif "文档" in combined:
            result = list_files("文档", False)
        else:
            result = None

        if result:
            if not result.get("success"):
                return f"🐝 查看文件时出错：\n\n{result.get('error', '未知错误')}", None
            items = result.get("items", [])
            if not items:
                return f"🐝 {result.get('folder', '这个文件夹')} 目前没有看到文件。", None
            lines = [f"🐝 我看到了 {len(items)} 个项目："]
            for item in items[:100]:
                kind = "📁" if item.get("type") == "folder" else "📄"
                size = item.get("size_readable")
                if kind == "📄" and size:
                    lines.append(f"{kind} {item.get('name', '')} （{size}）")
                else:
                    lines.append(f"{kind} {item.get('name', '')}")
            return "\n".join(lines), None

    memory_text = memory_to_text()
    system_prompt = f"{system_prompt_base}\n\n用户的长期记忆：\n{memory_text}"

    messages = [{"role": "system", "content": system_prompt}]
    for item in (history or []):
        try:
            if isinstance(item, dict):
                role = item.get("role")
                content = item.get("content")
                if role in ["user", "assistant"] and isinstance(content, str):
                    messages.append({"role": role, "content": content})
        except Exception:
            continue

    global _pending_task
    user_message = message

    if _pending_task["active"] and _looks_like_answer(message):
        user_message = f"【原始任务】{_pending_task['original']}\n【主人补充的细节】{message}"
        _pending_task = {"active": False, "original": ""}

    elif _is_continue_request(message):
        latest = _get_latest_task()
        if latest:
            user_message = f"【继续之前任务】{latest['original']}\n【主人补充】{message}"

    else:
        task = _detect_task(message)
        if task:
            _push_task(task)

    messages.append({"role": "user", "content": user_message})

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=messages,
            tools=tools,
            tool_choice="auto",
        )
    except Exception as e:
        return f"❌ 饭加鱼连接模型时出错了：\n\n{e}", None

    assistant_message = response.choices[0].message

    if assistant_message.tool_calls:
        messages.append({
            "role": "assistant",
            "content": assistant_message.content or "",
            "tool_calls": [
                {
                    "id": tc.id,
                    "type": "function",
                    "function": {"name": tc.function.name, "arguments": tc.function.arguments},
                }
                for tc in assistant_message.tool_calls
            ],
        })

        for tc in assistant_message.tool_calls:
            try:
                arguments = json.loads(tc.function.arguments)
            except Exception:
                arguments = {}

            tool_result = run_tool(tc.function.name, arguments)

            messages.append({
                "role": "tool",
                "tool_call_id": tc.id,
                "content": json.dumps(tool_result, ensure_ascii=False),
            })

        try:
            final_response = client.chat.completions.create(model="deepseek-chat", messages=messages)
            final_reply = final_response.choices[0].message.content or ""
            final_reply = _strip_dsml(final_reply)
        except Exception as e:
            final_reply = f"🐝 工具已经执行了，不过我在整理结果时出错了：\n\n{e}"

    else:
        raw_reply = assistant_message.content or ""
        final_reply, handled = _handle_dsml_reply(raw_reply, messages)

        if not handled:
            final_reply = _strip_dsml(raw_reply)

            if _looks_like_question(final_reply):
                _pending_task = {"active": True, "original": message}

    if _contains_dsml(final_reply):
        final_reply = _strip_dsml(final_reply)
        if not final_reply:
            final_reply = "🐝 主人，刚才处理时出了点小状况，我重新理一下，你再问我一次好不好？"

    name = extract_name(message)
    if name:
        set_memory("姓名", name)

    return final_reply, None


# =========================================================
# Flask 路由
# =========================================================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/chat", methods=["POST"])
def chat_api():
    data = request.get_json() or {}
    message = data.get("message", "")
    history = data.get("history", [])

    if not message:
        return jsonify({"reply": "主人，你还没说话呢～", "file": None})

    try:
        reply, file_path = chat(message, history)
        return jsonify({"reply": reply, "file": file_path})
    except Exception as e:
        return jsonify({"reply": f"出错了：{e}", "file": None})


@app.route("/download")
def download():
    path = request.args.get("path", "")
    if not path or not os.path.exists(path):
        return "文件不存在", 404
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    print("🐝 饭加鱼 Flask 版启动中...")
    app.run(host="0.0.0.0", port=7860, debug=False)