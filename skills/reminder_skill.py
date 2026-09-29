# -*- coding: utf-8 -*-
"""
饭加鱼: 定时提醒技能 v2
"""

import json
import re
from datetime import datetime, timedelta
from pathlib import Path

REMINDER_FILE = Path(__file__).parent / "reminders.json"


def _load():
    if not REMINDER_FILE.exists():
        return []
    try:
        with open(REMINDER_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception:
        return []


def _save(reminders):
    try:
        with open(REMINDER_FILE, "w", encoding="utf-8") as f:
            json.dump(reminders, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def add_reminder(time_str: str, content: str):
    """
    添加提醒。
    如果 time_str 是相对时间（例如 "2分钟后"），会自动转成真实时间。
    """
    if not time_str or not content:
        return {"success": False, "error": "提醒时间和内容不能为空。"}

    # 如果 time_str 不是标准时间，尝试解析相对时间
    if not re.match(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", time_str):
        now = datetime.now()
        if "分钟后" in time_str:
            match = re.search(r"(\d+)", time_str)
            minutes = int(match.group(1)) if match else 2
            real_time = now + timedelta(minutes=minutes)
            time_str = real_time.strftime("%Y-%m-%d %H:%M")
        elif "小时后" in time_str:
            match = re.search(r"(\d+)", time_str)
            hours = int(match.group(1)) if match else 1
            real_time = now + timedelta(hours=hours)
            time_str = real_time.strftime("%Y-%m-%d %H:%M")
        else:
            real_time = now + timedelta(minutes=2)
            time_str = real_time.strftime("%Y-%m-%d %H:%M")

    reminders = _load()
    reminders.append({
        "time": time_str,
        "content": content,
        "done": False,
    })
    if _save(reminders):
        return {
            "success": True,
            "message": f"已添加提醒：{time_str} - {content}",
        }
    return {"success": False, "error": "提醒保存失败。"}


def list_reminders():
    """查看所有提醒。"""
    reminders = _load()
    if not reminders:
        return {"success": True, "reminders": [], "message": "当前没有提醒。"}
    return {
        "success": True,
        "reminders": reminders,
        "message": f"共有 {len(reminders)} 条提醒。",
    }


def delete_reminder(index: int):
    """删除指定序号的提醒。"""
    reminders = _load()
    if index < 0 or index >= len(reminders):
        return {"success": False, "error": "提醒序号不存在。"}
    removed = reminders.pop(index)
    _save(reminders)
    return {"success": True, "message": f"已删除提醒：{removed['time']} - {removed['content']}"}