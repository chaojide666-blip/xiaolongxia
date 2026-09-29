# -*- coding: utf-8 -*-
from datetime import datetime

def get_current_time():
    now = datetime.now()
    wm = ["星期一","星期二","星期三","星期四","星期五","星期六","星期日"]
    return {"success": True, "datetime": now.strftime("%Y-%m-%d %H:%M:%S"), "date": now.strftime("%Y-%m-%d"), "time": now.strftime("%H:%M:%S"), "weekday": wm[now.weekday()], "year": now.year, "month": now.month, "day": now.day}
