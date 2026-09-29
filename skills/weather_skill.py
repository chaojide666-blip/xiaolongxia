# -*- coding: utf-8 -*-
"""
饭加鱼：天气查询技能（和风天气版）
"""

import os
import requests


def get_weather(city: str = "深圳"):
    """
    查询指定城市的当前天气。
    city 可以是中文城市名，如"深圳"、"北京"、"上海"。
    """
    if not city or not city.strip():
        city = "深圳"

    city = city.strip()

    api_key = os.getenv("QWEATHER_API_KEY")
    api_host = os.getenv("QWEATHER_API_HOST")

    if not api_key or not api_host:
        return {
            "success": False,
            "error": "没有找到 QWEATHER_API_KEY 或 QWEATHER_API_HOST 环境变量。",
        }

    try:
        # 第一步：通过 GeoAPI 查找城市 ID
        geo_url = f"https://{api_host}/geo/v2/city/lookup"
        geo_params = {"location": city, "key": api_key}
        geo_resp = requests.get(geo_url, params=geo_params, timeout=10)
        geo_resp.raise_for_status()
        geo_data = geo_resp.json()

        if geo_data.get("code") != "200" or not geo_data.get("location"):
            return {"success": False, "error": f"找不到城市：{city}"}

        location_id = geo_data["location"][0]["id"]
        location_name = geo_data["location"][0]["name"]

        # 第二步：查询实时天气
        weather_url = f"https://{api_host}/v7/weather/now"
        weather_params = {"location": location_id, "key": api_key}
        weather_resp = requests.get(weather_url, params=weather_params, timeout=10)
        weather_resp.raise_for_status()
        weather_data = weather_resp.json()

        if weather_data.get("code") != "200":
            return {"success": False, "error": f"查询天气失败：{weather_data.get('code')}"}

        now = weather_data.get("now", {})

        return {
            "success": True,
            "city": location_name,
            "temperature": now.get("temp", ""),
            "feels_like": now.get("feelsLike", ""),
            "weather": now.get("text", ""),
            "wind_direction": now.get("windDir", ""),
            "wind_scale": now.get("windScale", ""),
            "humidity": now.get("humidity", ""),
            "pressure": now.get("pressure", ""),
            "visibility": now.get("vis", ""),
            "update_time": weather_data.get("updateTime", ""),
        }

    except Exception as e:
        return {"success": False, "error": f"查询天气失败：{e}"}