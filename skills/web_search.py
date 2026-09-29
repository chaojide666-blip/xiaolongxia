# -*- coding: utf-8 -*-
"""
饭加鱼: 联网搜索技能 v2 (Tavily 版)
"""

import os


def web_search(query: str, max_results: int = 5):
    """
    使用 Tavily 搜索网页并返回结果摘要。
    """
    if not query or not query.strip():
        return {"success": False, "error": "搜索内容不能为空。"}

    api_key = os.getenv("TAVILY_API_KEY")
    if not api_key:
        return {
            "success": False,
            "error": "没有找到 TAVILY_API_KEY，请先设置环境变量。",
        }

    try:
        from tavily import TavilyClient

        client = TavilyClient(api_key=api_key)
        response = client.search(
            query=query,
            max_results=max_results,
        )

        results = []
        for r in response.get("results", []):
            results.append({
                "title": r.get("title", ""),
                "url": r.get("url", ""),
                "body": r.get("content", ""),
            })

        if not results:
            return {"success": False, "error": "没有搜索到相关内容。"}

        return {
            "success": True,
            "query": query,
            "results": results,
        }

    except Exception as e:
        return {"success": False, "error": f"搜索失败：{e}"}