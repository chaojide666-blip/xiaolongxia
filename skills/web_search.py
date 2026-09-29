# -*- coding: utf-8 -*-
"""
饭加鱼：联网搜索技能 v3（含抓正文）
"""

import os
import requests
from bs4 import BeautifulSoup


def _fetch_article(url, max_chars=3000):
    """抓取网页正文。"""
    if not url:
        return ""

    try:
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            )
        }
        resp = requests.get(url, headers=headers, timeout=10)
        resp.encoding = resp.apparent_encoding or "utf-8"

        soup = BeautifulSoup(resp.text, "html.parser")

        # 去掉没用的标签
        for tag in soup(["script", "style", "nav", "footer", "aside", "header", "form", "iframe"]):
            tag.decompose()

        # 优先找 article 或 main
        article = soup.find("article") or soup.find("main")

        if article:
            text = article.get_text(separator="\n", strip=True)
        else:
            # 退而求其次：所有 p 标签
            paragraphs = soup.find_all("p")
            text = "\n".join(p.get_text(strip=True) for p in paragraphs)

        # 清理多余空行
        lines = [line.strip() for line in text.split("\n") if line.strip()]
        text = "\n".join(lines)

        return text[:max_chars]

    except Exception as e:
        print(f"抓正文失败 {url}: {e}")
        return ""


def web_search(query: str, max_results: int = 5):
    """
    使用 Tavily 搜索网页，并抓取前几篇的正文。
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
        raw_results = response.get("results", [])

        # 前 3 篇抓正文，后面的只用摘要
        for i, r in enumerate(raw_results):
            url = r.get("url", "")
            title = r.get("title", "")
            summary = r.get("content", "")

            full_text = ""
            if i < 3 and url:
                print(f"📄 抓取正文：{url}")
                full_text = _fetch_article(url, max_chars=3000)

            results.append({
                "title": title,
                "url": url,
                "summary": summary,
                "content": full_text if full_text else summary,
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
