import os
import re

from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor


# =========================================================
# DeepSeek API
# =========================================================

api_key = os.getenv("DEEPSEEK_API_KEY")

if not api_key:
    raise ValueError(
        "没有找到 DEEPSEEK_API_KEY，请先设置 Windows 环境变量。"
    )

client = OpenAI(
    api_key=api_key,
    base_url="https://api.deepseek.com"
)


# =========================================================
# 页面尺寸
# =========================================================

SLIDE_W = 13.333
SLIDE_H = 7.5


# =========================================================
# 颜色
# =========================================================

BG = RGBColor(248, 249, 252)
WHITE = RGBColor(255, 255, 255)

DARK = RGBColor(32, 38, 48)
TEXT = RGBColor(70, 78, 90)
LIGHT_TEXT = RGBColor(150, 156, 166)

BLUE = RGBColor(57, 117, 225)
PURPLE = RGBColor(116, 103, 214)
GREEN = RGBColor(70, 157, 120)

BORDER = RGBColor(220, 224, 230)

LIGHT_BLUE = RGBColor(235, 242, 255)
LIGHT_PURPLE = RGBColor(241, 239, 255)
LIGHT_GREEN = RGBColor(237, 248, 242)



# =========================================================
# V1.2 PPT专业化主题系统
# =========================================================

PPT_STYLES = {
    "商务风": {"primary": BLUE, "bg": BG},
    "科技风": {"primary": PURPLE, "bg": RGBColor(245,245,252)},
    "简约风": {"primary": GREEN, "bg": RGBColor(250,250,250)},
}

def apply_style(style="商务风"):
    """应用PPT主题（兼容旧版本调用）"""
    global BLUE, BG
    cfg = PPT_STYLES.get(style, PPT_STYLES["商务风"])
    BLUE = cfg["primary"]
    BG = cfg["bg"]

# =========================================================
# 基础函数
# =========================================================

def add_text(
    slide,
    text,
    x,
    y,
    w,
    h,
    size=20,
    bold=False,
    color=TEXT,
    align=PP_ALIGN.LEFT
):
    """添加文字"""

    box = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h)
    )

    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP

    p = tf.paragraphs[0]
    p.text = str(text)
    p.alignment = align

    if p.runs:
        run = p.runs[0]
        run.font.name = "Microsoft YaHei"
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color

    return box


def set_background(slide, color=BG):
    """设置背景"""

    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_top_line(slide):
    """标题下面的装饰线"""

    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.65),
        Inches(1.12),
        Inches(1.15),
        Inches(0.06)
    )

    line.fill.solid()
    line.fill.fore_color.rgb = BLUE
    line.line.fill.background()


def add_page_title(slide, title, subtitle=""):
    """添加统一页面标题"""

    add_text(
        slide,
        title,
        0.65,
        0.35,
        11.5,
        0.65,
        size=27,
        bold=True,
        color=DARK
    )

    if subtitle:
        add_text(
            slide,
            subtitle,
            1.95,
            0.92,
            9.5,
            0.3,
            size=10,
            color=LIGHT_TEXT
        )

    add_top_line(slide)


def add_page_number(slide, page_number):
    """页码"""

    add_text(
        slide,
        f"{page_number:02d}",
        11.8,
        7.0,
        0.55,
        0.25,
        size=9,
        color=LIGHT_TEXT,
        align=PP_ALIGN.RIGHT
    )


def add_card(
    slide,
    x,
    y,
    w,
    h,
    title,
    body,
    accent=BLUE,
    background=WHITE
):
    """信息卡片"""

    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h)
    )

    card.fill.solid()
    card.fill.fore_color.rgb = background

    card.line.color.rgb = BORDER
    card.line.width = Pt(1)

    # 顶部色条
    stripe = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(0.08)
    )

    stripe.fill.solid()
    stripe.fill.fore_color.rgb = accent
    stripe.line.fill.background()

    add_text(
        slide,
        title,
        x + 0.22,
        y + 0.25,
        w - 0.44,
        0.48,
        size=17,
        bold=True,
        color=DARK
    )

    add_text(
        slide,
        body,
        x + 0.22,
        y + 0.82,
        w - 0.44,
        h - 1.0,
        size=12.5,
        color=TEXT
    )


def add_keyword_box(slide, keywords):
    """关键词区域"""

    if not keywords:
        return

    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.75),
        Inches(6.45),
        Inches(11.8),
        Inches(0.45)
    )

    box.fill.solid()
    box.fill.fore_color.rgb = LIGHT_BLUE
    box.line.fill.background()

    add_text(
        slide,
        "关键词  " + "   ·   ".join(keywords),
        1.0,
        6.53,
        11.25,
        0.28,
        size=10.5,
        bold=True,
        color=BLUE
    )


# =========================================================
# AI 生成内容
# =========================================================

def generate_content(topic, pages):
    """
    让 AI 生成更丰富的 PPT 内容
    """

    prompt = f"""
请制作一份主题为《{topic}》的中文 PPT，共 {pages} 页。

这是一份正式展示型 PPT，请不要只写简单的三条短句。
需要有完整、丰富、适合展示的内容。

必须严格按照以下格式输出，不要改变格式：

第1页：标题
副标题：一句简洁的副标题
核心观点：这一页最重要的一句话
要点1标题：小标题
要点1内容：对这个小标题进行2到3句话的解释
要点2标题：小标题
要点2内容：对这个小标题进行2到3句话的解释
要点3标题：小标题
要点3内容：对这个小标题进行2到3句话的解释
关键词：关键词1、关键词2、关键词3

第2页：标题
核心观点：这一页最重要的一句话
要点1标题：小标题
要点1内容：详细说明
要点2标题：小标题
要点2内容：详细说明
要点3标题：小标题
要点3内容：详细说明
关键词：关键词1、关键词2、关键词3

一直生成到第{pages}页。

特别要求：

1. 第1页必须是专业封面。
2. 最后一页必须是总结页。
3. 中间页面必须有清晰逻辑。
4. 每个要点必须有“标题 + 详细说明”。
5. 内容不能过于简短。
6. 每页的信息量要足够，但不要写成长文章。
7. 适合 PowerPoint 展示，而不是论文。
8. 尽量使用专业但容易理解的中文。
10. 根据主题适当设计流程、时间轴、对比分析、案例展示页面。
9. 不要输出任何解释。
"""

    print("🧠 正在向饭加鱼的大脑请求 PPT 内容...")

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )

    content = response.choices[0].message.content

    if not content:
        raise ValueError("AI 没有生成 PPT 内容。")

    return content


# =========================================================
# 解析 AI 内容
# =========================================================

def parse_content(content):
    """解析 AI 生成的结构化内容"""

    slides = []

    current = None

    for raw_line in content.splitlines():

        line = raw_line.strip()

        if not line:
            continue

        # -----------------------------
        # 新页面
        # -----------------------------

        if line.startswith("第") and "页：" in line:

            if current:
                slides.append(current)

            title = line.split("：", 1)[1].strip()

            current = {
                "title": title,
                "subtitle": "",
                "core": "",
                "cards": [],
                "keywords": []
            }

            continue

        # 没有当前页面时忽略
        if current is None:
            continue

        # -----------------------------
        # 副标题
        # -----------------------------

        if line.startswith("副标题："):

            current["subtitle"] = line.split(
                "：",
                1
            )[1].strip()

            continue

        # -----------------------------
        # 核心观点
        # -----------------------------

        if line.startswith("核心观点："):

            current["core"] = line.split(
                "：",
                1
            )[1].strip()

            continue

        # -----------------------------
        # 要点标题
        # -----------------------------

        match_title = re.match(
            r"要点([1-4])标题：(.*)",
            line
        )

        if match_title:

            title_text = match_title.group(2).strip()

            current["cards"].append(
                {
                    "title": title_text,
                    "content": ""
                }
            )

            continue

        # -----------------------------
        # 要点内容
        # -----------------------------

        match_content = re.match(
            r"要点([1-4])内容：(.*)",
            line
        )

        if match_content:

            content_text = match_content.group(2).strip()

            index = int(
                match_content.group(1)
            ) - 1

            if index >= len(current["cards"]):

                current["cards"].append(
                    {
                        "title": f"要点 {index + 1}",
                        "content": content_text
                    }
                )

            else:

                current["cards"][index]["content"] = (
                    content_text
                )

            continue

        # -----------------------------
        # 关键词
        # -----------------------------

        if line.startswith("关键词："):

            keyword_text = line.split(
                "：",
                1
            )[1].strip()

            current["keywords"] = [
                x.strip()
                for x in keyword_text.replace(
                    "；",
                    "、"
                ).split("、")
                if x.strip()
            ]

            continue

    if current:
        slides.append(current)

    return slides


# =========================================================
# 创建封面
# =========================================================

def create_cover(prs, info):
    """创建封面"""

    slide = prs.slides.add_slide(
        prs.slide_layouts[6]
    )

    set_background(
        slide,
        DARK
    )

    # 左侧装饰
    left_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(0.18),
        Inches(SLIDE_H)
    )

    left_bar.fill.solid()
    left_bar.fill.fore_color.rgb = BLUE
    left_bar.line.fill.background()

    # 顶部标签
    add_text(
        slide,
        "AI PRESENTATION",
        0.8,
        1.15,
        4.0,
        0.35,
        size=12,
        bold=True,
        color=BLUE
    )

    # 标题
    add_text(
        slide,
        info["title"],
        0.8,
        1.9,
        11.4,
        1.4,
        size=36,
        bold=True,
        color=WHITE
    )

    # 副标题
    if info["subtitle"]:

        add_text(
            slide,
            info["subtitle"],
            0.82,
            3.5,
            10.2,
            0.8,
            size=18,
            color=RGBColor(215, 220, 228)
        )

    # 装饰圆形
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(10.7),
        Inches(4.9),
        Inches(1.35),
        Inches(1.35)
    )

    circle.fill.solid()
    circle.fill.fore_color.rgb = BLUE
    circle.line.fill.background()

    add_text(
        slide,
        "饭加鱼",
        10.77,
        5.34,
        1.2,
        0.3,
        size=12,
        bold=True,
        color=WHITE,
        align=PP_ALIGN.CENTER
    )

    # 底部
    add_text(
        slide,
        "由饭加鱼 AI 生成",
        0.82,
        6.72,
        4.0,
        0.3,
        size=10,
        color=LIGHT_TEXT
    )

    return slide


# =========================================================
# 创建内容页
# =========================================================

def create_content_slide(prs, info, page_number):
    """创建内容页面"""

    slide = prs.slides.add_slide(
        prs.slide_layouts[6]
    )

    set_background(
        slide,
        BG
    )

    add_page_title(
        slide,
        info["title"],
        info.get("subtitle", "")
    )

    # -----------------------------------------------------
    # 核心观点区域
    # -----------------------------------------------------

    if info.get("core"):

        core_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.75),
            Inches(1.45),
            Inches(11.8),
            Inches(1.05)
        )

        core_box.fill.solid()
        core_box.fill.fore_color.rgb = LIGHT_BLUE
        core_box.line.fill.background()

        add_text(
            slide,
            "核心观点",
            1.0,
            1.67,
            1.2,
            0.3,
            size=11,
            bold=True,
            color=BLUE
        )

        add_text(
            slide,
            info["core"],
            2.1,
            1.58,
            9.95,
            0.55,
            size=15,
            bold=True,
            color=DARK
        )

    # -----------------------------------------------------
    # 三卡片
    # -----------------------------------------------------

    cards = info.get(
        "cards",
        []
    )

    cards = cards[:3]

    accent_colors = [
        (BLUE, LIGHT_BLUE),
        (PURPLE, LIGHT_PURPLE),
        (GREEN, LIGHT_GREEN)
    ]

    card_y = 2.85

    if len(cards) == 3:

        card_x = [
            0.65,
            4.55,
            8.45
        ]

        for i, card_data in enumerate(cards):

            accent, light_bg = accent_colors[i]

            add_card(
                slide,
                card_x[i],
                card_y,
                3.65,
                3.25,
                card_data["title"],
                card_data["content"],
                accent,
                light_bg
            )

    elif len(cards) == 2:

        add_card(
            slide,
            0.8,
            card_y,
            5.75,
            3.25,
            cards[0]["title"],
            cards[0]["content"],
            BLUE,
            LIGHT_BLUE
        )

        add_card(
            slide,
            6.8,
            card_y,
            5.75,
            3.25,
            cards[1]["title"],
            cards[1]["content"],
            PURPLE,
            LIGHT_PURPLE
        )

    elif len(cards) == 1:

        add_card(
            slide,
            1.0,
            card_y,
            11.3,
            3.25,
            cards[0]["title"],
            cards[0]["content"],
            BLUE,
            LIGHT_BLUE
        )

    # -----------------------------------------------------
    # 关键词
    # -----------------------------------------------------

    add_keyword_box(
        slide,
        info.get("keywords", [])
    )

    add_page_number(
        slide,
        page_number
    )

    return slide


# =========================================================
# 创建总结页
# =========================================================

def create_summary_slide(prs, info, page_number):
    """最后总结页"""

    slide = prs.slides.add_slide(
        prs.slide_layouts[6]
    )

    set_background(
        slide,
        DARK
    )

    add_text(
        slide,
        info["title"],
        0.75,
        0.65,
        11.5,
        0.7,
        size=30,
        bold=True,
        color=WHITE
    )

    add_text(
        slide,
        "关键结论与下一步",
        0.78,
        1.38,
        5.0,
        0.35,
        size=11,
        color=LIGHT_TEXT
    )

    cards = info.get(
        "cards",
        []
    )[:3]

    card_colors = [
        BLUE,
        PURPLE,
        GREEN
    ]

    y_positions = [
        2.0,
        3.5,
        5.0
    ]

    for i, card_data in enumerate(cards):

        if i >= 3:
            break

        # 小圆点
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(0.9),
            Inches(y_positions[i] + 0.08),
            Inches(0.32),
            Inches(0.32)
        )

        dot.fill.solid()
        dot.fill.fore_color.rgb = card_colors[i]
        dot.line.fill.background()

        add_text(
            slide,
            card_data["title"],
            1.45,
            y_positions[i] - 0.02,
            2.2,
            0.4,
            size=16,
            bold=True,
            color=WHITE
        )

        add_text(
            slide,
            card_data["content"],
            3.25,
            y_positions[i] - 0.05,
            8.5,
            0.85,
            size=12.5,
            color=RGBColor(215, 220, 228)
        )

    add_text(
        slide,
        "由饭加鱼 AI 生成",
        0.8,
        6.8,
        4.0,
        0.25,
        size=9,
        color=LIGHT_TEXT
    )

    add_page_number(
        slide,
        page_number
    )

    return slide


# =========================================================
# 生成 PPT
# =========================================================

def build_ppt(slides_data, filename):
    """把结构化数据生成 PPT"""

    prs = Presentation()

    # 16:9
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    total_pages = len(slides_data)

    for index, info in enumerate(slides_data):

        page_number = index + 1

        # 第一页：封面
        if page_number == 1:

            create_cover(
                prs,
                info
            )

            continue

        # 最后一页：总结
        if page_number == total_pages:

            create_summary_slide(
                prs,
                info,
                page_number
            )

            continue

        # 中间页
        create_content_slide(
            prs,
            info,
            page_number
        )

    prs.save(
        filename
    )


# =========================================================
# 主函数
# =========================================================

def make_ppt(topic, pages=6, style="商务风"):
    """生成 PPT（V1.2专业版）"""

    apply_style(style)

    try:
        pages = int(pages)
    except ValueError:
        pages = 6

    # 最少5页，最多15页
    pages = max(
        5,
        min(pages, 15)
    )

    print()
    print("=" * 50)
    print("🦞 饭加鱼 PPT 制作工具")
    print("=" * 50)

    print()
    print("主题：", topic)
    print("页数：", pages)

    print()
    print("🧠 第一步：生成详细内容...")

    ai_content = generate_content(
        topic,
        pages
    )

    print("✅ 内容生成完成")

    print()
    print("🔍 第二步：解析内容...")

    slides_data = parse_content(
        ai_content
    )

    if len(slides_data) < 3:

        raise ValueError(
            "AI 生成的页面数量太少，无法制作 PPT。"
        )

    # 如果 AI 超额生成
    slides_data = slides_data[:pages]

    print(
        f"✅ 成功解析 {len(slides_data)} 页"
    )

    # -----------------------------------------------------
    # 创建输出文件夹
    # -----------------------------------------------------

    output_dir = os.path.join(
        os.path.dirname(
            os.path.abspath(__file__)
        ),
        "生成的PPT"
    )

    os.makedirs(
        output_dir,
        exist_ok=True
    )

    # -----------------------------------------------------
    # 安全处理文件名
    # -----------------------------------------------------

    safe_topic = topic

    for char in '\\/:*?"<>|':

        safe_topic = safe_topic.replace(
            char,
            "_"
        )

    filename = os.path.join(
        output_dir,
        f"{safe_topic}.pptx"
    )

    # -----------------------------------------------------
    # 创建 PPT
    # -----------------------------------------------------

    print()
    print("🎨 第三步：设计 PPT...")

    build_ppt(
        slides_data,
        filename
    )

    print()
    print("🎉 PPT 制作完成！")
    print()
    print("📁 文件位置：")
    print(filename)
    print()
    print("=" * 50)

    return filename


# =========================================================
# 独立运行
# =========================================================

if __name__ == "__main__":

    topic = input(
        "请输入PPT主题："
    ).strip()

    pages = input(
        "请输入PPT页数："
    ).strip()

    if not topic:
        topic = "人工智能的发展"

    if not pages:
        pages = "6"

    make_ppt(
        topic,
        pages
    )