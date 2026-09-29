import json
import os


# ==================================================
# 记忆文件位置
# ==================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MEMORY_FILE = os.path.join(BASE_DIR, "memory.json")


# ==================================================
# 默认记忆结构
# ==================================================

DEFAULT_MEMORY = {
    "姓名": "",
    "工作": "",
    "爱好": [],
    "习惯": [],
    "其他重要信息": []
}


# ==================================================
# 读取记忆
# ==================================================

def get_memory():
    """读取小龙虾长期记忆"""

    # 没有文件就创建
    if not os.path.exists(MEMORY_FILE):

        save_memory(DEFAULT_MEMORY.copy())

        return DEFAULT_MEMORY.copy()


    try:

        with open(
            MEMORY_FILE,
            "r",
            encoding="utf-8"
        ) as f:

            memory = json.load(f)


        # 防止旧版本 memory.json 是列表
        if not isinstance(memory, dict):

            return DEFAULT_MEMORY.copy()


        return memory


    except Exception as e:

        print("❌ 读取记忆失败：", e)

        return DEFAULT_MEMORY.copy()


# ==================================================
# 保存整个记忆
# ==================================================

def save_memory(memory):
    """保存长期记忆"""

    try:

        with open(
            MEMORY_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            json.dump(
                memory,
                f,
                ensure_ascii=False,
                indent=4
            )

        return True


    except Exception as e:

        print("❌ 保存记忆失败：", e)

        return False


# ==================================================
# 设置某一项记忆
# ==================================================

def set_memory(key, value):
    """修改一项长期记忆"""

    memory = get_memory()

    memory[key] = value

    save_memory(memory)

    print(f"🧠 已记住：{key} = {value}")


# ==================================================
# 获取记忆文字
# ==================================================

def memory_to_text():
    """把记忆转换成 AI 容易理解的文字"""

    memory = get_memory()

    lines = []

    for key, value in memory.items():

        if value:

            lines.append(
                f"{key}：{value}"
            )

    if not lines:

        return "目前还没有长期记忆。"

    return "\n".join(lines)