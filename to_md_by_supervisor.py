import json, re, os, sys
from collections import Counter
from pathlib import Path

# 1. 需要检查的所有文件路径
REQUIRED_FILES = [
    Path("output/progress.jsonl"),
]

# 2. 找出所有缺失的文件
missing_files = [file for file in REQUIRED_FILES if not file.exists()]

# 3. 如果有缺失的文件，直接提示并退出
if missing_files:
    print("❌ 错误：缺少运行所需的关键文件！")
    print("缺失的文件如下：")
    for file in missing_files:
        print(f"  - {file}")
    sys.exit(1)

# 所有文件都存在，继续往下执行
print("✅ 所有必要文件已就绪，开始运行...")

# ===== 设置区 =====
# 想要哪些导师就写在这里，可以一个或多个。
# 不区分大小写，只要名字「包含」这段文字就算匹配（所以不用写 (Dr) 之类的后缀）。
# 留空 [] 表示不筛选导师，输出全部 Available 项目。
SUPERVISORS = [
    "Tan Shing Chiang",
    "Goh Pey Yun",
]

OUTPUT = "output/available_projects.md"
# ==================

# 清掉不可见的控制字符，保持文本干净
ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

def clean(text):
    return ILLEGAL.sub("", text).strip() if isinstance(text, str) else ""

# 1. 读取，只保留 Available（Status 原本长这样: "[ Available ]"，这里去掉括号和空格后做精确比较）
available = []
with open("output/progress.jsonl", encoding="utf-8") as f:
    for line in f:
        if not line.strip():
            continue
        try:
            r = json.loads(line)
        except json.JSONDecodeError:
            print("跳过一行无法读取的数据")
            continue
        if re.sub(r"[\[\]\s]", "", r.get("Status", "")).lower() == "available":
            available.append(r)

# 2. 显示 Available 项目里每位导师各有几个，方便你挑名字
counts = Counter(clean(r.get("Supervisor name:", "")) for r in available)
print(f"Available 项目共 {len(available)} 个，导师分布：")
for name, c in sorted(counts.items(), key=lambda x: (-x[1], x[0])):
    print(f"  {c:>3}  {name}")
print()

# 3. 按导师筛选
keywords = [s.lower().strip() for s in SUPERVISORS if s.strip()]
if keywords:
    projects = [
        r for r in available
        if any(k in clean(r.get("Supervisor name:", "")).lower() for k in keywords)
    ]
    # 提醒：哪些关键字一个项目都没匹配到（可能拼错，或该导师没有 Available 项目）
    for k, original in zip(keywords, [s for s in SUPERVISORS if s.strip()]):
        if not any(k in clean(r.get("Supervisor name:", "")).lower() for r in available):
            print(f"注意：找不到匹配「{original}」的 Available 项目")
else:
    projects = available

# 4. 输出 md，格式：「1. Project ID Project Title」换行接 Description
with open(OUTPUT, "w", encoding="utf-8") as f:
    for n, r in enumerate(projects, 1):
        pid = clean(r.get("Project ID", ""))
        title = clean(r.get("Project Title", ""))
        desc = clean(r.get("Description", ""))
        f.write(f"{n}. {pid} {title}\n")
        f.write(f"{desc}\n\n")

print(f"筛选后项目数: {len(projects)}")
print(f"已保存 {OUTPUT}")