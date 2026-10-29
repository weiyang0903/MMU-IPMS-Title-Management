import json, re, os, sys
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

# 清掉不可见的控制字符，保持文本干净
ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

def clean(text):
    return ILLEGAL.sub("", text).strip() if isinstance(text, str) else ""

projects = []
with open("output/progress.jsonl", encoding="utf-8") as f:
    for line in f:
        if not line.strip():
            continue
        try:
            r = json.loads(line)
        except json.JSONDecodeError:
            print("跳过一行无法读取的数据")
            continue
        # Status 原本长这样: "[ Available ]"，去掉括号和空格后做精确比较
        if re.sub(r"[\[\]\s]", "", r.get("Status", "")).lower() == "available":
            projects.append(r)

with open("output/available_projects.md", "w", encoding="utf-8") as f:
    for n, r in enumerate(projects, 1):
        pid = clean(r.get("Project ID", ""))
        title = clean(r.get("Project Title", ""))
        desc = clean(r.get("Description", ""))
        f.write(f"{n}. {pid} {title}\n")
        f.write(f"{desc}\n\n")

print("Available 项目数:", len(projects))
print("已保存 available_projects.md")