import pandas as pd
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

KEEP = ["Project ID", "Status", "Project Title", "Description", "Objectives",
        "Project type", "Specific area of research", "Skill set required",
        "Expected deliverable", "Project schedule (List of steps for the project)",
        "Innovative component/ project potential", "This project suitable for programme",
        "Proposed by", "Supervisor name:", "Email:"]

# Excel 不允许的控制字符（肉眼看不见，从 Word/PDF 复制粘贴时可能带进来）
ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

def clean(text):
    return ILLEGAL.sub("", text) if isinstance(text, str) else text

rows = []
with open("output/progress.jsonl", encoding="utf-8") as f:
    for line in f:
        if not line.strip():
            continue
        try:
            r = json.loads(line)
        except json.JSONDecodeError:
            print("跳过一行无法读取的数据")
            continue
        rows.append({k: clean(r.get(k, "")) for k in KEEP})

df = pd.DataFrame(rows)
df.to_excel("output/projects_with_description.xlsx", index=False)
print("已保存 projects_with_description.xlsx，共", len(df), "个")