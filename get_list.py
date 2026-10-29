import sys
from pathlib import Path
from playwright.sync_api import sync_playwright
import pandas as pd
import time, os, random

# 1. 需要检查的所有文件路径
REQUIRED_FILES = [
    Path("output/login_state.json"),
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

# 如果 output 文件夹不存在就自动创建（存在则自动忽略）
Path("output").mkdir(parents=True, exist_ok=True)

LIST_URL = "https://fist-fyp.mmu.edu.my/qazxsw/p_student/stud_project_list.php"

with sync_playwright() as p:
    browser = p.chromium.launch(channel="msedge", headless=False)
    context = browser.new_context(storage_state="output/login_state.json")
    page = context.new_page()
    page.goto(LIST_URL)

    input("选好专业（如 AI），把 Show 选到 100，然后回到这里按回车...")

    all_rows = []
    seen = set()
    while True:
        rows = page.query_selector_all("#dt_table tbody tr")
        for r in rows:
            cells = [c.inner_text().strip() for c in r.query_selector_all("td")]
            if len(cells) >= 7 and cells[0] not in seen:
                seen.add(cells[0])
                all_rows.append(cells[:7])

        print("已读取:", len(all_rows), "|", page.inner_text("#dt_table_info"))

        nxt = page.query_selector("#dt_table_next")
        if nxt is None or "disabled" in (nxt.get_attribute("class") or ""):
            break
        nxt.click()
        time.sleep(random.uniform(3, 6))

    df = pd.DataFrame(
        all_rows,
        columns=["Project ID", "Title", "Type", "Programme", "Supervisor", "View", "Status"],
    )
    df.to_excel("output/project_list_all.xlsx", index=False)
    print("总共:", len(df), "已保存 project_list_all.xlsx")

    browser.close()