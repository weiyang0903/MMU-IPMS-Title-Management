from pathlib import Path
from playwright.sync_api import sync_playwright

LIST_URL = "https://fist-fyp.mmu.edu.my/qazxsw/p_student/stud_project_list.php"

# 如果 output 文件夹不存在就自动创建（存在则自动忽略）
Path("output").mkdir(parents=True, exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch(channel="msedge", headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto(LIST_URL)

    input("请手动登录，看到 Project List 后，回到这里按回车...")

    context.storage_state(path="output/login_state.json")
    print("已保存 login_state.json")
    browser.close()