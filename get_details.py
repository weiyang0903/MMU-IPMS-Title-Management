from playwright.sync_api import sync_playwright
import pandas as pd
import time, random, json, os, sys
from pathlib import Path

# 1. 需要检查的所有文件路径
REQUIRED_FILES = [
    Path("output/login_state.json"),
    Path("output/project_list_all.xlsx"),
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

BASE = "https://fist-fyp.mmu.edu.my/qazxsw/p_student/stud_project_view.php?q={}&btn=back"
PROGRESS = "output/progress.jsonl"   # 每抓完一个就追加一行，中断后可接着抓
ONLY_AVAILABLE = False        # 改成 True 就只抓状态为 Available 的项目（分批期间不要改）
AUTO_BATCH_LIMIT = 100        # 单次运行最多抓多少个（手动输入的数字也不能超过它）
BROWSER = "msedge"            # 用 Chrome 就改成 "chrome"
DELAY = (3, 6)                # 每个项目之间随机等待的秒数范围，不要调得比 3 秒更小


def fix(text):
    # 修正 â€“ 之类的乱码，修不了就保持原样
    try:
        return text.encode("cp1252").decode("utf-8")
    except Exception:
        return text


df = pd.read_excel("output/project_list_all.xlsx")
if ONLY_AVAILABLE:
    df = df[df["Status"] == "Available"]
ids = df["Project ID"].astype(str).str.strip().tolist()

# 如果上次中断时最后一行没有写完（没有换行），先补一个换行，
# 避免新记录粘在残缺的那一行后面
if os.path.exists(PROGRESS) and os.path.getsize(PROGRESS) > 0:
    with open(PROGRESS, "rb") as f:
        f.seek(-1, 2)
        last = f.read(1)
    if last != b"\n":
        with open(PROGRESS, "a", encoding="utf-8") as f:
            f.write("\n")

# 读取已完成的进度（遇到无法读取的行就跳过，不会崩溃）
done = set()
bad_lines = 0
if os.path.exists(PROGRESS):
    with open(PROGRESS, encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            try:
                done.add(str(json.loads(line)["Project ID"]).strip())
            except (json.JSONDecodeError, KeyError, TypeError):
                bad_lines += 1
if bad_lines:
    print(f"⚠️ 注意：progress.jsonl 里有 {bad_lines} 行无法读取，已跳过。"
          "可能是上次中断时写了一半，用记事本打开检查一下。")

remaining = [i for i in ids if i not in done]

if not remaining:
    raise SystemExit("没有需要抓取的项目，已全部完成")

choice = input(
    f"本次要限制抓取数量吗？输入正整数限制本次运行；"
    f"直接按回车自动计算（最多 {AUTO_BATCH_LIMIT} 个）："
).strip()
if choice:
    try:
        batch = int(choice)
    except ValueError:
        raise SystemExit("请输入正整数，或直接按回车使用自动计算")
    if batch <= 0:
        raise SystemExit("抓取数量必须是正整数")
    if batch > AUTO_BATCH_LIMIT:
        raise SystemExit(f"本次最多只能抓取 {AUTO_BATCH_LIMIT} 个项目")
else:
    batch = min(len(remaining), AUTO_BATCH_LIMIT)

todo = remaining[:batch]
runs_after = -(-(len(remaining) - len(todo)) // AUTO_BATCH_LIMIT)  # 向上取整
print(f"总共 {len(ids)} 个，已完成 {len(done)} 个，还剩 {len(remaining)} 个")
print(f"本次要抓 {len(todo)} 个，抓完后大约还需要再运行 {runs_after} 次")

with sync_playwright() as p:
    browser = p.chromium.launch(channel=BROWSER, headless=False)
    context = browser.new_context(storage_state="output/login_state.json")
    page = context.new_page()

    for n, pid in enumerate(todo, 1):
        try:
            resp = page.goto(BASE.format(pid), timeout=30000)
        except Exception as e:
            print("打开页面出错，停止:", e)
            break

        if resp is None or resp.status != 200:
            print("异常状态码:", resp.status if resp else None, "，停止")
            break
        if page.query_selector("input[type=password]"):
            print("跳回登录页，登录状态可能过期，停止")
            break

        # 详情页是「左列字段名 | 右列内容」的表格
        record = {"Project ID": pid}
        for tr in page.query_selector_all("tr"):
            tds = tr.query_selector_all("td")
            if len(tds) == 2:
                key = fix(tds[0].inner_text().strip().replace("\n", " "))
                val = fix(tds[1].inner_text().strip())
                if key == "Project Information":
                    key = "Status"
                if key and key not in record:
                    record[key] = val

        if "Description" not in record:
            print(pid, "读不到 Description，页面可能异常，停止")
            break

        with open(PROGRESS, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")

        print(f"[{n}/{len(todo)}] {pid} 完成")
        time.sleep(random.uniform(*DELAY))

    browser.close()

print("本次结束。进度保存在", PROGRESS)