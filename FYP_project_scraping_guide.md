# IPMS FYP 项目数据爬取指南（个人记录）

> **仅供个人记录与使用，不可以恶意商用或者读取。**
> 目的：把 IPMS 里的 FYP 项目（Project ID、标题、描述等）整理成本地文件，方便自己比较选题。
> 抓取下来的数据、登录文件都只留在自己电脑上，不上传、不转发。

---

## 0. 总览

### 整体流程

```
手动登录一次 → 保存登录状态 → 抓取项目列表(翻页) → 逐个抓取详情页 → 导出 Excel / Markdown / 用搜索页面查看 → 之后按需更新 → 清理
```

### 用到的文件

| 文件 | 作用 |
|---|---|
| `save_login.py` | 手动登录一次，保存登录状态到 `output/login_state.json` |
| `get_list.py` | 抓取全部项目列表（翻页），生成 `output/project_list_all.xlsx` |
| `get_details.py` | 逐个打开详情页，抓取描述等字段，边抓边存进 `output/progress.jsonl` |
| `export.py` | 把 `output/progress.jsonl` 整理成 `output/projects_with_description.xlsx` |
| `to_md.py` | 只提取 Available 的项目，生成 `available_projects.md`（第 7 节） |
| `to_md_by_supervisor.py` | 在 Available 的基础上再按导师筛选，生成 `output/available_projects.md`（第 8 节） |
| `fyp_search.html` | 读取 `output/progress.jsonl` 的搜索 / 筛选 / 收藏页面（第 9 节） |
| `update_status.py` | 用最新列表同步状态，处理新增和被删除的项目（第 10 节） |

### 关键信息

- 项目列表页：`https://fist-fyp.mmu.edu.my/qazxsw/p_student/stud_project_list.php`
- 详情页：`https://fist-fyp.mmu.edu.my/qazxsw/p_student/stud_project_view.php?q=项目ID&btn=back`（`q=` 后面就是 Project ID）
- 列表表格的 id 是 `dt_table`，分页信息的 id 是 `dt_table_info`，下一页按钮的 id 是 `dt_table_next`
- 使用电脑自带的 **Edge / Chrome** 浏览器（`channel="msedge" / channel="chrome"`），**不需要另外下载浏览器**
- 当时一个专业共 391 个项目，Show 最多只能选 100，所以需要翻 4 页

---

## 1. 安全原则（每次使用前先看一遍）

1. **只用自己的账号**，手动登录，密码不写进代码。
2. **顺序抓取、不并发**，每个详情页之间随机间隔 3～6 秒。
3. **只读不改**：脚本只打开页面、读取文字，不点任何按钮（申请、提交、保存等）。
4. **分批抓取**：每次最多 100 个，批与批之间隔几个小时，避开选题高峰期。
5. **出现异常就停**：403、429、跳回登录页、读不到 Description，都不要硬重试。
6. `login_state.json` 等同于登录凭证：**不分享、不上传，用完就删**。
7. 要诚实地知道：学校服务器的访问日志和项目的 View 浏览次数你这边删不掉，所以风险是"很低"而不是"零"。
8. **不能影响网站运行**：详细规则见第 11 节，那是整份指南里最重要的一节。

---

## 2. 环境准备（只需做一次）

### 2.1 确认 Python

打开命令行（`Win + R` → 输入 `cmd` → 回车）：

```
python --version
```

显示 `Python 3.x.x` 即可。如果提示找不到，到 https://www.python.org/downloads/ 安装，**安装时务必勾选 "Add Python to PATH"**。

### 2.2 建立工作文件夹
前往你所在的文件夹位置。例子：
```
cd 你所在的文件夹位置
```

例如 `C:\Users\User\Desktop\t`，脚本放在这个文件夹里，数据会自动存进它下面的`output`子文件夹。

```
cd C:\Users\User\Desktop\t
```

### 2.3 建立虚拟环境并安装依赖

```
python -m venv myenv
myenv\Scripts\activate
pip install playwright pandas openpyxl
```

- 成功激活后，命令行最前面会出现 `(myenv)`。
- **以后每次打开新的命令行窗口，都要先 `cd` 到文件夹，再执行 `myenv\Scripts\activate`，然后才运行脚本。**
- 不需要执行 `playwright install`，因为脚本直接使用电脑自带的 Chrome\Edge。

---

## 3. 第 1 步：保存登录状态

### 3.1 脚本 `save_login.py`

```python
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
```

### 3.2 运行

```
python save_login.py
```

会弹出 Edge 窗口。用自己的账号手动登录，看到 Project List 表格后，回到命令行按回车。文件夹里会出现 `login_state.json`。

> 这里要注意的是如果要使用的是chrome浏览器就把`msedge`改成`chrome`。
例子：
```python
browser = p.chromium.launch(channel="msedge", headless=False)
```
改成
```python
browser = p.chromium.launch(channel="chrome", headless=False)
```
>要修改的地方有三个，分别是`save_login.py`和`get_list.py`里的`channel="msedge"`，以及`get_details.py`开头的`BROWSER = "msedge"`。

### 3.3 注意

- 登录状态**会过期**。以后如果脚本提示"跳回登录页"，重新运行一次 `save_login.py` 即可。
- 在网站里点了 Logout，保存的登录状态也会失效。

---

## 4. 第 2 步：抓取项目列表（含翻页）

### 4.1 脚本 `get_list.py`

```python
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
```

### 4.2 运行

```
python get_list.py
```

浏览器打开后，在网页里手动做两件事，再回命令行按回车：

1. **Select Programme** 选成自己的专业（例如 Artificial Intelligence (AI)）
2. **Show** 选到 100

### 4.3 检查结果

最后一行 `总共: N` 的 N 应该等于页面右下角 "of N entries" 的数字（当时是 391）。数字对得上，说明翻页没有漏。

---

## 5. 第 3 步：抓取详情页（描述等字段）

列表页没有描述，描述在每个项目的详情页里，所以要逐个打开。

### 5.1 脚本 `get_details.py`

```python
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
DELAY = (3, 6)                # 每个项目之间随机等待的秒数范围，不要调得比 2 秒更小


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
```

### 5.2 先小规模测试（强烈建议）

第一次运行前，运行脚本时在数量提示中输入 `3`，运行：

```
python get_details.py
```

然后打开 `output/progress.jsonl`，确认 `Description`、`Objectives` 等字段内容正确。没问题后，下一次运行时直接按回车使用自动计算；这 3 个已经存好，之后会自动跳过。

### 5.3 正式分批抓取

```
python get_details.py
```

- 每次最多抓 100 个，大约 **9～10 分钟**（这是估算，实际取决于网速和服务器）。
- 391 个大约分 4 次跑完，**每批之间隔几个小时**，或者分早上、晚上、不同天。
- 每次运行都会询问数量：输入正整数只限制这一次运行；直接按回车则自动取“剩余项目数”和 100 中较小的值。
- 每次运行开头会显示 总共 391 个，已完成 XXX 个，还剩 YYY 个，以及 本次要抓 N 个，抓完后大约还需要再运行 M 次
- 全部抓完后，再运行会显示没有需要抓取的项目，已全部完成，并且不会打开浏览器
> 注意！！之所以不要一次过抓取是怕被误判成机器人。

### 5.4 分批注意事项

1. **`progress.jsonl` 不要删、不要改名、不要移动**，它必须和脚本在同一个文件夹，否则会从头开始抓。
2. 中途想停，按 `Ctrl+C`。最多只丢掉正在抓的那一个，下次会重新抓。
3. `ONLY_AVAILABLE` 在整个分批期间保持一致。
4. 隔了很久再跑，如果提示"跳回登录页"，先重新运行 `save_login.py`。

### 5.5 脚本会自动停止的情况

| 情况 | 屏幕提示 | 常见原因 |
|---|---|---|
| 抓取数量 > AUTO_BATCH_LIMIT | `本次最多只能抓取 AUTO_BATCH_LIMIT 个项目` | 输入的数字超过 AUTO_BATCH_LIMIT |
| 页面打不开或超过 30 秒 | `打开页面出错，停止` | 断网、网站卡住 |
| 状态码不是 200 | `异常状态码: xxx ，停止` | 被拒绝、请求太快、服务器出错 |
| 页面出现密码输入框 | `跳回登录页...停止` | 登录过期、被踢出 |
| 读不到 Description | `读不到 Description...停止` | 页面结构变了、返回了错误页或验证页 |

停止后的处理：

- **403 / 429**：至少等几个小时甚至隔天再继续，不要硬试。
- **跳回登录页**：重新运行 `save_login.py`，再运行 `get_details.py`。
- **读不到 Description**：先在浏览器里手动打开那个项目看看是否正常。
- **断网 / 超时**：网络恢复后直接重新运行。

以下情况不会自动停，要自己留意：电脑睡眠或关闭窗口（下次接着抓即可）、网站只是变慢、某个项目 Description 为空但字段存在（抓完后在 Excel 里扫一眼有没有明显空白行）。

---

## 6. 第 4 步：导出 Excel

### 6.1 脚本 `export.py`

这一步只读本地的 `output/progress.jsonl`，不会访问网站。

```python
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
```

### 6.2 运行与检查

```
python export.py
```

最后一行应显示 `共 391 个`（或你实际的数量）。打开 Excel 随便扫几行，看 Description 有没有明显空白。

> 如果直接把原始数据写进 Excel 时报 `IllegalCharacterError`，就是因为数据里有隐藏的控制字符，上面的 `clean()` 就是用来解决这个问题的。

---

## 7. 第 5 步：导出 Available 项目为 Markdown

### 7.1 脚本 `to_md.py`

这一步只读本地的 `output/progress.jsonl`，不会访问网站。

```python
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
```

### 7.2 运行

```
python to_md.py
```

### 7.3 输出格式

```
1. T493015 ALGORITHMS COMPARISON OF SELF-ORGANIZATION MAPS IN DATA CLASSIFICATION
This system is intends to compare the performance ...

2. T493056 INTERPRETING A NEURO-FUZZY MODEL WITH A FUZZY REASONING MECHANISM
...
```

### 7.4 可调整

- 编号是流水号 1、2、3…，不是 Project ID。
- 想让每个描述变成一整段（去掉换行），把 `desc = clean(...)` 那行改成：
  ```python
  desc = " ".join(clean(r.get("Description", "")).split())
  ```
- 如果 Available 数量和列表页看到的对不上，检查详情页里 Status 字段的实际写法。

---

## 8. 按导师筛选的 Markdown（`to_md_by_supervisor.py`）

第 7 节的 `to_md.py` 输出全部 Available 项目。如果只想看某几位导师的，用这个版本。它同样只读本地的 `progress.jsonl`，不会访问网站。

### 8.1 脚本

```python
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
```

### 8.2 怎么用

1. **先直接运行一次**：`python to_md_by_supervisor.py`，屏幕会列出 Available 项目里每位导师各有几个，方便挑名字。
2. **在设置区填想要的导师**：`SUPERVISORS = ["Tan Shing Chiang", "Goh Pey Yun"]`，一个或多个都可以。
3. **再运行一次**，生成 `output/available_projects.md`。

### 8.3 修改设置后会变成什么

| 设置 | 修改后的效果 |
|---|---|
| `SUPERVISORS = ["A", "B"]` | 只输出这两位导师的 Available 项目 |
| `SUPERVISORS = []` | 不筛选导师，输出全部 Available 项目 |
| `OUTPUT = "output/别的名字.md"` | 输出到别的文件，不覆盖旧的 `available_projects.md` |

### 8.4 需要知道的

- 导师名字用**包含**匹配、不区分大小写，所以不用写 `(Dr)`。名字以第 1 步打印出来的为准（来自详情页的 `Supervisor name:`，可能和列表页写法略有不同）。
- 某个名字匹配不到任何 Available 项目时，屏幕会提示"找不到匹配…"，通常是拼错了，或该导师目前没有 Available 项目。
- **状态用精确匹配**：只有状态正好是 `Available` 才算。以后如果出现 `Unavailable` 之类的状态，不会被误当成 Available。
- 状态是 `Removed`（网站已删除，见第 10 节）的项目不会被输出。

---

## 9. 搜索页面（`fyp_search.html`）

一个单独的 HTML 文件，用来读取 `progress.jsonl`，方便搜索和筛选项目。**它不访问任何网站**，数据只在你打开文件时读进页面，不上传，也不保存到浏览器（所以不会留下残留）。

### 9.1 使用步骤

1. 把 `fyp_search.html` 放在电脑上，**双击用浏览器打开**（Edge、Chrome 都可以）。
2. 把 `output/progress.jsonl` **拖进页面**，或点右上角「选择 progress.jsonl」。
3. 搜索、筛选、收藏。

### 9.2 功能说明

| 功能 | 说明 |
|---|---|
| 搜索框 | 多个关键词用空格分隔，表示**同时包含**；命中的词会高亮；按键盘 `/` 快速定位 |
| 搜索范围 | 全部内容 / 仅标题 / 标题 + 描述 |
| 状态 | 按钮上显示每种状态的数量（Available、Taken、Removed，以及无法识别的状态）；默认只看 Available |
| 导师（可多选） | 每位导师后面显示项目数，上面有小搜索框 |
| 项目类型 / 适合专业 | 下拉选择 |
| 研究领域 / 所需技能 | 输入文字筛选，研究领域有自动提示 |
| 只看已收藏 | 配合 ★ 使用 |
| 排序 | ID、ID 倒序、标题 A-Z、导师 A-Z |
| 展开卡片 | 点卡片展开全部字段（目标、技能、交付物、时间表、导师邮箱等），可一键复制单个项目 |
| 收藏 ★ 与导出 | 「复制收藏 (md)」「下载收藏 (md)」，格式与第 7 节相同 |
| 深色模式 | 跟随系统自动切换 |

### 9.3 网站上已删除的项目怎么显示

页面**只能显示 `progress.jsonl` 里有的内容**，它自己不会联网，所以不会知道网站上发生了什么。要让它显示"已删除"，需要先按第 10 节更新数据，并把 `GONE_ACTION` 设成 `"mark"`，被删除的项目在数据里状态会变成 `[ Removed ]`，然后：

- 状态按钮里会出现 **Removed 及数量**，默认的 Available 视图**不会显示**它们。
- 点「Removed」或「全部」才会看到，卡片变淡、标题有删除线、状态标签显示 `Removed`。
- 如果你**收藏的项目**后来变成 Removed 或 Taken，页面顶部提示条会告诉你"收藏里有 N 个已经不是 Available"，导出收藏时也会提醒"其中有已不可选的项目"。

### 9.3.1 每个项目是否「还在」

每张卡片的状态旁边都有一行小字：

- **✔ 最近确认仍在：日期**：那天运行 `update_status.py` 时，这个项目还在最新列表里。
- **✖ 网站已移除（日期 发现）**：那天发现它已不在列表里（需要 `GONE_ACTION = "mark"`）。
- **未确认（更新数据后会显示）**：这是旧数据，还没有用 `update_status.py` 确认过。

注意：这个日期是**你上次更新时**确认的，不是现在这一刻；决定申请前仍要回网站核对。

### 9.4 页面顶部的提示条

读取文件后，结果上方会出现一个提示条，可能包含：

- **数据文件最后修改时间**，超过 3 天会提示"建议先更新"。这个时间来自文件本身，不是联网获取的。
- **永远显示的提醒**：数据只是快照，决定申请前请回 IPMS 核对状态。
- 有 Removed 项目、收藏里有已不可选项目、状态无法识别、描述为空、无法读取的行、重复的 Project ID（保留最后出现的那条），都会在这里提示。

### 9.5 局限

- 刷新页面后要重新选文件，**收藏会清空**，需要的话先导出。
- 页面不能更新数据，更新要走第 10 节的流程。
- 搜索和筛选的结果只是本地数据，**不等于网站此刻的状态**。

---

## 10. 更新数据（`update_status.py`）

### 10.1 要先弄清楚：什么会变

| 内容 | 会变吗 | 怎么处理 |
|---|---|---|
| Status（Available / Taken） | **会**，选题期间天天有人选走 | 抓一次列表就能更新，不用进详情页 |
| 新增项目 | 偶尔有 | 只对新增的项目抓详情页 |
| 网站上被删除的项目 | 可能 | 用 `GONE_ACTION` 处理 |
| 描述、目标等内容 | 基本不变 | 需要时用 `REFRESH_IDS` 单独重抓 |

所以**不需要整批重抓 391 个详情页**。列表页本身就有状态，只抓列表（4 页）加上新增项目的详情页，访问量比第一次小得多。

> 直接重跑 `get_details.py` 不会更新已有项目，因为它会跳过 `progress.jsonl` 里已有的项目。这是需要 `update_status.py` 的原因。

### 10.2 更新流程

**步骤 0：（如果环境已清理）重新准备**

```
cd C:\Users\User\Desktop\t
python -m venv myenv
myenv\Scripts\activate
pip install playwright pandas openpyxl
python save_login.py
```

**步骤 1：更新列表（只访问 4 个列表页）**

```
python get_list.py
```

**专业和 Show 必须和上次一致**，否则列表不完整，后面会误判有项目被删除。

**步骤 2：本地同步状态（不访问网站）**

```
python update_status.py
```

它会先自动备份，再把 `progress.jsonl` 里的状态更新成最新的，并打印：状态变化、新增项目、网站已删除的项目。

**步骤 3：仅当有新增项目或有 `REFRESH_IDS` 时**

```
python get_details.py
```

只会抓新增和指定重抓的项目。

**步骤 4：重新生成你要的东西**

```
python to_md_by_supervisor.py
```

或者把更新后的 `progress.jsonl` 重新拖进 `fyp_search.html`。

**步骤 5：收尾**：删除 `output\login_state.json`，在 IPMS 点 Logout（见第 13 节）。

### 10.3 脚本

```python
import json, os, shutil, time
import pandas as pd

PROGRESS = "output/progress.jsonl"
LIST_FILE = "output/project_list_all.xlsx"   # 刚用 get_list.py 重新抓的最新列表

# 想重新抓详情页的项目（比如导师改了描述），把 Project ID 写在这里；不需要就留空
REFRESH_IDS = []
# 例如: REFRESH_IDS = ["T493015", "T493056"]

# 网站上已经删除的项目（本地有、最新列表里没有）怎么处理：
#   "report" = 只提醒，不改任何数据（默认，最安全）
#   "mark"   = 把它的 Status 改成 [ Removed ]，数据保留，但不会再出现在 Available 里
#   "delete" = 直接从 progress.jsonl 删掉（有备份，但仍建议优先用 mark）
GONE_ACTION = "report"

# 1. 读取最新列表：Project ID -> Status
df = pd.read_excel(LIST_FILE)
live = dict(zip(df["Project ID"].astype(str).str.strip(),
                df["Status"].astype(str).str.strip()))

# 2. 读取现有数据
records = []
with open(PROGRESS, encoding="utf-8") as f:
    for n, line in enumerate(f, 1):
    if not line.strip():
        continue
    try:
        records.append(json.loads(line))
    except json.JSONDecodeError:
        raise SystemExit(f"progress.jsonl 第 {n} 行无法读取（可能是中断时写了一半）。"
                         "请用记事本打开删掉这一行再运行，本脚本没有改动任何文件。")

# 3. 先备份，再修改
backup = f"output/progress_backup_{time.strftime('%Y%m%d_%H%M%S')}.jsonl"
shutil.copy(PROGRESS, backup)
print("已备份到", backup)

refresh = set(REFRESH_IDS)
today = time.strftime("%Y-%m-%d")

# 安全检查：最新列表比本地少很多，多半是抓列表时选错了专业或 Show，
# 这种情况下不执行 mark / delete，避免把正常的项目误标成已删除
action = GONE_ACTION
if action in ("mark", "delete") and len(live) < 0.8 * len(records):
    print(f"警告：最新列表只有 {len(live)} 个，本地有 {len(records)} 个，"
          "差距太大，可能是专业或 Show 选错了。本次只提醒，不处理已删除的项目。\n")
    action = "report"

kept, changed, removed = [], [], []
for r in records:
    pid = r["Project ID"]
    if pid in refresh:
        continue                      # 删掉旧记录，下次 get_details.py 会重新抓
    new = live.get(pid)
    if new:
        r["Last seen"] = today            # 记录：这天确认过它还在最新列表里
        r.pop("Removed on", None)         # 如果之前被标记过移除，现在又出现了，就清掉标记
        old = r.get("Status", "")
        old_clean = old.replace("[", "").replace("]", "").strip()
        if new.lower() != old_clean.lower():
            changed.append((pid, old_clean, new))
            r["Status"] = f"[ {new} ]"
    elif action == "delete":
        removed.append(pid)
        continue
    elif action == "mark":
        if "removed" not in r.get("Status", "").lower():
            removed.append(pid)
            r["Removed on"] = today       # 记录：这天发现它已不在列表里
        r["Status"] = "[ Removed ]"
    kept.append(r)

# 4. 写回（先写临时文件再替换，中途出错也不会弄坏原文件）
with open(PROGRESS + ".tmp", "w", encoding="utf-8") as f:
    for r in kept:
        f.write(json.dumps(r, ensure_ascii=False) + "\n")
os.replace(PROGRESS + ".tmp", PROGRESS)

# 5. 报告
have = {r["Project ID"] for r in kept}
new_ids = [i for i in live if i not in have]
gone = [i for i in have if i not in live and i not in removed
        and not any(r["Project ID"] == i and "removed" in r.get("Status", "").lower() for r in kept)]

print(f"\n状态变化: {len(changed)} 个")
for pid, old, new in changed:
    print(f"  {pid}: {old} -> {new}")
if removed:
    print(f"\n网站已删除的项目，本次已{'标记为 Removed' if action == 'mark' else '从本地删除'}: {len(removed)} 个")
    for i in removed:
        print("  ", i)
print(f"\n新增项目（还没有详情）: {len(new_ids)} 个")
for i in new_ids:
    print("  ", i)
print(f"\n本地有、但最新列表里没有: {len(gone)} 个（不会删除，只是提醒）")
if gone:
    print("  可能是项目被移除，或抓列表时选的专业 / Show 数量和上次不同")
    for i in gone[:20]:
        print("  ", i)

if new_ids or refresh:
    print("\n下一步：运行 python get_details.py，只会抓新增和你指定要重抓的项目")
else:
    print("\n没有需要抓详情页的项目，更新完成")
```

### 10.4 三个设置，改了会怎样

**`REFRESH_IDS`（最常用）**：指定哪些项目要重新抓详情页。

| 修改 | 效果 |
|---|---|
| `REFRESH_IDS = []`（默认） | 只更新状态，其他内容不动 |
| `REFRESH_IDS = ["T493015"]` | T493015 的旧记录被移除，再运行 `get_details.py` 就会重新抓它 |

ID 要加引号，多个用英文逗号隔开。**用完记得改回 `[]`**，否则下次运行又会移除并重抓。因为每打开一次详情页，对应项目的浏览次数可能 +1，所以不要滥用。

**`GONE_ACTION`**：网站上已删除的项目（本地有、最新列表里没有）怎么处理。

| 修改 | 效果 |
|---|---|
| `"report"`（默认） | 只提醒，不改任何数据 |
| `"mark"` | 状态改成 `[ Removed ]`，数据保留，不再出现在 Available 里，搜索页面里会显示为已移除 |
| `"delete"` | 直接从 `progress.jsonl` 删除（有备份，但建议优先用 `mark`） |

**`LIST_FILE` / `PROGRESS`**：最新列表文件和数据文件的名字，只有你重命名过文件才需要改。

### 10.5 确认项目真的被删除了再标记

脚本提示"本地有、但最新列表里没有"时，不要马上处理：

1. 先检查抓列表时专业和 Show 是否和上次一致。
2. 手动在浏览器里打开那个项目的详情页（网址里把 `q=` 后面换成项目 ID），确认是否真的打不开或不存在。只手动看几个，不要用脚本逐个访问。
3. 确认后把 `GONE_ACTION` 改成 `"mark"`，再运行一次 `update_status.py`。
4. 用完改回 `"report"`。

### 10.6 脚本内置的安全措施

- **每次运行前自动备份**成 `output/progress_backup_日期时间.jsonl`。
- **写入先写临时文件再替换**，中途出错不会弄坏原文件。
- **数量差距保护**：如果最新列表比本地少 20% 以上，大概率是专业或 Show 选错了，这时 `mark` 和 `delete` 不会执行，只提醒。
- **状态用精确比较**，不会把 `Unavailable` 当成 `Available`。

---

## 11. 保护网站的原则（最重要）

这些脚本的目的**只是把项目资料抄下来方便自己看题目**，不能为此影响 IPMS 的正常运行，也不能影响其他同学和老师使用。下面这些是底线：

### 11.1 只读，不改任何东西

- 脚本只做"打开页面、读取文字"，**不点任何按钮，不填任何表单**，不会申请项目、提交、保存、删除。
- 只访问你登录后正常能看到的页面：列表页和项目详情页。
- 不要修改网址参数去猜别的页面，不要去碰别人的资料（比如别的同学的 My Project）。

### 11.2 请求要少、要慢、要一个一个来

- **顺序访问，不并发**：一次只开一个页面，不要同时开多个窗口或多个脚本。
- **间隔 3～6 秒随机**（`time.sleep(random.uniform(3, 6))`）。这一行**不要删，也不要把间隔调得比 3 秒更小**；想更温和可以调大，例如 `(3, 6)`。
- **分批**：每次最多 100 个，批与批之间隔几个小时。
- **更新只抓列表 + 新增项目**，不要整批重抓详情页。
- **不要频繁运行**：选题期间每隔几天更新一次就够，不需要每天。
- **避开高峰**：选题截止前几天、大家都在用系统时不要跑，早上或深夜更好。

### 11.3 出现异常就停，不要硬试

- 脚本遇到异常状态码（403、429、500…）、超时、跳回登录页、读不到 Description，都会**自己停下**。
- 停下之后**不要立刻重试**：网站变慢或报错，说明它可能正在承受压力，应该等几个小时甚至隔天再继续。
- 不要把脚本改成"出错自动重试"或"无限循环"，也不要让它无人看管地跑。
- 如果网站管理员或老师要求你停止，**立刻停止**。

### 11.4 已知的小影响（要诚实知道）

- 每打开一个详情页，该项目的 **View 浏览次数可能 +1**，这会让老师看到的浏览数略有偏高，所以不要反复重抓详情页。
- 学校服务器会留下访问日志，你这边删不掉。管理员如果去看，能看出账号在短时间内规律地打开了很多页面。所以风险是"低"，不是"零"。
- 最稳妥的做法仍然是**事先问一下 FYP 协调员**，一句话就够，例如：*"I would like to copy the project list into a spreadsheet for my own comparison when choosing a topic. Is that okay?"*

### 11.5 数据与账号

- 只用自己的账号，手动登录，**不把密码写进代码**。
- `login_state.json` 等同于登录凭证，用完就删，不分享、不上传。
- 抓下来的资料只留在自己电脑上，**不上传、不转发、不分享**。

---

## 12. 没想到的情况与处理

| 情况 | 会发生什么 | 怎么处理 |
|---|---|---|
| 项目在网站上被删除 | 本地仍是旧状态，可能被当成可选项目 | 按第 10 节，用 `GONE_ACTION = "mark"` |
| 项目被别人选走 | 状态 Available → Taken | 抓列表后运行 `update_status.py` 自动更新 |
| 项目被释放（Taken → Available） | 同上 | 同上，会自动更新 |
| 新增了项目 | 本地没有 | `update_status.py` 会列出新增 ID，再运行 `get_details.py` 抓取 |
| 导师修改了描述 | 脚本发现不了（只比较状态） | 手动填 `REFRESH_IDS` 重抓 |
| 新学期 / 新一轮 FYP，列表整批更换 | 大量"本地有、列表没有"，触发 20% 保护 | 把旧 `progress.jsonl` 改名存档，当作第一次重新分批抓取，不要在旧数据上硬更新 |
| 抓列表时专业或 Show 选错 | 列表不完整，被误判为"已删除" | 数量保护会拦住；重新抓列表，保持专业和 Show 一致 |
| 出现新的状态（如 Unavailable） | 搜索页面会单独显示这个状态，不会当成 Available | 到页面「状态」按钮里查看，必要时告诉我来调整 |
| 网站改版 | `get_details.py` 读不到 Description 而自动停下 | 重新查看页面结构，再改脚本，不是改设置能解决的 |
| 登录过期（批次中途） | 提示"跳回登录页"并停止 | 重新运行 `save_login.py`，再接着抓 |
| 抓取中途中断（Ctrl+C、断电） | `progress.jsonl` 最后一行可能写了一半 | 搜索页面会跳过并提示"无法读取"；`get_details.py` 会自动跳过并提示坏行，下次抓取不会粘连；`update_status.py` 会停下并告诉你是第几行，用记事本删掉那一行再运行 |
| Excel 正开着时写 xlsx | `PermissionError` | 关掉 Excel 再运行 |
| 同时开两个窗口跑 `get_details.py` | 同时写同一个文件，可能错乱，也违反第 11 节 | 一次只跑一个 |
| 同一项目出现两次 | 搜索页面自动合并，保留最后一条 | 无需处理 |
| 备份文件越积越多 | 里面是项目资料和导师邮箱 | 确认更新无误后删除 `progress_backup_*.jsonl` |
| 一个项目属于多个专业（如 `DCN,ST,AI`） | 换专业抓列表时，列表内容会不同 | 每次抓列表都选同一个专业 |
| 数据太旧 | 搜索页面看到的状态已经过期 | 搜索页面会提示文件已过几天；**申请前一定回网站核对** |

---

## 13. 彻底清理（不留残留）

**保留**：`.py` 脚本，以及导出的资料（`available_projects.md`、`projects_with_description.xlsx` 等）。
**删除**：登录凭证、依赖、缓存、运行残留。

### 13.0 清理前：备份要保留的东西（建议）

把脚本和成品（`.md`、`.xlsx`）复制到另一个文件夹，例如 `Desktop\ipms_keep`，避免手滑误删。

### 13.1 删除登录凭证（最重要，最先做）

```
cd C:\Users\User\Desktop\t
del output\login_state.json
```

再回到 IPMS 网站点一次 **Logout**，让服务器那边的登录会话也失效。

### 13.2 删除虚拟环境（所有 pip 依赖）

```
deactivate
rmdir /s /q myenv
```

playwright、pandas、openpyxl 以及它们的所有依赖都装在 `myenv` 里，删掉这个文件夹就全部清掉，没有装到系统 Python。

### 13.3 清理 pip 下载缓存

退出虚拟环境后 `pip` 命令可能不可用，所以用：

```
python -m pip cache purge
```

### 13.4 删除运行残留文件

```
del output\login_state.json
del output\progress.jsonl
del output\progress_backup_*.jsonl
del output\project_list_all.xlsx
del output\*.tmp
rmdir /s /q __pycache__
```

- `progress.jsonl` 是原始抓取数据，`.md` 和 Excel 都是从它生成的。**确认两个成品都满意后再删**；想保险就先留着。
- `progress_backup_*.jsonl` 是 `update_status.py` 每次运行自动生成的备份，里面同样是项目资料和导师邮箱，确认更新无误后全部删除。
- `project_list_all.xlsx` 是 391 个项目的列表，想留就不删。
- 如果还想继续用 `fyp_search.html` 或以后再更新，就先保留 `progress.jsonl`；它是搜索页面和更新流程的数据来源。
- 如果过程中保存过网页源码（如 `list_page.html`，里面带有姓名和专业）或测试文件（如 `test_details.json`），也一并删除。
- `__pycache__` 如果不存在，会提示找不到，忽略即可。

### 13.5 检查是否干净

在工作文件夹里输入：

```
dir
```

应该只剩：

- 脚本：`save_login.py`、`get_list.py`、`get_details.py`、`export.py`、`to_md.py`、`to_md_by_supervisor.py`、`update_status.py`
- 页面：`fyp_search.html`
- 资料：`available_projects.md`、`projects_with_description.xlsx`（以及你选择保留的文件）

不应该再有：`login_state.json`、`myenv`、`progress.jsonl`（如果已决定删）、`progress_backup_*.jsonl`、`__pycache__`、网页源码和测试文件。

### 13.6 可选：清 Edge 浏览记录

脚本是用电脑自带的 Edge 打开的。如果介意，可在 Edge 的 设置 → 隐私 → 清除浏览数据 里清理。

### 13.7 不需要删的

- **Edge**：系统自带，脚本只是调用它，没有另外安装任何浏览器。
- **Python 本身**：除非以后完全不用。要删的话在 Windows 设置 → 应用 里卸载。

### 13.8 无法清除的

- Windows 自身的痕迹（命令行历史、最近使用的文件、系统日志）。
- 学校服务器上的访问记录和项目 View 浏览次数。

---

## 14. 常见问题速查

| 现象 | 原因 / 处理 |
|---|---|
| `'pip' is not recognized` | 没有激活虚拟环境，先执行 `myenv\Scripts\activate`；清缓存时用 `python -m pip cache purge` |
| 运行脚本提示找不到 `login_state.json` | 还没运行 `save_login.py`，或 output 文件夹里没有这个文件 |
| 打开后是登录页 / 提示跳回登录页 | 登录状态过期，重新运行 `save_login.py` |
| 列表抓到的数量不对 | 确认已选对专业，且 Show 选到 100；检查最后的 "of N entries" |
| 详情页中文或符号出现 `â€“` 乱码 | 网页本身的编码问题，`get_details.py` 里的 `fix()` 会自动修正 |
| 写 Excel 报 `IllegalCharacterError` | 数据里有隐藏控制字符，用 `export.py`（带 `clean()`） |
| 搜索页面看不到某个项目 | 默认只显示 Available；到左侧「状态」选「全部」，或检查是否被标记为 Removed |
| `update_status.py` 提示列表比本地少很多 | 抓列表时专业或 Show 选错，重新抓列表，不要强行处理 |
| 第二次运行好像从头开始 | `progress.jsonl` 被删、改名或不在 output 文件夹里 |
| 出现 403 / 429 | 被限制了，等几个小时或隔天再继续，不要硬试 |

---

## 15. 完整使用顺序（速查）

**第一次使用**

```
cd C:\Users\User\Desktop\t
python -m venv myenv                      （仅第一次）
myenv\Scripts\activate
pip install playwright pandas openpyxl    （仅第一次）

python save_login.py                      手动登录一次
python get_list.py                        选专业、Show 100，抓列表
python get_details.py                     输入 3 测试；之后按回车自动分批，隔几小时重复
python export.py                          导出 Excel
python to_md_by_supervisor.py             导出 Available（可按导师筛选）的 Markdown
```

然后用浏览器打开 `fyp_search.html`，把 `progress.jsonl` 拖进去搜索、收藏。

**之后更新（选题期间每隔几天一次，不要每天）**

```
myenv\Scripts\activate
python save_login.py                      登录状态过期时才需要
python get_list.py                        专业和 Show 与上次一致
python update_status.py                   同步状态，查看新增和被删除的项目
python get_details.py                     仅当有新增项目时
python to_md_by_supervisor.py             或重新把 progress.jsonl 拖进 fyp_search.html
```

**决定申请某个项目前：一定回 IPMS 网站核对它是否仍是 Available。**

**全部用完后：按第 13 节清理。**
