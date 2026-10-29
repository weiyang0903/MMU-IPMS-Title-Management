import json, os, shutil, time
import pandas as pd

PROGRESS = "output/progress.jsonl"
LIST_FILE = "output/project_list_all.xlsx"   # 刚用 get_list.py 重新抓的最新列表

# 想重新抓详情页的项目（比如导师改了描述），把 Project ID 写在这里；不需要就留空
REFRESH_IDS = []
# 例如: REFRESH_IDS = ["T493015", "T493056"]

# 网站上已经删除的项目（本地有、最新列表里没有）怎么处理：
#   "report" = 对缺失项目只提醒，不删除或标记它们；其他状态和日期仍会更新（默认）
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
                             "请先手动备份原文件，再检查并删除损坏行后运行；本脚本没有改动任何文件。")

# 3. 先备份，再修改
backup = f"output/progress_backup_{time.strftime('%Y%m%d_%H%M%S')}.jsonl"
shutil.copy(PROGRESS, backup)
print("已备份到", backup)

refresh = set(REFRESH_IDS)
today = time.strftime("%Y-%m-%d")

# 数量检查：只比较列表 ID 数与本地记录行数，不验证 ID 重合度或列表完整性。
# 最新列表少于本地记录行数的 80% 时，本次不执行 mark / delete；仍会同步其他状态和日期。
action = GONE_ACTION
if action in ("mark", "delete") and len(live) < 0.8 * len(records):
    print(f"警告：最新列表只有 {len(live)} 个，本地有 {len(records)} 个，"
          "差距太大，可能是专业或筛选范围不同、列表未抓完整。本次不删除或标记缺失项目，其他同步仍会进行。\n")
    action = "report"

kept, changed, removed = [], [], []
for r in records:
    pid = r["Project ID"]
    if pid in refresh:
        continue                      # 先移除旧记录；后续是否补抓取决于列表、筛选和批次
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
    print("  可能是项目被移除，或专业 / 筛选范围不同、列表未抓完整")
    for i in gone[:20]:
        print("  ", i)

if new_ids or refresh:
    print("\n下一步：运行 python get_details.py，按当前列表、筛选和批次补抓缺少详情的项目")
else:
    print("\n没有需要抓详情页的项目，更新完成")