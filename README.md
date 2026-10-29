# 🎓 MMU IPMS · FYP Title Management

用于在本地查看、整理和比较 MMU IPMS 的 FYP 题目。项目由 Python 数据整理脚本与一个独立 HTML 搜索页面组成。大家可以用自己的账号获取正常可见的题目资料，再在本地搜索、筛选、收藏和导出。

- 🔎 **查找题目**：按关键词、导师、状态、专业、类型、领域与技能组合筛选。
- ⭐ **保留选择**：用 JSON 保存各清单的筛选和收藏，区分同一清单更新与全新清单。
- 📄 **整理资料**：查看项目详情，点击邮箱复制地址，导出 Excel 或 Markdown。

## 🚀 快速启动

### 已有项目资料：直接打开网页

此方式不需要运行 Python、安装 Node.js 或启动服务器，也不会访问 IPMS。

1. 用浏览器打开 [fyp_search.html](fyp_search.html)。
2. 点击「导入数据」，或拖入已有的 `progress.jsonl`。
3. 核对项目数量、状态和文件校验提示，再搜索、筛选或点击星标收藏。
4. 需要保留筛选与收藏时，在「使用记录」中新建或打开记录 JSON。建议保存为 `output/fyp_search_state.json`。

**关闭前确认保存状态。** 不支持直接写回文件时，页面会提供下载模式，请自行保留最新 JSON 副本。

### 首次获取资料：准备 Python 环境

> **访问网站前，请先获得 FYP 负责人的许可。未获得明确答复前，不要运行访问网站的脚本。** 已有授权资料的本地查看、状态整理和导出不需要重新访问网站。

打开 **Windows CMD** 运行以下指令。请把示例路径替换为自己的项目路径，并始终从项目根目录运行脚本。

运行前先确认 Python 已安装，且 `python --version` 能正常输出版本。指南以 Python 3.14 为环境准备示例，已有可运行环境时，优先保留其版本。

```cmd
cd /d C:\path\to\MMU-IPMS-Title-Management
python -m venv myenv
myenv\Scripts\activate
python -m pip install playwright pandas openpyxl
```

脚本默认调用电脑已安装的 **Microsoft Edge**，无需另外运行 `playwright install`。若改用 **Chrome**，需要同步修改三个浏览器设置，见[使用指南第 3 节](FYP_project_scraping_guide.md#3-第-1-步保存登录状态)。

准备好后，按下方顺序登录、获取列表，再分批获取详情。完整环境准备、校验和清理步骤见 [FYP 项目使用指南](FYP_project_scraping_guide.md)。

## 🧭 如何运行

### 1. 获取项目资料

| 顺序 | 命令 | 需要完成的操作 |
|---|---|---|
| a. 保存登录状态 | `python save_login.py` | 在弹出的浏览器中手动登录，看到 Project List 后，回到命令行按回车 |
| b. 获取项目列表 | `python get_list.py` | 手动选择专业及范围筛选，建议将 Show 设为 100，按回车后脚本遍历分页 |
| c. 分批获取详情 | `python get_details.py` | 首次在数量提示中输入 `3`，核对字段后再分批继续 |

登录状态、列表和详情分别写入 `output/login_state.json`、`output/project_list_all.xlsx`、`output/progress.jsonl`。登录脚本会创建 `output/`，项目资料不随代码提供。

- 当前详情脚本默认 `ONLY_AVAILABLE = False`，会处理列表中尚未保存的各状态项目。
- 每次最多 100 个，项目之间随机间隔 3～6 秒，批次之间留出几个小时，不并发运行。
- 再次运行会跳过已有 Project ID。已有旧进度和列表时，先将它们放入 `output/`，避免从头重新获取。
- 列表总数需与网页分页信息核对，同时确认专业、范围与抽查记录正确，总数一致不能单独证明资料完整。

### 2. 搜索、查看与保存收藏

打开 `fyp_search.html` 并导入 `output/progress.jsonl` 后：

| 操作 | 用途 |
|---|---|
| 搜索与筛选 | 组合多个条件，关键词用空格分隔表示同时包含，按 `/` 定位搜索框 |
| 展开卡片 | 查看目标、技能、交付内容、时间安排等详情，字段标题配有图标，点击邮箱地址可复制 |
| 全部收藏 | 查看所有状态及未匹配收藏，不受主列表筛选影响，支持复制或下载 Markdown |
| 使用记录 | 管理记录 JSON、独立清单、文件权限与浏览器草稿 |
| 更新当前清单 | 为原清单载入新版资料，保留其筛选和收藏，缺失或身份变化的收藏需要另行核对 |
| 导入新清单 | 为新学期或不同范围建立独立档案，保留旧档案，不自动复制收藏 |

`progress.jsonl` 是项目资料，`fyp_search_state.json` 是个人使用记录，`login_state.json` 是登录凭证。**三者用途不同，不能互相替代。**

首次保存使用记录需要选择位置并授权，双击 HTML 不会直接在磁盘上创建 JSON。浏览器缓存只作辅助，换浏览器、移动 HTML 或清理缓存后，可以重新打开记录 JSON。保存模式、清单识别与校验规则见[指南第 9 节](FYP_project_scraping_guide.md#9-搜索页面与使用记录fyp_searchhtml)。

### 3. 导出 Excel 或 Markdown

这些操作只处理本地 `progress.jsonl`，不访问 IPMS。

| 命令 | 输出 | 内容 |
|---|---|---|
| `python export.py` | `output/projects_with_description.xlsx` | 导出脚本选定的项目字段，包含各状态 |
| `python to_md.py` | `output/available_projects.md` | 仅导出 Available 项目的编号、标题与描述 |
| `python to_md_by_supervisor.py` | `output/available_projects.md` | 在 Available 项目中，按脚本内的 `SUPERVISORS` 设置筛选导师 |

**两个 Markdown 脚本默认写入同一个文件，后运行的会覆盖前一次结果。** 同时保留两个版本时，先修改导师筛选脚本的 `OUTPUT`。写入 Excel 前，请关闭正在打开的同名文件。

网页中的收藏 Markdown 是另一份导出，包含当前清单的全部收藏及必要的快照说明，浏览器通常将其下载为 `shortlist.md`，不一定放在 `output/`。

### 4. 更新已有资料

1. 登录状态失效或已删除时，先运行 `python save_login.py`。
2. 运行 `python get_list.py`，保持与原清单相同的专业和范围，核对分页与总数。
3. 运行 `python update_status.py`，在本地备份并同步状态。默认 `GONE_ACTION = "report"` 只报告缺失项，不自动标记或删除它们。
4. 仍有缺少详情的项目时，再按批次运行 `python get_details.py`。
5. 按需重新导出文件，或在网页中选择「更新当前清单」，最后确认个人记录已保存。

`update_status.py` 不访问网站，也不会自动获取新增项目的详情。使用 `REFRESH_IDS` 或批量处理缺失项目之前，请先阅读[指南第 10 节](FYP_project_scraping_guide.md#10-更新数据update_statuspy)，避免错误移除原记录。

## 🔒 隐私与安全

- **账号与权限**：只使用自己的账号，手动登录，不绕过权限、不猜测其他数据入口、不操作他人的资料。
- **登录凭证**：`login_state.json` 包含可复用的登录状态，应按凭证保管，不上传、不分享。删除文件不等于注销服务器会话，收尾方法见[指南第 13 节](FYP_project_scraping_guide.md#13-清理本项目的本地文件)。
- **项目与个人记录**：项目资料、导出文件、使用记录和备份仅供本人本地整理，不上传或转发。记录 JSON 与浏览器草稿也可能含标题、描述、导师等片段。
- **Git 与云盘**：现有 `.gitignore` 排除了 `output/` 中的文件，但不会替你配置云盘，也不会清除已经提交过的文件。另存到其他位置的 JSON、导出和下载副本也要单独排除。
- **网站负载**：顺序、低频、分批访问，遇到 403、429、登录异常、网站变慢或资料异常时停止，不反复重试。脚本读取页面仍可能产生服务器日志及项目浏览次数。
- **状态核对**：网页显示的是本地快照。决定申请题目前，请回到 IPMS 核对其当前状态。

## 📁 项目结构

```text
MMU-IPMS-Title-Management/
├── README.md                           # 项目介绍与快速入口
├── FYP_project_scraping_guide.md       # 完整流程、规则、校验与清理说明
├── fyp_search.html                     # 本地搜索、筛选、详情与收藏页面
├── save_login.py                       # 手动登录并保存登录状态
├── get_list.py                         # 遍历分页，获取当前范围的项目列表
├── get_details.py                      # 分批获取详情，保留断点进度
├── update_status.py                    # 备份进度并同步本地状态
├── export.py                           # 导出 Excel
├── to_md.py                            # 导出 Available 项目 Markdown
├── to_md_by_supervisor.py              # 按导师筛选后导出 Markdown
├── .gitignore                          # 本地数据与环境排除规则
├── LICENSE                             # MIT 源码许可
└── output/                             # 运行后生成或由本人放入的数据
    ├── login_state.json                # 登录凭证
    ├── project_list_all.xlsx           # 项目列表
    ├── progress.jsonl                  # 项目详情与抓取进度
    ├── progress_backup_*.jsonl         # 更新前备份
    ├── projects_with_description.xlsx
    ├── available_projects.md
    └── fyp_search_state.json           # 网页使用记录，需自行选择保存位置
```

上面列出的数据文件按相应步骤生成，不代表仓库附带这些资料。网页记录与下载文件的实际位置，以你选择的位置或浏览器下载设置为准。

## 📝 备注

- **运行位置**：脚本使用相对路径，请先进入项目根目录。重新打开 CMD 后，需要再次激活 `myenv`。
- **环境版本**：本项目没有锁定依赖版本文件。优先保留已跑通的环境，环境记录与重建方式见指南第 2.4 节。
- **网页变化**：IPMS 页面结构、字段或登录流程改变后，脚本可能需要调整。文档中的数量示例不代表当前实际项目数。
- **可选设置**：`ONLY_AVAILABLE`、批次上限、浏览器、导师筛选及更新规则都在对应 Python 文件中配置，修改前先看指南，不要缩短访问间隔或增加并发。
- **已知边界**：页面不会实时监听磁盘文件，也不会自动核对网站。记录 JSON 不包含完整项目库，Markdown 不能替代原始进度备份。
- **源码许可**：见 [LICENSE](LICENSE)。开源代码不等于已获得 IPMS 的访问、资料使用或再分发许可。

## 🙏 致谢

本项目由开发者本人与 AI 工具共同完成。功能设计和整体方向由开发者规划主导，AI 在代码编写和审查过程中提供了辅助支持。

## 📢 声明

本项目为个人整理工具，非学校或 IPMS 的官方项目。

> 本项目仅用于更方便地查看、整理和比较 FYP 题目，不以恶意获取资料、访问未经授权的数据或影响 IPMS 正常运行为目的。若学校、系统管理方、资料相关人员或其他有关人员对本项目或其内容有异议，请联系我，收到通知后，我会立即停止相关使用，并删除本项目及相关资料。