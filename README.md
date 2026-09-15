# 施工排期甘特图工具

[中文](#中文说明) · [English](#english)

> 给工程人的、双击就能用的施工进度排期工具。
> A construction scheduling tool for civil engineers — download and double-click, that's it.

基于 Python + Tkinter 的**单机绿色工具**，打包后**单个 exe 双击即用** ——
不用装环境、不用联网、U 盘带走就能在项目部电脑上跑。

---

## 中文说明

### 为什么做这个

工程上的进度排期，常见几种选择，各有各的别扭：

| 方案 | 别扭在哪 |
|------|---------|
| Microsoft Project | 要钱、要装、要学；项目部电脑未必有 |
| 在线甘特图平台 | 要联网、要注册；工地上网络未必好；数据还要传到别人服务器 |
| 通用项目管理软件 | 下拉菜单里塞着 "FS/SS/FF + Lag"，不认得"搭接""分部工程" |
| Excel 画横道图 | 能画，但改一个工期要手动挪一堆格子 |

所以做了这个：**一个懂工程语境、又不添麻烦的现场工具**。

### 特点

**一、工程语义（而不是通用项目管理术语）**

| 本工具 | 通用软件的叫法 | 为什么用工程说法 |
|--------|--------------|----------------|
| **搭接关系 / 搭接时间** | 前置任务 / Lead-Lag | 工地上就是这么说的 |
| **分部工程** | WBS / 分组 | 《建筑工程施工质量验收统一标准》里的正式概念 |
| **停歇期** | 日历 / 非工作日 | 春节、雨季、冬歇 —— 工程特有 |
| **标准层自动生成** | 手动加 N 行 | 2F~18F 一键铺开，不用重复劳动 |

**二、零门槛**
- 单个 exe，双击就跑
- 不写注册表、不装服务；删掉文件夹 = 完全卸载
- 全程离线，数据只在自己电脑上

**三、两个视图**
- **甘特图**：按周/按天显示，关键路径高亮，Ctrl+滚轮缩放
- **搭接关系图**：任务依赖网络图，看清工序流向

### 功能一览

**任务管理**
- 主任务 / 子任务两级结构
- 分部工程：自定义分类层（如"地基与基础工程""主体结构工程"），可折叠、可拖拽收纳
- 标准层批量生成：填一层参数 → 自动生成 N 个子任务（带楼层号与递增日期）
- 搭接关系：前置任务 + 搭接时间（正数延后、负数提前）
- 停歇期管理：春节、雨季等停工期，自动顺延

**视图**
- 甘特图：按周/按天，关键路径红框标示
- 搭接关系图：自动分层布局 + 蛇形折行，重复连线自动合并
- 搜索定位：实时匹配下拉，自动展开、标黄、滚动居中
- 显示今天：红色虚线标出当前日期
- 画布缩放：Ctrl+滚轮，25% ~ 400%

**数据**
- 自动保存（每次改动即存）+ 历史备份
- 保存 / 导入：`.json` 格式
- 导出 Excel：两个工作表（主任务计划 / 全部任务计划），带甘特条

### 快速开始

**方式一：下载 exe（推荐给不装开发环境的人）**

到 [Releases](../../releases) 页面下载 `施工排期甘特图工具_单文件版.exe`，双击运行。

**方式二：从源码运行**

```bash
python gantt_tool.py
```

- 环境要求：Python 3.8+（Tkinter 是标准库，无需额外安装）
- 导出 Excel 需要：`pip install openpyxl`

**方式三：自己打包成 exe**

```bash
pip install pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed \
    --name "施工排期甘特图工具" gantt_tool.py
```

输出在 `dist/` 下，双击 exe 即可。

### 📖 使用教程

**第一次用？强烈建议先看 [使用教程](docs/教程.md)** —— 里面有：

- 每个按钮是干什么的（配图）
- **搭接关系怎么设**（这是关键，直接决定图表准不准）
- 高效录入任务的技巧
- 一个完整的建计划示例（主任务 / 子任务 / 分部工程）
- 常见问题解答

### 数据存在哪

- **源码运行**：数据存在**程序所在文件夹**（跟着工程走）
- **打包成 exe 运行**：数据存在 **`我的文档\Project Gante\`**（自动创建）

> 为什么不一样？因为非技术用户容易把 exe 旁边的 `.json` 当垃圾文件删掉。
> 如果你希望 exe 版也把数据放程序旁边，在 exe 同目录建个空文件 `portable.flag` 即可。

### 使用说明（要点）

**搭接时间**

> 开始日期 = 所选前置中"最晚结束的日期" + 搭接时间

- 填 `0`：前置干完第二天就开工
- 填 `5`：前置干完再等 5 天
- 填 `-10`：前置还没干完就提前 10 天进场（搭接施工）

**分部工程**
不是任务的父子关系，只是"归类"。一个任务随时可以换分部。

**停歇期**
如春节放假，设置后落在停歇期内的任务会自动顺延。

### 已知限制

- 目前只支持"完成-开始"（FS）一种搭接类型，用搭接时间模拟其他情况
- 关键路径是"最晚结束链"高亮，不是完整 CPM 算法（无浮动时间计算）
- 未做虚拟滚动，任务数特别大（万级）时可能变慢
- 单机工具，无多人协作
- 主要面向 Windows（用了 DPI 感知 API）

### 版本记录

| 版本 | 主要变化 |
|------|---------|
| **v1.2.0** | 界面重构为**菜单栏**（任务/数据/视图/帮助），工具条只留搜索等高频控件；列表层级缩进（主任务 1 格、子任务 2 格）；**双击折叠/展开**、**右键编辑**；「新建分部工程」提到任务菜单首位；新增快捷键（Ctrl+N/S/O/E/F） |
| v1.1.0 | 新增**分部工程**分类收纳（可拖拽）；**搜索定位**（实时匹配、自动展开、标黄、居中）；画布 **Ctrl+滚轮缩放**；前置选择器支持搜索；**显示今天**竖线；网络图连线治理；「延迟进场」更名为「搭接时间」 |
| v1.0.0 | 首个正式版本：任务管理、甘特图、搭接关系图、搭接关系与停歇期、导出 Excel |

---

### 许可

**GPL-3.0**（含作者附加条款）—— 你可以自由使用、修改、分发，但：

1. **必须署名** —— 不得删除作者信息，不得宣称是自己原创
2. **须告知作者** —— 分发或用于公开产品时，请知会一声
3. **衍生作品必须开源** —— 改了之后分发，必须同样以 GPL-3.0 公开源码
4. **保留授权文件** —— 不得移除 LICENSE

详见 [LICENSE](LICENSE)。

---

## English

> A lightweight construction project scheduling tool for Chinese civil engineers.
> Built with Python + Tkinter. Ships as a single standalone `.exe` — no install, no internet needed.

### Why this exists

Scheduling tools for construction sites tend to fall into two traps: either they are
generic project-management software that speaks "FS/SS/FF + Lag" instead of the
vocabulary engineers actually use, or they are cloud platforms that require
registration and an internet connection your site may not have.

This tool speaks the language of the site:

| This tool | Generic PM software | Why |
|-----------|--------------------|-----|
| **搭接关系 / 搭接时间** (overlap relation / overlap time) | Predecessor / Lead-Lag | Site vocabulary |
| **分部工程** (division of works) | WBS / Grouping | A formal concept in Chinese construction quality-acceptance standards |
| **停歇期** (shutdown period) | Calendar / Non-working days | Spring Festival, rainy season, winter break |
| **标准层自动生成** (standard-floor auto-generation) | Add N rows manually | Generate floors 2F–18F in one click |

### Features

- **Task management** — parent/sub-task hierarchy; custom "division of works" layers
  (collapsible, drag-to-assign); bulk generation of standard floors; overlap
  relations with lead/lag; shutdown periods (holidays, rainy season).
- **Views** — Gantt chart (weekly/daily, critical path highlighted, Ctrl+wheel zoom);
  precedence network diagram (auto-layered layout, duplicate links merged);
  search with live suggestions, auto-expand and highlight; "today" marker line.
- **Data** — autosave with rolling backups; JSON import/export;
  Excel export (two sheets, with Gantt bars).

### Quick start

**Option 1 — download the exe** (recommended for non-developers)

Grab `施工排期甘特图工具_单文件版.exe` from the [Releases](../../releases) page and double-click it.

**Option 2 — run from source**

```bash
python gantt_tool.py
```

Requires Python 3.8+ (Tkinter ships with Python). Excel export needs `pip install openpyxl`.

**Option 3 — build your own exe**

```bash
pip install pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed \
    --name "施工排期甘特图工具" gantt_tool.py
```

### 📖 Tutorial

New to this tool? See the **[tutorial (Chinese)](docs/教程.md)** with annotated screenshots —
it covers every button, how to set up overlap relations, tips for fast data entry,
and a complete worked example.

### Where data is stored

- **Run from source**: data lives next to the script.
- **Run as packaged exe**: data lives in `Documents\Project Gante\` (created automatically).

> The split exists because non-technical users tend to delete stray `.json` files
> next to an exe. To make the packaged version keep data beside the exe instead,
> create an empty file named `portable.flag` in the same folder.

### Known limitations

- Only Finish-to-Start (FS) relations; other types are simulated with lead/lag.
- "Critical path" is a longest-chain highlight, not a full CPM with float calculation.
- No virtual scrolling — very large projects (10k+ tasks) may feel slow.
- Single-user desktop tool; no collaboration features.
- Primarily Windows (uses DPI-awareness APIs).

### License

**GPL-3.0** with additional terms. You may use, modify and redistribute it freely, provided that:

1. **Attribution is kept** — do not remove the author's copyright or claim the work as your own.
2. **The author is notified** — please let the author know when you distribute it or ship it in a public product.
3. **Derivatives stay open** — modified versions must be distributed under GPL-3.0 with full source.
4. **This license file is kept** — do not remove LICENSE.

See [LICENSE](LICENSE).

---

## 贡献 / Contributing

If you work in construction and find something awkward, or want a feature,
feel free to open an Issue. This tool grew out of real scheduling pain —
**knowing your pain point is worth more than anything**.

## 开发方式 / How this was built

本项目的需求、工程逻辑与验收标准由开发者确定；代码实现与调试由开发者与
AI 编程助手协作完成。

这也解释了项目里为什么保留了一套「版本检查点 + 一键重新打包」流程 ——
在这种快速迭代的开发方式下，随时能回退到上一个可用版本很重要。

> The requirements, engineering logic and acceptance criteria came from the developer.
> Implementation and debugging were done in collaboration with an AI coding assistant.
> That workflow is also why this project keeps a lightweight version-checkpoint and
> repackaging setup — being able to roll back to the last working build matters when
> you iterate quickly.

## 致谢 / Credits

- Python 标准库 `tkinter` —— 图形界面
- [`openpyxl`](https://openpyxl.readthedocs.io/) —— Excel 导出
- [`PyInstaller`](https://pyinstaller.org/) —— 打包成单文件 exe
