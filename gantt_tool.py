# -*- coding: utf-8 -*-
"""
施工进度甘特图工具 (桌面版)
=======================
基于 Python 自带 Tkinter 的本地甘特图程序，不需要额外安装环境。
功能：任务清单管理 + 甘特图绘制 + 搭接关系(前置任务+搭接时间) + 父子任务/标准层自动生成 + 保存/导入(json) + 导出Excel。
运行：双击桌面快捷方式，或执行  python gantt_tool.py
版本管理：改动源码后，双击「重新打包.bat」生成新版 exe；
          历史检查点见「版本管理.bat」，自动保留最近 10 个版本。
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
import os
import sys
from datetime import datetime, date, timedelta

VERSION = "1.1.0"
TITLE = "施工进度甘特图工具"

# 列表层级缩进（用空格模拟，增强可读性）
#   分部工程 -> 1 格 / 主任务 -> 2 格 / 子任务 -> 3 格
IND_SECTION = "    "
IND_MAIN = "        "
IND_SUB = "            "
APP_TITLE = "%s v%s" % (TITLE, VERSION)
COLORS = ["#4FA892", "#6A9FCB", "#DFA0B4", "#C9B458", "#8A7FB8", "#4FA8A0",
          "#B47A50", "#5FA8D3", "#C75B7A", "#6B8E5A", "#9B7EBD", "#D3A05F"]


def _app_dir():
    """程序所在的安装目录（只用于找 exe/源码，不再用来存数据）。
    - 源码运行：脚本所在目录。
    - 打包成 exe(PyInstaller) 后：exe 所在目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# 数据目录名（仅"发布版"使用；放在用户"文档"下，避免用户误删程序旁边的文件）
DATA_DIR_NAME = "Project Gante"

# 发布版标记文件：打包发布版时会把 IDE_DATA_IN_DOCUMENTS 写进 exe；
# 也可以用"程序旁边放一个此文件"来标记（便于手动区分）。
PORTABLE_FLAG = "portable.flag"


def _is_frozen():
    """是否运行在打包后的 exe 里（PyInstaller 会设 sys.frozen）"""
    return bool(getattr(sys, "frozen", False))


def _is_release_build():
    """是否为【发布版】（数据存到"我的文档\\Project Gante"）。

    区分规则（按优先级）：
      ① 源码直接运行（python gantt_tool.py）        → 开发版，数据在程序旁
      ② 程序旁边有 portable.flag 文件               → 便携版，数据在程序旁
      ③ 打包运行且没有该标记                        → 发布版，数据在文档目录

    这样"绿色版/开发版"和"单文件发布版"就能各走各的数据策略，
    哪怕两者都是打包成 exe 的。
    """
    if not _is_frozen():
        return False                      # ① 源码运行 = 开发版
    if os.path.exists(os.path.join(_app_dir(), PORTABLE_FLAG)):
        return False                      # ② 有标记 = 便携/开发版
    return True                           # ③ 打包且无标记 = 发布版


def _data_dir():
    """数据目录。**开发版与发布版策略不同**：

    · 开发版 / 绿色版（源码运行，或程序旁有 portable.flag）
      → 数据放【程序所在文件夹】，跟着工程走。
        便于整体拷贝、调试，也便于当作模块嵌进更大的项目。

    · 发布版（单文件 exe，无标记）
      → 数据放【我的文档\\Project Gante】，自动创建。
        原因：非技术用户容易把 exe 旁边的 json 当垃圾文件删掉；
        而且 exe 可以随便移动/换版本，数据位置始终不变。

    若"文档"目录不可用，发布版会依次回退：Documents → 文档 → 用户主目录 → 程序目录。
    """
    if not _is_release_build():
        return _app_dir()          # 开发版/绿色版：数据就在程序旁边

    candidates = []
    home = os.path.expanduser("~")
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
        val, _ = winreg.QueryValueEx(key, "Personal")
        winreg.CloseKey(key)
        val = os.path.expandvars(val)
        if val:
            candidates.append(val)
    except Exception:
        pass
    candidates.append(os.path.join(home, "Documents"))     # 兜底
    candidates.append(os.path.join(home, "文档"))          # 中文系统兜底
    candidates.append(home)                                # 再兜底
    candidates.append(_app_dir())                          # 最后退回程序目录

    for base in candidates:
        try:
            d = os.path.join(base, DATA_DIR_NAME)
            os.makedirs(d, exist_ok=True)
            # 验证真的可写
            probe = os.path.join(d, ".write_test")
            with open(probe, "w", encoding="utf-8") as f:
                f.write("ok")
            os.remove(probe)
            return d
        except Exception:
            continue
    return _app_dir()


def _migrate_old_data(new_dir):
    """把旧版本留在程序旁边的数据搬到新位置（**仅发布版会发生**）。

    只在目标目录还没有数据时才搬，避免覆盖用户正在用的数据。"""
    target = os.path.join(new_dir, "autosave.json")
    if os.path.exists(target):
        return None
    if os.path.abspath(new_dir) == os.path.abspath(_app_dir()):
        return None                      # 目标就是程序目录，无需搬
    old_candidates = [
        os.path.join(_app_dir(), "autosave.json"),
        os.path.join(os.path.expanduser("~"), "autosave.json"),
    ]
    for old in old_candidates:
        try:
            if os.path.exists(old) and os.path.abspath(old) != os.path.abspath(target):
                import shutil
                shutil.copy2(old, target)
                return old
        except Exception:
            continue
    return None


def _parse_date(text):
    try:
        return datetime.strptime(text.strip(), "%Y-%m-%d").date()
    except ValueError:
        return None


def _end_of_task(t):
    """任务结束日期 = 开始 + 工期 - 1"""
    s = _parse_date(t.get("start", ""))
    if not s:
        return None
    return s + timedelta(days=int(t.get("days", 1)) - 1)


def _effective_end(t, all_tasks):
    """任务的实际结束日期：父任务有子任务时用【汇总结束](最晚子任务结束)，
    否则用自身 start+days。all_tasks 为任务列表(用于查找父任务的子任务)。"""
    if t.get("level", 0) == 0:
        # 找该父任务的子任务
        children = [c for c in all_tasks
                    if c.get("level", 0) == 1 and c.get("parent") == t.get("name")]
        if children:
            ends = [_end_of_task(c) for c in children]
            ends = [e for e in ends if e]
            if ends:
                return max(ends)
    return _end_of_task(t)


def _next_group_id():
    """生成一个唯一 group id（子任务族标识）。用固定前缀+递增。"""
    _next_group_id.counter = getattr(_next_group_id, "counter", 0) + 1
    return "g%d" % _next_group_id.counter


def _new_group_id(tasks):
    """生成不与现有任务 group 冲突的 group id：取现有 g 前缀数字最大值+1。"""
    nums = []
    for t in tasks:
        g = t.get("group", "")
        if isinstance(g, str) and g.startswith("g") and g[1:].isdigit():
            nums.append(int(g[1:]))
    return "g%d" % ((max(nums) if nums else 0) + 1)


def _parent_id_split(t):
    """返回任务的 (父id, 子序号)，父任务返回 (id, None)。"""
    iid = t.get("_id", "")
    if "@" in iid:
        pid, sub = iid.split("@", 1)
        return pid, sub
    return iid, None


def _new_parent_id(tasks):
    """生成新的父任务 id：取现有纯数字父 id 最大值+1。"""
    nums = []
    for t in tasks:
        pid, _ = _parent_id_split(t)
        if pid and pid.isdigit():
            nums.append(int(pid))
    return str((max(nums) if nums else 0) + 1)


def _new_sub_id(parent_id, tasks):
    """生成子任务 id = '父id@序号'，序号 = 该父任务下已有子任务最大序号+1。"""
    max_num = 0
    for t in tasks:
        pid, sub = _parent_id_split(t)
        if pid == parent_id and sub is not None and sub.isdigit():
            max_num = max(max_num, int(sub))
    return "%s@%d" % (parent_id, max_num + 1)


def _first_floor(group, tasks):
    """返回 group 子任务族的最小层号（用于判断标准层首层）"""
    floors = [t.get("floor") for t in tasks if t.get("group") == group and t.get("floor") is not None]
    return min(floors) if floors else None


def _floor_label(floor):
    """楼层显示：负为地下，正为地上。用 B1/B2 表示地下，直接用数字表示地上。"""
    if floor < 0:
        return "B%d" % abs(floor)
    return "%dF" % floor


class PredecessorSelector(ttk.Frame):
    """三组联动前置任务选择器：父任务组 / 子任务组(点父才显示) / 已选中组。
    双击任务从原组移到已选中组，双击已选组移回原组。"""

    def __init__(self, master, all_tasks, exclude_names=None, forbid_parent=None):
        """
        all_tasks: 全部任务(含父任务和子任务)
        exclude_names: 需要排除的任务名(如自身)，避免循环依赖
        forbid_parent: 当前父任务名——该父任务本身不能被双击选中(避免循环)，
                       但可以单击它来显示其下的子任务。
        """
        super().__init__(master)
        self.all_tasks = all_tasks
        self.exclude_names = exclude_names or []
        self.forbid_parent = forbid_parent   # 当前父任务名，禁止作为前置
        self.sel_parent_map = {}   # 子任务组的父任务名 -> 子任务列表
        self.selected = []         # 已选中任务名列表

        # ---- 搜索框（筛选父任务/子任务列表，快速定位） ----
        sf = ttk.Frame(self)
        sf.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 4))
        ttk.Label(sf, text="🔍 搜索：").pack(side="left")
        self.var_search = tk.StringVar()
        ent = ttk.Entry(sf, textvariable=self.var_search, width=24)
        ent.pack(side="left", padx=(2, 4))
        ent.bind("<KeyRelease>", lambda e: self._apply_filter())
        ttk.Button(sf, text="✕", width=3,
                   command=self._clear_search).pack(side="left")
        ttk.Label(sf, text="（输入即筛选；清空恢复全部）",
                  foreground="#6B7280").pack(side="left", padx=6)

        # 三个 Listbox + 标签
        self.lbl_parent = ttk.Label(self, text="父任务（双击选中）")
        self.lbl_parent.grid(row=1, column=0, sticky="w")
        self.lst_parent = tk.Listbox(self, height=7, exportselection=False)
        self.lst_parent.grid(row=2, column=0, sticky="nsew", padx=(0, 6))
        self.lst_parent.bind("<Double-Button-1>", lambda e: self._pick_parent())
        self.lst_parent.bind("<ButtonRelease-1>", lambda e: self._on_parent_click())

        self.lbl_sub = ttk.Label(self, text="子任务（点击父任务后显示，双击选中）")
        self.lbl_sub.grid(row=1, column=1, sticky="w")
        self.lst_sub = tk.Listbox(self, height=7, exportselection=False)
        self.lst_sub.grid(row=2, column=1, sticky="nsew", padx=(0, 6))
        self.lst_sub.bind("<Double-Button-1>", lambda e: self._pick_sub())

        self.lbl_sel = ttk.Label(self, text="已选中（双击移回）")
        self.lbl_sel.grid(row=1, column=2, sticky="w")
        self.lst_sel = tk.Listbox(self, height=7, exportselection=False)
        self.lst_sel.grid(row=2, column=2, sticky="nsew")
        self.lst_sel.bind("<Double-Button-1>", lambda e: self._unpick())

        for i in range(3):
            self.grid_columnconfigure(i, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self._refresh_parent()
        self._refresh_selected()

    # ---- 搜索筛选 ----
    def _kw(self):
        return (self.var_search.get() or "").strip().lower()

    def _clear_search(self):
        self.var_search.set("")
        self._apply_filter()

    def _apply_filter(self):
        """按关键字筛选父任务与子任务列表（已选中组不受影响）

        贴心之处：若关键字只匹配子任务（父任务名不含关键字），
        就把这些子任务所在的父任务也一并列出来 —— 否则会出现
        "父任务列表空着、子任务却搜到了却看不到"的尴尬。
        """
        kw = self._kw()
        self.lst_parent.delete(0, tk.END)
        self.lst_sub.delete(0, tk.END)
        self.sel_parent_map = {}
        if not kw:
            for t in self._parent_items():
                self.lst_parent.insert(tk.END, t["name"])
            return

        # ① 名字本身命中的父任务
        hit_names = [t["name"] for t in self._parent_items()
                     if kw in t["name"].lower()]
        # ② 有子任务名命中的父任务（即使父任务名不含关键字）
        for t in self._parent_items():
            pn = t.get("name")
            if pn in hit_names:
                continue
            if any(kw in c.get("name", "").lower()
                   for c in self.all_tasks
                   if c.get("level", 0) == 1 and c.get("parent") == pn
                   and c.get("name") not in self.exclude_names):
                hit_names.append(pn)

        for n in hit_names:
            self.lst_parent.insert(tk.END, n)

        # ③ 若只有一个父任务，自动选中它并列出匹配的子任务（省一次点击）
        if len(hit_names) == 1:
            self.lst_parent.selection_set(0)
            self._on_parent_click()

    def _refresh_sub(self):
        """重刷子任务组（应用当前筛选）"""
        kw = self._kw()
        self.lst_sub.delete(0, tk.END)
        self.sel_parent_map = {}
        for pname in self._current_parents():
            subs = [t for t in self.all_tasks
                    if t.get("level", 0) == 1 and t.get("parent") == pname
                    and t.get("name") not in self.exclude_names
                    and t.get("name") not in self.selected
                    and (not kw or kw in t["name"].lower())]
            if subs:
                self.sel_parent_map[pname] = subs
                for t in subs:
                    self.lst_sub.insert(tk.END, t["name"])

    def _current_parents(self):
        """当前在父任务列表里显示的父任务名（考虑筛选与选中）"""
        sel = self.lst_parent.curselection()
        if sel:
            return [self.lst_parent.get(sel[0])]
        return [self.lst_parent.get(i) for i in range(self.lst_parent.size())]

    def _parent_items(self):
        return [t for t in self.all_tasks
                if t.get("level", 0) == 0 and t.get("name") not in self.exclude_names]

    def _refresh_parent(self):
        """重刷父任务组（应用当前搜索筛选）"""
        kw = self._kw() if hasattr(self, "var_search") else ""
        self.lst_parent.delete(0, tk.END)
        for t in self._parent_items():
            if not kw or self.var_search.get().strip().lower() in t["name"].lower():
                self.lst_parent.insert(tk.END, t["name"])
        self.lst_sub.delete(0, tk.END)
        self.sel_parent_map = {}

    def _on_parent_click(self):
        """点击父任务：子任务组显示该父任务下的子任务（套用当前搜索筛选）"""
        sel = self.lst_parent.curselection()
        if not sel:
            return
        parent_name = self.lst_parent.get(sel[0])
        kw = self._kw()
        self.lst_sub.delete(0, tk.END)
        self.sel_parent_map = {}
        subs = [t for t in self.all_tasks
                if t.get("level", 0) == 1 and t.get("parent") == parent_name
                and t.get("name") not in self.exclude_names
                and t.get("name") not in self.selected
                and (not kw or kw in t.get("name", "").lower())]
        self.sel_parent_map[parent_name] = subs
        for t in subs:
            self.lst_sub.insert(tk.END, t["name"])

    def _pick_parent(self):
        sel = self.lst_parent.curselection()
        if not sel:
            return
        name = self.lst_parent.get(sel[0])
        # 禁止把当前父任务本身作为前置（避免循环）
        if self.forbid_parent and name == self.forbid_parent:
            messagebox.showinfo("提示", "不能把当前父任务本身作为前置条件（会形成循环）",
                                parent=self.winfo_toplevel())
            return
        self._move_to_selected(name)
        self._on_parent_click()  # 刷新子任务组

    def _pick_sub(self):
        sel = self.lst_sub.curselection()
        if not sel:
            return
        name = self.lst_sub.get(sel[0])
        self._move_to_selected(name)
        # 刷新子任务组（隐藏已选中）
        if self.sel_parent_map:
            for pname in self.sel_parent_map:
                self._on_parent_click()
                break

    def _pick_selected(self):
        pass

    def _unpick(self):
        sel = self.lst_sel.curselection()
        if not sel:
            return
        name = self.lst_sel.get(sel[0])
        if name in self.selected:
            self.selected.remove(name)
        self._refresh_selected()
        self._refresh_parent()  # 重新显示父任务列表

    def _move_to_selected(self, name):
        if name not in self.selected:
            self.selected.append(name)
        self._refresh_selected()
        self._refresh_parent()

    def _refresh_selected(self):
        self.lst_sel.delete(0, tk.END)
        for n in self.selected:
            self.lst_sel.insert(tk.END, n)


class ShutdownDialog(tk.Toplevel):
    """管理停歇期：列表显示 + 添加/删除。返回更新后的 shutdowns。"""

    def __init__(self, parent, shutdowns):
        super().__init__(parent)
        self.title("停歇期管理")
        self.resizable(False, False)
        self.shutdowns = [dict(s) for s in (shutdowns or [])]
        self.result = None

        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="已添加的停歇期：").grid(row=0, column=0, columnspan=3, sticky="w")

        self.lst = tk.Listbox(frm, width=48, height=8, exportselection=False)
        self.lst.grid(row=1, column=0, columnspan=3, sticky="we", pady=(4, 6))
        sb = ttk.Scrollbar(frm, orient="vertical", command=self.lst.yview)
        sb.grid(row=1, column=3, sticky="ns")
        self.lst.config(yscrollcommand=sb.set)
        self._refresh_list()

        # 添加入口
        addf = ttk.Frame(frm)
        addf.grid(row=2, column=0, columnspan=3, sticky="we", pady=(4, 0))
        ttk.Label(addf, text="名称：").pack(side="left")
        self.var_name = tk.StringVar()
        ttk.Entry(addf, textvariable=self.var_name, width=12).pack(side="left", padx=(2, 8))
        ttk.Label(addf, text="开始：").pack(side="left")
        self.var_start = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(addf, textvariable=self.var_start, width=11).pack(side="left", padx=(2, 8))
        ttk.Label(addf, text="结束：").pack(side="left")
        self.var_end = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(addf, textvariable=self.var_end, width=11).pack(side="left", padx=(2, 8))
        ttk.Button(addf, text="添加", command=self._add).pack(side="left", padx=6)

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=3, pady=(14, 0), sticky="e")
        ttk.Button(btns, text="删除选中", command=self._del).pack(side="left", padx=4)
        ttk.Button(btns, text="应用并关闭", command=self._ok).pack(side="left", padx=4)
        ttk.Button(btns, text="取消", command=self._cancel).pack(side="left", padx=4)

        ttk.Label(frm, text="说明：应用后，开始/结束落在停歇期内的任务，其源头工作移到停歇结束日+1。",
                  foreground="#888888", wraplength=360).grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        self.grab_set()
        self.transient(parent)
        self.wait_window()

    def _refresh_list(self):
        self.lst.delete(0, tk.END)
        for s in self.shutdowns:
            self.lst.insert(tk.END, "%s：%s ~ %s" % (s.get("name", ""), s.get("start"), s.get("end")))

    def _add(self):
        name = self.var_name.get().strip() or "停歇期"
        s = _parse_date(self.var_start.get())
        e = _parse_date(self.var_end.get())
        if not s or not e:
            messagebox.showwarning("提示", "开始/结束日期格式应为 YYYY-MM-DD", parent=self)
            return
        if e < s:
            messagebox.showwarning("提示", "结束日期不能早于开始日期", parent=self)
            return
        self.shutdowns.append({"name": name, "start": s.isoformat(), "end": e.isoformat()})
        self._refresh_list()

    def _del(self):
        sel = self.lst.curselection()
        if not sel:
            messagebox.showinfo("提示", "请先选中一个停歇期", parent=self)
            return
        del self.shutdowns[sel[0]]
        self._refresh_list()

    def _ok(self):
        self.result = self.shutdowns
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class SubTaskDialog(tk.Toplevel):
    """从父任务新建子任务的对话框：特殊层(1层) / 标准层(层数+起始层号+每层工期 + 主前置任务)"""

    def __init__(self, parent, parent_task, candidates):
        super().__init__(parent)
        self.title("新建子任务 - " + (parent_task.get("name", "")))
        self.resizable(False, False)
        self.result = None
        self.parent_task = parent_task
        # candidates = 全部任务(含子任务)，供前置选择器使用
        self.candidates = candidates or []
        self.all_tasks = self.candidates

        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="子任务名称：").grid(row=0, column=0, sticky="w", pady=4)
        self.var_name = tk.StringVar(value="")
        ttk.Entry(frm, textvariable=self.var_name, width=32).grid(row=0, column=1, columnspan=2, pady=4, padx=(6, 0), sticky="we")

        # 模式选择
        ttk.Label(frm, text="类型：").grid(row=1, column=0, sticky="w", pady=4)
        self.var_mode = tk.StringVar(value="standard")
        mode_frame = ttk.Frame(frm)
        mode_frame.grid(row=1, column=1, columnspan=2, sticky="w", pady=4, padx=(6, 0))
        ttk.Radiobutton(mode_frame, text="标准层(生成多层)", variable=self.var_mode, value="standard",
                        command=self._on_mode).pack(side="left", padx=4)
        ttk.Radiobutton(mode_frame, text="特殊层(单层)", variable=self.var_mode, value="special",
                        command=self._on_mode).pack(side="left", padx=4)

        # ---- 起止方式选择：搭接关系 / 手动日期（单选互斥）----
        ttk.Separator(frm).grid(row=2, column=0, columnspan=3, sticky="we", pady=6)
        # 方式单选按钮
        self.var_method = tk.StringVar(value="a")
        method_frame = ttk.Frame(frm)
        method_frame.grid(row=3, column=0, columnspan=3, sticky="w")
        ttk.Radiobutton(method_frame, text="方式A：搭接关系+工期+延迟",
                        variable=self.var_method, value="a", command=self._on_method).pack(side="left", padx=4)
        ttk.Radiobutton(method_frame, text="方式B：手动输入起止日期+工期",
                        variable=self.var_method, value="b", command=self._on_method).pack(side="left", padx=4)

        # 起始层号 / 结束层号 / 每层工期
        ttk.Label(frm, text="起始层号：").grid(row=4, column=0, sticky="w", pady=4)
        self.var_start_floor = tk.StringVar(value="2")
        self.ent_start_floor = ttk.Entry(frm, textvariable=self.var_start_floor, width=12)
        self.ent_start_floor.grid(row=4, column=1, sticky="w", pady=4, padx=(6, 0))

        ttk.Label(frm, text="结束层号：").grid(row=5, column=0, sticky="w", pady=4)
        self.var_end_floor = tk.StringVar(value="20")
        self.ent_end_floor = ttk.Entry(frm, textvariable=self.var_end_floor, width=12)
        self.ent_end_floor.grid(row=5, column=1, sticky="w", pady=4, padx=(6, 0))

        ttk.Label(frm, text="每层工期(天)：").grid(row=6, column=0, sticky="w", pady=4)
        self.var_days_per_floor = tk.StringVar(value="7")
        self.ent_days_per_floor = ttk.Entry(frm, textvariable=self.var_days_per_floor, width=12)
        self.ent_days_per_floor.grid(row=6, column=1, sticky="w", pady=4, padx=(6, 0))

        # 特殊层号（仅特殊层可填，标准层禁用）
        ttk.Label(frm, text="特殊层号（选填）：").grid(row=7, column=0, sticky="w", pady=4)
        self.var_special_floor = tk.StringVar(value="1")
        self.ent_special_floor = ttk.Entry(frm, textvariable=self.var_special_floor, width=12)
        self.ent_special_floor.grid(row=7, column=1, sticky="w", pady=4, padx=(6, 0))
        ttk.Label(frm, text="（填层号→如「垫层1F」；留空→不带层号，如「垫层」）",
                  foreground="#888888").grid(row=7, column=2, sticky="w", pady=4, padx=(6, 0))

        ttk.Label(frm, text="搭接时间(天)：").grid(row=8, column=0, sticky="w", pady=4)
        self.var_lag = tk.StringVar(value="0")
        self.ent_lag = ttk.Entry(frm, textvariable=self.var_lag, width=12)
        self.ent_lag.grid(row=8, column=1, sticky="w", pady=4, padx=(6, 0))

        # 前置任务选择器（三组联动：父任务/子任务/已选中）
        ttk.Label(frm, text="前置任务：").grid(row=9, column=0, sticky="nw", pady=4)
        self.pre_selector = PredecessorSelector(frm, self.all_tasks,
                                                exclude_names=[],
                                                forbid_parent=parent_task.get("name"))
        self.pre_selector.grid(row=9, column=1, columnspan=2, sticky="we", pady=4, padx=(6, 0))

        # ---- 方式B：手动日期 ----
        ttk.Label(frm, text="【方式B：手动输入起止日期】", foreground="#555555").grid(
            row=11, column=0, columnspan=3, sticky="w")

        ttk.Label(frm, text="开始日期：").grid(row=12, column=0, sticky="w", pady=4)
        self.var_m_start = tk.StringVar(value="")
        self.ent_m_start = ttk.Entry(frm, textvariable=self.var_m_start, width=14)
        self.ent_m_start.grid(row=12, column=1, sticky="w", pady=4, padx=(6, 0))

        ttk.Label(frm, text="结束日期：").grid(row=13, column=0, sticky="w", pady=4)
        self.var_m_end = tk.StringVar(value="")
        self.ent_m_end = ttk.Entry(frm, textvariable=self.var_m_end, width=14)
        self.ent_m_end.grid(row=13, column=1, sticky="w", pady=4, padx=(6, 0))

        ttk.Label(frm, text="工期(天)：").grid(row=14, column=0, sticky="w", pady=4)
        self.var_m_days = tk.StringVar(value="")
        self.ent_m_days = ttk.Entry(frm, textvariable=self.var_m_days, width=14)
        self.ent_m_days.grid(row=14, column=1, sticky="w", pady=4, padx=(6, 0))

        ttk.Label(frm, text="手动日期与工期联动。", foreground="#888888").grid(
            row=15, column=0, columnspan=3, sticky="w", pady=(4, 0))

        btns = ttk.Frame(frm)
        btns.grid(row=16, column=0, columnspan=3, pady=(14, 0), sticky="e")
        ttk.Button(btns, text="生成", command=self._ok).pack(side="left", padx=4)
        ttk.Button(btns, text="取消", command=self._cancel).pack(side="left", padx=4)

        # 手动日期联动
        self.var_m_start.trace_add("write", lambda *a: self._sync_manual_from_dates())
        self.var_m_end.trace_add("write", lambda *a: self._sync_manual_from_dates())
        self.var_m_days.trace_add("write", lambda *a: self._sync_manual_from_days())

        self._on_mode()
        self.grab_set()
        self.transient(parent)
        self.wait_window()

    def _on_mode(self):
        std = self.var_mode.get() == "standard"
        self.ent_start_floor.config(state="normal" if std else "disabled")
        self.ent_end_floor.config(state="normal" if std else "disabled")
        self.ent_days_per_floor.config(state="normal")  # 特殊层也要工期，保持可用
        self.ent_special_floor.config(state="normal" if not std else "disabled")
        self._on_method()

    def _on_method(self):
        """方式A/方式B 互斥：点亮哪个，另一个变灰不可填。
        方式A控件 = 起始层号/层数/每层工期/特殊层号/延迟/前置选择器；
        方式B控件 = 手动日期三框。
        注意：起始层号/层数仅在【标准层】模式下可填（特殊层禁用）。"""
        method = getattr(self, "var_method", tk.StringVar(value="a"))
        is_a = method.get() == "a"
        std = self.var_mode.get() == "standard"
        # 方式A控件状态
        a_state = "normal" if is_a else "disabled"
        # 起始层号/结束层号：需方式A 且 标准层
        self.ent_start_floor.config(state="normal" if (is_a and std) else "disabled")
        self.ent_end_floor.config(state="normal" if (is_a and std) else "disabled")
        # 特殊层号：仅方式A 且 特殊层 可填
        self.ent_special_floor.config(state="normal" if (is_a and not std) else "disabled")
        # 每层工期：仅需方式A
        self.ent_days_per_floor.config(state=a_state)
        # 搭接时间 + 前置选择器（若无引用则跳过）
        if hasattr(self, 'ent_lag'):
            self.ent_lag.config(state=a_state)
        if hasattr(self, 'pre_selector'):
            for lb in [self.pre_selector.lst_parent, self.pre_selector.lst_sub,
                       self.pre_selector.lst_sel]:
                lb.config(state=a_state)
        # 方式B控件状态
        b_state = "normal" if not is_a else "disabled"
        for w in [self.ent_m_start, self.ent_m_end, self.ent_m_days]:
            w.config(state=b_state)

    def _sync_manual_from_dates(self):
        """手动方式：开始/结束 -> 工期"""
        if getattr(self, "_m_internal", False):
            return
        s = _parse_date(self.var_m_start.get())
        e = _parse_date(self.var_m_end.get())
        if s and e and e >= s:
            self._m_internal = True
            self.var_m_days.set(str((e - s).days + 1))
            self._m_internal = False

    def _sync_manual_from_days(self):
        """手动方式：开始+工期 -> 结束"""
        if getattr(self, "_m_internal", False):
            return
        s = _parse_date(self.var_m_start.get())
        try:
            days = int(self.var_m_days.get())
        except ValueError:
            return
        if s and days >= 1:
            self._m_internal = True
            self.var_m_end.set((s + timedelta(days=days - 1)).isoformat())
            self._m_internal = False

    def _ok(self):
        name = self.var_name.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入子任务名称", parent=self)
            return
        mode = self.var_mode.get()
        try:
            lag = int(self.var_lag.get())
        except ValueError:
            lag = 0
        # 前置 = 三组选择器已选中的任务名列表
        pre_list = list(self.pre_selector.selected) if hasattr(self, 'pre_selector') else []

        # 手动日期：仅当选择方式B时生效（方式A/B互斥）
        manual = None
        if hasattr(self, 'var_method') and self.var_method.get() == "b":
            manual_start = self.var_m_start.get().strip() if hasattr(self, 'var_m_start') else ""
            manual_days = self.var_m_days.get().strip() if hasattr(self, 'var_m_days') else ""
            if not manual_start:
                messagebox.showwarning("提示", "方式B：请填写开始日期", parent=self)
                return
            s = _parse_date(manual_start)
            if s is None:
                messagebox.showwarning("提示", "方式B：开始日期格式应为 YYYY-MM-DD", parent=self)
                return
            try:
                mdays = int(manual_days) if manual_days else 7
            except ValueError:
                mdays = 7
            manual = {"start": s.isoformat(), "days": mdays}

        if mode == "special":
            # 特殊层号：非必填。填了按层号命名(如"垫层1F")；留空则创建不带层号的子任务(如"垫层")
            sf_text = self.var_special_floor.get().strip()
            if sf_text:
                try:
                    sfloor = int(sf_text)
                except ValueError:
                    messagebox.showwarning("提示", "特殊层号应为整数，或留空表示不带层号", parent=self)
                    return
            else:
                sfloor = None   # 不带层号
            try:
                days = int(self.var_days_per_floor.get())
            except ValueError:
                days = 7
            self.result = {"mode": "special", "name": name, "floor": sfloor, "days": days,
                           "lag": lag, "pre": pre_list, "manual": manual}
        else:
            try:
                sfloor = int(self.var_start_floor.get())
                efloor = int(self.var_end_floor.get())
            except ValueError:
                messagebox.showwarning("提示", "起始层号和结束层号应为整数", parent=self)
                return
            try:
                days = int(self.var_days_per_floor.get())
            except ValueError:
                days = 7
            if efloor < sfloor:
                messagebox.showwarning("提示", "结束层号不能小于起始层号", parent=self)
                return
            count = efloor - sfloor + 1
            # 记录 start_floor 和 count（count = 结束-起始+1）
            self.result = {"mode": "standard", "name": name, "start_floor": sfloor,
                           "end_floor": efloor, "count": count, "days": days,
                           "lag": lag, "pre": pre_list, "manual": manual}
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class TaskDialog(tk.Toplevel):
    """自定义任务编辑对话框：名称/开始/结束/工期 + 前置任务/搭接时间 在一个窗口内填"""

    def __init__(self, parent, title=None, default=None, candidates=None):
        super().__init__(parent)
        self.title(title or "任务信息")
        self.resizable(False, False)
        self.result = None
        self.candidates = candidates or []

        default = default or {}
        init_start = default.get("start") or date.today().isoformat()
        init_days = int(default.get("days", 1))
        init_start_d = _parse_date(init_start) or date.today()
        init_end_d = init_start_d + timedelta(days=init_days - 1)
        init_pre = default.get("pre", [])
        init_lag = default.get("lag", 0)

        frm = ttk.Frame(self, padding=14)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="任务名称：").grid(row=0, column=0, sticky="w", pady=4)
        self.var_name = tk.StringVar(value=default.get("name", ""))
        ttk.Entry(frm, textvariable=self.var_name, width=34).grid(row=0, column=1, columnspan=2, pady=4, padx=(6, 0), sticky="we")

        ttk.Label(frm, text="开始日期：").grid(row=1, column=0, sticky="w", pady=4)
        self.var_start = tk.StringVar(value=init_start_d.isoformat())
        ttk.Entry(frm, textvariable=self.var_start, width=34).grid(row=1, column=1, columnspan=2, pady=4, padx=(6, 0), sticky="we")

        ttk.Label(frm, text="结束日期：").grid(row=2, column=0, sticky="w", pady=4)
        self.var_end = tk.StringVar(value=init_end_d.isoformat())
        ttk.Entry(frm, textvariable=self.var_end, width=34).grid(row=2, column=1, columnspan=2, pady=4, padx=(6, 0), sticky="we")

        ttk.Label(frm, text="工期(天)：").grid(row=3, column=0, sticky="w", pady=4)
        self.var_days = tk.StringVar(value=str(init_days))
        ttk.Entry(frm, textvariable=self.var_days, width=34).grid(row=3, column=1, columnspan=2, pady=4, padx=(6, 0), sticky="we")

        ttk.Separator(frm).grid(row=4, column=0, columnspan=3, sticky="we", pady=8)

        ttk.Label(frm, text="前置任务(可多选，点击父任务看子任务)：").grid(row=5, column=0, sticky="nw", pady=4)
        self.pre_selector = PredecessorSelector(frm, self.candidates, exclude_names=[default.get("name")])
        self.pre_selector.grid(row=5, column=1, columnspan=2, sticky="we", pady=4, padx=(6, 0))
        # 预选：编辑时把已有前置加入已选中组
        for pre in init_pre:
            if pre not in self.pre_selector.selected:
                self.pre_selector.selected.append(pre)
        self.pre_selector._refresh_selected()
        self.pre_selector._refresh_parent()

        ttk.Label(frm, text="搭接时间(天)：").grid(row=6, column=0, sticky="w", pady=4)
        self.var_lag = tk.StringVar(value=str(init_lag))
        ttk.Entry(frm, textvariable=self.var_lag, width=10).grid(row=6, column=1, sticky="w", pady=4, padx=(6, 0))

        btn_small = ttk.Frame(frm)
        btn_small.grid(row=6, column=2, sticky="e", pady=4)
        ttk.Button(btn_small, text="应用搭接", width=10, command=self._apply_pred).pack(side="left", padx=2)
        ttk.Button(btn_small, text="清空前置", width=10, command=self._clear_pred).pack(side="left", padx=2)

        ttk.Label(frm, text="自动绑定：开始日期 = 所选前置中最晚结束的日期 + 搭接时间(天)",
                  foreground="#888888").grid(row=7, column=0, columnspan=3, sticky="w", pady=(4, 0))

        btns = ttk.Frame(frm)
        btns.grid(row=8, column=0, columnspan=3, pady=(14, 0), sticky="e")
        ttk.Button(btns, text="确定", command=self._ok).pack(side="left", padx=4)
        ttk.Button(btns, text="取消", command=self._cancel).pack(side="left", padx=4)

        self.var_start.trace_add("write", lambda *a: self._sync_from_dates())
        self.var_end.trace_add("write", lambda *a: self._sync_from_dates())
        self.var_days.trace_add("write", lambda *a: self._sync_from_days())
        self.var_lag.trace_add("write", lambda *a: self._apply_pred(silent=True))

        self.grab_set()
        self.transient(parent)
        self.wait_window()

    def _sync_from_dates(self):
        if getattr(self, "_internal", False):
            return
        s = _parse_date(self.var_start.get())
        e = _parse_date(self.var_end.get())
        if s and e and e >= s:
            self._internal = True
            self.var_days.set(str((e - s).days + 1))
            self._internal = False

    def _sync_from_days(self):
        if getattr(self, "_internal", False):
            return
        s = _parse_date(self.var_start.get())
        try:
            days = int(self.var_days.get())
        except ValueError:
            return
        if s and days >= 1:
            self._internal = True
            self.var_end.set((s + timedelta(days=days - 1)).isoformat())
            self._internal = False

    def _selected_pre_names(self):
        if hasattr(self, 'pre_selector'):
            return list(self.pre_selector.selected)
        return []

    def _clear_pred(self):
        if hasattr(self, 'pre_selector'):
            self.pre_selector.selected = []
            self.pre_selector._refresh_selected()
            self.pre_selector._refresh_parent()

    def _apply_pred(self, silent=False):
        if not silent and getattr(self, "_internal", False):
            return
        names = self._selected_pre_names()
        if not names:
            if not silent:
                messagebox.showinfo("提示", "请先选择至少一个前置任务", parent=self)
            return
        last_end = None
        for c in self.candidates:
            if c.get("name") in names:
                end = _effective_end(c, self.candidates)
                if end and (last_end is None or end > last_end):
                    last_end = end
        if last_end is None:
            if not silent:
                messagebox.showinfo("提示", "未找到所选前置任务的结束日期", parent=self)
            return
        try:
            lag = int(self.var_lag.get())
        except ValueError:
            lag = 0
        start = last_end + timedelta(days=lag + 1)
        self._internal = True
        self.var_start.set(start.isoformat())
        self._sync_from_days()
        self._internal = False

    def _ok(self):
        name = self.var_name.get().strip()
        s = _parse_date(self.var_start.get())
        e = _parse_date(self.var_end.get())
        try:
            days = int(self.var_days.get())
        except ValueError:
            days = None
        if not name:
            messagebox.showwarning("提示", "请输入任务名称", parent=self)
            return
        if not s:
            messagebox.showwarning("提示", "开始日期格式应为 YYYY-MM-DD", parent=self)
            return
        if not e:
            messagebox.showwarning("提示", "结束日期格式应为 YYYY-MM-DD", parent=self)
            return
        if e < s:
            messagebox.showwarning("提示", "结束日期不能早于开始日期", parent=self)
            return
        if days is None or days < 1:
            messagebox.showwarning("提示", "工期应为≥1的整数", parent=self)
            return
        self.result = {
            "name": name,
            "start": s.isoformat(),
            "days": days,
            "pre": self._selected_pre_names(),
            "lag": self._safe_int(self.var_lag.get()),
            "level": 0,
        }
        self.destroy()

    @staticmethod
    def _safe_int(text):
        try:
            return int(text)
        except ValueError:
            return 0

    def _cancel(self):
        self.result = None
        self.destroy()


class GanttApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1280x720")
        self.tasks = []
        self.cur_color_idx = 0
        self.data_file = ""
        self.collapsed = set()   # 被收起子任务的父任务名集合
        self.critical_tasks = set()  # 关键路径上的任务名集合
        self.shutdowns = []      # 停歇期列表 [{name, start, end}]
        # 分部工程：自定义的"类"层级（比任务高一层，非父子关系）
        # 每个元素 {"name": 分部名, "collapsed": 是否折叠}
        self.sections = []
        self.section_collapsed = set()   # 被折叠的分部名（会话内状态即可）
        self._sections_inited = False     # 首次打开时是否已"默认全折叠"
        # 画布缩放系数（Ctrl+滚轮调整；1.0 = 默认大小）
        self.gantt_zoom = 1.0
        self.net_zoom = 1.0
        self._sort_state = {"col": None, "asc": True}  # 列表排序状态
        self.undo_stack = []     # 撤销历史（每次操作前的 tasks 深拷贝）
        self.undo_max = 50       # 最多保留 50 步撤销
        # 自动保存文件（程序同目录）
        # 数据目录：<我的文档>\Project Gante（不放程序旁边，免得被误删）
        self.data_dir = _data_dir()
        self._migrated_from = _migrate_old_data(self.data_dir)   # 旧数据自动搬过来
        self.autosave_path = os.path.join(self.data_dir, "autosave.json")
        # 缓存：规则计算结果存版本化缓存文件，展示/导出用缓存
        self.cache_dir = self.data_dir
        self.cache_ver = 0         # 当前缓存版本号
        self.cache_data = None     # 当前缓存 {任务名: {name,start,end}} 
        self.rule_changed = True   # 规则是否有改动(需重算)
        # DPI 缩放系数(行高/字体自适应)
        try:
            self.dpi_scale = float(os.environ.get("TK_DPI_SCALING", "1.0"))
        except Exception:
            self.dpi_scale = 1.0
        if self.dpi_scale < 1.0:
            self.dpi_scale = 1.0
        # 字体缩放用更温和的系数(避免字体过大)：行距/列宽随 dpi_scale，字体最多×1.2
        self.dpi_font = min(self.dpi_scale, 1.2)

        # ---- 现代工具栏 ----
        from tkinter import font as tkfont
        _scale = getattr(self, "dpi_scale", 1.0)
        _font = getattr(self, "dpi_font", 1.0)
        btnsz = max(10, int(10 * _font))
        try:
            style = ttk.Style(self.root)
            # 现代按钮：圆角、主蓝/次浅/危险红
            style.configure("Tool.TButton",
                            font=("TkDefaultFont", btnsz),
                            padding=(int(12*_scale), int(5*_scale)),
                            background="#3B82F6", foreground="white",
                            relief="flat", borderwidth=0)
            style.map("Tool.TButton",
                      background=[("active", "#2563EB"), ("pressed", "#1D4ED8")],
                      foreground=[("disabled", "#B0BEC5")])
            style.configure("ToolSecondary.TButton",
                            font=("TkDefaultFont", btnsz),
                            padding=(int(10*_scale), int(5*_scale)),
                            background="#E5E7EB", foreground="#1F2937",
                            relief="flat", borderwidth=0)
            style.map("ToolSecondary.TButton",
                      background=[("active", "#D1D5DB"), ("pressed", "#9CA3AF")])
            style.configure("ToolDanger.TButton",
                            font=("TkDefaultFont", btnsz),
                            padding=(int(10*_scale), int(5*_scale)),
                            background="#EF4444", foreground="white",
                            relief="flat", borderwidth=0)
            style.map("ToolDanger.TButton",
                      background=[("active", "#DC2626"), ("pressed", "#B91C1C")])
            # 工具栏容器背景
            style.configure("Toolbar.TFrame", background="#F3F4F6")
        except Exception:
            pass

        # ============================================================
        # 菜单栏（替代原来的三行按钮工具栏）
        #  任务 / 数据 / 视图 / 帮助
        # ============================================================
        # 注意：var_today 要在菜单之前建好，因为「视图」菜单要引用它
        self.var_today = tk.BooleanVar(value=False)
        self.menubar = tk.Menu(self.root)

        m_task = tk.Menu(self.menubar, tearoff=0)
        m_task.add_command(label="新建分部工程…", command=self.new_section)
        m_task.add_separator()
        m_task.add_command(label="添加任务…", accelerator="Ctrl+N",
                           command=self.add_task)
        m_task.add_command(label="新建子任务…", accelerator="Ctrl+Shift+N",
                           command=self.add_subtask)
        m_task.add_separator()
        m_task.add_command(label="编辑任务…", accelerator="Enter",
                           command=self.edit_task)
        m_task.add_command(label="删除任务", accelerator="Delete",
                           command=self.delete_task)
        self.menubar.add_cascade(label="任务(T)", menu=m_task)

        m_data = tk.Menu(self.menubar, tearoff=0)
        m_data.add_command(label="保存…", accelerator="Ctrl+S", command=self.save)
        m_data.add_command(label="导入…", accelerator="Ctrl+O",
                           command=self.import_config)
        m_data.add_separator()
        m_data.add_command(label="导出 Excel…", accelerator="Ctrl+E",
                           command=self.export_excel)
        m_data.add_separator()
        m_data.add_command(label="停歇期设置…", command=self.manage_shutdowns)
        m_data.add_separator()
        m_data.add_command(label="载入示例数据", command=self.load_demo)
        m_data.add_command(label="清空全部任务", command=self.clear_all)
        self.menubar.add_cascade(label="数据(D)", menu=m_data)

        m_view = tk.Menu(self.menubar, tearoff=0)
        m_view.add_checkbutton(label="显示今天竖线", variable=self.var_today,
                               command=self._on_toggle_today)
        m_view.add_separator()
        m_view.add_command(label="放大 (Ctrl+滚轮上)", command=lambda: self._zoom_by(1))
        m_view.add_command(label="缩小 (Ctrl+滚轮下)", command=lambda: self._zoom_by(-1))
        m_view.add_command(label="恢复 100%", command=self._zoom_reset)
        m_view.add_separator()
        m_view.add_command(label="展开全部分部", command=lambda: self._set_all_sections(False))
        m_view.add_command(label="折叠全部分部", command=lambda: self._set_all_sections(True))
        self.menubar.add_cascade(label="视图(V)", menu=m_view)

        m_help = tk.Menu(self.menubar, tearoff=0)
        m_help.add_command(label="使用说明", command=self._show_help)
        m_help.add_command(label="关于 / 版本信息", command=self._show_about)
        m_help.add_separator()
        m_help.add_command(label="打开数据文件夹…", command=self.open_data_folder)
        self.menubar.add_cascade(label="帮助(H)", menu=m_help)

        self.root.config(menu=self.menubar)

        # ============================================================
        # 工具条（一行）：搜索框 + 常用开关
        # 常用操作都在菜单栏里，这里只保留搜索这类高频交互控件
        # ============================================================
        top = ttk.Frame(self.root, style="Toolbar.TFrame", padding=(8, int(5*_scale)))
        top.pack(fill="x")

        # ---------- 搜索框 ----------
        sframe = ttk.Frame(top)
        sframe.pack(side="left", padx=(2, 4))
        ttk.Label(sframe, text="🔍 搜索：").pack(side="left", padx=(0, 2))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(sframe, textvariable=self.search_var,
                                      width=24)
        self.search_entry.pack(side="left")
        ttk.Button(sframe, text="✕", width=3,
                   command=self.clear_search).pack(side="left", padx=(3, 0))
        self.search_entry.bind("<Return>", lambda e: self.do_search())
        self.search_entry.bind("<KeyRelease>", self._on_search_typed)
        self.search_entry.bind("<Escape>", lambda e: self._hide_suggest())
        # 匹配项下拉（用 Toplevel 承载 Listbox，浮在搜索框下方）
        self.suggest_win = None
        self.suggest_box = None

        # ---------- 显示"当前时间"竖线（勾选框放工具栏右侧） ----------
        chk = ttk.Checkbutton(top, text="显示今天", variable=self.var_today,
                              command=self._on_toggle_today)
        chk.pack(side="left", padx=(12, 4))
        # 一键打开数据目录（同事找不到数据在哪时用）
        btn_data = ttk.Button(top, text="📁 数据",
                              command=self.open_data_folder)
        btn_data.pack(side="left", padx=(6, 4))
        # 定时刷新（每分钟更新一次竖线位置，跨天也能自动走到新的一天）
        self._today_job = None
        self._schedule_today_refresh()

        # ---------- 快捷键 ----------
        self.root.bind("<Control-n>", lambda e: self.add_task())
        self.root.bind("<Control-N>", lambda e: self.add_subtask())
        self.root.bind("<Control-s>", lambda e: self.save())
        self.root.bind("<Control-o>", lambda e: self.import_config())
        self.root.bind("<Control-e>", lambda e: self.export_excel())
        self.root.bind("<Control-f>", lambda e: self._focus_search())

        paned = ttk.PanedWindow(self.root, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left = ttk.Frame(paned)
        self.tree = ttk.Treeview(left, columns=("name", "start", "end", "days", "pre", "type"),
                                 show="headings", height=20)
        # 列表行距/字距自适应 DPI 缩放(字体温和，行距/列宽随缩放)
        _scale = getattr(self, "dpi_scale", 1.0)
        _font = getattr(self, "dpi_font", 1.0)
        try:
            style = ttk.Style(self.root)
            style.configure("Treeview", rowheight=int(24 * _scale), font=("TkDefaultFont", max(9, int(9 * _font))))
            style.configure("Treeview.Heading", font=("TkDefaultFont", max(9, int(9 * _font))))
        except Exception:
            pass
        # 列宽随 DPI 放大(保证文字不被截断)
        self.tree.column("name", width=int(210 * _scale))
        self.tree.column("start", width=int(110 * _scale), anchor="center")
        self.tree.column("end", width=int(110 * _scale), anchor="center")
        self.tree.column("days", width=int(90 * _scale), anchor="center")
        self.tree.column("pre", width=int(160 * _scale))
        self.tree.column("type", width=int(90 * _scale), anchor="center")
        self.tree.heading("name", text="任务名称", command=lambda: self._on_column_sort("name"))
        self.tree.heading("start", text="开始日期", command=lambda: self._on_column_sort("start"))
        self.tree.heading("end", text="结束日期", command=lambda: self._on_column_sort("end"))
        self.tree.heading("days", text="工期(天)", command=lambda: self._on_column_sort("days"))
        self.tree.heading("pre", text="搭接关系", command=lambda: self._on_column_sort("pre"))
        self.tree.heading("type", text="类型", command=lambda: self._on_column_sort("type"))
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-1>", self._on_left_click)  # 单击名称左端箭头切换展开
        # 拖拽：把任务拖到「分部」行上，即归入该分部
        self._drag_from = None      # 拖拽起点行 iid
        self._drag_moved = False    # 是否真的移动过（区分"点击"和"拖拽"）
        self.tree.bind("<B1-Motion>", self._on_tree_drag)
        self.tree.bind("<ButtonRelease-1>", self._on_tree_drop)
        self.tree.tag_configure("droptarget", background="#D6E4FF")   # 落点高亮
        # 搜索命中：主目标用纯黄（#FFFF00，与甘特图土黄 #C9B458 感知距离 227，不混淆）
        self.tree.tag_configure("searchhit", background="#FFFF00",
                                foreground="#212121")
        # 其它也含关键字的行：淡黄提示（存在感弱于主目标，不抢视线）
        self.tree.tag_configure("searchdim", background="#FFF9C4",
                                foreground="#5D4037")
        # 右键菜单：新建子任务
        self._menu = tk.Menu(self.tree, tearoff=0)
        self._menu.add_command(label="新建子任务", command=self.add_subtask)
        self._menu.add_command(label="展开/收起子任务", command=self._toggle_right_click)
        self._menu.add_command(label="编辑任务", command=self.edit_task)
        self._menu.add_command(label="删除任务", command=self.delete_task)
        self.tree.bind("<Button-3>", self._on_right_click)  # 右键弹出
        self.tree.pack(fill="both", expand=True)
        paned.add(left, weight=1)

        right = ttk.Frame(paned)
        # 右侧改为选项卡：甘特图 / 搭接关系图(网络图)
        self.notebook = ttk.Notebook(right)
        self.notebook.pack(fill="both", expand=True)

        # ===== 页签1：甘特图 =====
        gantt_page = ttk.Frame(self.notebook)
        self.notebook.add(gantt_page, text="  甘特图  ")
        canvas_frame = ttk.Frame(gantt_page)
        canvas_frame.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(canvas_frame, bg="#FCFEFD", highlightthickness=0)
        self.canvas.configure(width=700, height=700, background="#FCFEFD")
        hbar = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self.canvas.xview)
        vbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.config(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        # 甘特图鼠标左键拖动平移(默认箭头指针，按下拖动时显示移动光标)
        self._gantt_drag = [None, None]
        self.canvas.bind("<ButtonPress-1>", self._on_gantt_press)
        self.canvas.bind("<B1-Motion>", self._on_gantt_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_gantt_release)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        hbar.grid(row=1, column=0, sticky="ew")
        vbar.grid(row=0, column=1, sticky="ns")
        canvas_frame.grid_rowconfigure(0, weight=1)
        canvas_frame.grid_columnconfigure(0, weight=1)

        # ===== 页签2：搭接关系图(网络图) =====
        net_page = ttk.Frame(self.notebook)
        self.notebook.add(net_page, text="  搭接关系图  ")
        net_frame = ttk.Frame(net_page)
        net_frame.pack(fill="both", expand=True)
        self.net_canvas = tk.Canvas(net_frame, bg="#FCFEFD", highlightthickness=0)
        self.net_canvas.configure(background="#FCFEFD")
        nhbar = ttk.Scrollbar(net_frame, orient="horizontal", command=self.net_canvas.xview)
        nvbar = ttk.Scrollbar(net_frame, orient="vertical", command=self.net_canvas.yview)
        self.net_canvas.config(xscrollcommand=nhbar.set, yscrollcommand=nvbar.set)
        self.net_canvas.grid(row=0, column=0, sticky="nsew")
        nhbar.grid(row=1, column=0, sticky="ew")
        nvbar.grid(row=0, column=1, sticky="ns")
        net_frame.grid_rowconfigure(0, weight=1)
        net_frame.grid_columnconfigure(0, weight=1)
        # 网络图也支持鼠标拖动平移
        self._net_drag = [None, None]
        self.net_canvas.bind("<ButtonPress-1>", self._on_net_press)
        self.net_canvas.bind("<B1-Motion>", self._on_net_drag)
        self.net_canvas.bind("<ButtonRelease-1>", self._on_net_release)
        # 滚轮：默认滚动；按住 Ctrl 滚动 = 缩放
        self._bind_canvas_wheel(self.canvas, "gantt")
        self._bind_canvas_wheel(self.net_canvas, "net")
        # 切换到网络图时，若未绘制过则绘制
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        paned.add(right, weight=5)
        # 让右侧甘特图区域更宽：初始分割线放靠左位置
        try:
            paned.sashpos(0, 380)
        except Exception:
            pass

        self.status = tk.StringVar(value="就绪")
        ttk.Label(self.root, textvariable=self.status, relief="sunken", anchor="w").pack(fill="x", side="bottom")

        self.refresh_list()
        # 启动：优先加载自动保存数据（若存在且非空），否则载入示例数据
        if os.path.exists(self.autosave_path):
            try:
                with open(self.autosave_path, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                # 兼容旧格式（纯数组）和新格式（{tasks, shutdowns}）
                if isinstance(saved, dict):
                    self.tasks = saved.get("tasks", [])
                    self.shutdowns = saved.get("shutdowns", [])
                    self.sections = saved.get("sections", []) or []
                else:
                    self.tasks = saved
                    self.shutdowns = []
                    self.sections = []
                if self.tasks:
                    self.status.set(f"已自动恢复上次数据（{len(self.tasks)} 个任务）")
            except Exception:
                self.load_demo()
                self.status.set("自动保存读取失败，已载入示例数据")
        else:
            self.load_demo()
        # 旧数据自动搬迁的提示（一次性，让用户知道数据换地方了）
        if getattr(self, "_migrated_from", None):
            self.root.after(600, lambda: messagebox.showinfo(
                "数据已归位",
                "你的排期数据已自动从旧位置迁移到：\n\n%s\n\n"
                "（旧文件：%s）\n\n"
                "以后请到这里找数据；把这个文件夹备份或同步即可。"
                % (self.data_dir, self._migrated_from),
                parent=self.root))
        # 关闭窗口时自动保存一次
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        # 启动时加载最新缓存(若有)，无则 rule_changed=True 触发重算
        if self._load_latest_cache():
            self.rule_changed = False   # 已有缓存，直接用
        # 默认折叠所有有子任务的父任务
        self._collapse_all_parents()
        # 默认折叠所有分部工程（首次打开只看到分部列表，清爽）
        self.section_collapsed = set(self._section_names())
        self._sections_inited = True
        self.refresh_all()
        # 窗口完全渲染后重绘一次，确保甘特图按真实宽度绘制(4px/天,可横向滚动)
        self.root.after(120, self.draw_gantt)
        # 启动数据体检(延迟弹出，不阻塞启动)
        self.root.after(400, lambda: self._check_data_health("启动"))

    def _on_right_click(self, event):
        """右键：按点中的行类型（分部 / 任务）重建菜单再弹出"""
        row = self.tree.identify_row(event.y)
        self._menu.delete(0, "end")
        if row and self._is_section_row(row):
            sec = self._section_of_row(row)
            self.tree.selection_set(row)
            self.tree.focus(row)
            self._menu.add_command(label="新建分部工程…", command=self.new_section)
            self._menu.add_separator()
            if sec == self.UNSORTED_SEC:
                self._menu.add_command(label="展开/折叠「未分类」",
                                       command=lambda: self._toggle_section(sec))
            else:
                self._menu.add_command(label="展开/折叠「%s」" % sec,
                                       command=lambda: self._toggle_section(sec))
                self._menu.add_command(label="重命名「%s」…" % sec,
                                       command=lambda: self.rename_section(sec))
                self._menu.add_command(label="删除「%s」" % sec,
                                       command=lambda: self.delete_section(sec))
        else:
            if row:
                self.tree.selection_set(row)
                self.tree.focus(row)
            self._menu.add_command(label="新建子任务", command=self.add_subtask)
            self._menu.add_command(label="展开/收起子任务", command=self._toggle_right_click)
            self._menu.add_command(label="编辑任务", command=self.edit_task)
            self._menu.add_command(label="删除任务", command=self.delete_task)
            self._menu.add_separator()
            sub = tk.Menu(self._menu, tearoff=0)
            sub.add_command(label="（未分类）",
                            command=lambda: self._move_selected_to_section(self.UNSORTED_SEC))
            names = self._section_names()
            if names:
                sub.add_separator()
                for s in names:
                    sub.add_command(label=s,
                                    command=lambda s=s: self._move_selected_to_section(s))
            sub.add_separator()
            sub.add_command(label="新建分部…", command=self.new_section)
            self._menu.add_cascade(label="移入分部 ▸", menu=sub)
            self._menu.add_command(label="新建分部工程…", command=self.new_section)
        try:
            self._menu.tk_popup(event.x_root, event.y_root)
        finally:
            self._menu.grab_release()

    def _move_selected_to_section(self, sec):
        """右键菜单：把选中的任务移入指定分部（批量，一次刷新）"""
        rows = list(self.tree.selection())
        real = "" if sec == self.UNSORTED_SEC else sec
        changed = 0
        self._push_undo()
        for r in rows:
            if self._is_section_row(r):
                continue
            t = self._task_by_iid(r)
            if t is None:
                continue
            if t.get("level", 0) == 1:      # 子任务：改它父任务
                pn = t.get("parent")
                t = next((p for p in self.tasks
                          if p.get("level", 0) == 0 and p.get("name") == pn), None)
                if t is None:
                    continue
            if (t.get("section") or "") != real:
                t["section"] = real
                changed += 1
        if changed:
            self.refresh_all()
            self.status.set("已把 %d 个任务归入：%s" % (changed, sec))
        elif self.undo_stack:
            self.undo_stack.pop()           # 没变化，撤销刚才压栈

    # ---------- 数据操作 ----------
    def add_task(self):
        dlg = TaskDialog(self.root, title="添加任务",
                         default={"name": "", "start": date.today().isoformat(), "days": 1},
                         candidates=self.tasks)
        if dlg.result:
            self._push_undo()
            t = dlg.result
            t["color"] = COLORS[self.cur_color_idx % len(COLORS)]
            self.cur_color_idx += 1
            t["_id"] = _new_parent_id(self.tasks)
            t["level"] = 0
            t["parent"] = None
            t["floor"] = None
            self.tasks.append(t)
            self.refresh_all()

    def edit_task(self):
        idx = self._selected_index()
        if idx is None:
            return
        t = self.tasks[idx]
        candidates = [c for i, c in enumerate(self.tasks) if i != idx]
        # 父任务：编辑时显示/调整其"最早子任务"继承来的搭接关系
        default = t
        if t.get("level", 0) == 0 and self._children_of(t.get("name")):
            default = dict(t)
            default["pre"] = self._effective_pre(t)
            default["lag"] = self._effective_lag(t)
        dlg = TaskDialog(self.root, title="编辑任务", default=default, candidates=candidates)
        if dlg.result:
            r = dlg.result
            old_name = t.get("name")
            new_name = r.get("name")
            # 重名检测：新名字与其他任务冲突则阻止（重名会导致前置匹配/选中错乱）
            if new_name != old_name:
                clash = [x for i, x in enumerate(self.tasks) if i != idx and x.get("name") == new_name]
                if clash:
                    messagebox.showerror("重名冲突",
                                         f"已存在名为「{new_name}」的任务。\n任务名称必须唯一，否则搭接关系会错乱。",
                                         parent=self.root)
                    return
            self._push_undo()
            r["color"] = t.get("color", COLORS[0])
            r["_id"] = t.get("_id") or _new_parent_id([x for i, x in enumerate(self.tasks) if i != idx])
            r["level"] = t.get("level", 0)
            r["parent"] = t.get("parent")
            r["floor"] = t.get("floor")
            r["group"] = t.get("group")
            is_parent = (t.get("level", 0) == 0 and self._children_of(t.get("name")))
            if is_parent:
                # 父任务不存自身搭接：把编辑后的搭接写回"最早子任务"
                children = self._children_of(t.get("name"))
                dated = [c for c in children if _parse_date(c.get("start", ""))]
                first = min(dated, key=lambda c: (c.get("start", ""), c.get("name", ""))) if dated else children[0]
                first["pre"] = list(r.get("pre", []))
                first["lag"] = int(r.get("lag", 0) or 0)
                r["pre"] = []          # 父任务自身不保留前置
                r["lag"] = 0
            self.tasks[idx] = r
            # 改名时同步更新所有引用旧名的前置，防止搭接关系断裂
            if new_name != old_name:
                updated = 0
                for x in self.tasks:
                    if x is r:
                        continue
                    pre = x.get("pre", [])
                    if old_name in pre:
                        x["pre"] = [new_name if p == old_name else p for p in pre]
                        updated += 1
                    # 子任务改名时，其子任务的 parent 字段也要同步
                    if x.get("parent") == old_name:
                        x["parent"] = new_name
                if updated:
                    self.status.set(f"已改名并同步更新 {updated} 个任务的前置引用")
            self.refresh_all()

    # ---------- 子任务：标准层/特殊层生成 ----------
    def add_subtask(self):
        parent_idx = self._selected_index()
        if parent_idx is None:
            return
        parent = self.tasks[parent_idx]
        # 前置候选 = 全部任务(含父任务和子任务)，含当前父任务——供点击它选其子任务作前置
        candidates = [c for c in self.tasks]
        dlg = SubTaskDialog(self.root, parent, candidates)
        if not dlg.result:
            return
        r = dlg.result
        self._push_undo()
        group = _new_group_id(self.tasks)
        pre_list = r.get("pre", [])
        manual = r.get("manual")

        warnings = []
        if r["mode"] == "special":
            # 特殊层号为空(None)时不做层号冲突检测，直接创建不带层号的子任务
            if r["floor"] is None:
                self._gen_one_sub(parent, r["name"], None, r["days"], r["lag"],
                                  pre_list, group, warnings, manual, is_first=True)
                self.refresh_all()
                if warnings:
                    messagebox.showwarning("提示", "\n".join(warnings), parent=self.root)
                return
            existing = [
                x for x in self.tasks
                if x.get("level", 0) == 1 and x.get("parent") == parent.get("name")
                and x.get("floor") == r["floor"] and x.get("group") != group
            ]
            action = "create"
            if existing:
                existing_names = "、".join(x.get("name", "") for x in existing)
                choice = messagebox.askyesnocancel(
                    "特殊层号冲突",
                    f"特殊层号 {r['floor']} 与已有子任务（{existing_names}）冲突。\n"
                    f"是否继续新建？\n"
                    f"· 是 = 继续新建（平行存在）\n· 否 = 替换已有层\n· 取消 = 放弃新建",
                    parent=self.root)
                if choice is None:   # 取消
                    return
                elif choice is False:  # 否 = 替换已有层
                    action = "replace"
            if action == "replace" and existing:
                self.tasks = [x for x in self.tasks
                              if not (x.get("level", 0) == 1 and x.get("parent") == parent.get("name")
                                      and x.get("floor") == r["floor"] and x.get("group") != group)]
                for ex in existing:
                    en = ex.get("name")
                    for x in self.tasks:
                        if en in x.get("pre", []):
                            x["pre"] = [p for p in x["pre"] if p != en]
            # 特殊层为单层创建，直接关联所选前置（is_first=True）
            self._gen_one_sub(parent, r["name"], r["floor"], r["days"], r["lag"],
                              pre_list, group, warnings, manual, is_first=True)
        else:
            floor = r["start_floor"]
            first_floor = r["start_floor"]
            for _ in range(r["count"]):
                self._gen_one_sub(parent, r["name"], floor, r["days"], r["lag"], pre_list, group, warnings, manual,
                                  is_first=(floor == first_floor))
                floor += 1
        self.refresh_all()
        if warnings:
            messagebox.showwarning("提示", "\n".join(warnings), parent=self.root)

    def _gen_one_sub(self, parent, base_name, floor, days, lag, pre_list, group, warnings, manual=None, is_first=False):
        """生成单个子任务，命名 基名+层号（无层号时用基名）。
        pre_list: 用户选中的前置任务名列表(可为父任务或子任务)。
        manual: 手动日期 {start, days} 或 None；有则优先采用。
        is_first: 是否标准层的首层(用于"标准层选子任务仅首层关联")。"""
        if floor is None:
            sub_name = base_name                    # 不带层号：直接用基名
        else:
            sub_name = f"{base_name}{_floor_label(floor)}"
        pres = []
        # 本子任务族上一层（同父任务、同 group、floor-1 的子任务）；无层号时不找上一层
        prev_task = None
        if floor is not None:
            for c in self.tasks:
                if (c.get("group") == group and c.get("floor") == floor - 1
                        and c.get("parent") == parent.get("name")):
                    prev_task = c
                    break
        if prev_task:
            pres.append(prev_task["name"])

        # 处理用户选中的前置列表
        for pname in (pre_list or []):
            # 判断该前置是父任务还是子任务
            pre_task = None
            for c in self.tasks:
                if c.get("name") == pname:
                    pre_task = c
                    break
            if pre_task is None:
                warnings.append(f"未找到前置任务「{pname}」，已跳过。")
                continue
            if pre_task.get("level", 0) == 1:
                # 前置是子任务：只关联该子任务（不做层映射）
                # 标准层选子任务：仅首层关联该子任务，其余层仅关联上一层(prev已处理)
                if is_first:
                    pres.append(pname)
                # 非首层时，该子任务前置跳过（只用上一层）
            else:
                # 前置是父任务：沿用"同层/整任务"映射（仅父任务前置且本任务有层号时）
                self._map_parent_pre(pre_task, floor, sub_name, pres, warnings)

        t = {
            "name": sub_name,
            "start": date.today().isoformat(),
            "days": days,
            "pre": pres,
            "lag": lag,
            "color": COLORS[self.cur_color_idx % len(COLORS)],
            "_id": _new_sub_id(parent.get("_id"), self.tasks),
            "level": 1,
            "parent": parent.get("name"),
            "floor": floor,
            "group": group,
        }

        # 手动日期优先：直接采用
        if manual:
            t["start"] = manual["start"]
            t["days"] = manual["days"]
        else:
            # 自动绑定开始日期 = 最晚前置结束 + lag
            if pres:
                last_end = None
                for c in self.tasks:
                    if c.get("name") in pres:
                        end = self._effective_end(c)
                        if end and (last_end is None or end > last_end):
                            last_end = end
                if last_end is not None:
                    t["start"] = (last_end + timedelta(days=lag + 1)).isoformat()
        self.cur_color_idx += 1
        self.tasks.append(t)

    def _map_parent_pre(self, pre_parent, floor, sub_name, pres, warnings):
        """父任务前置的层映射：已分层取同层，未分层取整任务，超层号提示。
        本任务无层号(floor=None)时，直接取前置整任务结束。"""
        pname = pre_parent.get("name")
        if floor is None:
            warnings.append(f"「{sub_name}」无层号，已将该主前置「{pname}」整任务结束作为前置。")
            pres.append(pname)
            return
        sub_children = [x for x in self.tasks if x.get("parent") == pname and x.get("level", 0) == 1]
        if sub_children:
            max_floor = max((x.get("floor", 0) for x in sub_children), default=None)
            same_floor = [x for x in sub_children if x.get("floor") == floor]
            if same_floor:
                pres.append(same_floor[0]["name"])
            else:
                if max_floor is not None and floor > max_floor:
                    warnings.append(f"「{sub_name}」层号 {floor} 超过主前置「{pname}」的最大层号 {max_floor}，"
                                    f"已将该主前置整任务结束作为前置。")
                else:
                    warnings.append(f"主前置「{pname}」无对应 {floor} 层，已将该主前置整任务结束作为前置。")
                if pre_parent:
                    pres.append(pre_parent["name"])
        else:
            warnings.append(f"主前置「{pname}」未分层，已将其整任务结束作为前置。")
            pres.append(pre_parent["name"])

    def delete_task(self):
        idx = self._selected_index()
        if idx is None:
            return
        self._push_undo()
        t = self.tasks[idx]
        removed_name = t.get("name")
        # 若删除父任务，连带删除其所有子任务；记录所有被删任务名
        removed_names = {removed_name}
        if t.get("level", 0) == 0:
            for x in self.tasks:
                if x.get("parent") == removed_name:
                    removed_names.add(x.get("name"))
            self.tasks = [x for x in self.tasks if x.get("parent") != removed_name and x is not t]
        else:
            del self.tasks[idx]
        # 清理其它任务 pre 里指向任何被删任务的引用（防止悬空前置）
        for x in self.tasks:
            if removed_names & set(x.get("pre", [])):
                x["pre"] = [p for p in x["pre"] if p not in removed_names]
        # 折叠状态里的旧名也清理
        self.collapsed -= removed_names
        self.refresh_all()

    def _selected_index(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在左侧选中一行，再执行操作")
            return None
        # 用 Treeview 的 iid(即任务的 _id) 反查数据列表索引，避免父/子任务层级导致索引错位
        iid = sel[0]
        for i, t in enumerate(self.tasks):
            if t.get("_id") == iid:
                return i
        # 若无 _id 匹配，回退到名称匹配（兜底）
        vals = self.tree.item(iid, "values")
        name = vals[0] if vals else ""
        for prefix in ("▸ ", "▾ ", "└ "):
            if name.startswith(prefix):
                name = name[len(prefix):]
                break
        for i, t in enumerate(self.tasks):
            if t.get("name") == name:
                return i
        return None

    # ---------- 绘制 ----------
    def _children_of(self, parent_name):
        return [x for x in self.tasks if x.get("parent") == parent_name and x.get("level", 0) == 1]

    def _effective_pre(self, t):
        """任务【有效搭接】：父任务动态继承其【最早开始子任务】的前置；
        子任务/无子任务父任务返回自身 pre。
        规则：最早子任务没有前置时→父任务视为无搭接（不继承其他子任务的前置）。"""
        if t.get("level", 0) == 0:
            children = self._children_of(t.get("name"))
            if children:
                # 找最早开始的子任务
                dated = [c for c in children if _parse_date(c.get("start", ""))]
                if not dated:
                    return []
                first = min(dated, key=lambda c: (c.get("start", ""), c.get("name", "")))
                return list(first.get("pre", []))
        return list(t.get("pre", []))

    def _effective_lag(self, t):
        """任务【有效延迟】：与 _effective_pre 配套，父任务取最早子任务的 lag。"""
        if t.get("level", 0) == 0:
            children = self._children_of(t.get("name"))
            if children:
                dated = [c for c in children if _parse_date(c.get("start", ""))]
                if not dated:
                    return 0
                first = min(dated, key=lambda c: (c.get("start", ""), c.get("name", "")))
                return int(first.get("lag", 0) or 0)
        return int(t.get("lag", 0) or 0)

    def _sub_with_links(self, parent):
        """返回父任务下【需要单独表现的对外搭接子任务】（用于网络图表现）。

        判定口径（关键）：先把子任务的对外前置【按图上可见节点归并】，
        再和父任务自身的搭接对比——若已被父任务那条线覆盖（是其子集），
        就不再单独画、也不计数。

        为什么必须归并后比较：
          折叠视图里子任务会归并到父节点。例如「室内抹灰2F~18F」各自
          搭接「窗框安装2F~18F」，归并后全都变成「窗框安装」，
          和父任务「室内抹灰」的搭接（也是「窗框安装」）完全重复 ——
          不剔除就会在同一对节点间叠出十几条平行线。

        同时排除最早子任务（它的搭接已由父任务节点继承体现）。
        """
        children = self._children_of(parent.get("name"))
        if not children:
            return []

        sibling = {c.get("name") for c in children}
        index = {t.get("name"): t for t in self.tasks}

        def _root(n):
            """归并到图上可见节点：子任务 → 其父任务名"""
            t = index.get(n)
            if t is not None and t.get("level", 0) == 1:
                return t.get("parent") or n
            return n

        def _external(pre_list):
            """对外前置（归并后），排除同父下的兄弟"""
            return {_root(p) for p in pre_list if p not in sibling}

        dated = [c for c in children if _parse_date(c.get("start", ""))]
        if not dated:
            # 没有日期的场景：以"空搭接"为基准，凡有对外搭接的都保留
            return [c for c in children if c.get("pre") and _external(c.get("pre", []))]

        first = min(dated, key=lambda c: (c.get("start", ""), c.get("name", "")))
        base = _external(list(first.get("pre", [])))      # 父任务在网络图上的对外搭接

        out = []
        for c in children:
            if c is first or not c.get("pre"):
                continue
            ext = _external(c.get("pre", []))
            if ext and not ext.issubset(base):
                out.append(c)
        return out

    def _parent_summary(self, parent):
        """父任务有子任务时的汇总：开始=最早子开始，结束=最晚子结束，工期=子工期之和。
        若无子任务返回 None（按独立任务显示）。"""
        children = self._children_of(parent.get("name"))
        if not children:
            return None
        starts = [_parse_date(c.get("start", "")) for c in children]
        starts = [s for s in starts if s]
        ends = [_end_of_task(c) for c in children]
        ends = [e for e in ends if e]
        if not starts or not ends:
            return None
        return {
            "start": min(starts).isoformat(),
            "end": max(ends).isoformat(),
            "days": sum(int(c.get("days", 0)) for c in children),
        }

    def _display_task(self, t):
        """返回界面展示用的任务（父任务若有子任务则用汇总值覆盖）"""
        if t.get("level", 0) == 0:
            s = self._parent_summary(t)
            if s:
                # 返回一个副本，展示汇总值
                disp = dict(t)
                disp["start"] = s["start"]
                disp["days"] = s["days"]
                disp["_summary_end"] = s["end"]
                return disp
        return t

    def _effective_end(self, t):
        """计算任务用于搭接/关键路径的【实际结束日期】。
        父任务有子任务时用汇总结束(最晚子结束)；否则用自身 start+days。"""
        if t.get("level", 0) == 0:
            s = self._parent_summary(t)
            if s:
                e = _parse_date(s["end"])
                if e:
                    return e
        return _end_of_task(t)

    def _toggle_collapse(self, parent_name):
        if parent_name in self.collapsed:
            self.collapsed.discard(parent_name)
        else:
            self.collapsed.add(parent_name)
        self.refresh_list()
        self.draw_gantt()   # 甘特图随折叠状态同步刷新

    def _check_data_health(self, source_label="数据"):
        """数据体检：检测悬空前置/孤儿子任务/重名，发现问题在状态栏提示并弹窗汇总。
        只提示不修改数据。返回问题列表。"""
        names = [t.get("name") for t in self.tasks]
        nameset = set(names)
        problems = []
        # 重名
        from collections import Counter as _Counter
        dup = [k for k, v in _Counter(names).items() if v > 1]
        if dup:
            problems.append(f"重复任务名 {len(dup)} 个：{'、'.join(dup[:3])}…")
        # 悬空前置
        dangling = {}
        for t in self.tasks:
            bad = [p for p in t.get("pre", []) if p not in nameset]
            if bad:
                dangling[t.get("name")] = bad
        if dangling:
            first = next(iter(dangling.items()))
            problems.append(
                f"悬空前置 {len(dangling)} 处（前置任务已不存在，搭接失效）：\n"
                + "\n".join(f"  · 「{k}」的前置「{'、'.join(v)}」不存在" for k, v in list(dangling.items())[:5])
                + ("\n  …" if len(dangling) > 5 else ""))
        # 孤儿子任务
        parent_names = {t.get("name") for t in self.tasks if t.get("level", 0) == 0}
        orphan = [t.get("name") for t in self.tasks
                  if t.get("level") == 1 and t.get("parent") not in parent_names]
        if orphan:
            problems.append(f"孤儿子任务 {len(orphan)} 个（父任务不存在）：{'、'.join(orphan[:3])}…")
        if problems:
            try:
                self.status.set(f"⚠ {source_label}体检发现 {len(problems)} 类问题，点击查看详情")
                messagebox.showwarning(
                    f"{source_label}体检",
                    "发现以下数据问题（不影响继续使用，但搭接可能失效）：\n\n"
                    + "\n\n".join(problems)
                    + "\n\n提示：可用『编辑任务』把悬空前置改成正确的任务名。",
                    parent=self.root)
            except Exception:
                pass
        return problems

    def _ensure_unique_ids(self):
        """清洗所有任务 _id 为层级规则：父任务=纯数字递增(1,2,3...)，子任务=父id@序号(1@1,1@2...)。"""
        tasks = self.tasks
        # 父任务：按当前顺序分配唯一纯数字 id（保持已有合法数字 id）
        parents = [t for t in tasks if t.get("level", 0) == 0]
        used_pids = set()
        for t in parents:
            pid = t.get("_id", "")
            if pid.isdigit() and pid not in used_pids:
                used_pids.add(pid)
        # 给父任务稳定分配：保留合法唯一数字 id，其余按 1,2,3... 补齐(跳过已用的)
        next_num = 1
        used = set()
        for t in parents:
            pid = t.get("_id", "")
            if pid.isdigit() and pid not in used:
                t["_id"] = pid
            else:
                while str(next_num) in used:
                    next_num += 1
                t["_id"] = str(next_num)
            used.add(t["_id"])
        # 父id -> 名字 映射
        pid_by_name = {t.get("name"): t.get("_id") for t in parents}
        # 子任务：id = 父id@序号
        sub_counter = {}
        for t in tasks:
            if t.get("level", 0) == 1:
                pid = pid_by_name.get(t.get("parent"), "0")
                sub_counter[pid] = sub_counter.get(pid, 0) + 1
                t["_id"] = "%s@%d" % (pid, sub_counter[pid])

    def _column_sort_val(self, task, col):
        """取任务在指定列的排序值。"""
        if col == "name":
            return task.get("name", "")
        if col == "start":
            return task.get("start", "")
        if col == "end":
            e = _effective_end(task, self.tasks)
            return e.isoformat() if e else ""
        if col == "days":
            return int(task.get("days", 0))
        if col == "pre":
            return "、".join(self._effective_pre(task))
        if col == "type":
            return self._row_values(task)[5]
        return ""

    def _sorted_tasks(self, by_section=False):
        """按当前排序状态生成显示顺序：先排父任务，再在各父任务内部排子任务。

        by_section=True 时（甘特图用），父任务先按【分部工程】聚类 ——
        同一个分部的任务相邻，分部之间按分部表顺序排列，
        这样甘特图上同一分部的工作也聚在一起，跟着用户的分部调整走。
        """
        col = self._sort_state.get("col")
        parents = [t for t in self.tasks if t.get("level", 0) == 0]
        children = [t for t in self.tasks if t.get("level", 0) == 1]

        def cluster(items):
            """按分部把父任务分组，返回展平后的顺序（仅 by_section 时生效）"""
            if not by_section:
                return items
            idx = {n: i for i, n in enumerate(self._section_names())}
            un = len(idx) + 1                       # 未分类排最后
            return sorted(items, key=lambda t: idx.get(self._section_of(t), un))

        if not col:
            # 默认顺序：父任务保持原序（甘特图下先按分部聚类），
            # 各父任务下的子任务按开始时间升序排
            result = []
            for p in cluster(parents):
                result.append(p)
                subs = [c for c in children if c.get("parent") == p.get("name")]
                subs.sort(key=lambda t: t.get("start", ""))
                result.extend(subs)
            orphan = [c for c in children if not any(c.get("parent") == p.get("name") for p in parents)]
            result.extend(orphan)
            return result
        asc = self._sort_state.get("asc", True)
        parents.sort(key=lambda t: self._column_sort_val(t, col), reverse=not asc)
        parents = cluster(parents)
        # 保持父任务顺序，父任务下的子任务内部排序
        result = []
        for p in parents:
            result.append(p)
            subs = [c for c in children if c.get("parent") == p.get("name")]
            subs.sort(key=lambda t: self._column_sort_val(t, col), reverse=not asc)
            result.extend(subs)
        # 游离子任务(父已删/无父)追加到末尾
        orphan = [c for c in children if not any(c.get("parent") == p.get("name") for p in parents)]
        result.extend(orphan)
        return result

    def _on_column_sort(self, col):
        """点击表头列排序：升序 -> 降序 -> 恢复原列表 循环。"""
        cur = self._sort_state.get("col")
        if cur != col:
            self._sort_state = {"col": col, "asc": True}
        else:
            if self._sort_state["asc"]:
                self._sort_state["asc"] = False
            else:
                self._sort_state = {"col": None, "asc": True}  # 恢复原列表
        self.refresh_list()
        self.draw_gantt()   # 甘特图随列表排序同步刷新，保持两边行序一致

    def _collapse_all_parents(self):
        """把所有有子任务的父任务加入 collapsed(默认折叠)。"""
        for t in self.tasks:
            if t.get("level", 0) == 0 and self._children_of(t.get("name")):
                self.collapsed.add(t.get("name"))

    # ---------- 分部工程（自定义"类"层级，比任务高一层） ----------
    UNSORTED_SEC = "未分类"
    SEC_IID_PREFIX = "_sec_::"

    def _section_names(self):
        """用户已创建的分部名（按创建顺序）"""
        return [s.get("name") for s in self.sections
                if isinstance(s, dict) and s.get("name")]

    def _section_of(self, t):
        """任务所属分部名；空或无效 -> 未分类"""
        sec = (t.get("section") or "").strip()
        return sec if sec else self.UNSORTED_SEC

    def _is_section_row(self, iid):
        return bool(iid) and iid.startswith(self.SEC_IID_PREFIX)

    def _section_of_row(self, iid):
        return iid[len(self.SEC_IID_PREFIX):] if self._is_section_row(iid) else None

    def refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        sorted_order = self._sorted_tasks()

        # 每个分部下的（父）任务数
        counts = {}
        for t in sorted_order:
            if t.get("level", 0) == 0:
                sec = self._section_of(t)
                counts[sec] = counts.get(sec, 0) + 1

        # 分部顺序：用户创建的顺序；容错补上"任务里有、分部表里没有"的；未分类固定末尾
        names = list(self._section_names())
        for sec in counts:
            if sec != self.UNSORTED_SEC and sec not in names:
                names.append(sec)
        names.append(self.UNSORTED_SEC)

        sec_iid = {}
        for sec in names:
            n = counts.get(sec, 0)
            # 「未分类」为空时不显示（免得留个空壳占地方）
            if sec == self.UNSORTED_SEC and n == 0:
                continue
            coll = sec in self.section_collapsed
            arrow = "▸" if coll else "▾"
            iid = self.SEC_IID_PREFIX + sec
            self.tree.insert("", "end", iid=iid,
                             values=(IND_SECTION + "%s %s  (%d 项)" % (arrow, sec, n),
                                     "", "", "", "", "分部"),
                             open=not coll)
            sec_iid[sec] = iid

        # 父任务挂到所属分部；子任务挂到父任务（孤儿子任务挂到分部）
        parent_id_map = {}
        for t in sorted_order:
            if t.get("level", 0) == 0:
                disp = self._display_task(t)
                has_children = bool(self._children_of(t.get("name")))
                arrow = "▾" if (has_children and t.get("name") not in self.collapsed) else "▸"
                open_state = t.get("name") not in self.collapsed
                vals = list(self._row_values(disp))
                # 缩进：分部(1格) → 主任务(2格) → 子任务(3格)
                vals[0] = IND_MAIN + arrow + " " + vals[0]
                iid = t.get("_id") or _new_parent_id(self.tasks)
                pid = sec_iid.get(self._section_of(t), "")
                tid = self.tree.insert(pid, "end", iid=iid, values=vals, open=open_state)
                parent_id_map[t.get("name")] = tid
        for t in sorted_order:
            if t.get("level", 0) == 1:
                pid = parent_id_map.get(t.get("parent"))
                if pid is None:
                    pid = sec_iid.get(self._section_of(t), "")
                # 子任务：再深一级缩进 + └ 前缀表示包含关系
                vals = list(self._row_values(t))
                vals[0] = IND_SUB + "└ " + vals[0]
                iid = t.get("_id") or _new_sub_id("0", self.tasks)
                self.tree.insert(pid, "end", iid=iid, values=vals, open=True)

        # 按fold状态同步：折叠的分部不收子项（Treeview 的 open=False 已隐藏）
        # 供点击/拖拽判定使用
        self._sec_iid_map = sec_iid
        # 重新施加搜索高亮（refresh_list 会重建所有行，tag 需重置）
        self._apply_search_highlight()

    def _apply_search_highlight(self):
        """标黄：主目标（定位到的那一个）亮黄；其它含关键字的行淡黄"""
        try:
            self.tree.tag_remove("searchhit", "1.0", "end")
            self.tree.tag_remove("searchdim", "1.0", "end")
        except Exception:
            pass
        kw = (self.search_var.get() if hasattr(self, "search_var") else "").strip()
        if not kw:
            return
        low = kw.lower()
        main = getattr(self, "_search_main", None)      # 主目标行 iid
        for iid in self.tree.get_children(""):
            self._highlight_subtree(iid, low, main)

    def _highlight_subtree(self, iid, low, main):
        """递归标黄：命中主目标 -> searchhit（亮），其余命中 -> searchdim（淡）"""
        try:
            vals = self.tree.item(iid, "values")
        except Exception:
            return
        if vals:
            name = vals[0]
            for pre in ("▸ ", "▾ ", "└ "):
                if name.startswith(pre):
                    name = name[len(pre):]
                    break
            name = name.split("  (")[0]          # 分发行形如 "▾ 装饰装修工程  (16 项)"
            if low in name.lower():
                tag = "searchhit" if iid == main else "searchdim"
                self.tree.item(iid, tags=tuple(set(self.tree.item(iid, "tags"))
                                               | {tag}))
        for c in self.tree.get_children(iid):
            self._highlight_subtree(c, low, main)

    def _on_double_click(self, event):
        """双击任意列：编辑/查看任务属性。展开子任务改用单击箭头或右键菜单。"""
        self.edit_task()

    def _on_left_click(self, event):
        """单击：名称列左端箭头区 -> 切换展开/收起（分部行同理）；
        点到任务行 -> 选中并让甘特图居中。同时记录拖拽起点。"""
        region = self.tree.identify("region", event.x, event.y)
        col = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        # 记录拖拽起点（供 B1-Motion / ButtonRelease 判断是否发生拖拽）
        self._drag_from = row if (row and region == "cell") else None
        self._drag_moved = False
        if not row or region != "cell":
            return
        # 分部行：点名称列左端箭头区 -> 折叠/展开该分部
        if self._is_section_row(row):
            if col == "#1":
                bbox = self.tree.bbox(row, col)
                if bbox and 0 <= event.x - bbox[0] <= 22:
                    self._toggle_section(self._section_of_row(row))
                    self._drag_from = None
                    return "break"
            return None
        # 若点在名称列左端箭头区(前22px)且有子任务 -> 切换展开/收起
        if col == "#1":
            bbox = self.tree.bbox(row, col)
            if bbox:
                cell_x = event.x - bbox[0]
                if 0 <= cell_x <= 22:
                    raw = self.tree.item(row, "values")[0]
                    name = raw
                    for prefix in ("▸ ", "▾ ", "└ "):
                        if name.startswith(prefix):
                            name = name[len(prefix):]
                            break
                    for parent in self.tasks:
                        if parent.get("level", 0) == 0 and parent.get("name") == name:
                            if self._children_of(name):
                                self._toggle_collapse(name)
                                return "break"
        # 任意列点到任务 -> 居中显示在甘特图
        self._center_gantt_on_task(row)
        return None

    # ---------- 分部：折叠 / 拖拽归入 / 增删改 ----------
    def _toggle_section(self, sec):
        """折叠/展开某个分部"""
        if not sec:
            return
        if sec in self.section_collapsed:
            self.section_collapsed.discard(sec)
        else:
            self.section_collapsed.add(sec)
        self.refresh_list()

    def _task_by_iid(self, iid):
        """按 Treeview 行 iid(=任务 _id) 反查任务"""
        for t in self.tasks:
            if t.get("_id") == iid:
                return t
        return None

    def new_section(self):
        """新建分部工程"""
        name = simpledialog.askstring("新建分部工程",
                                      "分部名称（如：地基与基础工程）：",
                                      parent=self.root)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name == self.UNSORTED_SEC:
            messagebox.showinfo("提示", "「%s」是系统保留名称。" % self.UNSORTED_SEC, parent=self.root)
            return
        if name in self._section_names():
            messagebox.showinfo("提示", "分部「%s」已存在。" % name, parent=self.root)
            return
        self._push_undo()
        self.sections.append({"name": name})
        self.refresh_all()
        self.status.set("已新建分部：%s（把任务拖进去即可）" % name)

    def rename_section(self, sec):
        """重命名分部，并同步更新任务的所属分部"""
        if not sec or sec == self.UNSORTED_SEC:
            messagebox.showinfo("提示", "「未分类」不能重命名。", parent=self.root)
            return
        new = simpledialog.askstring("重命名分部", "新的分部名称：",
                                     initialvalue=sec, parent=self.root)
        if not new:
            return
        new = new.strip()
        if not new or new == sec:
            return
        if new == self.UNSORTED_SEC or new in self._section_names():
            messagebox.showinfo("提示", "名称「%s」已被占用。" % new, parent=self.root)
            return
        self._push_undo()
        for s in self.sections:
            if s.get("name") == sec:
                s["name"] = new
        for t in self.tasks:                     # 同步任务引用
            if (t.get("section") or "") == sec:
                t["section"] = new
        if sec in self.section_collapsed:         # 同步折叠状态
            self.section_collapsed.discard(sec)
            self.section_collapsed.add(new)
        self.refresh_all()
        self.status.set("分部已重命名为：%s" % new)

    def delete_section(self, sec):
        """删除分部：里面的任务回到「未分类」（任务本身不删）"""
        if not sec or sec == self.UNSORTED_SEC:
            messagebox.showinfo("提示", "「未分类」不能删除。", parent=self.root)
            return
        n = sum(1 for t in self.tasks
                if (t.get("section") or "") == sec and t.get("level", 0) == 0)
        if not messagebox.askyesno(
                "确认删除",
                "删除分部「%s」？\n里面 %d 个任务会回到「未分类」（任务本身不删）。" % (sec, n),
                parent=self.root):
            return
        self._push_undo()
        self.sections = [s for s in self.sections if s.get("name") != sec]
        for t in self.tasks:
            if (t.get("section") or "") == sec:
                t["section"] = ""
        self.section_collapsed.discard(sec)
        self.refresh_all()
        self.status.set("已删除分部：%s（任务已回到未分类）" % sec)

    def _assign_to_section(self, iid, sec):
        """把某行归入分部；拖的是子任务时改它父任务的分部（子任务跟随父任务）"""
        if not iid or self._is_section_row(iid):
            return False
        t = self._task_by_iid(iid)
        if t is None:
            return False
        if t.get("level", 0) == 1:
            parent_name = t.get("parent")
            t = next((p for p in self.tasks
                      if p.get("level", 0) == 0 and p.get("name") == parent_name), None)
            if t is None:
                return False
        real = "" if sec == self.UNSORTED_SEC else sec
        if (t.get("section") or "") == real:
            return False
        self._push_undo()
        t["section"] = real
        self.refresh_all()
        self.status.set("已把「%s」归入：%s" % (t.get("name"), sec))
        return True

    def _on_tree_drag(self, event):
        """拖拽中：高亮潜在落点（分部行）"""
        if not self._drag_from:
            return
        self._drag_moved = True
        try:
            self.tree.config(cursor="hand2")
        except Exception:
            pass
        for iid in self.tree.tag_has("droptarget"):
            self.tree.item(iid, tags=())
        target = self.tree.identify_row(event.y)
        if target and self._is_section_row(target):
            self.tree.item(target, tags=("droptarget",))

    def _on_tree_drop(self, event):
        """松开鼠标：落在分部行上 -> 归入该分部"""
        try:
            self.tree.config(cursor="")
        except Exception:
            pass
        for iid in self.tree.tag_has("droptarget"):
            self.tree.item(iid, tags=())
        src, moved = self._drag_from, self._drag_moved
        self._drag_from, self._drag_moved = None, False
        if not moved or not src:
            return
        target = self.tree.identify_row(event.y)
        if not target or not self._is_section_row(target):
            return
        self._assign_to_section(src, self._section_of_row(target))

    # ---------- 搜索：匹配下拉 / 展开定位 / 高亮 ----------
    SEARCH_TAG = "searchhit"

    def _search_matches(self, kw):
        """按关键字匹配任务名，返回 [(显示文本, 任务名)]，父任务优先、按原序"""
        kw = (kw or "").strip()
        if not kw:
            return []
        low = kw.lower()
        out = []
        for lv in (0, 1):                       # 先父任务、后子任务
            for t in self._sorted_tasks():
                if t.get("level", 0) != lv:
                    continue
                n = t.get("name", "")
                if low in n.lower():
                    sec = self._section_of(t) if lv == 0 else ""
                    label = n if lv == 0 else ("    └ " + n)
                    out.append((label, n))
        return out

    def _on_search_typed(self, event=None):
        """输入时实时刷新匹配下拉"""
        if event is not None and getattr(event, "keysym", "") in (
                "Up", "Down", "Return", "Escape", "Tab"):
            return
        kw = self.search_var.get().strip()
        if not kw:
            self._hide_suggest()
            return
        self._show_suggest(self._search_matches(kw))

    def _hide_suggest(self):
        if self.suggest_win is not None:
            try:
                self.suggest_win.destroy()
            except Exception:
                pass
        self.suggest_win = None
        self.suggest_box = None

    def _show_suggest(self, items):
        """在搜索框正下方弹出匹配列表"""
        self._hide_suggest()
        if not items:
            return
        e = self.search_entry
        x = e.winfo_rootx()
        y = e.winfo_rooty() + e.winfo_height() + 2
        w = max(e.winfo_width() + 150, 240)
        h = min(len(items), 12) * 20 + 4

        win = tk.Toplevel(self.root)
        win.wm_overrideredirect(True)          # 无标题栏的浮层
        win.wm_geometry("%dx%d+%d+%d" % (w, h, x, y))
        win.attributes("-topmost", True)
        box = tk.Listbox(win, activestyle="dotbox", font=("", 10),
                         selectbackground="#D6E4FF", selectforeground="#111111")
        box.pack(fill="both", expand=True)
        for label, _ in items:
            box.insert("end", label)
        box.selection_set(0)

        # 单击 -> 填入搜索框；双击 -> 直接搜索定位
        box.bind("<Button-1>", lambda ev: self._suggest_click(ev, items))
        box.bind("<Double-Button-1>", lambda ev: self._suggest_dblclick(ev, items))
        box.bind("<Return>", lambda ev: self._suggest_dblclick(ev, items))
        self.suggest_win = win
        self.suggest_box = box

    def _suggest_index(self, event):
        if self.suggest_box is None:
            return None
        i = self.suggest_box.nearest(event.y)
        if i < 0 or i >= self.suggest_box.size():
            return None
        return i

    def _suggest_click(self, event, items):
        """单击匹配项 -> 把内容填进搜索框（不立即定位）"""
        i = self._suggest_index(event)
        if i is None:
            return
        self.search_var.set(items[i][1])
        self.search_entry.icursor("end")
        return "break"

    def _suggest_dblclick(self, event, items):
        """双击匹配项 -> 直接搜索并定位到它"""
        i = self._suggest_index(event)
        if i is None:
            return "break"
        self.search_var.set(items[i][1])
        self.do_search()
        return "break"

    def do_search(self):
        """执行搜索：展开到位 + 标黄 + 滚动到尽量居中"""
        kw = self.search_var.get().strip()
        if not kw:
            return
        self._hide_suggest()
        low = kw.lower()
        hit = None
        # 优先精确命中，其次包含（父任务优先）
        for lv in (0, 1):
            for t in self._sorted_tasks():
                if t.get("level", 0) == lv and t.get("name", "").lower() == low:
                    hit = t
                    break
            if hit:
                break
        if hit is None:
            for lv in (0, 1):
                for t in self._sorted_tasks():
                    if t.get("level", 0) == lv and low in t.get("name", "").lower():
                        hit = t
                        break
                if hit:
                    break
        if hit is None:
            self.status.set("未找到匹配「%s」的任务" % kw)
            return

        # ① 记录主目标（亮黄），必须在 refresh_list 之前 —— 否则高亮时还不知主目标是谁
        self._search_main = hit.get("_id")

        # ② 逐级展开：父任务展开 + 其所属分部展开
        if hit.get("level", 0) == 1:
            self.collapsed.discard(hit.get("parent"))
        else:
            self.collapsed.discard(hit.get("name"))
        self.section_collapsed.discard(self._section_of(
            next((p for p in self.tasks
                  if p.get("level", 0) == 0 and p.get("name") == hit.get("parent")), hit)))
        self.refresh_list()

        # ③ 定位到该行：展开树节点 + 滚动
        iid = hit.get("_id")
        if not iid or not self.tree.exists(iid):
            self.status.set("未找到「%s」所在行" % kw)
            return
        # 展开从根到该行的所有祖先（Treeview 层）
        parent = self.tree.parent(iid)
        chain = []
        while parent:
            chain.append(parent)
            parent = self.tree.parent(parent)
        for p in chain:
            self.tree.item(p, open=True)
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self._scroll_row_into_view(iid)
        # ③ 甘特图同步：重绘（含命中高亮）并滚动到该任务（垂直居中 + 水平对齐开始时间）
        try:
            self.draw_gantt()
            self.root.update_idletasks()
            self._center_gantt_on_task(iid)
        except Exception:
            pass
        self.status.set("已定位到：%s" % hit.get("name"))

    def _scroll_row_into_view(self, iid):
        """把某行滚动到可视区 —— 保证可见优先，尽量居中。"""
        vis = self._visible_tree_items()
        if iid not in vis:
            return
        idx = vis.index(iid)
        total = len(vis)

        # 用实际行高估算一屏能显示几行
        row_h = 24
        try:
            bb = self.tree.bbox(iid)
            if bb and bb[3] > 0:
                row_h = bb[3]
            else:
                # 该行尚未渲染时，取任一已渲染行的高度
                for v in vis:
                    b2 = self.tree.bbox(v)
                    if b2 and b2[3] > 0:
                        row_h = b2[3]
                        break
        except Exception:
            pass
        n_show = max(int(self.tree.winfo_height() / max(row_h, 1)), 1)

        # 目标：命中行靠近中间；上方不够则贴顶，下方不够则贴底
        want = idx - n_show // 2
        want = max(0, min(want, max(0, total - n_show)))
        # 先让 Tk 滚到该行，再按我们的目标位置精调
        self.tree.see(iid)
        try:
            self.tree.update_idletasks()
            self.tree.yview_moveto(max(0.0, min(1.0, float(want) / max(total, 1))))
            self.tree.update_idletasks()
        except Exception:
            pass

        # 兜底：若仍未进入可视区（Treeview 的 yview 与行数不完全对应），
        # 用 see + 前后微调逼近
        for _ in range(40):
            bb = self.tree.bbox(iid)
            if bb and bb[1] >= 0 and (bb[1] + bb[3]) <= self.tree.winfo_height():
                break
            first_iid = vis[0] if vis else None
            b0 = self.tree.bbox(first_iid) if first_iid else None
            cur_top = self.tree.yview()[0]
            step = 1.0 / max(total, 1)
            if bb and bb[1] < 0:
                self.tree.yview_moveto(max(0.0, cur_top - step * 3))
            else:
                self.tree.yview_moveto(min(1.0, cur_top + step * 3))
            self.tree.update_idletasks()
        # 最后再 see 一次，确保 Tk 认为它可见
        self.tree.see(iid)

    def _visible_tree_items(self, parent=""):
        """当前树上真正可见的行（展开的才递归进去）"""
        out = []
        for iid in self.tree.get_children(parent):
            out.append(iid)
            if self.tree.item(iid, "open"):
                out.extend(self._visible_tree_items(iid))
        return out

    def clear_search(self):
        """清空搜索框与高亮"""
        self.search_var.set("")
        self._search_main = None
        self._hide_suggest()
        try:
            self.tree.tag_remove("searchhit", "1.0", "end")
            self.tree.tag_remove("searchdim", "1.0", "end")
        except Exception:
            pass
        try:
            if hasattr(self, "canvas"):
                self.canvas.delete("searchhit")
        except Exception:
            pass
        self.status.set("已清除搜索")

    # ---------- 画布滚轮：滚动 + Ctrl 缩放 ----------
    ZOOM_MIN, ZOOM_MAX = 0.25, 4.0
    ZOOM_STEP = 1.15

    def _bind_canvas_wheel(self, canvas, which):
        """给画布绑滚轮事件。
        Windows/macOS → <MouseWheel>；X11 → <Button-4>/<Button-5>"""
        def on_wheel(event):
            # delta>0 = 向上滚
            delta = getattr(event, "delta", 0)
            if delta == 0:
                delta = 120 if getattr(event, "num", 0) == 4 else -120
            ctrl = bool(event.state & 0x0004)      # Ctrl 键
            if ctrl:
                self._zoom_canvas(canvas, which, delta, event.x, event.y)
            else:
                self._scroll_canvas(canvas, delta)
            return "break"
        canvas.bind("<MouseWheel>", on_wheel)
        canvas.bind("<Button-4>", on_wheel)
        canvas.bind("<Button-5>", on_wheel)

    def _scroll_canvas(self, canvas, delta):
        """普通滚轮：垂直滚动（Shift 按住时横向滚动由系统处理，这里只做纵向）"""
        canvas.yview_scroll(-1 if delta > 0 else 1, "units")

    def _zoom_canvas(self, canvas, which, delta, mx, my):
        """Ctrl+滚轮：缩放。以鼠标位置为锚点，缩放后鼠标指着的那个点尽量不动。"""
        old = getattr(self, "gantt_zoom" if which == "gantt" else "net_zoom", 1.0)
        step = self.ZOOM_STEP if delta > 0 else (1.0 / self.ZOOM_STEP)
        new = max(self.ZOOM_MIN, min(self.ZOOM_MAX, old * step))
        if abs(new - old) < 1e-6:
            return
        # 记录鼠标处的"内容坐标"（0~1 归一化），缩放后仍滚到该位置
        cw = max(canvas.winfo_width(), 1)
        ch = max(canvas.winfo_height(), 1)
        fx = canvas.xview()[0] + (mx / float(cw)) * (canvas.xview()[1] - canvas.xview()[0])
        fy = canvas.yview()[0] + (my / float(ch)) * (canvas.yview()[1] - canvas.yview()[0])

        if which == "gantt":
            self.gantt_zoom = new
            self.draw_gantt()
        else:
            self.net_zoom = new
            self._draw_network()
        canvas.update_idletasks()

        # 缩放后把鼠标处那块内容拉回原位
        vx0, vx1 = canvas.xview()
        vy0, vy1 = canvas.yview()
        span_x = max(vx1 - vx0, 1e-6)
        span_y = max(vy1 - vy0, 1e-6)
        canvas.xview_moveto(max(0.0, min(1.0, fx - (mx / float(cw)) * span_x)))
        canvas.yview_moveto(max(0.0, min(1.0, fy - (my / float(ch)) * span_y)))
        pct = int(round(new * 100))
        self.status.set("缩放 %d%%（Ctrl+滚轮缩放，滚轮滚动）" % pct)

    # ---------- "今天"竖线（当前时间指示） ----------
    # ---------- 数据目录 ----------
    def open_data_folder(self):
        """打开数据所在文件夹（Windows 资源管理器）"""
        try:
            if os.name == "nt":
                os.startfile(self.data_dir)
            else:
                import subprocess as _sp
                _sp.Popen(["xdg-open", self.data_dir])
            self.status.set("数据目录：%s" % self.data_dir)
        except Exception as e:
            messagebox.showinfo(
                "数据目录",
                "你的排期数据在这里：\n\n%s\n\n(自动打开失败：%s)" % (self.data_dir, e),
                parent=self.root)

    # ---------- 菜单栏用到的辅助方法 ----------
    def _focus_search(self):
        """Ctrl+F：跳到搜索框"""
        try:
            self.search_entry.focus_set()
            self.search_entry.selection_range(0, "end")
        except Exception:
            pass

    def _zoom_by(self, direction):
        """菜单「放大/缩小」：等价于 Ctrl+滚轮一格"""
        try:
            c = self.canvas
            cx = max(c.winfo_width() // 2, 1)
            cy = max(c.winfo_height() // 2, 1)
            delta = 120 if direction > 0 else -120
            self._zoom_canvas(c, "gantt", delta, cx, cy)
        except Exception as e:
            self.status.set("缩放失败：%s" % e)

    def _zoom_reset(self):
        """菜单「恢复 100%」"""
        try:
            self.gantt_zoom = 1.0
            self.net_zoom = 1.0
            self.draw_gantt()
            self._draw_network()
            self.status.set("缩放已恢复 100%")
        except Exception as e:
            self.status.set("恢复缩放失败：%s" % e)

    def _set_all_sections(self, collapse):
        """菜单「展开/折叠全部分部」"""
        if collapse:
            self.section_collapsed = set(self._section_names())
        else:
            self.section_collapsed = set()
        self.refresh_list()
        self.status.set("已%s所有分部" % ("折叠" if collapse else "展开"))

    def _show_help(self):
        """菜单「使用说明」：弹出简要操作指引"""
        msg = """【怎么用】

加任务 → 菜单「任务 → 添加任务」，或工具栏「＋ 添加任务」
改任务 → 双击那一行
删任务 → 选中后按 Delete

【核心建议】
建任务时请用「前置任务 + 搭接时间」定义工序关系，
而不是只填开始/结束日期 —— 这样：
  · 搭接关系图才画得对
  · 改工期时后面任务会自动顺延

【常用操作】
  分部工程 → 把任务拖到分部行上即可归类
  搜索     → 工具栏搜索框，双击结果直接跳转
  显示今天 → 勾选后甘特图出现红色虚线
  缩放     → Ctrl + 滚轮

【数据在哪】
点菜单「帮助 → 打开数据文件夹」

更详细的教程见项目文档。"""
        messagebox.showinfo("使用说明", msg, parent=self.root)

    def _show_about(self):
        """菜单「关于」"""
        try:
            import gantt_tool as _self
            ver = getattr(_self, "VERSION", "?")
        except Exception:
            ver = "?"
        messagebox.showinfo(
            "关于",
            "施工排期甘特图工具  v%s\n\n"
            "给工程人的、双击就能用的施工进度排期工具。\n\n"
            "数据目录：\n%s\n\n"
            "授权：GPL-3.0（含作者附加条款）" % (ver, self.data_dir),
            parent=self.root)

    def _on_toggle_today(self):
        """勾选/取消「显示今天」"""
        self.draw_gantt()
        if self.var_today.get():
            self.status.set("已显示当前时间竖线（红色虚线）")
        else:
            self.status.set("已隐藏当前时间竖线")

    def _schedule_today_refresh(self):
        """每分钟刷新一次竖线：跨零点时自动走到新的一天"""
        def tick():
            if self.var_today.get():
                try:
                    self.draw_gantt()
                except Exception:
                    pass
            self._today_job = self.root.after(60000, tick)
        self._today_job = self.root.after(60000, tick)

    def _draw_today_line(self, d0, d1, scale, left_margin, top_margin,
                         rows, chart_w):
        """在甘特图上画一条"今天"的竖线。
        d0/d1: 图表起止日期；scale: 每天像素；rows: 可见行数。
        竖线落在今天对应的 x 位置；今天不在图内时不画。"""
        today = date.today()
        if today < d0 or today > d1:
            return
        x = left_margin + (today - d0).days * scale
        y0 = top_margin
        y1 = top_margin + rows * (self._gantt_row_h or 30)
        c = self.canvas
        # 竖线：红色虚线，够显眼又不盖住进度条（画完会降到色块下方）
        c.create_line(x, y0, x, y1, fill="#E53935", width=2,
                      dash=(6, 4), tags="todayline")
        # 顶部小三角 + 日期文字
        c.create_polygon(x - 6, y0 - 10, x + 6, y0 - 10, x, y0 - 1,
                         fill="#E53935", outline="", tags="todayline")
        label = "今天 %d/%d" % (today.month, today.day)
        c.create_text(x + 6, y0 - 18, text=label, anchor="w",
                      fill="#E53935", font=("", max(8, int(9 * self.gantt_zoom)), "bold"),
                      tags="todayline")
        # 竖线垫到进度条下方（不遮挡任务条），但顶部标记保持可见
        try:
            c.tag_lower("todayline")
        except Exception:
            pass

    def _center_gantt_on_task(self, row_iid):
        """点击列表任务后，甘特图自动滚动到该任务，让其垂直(和水平)居中显示。"""
        raw = self.tree.item(row_iid, "values")
        if not raw:
            return
        name = raw[0]
        for prefix in ("▸ ", "▾ ", "└ "):
            if name.startswith(prefix):
                name = name[len(prefix):]
                break
        visible = getattr(self, "_gantt_visible", None)
        if not visible:
            return
        for i, (orig, disp) in enumerate(visible):
            if orig.get("name") == name or disp.get("name") == name:
                total = len(visible)
                # 垂直居中：把该行放到窗口垂直中点附近
                self.canvas.yview_moveto(max(0.0, min(1.0, (i + 0.5) / max(total, 1))))
                # 水平居中：滚动到任务开始时间的 x 位置
                s = _parse_date(disp.get("start", ""))
                if s:
                    d0 = getattr(self, "_gantt_d0", None)
                    total_days = getattr(self, "_gantt_total_days", None)
                    chart_w = getattr(self, "_gantt_chart_w", None)
                    if d0 and total_days and chart_w:
                        frac = (s - d0).days / max(total_days, 1)
                        self.canvas.xview_moveto(max(0.0, min(1.0, frac - 0.15)))
                return

    def _on_gantt_press(self, event):
        """按下鼠标：记录拖动起点，切换为移动光标。"""
        self._gantt_drag = [event.x, event.y]
        try:
            self.canvas.config(cursor="fleur")
        except Exception:
            pass

    def _on_gantt_drag(self, event):
        """按住左键拖动：按"像素/画布尺寸"比例精确平移，不跳动。"""
        if self._gantt_drag[0] is None:
            self._gantt_drag = [event.x, event.y]
            return
        dx = event.x - self._gantt_drag[0]
        dy = event.y - self._gantt_drag[1]
        self._gantt_drag = [event.x, event.y]
        cw = max(self.canvas.winfo_width(), 1)
        ch = max(self.canvas.winfo_height(), 1)
        # xview/moveto 是[0,1]比例，移动像素/画布尺寸 就是应移动的比例
        xfrac = -dx / cw
        yfrac = -dy / ch
        try:
            self.canvas.xview_moveto(max(0.0, min(1.0, self.canvas.xview()[0] + xfrac)))
        except Exception:
            pass
        try:
            self.canvas.yview_moveto(max(0.0, min(1.0, self.canvas.yview()[0] + yfrac)))
        except Exception:
            pass

    def _on_gantt_release(self, event):
        """松开鼠标：恢复默认指针。"""
        self._gantt_drag = [None, None]
        try:
            self.canvas.config(cursor="")
        except Exception:
            pass

    # ---------- 搭接关系图(网络图) ----------
    def _on_tab_changed(self, event=None):
        """切换到搭接关系图页签时绘制。"""
        try:
            idx = self.notebook.index(self.notebook.select())
        except Exception:
            return
        if idx == 1:
            self._draw_network()

    def _draw_network(self):
        """绘制搭接关系图(网络图)：节点=任务，有向连线=前置→后继(标注延迟)。
        布局：按依赖深度分列；节点尺寸自适应文字；列间留出连线通道；同列垂直居中。"""
        c = self.net_canvas
        c.delete("all")
        tasks = self.tasks
        if not tasks:
            c.create_text(20, 20, anchor="nw", text="暂无任务", fill="#888888")
            c.configure(scrollregion=(0, 0, 400, 200))
            return

        by_name = {t.get("name"): t for t in tasks}
        # 只画当前可见序列（与列表一致）
        visible = [t for t in self._sorted_tasks()
                   if t.get("level", 0) == 0 or t.get("parent") not in self.collapsed]
        if not visible:
            return
        vis_names = {t.get("name") for t in visible}

        # --- 分层：按依赖深度 ---
        depth = {}

        def calc_depth(t, stack):
            n = t.get("name")
            if n in depth:
                return depth[n]
            if n in stack:
                return 0
            stack.add(n)
            d = 0
            # 用【有效搭接】(父任务=最早子任务的前置)分层，与列表显示一致
            for p in self._effective_pre(t):
                # 前置若被折叠(子任务不可见)，回溯到其所属父任务
                pn = p
                if pn not in vis_names:
                    pt = by_name.get(p)
                    if pt is not None and pt.get("level", 0) == 1:
                        pn = pt.get("parent")
                if pn in vis_names and pn in by_name and pn != n:
                    d = max(d, calc_depth(by_name[pn], stack) + 1)
            stack.discard(n)
            depth[n] = d
            return d

        for t in visible:
            calc_depth(t, set())

        # --- 节点尺寸自适应文字 ---
        try:
            import tkinter.font as _tf
            name_font = _tf.Font(family="TkDefaultFont", size=max(7, int(11 * self.net_zoom)), weight="bold")
            info_font = _tf.Font(family="TkDefaultFont", size=max(6, int(9 * self.net_zoom)))
        except Exception:
            name_font = info_font = None

        node_metrics = {}     # name -> (width, height, name, info)
        z = self.net_zoom     # 缩放系数：节点尺寸、间距都跟着变
        for t in visible:
            n = t.get("name", "")
            e = _end_of_task(t)
            info = "%s~%s  %dd" % (t.get("start", "")[5:],
                                   (e.isoformat()[5:] if e else ""),
                                   int(t.get("days", 1)))
            nw = name_font.measure(n) if name_font else len(n) * 8 * z
            iw = info_font.measure(info) if info_font else len(info) * 6 * z
            w = max(160 * z, min(400 * z, max(nw, iw) + 40 * z))   # 左右各留 20px 内边距
            node_metrics[n] = (w, max(30, 62 * z), n, info)         # 节点加高：62px

        # --- 列布局（S形/蛇形折行）：第1行从左到右、第2行从右到左，依次交替 ---
        cols = {}
        for t in visible:
            cols.setdefault(depth.get(t.get("name"), 0), []).append(t)
        col_keys = sorted(cols)
        # 列内按开始时间排序，视觉更有序
        for d in col_keys:
            cols[d].sort(key=lambda x: (x.get("start", ""), x.get("name", "")))
        col_w = {}
        for d in col_keys:
            col_w[d] = max(node_metrics[t.get("name")][0] for t in cols[d]) + 110 * z

        LEFT, TOP = 90 * z, 50 * z
        ROW_H = 130 * z                  # 行距(带内节点垂直间距)：加大，避免重叠
        BAND_GAP = 130 * z               # 带与带之间的额外间距：加大
        # 每行最多放几列：让画布接近方形
        n_cols = len(col_keys)
        if n_cols <= 8:
            COLS_PER_ROW = n_cols
        else:
            import math
            COLS_PER_ROW = max(4, int(math.ceil(math.sqrt(n_cols * 1.6))))

        # 带内列顺序：偶数带(0,2,4...)正序；奇数带(1,3,5...)倒序 → 形成 S 形回流
        band_rows_list = []      # 每带的列列表
        for bi in range(0, n_cols, COLS_PER_ROW):
            keys = list(col_keys[bi:bi + COLS_PER_ROW])
            if (bi // COLS_PER_ROW) % 2 == 1:
                keys = list(reversed(keys))
            band_rows_list.append(keys)

        col_x = {}          # 列 -> 左边缘 x
        col_band_top = {}   # 列 -> 该带顶部 y
        band_top = TOP
        for keys in band_rows_list:
            bx = LEFT
            for d in keys:
                col_x[d] = bx
                col_band_top[d] = band_top
                bx += col_w[d]
            band_rows = max(len(cols[d]) for d in keys)
            band_top += band_rows * ROW_H + BAND_GAP
        # 画布宽度：取最宽的一带
        widest = 0
        for keys in band_rows_list:
            widest = max(widest, sum(col_w[d] for d in keys))
        total_w = LEFT + widest + 80
        total_h = band_top + 40

        pos = {}     # name -> (cx, cy, w, h)
        name_band = {}   # 节点名 -> 所属带号(0起)
        band_dir = {}    # 带号 -> +1(正向,左→右) / -1(反向,右→左)
        name_no = {}     # 节点名 -> 编号(按依赖深度/列序，用于判定连线顺向/反向)
        for bi, keys in enumerate(band_rows_list):
            band_dir[bi] = 1 if bi % 2 == 0 else -1     # 偶数带正向、奇数带反向
            band_rows = max(len(cols[d]) for d in keys)
            for d in keys:
                items = cols[d]
                n_rows = len(items)
                offset = (band_rows - n_rows) / 2.0     # 带内垂直居中
                for i, t in enumerate(items):
                    n = t.get("name")
                    w, h, _, _ = node_metrics[n]
                    cx = col_x[d] + w / 2 + (col_w[d] - 110 - w) / 2
                    cy = col_band_top[d] + h / 2 + (offset + i) * ROW_H
                    pos[n] = (cx, cy, w, h)
                    name_band[n] = bi
                    name_no[n] = d          # 用依赖深度作为编号(小=靠前/左侧)

        # --- 先画连线(在下层)，再画节点 ---
        LANE = ["#2F5597", "#C00000", "#548235", "#BF8F00", "#7030A0"]
        # 统计每个节点的"出边数/入边数"及序号，用于出线/入线端口均匀分布
        out_degree = {}
        in_degree = {}
        out_index = {}     # 源 -> {目标: 第几条出边}
        in_index = {}      # 目标 -> {源: 第几条入边}
        for _t in visible:
            _tn = _t.get("name")
            _grp = (list(self._effective_pre(_t))
                    + [p for _s in self._sub_with_links(_t) for p in _s.get("pre", [])])
            for _p in _grp:
                out_degree[_p] = out_degree.get(_p, 0) + 1
                out_index.setdefault(_p, {})[_tn] = out_degree[_p] - 1
                in_degree[_tn] = in_degree.get(_tn, 0) + 1
                in_index.setdefault(_tn, {})[_p] = in_degree[_tn] - 1
        occupied = []   # 已占用矩形 [(x0,y0,x1,y1)]，用于碰撞检测(供连线标注避让)
        drawn_segs = []  # 已画线段 [(方向,固定坐标,起,止)]，用于平行重叠错开
        ARROW_GAP = 6    # 重叠时错开的半箭头距离

        def hit(x0, y0, x1, y1):
            """检测矩形是否与已占用区域相交。"""
            for (a0, b0, a1, b1) in occupied:
                if x0 < a1 and a0 < x1 and y0 < b1 and b0 < y1:
                    return True
            return False

        def offset_if_parallel(orient, fixed, lo, hi, prefer=1):
            """若与已画线段【平行且重叠】，返回应错开的距离；否则 0。
            orient: 'h'横线(固定y) / 'v'竖线(固定x)；fixed: 该线固定坐标；lo~hi: 线段区间。
            prefer: 优先朝哪边错开（+1 右/下，-1 左/上）。
                    由调用方按【这段线在色块哪一侧】传入 —— 错开朝外走，
                    才不会挤向色块、压住箭头，或与相邻线继续打架。
            错开量按 ARROW_GAP 逐级增加：prefer 方向先找最近空位，找不到再反向兜底。"""
            def conflict(f):
                for (o, f2, a, b) in drawn_segs:
                    if o != orient:
                        continue
                    # 平行(同方向) 且 固定坐标接近(<=2px视作落在同一条线上) 且 区间重叠
                    if abs(f2 - f) <= 2 and not (hi <= a or lo >= b):
                        return True
                return False

            if not conflict(fixed):
                return 0
            # ① 先沿 prefer 方向由近及远找空位
            for k in range(1, 16):
                d = prefer * k * ARROW_GAP
                if not conflict(fixed + d):
                    return d
            # ② 兜底：反方向再找
            for k in range(1, 16):
                d = -prefer * k * ARROW_GAP
                if not conflict(fixed + d):
                    return d
            return 0

        # 辅助：画一条 前置→后继 的折线
        def draw_edge(t_from, t_to, lag, lane_idx, dashed=False):
            nfrom, nto = t_from.get("name"), t_to.get("name")
            if nfrom not in pos or nto not in pos:
                return
            fc_x, fc_y, fw, fh = pos[nfrom]
            tc_x, tc_y, tw, th = pos[nto]
            col = LANE[lane_idx % len(LANE)]
            b_from = name_band.get(nfrom, 0)
            b_to = name_band.get(nto, 0)
            cross_band = (b_from != b_to)
            dir_from = band_dir.get(b_from, 1)
            no_from = name_no.get(nfrom, 0)
            no_to = name_no.get(nto, 0)
            reversed_dep = (no_to < no_from)

            # --- 连接端口选择(沿用带方向/反向依赖规则) ---
            if reversed_dep:
                x_from = fc_x - fw / 2; x_to = tc_x + tw / 2 + 8; dir_out = -1
            elif not cross_band:
                if dir_from > 0:
                    x_from, x_to = fc_x + fw / 2, tc_x - tw / 2 - 8; dir_out = +1
                else:
                    x_from, x_to = fc_x - fw / 2, tc_x + tw / 2 + 8; dir_out = -1
            else:
                if dir_from > 0:
                    x_from = fc_x + fw / 2; x_to = tc_x + tw / 2 + 8; dir_out = +1
                else:
                    x_from = fc_x - fw / 2; x_to = tc_x - tw / 2 - 8; dir_out = -1

            # --- 端口纵向定位：源/目标用节点中心(行内已错高) ---
            y_from, y_to = fc_y, tc_y
            # 源节点多条出边时轻微错开
            oc = out_degree.get(nfrom, 1)
            oi = out_index.get(nfrom, {}).get(nto, 0)
            if oc > 1:
                y_from = fc_y + (oi - (oc - 1) / 2.0) * 9
            ic = in_degree.get(nto, 1)
            ii = in_index.get(nto, {}).get(nfrom, 0)
            if ic > 1:
                y_to = tc_y + (ii - (ic - 1) / 2.0) * 9

            # --- 垂直通道 + 平行重叠错开(上一版逻辑) ---
            if not cross_band:
                chan_x = x_from + dir_out * (max(16, abs(x_to - x_from) * 0.45) + lane_idx * 10)
            else:
                chan_x = x_from + dir_out * (26 + lane_idx * 10)
            v_lo, v_hi = min(y_from, y_to), max(y_from, y_to)
            # 通道竖线：按"从色块哪一侧出去"决定错开方向 ——
            # 右侧出去就朝右错开、左侧出去就朝左错开（朝外走，不挤向色块、不压箭头）
            chan_x += offset_if_parallel("v", chan_x, v_lo, v_hi, prefer=dir_out)
            h_lo, h_hi = min(x_from, chan_x), max(x_from, chan_x)
            y_from += offset_if_parallel("h", y_from, h_lo, h_hi)
            h2_lo, h2_hi = min(chan_x, x_to), max(chan_x, x_to)
            y_to += offset_if_parallel("h", y_to, h2_lo, h2_hi)
            drawn_segs.append(("h", y_from, *sorted((x_from, chan_x))))
            drawn_segs.append(("v", chan_x, *sorted((y_from, y_to))))
            drawn_segs.append(("h", y_to, *sorted((chan_x, x_to))))

            dash = (5, 3) if dashed else None
            c.create_line(x_from, y_from, chan_x, y_from, chan_x, y_to, x_to, y_to,
                          arrow="last", fill=col, width=2, dash=dash,
                          arrowshape=(11, 13, 5), tags="edge")

            # --- 延迟标注：放在通道旁 ---
            if lag:
                label = ("+%dd" % lag) if lag > 0 else ("%dd" % lag)
                lx = chan_x + 5
                ly = (y_from + y_to) / 2
                lw, lh = 46, 18
                guard = 0
                while hit(lx - 2, ly - lh / 2, lx + lw, ly + lh / 2) and guard < 30:
                    ly += 14; guard += 1
                    if guard % 4 == 0:
                        lx += 12
                c.create_text(lx, ly, text=label, anchor="w",
                              fill=col, font=("", 9, "bold"), tags="edge")
                occupied.append((lx - 2, ly - lh / 2, lx + lw, ly + lh / 2))

        def _eff_pre(t):
            return self._effective_pre(t)

        for t in visible:
            n = t.get("name")
            if n not in pos:
                continue
            # 辅助：把前置名解析为【图上可见的节点】——
            # 若前置本身可见直接用；否则(折叠/子任务)回溯到其所属父任务。
            def resolve_node(name):
                if name in pos:
                    return name
                pt = by_name.get(name)
                # 前置是子任务但被折叠 → 归并到其父任务
                if pt is not None and pt.get("level", 0) == 1:
                    pn = pt.get("parent")
                    if pn in pos:
                        return pn
                # 兜底：同名父任务
                if name in pos:
                    return name
                return None

            # 1) 有效搭接（父任务=最早子任务的前置；子任务=自身前置）
            for idx, p in enumerate(self._effective_pre(t)):
                rn = resolve_node(p)
                if rn is None or rn == n:          # 跳过自身环
                    continue
                draw_edge(by_name.get(rn, {"name": rn}), t, self._effective_lag(t), idx)
            # 2) 折叠的父任务：其"有搭接关系"的其余子任务，用虚线+半透明块表现
            if t.get("level", 0) == 0 and n in self.collapsed:
                for sub in self._sub_with_links(t):
                    for idx2, p2 in enumerate(sub.get("pre", [])):
                        rn2 = resolve_node(p2)
                        if rn2 is None or rn2 == n:
                            continue
                        draw_edge(by_name.get(rn2, {"name": rn2}), t,
                                  int(sub.get("lag", 0) or 0),
                                  len(self._effective_pre(t)) + idx2, dashed=True)

        # --- 第一遍：画全部节点 + 记录占位(含文字实际边界) ---
        for t in visible:
            n = t.get("name")
            if n not in pos:
                continue
            cx, cy, w, h = pos[n]
            x0, y0 = cx - w / 2, cy - h / 2
            x1, y1 = cx + w / 2, cy + h / 2
            is_sub = t.get("level", 0) == 1
            color = (t.get("color") or "#4472C4")
            is_crit = n in getattr(self, "critical_tasks", set())
            c.create_rectangle(x0, y0, x1, y1,
                               fill="#F7FBFF" if is_sub else color,
                               outline="#E53935" if is_crit else "#4A5568",
                               width=3 if is_crit else 1)
            c.create_rectangle(x0, y0, x0 + 5, y1, fill=color, outline="")
            txt = "#1F2937" if is_sub else "#FFFFFF"
            tname = c.create_text(x0 + 13, cy - 12, anchor="w", text=n,
                                  fill=txt, font=("", max(7, int(11 * self.net_zoom)), "bold"))
            tinfo = c.create_text(x0 + 13, cy + 12, anchor="w", text=node_metrics[n][3],
                                  fill="#6B7280" if is_sub else "#EAF2FB", font=("", max(6, int(9 * self.net_zoom))))
            occupied.append((x0 - 2, y0 - 2, x1 + 2, y1 + 2))
            # 节点内文字的实际占位也纳入，避免色块压住文字
            for ti in (tname, tinfo):
                bb = c.bbox(ti)
                if bb:
                    occupied.append((bb[0] - 4, bb[1] - 3, bb[2] + 4, bb[3] + 3))

        # --- 第二遍：为每个"有对外搭接"的子任务画独立小色块，虚线直连子任务 ---
        # --- 第二遍：折叠的父任务用半透明块表现"其余子任务的搭接" ---
        for t in visible:
            n = t.get("name")
            if n not in pos:
                continue
            if t.get("level", 0) != 0 or n not in self.collapsed:
                continue
            subs = self._sub_with_links(t)
            if not subs:
                continue
            cx, cy, w, h = pos[n]
            x1, y1 = cx + w / 2, cy + h / 2
            color = (t.get("color") or "#4472C4")
            bw = max(120 * self.net_zoom, min(w, 190 * self.net_zoom))
            bh = max(24, 46 * self.net_zoom)
            bx = cx - bw / 2                      # 与节点同列居中
            by = y1 + 10                          # 挂在节点正下方
            # 碰撞避让：向下试位，再左右微调
            guard = 0
            while hit(bx, by, bx + bw, by + bh) and guard < 60:
                by += 12
                guard += 1
                if guard % 5 == 0:
                    bx += 14
            bx2, by2 = bx + bw, by + bh
            light = self._lighten(color, 0.78)
            c.create_rectangle(bx, by, bx2, by2,
                               fill=light, outline=color, width=1, dash=(4, 2))
            c.create_line(cx, y1, bx + bw / 2, by, arrow="last", fill=color,
                          width=1, arrowshape=(6, 8, 3), tags="edge")
            c.create_text(bx + bw / 2, by + 15, anchor="center",
                          text="子任务搭接", fill="#374151", font=("", max(6, int(8 * self.net_zoom))))
            c.create_text(bx + bw / 2, by + 33, anchor="center",
                          text="%d 项" % len(subs), fill="#4B5563",
                          font=("", max(7, int(10 * self.net_zoom)), "bold"))
            occupied.append((bx - 2, by - 2, bx2 + 2, by2 + 2))

        c.configure(scrollregion=(0, 0, max(total_w, 400), max(total_h, 200)))
        # 连线置于节点之上，确保箭头始终可见(不被节点框盖住)
        try:
            c.tag_raise("edge")
        except Exception:
            pass

    @staticmethod
    def _lighten(hex_color, ratio):
        """把颜色变浅(ratio=0保持原色, 1变白)，用于模拟半透明同色块。"""
        h = (hex_color or "#4472C4").lstrip("#")
        if len(h) != 6:
            return "#D0D8E8"
        try:
            r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        except ValueError:
            return "#D0D8E8"
        r = int(r + (255 - r) * ratio)
        g = int(g + (255 - g) * ratio)
        b = int(b + (255 - b) * ratio)
        return "#%02X%02X%02X" % (r, g, b)

    def _on_net_press(self, event):
        self._net_drag = [event.x, event.y]
        try:
            self.net_canvas.config(cursor="fleur")
        except Exception:
            pass

    def _on_net_drag(self, event):
        if self._net_drag[0] is None:
            self._net_drag = [event.x, event.y]
            return
        dx = event.x - self._net_drag[0]
        dy = event.y - self._net_drag[1]
        self._net_drag = [event.x, event.y]
        cw = max(self.net_canvas.winfo_width(), 1)
        ch = max(self.net_canvas.winfo_height(), 1)
        try:
            self.net_canvas.xview_moveto(max(0.0, min(1.0, self.net_canvas.xview()[0] - dx / cw)))
        except Exception:
            pass
        try:
            self.net_canvas.yview_moveto(max(0.0, min(1.0, self.net_canvas.yview()[0] - dy / ch)))
        except Exception:
            pass

    def _on_net_release(self, event):
        self._net_drag = [None, None]
        try:
            self.net_canvas.config(cursor="")
        except Exception:
            pass

    def _toggle_right_click(self):
        """右键菜单『展开/收起子任务』：对当前选中父任务切换。"""
        idx = self._selected_index()
        if idx is None:
            return
        parent = self.tasks[idx]
        if parent.get("level", 0) != 0:
            messagebox.showinfo("提示", "请先选中一个有子任务的父任务", parent=self.root)
            return
        if self._children_of(parent.get("name")):
            self._toggle_collapse(parent.get("name"))
        else:
            messagebox.showinfo("提示", "该任务没有子任务", parent=self.root)

    def _row_values(self, t):
        end = _end_of_task(t)
        # 若 t 带汇总结束字段（父任务汇总），优先用该值
        if t.get("_summary_end"):
            end = _parse_date(t.get("_summary_end"))
        sub = "子" if t.get("level", 0) == 1 else ""
        floor = _floor_label(t["floor"]) if t.get("floor") is not None else ""
        summary = "汇总" if (t.get("level", 0) == 0 and self._children_of(t.get("name"))) else ""
        parts = [p for p in [sub, summary, floor] if p]
        typ = "｜".join(parts) if parts else "主"
        # 搭接关系：父任务显示其"最早子任务"继承来的前置；子任务显示自身前置
        raw_pre = self._effective_pre(t)
        if raw_pre:
            lag = self._effective_lag(t)
            rel = "、".join(raw_pre)
            if lag:
                rel += (" (延迟%d天)" % lag) if lag > 0 else (" (提前%d天)" % abs(lag))
            pre_txt = rel
        else:
            pre_txt = "—"
        return (t["name"], t["start"], end.isoformat() if end else "", t["days"], pre_txt, typ)

    # ---------- 撤销 + 自动保存 ----------
    def _push_undo(self):
        """把当前 tasks 深拷贝压入撤销栈（在每次修改前调用）"""
        import copy
        self.undo_stack.append(copy.deepcopy(self.tasks))
        if len(self.undo_stack) > self.undo_max:
            self.undo_stack.pop(0)
        self.rule_changed = True   # 规则有改动，需要重算缓存

    def _auto_save(self):
        """自动保存到 autosave.json（每次变更后调用）。
        安全保险：若新数据比旧文件大幅缩水(骤减过半)，先备份旧文件再覆盖，防止误操作丢数据。"""
        try:
            old_count = None
            if os.path.exists(self.autosave_path):
                try:
                    with open(self.autosave_path, "r", encoding="utf-8") as f:
                        old = json.load(f)
                    old_tasks = old.get("tasks", old) if isinstance(old, dict) else old
                    old_count = len(old_tasks) if isinstance(old_tasks, list) else None
                except Exception:
                    old_count = None
            new_count = len(self.tasks)
            # 骤减过半且原本有较多数据 -> 先备份旧文件
            if old_count and old_count >= 5 and new_count < old_count * 0.5:
                try:
                    import time as _time
                    bak = self.autosave_path.replace(".json", "_backup_%s.json" % _time.strftime("%Y%m%d_%H%M%S"))
                    with open(self.autosave_path, "r", encoding="utf-8") as src, \
                         open(bak, "w", encoding="utf-8") as dst:
                        dst.write(src.read())
                except Exception:
                    pass
            with open(self.autosave_path, "w", encoding="utf-8") as f:
                json.dump({"tasks": self.tasks, "shutdowns": self.shutdowns,
                           "sections": self.sections},
                          f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.status.set(f"自动保存失败：{e}")

    def _cache_filename(self, ver):
        return os.path.join(self.cache_dir, "cache_%d.json" % ver)

    def _write_cache(self):
        """计算后写版本化缓存：存每个任务的 名称/开始/结束(计算后值)。删除旧缓存。"""
        self.cache_ver += 1
        data = {}
        for t in self.tasks:
            e = _effective_end(t, self.tasks)
            data[t.get("name")] = {
                "name": t.get("name"),
                "start": t.get("start"),
                "end": e.isoformat() if e else "",
                "days": t.get("days"),
                "parent": t.get("parent"),
                "level": t.get("level", 0),
                "floor": t.get("floor"),
            }
        self.cache_data = data
        try:
            with open(self._cache_filename(self.cache_ver), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
        self._clean_old_cache()

    def _clean_old_cache(self):
        """删除除最新版本外的所有旧缓存文件。"""
        try:
            for fn in os.listdir(self.cache_dir):
                if fn.startswith("cache_") and fn.endswith(".json"):
                    try:
                        ver = int(fn[6:-5])
                    except ValueError:
                        continue
                    if ver != self.cache_ver:
                        try:
                            os.remove(os.path.join(self.cache_dir, fn))
                        except Exception:
                            pass
        except Exception:
            pass

    def _load_latest_cache(self):
        """启动时加载最新版本缓存，存在则直接用(展示缓存结果)。无则 None。"""
        best = None
        try:
            for fn in os.listdir(self.cache_dir):
                if fn.startswith("cache_") and fn.endswith(".json"):
                    try:
                        ver = int(fn[6:-5])
                    except ValueError:
                        continue
                    if best is None or ver > best[0]:
                        best = (ver, os.path.join(self.cache_dir, fn))
            if best:
                with open(best[1], encoding="utf-8") as f:
                    self.cache_data = json.load(f)
                self.cache_ver = best[0]
                return self.cache_data
        except Exception:
            return None
        return None

    def manage_shutdowns(self):
        """打开停歇期管理对话框。停歇期在每次重算(_recalculate_all)时自动代入，
        与手动改任务走同一套逻辑，无需单独顺延，也不会留下手工固定标记。"""
        dlg = ShutdownDialog(self.root, self.shutdowns)
        if dlg.result is None:
            return
        self._push_undo()
        self.shutdowns = dlg.result
        # 清除历史遗留的 manual_fixed 标记（旧算法顺延停歇期时留下的，非用户设置），
        # 让所有任务统一按"工作关系 + 停歇期"重新计算。
        for t in self.tasks:
            t.pop("manual_fixed", None)
        self.rule_changed = True
        self.refresh_all()
        self.status.set(f"已应用 {len(self.shutdowns)} 个停歇期（重算时自动代入）")

    def undo(self):
        """撤销最近一次任务修改"""
        if not self.undo_stack:
            messagebox.showinfo("提示", "没有可撤销的操作", parent=self.root)
            return
        self.tasks = self.undo_stack.pop()
        self.rule_changed = True   # 撤销后需重算缓存
        self.refresh_all()
        self.status.set("已撤销上一步操作")

    def on_close(self):
        """关闭窗口时自动保存，再退出"""
        self._auto_save()
        self.root.destroy()

    def _recalculate_all(self):
        """CPM 正推 + 停歇期代入，循环迭代直到稳定。
        规则：
        - 有前置的任务：start = max(所有前置任务结束日) + 1 + lag，自动重算
        - 无前置或手动固定(manual_fixed=True)的任务：保持现有 start（锚点）
        - **停歇期代入**：任何任务(含锚点)只要与停歇期【区间相交】(开始在内/结束在内/跨越)
          就整体顺延到停歇期结束日+1；多个停歇期相邻则连续推。
        - 每轮：先按工作关系(搭接)正推，再整体代入停歇期，重复直到无变化。
        """
        tasks = self.tasks
        if not tasks:
            return
        name_to_task = {}
        for t in tasks:
            name_to_task[t.get("name")] = t

        # 有子任务的父任务名：其日期由子任务汇总，不单独顺延
        parent_names = {t.get("parent") for t in tasks
                        if t.get("level") == 1 and t.get("parent")}

        # 停歇期区间（按开始排序）
        shutdowns = []
        for sd in self.shutdowns:
            s0 = _parse_date(sd.get("start"))
            s1 = _parse_date(sd.get("end"))
            if s0 and s1 and s0 <= s1:
                shutdowns.append((s0, s1))
        shutdowns.sort()

        def push_out(start_date, days):
            """任务区间 [start, start+days-1] 与任一停歇期相交，则整体推到该期结束+1；
            多个停歇期相邻/重叠时连续推，直到完全避开所有停歇期。"""
            if start_date is None or not shutdowns:
                return start_date
            days = max(1, int(days or 1))
            guard = 0
            while guard < 100:
                guard += 1
                end_date = start_date + timedelta(days=days - 1)
                hit = None
                for (s0, s1) in shutdowns:
                    if start_date <= s1 and end_date >= s0:   # 区间相交
                        cand = s1 + timedelta(days=1)
                        if hit is None or cand > hit:          # 推得最靠后的那个
                            hit = cand
                if hit is None:
                    break
                start_date = hit
            return start_date

        # 迭代：先算工作关系(搭接)，再代入停歇期，直到稳定
        max_iter = len(tasks) * 3 + 50
        changed = True
        iterations = 0
        while changed and iterations < max_iter:
            changed = False
            iterations += 1
            for t in tasks:
                # 汇总父任务：日期由子任务汇总，不单独计算
                if t.get("level", 0) == 0 and t.get("name") in parent_names:
                    continue
                days = t.get("days", 1)
                new_start = _parse_date(t.get("start", ""))
                # 1) 工作关系：搭接正推（有前置且非手动固定）
                if not t.get("manual_fixed") and t.get("pre"):
                    last_end = None
                    for pname in t.get("pre", []):
                        pt = name_to_task.get(pname)
                        if pt is None:
                            continue
                        e = self._effective_end(pt)
                        if e and (last_end is None or e > last_end):
                            last_end = e
                    if last_end is not None:
                        new_start = last_end + timedelta(days=int(t.get("lag", 0)) + 1)
                if new_start is None:
                    continue
                # 2) 停歇期代入（所有任务，含锚点）
                new_start = push_out(new_start, days)
                if new_start.isoformat() != t.get("start"):
                    t["start"] = new_start.isoformat()
                    changed = True
        # 重算后，父任务汇总自动由 _parent_summary 计算

    def _compute_critical_path(self):
        """计算关键路径：找结束时间最晚的任务（排除孤立的手动远期锚点），
        沿其前置链回溯到锚点，这条链上的任务为关键路径。"""
        tasks = self.tasks
        name_to_task = {t.get("name"): t for t in tasks}
        if not name_to_task:
            return set()
        # 结束时间（ordinal）
        def end_ord(t):
            e = self._effective_end(t)
            return e.toordinal() if e else 0
        # 找项目最晚结束的任务（仅考虑有前置链或真实项目任务）
        # 用 memo 算每个任务的"最长依赖结束"
        memo = {}
        def longest_end(t):
            if t.get("name") in memo:
                return memo[t.get("name")]
            pres = t.get("pre", [])
            if not pres:
                val = end_ord(t)
            else:
                max_prev = 0
                for pname in pres:
                    pt = name_to_task.get(pname)
                    if pt:
                        max_prev = max(max_prev, longest_end(pt))
                # 自身结束 >= 前置最晚结束
                val = max(end_ord(t), max_prev)
            memo[t.get("name")] = val
            return val
        for t in tasks:
            longest_end(t)
        if not memo:
            return set()
        max_val = max(v for name, v in memo.items()
                      if not name_to_task.get(name, {}).get("manual_fixed")) if memo else 0
        # 找结束达到 max_val 且【有前置】且【非手动固定】的任务作为关键终点
        # (手动固定任务不参与搭接重算，也不应成为关键路径终点)
        candidates = [t for t in tasks
                      if memo.get(t.get("name")) == max_val and t.get("pre")
                      and not t.get("manual_fixed")]
        if not candidates:
            candidates = [t for t in tasks
                          if memo.get(t.get("name")) == max_val and not t.get("manual_fixed")]
        critical = set()
        # 沿前置链回溯
        stack = list(candidates)
        visited = set()
        while stack:
            t = stack.pop()
            if t.get("name") in visited:
                continue
            visited.add(t.get("name"))
            critical.add(t.get("name"))
            for pname in t.get("pre", []):
                pt = name_to_task.get(pname)
                if pt:
                    stack.append(pt)
        return critical

    def refresh_all(self):
        self._ensure_unique_ids()
        if self.rule_changed or self.cache_data is None:
            self._recalculate_all()
            self.critical_tasks = self._compute_critical_path()
            self._write_cache()
            self.rule_changed = False
        if self.cache_data is None:
            # 仍无缓存(异常)，兜底重算
            self._recalculate_all()
            self.critical_tasks = self._compute_critical_path()
            self._write_cache()
        self.refresh_list()
        self.draw_gantt()
        # 若当前正显示搭接关系图，同步刷新
        try:
            if self.notebook.index(self.notebook.select()) == 1:
                self._draw_network()
        except Exception:
            pass
        self.status.set(f"任务数：{len(self.tasks)}  ·  就绪")
        self._auto_save()

    def draw_gantt(self):
        c = self.canvas
        c.delete("all")
        if not self.tasks:
            c.create_text(20, 20, anchor="nw", text="暂无任务，点左上角【添加任务】开始", fill="#888888")
            return

        # 构建可见任务序列：与【列表显示顺序】一致（甘特图额外按分部聚类）
        # 用 _sorted_tasks(by_section=True)：同一分部的任务在甘特图上也聚在一起
        visible = []
        for t in self._sorted_tasks(by_section=True):
            if t.get("level", 0) == 0:
                visible.append((t, self._display_task(t)))
            else:
                # 子任务：父任务折叠时隐藏；展开时按列表顺序显示
                if t.get("parent") in self.collapsed:
                    continue
                visible.append((t, t))

        self._gantt_visible = visible   # 供点击任务居中甘特图使用

        # 计算时间范围（用展示值）
        starts = []
        ends = []
        for _, d in visible:
            s = _parse_date(d.get("start", ""))
            if s:
                starts.append(s)
            # 汇总父任务用 _summary_end，否则用 _end_of_task
            if d.get("_summary_end"):
                e = _parse_date(d.get("_summary_end"))
            else:
                e = _end_of_task(d)
            if e:
                ends.append(e)
        if not starts or not ends:
            return
        d0, d1 = min(starts), max(ends)
        total_days = (d1 - d0).days + 1
        total_days = max(total_days, 10)
        # 供点击居中水平定位使用
        self._gantt_d0 = d0
        self._gantt_total_days = total_days

        left_margin, top_margin = 70, 50
        row_h = int(30 * self.dpi_scale * self.gantt_zoom)
        self._gantt_row_h = row_h          # 供"今天"竖线使用
        base_font_size = max(9, int(9 * self.dpi_font * self.gantt_zoom))   # 甘特图进度条字体(温和)
        # 文字宽度测量字体(用于判断是否超出色块)
        try:
            import tkinter.font as _tkfont
            _name_font = _tkfont.Font(family="TkDefaultFont", size=base_font_size, weight="bold")
        except Exception:
            _name_font = None
        tick_font_size = max(8, int(8 * self.dpi_font * self.gantt_zoom))   # 时间轴刻度字体
        # 内容宽度：按"每天固定像素"让进度条清晰可读，超出画布靠横向滚动。
        # 每天 px 约 4（随缩放系数变化），保证进度条有合理宽度；至少占满画布。
        px_per_day = 4.0 * self.gantt_zoom
        content_w = int(total_days * px_per_day)
        avail_w = max(self.canvas.winfo_width() - left_margin - 30, 200)
        # 若内容比画布窄，则铺满画布(scale 放大)
        chart_w = max(content_w, avail_w)
        self._gantt_chart_w = chart_w
        scale = chart_w / total_days

        # 智能时间轴刻度：根据总天数自动选间隔，避免日期文字重叠
        # 估算能容纳的刻度数(每刻度约需 60px)
        max_ticks = max(int(chart_w / 60), 2)
        # 候选间隔(天)：按粒度递增
        candidates = [1, 2, 3, 7, 14, 30, 60, 90, 180, 365]
        interval = 1
        for cand in candidates:
            if total_days / cand <= max_ticks:
                interval = cand
                break
        else:
            interval = 365
        month_labels = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]

        c.create_line(left_margin, top_margin, left_margin + chart_w, top_margin, fill="#CCCCCC")
        i = 0
        while i <= total_days:
            x = left_margin + i * scale
            d = d0 + timedelta(days=i)
            # 刻度线
            c.create_line(x, top_margin - 6, x, top_margin, fill="#CCCCCC")
            # 标签：若间隔>=14天显示"年-月"，否则显示"月/日"
            if interval >= 14:
                label = "%d-%02d" % (d.year, d.month)
            else:
                label = "%d/%d" % (d.month, d.day)
            c.create_text(x, top_margin - 12, text=label, fill="#666666", font=("", tick_font_size))
            i += interval
        # 纵向网格线(按刻度间隔)
        i = 0
        while i <= total_days:
            x = left_margin + i * scale
            c.create_line(x, top_margin, x, top_margin + len(visible) * row_h, fill="#EFEFEF")
            i += interval

        # 绘制：父任务用汇总条(若有子任务)，子任务独立彩条
        for i, (orig, disp) in enumerate(visible):
            s = _parse_date(disp.get("start", ""))
            if not s:
                continue
            y = top_margin + i * row_h
            # 父任务若有子任务：汇总条从最早子开始到最晚子结束
            if orig.get("level", 0) == 0 and self._children_of(orig.get("name")):
                summary = self._parent_summary(orig)
                if summary is None:
                    continue
                s = _parse_date(summary["start"])
                e = _parse_date(summary["end"])
                x0 = left_margin + (s - d0).days * scale
                x1 = left_margin + ((e - d0).days + 1) * scale
                color = orig.get("color", COLORS[0])
                # 关键路径任务：红色描边标记
                is_crit = orig.get("name") in self.critical_tasks
                outline = "#E53935" if is_crit else ""
                # 汇总条：用空心/细长条区分
                c.create_rectangle(x0, y + 9, x1, y + row_h - 9, fill=color, outline=outline,
                                   width=2 if is_crit else 0)
                _name = orig.get("name", "")
                _fname = "#000000" if (_name_font and _name_font.measure(_name) > (x1 - x0)) else "#FFFFFF"
                c.create_text(x0 + 3, y + row_h / 2, text=_name, anchor="w",
                              fill=_fname, font=("", base_font_size, "bold"))
            else:
                e = _end_of_task(disp)
                if not e:
                    continue
                x0 = left_margin + (s - d0).days * scale
                x1 = x0 + int(disp.get("days", 1)) * scale
                color = orig.get("color", COLORS[0])
                is_crit = orig.get("name") in self.critical_tasks
                outline = "#E53935" if is_crit else ""
                c.create_rectangle(x0, y + 6, x1, y + row_h - 6, fill=color, outline=outline,
                                   width=2 if is_crit else 0)
                _name = orig.get("name", "")
                _fname = "#000000" if (_name_font and _name_font.measure(_name) > (x1 - x0)) else "#FFFFFF"
                c.create_text(x0 + 3, y + row_h / 2, text=_name, anchor="w",
                              fill=_fname, font=("", base_font_size, "bold"))
            c.create_line(left_margin, y + row_h, left_margin + chart_w, y + row_h, fill="#EEEEEE")
        # 搜索命中的任务：在甘特图上描柠檬黄边框 + 铺淡黄底，与列表的标黄呼应
        self._mark_gantt_hits(visible, left_margin, top_margin, row_h, chart_w)
        # "今天"竖线（勾选后显示）
        if getattr(self, "var_today", None) is not None and self.var_today.get():
            self._draw_today_line(d0, d1, scale, left_margin, top_margin,
                                  len(visible), chart_w)
        c.configure(scrollregion=(0, 0, left_margin + chart_w + 20, top_margin + len(visible) * row_h + 20))

    def _mark_gantt_hits(self, visible, left_margin, top_margin, row_h, chart_w):
        """甘特图同步高亮：主目标行加醒目黄框，其它命中行只铺淡黄底。"""
        kw = (self.search_var.get() if hasattr(self, "search_var") else "").strip()
        if not kw:
            return
        low = kw.lower()
        main_id = getattr(self, "_search_main", None)
        c = self.canvas
        for i, (orig, disp) in enumerate(visible):
            if low not in orig.get("name", "").lower():
                continue
            y = top_margin + i * row_h
            # 命中行都铺一层淡黄底
            c.create_rectangle(left_margin, y, left_margin + chart_w, y + row_h,
                               fill="#FFFDE7", outline="", tags="searchhit")
            # 主目标额外描醒目黄框（不抢色块、也不跟进度条撞色）
            if orig.get("_id") == main_id:
                c.create_rectangle(left_margin, y + 1,
                                   left_margin + chart_w, y + row_h - 1,
                                   outline="#FFC400", width=2, tags="searchhit")
        # 黄底要垫在进度条下面，避免盖住色块
        try:
            c.tag_lower("searchhit")
        except Exception:
            pass

    # ---------- 保存 / 打开 ----------
    def save(self):
        if not self.data_file:
            path = filedialog.asksaveasfilename(defaultextension=".json",
                                                filetypes=[("JSON 数据", "*.json")])
            if not path:
                return
            self.data_file = path
        with open(self.data_file, "w", encoding="utf-8") as f:
            json.dump({"tasks": self.tasks, "shutdowns": self.shutdowns,
                       "sections": self.sections},
                      f, ensure_ascii=False, indent=2)
        self.status.set(f"已保存到 {self.data_file}")

    def import_config(self):
        """导入：读取『保存』生成的配置文件({tasks, shutdowns})，兼容旧格式(纯任务数组)。
        带文件校验、导入前确认、撤销支持。"""
        path = filedialog.askopenfilename(
            title="导入配置文件",
            filetypes=[("甘特图配置文件", "*.json"), ("JSON 数据", "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        # 读取 + 校验
        try:
            with open(path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
        except json.JSONDecodeError as e:
            messagebox.showerror("导入失败", f"文件不是合法的 JSON：\n{e}", parent=self.root)
            return
        except Exception as e:
            messagebox.showerror("导入失败", f"无法读取文件：\n{e}", parent=self.root)
            return

        if isinstance(loaded, dict):
            new_tasks = loaded.get("tasks", [])
            new_shutdowns = loaded.get("shutdowns", [])
            new_sections = loaded.get("sections", []) or []
        elif isinstance(loaded, list):
            new_tasks = loaded
            new_shutdowns = []
            new_sections = []
        else:
            messagebox.showerror("导入失败", "文件格式不正确：既不是配置对象也不是任务数组。", parent=self.root)
            return
        if not isinstance(new_tasks, list):
            messagebox.showerror("导入失败", "配置中的 tasks 不是列表，文件可能已损坏。", parent=self.root)
            return
        # 简单校验每个任务是 dict
        bad = [i for i, t in enumerate(new_tasks) if not isinstance(t, dict)]
        if bad:
            messagebox.showerror("导入失败", f"第 {bad[:5]} 项不是任务对象，文件可能已损坏。", parent=self.root)
            return

        # 导入前确认
        if not messagebox.askyesno(
                "确认导入",
                f"即将导入 {len(new_tasks)} 个任务、{len(new_shutdowns)} 个停歇期。\n"
                f"当前有 {len(self.tasks)} 个任务，将被替换（可用『撤销』回退）。\n\n确定导入吗？",
                parent=self.root):
            return

        self._push_undo()
        self.tasks = new_tasks
        self.shutdowns = new_shutdowns
        self.sections = new_sections
        self.data_file = path
        self.rule_changed = True
        self.refresh_all()
        self.status.set(f"已导入 {len(self.tasks)} 个任务 ← {path}")
        self._check_data_health("导入")

    # 兼容旧名
    open = import_config

    def clear_all(self):
        if messagebox.askyesno("确认", "确定清空所有任务？"):
            self._push_undo()
            self.tasks = []
            self.refresh_all()

    # ---------- 导出 Excel ----------
    def export_excel(self):
        """导出两个 sheet：主任务计划(按开始升序)+主任务甘特图；全部任务+完整甘特图。
        样式仿参考工作表：微软雅黑、深蓝表头白字、medium边框、行高30、居中。"""
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
        except ImportError:
            messagebox.showerror("错误", "需要 openpyxl 库，请先运行 pip install openpyxl")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return

        # ---- 样式：深蓝表头 + 微软雅黑（国内工程表格常见风格）----
        FN = "微软雅黑"
        f_title = Font(name=FN, size=14, bold=True)
        f_head = Font(name=FN, size=12, bold=True, color="FFF9FBFA")
        f_data = Font(name=FN, size=11)
        f_data_b = Font(name=FN, size=11, bold=True)
        f_sub = Font(name=FN, size=10, color="FF4B5563")
        f_gm = Font(name=FN, size=9, bold=True)
        f_gw = Font(name=FN, size=7)
        fill_head = PatternFill("solid", fgColor="2F5597")
        fill_gm = PatternFill("solid", fgColor="D9E2F3")
        fill_gw = PatternFill("solid", fgColor="EDF2FA")
        al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
        al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
        med = Side(style="medium", color="FF000000")
        thin = Side(style="thin", color="FFB0B7C3")
        bd_med = Border(left=med, right=med, top=med, bottom=med)
        bd_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

        # ---- 任务数据(用计算结果，与界面一致) ----
        def eff(t):
            if t.get("level", 0) == 0 and self._children_of(t.get("name")):
                s = self._parent_summary(t)
                if s:
                    return _parse_date(s["start"]), _parse_date(s["end"])
            return _parse_date(t.get("start", "")), _effective_end(t, self.tasks)

        def pre_desc(t):
            pres = t.get("pre", [])
            if not pres:
                return "—"
            lag = int(t.get("lag", 0) or 0)
            return "、".join(pres) + ("后%dd" % (lag + 1))

        parents = [t for t in self.tasks if t.get("level", 0) == 0]
        children = [t for t in self.tasks if t.get("level", 0) == 1]

        # sheet1 行：仅主任务，按开始时间升序
        rows1 = [(t, *eff(t), 0) for t in
                 sorted(parents, key=lambda x: (eff(x)[0] or date(2999, 1, 1)))]
        # sheet2 行：层级顺序(父任务 + 其子任务按开始升序)
        rows2 = []
        for p in parents:
            s, e = eff(p)
            rows2.append((p, s, e, 0))
            subs = sorted([c for c in children if c.get("parent") == p.get("name")],
                          key=lambda x: x.get("start", ""))
            for c in subs:
                cs, ce = eff(c)
                rows2.append((c, cs, ce, 1))
        for c in children:
            if not any(c.get("parent") == p.get("name") for p in parents):
                cs, ce = eff(c)
                rows2.append((c, cs, ce, 1))

        # ---- 甘特图周列范围 ----
        def monday(d):
            return d - timedelta(days=d.weekday())
        all_s = [r[1] for r in rows1 + rows2 if r[1]]
        all_e = [r[2] for r in rows1 + rows2 if r[2]]
        if not all_s or not all_e:
            messagebox.showwarning("提示", "没有可导出的任务日期", parent=self.root)
            return
        g0 = monday(min(all_s))
        g1 = monday(max(all_e))
        n_weeks = (g1 - g0).days // 7 + 1

        HEAD_COLS = 6   # A序号 B名称 C开始 D完成 E工期 F搭接
        GCOL = HEAD_COLS + 1  # 甘特图起始列 G

        def build_sheet(ws, title, rows, numbered_subs):
            last_col = HEAD_COLS + n_weeks
            last_L = get_column_letter(last_col)
            # 行1 标题
            ws.merge_cells("A1:%s1" % last_L)
            c1 = ws.cell(1, 1, title)
            c1.font = f_title
            c1.alignment = al_c
            ws.row_dimensions[1].height = 30
            # 行2 甘特年月(合并同月周列) / A2:F3 合并"任务信息"
            ws.merge_cells("A2:F3")
            c2 = ws.cell(2, 1, "任务信息")
            c2.font = f_head
            c2.fill = fill_head
            c2.alignment = al_c
            run_start = 0
            for i in range(1, n_weeks + 1):
                wk = g0 + timedelta(days=7 * i)
                prev_wk = g0 + timedelta(days=7 * (i - 1))
                if i == n_weeks or (wk.year, wk.month) != (prev_wk.year, prev_wk.month):
                    a = GCOL + run_start
                    b = GCOL + i - 1
                    if b > a:
                        ws.merge_cells(start_row=2, start_column=a, end_row=2, end_column=b)
                    cellm = ws.cell(2, a, "%d年%d月" % (prev_wk.year, prev_wk.month))
                    cellm.font = f_gm
                    cellm.fill = fill_gm
                    cellm.alignment = al_c
                    run_start = i
            for cc in range(GCOL, last_col + 1):
                cell = ws.cell(2, cc)
                cell.fill = fill_gm
                cell.border = bd_thin
            ws.row_dimensions[2].height = 18
            # 行3 周日期
            for i in range(n_weeks):
                wk = g0 + timedelta(days=7 * i)
                cell = ws.cell(3, GCOL + i, "%d/%d" % (wk.month, wk.day))
                cell.font = f_gw
                cell.fill = fill_gw
                cell.alignment = al_c
                cell.border = bd_thin
            ws.row_dimensions[3].height = 16
            # 行4 表头
            heads = ["序号", "任务名称", "开始日期", "计划完成日期", "工期(天)", "搭接关系"]
            for i, h in enumerate(heads):
                cell = ws.cell(4, 1 + i, h)
                cell.font = f_head
                cell.fill = fill_head
                cell.alignment = al_c
                cell.border = bd_med
            ws.merge_cells(start_row=4, start_column=GCOL, end_row=4, end_column=last_col)
            cg = ws.cell(4, GCOL, "进度甘特图（按周）")
            cg.font = f_head
            cg.fill = fill_head
            cg.alignment = al_c
            for cc in range(GCOL, last_col + 1):
                ws.cell(4, cc).fill = fill_head
                ws.cell(4, cc).border = bd_med
            ws.row_dimensions[4].height = 28
            # 数据行
            r = 5
            main_no = 0
            sub_no = 0
            for (t, s, e, lvl) in rows:
                if lvl == 0:
                    main_no += 1
                    sub_no = 0
                    no = main_no
                else:
                    sub_no += 1
                    no = ("%d-%d" % (main_no, sub_no)) if numbered_subs else (main_no * 100 + sub_no)
                ws.cell(r, 1, no)
                nm = t.get("name", "")
                ws.cell(r, 2, ("      └ " + nm) if lvl == 1 else nm)
                ws.cell(r, 3, s if s else "")
                ws.cell(r, 4, e if e else "")
                ws.cell(r, 5, ((e - s).days + 1) if (s and e) else "")
                ws.cell(r, 6, pre_desc(t))
                for ci in range(1, HEAD_COLS + 1):
                    cell = ws.cell(r, ci)
                    cell.border = bd_med
                    cell.alignment = al_l if ci in (2, 6) else al_c
                    if lvl == 1:
                        cell.font = f_sub
                    else:
                        cell.font = f_data_b if ci == 2 else f_data
                    if ci in (3, 4) and cell.value:
                        cell.number_format = "yyyy/m/d"
                # 甘特色块
                if s and e:
                    color = (t.get("color") or "#4472C4").lstrip("#")
                    fill_bar = PatternFill("solid", fgColor=color.upper())
                    w0 = max(0, (monday(s) - g0).days // 7)
                    w1 = min(n_weeks - 1, (monday(e) - g0).days // 7)
                    for wi in range(w0, w1 + 1):
                        cell = ws.cell(r, GCOL + wi)
                        cell.fill = fill_bar
                        cell.border = bd_thin
                # 甘特空列也加细边框(色块已加)
                for wi in range(n_weeks):
                    cell = ws.cell(r, GCOL + wi)
                    if cell.fill is None or cell.fill.fill_type != "solid":
                        cell.border = bd_thin
                ws.row_dimensions[r].height = 26
                r += 1
            # 列宽
            ws.column_dimensions["A"].width = 7
            ws.column_dimensions["B"].width = 40
            ws.column_dimensions["C"].width = 12.5
            ws.column_dimensions["D"].width = 12.5
            ws.column_dimensions["E"].width = 9
            ws.column_dimensions["F"].width = 36
            for i in range(n_weeks):
                ws.column_dimensions[get_column_letter(GCOL + i)].width = 3.2
            ws.freeze_panes = "G5"

        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "主任务计划"
        build_sheet(ws1, "施工总进度计划 · 主任务（按开始时间升序）", rows1, numbered_subs=False)
        ws2 = wb.create_sheet("全部任务计划")
        build_sheet(ws2, "施工进度计划 · 全部任务（主任务+子任务）", rows2, numbered_subs=True)
        try:
            wb.save(path)
        except PermissionError:
            messagebox.showerror("错误", "文件被占用，请先关闭同名 Excel 文件再导出", parent=self.root)
            return
        messagebox.showinfo("完成", f"已导出到 {path}\nSheet1=主任务计划(升序+甘特)\nSheet2=全部任务计划(层级+甘特)")
        self.status.set(f"已导出 {path}")

    # ---------- 示例数据 ----------
    def load_demo(self):
        self._push_undo()
        self.tasks = [
            {"name": "三通一平/场地平整", "start": "2026-03-01", "days": 15, "color": COLORS[0], "pre": [], "lag": 0, "level": 0, "parent": None, "floor": None, "_id": ""},
            {"name": "土方开挖", "start": "2026-03-20", "days": 20, "color": COLORS[1], "pre": [], "lag": 0, "level": 0, "parent": None, "floor": None, "_id": ""},
            {"name": "桩基施工", "start": "2026-04-10", "days": 25, "color": COLORS[2], "pre": [], "lag": 0, "level": 0, "parent": None, "floor": None, "_id": ""},
            {"name": "基础底板施工", "start": "2026-05-08", "days": 20, "color": COLORS[3], "pre": [], "lag": 0, "level": 0, "parent": None, "floor": None, "_id": ""},
            {"name": "主体结构(±0.00以上)", "start": "2026-05-30", "days": 90, "color": COLORS[4], "pre": [], "lag": 0, "level": 0, "parent": None, "floor": None, "_id": ""},
            {"name": "砌体工程", "start": "2026-08-20", "days": 30, "color": COLORS[5], "pre": [], "lag": 0, "level": 0, "parent": None, "floor": None, "_id": ""},
            {"name": "机电安装", "start": "2026-09-20", "days": 60, "color": COLORS[6], "pre": [], "lag": 0, "level": 0, "parent": None, "floor": None, "_id": ""},
            {"name": "精装修工程", "start": "2026-11-20", "days": 60, "color": COLORS[7], "pre": [], "lag": 0, "level": 0, "parent": None, "floor": None, "_id": ""},
            {"name": "室外配套/园林景观", "start": "2027-01-20", "days": 45, "color": COLORS[8], "pre": [], "lag": 0, "level": 0, "parent": None, "floor": None, "_id": ""},
            {"name": "竣工验收", "start": "2027-03-10", "days": 30, "color": COLORS[9], "pre": [], "lag": 0, "level": 0, "parent": None, "floor": None, "_id": ""},
        ]
        self.refresh_all()
        self.status.set("已载入示例数据（可清空后录入真实任务）")


def main():
    # Windows DPI 感知：让 Tkinter 按系统缩放渲染，文字不糊
    try:
        import ctypes
        # 尝试高精度 DPI 感知(推荐)
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass
        # 读取当前 DPI 缩放，存入环境变量供后续调整
        try:
            dpi = ctypes.windll.user32.GetDpiForSystem()
        except Exception:
            dpi = 96
        try:
            scaling = dpi / 96.0
            os.environ["TK_DPI_SCALING"] = str(round(scaling, 2))
        except Exception:
            pass
    except Exception:
        pass

    root = tk.Tk()
    # 用 Tk 自带的 ttk 缩放适配 DPI(若支持)
    try:
        import tkinter.font as tkfont
        # 读取系统默认字体缩放系数
        scaling = float(os.environ.get("TK_DPI_SCALING", "1.0"))
        root.tk.call("tk", "scaling", scaling * 1.333)  # Tk默认缩放基准
    except Exception:
        pass
    GanttApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
