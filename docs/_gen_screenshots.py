# -*- coding: utf-8 -*-
"""教程截图生成器 v2 —— 避开模态对话框阻塞问题。

思路：不用 app.add_task() / app.add_subtask() 这类会 wait_window 的入口，
      而是直接构造对话框对象（构造 ≠ 阻塞），截完就 destroy。
"""
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import os
import sys
import time
import tkinter as tk
from datetime import date, timedelta
from PIL import ImageGrab

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
OUT = os.path.join(HERE, "_tutorial_shots")
os.makedirs(OUT, exist_ok=True)

import gantt_tool as G

# ============================================================
# 补丁：TaskDialog.__init__ 结尾会 wait_window() 阻塞，
# 导致截图脚本卡死。这里把 wait_window / grab_set 变成空操作，
# 这样对话框可以构造出来展示，但不会进入事件等待。
# 仅影响本脚本进程，不修改源码文件。
# ============================================================
import tkinter as _tk
_orig_wait = _tk.Misc.wait_window
_orig_grab = _tk.Misc.grab_set


def _no_wait(self, window=None):
    return None


def _no_grab(self):
    return None


_tk.Misc.wait_window = _no_wait
_tk.Misc.grab_set = _no_grab
_tk.Toplevel.wait_window = _no_wait
try:
    _tk.Toplevel.grab_set = _no_grab
except Exception:
    pass


BS = chr(92)
shot_log = []


def grab(root, name, region=None, wait=0.6):
    root.update()
    time.sleep(wait)
    if region == "top":
        x, y = root.winfo_rootx(), root.winfo_rooty()
        box = (x, y, x + root.winfo_width(), y + 64)
    elif isinstance(region, tuple):
        box = region
    else:
        x, y = root.winfo_rootx(), root.winfo_rooty()
        box = (x, y, x + root.winfo_width(), y + root.winfo_height())
    img = ImageGrab.grab(bbox=box)
    p = os.path.join(OUT, name + ".png")
    img.save(p)
    shot_log.append((name, img.size))
    print("  [shot] %-26s %s" % (name, img.size))
    return p


def grab_widget(root, wdg, name, pad=8, wait=0.5):
    root.update()
    time.sleep(wait)
    x, y = wdg.winfo_rootx() - pad, wdg.winfo_rooty() - pad
    bw, bh = wdg.winfo_width() + pad * 2, wdg.winfo_height() + pad * 2
    img = ImageGrab.grab(bbox=(x, y, x + bw, y + bh))
    p = os.path.join(OUT, name + ".png")
    img.save(p)
    shot_log.append((name, img.size))
    print("  [shot] %-26s %s" % (name, img.size))
    return p


def main():
    root = tk.Tk()
    app = G.GanttApp(root)
    # 拦住一切写盘/关闭
    app._auto_save = lambda: None
    app.on_close = lambda: None
    app.data_file = ""

    # ★ 最大化窗口：既隐藏桌面隐私，又让各区域比例正常、信息完整
    root.deiconify()
    try:
        root.state("zoomed")          # Windows 上 = 最大化
    except Exception:
        try:
            sw = root.winfo_screenwidth()
            sh = root.winfo_screenheight()
            root.geometry("%dx%d+0+0" % (sw, sh))
        except Exception:
            root.geometry("1600x1000+0+0")
    root.attributes("-topmost", True)
    root.lift()
    root.focus_force()
    root.update()
    time.sleep(1.5)

    W, H = root.winfo_width(), root.winfo_height()
    RX, RY = root.winfo_rootx(), root.winfo_rooty()

    # ===== 1) 示例数据看全貌 =====
    print("\n[1] 全貌")
    app.load_demo()
    app.add_task = app.add_task          # 占位
    app.section_collapsed = set()
    app.refresh_all()
    root.update()
    time.sleep(0.8)
    grab(root, "01-主界面全貌")
    grab(root, "02-工具栏", region="top")
    # 左侧列表区域
    try:
        tree = app.tree
        tx, ty = tree.winfo_rootx(), tree.winfo_rooty()
        tw, th = tree.winfo_width(), tree.winfo_height()
        grab(root, "03-任务列表", region=(tx - 4, ty - 28, tx + tw + 4, ty + min(th, 420)))
    except Exception as e:
        print("   列表截图失败:", e)

    # ===== 2) 清空 =====
    print("\n[2] 空白起点")
    app.tasks = []
    app.sections = []
    app.shutdowns = []
    app.collapsed = set()
    app.section_collapsed = set()
    app.refresh_all()
    root.update()
    time.sleep(0.6)
    grab(root, "04-空白起点")

    # ===== 3) 任务信息对话框（构造但不阻塞）=====
    print("\n[3] 任务对话框")
    try:
        dlg = G.TaskDialog(root, title="新建任务", default=None, candidates=[])
        dlg.geometry("+%d+%d" % (RX + 260, RY + 120))
        root.update()
        time.sleep(0.8)
        try:
            dlg.var_name.set("土方开挖")
            dlg.var_start.set("2026-03-20")
            dlg.var_end.set("2026-04-08")
            dlg.var_days.set("20")
            dlg.var_lag.set("4")
        except Exception:
            pass
        root.update()
        time.sleep(0.4)
        grab_widget(root, dlg, "05-任务信息对话框")
        dlg.destroy()
    except Exception as e:
        print("   任务对话框失败:", e)
    root.update()

    # ===== 4) 录入主任务 =====
    print("\n[4] 录入主任务")
    demo = [
        ("三通一平/场地平整", 15, []),
        ("土方开挖", 20, ["三通一平/场地平整"]),
        ("桩基施工", 25, ["土方开挖"]),
        ("基础底板施工", 20, ["桩基施工"]),
        ("主体结构", 90, ["基础底板施工"]),
        ("砌体工程", 30, ["主体结构"]),
        ("机电安装", 60, ["砌体工程"]),
        ("竣工备案", 30, ["机电安装"]),
    ]
    d = date(2026, 3, 1)
    cur_end = d
    for i, (nm, dy, pre) in enumerate(demo):
        if pre:
            # 接着前一个的结束日
            pass
        start = d if i == 0 else cur_end + timedelta(days=1)
        app.tasks.append({
            "name": nm, "start": start.isoformat(), "days": dy,
            "color": G.COLORS[i % len(G.COLORS)],
            "pre": pre, "lag": 0, "level": 0, "parent": None,
            "floor": None, "_id": str(i + 1),
        })
        cur_end = start + timedelta(days=dy - 1)
    app.refresh_all()
    root.update()
    time.sleep(0.9)
    grab(root, "06-录入主任务")

    # ===== 5) 建分部 =====
    print("\n[5] 分部工程")
    app.sections = [
        {"name": "地基与基础工程"},
        {"name": "主体结构工程"},
        {"name": "装饰装修工程"},
    ]
    mg = {
        "三通一平/场地平整": "地基与基础工程",
        "土方开挖": "地基与基础工程",
        "桩基施工": "地基与基础工程",
        "基础底板施工": "地基与基础工程",
        "主体结构": "主体结构工程",
        "砌体工程": "主体结构工程",
    }
    for t in app.tasks:
        t["section"] = mg.get(t["name"], "")
    app.refresh_all()
    root.update()
    time.sleep(0.9)
    grab(root, "07-建好分部工程")
    try:
        tree = app.tree
        tx, ty = tree.winfo_rootx(), tree.winfo_rooty()
        tw, th = tree.winfo_width(), tree.winfo_height()
        grab(root, "08-分部下的任务", region=(tx - 4, ty - 28, tx + tw + 4, ty + min(th, 400)))
    except Exception:
        pass

    # ===== 6) 折叠分部 =====
    print("\n[6] 折叠分部")
    app.section_collapsed = set(app._section_names())
    app.refresh_list()
    root.update()
    time.sleep(0.7)
    try:
        tree = app.tree
        tx, ty = tree.winfo_rootx(), tree.winfo_rooty()
        tw, th = tree.winfo_width(), tree.winfo_height()
        grab(root, "09-分部折叠后", region=(tx - 4, ty - 28, tx + tw + 4, ty + min(th, 300)))
    except Exception:
        pass

    # ===== 7) 子任务对话框 =====
    print("\n[7] 新建子任务对话框")
    parent = next(t for t in app.tasks if t["name"] == "主体结构")
    try:
        sd = G.SubTaskDialog(root, parent, app.tasks)
        sd.geometry("+%d+%d" % (RX + 200, RY + 60))
        root.update()
        time.sleep(0.9)
        try:
            sd.var_name.set("主体结构")
            sd.var_mode.set("standard")
            sd.var_start_floor.set("2")
            sd.var_end_floor.set("18")
            sd.var_days_per_floor.set("7")
        except Exception as e:
            print("   设参失败:", e)
        root.update()
        time.sleep(0.5)
        grab_widget(root, sd, "10-新建子任务对话框")
        sd.destroy()
    except Exception as e:
        print("   子任务对话框失败:", e)
    root.update()

    # ===== 8) 生成标准层子任务 =====
    print("\n[8] 标准层子任务")
    pid = parent["_id"]
    floors = ["B1", "1F", "2F", "3F", "4F", "5F", "6F"]
    pstart = date(2026, 6, 1)
    for i, fl in enumerate(floors):
        st = pstart + timedelta(days=i * 14)
        app.tasks.append({
            "name": "主体结构" + fl, "start": st.isoformat(), "days": 12,
            "color": parent["color"], "pre": [], "lag": 0,
            "level": 1, "parent": "主体结构", "floor": i,
            "_id": "%s@%d" % (pid, i + 1),
        })
    for i in range(1, len(floors)):
        for t in app.tasks:
            if t["name"] == "主体结构" + floors[i]:
                t["pre"] = ["主体结构" + floors[i - 1]]
    app.collapsed.discard("主体结构")
    app.section_collapsed.discard("主体结构工程")
    app.refresh_all()
    root.update()
    time.sleep(1.0)
    grab(root, "11-生成标准层子任务")

    # ===== 9) 搭接关系图 =====
    print("\n[9] 搭接关系图")
    try:
        app.notebook.select(1)
        root.update()
        app._draw_network()
        root.update()
        time.sleep(1.2)
        grab(root, "12-搭接关系图")
    except Exception as e:
        print("   网络图失败:", e)
    try:
        app.notebook.select(0)
        root.update()
    except Exception:
        pass

    # ===== 10) 搜索 =====
    print("\n[10] 搜索")
    app.search_var.set("主体")
    app.do_search()
    root.update()
    time.sleep(1.0)
    grab(root, "13-搜索定位")
    app.clear_search()
    root.update()

    # ===== 11) 显示今天 =====
    print("\n[11] 显示今天")
    if app.tasks:
        app.tasks[0]["start"] = (date.today() - timedelta(days=20)).isoformat()
    app.var_today.set(True)
    app.rule_changed = True
    app.refresh_all()
    root.update()
    time.sleep(1.0)
    grab(root, "14-显示今天")
    app.var_today.set(False)

    # ===== 12) 缩放 =====
    print("\n[12] 缩放")
    app.gantt_zoom = 1.5
    app.refresh_all()
    root.update()
    time.sleep(0.9)
    grab(root, "15-画布放大")

    print("\n=== 完成，共 %d 张 ===" % len(shot_log))
    for n, s in shot_log:
        print("   %-26s %s" % (n, s))
    print("\n输出:", OUT)

    root.attributes("-topmost", False)
    root.destroy()


if __name__ == "__main__":
    main()
