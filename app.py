"""谁没交？—— 文件夹名单核对器

左边拖 Excel 名单，右边拖文件夹，点一下就知道谁没交。

匹配逻辑全部在 matcher.py；这个文件只管界面。
只用标准库（tkinter + ctypes），打包出来约 15MB。
"""

from __future__ import annotations

import calendar
import ctypes
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import matcher as M

# --------------------------------------------------------------------------
# 外观
# --------------------------------------------------------------------------

APP_TITLE = "谁没交？—— 文件夹名单核对器"

BG = "#eef1f5"
CARD = "#ffffff"
LINE = "#d3d9e0"
TEXT = "#1f2328"
MUTED = "#6b7280"
ACCENT = "#2f6fed"
OK = "#1a7f37"
WARN = "#bf6a02"
DANGER = "#c0392b"

F = "Microsoft YaHei UI"
F_TITLE = (F, 17, "bold")
F_SUB = (F, 9)
F_BODY = (F, 10)
F_BOLD = (F, 10, "bold")
F_BIG = (F, 15, "bold")
F_MONO = ("Consolas", 9)


# --------------------------------------------------------------------------
# 拖放
#
# 原来是拿 ctypes 子类化窗口过程接 WM_DROPFILES，**这条路走不通**：
# 真实拖放是 PostMessage 投递的，消息在 Tk 的 mainloop 里被处理，
# 此时 Python 已经放开了 GIL；ctypes 回调再进来会让线程状态错乱，
# 直接 Fatal Python error: PyEval_RestoreThread（0xC000041D）把进程带走。
# 最小复现（Tk + 子类化 + PostMessage）稳定崩溃，跟程序逻辑无关。
#
# 所以改用 tkinterdnd2：拖放由 Tcl 的 tkdnd 扩展在 C 层处理，
# 不经过任何 Python 回调，也就没有 GIL 这回事。
# --------------------------------------------------------------------------

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    _HAS_DND = False


def make_root() -> tk.Tk:
    """有 tkinterdnd2 就用它，没有就退回普通 Tk（拖放不可用，按钮还能用）。"""
    if _HAS_DND:
        return TkinterDnD.Tk()
    return tk.Tk()


def parse_drop(data: str) -> list[str]:
    """把 tkinterdnd2 给的 event.data 拆成路径列表。

    格式是 Tcl 列表：带空格的路径会用 {} 包起来，
    比如 `{D:/我的 作业/张三.docx} D:/b.docx`。
    """
    out: list[str] = []
    buf = ""
    depth = 0
    for ch in data or "":
        if ch == "{":
            depth += 1
            if depth == 1:
                continue
        elif ch == "}":
            depth -= 1
            if depth == 0:
                out.append(buf)
                buf = ""
                continue
        elif ch == " " and depth == 0:
            if buf:
                out.append(buf)
                buf = ""
            continue
        buf += ch
    if buf:
        out.append(buf)
    return [p for p in out if p.strip()]


# --------------------------------------------------------------------------
# 导出（抽成纯函数，方便无头自测）
# --------------------------------------------------------------------------


def write_result_xlsx(path: Path, rows, missing) -> None:
    """rows: [(姓名, 编号, 状态, 依据, tag), ...]"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "核对结果"
    ws.append(["姓名", "编号", "状态", "依据"])
    for c in range(1, 5):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="E3E8EF")

    color = {"未交": "C0392B", "待确认": "BF6A02",
             "已交": "1A7F37", "认不出": "6B7280"}
    for name, sid, status, ev, _tag in rows:
        ws.append([name, sid, status, ev])
        ws.cell(row=ws.max_row, column=3).font = Font(color=color.get(status, "000000"))

    for col, w in (("A", 16), ("B", 16), ("C", 10), ("D", 60)):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        row[3].alignment = Alignment(vertical="center")

    # 第二页：只留未交，方便直接打印或转发
    ws2 = wb.create_sheet("未交名单")
    ws2.append(["姓名", "编号"])
    for p in missing:
        ws2.append([p.name, p.sid])
    ws2.column_dimensions["A"].width = 16
    ws2.column_dimensions["B"].width = 16

    wb.save(str(path))


# --------------------------------------------------------------------------
# 日历弹窗：tkinter 没有内置的，为这点功能把 tkcalendar（连 babel）装进来
# 不划算，自己画一个够用的
# --------------------------------------------------------------------------


class DatePicker(tk.Toplevel):
    WEEK = ("一", "二", "三", "四", "五", "六", "日")

    def __init__(self, parent, initial: datetime | None = None, on_pick=None):
        super().__init__(parent)
        self.on_pick = on_pick
        self.view = (initial or datetime.now()).replace(day=1)

        self.title("选个日期")
        self.configure(bg=CARD)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        head = tk.Frame(self, bg=CARD)
        head.pack(fill="x", padx=12, pady=(12, 4))
        tk.Button(head, text="‹", font=(F, 13), bg=CARD, fg=ACCENT,
                  relief="flat", bd=0, cursor="hand2", activebackground=CARD,
                  command=lambda: self._shift(-1)).pack(side="left")
        tk.Button(head, text="›", font=(F, 13), bg=CARD, fg=ACCENT,
                  relief="flat", bd=0, cursor="hand2", activebackground=CARD,
                  command=lambda: self._shift(1)).pack(side="right")
        self.lbl_month = tk.Label(head, text="", font=F_BOLD, bg=CARD, fg=TEXT)
        self.lbl_month.pack(side="left", expand=True)

        self.grid_box = tk.Frame(self, bg=CARD)
        self.grid_box.pack(padx=12)

        foot = tk.Frame(self, bg=CARD)
        foot.pack(fill="x", padx=12, pady=12)
        ttk.Button(foot, text="今天", command=self._today).pack(side="left")
        ttk.Button(foot, text="清空", command=lambda: self._pick(None)).pack(
            side="left", padx=(6, 0))
        ttk.Button(foot, text="取消", command=self.destroy).pack(side="right")

        self._render()
        self.update_idletasks()
        x = parent.winfo_rootx() + parent.winfo_width() - self.winfo_width() - 80
        y = parent.winfo_rooty() + 220
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _shift(self, months: int) -> None:
        m = self.view.month - 1 + months
        self.view = datetime(self.view.year + m // 12, m % 12 + 1, 1)
        self._render()

    def _today(self) -> None:
        self._pick(datetime.now().replace(hour=0, minute=0, second=0, microsecond=0))

    def _pick(self, d: datetime | None) -> None:
        if self.on_pick:
            self.on_pick(d)
        self.destroy()

    def _render(self) -> None:
        for w in self.grid_box.winfo_children():
            w.destroy()
        self.lbl_month.config(text=f"{self.view.year} 年 {self.view.month} 月")

        for c, d in enumerate(self.WEEK):
            tk.Label(self.grid_box, text=d, font=(F, 8), bg=CARD, fg=MUTED,
                     width=4).grid(row=0, column=c, pady=(0, 2))

        cal = calendar.Calendar(firstweekday=0)   # 周一开头，跟中文习惯一致
        today = datetime.now().date()
        for r, week in enumerate(
                cal.monthdayscalendar(self.view.year, self.view.month), start=1):
            for c, day in enumerate(week):
                if day == 0:
                    tk.Label(self.grid_box, text="", bg=CARD, width=4).grid(
                        row=r, column=c)
                    continue
                d = datetime(self.view.year, self.view.month, day)
                hot = d.date() == today
                tk.Button(
                    self.grid_box, text=str(day), font=F_SUB, width=4,
                    bg="#d7e3ff" if hot else CARD,
                    fg=ACCENT if hot else TEXT,
                    relief="flat", bd=0, cursor="hand2",
                    activebackground="#c3d6ff",
                    command=lambda dd=d: self._pick(dd),
                ).grid(row=r, column=c, padx=1, pady=1)


# --------------------------------------------------------------------------
# 主程序
# --------------------------------------------------------------------------


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.roster_path: Path | None = None
        self.folder_path: Path | None = None
        self.result: M.Result | None = None
        self.rows: list[tuple[str, str, str, str]] = []
        self.name_issues: list = []
        self.need_sid = False
        self.need_name = False
        self.order = None
        self.stats: dict = {}
        self._pending = None      # 自动核对的防抖句柄
        self.cut_kind = 0         # 被类型筛选滤掉的数量
        self.cut_time = 0         # 被时间筛选滤掉的数量
        self._running = False     # 正在核对中，别并发跑第二遍

        root.title(APP_TITLE)
        root.configure(bg=BG)
        # 别写死高度：1080 屏上 880 会把底下的结果表挤没
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w, h = 980, min(940, sh - 110)
        root.geometry(f"{w}x{h}+{(sw - w) // 2}+30")
        root.minsize(900, 700)

        self._style()
        self._build()
        self._wire_drop()

    # ---------------------------------------------------------------- 样式

    def _style(self) -> None:
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure(".", background=BG, foreground=TEXT, font=F_BODY)
        st.configure("Treeview", background=CARD, fieldbackground=CARD,
                     rowheight=30, borderwidth=0, font=F_BODY)
        st.configure("Treeview.Heading", font=F_BOLD, background="#e3e8ef",
                     foreground=TEXT, relief="flat", padding=6)
        st.map("Treeview", background=[("selected", "#d7e3ff")],
               foreground=[("selected", TEXT)])
        st.configure("TButton", font=F_BODY, padding=(12, 7))
        st.configure("Accent.TButton", font=F_BOLD, padding=(20, 10))

    # ---------------------------------------------------------------- 布局

    def _build(self) -> None:
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True, padx=18, pady=14)

        # --- 标题 ---
        head = tk.Frame(outer, bg=BG)
        head.pack(fill="x")
        tk.Label(head, text="谁没交？", font=F_TITLE, bg=BG, fg=TEXT).pack(anchor="w")
        tk.Label(
            head,
            text="左边放名单，右边放收上来的文件，点一下就知道谁还没交。",
            font=F_SUB, bg=BG, fg=MUTED,
        ).pack(anchor="w", pady=(2, 0))

        # --- 两个投放卡片 ---
        cards = tk.Frame(outer, bg=BG)
        cards.pack(fill="x", pady=(14, 0))
        cards.columnconfigure(0, weight=1, uniform="c")
        cards.columnconfigure(1, weight=1, uniform="c")

        self.card_roster = self._drop_card(
            cards, 0, "①  名单", "把 Excel 拖到这里",
            "支持 .xlsx / .xlsm；有「姓名」列就行，有「学号」列更准",
            "选择 Excel 名单…", self.pick_roster,
        )
        self.card_folder = self._drop_card(
            cards, 1, "②  收上来的文件", "把文件夹拖到这里",
            "会递归翻子文件夹；Word 临时文件、系统文件自动忽略",
            "选择文件夹…", self.pick_folder,
        )

        # --- 开始按钮 ---
        # --- 顶部状态条：条件一变就自动核对，所以这里不再放大按钮 ---
        bar = tk.Frame(outer, bg=BG)
        bar.pack(fill="x", pady=(10, 0))
        self.lbl_hint = tk.Label(bar, text="把 Excel 和文件夹拖进来，会自动开始核对",
                                 font=F_SUB, bg=BG, fg=MUTED)
        self.lbl_hint.pack(side="left")
        self.btn_run = ttk.Button(bar, text="重新核对", command=self.run_match)
        self.btn_run.pack(side="right")
        opts = tk.Frame(outer, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        opts.pack(fill="x", pady=(14, 0))
        pad = tk.Frame(opts, bg=CARD)
        pad.pack(fill="x", padx=14, pady=10)

        def _chk(parent, text, var):
            return tk.Checkbutton(
                parent, text=text, variable=var, bg=CARD, fg=TEXT, font=F_SUB,
                activebackground=CARD, activeforeground=ACCENT,
                selectcolor=CARD, bd=0, highlightthickness=0, cursor="hand2",
            )

        self.var_deep = tk.BooleanVar(value=True)
        _chk(pad, "文件名认不出时，打开文件读署名"
                  "（能救回「新建 Microsoft Word 文档.docx」这类）",
             self.var_deep).pack(anchor="w")

        r2 = tk.Frame(pad, bg=CARD)
        r2.pack(fill="x", pady=(6, 0))
        tk.Label(r2, text="要求文件名含：", font=F_SUB, bg=CARD,
                 fg=TEXT).pack(side="left")
        self.var_need_sid = tk.BooleanVar(value=False)
        self.var_need_name = tk.BooleanVar(value=False)
        _chk(r2, "学号", self.var_need_sid).pack(side="left", padx=(6, 0))
        _chk(r2, "姓名", self.var_need_name).pack(side="left", padx=(12, 0))
        tk.Label(r2, text="顺序：", font=F_SUB, bg=CARD,
                 fg=TEXT).pack(side="left", padx=(18, 0))
        self.cmb_order = ttk.Combobox(r2, width=9, state="readonly", font=F_SUB,
                                      values=["不限", "学号在前", "姓名在前"])
        self.cmb_order.current(0)
        self.cmb_order.pack(side="left", padx=(4, 0))

        r3 = tk.Frame(pad, bg=CARD)
        r3.pack(fill="x", pady=(8, 0))
        tk.Label(r3, text="只看", font=F_SUB, bg=CARD, fg=TEXT).pack(side="left")
        self.ent_since = self._date_entry(r3)
        tk.Label(r3, text="到", font=F_SUB, bg=CARD, fg=TEXT).pack(
            side="left", padx=(14, 0))
        self.ent_until = self._date_entry(r3)
        tk.Label(r3, text="之间收到的文件（可留空）",
                 font=(F, 8), bg=CARD, fg=MUTED).pack(side="left", padx=(14, 0))

        r4 = tk.Frame(pad, bg=CARD)
        r4.pack(fill="x", pady=(8, 0))
        tk.Label(r4, text="只要这些类型：", font=F_SUB, bg=CARD,
                 fg=TEXT).pack(side="left")
        self.kind_vars: dict[str, tk.BooleanVar] = {}
        for key in M.KIND_ORDER:
            var = tk.BooleanVar(value=key in M.DEFAULT_KINDS)
            self.kind_vars[key] = var
            _chk(r4, M.FILE_KINDS[key][0], var).pack(side="left", padx=(9, 0))
        tk.Label(r4, text="（都不勾 = 全部）", font=(F, 8), bg=CARD,
                 fg=MUTED).pack(side="left", padx=(10, 0))

        self.lbl_span = tk.Label(pad, text="", font=(F, 8), bg=CARD, fg=MUTED)
        self.lbl_span.pack(anchor="w", pady=(4, 0))

        # 条件一变就自动核对，这些控件都得挂上触发器
        for var in (self.var_deep, self.var_need_sid, self.var_need_name):
            var.trace_add("write", lambda *_: self.auto_run())
        for var in self.kind_vars.values():
            var.trace_add("write", lambda *_: self.auto_run())
        self.cmb_order.bind("<<ComboboxSelected>>", lambda e: self.auto_run())
        for ent in (self.ent_since, self.ent_until):
            ent.bind("<KeyRelease>", lambda e: self.auto_run())
            ent.bind("<<CalendarPicked>>", lambda e: self.auto_run())

        # --- 统计条 ---
        self.summary = tk.Frame(outer, bg=CARD, highlightbackground=LINE,
                                highlightthickness=1)
        self.summary.pack(fill="x", pady=(14, 0))
        inner = tk.Frame(self.summary, bg=CARD)
        inner.pack(fill="x", padx=16, pady=12)
        self.stat_labels = {}
        for key, label in (("submitted", "已交"), ("unsure", "待确认"),
                           ("missing", "未交"), ("unidentified", "认不出")):
            box = tk.Frame(inner, bg=CARD)
            box.pack(side="left", expand=True)
            num = tk.Label(box, text="–", font=F_BIG, bg=CARD, fg=TEXT)
            num.pack()
            tk.Label(box, text=label, font=F_SUB, bg=CARD, fg=MUTED).pack()
            self.stat_labels[key] = num
        self.verdict = tk.Label(self.summary, text="", font=F_BOLD, bg=CARD,
                                fg=MUTED, wraplength=900, justify="left")
        self.verdict.pack(fill="x", padx=16, pady=(0, 6))

        nrow = tk.Frame(self.summary, bg=CARD)
        nrow.pack(fill="x", padx=16, pady=(0, 12))
        self.lbl_names = tk.Label(nrow, text="", font=F_BOLD, bg=CARD, fg=MUTED)
        self.lbl_names.pack(side="left")
        self.btn_names = ttk.Button(nrow, text="查看清单",
                                    command=self.show_name_issues, state="disabled")
        self.btn_names.pack(side="left", padx=(10, 0))

        # --- 结果表 ---
        mid = tk.Frame(outer, bg=BG)
        mid.pack(fill="both", expand=True, pady=(14, 0))

        wrap = tk.Frame(mid, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        wrap.pack(fill="both", expand=True)

        cols = ("name", "sid", "status", "evidence")
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="browse")
        for c, t, w, anchor in (
            ("name", "姓名", 150, "w"),
            ("sid", "编号", 130, "w"),
            ("status", "状态", 100, "center"),
            ("evidence", "依据（命中的文件）", 480, "w"),
        ):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=anchor,
                             stretch=(c == "evidence"))
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")

        self.tree.tag_configure("missing", foreground=DANGER)
        self.tree.tag_configure("unsure", foreground=WARN)
        self.tree.tag_configure("submitted", foreground=OK)
        self.tree.tag_configure("unidentified", foreground=MUTED)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # --- 底部操作 ---
        foot = tk.Frame(outer, bg=BG)
        foot.pack(fill="x", pady=(12, 0))
        self.btn_copy = ttk.Button(foot, text="复制未交名单", command=self.copy_missing,
                                   state="disabled")
        self.btn_copy.pack(side="left")
        self.btn_export = ttk.Button(foot, text="导出结果 Excel", command=self.export,
                                     state="disabled")
        self.btn_export.pack(side="left", padx=(8, 0))
        self.status = tk.Label(foot, text="", font=F_SUB, bg=BG, fg=MUTED)
        self.status.pack(side="right")

    def _drop_card(self, parent, col, title, big, note, btn_text, command):
        card = tk.Frame(parent, bg=CARD, highlightbackground=LINE, highlightthickness=1)
        card.grid(row=0, column=col, sticky="nsew",
                  padx=(0, 8) if col == 0 else (8, 0))
        pad = tk.Frame(card, bg=CARD)
        pad.pack(fill="both", expand=True, padx=16, pady=14)

        tk.Label(pad, text=title, font=F_BOLD, bg=CARD, fg=TEXT).pack(anchor="w")
        zone = tk.Label(pad, text=big, font=(F, 12), bg="#f7f9fc", fg=ACCENT,
                        height=2, relief="flat")
        zone.pack(fill="x", pady=(8, 0))
        picked = tk.Label(pad, text="", font=F_SUB, bg=CARD, fg=MUTED,
                          wraplength=400, justify="left")
        picked.pack(fill="x", pady=(6, 0))
        tk.Label(pad, text=note, font=(F, 8), bg=CARD, fg=MUTED,
                 wraplength=400, justify="left").pack(fill="x", pady=(4, 8))
        ttk.Button(pad, text=btn_text, command=command).pack(anchor="w")

        card.zone = zone
        card.picked = picked
        return card

    def _date_entry(self, parent) -> tk.Entry:
        """日期输入框 + 开日历的按钮。

        之前按钮只有一个 8pt 的小三角，扔在灰底上肉眼根本找不着。
        现在给它一整块浅蓝底、够大的点击区，文字也写明白。
        """
        box = tk.Frame(parent, bg=LINE)      # 借用 1px 边框色包一圈
        box.pack(side="left", padx=(8, 0))
        inner = tk.Frame(box, bg=CARD)
        inner.pack(padx=1, pady=1)

        ent = tk.Entry(inner, width=11, font=F_BODY, relief="flat", bd=0,
                       bg=CARD, fg=TEXT, insertbackground=ACCENT)
        ent.pack(side="left", padx=(8, 6), pady=5)

        def set_date(d):
            ent.delete(0, "end")
            if d is not None:
                ent.insert(0, d.strftime("%Y-%m-%d"))
            self.auto_run()

        tk.Button(
            inner, text="日历", font=F_SUB, bg="#bed2f7", fg="#14449b",
            relief="flat", bd=0, padx=11, pady=5, cursor="hand2",
            activebackground=ACCENT, activeforeground="white",
            command=lambda: DatePicker(self.root, M.parse_date(ent.get()), set_date),
        ).pack(side="left", fill="y")
        return ent

    # ---------------------------------------------------------------- 自动核对

    def auto_run(self) -> None:
        """条件一变就自动核对 —— 带防抖，免得连点几下就跑几遍。"""
        if self._pending is not None:
            try:
                self.root.after_cancel(self._pending)
            except Exception:
                pass
        self._pending = self.root.after(400, self._run_now)

    def _run_now(self) -> None:
        self._pending = None
        if not self.roster_path or not self.folder_path:
            return          # 两样都没齐，跑了也没意义
        if self._running:
            return          # 上一遍还没跑完
        self.run_match()

    # ---------------------------------------------------------------- 拖放

    def _wire_drop(self) -> None:
        if not _HAS_DND:
            self.lbl_hint.config(
                text="提示：没装 tkinterdnd2，拖放用不了，请用每张卡片下的按钮选择。")
            return
        targets = [
            (self.root, self.on_drop_any),
            (self.card_roster, self.on_drop_roster),
            (self.card_roster.zone, self.on_drop_roster),
            (self.card_roster.picked, self.on_drop_roster),
            (self.card_folder, self.on_drop_folder),
            (self.card_folder.zone, self.on_drop_folder),
            (self.card_folder.picked, self.on_drop_folder),
        ]
        for widget, cb in targets:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", lambda e, f=cb: f(parse_drop(e.data)))
            except Exception:
                pass

    def on_drop_any(self, paths: list[str]) -> None:
        """拖到窗口空白处：按类型自动分流。"""
        if not paths:
            return
        p = Path(paths[0])
        if p.is_dir():
            self.on_drop_folder(paths)
        else:
            self.on_drop_roster(paths)

    def on_drop_roster(self, paths: list[str]) -> None:
        if not paths:
            return
        p = Path(paths[0])
        self.roster_path = p
        self.card_roster.picked.config(text=f"✓  {p.name}", fg=OK)
        self.root.after(60, self._recolor_cards)
        self.auto_run()

    def on_drop_folder(self, paths: list[str]) -> None:
        if not paths:
            return
        p = Path(paths[0])
        if p.is_file():
            p = p.parent  # 拖了个文件进来，就当他指的是那个文件夹
        self.folder_path = p
        self.card_folder.picked.config(text=f"✓  {p}", fg=OK)
        self.root.after(60, self._recolor_cards)
        self._peek_times()
        self.auto_run()

    def _peek_times(self) -> None:
        """轻量看一眼这批文件 —— **必须放后台线程**。

        以前在主线程里同步扫，拖个 C:\\Windows 进来界面会冻十几秒，
        Windows 随即弹「程序未响应」，用户一点关闭，程序就退出了。
        所以：后台扫 + 限流（quick_scan 最多取 3000 个，边遍历边停）。
        """
        if not self.folder_path:
            return
        path = self.folder_path
        self.lbl_span.config(text="正在看这批文件…", fg=MUTED)

        def work():
            try:
                real = M.locate_file_dir(path)
                files, cut = M.quick_scan(real)
                times = []
                for f in files:
                    try:
                        times.append(f.stat().st_mtime)
                    except OSError:
                        pass
                info = (real, files, cut, M.kind_counts(files), times)
            except Exception as exc:      # 后台线程里绝不能把异常漏出去
                info = exc
            try:
                self.root.after(0, self._peek_done, path, info)
            except Exception:
                pass

        threading.Thread(target=work, daemon=True).start()

    def _peek_done(self, path: Path, info) -> None:
        """回到主线程更新界面。"""
        try:
            if not self.root.winfo_exists():
                return
        except Exception:
            return
        if not isinstance(info, tuple):
            self.lbl_span.config(text="")
            return

        real, files, cut, counts, times = info
        if not times:
            self.lbl_span.config(text="")
            return

        kinds_txt = "　".join(f"{M.FILE_KINDS[k][0]} {counts[k]}"
                             for k in M.KIND_ORDER if counts[k])
        head = (f"这批文件已扫 {len(files)} 个（还有更多，没扫完）"
                if cut else f"这批文件一共 {len(files)} 个")

        lo = datetime.fromtimestamp(min(times))
        hi = datetime.fromtimestamp(max(times))
        msg = f"{head}：{kinds_txt}"
        if lo.date() != hi.date():
            msg += (f"　|　时间从 {lo:%Y-%m-%d} 到 {hi:%Y-%m-%d}，"
                    f"填上面两个框就能把别的会话滤掉")
        else:
            msg += f"　|　都是 {lo:%Y-%m-%d} 的"
        if real != self.folder_path:
            msg += f"（已自动定位到 {real.name}）"
        if cut or len(files) > 5000:
            msg += "　⚠ 文件挺多，核对会慢一点"
        self.lbl_span.config(text=msg, fg=MUTED)

    def _recolor_cards(self) -> None:
        # 子类化会接管绘制，这里只是视觉反馈，失败无所谓
        for card, ok in ((self.card_roster, self.roster_path),
                         (self.card_folder, self.folder_path)):
            try:
                card.zone.config(bg="#e8f5ec" if ok else "#f7f9fc")
            except tk.TclError:
                pass

    # ---------------------------------------------------------------- 选择

    def pick_roster(self) -> None:
        p = filedialog.askopenfilename(
            title="选择名单",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        self.on_drop_roster([p] if p else [])

    def pick_folder(self) -> None:
        p = filedialog.askdirectory(title="选择收上来的文件夹")
        self.on_drop_folder([p] if p else [])

    # ---------------------------------------------------------------- 核对

    def run_match(self) -> None:
        if not self.roster_path:
            messagebox.showwarning("还差一步", "请先选择或拖入 Excel 名单。")
            return
        if not self.folder_path:
            messagebox.showwarning("还差一步", "请先选择或拖入文件夹。")
            return

        # 下面这些必须在主线程读 —— Tk 不是线程安全的
        roster_path = self.roster_path
        folder_path = self.folder_path
        deep = self.var_deep.get()
        need_sid = self.var_need_sid.get()
        need_name = self.var_need_name.get()
        order = {"不限": None, "学号在前": "sid_first",
                 "姓名在前": "name_first"}.get(self.cmb_order.get())
        kinds = {k for k, v in self.kind_vars.items() if v.get()}
        since_text = self.ent_since.get().strip()
        until_text = self.ent_until.get().strip()
        since = M.parse_date(since_text) if since_text else None
        until = M.parse_date(until_text) if until_text else None

        for which, text, got in (("开始", since_text, since),
                                 ("结束", until_text, until)):
            if text and got is None:
                messagebox.showwarning(
                    "日期看不懂",
                    f"{which}日期「{text}」解析不了。\n\n"
                    "可以写 2026-01-03、20260103 或 2026-01，"
                    "也可以点右边的「日历」挑。")
                return
        if since and until and since > until:
            messagebox.showwarning("日期反了", "开始日期比结束日期还晚。")
            return

        self.btn_run.config(state="disabled")
        self.btn_copy.config(state="disabled")
        self.btn_export.config(state="disabled")
        self.btn_names.config(state="disabled")
        self._running = True
        self.status.config(text="正在读名单、翻文件夹…")

        def work():
            # 全是纯计算，不碰任何 Tk 对象
            try:
                out = self._compute(roster_path, folder_path, deep, need_sid,
                                    need_name, order, kinds, since, until)
            except Exception as exc:
                out = exc
            try:
                self.root.after(0, self._match_done, out, need_sid, need_name, order)
            except Exception:
                pass

        threading.Thread(target=work, daemon=True).start()

    def _compute(self, roster_path, folder_path, deep, need_sid, need_name,
                 order, kinds, since, until):
        """纯计算，跑在后台线程里。任何问题都用异常抛回去。"""
        roster = M.read_roster(roster_path)
        if not roster:
            raise ValueError("没从这份名单里找到人名。\n\n"
                             "请确认表格里有一列叫「姓名」，或者整个表格就是人名。")

        real = M.locate_file_dir(folder_path)
        files = M.collect_files(real)
        if not files:
            raise ValueError("这个文件夹里没有可用的文件。")

        total = len(files)
        cut_kind = cut_time = 0

        if kinds:
            files = M.filter_by_kind(files, kinds)
            cut_kind = total - len(files)
            if not files:
                raise ValueError("这个文件夹里没有你勾选的那几类文件。\n\n"
                                 "（类型一个都不勾就是不限制）")

        if since or until:
            before = len(files)
            files = M.filter_by_time(files, since, until)
            cut_time = before - len(files)
            if not files:
                raise ValueError("没有文件的修改时间落在这个区间里。")

        result = M.match_files(roster, files, deep=deep)

        issues, stats = [], {}
        if need_sid or need_name:
            stats = M.name_stats(roster, files)
            issues = M.check_names(roster, files, need_sid, need_name, order)

        return roster, files, result, issues, stats, total, cut_kind, cut_time

    def _match_done(self, out, need_sid, need_name, order) -> None:
        """回到主线程更新界面。"""
        self._running = False
        self.btn_run.config(state="normal")
        try:
            if not self.root.winfo_exists():
                return
        except Exception:
            return

        if isinstance(out, Exception):
            self.status.config(text="")
            if isinstance(out, ValueError):
                messagebox.showerror("没法核对", str(out))
            else:
                messagebox.showerror("出错了", f"{type(out).__name__}: {out}")
            return

        (roster, files, result, issues, stats,
         total, cut_kind, cut_time) = out
        self.result = result
        self.name_issues = issues
        self.stats = stats
        self.need_sid = need_sid
        self.need_name = need_name
        self.order = order
        self.cut_kind = cut_kind
        self.cut_time = cut_time

        if cut_kind or cut_time:
            bits = []
            if cut_kind:
                bits.append(f"类型 {cut_kind}")
            if cut_time:
                bits.append(f"时间 {cut_time}")
            self.lbl_hint.config(
                text=f"筛选：{total} 个文件 → 留下 {len(files)} 个"
                     f"（滤掉 {'、'.join(bits)} 个）")
        else:
            self.lbl_hint.config(text="改上面的条件会自动重新核对")

        self.render(roster, files)
        self.btn_copy.config(state="normal")
        self.btn_export.config(state="normal")
        self.status.config(text=f"完成 · {datetime.now():%H:%M:%S}")

    # ---------------------------------------------------------------- 展示

    def render(self, roster, files) -> None:
        res = self.result
        assert res is not None

        self.stat_labels["submitted"].config(text=str(len(res.submitted)), fg=OK)
        self.stat_labels["unsure"].config(text=str(len(res.unsure)),
                                          fg=WARN if res.unsure else MUTED)
        self.stat_labels["missing"].config(text=str(len(res.missing)),
                                           fg=DANGER if res.missing else MUTED)
        blind = len(res.unidentified)
        ratio = blind / len(files) if files else 0
        self.stat_labels["unidentified"].config(
            text=f"{blind} ({ratio:.0%})",
            fg=DANGER if ratio > 0.3 else (WARN if ratio > 0.1 else MUTED),
        )

        # 筛选会把一部分文件滤掉，让「未交」虚高 —— 必须当面说清楚，
        # 否则用户勾个「只要 Word」，交 PDF 的人就全变成未交了。
        cut_txt = ""
        if self.cut_kind or self.cut_time:
            bits = []
            if self.cut_kind:
                bits.append(f"类型滤掉 {self.cut_kind} 个")
            if self.cut_time:
                bits.append(f"时间滤掉 {self.cut_time} 个")
            cut_txt = ("\n⚠ 有筛选生效（" + "、".join(bits) + "），"
                       "未交名单里可能有人只是文件被滤掉了 —— "
                       "想确认就把筛选去掉再看一次。")

        # 结论提示：这是整个工具最该说清楚的一句话
        if ratio > 0.3:
            self.verdict.config(
                text=f"⚠  有 {ratio:.0%} 的文件认不出是谁（{blind} 个）。"
                     f"这批文件名太乱，「未交」名单不可信，请人工核对后再用。"
                     + cut_txt,
                fg=DANGER,
            )
        elif ratio > 0.1:
            self.verdict.config(
                text=f"注意：{ratio:.0%} 的文件认不出是谁（{blind} 个）。"
                     f"建议扫一眼「认不出」那些，再拿未交名单去催。" + cut_txt,
                fg=WARN,
            )
        elif res.missing:
            self.verdict.config(
                text=f"结果可信。{len(res.missing)} 人没交，"
                     f"复制名单就能直接发出去。" + cut_txt,
                fg=WARN if cut_txt else OK,
            )
        else:
            self.verdict.config(text="全员交齐 ✓" + cut_txt,
                                fg=WARN if cut_txt else OK)

        # 表格
        self.tree.delete(*self.tree.get_children())
        self.rows = []

        for p in res.missing:
            self.rows.append((p.name, p.sid, "未交", "", "missing"))
        for key, hits in sorted(res.unsure.items()):
            who = hits[0].person
            ev = "、".join(h.path.name for h in hits)
            self.rows.append((who.name, who.sid, "待确认", ev, "unsure"))
        for key, hits in sorted(res.submitted.items()):
            who = hits[0].person
            ev = "、".join(h.path.name for h in hits)
            self.rows.append((who.name, who.sid, "已交", ev, "submitted"))
        for f in res.unidentified:
            self.rows.append((f.name, "", "认不出", str(f), "unidentified"))

        for i, (name, sid, status, ev, tag) in enumerate(self.rows):
            self.tree.insert("", "end", iid=str(i),
                             values=(name, sid, status, ev), tags=(tag,))

        # 命名规范
        if self.need_sid or self.need_name:
            n = len(self.name_issues)
            st = self.stats or {}
            tail = ""
            if st.get("total"):
                tail = (f"　（这 {st['total']} 个文件里，"
                        f"含学号 {st['has_sid']}、含姓名 {st['has_name']}）")
            if n:
                self.lbl_names.config(text=f"命名规范：✕ {n} 个不符合{tail}", fg=WARN)
                self.btn_names.config(state="normal")
            else:
                self.lbl_names.config(text=f"命名规范：✓ 全部符合{tail}", fg=OK)
                self.btn_names.config(state="disabled")
        else:
            self.lbl_names.config(
                text="命名检查：勾上「要求文件名含」里的学号或姓名才会检查",
                fg=MUTED)
            self.btn_names.config(state="disabled")

    # ------------------------------------------------------------ 命名清单

    def _name_lines(self) -> list[str]:
        out = []
        for it in self.name_issues:
            who = f"{it.person.name}（{it.person.sid}）" if it.person and it.person.sid \
                else (it.person.name if it.person else "认不出是谁")
            out.append(f"{it.path.name}\t{it.why(self.need_sid, self.need_name)}\t{who}")
        return out

    def show_name_issues(self) -> None:
        if not self.name_issues:
            return
        win = tk.Toplevel(self.root)
        win.title(f"文件名不规范（{len(self.name_issues)} 个）")
        win.configure(bg=CARD)
        win.geometry("760x460")
        win.transient(self.root)

        tk.Label(win, text="这些文件的命名不符合你的要求。"
                          "右边是它对应的人 —— 认得出人的可以直接让他改名，"
                          "认不出的得先弄清楚是谁交的。",
                 font=F_SUB, bg=CARD, fg=MUTED, wraplength=700,
                 justify="left").pack(anchor="w", padx=16, pady=(14, 8))

        body = tk.Frame(win, bg=CARD)
        body.pack(fill="both", expand=True, padx=16)
        cols = ("file", "why", "who")
        tree = ttk.Treeview(body, columns=cols, show="headings")
        for c, t, w in (("file", "文件名", 340), ("why", "问题", 100),
                        ("who", "对应的人", 180)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor="w")
        vs = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vs.set)
        tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        for i, line in enumerate(self._name_lines()):
            f, why, who = line.split("\t", 2)
            tree.insert("", "end", iid=str(i), values=(f, why, who))

        btns = tk.Frame(win, bg=CARD)
        btns.pack(fill="x", padx=16, pady=12)

        def copy_it():
            text = "\n".join(self._name_lines()).replace("\t", "  ")
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update()
            messagebox.showinfo("已复制", f"{len(self.name_issues)} 条已复制到剪贴板。")

        def export_it():
            path = filedialog.asksaveasfilename(
                title="导出命名不规范清单", defaultextension=".xlsx",
                initialfile=f"命名不规范_{datetime.now():%Y%m%d_%H%M}.xlsx",
                filetypes=[("Excel 文件", "*.xlsx")])
            if not path:
                return
            try:
                from openpyxl import Workbook
                wb = Workbook()
                ws = wb.active
                ws.title = "命名不规范"
                ws.append(["文件名", "问题", "对应的人"])
                for line in self._name_lines():
                    ws.append(line.split("\t"))
                for col, w in (("A", 46), ("B", 14), ("C", 22)):
                    ws.column_dimensions[col].width = w
                wb.save(path)
            except Exception as exc:
                messagebox.showerror("导出失败", str(exc))
                return
            messagebox.showinfo("导出成功", f"已保存到：\n{path}")

        ttk.Button(btns, text="复制清单", command=copy_it).pack(side="left")
        ttk.Button(btns, text="导出 Excel", command=export_it).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="关掉", command=win.destroy).pack(side="right")

    def on_select(self, _event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if idx >= len(self.rows):
            return
        _, _, status, ev, tag = self.rows[idx]

        if tag == "missing":
            tip = "没有找到任何对得上这个人的文件。"
            if self.result and self.result.unidentified:
                tip += (f"\n\n提醒：有 {len(self.result.unidentified)} 个文件认不出是谁，"
                        f"去「认不出」那几条看看，别漏了。")
            detail = tip
        elif tag == "unsure":
            detail = ("文件名匹配上了，但这个人在名单里不是唯一的"
                      "（重名），所以不敢确定。\n\n"
                      "去打开这些文件看一眼就知道是谁了：\n"
                      + "\n".join(f"  · {n}" for n in ev.split("、")))
        elif tag == "submitted":
            detail = "判定依据（命中的文件）：\n" + "\n".join(
                f"  · {n}" for n in ev.split("、"))
        else:
            detail = ("这个文件名里找不到任何名单上的人。\n\n"
                      "它可能是一份作业，也可能是无关文件。\n"
                      f"位置：\n  {ev}")

        self._show_detail(detail)

    def _show_detail(self, text: str) -> None:
        win = getattr(self, "_detail_win", None)
        if win is not None and win.winfo_exists():
            win.destroy()
        win = tk.Toplevel(self.root)
        self._detail_win = win
        win.title("判定依据")
        win.configure(bg=CARD)
        win.geometry("620x320")
        win.transient(self.root)
        txt = tk.Text(win, wrap="word", font=F_BODY, bg=CARD, fg=TEXT,
                      relief="flat", padx=16, pady=14)
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", text)
        txt.config(state="disabled")
        ttk.Button(win, text="关掉", command=win.destroy).pack(pady=(0, 12))

    # ---------------------------------------------------------------- 输出

    def copy_missing(self) -> None:
        if not self.result:
            return
        if not self.result.missing:
            messagebox.showinfo("不用复制", "没人缺交。")
            return
        # 有学号就带上，重名时老师能对上人
        lines = []
        for p in self.result.missing:
            lines.append(f"{p.name}（{p.sid}）" if p.sid else p.name)
        text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        self.status.config(text=f"已复制 {len(lines)} 人到剪贴板")
        messagebox.showinfo("已复制", f"{len(lines)} 人已复制到剪贴板：\n\n{text}")

    def export(self) -> None:
        if not self.result:
            return
        default = f"核对结果_{datetime.now():%Y%m%d_%H%M}.xlsx"
        path = filedialog.asksaveasfilename(
            title="导出结果", defaultextension=".xlsx",
            initialfile=default, filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            write_result_xlsx(Path(path), self.rows, self.result.missing)
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
            return
        self.status.config(text=f"已导出 {Path(path).name}")
        messagebox.showinfo("导出成功", f"已保存到：\n{path}")


# --------------------------------------------------------------------------
# 无头自测：不弹窗，只验业务逻辑
# --------------------------------------------------------------------------


def _sample_dir() -> Path | None:
    """找「模拟数据」。打包成 exe 后 __file__ 指向临时解压目录，
    所以还要额外看当前工作目录和 exe 所在目录。"""
    cands = []
    if not getattr(sys, "frozen", False):
        cands.append(Path(__file__).resolve().parent / "模拟数据")
    cands.append(Path.cwd() / "模拟数据")
    cands.append(Path(sys.executable).resolve().parent / "模拟数据")
    for c in cands:
        if c.is_dir():
            return c
    return None


def selftest() -> int:
    """自测。既打印，也写一份到临时文件 ——
    --noconsole 打包出来的 exe 没有 stdout，只能靠那个文件看结果。"""
    import tempfile

    lines: list[str] = []

    def say(s: str = "") -> None:
        lines.append(s)
        try:
            print(s)
        except Exception:
            pass

    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    report = Path(tempfile.gettempdir()) / "谁没交_selftest.txt"

    def dump() -> None:
        try:
            report.write_text("\n".join(lines), encoding="utf-8")
        except Exception:
            pass

    base = _sample_dir()
    if base is None:
        say("找不到「模拟数据」目录。把 exe 和 模拟数据 放在同一层再跑。")
        dump()
        return 1

    bad = 0
    cases = [
        ("场景A_大学班级收作业", "全班名单.xlsx", "收到的作业", 46, 0, 6, 3),
        ("场景B_公司收身份证扫描件", "员工名单.xlsx", "收到的材料", 26, 1, 3, 0),
        ("场景C_微信群收报名材料", "报名名单.xlsx", "收到的文件", 22, 0, 3, 9),
        ("场景D_拼音命名", "名单.xlsx", "收到的作业", 17, 0, 3, 0),
    ]
    for folder, roster_name, files_name, s, u, m, x in cases:
        d = base / folder
        roster = M.read_roster(d / roster_name)
        files = M.collect_files(d / files_name)
        r = M.match_files(roster, files, deep=True)
        got = (len(r.submitted), len(r.unsure), len(r.missing), len(r.unidentified))
        ok = got == (s, u, m, x)
        bad += 0 if ok else 1
        say(f"  {'PASS' if ok else 'FAIL'}  {folder}: {got}"
            f"{'' if ok else f'  预期 {(s, u, m, x)}'}")

    # 时间过滤 + 命名规范检查
    say()
    say("时间过滤 / 命名检查：")
    try:
        d = base / "场景A_大学班级收作业"
        roster = M.read_roster(d / "全班名单.xlsx")
        allf = M.collect_files(d / "收到的作业")

        kept = M.filter_by_time(allf, M.parse_date("2026-01-03"),
                                M.parse_date("2026-01-04"))
        r = M.match_files(roster, kept, deep=True)
        ok1 = len(kept) == len(allf) - 2 and len(r.unidentified) == 1
        say(f"  {'PASS' if ok1 else 'FAIL'}  时间过滤：{len(allf)} → {len(kept)} 个，"
            f"认不出 {len(r.unidentified)}（预期 1）")
        bad += 0 if ok1 else 1

        issues = M.check_names(roster, allf, need_sid=True, need_name=False)
        ok2 = len(issues) == len(allf) - 2
        say(f"  {'PASS' if ok2 else 'FAIL'}  命名检查：{len(issues)} 个不合规（缺学号），"
            f"预期 {len(allf) - 2}")
        bad += 0 if ok2 else 1

        # 类型筛选：一个不勾 = 不限制；只勾 Word = 只剩 docx
        counts = M.kind_counts(allf)
        say("  类型分布：" + "　".join(
            f"{M.FILE_KINDS[k][0]} {counts[k]}"
            for k in M.KIND_ORDER if counts[k]))
        kept_word = M.filter_by_kind(allf, {"word"})
        ok3 = len(kept_word) == counts["word"] and counts["word"] > 0
        say(f"  {'PASS' if ok3 else 'FAIL'}  只勾 Word：{len(allf)} → "
            f"{len(kept_word)}（应该全是 .docx，共 {counts['word']} 个）")
        bad += 0 if ok3 else 1
    except Exception as exc:
        say(f"  FAIL  抛异常：{type(exc).__name__}: {exc}")
        bad += 1

    # 导出也要能跑，不然交付物是残的
    say()
    say("导出测试：")
    try:
        from openpyxl import load_workbook

        d = base / "场景A_大学班级收作业"
        r = M.match_files(M.read_roster(d / "全班名单.xlsx"),
                          M.collect_files(d / "收到的作业"), deep=True)
        rows = [(p.name, p.sid, "未交", "", "missing") for p in r.missing]
        for key, hits in sorted(r.unsure.items()):
            who = hits[0].person
            rows.append((who.name, who.sid, "待确认",
                         "、".join(h.path.name for h in hits), "unsure"))

        out = Path(tempfile.gettempdir()) / "dsh_export_probe.xlsx"
        write_result_xlsx(out, rows, r.missing)
        wb = load_workbook(out)
        n1, n2 = wb["核对结果"].max_row - 1, wb["未交名单"].max_row - 1
        ok = n1 == len(rows) and n2 == len(r.missing)
        say(f"  {'PASS' if ok else 'FAIL'}  核对结果 {n1} 行 / 未交名单 {n2} 行")
        bad += 0 if ok else 1
        out.unlink(missing_ok=True)
    except Exception as exc:
        say(f"  FAIL  导出抛异常：{type(exc).__name__}: {exc}")
        bad += 1

    say()
    say("全部通过" if not bad else f"{bad} 项不符")
    say(f"（报告写到 {report}）")
    dump()
    return bad


def main() -> None:
    # 兜底：任何没被捕获的异常都留个记录。
    # --noconsole 打包后连 stderr 都没有，不写文件就什么都看不见。
    def hook(exc_type, exc, tb):
        import traceback
        _write_crash("".join(traceback.format_exception(exc_type, exc, tb)))

    sys.excepthook = hook
    try:
        threading.excepthook = lambda a: hook(a.exc_type, a.exc_value,
                                               a.exc_traceback)
    except Exception:
        pass

    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 高分屏不糊
        except Exception:
            pass
    root = make_root()
    App(root)
    root.mainloop()


def droptest() -> int:
    """验证拖放链路。

    tkinterdnd2 的拖放由 Tcl 扩展在 C 层处理，没法从外面投消息模拟，
    所以这里测的是「拖放数据解析 + 落地处理」这两段 ——
    真正的事件投递只能靠人手拖一次。
    """
    import tempfile

    lines: list[str] = []

    def say(s: str = "") -> None:
        lines.append(s)
        try:
            print(s)
        except Exception:
            pass

    report = Path(tempfile.gettempdir()) / "谁没交_droptest.txt"
    if getattr(sys, "frozen", False):
        report = Path.cwd() / "谁没交_droptest.txt"

    bad = 0

    # --- 1. 拖放数据解析 ---
    say("拖放数据解析：")
    cases = [
        ("D:/a/b.docx", ["D:/a/b.docx"]),
        ("{D:/我的 作业/张三.docx}", ["D:/我的 作业/张三.docx"]),
        ("{D:/有 空 格/a.docx} D:/b.zip",
         ["D:/有 空 格/a.docx", "D:/b.zip"]),
        ("D:/x.docx D:/y.docx", ["D:/x.docx", "D:/y.docx"]),
        ("", []),
    ]
    for raw, want in cases:
        got = parse_drop(raw)
        ok = got == want
        bad += 0 if ok else 1
        say(f"  {'PASS' if ok else 'FAIL'} {raw!r} -> {got}")

    # --- 2. 落地处理 ---
    say("落地处理：")
    try:
        if sys.platform == "win32":
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    base = _sample_dir()
    target = (base / "场景A_大学班级收作业" / "收到的作业") if base else None

    root = make_root()
    a = App(root)
    root.attributes("-topmost", True)
    say(f"  拖放可用：{_HAS_DND}")

    if target is None or not target.exists():
        say("  找不到模拟数据，跳过")
        report.write_text("\n".join(lines), encoding="utf-8")
        root.destroy()
        return 1

    state = {"i": 0}
    steps = [("文件夹 -> 右侧卡片", a.on_drop_folder, str(target), "folder_path"),
             ("文件夹 -> 窗口空白", a.on_drop_any, str(target), "folder_path")]

    def next_step():
        if state["i"] >= len(steps):
            return finish()
        name, fn, arg, attr = steps[state["i"]]
        a.folder_path = None
        a.roster_path = None
        say(f"  {name}…")
        fn([arg])
        root.after(1200, lambda: check(name, attr))

    def check(name, attr):
        got = getattr(a, attr)
        ok = got is not None
        state["i"] += 1
        nonlocal_bad[0] += 0 if ok else 1
        say(f"    {'PASS' if ok else 'FAIL'} -> {got}")
        root.after(80, next_step)

    def finish():
        say()
        total_bad = bad + nonlocal_bad[0]
        say("全部通过" if not total_bad else f"{total_bad} 项失败")
        report.write_text("\n".join(lines), encoding="utf-8")
        say(f"（报告写到 {report}）")
        root.after(150, root.destroy)

    nonlocal_bad = [0]
    root.after(300, next_step)
    root.mainloop()
    return bad + nonlocal_bad[0]


def _write_crash(text: str) -> None:
    """打包成 --noconsole 之后，崩溃是彻底看不见的，至少留个文件。"""
    try:
        import tempfile
        (Path(tempfile.gettempdir()) / "谁没交_crash.txt").write_text(
            text, encoding="utf-8")
    except Exception:
        pass


if __name__ == "__main__":
    try:
        if "--selftest" in sys.argv:
            raise SystemExit(selftest())
        if "--droptest" in sys.argv:
            raise SystemExit(droptest())
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        _write_crash(traceback.format_exc())
        raise
