# -*- coding: utf-8 -*-
"""
DLP 批量解密工具 - 图形界面（企业极简风格）
拖入文件/文件夹 → 开始解密 → 解密副本保存到明文区 C:\解密文件
设计方向：Enterprise Minimal（近黑主色、语义状态、8px 网格、克制留白）
"""
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import decrypt_core as dc

# ---------- 设计令牌 ----------
PAGE_BG = "#F7F8FA"      # 页面背景
SURFACE_BG = "#FFFFFF"   # 卡片/表面
HOVER_BG = "#F3F4F6"     # 悬浮/选中
BORDER_C = "#E5E7EB"     # 描边
TEXT_1 = "#111827"       # 主文字
TEXT_2 = "#6B7280"       # 次要文字
TEXT_3 = "#9CA3AF"       # 弱文字
ACCENT = "#111827"       # 主操作色（近黑）
ACCENT_HOVER = "#1F2937"
ACCENT_PRESS = "#0B0F19"

# 语义状态（浅底 + 深字）
OK_BG, OK_FG = "#EAF3DE", "#3B6D11"
RUN_BG, RUN_FG = "#FAEEDA", "#854F0B"
FAIL_BG, FAIL_FG = "#FCEBEB", "#A32D2D"
IDLE_BG, IDLE_FG = "#F3F4F6", "#6B7280"

# 文件类型 → (图标字母, 品牌色)
TYPE_STYLE = {
    "excel": ("X", "#3B6D11"),
    "word": ("W", "#185FA5"),
    "ppt": ("P", "#854F0B"),
    "pdf": ("F", "#A32D2D"),
}

FONT = "Microsoft YaHei UI"
FONT_MONO = "Consolas"

# ---------- 首次启动引导 ----------
# 标记文件：存在 = 用户已完成初始配置确认，不再自动弹出引导
GUIDE_FLAG_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "DLPDecryptTool")
GUIDE_FLAG_FILE = os.path.join(GUIDE_FLAG_DIR, "initialized.flag")

GUIDE_TEXT = (
    "首次使用本工具前，请先完成一次初始配置（只需一次）：\n"
    "\n"
    "第 1 步　使用绿盾正常流程申请解密一个文件，申请成功后保存。\n"
    "\n"
    "第 2 步　导出方式选择「另存为」，保存位置指定为：\n"
    "        C:\\解密文件\n"
    "        （该位置无需手动创建，绿盾会自动创建。）\n"
    "\n"
    "第 3 步　完成上述配置后，C:\\解密文件 即为本工具后续的解密目录，\n"
    "        之后即可正常使用本工具进行批量解密。"
)


def _is_first_run():
    return not os.path.exists(GUIDE_FLAG_FILE)


def _mark_initialized():
    try:
        os.makedirs(GUIDE_FLAG_DIR, exist_ok=True)
        with open(GUIDE_FLAG_FILE, "w", encoding="utf-8") as f:
            f.write("initialized")
    except Exception:
        pass


# ---------- 单实例 ----------
# 锁文件：保证全系统只有一个主窗口（右键多个文件也只开一个窗口）
# 用"独占创建锁文件 + 写 PID + 检测 PID 存活"实现，纯标准库，打包零依赖风险
LOCK_FILE = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")),
                         "dlp_decrypt.lock")
# 队列文件：第二实例把文件路径写进来，主实例轮询拾取（实现"转发给已开窗口"）
QUEUE_FILE = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")),
                          "dlp_decrypt_queue.txt")
_MAIN_LOCK = False  # 主实例持有的锁标记（退出时据此释放锁文件）


def _pid_alive(pid):
    """检查 PID 是否存活（Windows，OpenProcess + GetExitCodeProcess）。"""
    try:
        import ctypes
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        STILL_ACTIVE = 259
        h = ctypes.windll.kernel32.OpenProcess(
            PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not h:
            return False
        try:
            code = ctypes.c_ulong()
            ctypes.windll.kernel32.GetExitCodeProcess(h, ctypes.byref(code))
            return code.value == STILL_ACTIVE
        finally:
            ctypes.windll.kernel32.CloseHandle(h)
    except Exception:
        return True  # 无法判断时保守认为存活


def _is_main_instance():
    """
    锁文件方式获取单实例。
    返回 None = 已有实例在运行；返回 True = 本进程是主实例。
    主实例异常退出会残留锁文件——启动时检测 PID 不存活则自动清理重试。
    """
    for _ in range(2):
        try:
            fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, str(os.getpid()).encode("ascii"))
            finally:
                os.close(fd)
            return True
        except FileExistsError:
            alive = True
            try:
                with open(LOCK_FILE, "r", encoding="ascii") as f:
                    pid = int(f.read().strip() or "0")
                alive = _pid_alive(pid) if pid > 0 else False
            except Exception:
                alive = True
            if alive:
                return None  # 已有实例存活
            # 残留锁：PID 不存在，删除后重试
            try:
                os.remove(LOCK_FILE)
            except Exception:
                return None
        except Exception:
            return True  # 拿不到锁不阻塞（宽松模式）
    return None


def _enqueue_paths(paths):
    """把文件路径追加到队列文件，供主实例轮询拾取。"""
    try:
        with open(QUEUE_FILE, "a", encoding="utf-8") as f:
            for p in paths:
                f.write(p + "\n")
    except Exception:
        pass


STATUS_META = {  # 状态 → 元信息文案
    "idle": "等待处理",
    "run": "正在解密",
    "ok": "已解密",
    "fail": "解密失败",
}


def human_size(path):
    try:
        b = os.path.getsize(path)
    except OSError:
        return "—"
    if b < 1024:
        return f"{b} B"
    if b < 1024 * 1024:
        return f"{b / 1024:.0f} KB"
    return f"{b / 1024 / 1024:.1f} MB"


class DecryptApp:
    def __init__(self, root):
        self.root = root
        root.title("DLP 解密工具")
        root.geometry("780x620")
        root.minsize(720, 540)
        root.configure(bg=PAGE_BG)

        # 拖拽支持（tkinterdnd2 可选）
        try:
            import tkinterdnd2  # noqa: F401
            self.dnd_available = True
        except Exception:
            self.dnd_available = False

        self.files = []           # 待处理文件列表
        self.running = False
        self._stop = None         # 解密停止信号（threading.Event）
        self._phase = ""          # 当前进度阶段：prepare / decrypt
        self._selected = set()    # 选中的文件下标
        self._rows = []           # 行容器列表
        self._badges = {}         # 下标 → 状态徽章
        self._idx_by_path = {}    # 绝对路径 → 下标
        self.log_visible = False

        self._setup_style()
        try:
            root.iconphoto(True, self._make_icon())
        except Exception:
            pass
        self._build_ui()
        self._check_env()
        self._bind_shortcuts()
        # 单实例队列轮询：拾取"第二实例转发来的文件"并自动解密
        self.root.after(600, self._poll_queue)

    # ---------- 主题 ----------
    def _setup_style(self):
        s = ttk.Style(self.root)
        try:
            s.theme_use("clam")
        except Exception:
            pass
        s.configure("Primary.TButton", background=ACCENT, foreground="#FFFFFF",
                    borderwidth=0, focusthickness=0, padding=(16, 7),
                    font=(FONT, 10))
        s.map("Primary.TButton",
              background=[("active", ACCENT_HOVER), ("pressed", ACCENT_PRESS),
                          ("disabled", "#9CA3AF")],
              foreground=[("disabled", "#F3F4F6")])
        s.configure("Secondary.TButton", background=SURFACE_BG,
                    foreground=TEXT_1, bordercolor=BORDER_C, lightcolor=BORDER_C,
                    darkcolor=BORDER_C, borderwidth=1, focusthickness=0,
                    padding=(14, 7), font=(FONT, 10))
        s.map("Secondary.TButton",
              background=[("active", HOVER_BG), ("pressed", "#E9EAEC"),
                          ("disabled", "#F9FAFB")],
              foreground=[("disabled", TEXT_3)])
        s.configure("Ghost.TButton", background=PAGE_BG, foreground=TEXT_2,
                    borderwidth=0, focusthickness=0, padding=(8, 7),
                    font=(FONT, 10))
        s.map("Ghost.TButton",
              background=[("active", HOVER_BG), ("pressed", "#E9EAEC")],
              foreground=[("disabled", TEXT_3)])
        s.configure("TProgressbar", troughcolor="#F1EFE8", background=ACCENT,
                    borderwidth=0, thickness=6)
        s.configure("Vertical.TScrollbar", background="#D8DADF",
                    troughcolor=SURFACE_BG, bordercolor=SURFACE_BG,
                    arrowcolor=TEXT_2, width=10)

    def _make_icon(self):
        """生成 32x32 应用图标（近黑底 + 白色盾牌）。"""
        img = tk.PhotoImage(width=32, height=32)
        img.put(ACCENT, to=(0, 0, 31, 31))
        # 白色盾牌
        for y in range(6, 27):
            for x in range(8, 24):
                img.put("#FFFFFF", (x, y))
        img.put(ACCENT, to=(12, 10, 20, 12))   # 盾牌缺口（钥匙孔）
        img.put(ACCENT, to=(14, 12, 18, 17))
        img.put(ACCENT, to=(13, 17, 19, 18))
        return img

    def _draw_logo(self, parent):
        c = tk.Canvas(parent, width=30, height=30, bg=ACCENT,
                      highlightthickness=0)
        c.pack()
        c.create_polygon(15, 6, 24, 10, 24, 16, 15, 24, 6, 16, 6, 10,
                         fill="", outline="#FFFFFF", width=1.6)
        c.create_line(12, 14, 18, 14, fill="#FFFFFF", width=1.6)
        c.create_rectangle(13.2, 14, 16.8, 19, fill="#FFFFFF", outline="")

    # ---------- UI ----------
    def _build_ui(self):
        self._header()
        self._toolbar()
        self._list_area()
        self._footer()
        self._log_area()

    def _header(self):
        hdr = tk.Frame(self.root, bg=PAGE_BG)
        hdr.pack(fill="x", padx=18, pady=(14, 12))
        logo = tk.Frame(hdr, width=30, height=30, bg=ACCENT)
        logo.pack(side="left")
        logo.pack_propagate(False)
        self._draw_logo(logo)
        t = tk.Frame(hdr, bg=PAGE_BG)
        t.pack(side="left", padx=(12, 0))
        tk.Label(t, text="DLP 解密工具", bg=PAGE_BG, fg=TEXT_1,
                 font=(FONT, 12, "bold")).pack(anchor="w")
        tk.Label(t, text="批量解密 · 输出到 C:\\解密文件", bg=PAGE_BG, fg=TEXT_3,
                 font=(FONT, 9)).pack(anchor="w")
        self.pill = tk.Frame(hdr, bg=HOVER_BG)
        self.pill.pack(side="right")
        self.pill_dot = tk.Frame(self.pill, width=7, height=7, bg="#9CA3AF")
        self.pill_dot.pack(side="left", padx=(9, 0), pady=6)
        self.pill_text = tk.Label(self.pill, text="就绪", bg=HOVER_BG,
                                  fg=TEXT_2, font=(FONT, 9))
        self.pill_text.pack(side="left", padx=(6, 10), pady=4)

    def _set_pill(self, state):
        m = {"ready": ("就绪", "#9CA3AF"),
             "running": ("解密中", "#F59E0B"),
             "done": ("完成", "#10B981"),
             "fail": ("完成", "#EF4444")}
        text, color = m.get(state, m["ready"])
        self.pill_text.config(text=text)
        self.pill_dot.config(bg=color)

    def _toolbar(self):
        bar = tk.Frame(self.root, bg=PAGE_BG)
        bar.pack(fill="x", padx=18, pady=(0, 10))
        # 主操作（左对齐）
        self.btn_start = self._btn(bar, "Primary.TButton", "开始解密", self.start_decrypt)
        self.btn_start.pack(side="left", padx=(0, 8))
        self.btn_stop = self._btn(bar, "Secondary.TButton", "停止", self.stop_decrypt)
        self.btn_stop.state(["disabled"])
        self.btn_stop.pack(side="left", padx=(0, 8))
        self._btn(bar, "Secondary.TButton", "打开明文区", self.open_plaintext_dir).pack(side="left", padx=(0, 8))
        # 文件管理（加细分隔，与主操作区分）
        self._btn(bar, "Ghost.TButton", "添加文件", self.add_files).pack(side="left", padx=(16, 0))
        self._btn(bar, "Ghost.TButton", "添加文件夹", self.add_folder).pack(side="left")
        self._btn(bar, "Ghost.TButton", "移除选中", self.remove_selected).pack(side="left")
        self._btn(bar, "Ghost.TButton", "清空", self.clear_list).pack(side="left")
        # 右：统计 + 右上角的设置入口（首次说明 / 注册卸载右键菜单）
        self.lbl_count = tk.Label(bar, text="共 0 个文件", bg=PAGE_BG,
                                  fg=TEXT_3, font=(FONT, 9))
        self.lbl_count.pack(side="right", padx=(8, 0))
        self.btn_settings = self._btn(bar, "Ghost.TButton", "⋯",
                                      self._open_settings_menu)
        self.btn_settings.pack(side="right")

    def _open_settings_menu(self):
        """设置菜单（首次使用说明 / 注册卸载右键 / 打开明文区）"""
        menu = tk.Menu(self.root, tearoff=0, bg=SURFACE_BG, fg=TEXT_1,
                       activebackground=HOVER_BG, activeforeground=TEXT_1,
                       relief="flat", bd=1, font=(FONT, 9))
        menu.add_command(label="  首次使用说明", command=self.show_first_run_guide)
        menu.add_command(label="  注册右键菜单", command=self.register_menu)
        menu.add_command(label="  卸载右键菜单", command=self.unregister_menu)
        menu.add_separator()
        menu.add_command(label="  打开明文区", command=self.open_plaintext_dir)
        try:
            x = self.btn_settings.winfo_rootx()
            y = self.btn_settings.winfo_rooty() + self.btn_settings.winfo_height() + 2
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _btn(self, parent, style, text, cmd):
        return ttk.Button(parent, text=text, command=cmd, style=style)

    def _list_area(self):
        # 外层卡片（带描边感）
        card = tk.Frame(self.root, bg=BORDER_C)
        card.pack(fill="both", expand=True, padx=18, pady=(2, 2))
        # 内层（去掉 1px 描边）
        inner = tk.Frame(card, bg=SURFACE_BG)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        # 用 grid 布局 canvas + 滚动条（grid 比 pack 在主体+滚动条场景更稳定）
        inner.grid_columnconfigure(0, weight=1)
        inner.grid_rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(inner, bg=SURFACE_BG, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(inner, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        # list_host 嵌入 canvas 的窗口项
        self.list_host = tk.Frame(self.canvas, bg=SURFACE_BG)
        self._win_id = self.canvas.create_window(
            (0, 0), window=self.list_host, anchor="nw", width=100)

        def _sync_width(event=None):
            # 把 canvas 当前宽度同步给 list_host，保证横向上撑满
            try:
                w = self.canvas.winfo_width()
                if w > 1:
                    self.canvas.itemconfigure(self._win_id, width=w)
            except Exception:
                pass

        self.canvas.bind("<Configure>", _sync_width)
        self.list_host.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        # 初始强制同步（处理 pack/grid 完成前的尺寸未就绪）
        self.root.after_idle(_sync_width)
        self.root.after(60, _sync_width)

        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Button-1>", lambda e: self._clear_selection())
        if self.dnd_available:
            self._enable_dnd()
        self._empty = None
        self._refresh_list()

    def _on_wheel(self, event):
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

    def _enable_dnd(self):
        try:
            from tkinterdnd2 import DND_FILES
            for w in (self.canvas, self.root):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass

    def _on_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        self._add_paths(paths)

    # ---------- 首次启动引导 ----------
    def show_first_run_guide(self):
        """首次启动引导弹窗（模态）。用户确认完成后写入标记，之后不再自动弹出。"""
        win = tk.Toplevel(self.root)
        win.title("首次使用 · 初始配置引导")
        win.configure(bg=SURFACE_BG)
        win.resizable(False, False)
        win.transient(self.root)

        head = tk.Frame(win, bg=SURFACE_BG)
        head.pack(fill="x", padx=26, pady=(24, 6))
        tk.Label(head, text="欢迎使用 DLP 批量解密工具", bg=SURFACE_BG, fg=TEXT_1,
                 font=(FONT, 13, "bold")).pack(anchor="w")
        tk.Label(head, text="检测到这是首次启动，请按以下步骤完成初始配置：",
                 bg=SURFACE_BG, fg=TEXT_2, font=(FONT, 10)).pack(anchor="w",
                                                                 pady=(4, 0))

        body = tk.Frame(win, bg=SURFACE_BG)
        body.pack(fill="x", padx=26, pady=(6, 2))
        tk.Label(body, text=GUIDE_TEXT, bg=SURFACE_BG, fg=TEXT_1,
                 font=(FONT, 10), justify="left", anchor="w").pack(fill="x")
        tk.Label(win, text="提示：说明内容可随时在右上角「⋯ → 首次使用说明」查看。",
                 bg=SURFACE_BG, fg=TEXT_3, font=(FONT, 8)).pack(anchor="w",
                                                                padx=26, pady=(2, 0))

        btns = tk.Frame(win, bg=SURFACE_BG)
        btns.pack(fill="x", padx=26, pady=(14, 22))

        def _done():
            _mark_initialized()
            win.destroy()

        ttk.Button(btns, text="我已完成配置，开始使用", style="Primary.TButton",
                   command=_done).pack(side="right")
        ttk.Button(btns, text="打开明文区", style="Secondary.TButton",
                   command=self.open_plaintext_dir).pack(side="left")

        # 居中于主窗口
        win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - win.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - win.winfo_height()) // 3
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        win.grab_set()
        self.root.wait_window(win)

    # ---------- 文件列表 ----------
    def _refresh_list(self):
        for r in self._rows:
            r.destroy()
        self._rows.clear()
        self._badges.clear()
        self._selected.clear()
        self._idx_by_path = {os.path.abspath(f): i for i, f in enumerate(self.files)}
        if self._empty is not None:
            self._empty.destroy()
            self._empty = None
        if not self.files:
            self._show_empty()
        else:
            for i, f in enumerate(self.files):
                self._make_row(i, f)
        self.lbl_count.config(text=f"共 {len(self.files)} 个文件")
        self.canvas.yview_moveto(0)

    def _show_empty(self):
        e = tk.Frame(self.list_host, bg=SURFACE_BG)
        e.pack(fill="both", expand=True, pady=60)
        tk.Label(e, text="＋", bg=SURFACE_BG, fg="#E5E7EB",
                 font=(FONT, 42)).pack()
        tk.Label(e, text="拖入文件或文件夹，或点击上方按钮", bg=SURFACE_BG,
                 fg=TEXT_3, font=(FONT, 10)).pack(pady=(4, 0))
        self._empty = e

    def _make_row(self, i, path):
        row = tk.Frame(self.list_host, bg=SURFACE_BG, cursor="hand2")
        row.pack(fill="x", padx=6, pady=2)
        ftype = dc.get_office_type(path)
        letter, color = TYPE_STYLE.get(ftype, ("?", "#6B7280"))
        parent = os.path.basename(os.path.dirname(path)) or path

        badge = tk.Label(row, text="待处理", bg=IDLE_BG, fg=IDLE_FG,
                         font=(FONT, 9), padx=9, pady=2)
        badge.pack(side="right")
        meta = tk.Label(row, text=f"{human_size(path)} · {parent}",
                        bg=SURFACE_BG, fg=TEXT_3, font=(FONT, 9))
        meta.pack(side="right", padx=(0, 8))
        ic = tk.Label(row, text=letter, width=3, bg=color, fg="#FFFFFF",
                      font=(FONT, 9, "bold"))
        ic.pack(side="left", padx=(4, 10), pady=4)
        name = tk.Label(row, text=os.path.basename(path), bg=SURFACE_BG,
                        fg=TEXT_1, font=(FONT, 10), anchor="w")
        name.pack(side="left", fill="x", expand=True)

        self._badges[i] = badge
        for w in (row, ic, name, meta, badge):
            w.bind("<Button-1>", lambda e, idx=i: self._toggle_select(idx))
        self._rows.append(row)

    def _toggle_select(self, idx):
        if idx in self._selected:
            self._selected.discard(idx)
            bg = SURFACE_BG
        else:
            self._selected.add(idx)
            bg = HOVER_BG
        row = self._rows[idx]
        for child in row.winfo_children():
            if isinstance(child, tk.Label) and child.cget("bg") in (SURFACE_BG, HOVER_BG):
                child.config(bg=bg)
        row.config(bg=bg)

    def _clear_selection(self):
        for idx in list(self._selected):
            self._selected.discard(idx)
            row = self._rows[idx]
            row.config(bg=SURFACE_BG)
            for child in row.winfo_children():
                if isinstance(child, tk.Label) and child.cget("bg") == HOVER_BG:
                    child.config(bg=SURFACE_BG)

    def _set_status(self, idx, status):
        badge = self._badges.get(idx)
        if badge is None:
            return
        cfg = {"idle": (IDLE_BG, IDLE_FG, "待处理"),
               "run": (RUN_BG, RUN_FG, "解密中"),
               "ok": (OK_BG, OK_FG, "成功"),
               "fail": (FAIL_BG, FAIL_FG, "失败")}[status]
        badge.config(bg=cfg[0], fg=cfg[1], text=cfg[2])

    # ---------- 文件操作 ----------
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择要解密的文件",
            filetypes=[("支持的文件", "*.xls *.xlsx *.xlsm *.xlsb *.xltx *.xlt "
                                     "*.doc *.docx *.docm *.dotx *.dotm *.dot "
                                     "*.ppt *.pptx *.pptm *.ppsx *.ppsm *.pps "
                                     "*.pdf"),
                       ("所有文件", "*.*")])
        if paths:
            self._add_paths(list(paths))

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择包含支持文件的文件夹")
        if folder:
            self._add_paths([folder])

    def _add_paths(self, paths):
        files = dc.collect_files(paths)
        if not files:
            self.log("未找到支持的 Excel/Word/PPT/PDF 文件。")
            return
        existing = {os.path.abspath(f).lower() for f in self.files}
        added = 0
        for f in files:
            key = os.path.abspath(f).lower()
            if key not in existing:
                self.files.append(f)
                existing.add(key)
                added += 1
        self._refresh_list()
        self.log(f"已添加 {added} 个文件（支持 Excel/Word/PPT/PDF）。")

    def remove_selected(self):
        if not self._selected:
            return
        for idx in sorted(self._selected, reverse=True):
            if 0 <= idx < len(self.files):
                self.files.pop(idx)
        self._refresh_list()
        self.log("已移除选中文件。")

    def clear_list(self):
        self.files = []
        self._refresh_list()

    def open_plaintext_dir(self):
        if os.path.isdir(dc.PLAINTEXT_DIR):
            os.startfile(dc.PLAINTEXT_DIR)
        else:
            messagebox.showwarning("提示", f"明文区不存在：{dc.PLAINTEXT_DIR}")

    # ---------- 右键菜单 ----------
    def register_menu(self):
        try:
            import context_menu as cm
        except Exception as e:
            messagebox.showerror("错误", f"右键菜单模块加载失败：{e}")
            return
        n, fails = cm.register()
        msg = f"右键菜单注册完成：成功 {n} 处。\n\n现在右键点击 Excel/Word/PPT/PDF 文件，即可看到「{cm.MENU_NAME}」。"
        if fails:
            msg += "\n\n失败项：\n" + "\n".join(fails[:8])
        messagebox.showinfo("注册右键菜单", msg)

    def unregister_menu(self):
        try:
            import context_menu as cm
        except Exception as e:
            messagebox.showerror("错误", f"右键菜单模块加载失败：{e}")
            return
        n, fails = cm.unregister()
        messagebox.showinfo("卸载右键菜单", f"右键菜单已移除：{n} 处。")
        if fails:
            messagebox.showwarning("卸载右键菜单", "部分移除失败：\n" + "\n".join(fails[:8]))

    # ---------- 解密 ----------
    def start_decrypt(self):
        if self.running:
            return
        if not self.files:
            self._toast("请先添加要解密的文件或文件夹", kind="warn")
            return
        if not os.path.isdir(dc.PLAINTEXT_DIR):
            try:
                os.makedirs(dc.PLAINTEXT_DIR)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建明文区 {dc.PLAINTEXT_DIR}:\n{e}")
                return
        self.running = True
        self._stop = threading.Event()
        self._phase = ""
        self.btn_start.config(state="disabled")
        self.btn_stop.state(["!disabled"])
        self._set_pill("running")
        self.lbl_sum.config(text="正在准备文件…")
        self.progress["maximum"] = max(len(self.files), 1)
        self.progress["value"] = 0
        self.lbl_prog.config(text="0 / 0")
        self.lbl_pct.config(text="0%")
        files_snapshot = list(self.files)
        t = threading.Thread(target=self._worker, args=(files_snapshot,),
                             daemon=True)
        t.start()

    def stop_decrypt(self):
        """请求停止：当前文件处理完后不再继续。"""
        if self.running and self._stop is not None:
            self._stop.set()
            self.btn_stop.state(["disabled"])
            self.lbl_sum.config(text="正在停止…")
            self.log("⏹ 已请求停止，等待当前文件完成…")

    def _worker(self, files):
        def progress_cb(done, total_, cur, phase="decrypt"):
            self.root.after(0, self._update_progress, done, total_, cur, phase)

        def log_cb(msg):
            self.root.after(0, self.log, msg)

        ok, fail, results = dc.decrypt_batch(
            files, progress_cb=progress_cb, log_cb=log_cb,
            should_stop=lambda: self._stop.is_set() if self._stop else False)
        self.root.after(0, self._finish, ok, fail, results)

    def _update_progress(self, done, total, cur, phase="decrypt"):
        # 阶段切换时重置进度条，避免"准备阶段已 100%，解密阶段却不动"的假死观感
        if phase != self._phase:
            self._phase = phase
            self.progress["value"] = 0
            if phase == "decrypt":
                self.lbl_sum.config(text="正在解密…")
        self.progress["maximum"] = max(total, 1)
        self.progress["value"] = done
        self.lbl_prog.config(text=f"{done} / {total}")
        pct = int(round(done / total * 100)) if total else 0
        self.lbl_pct.config(text=f"{pct}%")
        idx = self._idx_by_path.get(os.path.abspath(cur))
        if idx is not None:
            self._set_status(idx, "run")

    def _finish(self, ok, fail, results):
        self.running = False
        self.btn_start.config(state="normal")
        self.btn_stop.state(["disabled"])
        self.progress["value"] = self.progress["maximum"]
        self._set_pill("fail" if fail else "done")
        stopped = self._stop is not None and self._stop.is_set()
        summary = f"成功 {ok} · 失败 {fail}" + ("（已手动停止）" if stopped else "")
        self.lbl_sum.config(text=summary)
        # 按源文件路径回填状态（结果只含加密文件，不能用下标直接对应）
        for src, okf, dest, msg in results:
            idx = self._idx_by_path.get(os.path.abspath(src))
            if idx is not None:
                self._set_status(idx, "ok" if okf else "fail")
        if stopped:
            self._toast(f"已停止：{summary}", kind="warn")
        elif fail:
            self._toast(f"解密完成：{summary}", kind="fail")
        else:
            self._toast(f"全部解密成功，共 {ok} 个文件")
        self.log(f"完成：{summary}")

    # ---------- Toast / 日志 ----------
    def _footer(self):
        foot = tk.Frame(self.root, bg=PAGE_BG)
        foot.pack(fill="x", padx=18, pady=(12, 0))
        self.lbl_prog = tk.Label(foot, text="0 / 0", bg=PAGE_BG, fg=TEXT_2,
                                 font=(FONT, 9))
        self.lbl_prog.pack(side="left")
        self.progress = ttk.Progressbar(foot, style="TProgressbar",
                                        mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)
        self.lbl_pct = tk.Label(foot, text="0%", bg=PAGE_BG, fg=TEXT_1,
                                font=(FONT, 9))
        self.lbl_pct.pack(side="right")

    def _toast(self, msg, kind="ok"):
        bg = {"ok": "#111827", "warn": "#854F0B", "fail": "#A32D2D"}[kind]
        t = tk.Label(self.root, text=msg, bg=bg, fg="#FFFFFF",
                     font=(FONT, 10), padx=18, pady=10)
        self.root.update_idletasks()
        t.place(relx=0.5, y=16, anchor="n")
        self.root.after(4200, t.destroy)

    def _log_area(self):
        bar = tk.Frame(self.root, bg=PAGE_BG)
        bar.pack(fill="x", padx=18, pady=(10, 0))
        self.btn_log = ttk.Button(bar, text="日志 ▾", command=self.toggle_log,
                                  style="Ghost.TButton")
        self.btn_log.pack(side="left")
        self.lbl_sum = tk.Label(bar, text="就绪 — 拖入文件或选择添加", bg=PAGE_BG,
                                fg=TEXT_3, font=(FONT, 9))
        self.lbl_sum.pack(side="right")
        self.log_panel = tk.Frame(self.root, bg=BORDER_C)

    def toggle_log(self):
        if self.log_visible:
            self.log_panel.pack_forget()
            self.btn_log.config(text="日志 ▾")
            self.log_visible = False
        else:
            self.log_panel.pack(fill="both", expand=True, padx=18, pady=(8, 14))
            self.btn_log.config(text="日志 ▴")
            self.log_visible = True

    def log(self, msg):
        if not hasattr(self, "txt_log"):
            inner = tk.Frame(self.log_panel, bg=SURFACE_BG)
            inner.pack(fill="both", expand=True, padx=1, pady=1)
            self.txt_log = tk.Text(inner, height=7, bg=SURFACE_BG, fg=TEXT_2,
                                   font=(FONT_MONO, 9), relief="flat",
                                   wrap="none", state="disabled",
                                   highlightthickness=0)
            sb = ttk.Scrollbar(inner, orient="vertical", command=self.txt_log.yview)
            self.txt_log.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            self.txt_log.pack(side="left", fill="both", expand=True)
            self.txt_log.tag_configure("ok", foreground="#3B6D11")
            self.txt_log.tag_configure("fail", foreground="#A32D2D")
            self.txt_log.tag_configure("warn", foreground="#854F0B")
            self.txt_log.tag_configure("info", foreground="#6B7280")
        self.txt_log.config(state="normal")
        text = msg.strip()          # 明细行带缩进，按标记判定颜色时需去空白
        tag = "info"
        if text.startswith("✓"):
            tag = "ok"
        elif text.startswith("✗") or "失败" in text:
            tag = "fail"
        elif text.startswith("⏹"):
            tag = "warn"
        elif "提示" in text or "未找到" in text or "WPS" in text:
            tag = "warn"
        self.txt_log.insert(tk.END, msg + "\n", tag)
        self.txt_log.see(tk.END)
        self.txt_log.config(state="disabled")

    def _check_env(self):
        if not os.path.isdir(dc.PLAINTEXT_DIR):
            self.log(f"提示：明文区不存在（{dc.PLAINTEXT_DIR}），开始解密时会自动创建。")
        wps = dc.find_wps_exe()
        if wps:
            self.log(f"已检测到 WPS：{wps}\n（PDF 解密将通过 WPS 触发）")
        else:
            self.log("提示：未检测到 WPS(wps.exe)，PDF 文件将无法解密（Excel/Word/PPT 不受影响）。")
            self._toast("未检测到 WPS：PDF 解密将不可用", kind="warn")

    def _bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self.add_files())
        self.root.bind("<Control-Return>", lambda e: self.start_decrypt())
        self.root.bind("<Delete>", lambda e: self.remove_selected())
        self.root.bind("<Control-l>", lambda e: self.toggle_log())

    # ---------- 单实例队列 ----------
    def _poll_queue(self):
        """轮询队列文件：主实例拾取其他实例转发来的文件，加入列表并自动解密。"""
        try:
            if os.path.exists(QUEUE_FILE):
                try:
                    size = os.path.getsize(QUEUE_FILE)
                except OSError:
                    size = 0
                if size > 0:
                    with open(QUEUE_FILE, encoding="utf-8") as f:
                        lines = [ln.strip() for ln in f if ln.strip()]
                    # 清空队列（先截断再处理，避免重复拾取）
                    try:
                        open(QUEUE_FILE, "w").close()
                    except Exception:
                        pass
                    if lines:
                        self._add_paths(lines)
                        if not self.running and self.files:
                            self.start_decrypt()
        except Exception:
            pass
        self.root.after(600, self._poll_queue)


def _run_register_cli(args):
    """无界面模式：注册/卸载右键菜单（即时命令，不排队）。"""
    root = tk.Tk()
    root.withdraw()
    try:
        import context_menu as cm
        if "--register" in args:
            n, fails = cm.register()
            msg = f"右键菜单注册完成：成功 {n} 处。"
            if fails:
                msg += "\n失败：\n" + "\n".join(fails[:8])
            messagebox.showinfo("注册右键菜单", msg)
        if "--unregister" in args:
            n, fails = cm.unregister()
            msg = f"右键菜单已移除：{n} 处。"
            if fails:
                msg += "\n失败：\n" + "\n".join(fails[:8])
            messagebox.showinfo("卸载右键菜单", msg)
    except Exception as e:
        messagebox.showerror("错误", f"右键菜单模块加载失败：{e}")
    finally:
        root.destroy()


def main():
    global _MAIN_LOCK
    args = sys.argv[1:]

    # 无界面模式：注册/卸载右键菜单（即时命令，不排队）
    if any(a in args for a in ("--register", "--unregister")):
        _run_register_cli(args)
        return

    # 命令行传入的文件/文件夹路径（右键"一键解密"传参）
    path_args = [a for a in args if not a.startswith("--") and os.path.exists(a)]

    # 单实例锁：已有实例在运行 → 把文件转发给主实例后退出（不弹第二个窗口）
    if _is_main_instance() is None:
        if path_args:
            _enqueue_paths(path_args)
        return
    _MAIN_LOCK = True

    # 尝试启用拖拽（tkinterdnd2 可选依赖）
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
        app = DecryptApp(root)
        app.dnd_available = True
        app._enable_dnd()
    except Exception:
        root = tk.Tk()
        app = DecryptApp(root)

    # 首次启动：显示初始配置引导（用户确认后写入标记，之后不再自动弹出；
    # 未确认直接关窗则下次启动仍会提示）
    if _is_first_run():
        app.show_first_run_guide()

    # 命令行传入的文件/文件夹：自动加入列表，并自动开始解密（右键"一键解密"流程）
    if path_args:
        app._add_paths(path_args)
        if app.files:
            app.root.after(400, app.start_decrypt)

    # 窗口关闭时释放单实例锁
    def _on_close():
        try:
            os.remove(LOCK_FILE)
        except Exception:
            pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", _on_close)
    try:
        root.mainloop()
    finally:
        # 兜底释放（异常/正常退出都清理）
        if _MAIN_LOCK:
            try:
                os.remove(LOCK_FILE)
            except Exception:
                pass


if __name__ == "__main__":
    main()
