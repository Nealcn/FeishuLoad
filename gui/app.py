"""XTF desktop sync assistant.

A small Tkinter front-end for non-technical users.  It builds a strict v2 YAML
file and runs the existing ``XTF.py sync`` CLI.  It never reimplements the
sync engine and never stores the app secret.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter.scrolledtext import ScrolledText

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

try:
    from gui import core as gui_core  # noqa: E402
except Exception as exc:  # noqa: BLE001 - dependency bootstrap
    tk_root = tk.Tk()
    tk_root.withdraw()
    messagebox.showerror(
        "XTF 界面启动失败",
        "缺少运行依赖或项目模块。\n\n"
        f"{exc}\n\n"
        "请在项目目录执行：\n"
        "python -m pip install -r requirements.txt",
    )
    raise SystemExit(1) from exc


APP_TITLE = "XTF 同步助手 — Excel → 飞书多维表格"


class ScrollableFrame(ttk.Frame):
    """A vertical-scrolling frame for long column lists and long panels.

    The inner frame keeps its natural height so tall content stays reachable
    instead of being clipped by the window.
    """

    def __init__(self, master: tk.Misc, **kwargs):
        height = int(kwargs.pop("height", 180))
        super().__init__(master, height=height, **kwargs)
        self.pack_propagate(False)
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, height=height)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)
        self.inner.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind(
            "<Enter>",
            lambda event: canvas.bind_all(
                "<MouseWheel>",
                lambda wheel: canvas.yview_scroll(
                    int(-1 * (wheel.delta / 120)), "units"
                ),
            ),
        )
        canvas.bind("<Leave>", lambda event: canvas.unbind_all("<MouseWheel>"))


class ScrollablePanel(ttk.Frame):
    """A grid-based panel whose content scrolls when it exceeds the window.

    Unlike :class:`ScrollableFrame` the content uses ``grid`` inside
    ``inner`` and the visible height follows the parent, so a long form is
    never cut off at the window edge.
    """

    def __init__(self, master: tk.Misc, **kwargs):
        super().__init__(master, **kwargs)
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)
        self._canvas = canvas
        self._window = canvas.create_window((0, 0), window=self.inner, anchor="nw")
        # The inner frame keeps its natural height; only its width follows
        # the canvas, so a long form scrolls instead of being clipped.
        self.inner.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(self._window, width=event.width),
        )
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind(
            "<Enter>",
            lambda event: canvas.bind_all(
                "<MouseWheel>",
                lambda wheel: canvas.yview_scroll(
                    int(-1 * (wheel.delta / 120)), "units"
                ),
            ),
        )
        canvas.bind("<Leave>", lambda event: canvas.unbind_all("<MouseWheel>"))


class XTFGui(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1160x820")
        self.minsize(980, 700)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.report_callback_exception = self._report_callback_exception

        self._busy = False
        self._proc: subprocess.Popen[str] | None = None
        self._temp_dirs: list[Path] = []
        self._target_info: gui_core.TargetInfo | None = None
        self._target_ref: gui_core.TargetRef | None = None
        self._table_items: dict[str, dict[str, str]] = {}
        self._table_labels: dict[str, str] = {}
        self._updating_url = False
        self._source_info: gui_core.SourceInfo | None = None
        self._index_candidates: list[gui_core.IndexCandidate] = []
        self._column_vars: dict[str, tk.BooleanVar] = {}
        self._column_checks: dict[str, ttk.Checkbutton] = {}
        self._plan_payload: dict | None = None
        self._preview_config: dict | None = None
        self._preview_config_path: Path | None = None
        self._preview_signature: str | None = None
        self._loaded_secret: str | None = None
        self._action_buttons: list[ttk.Button] = []

        self._init_style()
        self._init_vars()
        self._build_ui()
        self._toggle_dangerous()
        for key in ("incremental", "full"):
            self.mode_buttons[key].state(["disabled"])
        self._load_saved_settings()

    # ------------------------------------------------------------------
    # Setup
    # ------------------------------------------------------------------

    @staticmethod
    def _enable_dpi() -> None:
        if sys.platform != "win32":
            return
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            pass

    def _init_style(self) -> None:
        style = ttk.Style(self)
        for theme in ("vista", "clam", "default"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        base_font = ("Microsoft YaHei UI", 10)
        self.option_add("*Font", base_font)
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 17, "bold"))
        style.configure("Sub.TLabel", foreground="#666666")
        style.configure("Hint.TLabel", foreground="#666666")
        style.configure("Danger.TLabel", foreground="#b00020")
        style.configure("Success.TLabel", foreground="#0b6b2f")
        style.configure("Treeview", rowheight=25, font=base_font)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TButton", padding=(10, 5))
        style.configure("Accent.TButton", font=("Microsoft YaHei UI", 10, "bold"))

    def _init_vars(self) -> None:
        self.app_id_var = tk.StringVar()
        self.secret_var = tk.StringVar()
        self.show_secret_var = tk.BooleanVar(value=False)
        self.target_url_var = tk.StringVar()
        self.table_name_var = tk.StringVar()
        self.backend_var = tk.StringVar(value="base_v3")
        self.source_path_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="append")
        self.index_var = tk.StringVar()
        self.verify_var = tk.BooleanVar(value=True)
        self.create_missing_var = tk.BooleanVar(value=True)
        self.field_strategy_var = tk.StringVar(value="base")
        self.batch_size_var = tk.StringVar(value="500")
        self.show_dangerous_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="请先读取目标表，再选择 Excel。")
        self.target_summary_var = tk.StringVar(value="尚未读取目标表")
        self.source_summary_var = tk.StringVar(value="尚未选择数据文件")
        self.mode_summary_var = tk.StringVar(value="")
        self.mode_risk_var = tk.StringVar(value="")
        self.mode_description_var = tk.StringVar(value="")
        self.loaded_secret_var = tk.StringVar(value="")
        self.parsed_target_var = tk.StringVar(value="")
        self.result_var = tk.StringVar(value="")
        self.plan_risk_var = tk.StringVar(value="")

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(14, 10, 14, 4))
        header.grid(row=0, sticky="ew")
        ttk.Label(header, text="XTF 同步助手", style="Title.TLabel").pack(side="left")
        ttk.Label(
            header,
            text="把 Excel / CSV 安全同步到飞书多维表格",
            style="Sub.TLabel",
        ).pack(side="left", padx=14)
        ttk.Label(
            header,
            text="界面只做配置与预览，实际同步由 XTF CLI 执行",
            style="Hint.TLabel",
        ).pack(side="right")

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=1, sticky="nsew", padx=14, pady=4)
        self._build_tab_connection()
        self._build_tab_source()
        self._build_tab_settings()
        self._build_tab_preview()

        status = ttk.Frame(self, padding=(14, 4, 14, 10))
        status.grid(row=2, sticky="ew")
        ttk.Label(status, textvariable=self.status_var, style="Hint.TLabel").pack(
            side="left"
        )

    # ------------------------------------------------------------------
    # Tab 1: connection
    # ------------------------------------------------------------------

    def _build_tab_connection(self) -> None:
        tab = ttk.Frame(self.notebook, padding=14)
        self.notebook.add(tab, text="1  连接飞书")
        tab.columnconfigure(0, weight=1)

        ttk.Label(
            tab,
            text=(
                "准备：在飞书开放平台创建“自建应用”，开通多维表格读写权限并发布版本；"
                "再把这个应用添加为 Base 的文档协作者（建议“可管理”）。"
            ),
            style="Hint.TLabel",
            wraplength=1050,
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        credentials = ttk.LabelFrame(tab, text="应用凭据", padding=12)
        credentials.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        credentials.columnconfigure(1, weight=1)

        ttk.Label(credentials, text="App ID").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4
        )
        ttk.Entry(credentials, textvariable=self.app_id_var).grid(
            row=0, column=1, sticky="ew", pady=4
        )
        ttk.Label(credentials, text="App Secret").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=4
        )
        self.secret_entry = ttk.Entry(
            credentials, textvariable=self.secret_var, show="*"
        )
        self.secret_entry.grid(row=1, column=1, sticky="ew", pady=4)
        secret_row = ttk.Frame(credentials)
        secret_row.grid(row=1, column=2, sticky="w", padx=(8, 0))
        ttk.Checkbutton(
            secret_row,
            text="显示",
            variable=self.show_secret_var,
            command=self._toggle_secret,
        ).pack(side="left")
        ttk.Button(
            secret_row,
            text="清除已载入的 secret",
            command=self._clear_loaded_secret,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(
            credentials,
            textvariable=self.loaded_secret_var,
            style="Hint.TLabel",
        ).grid(row=2, column=1, sticky="w")
        ttk.Label(
            credentials,
            text="secret 只保存在本次运行内存中，不会写入配置文件。",
            style="Hint.TLabel",
        ).grid(row=3, column=1, sticky="w")

        target = ttk.LabelFrame(tab, text="目标多维表格", padding=12)
        target.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        target.columnconfigure(1, weight=1)

        ttk.Label(target, text="Base 链接").grid(
            row=0, column=0, sticky="w", padx=(0, 10), pady=4
        )
        ttk.Entry(target, textvariable=self.target_url_var).grid(
            row=0, column=1, sticky="ew", pady=4
        )
        self.target_url_var.trace_add("write", lambda *_: self._target_url_changed())
        ttk.Button(target, text="解析", command=self._parse_target).grid(
            row=0, column=2, padx=(8, 0)
        )
        ttk.Label(
            target, textvariable=self.parsed_target_var, style="Hint.TLabel"
        ).grid(row=1, column=1, columnspan=2, sticky="w")

        # The data table is chosen from a plain dropdown on its own row, so
        # no separate selection window has to be opened.
        ttk.Label(target, text="数据表").grid(
            row=2, column=0, sticky="w", padx=(0, 10), pady=4
        )
        self.table_combo = ttk.Combobox(
            target,
            textvariable=self.table_name_var,
            values=(),
            state="readonly",
            width=46,
        )
        self.table_combo.grid(row=2, column=1, sticky="ew", pady=4)
        self.table_combo.bind("<<ComboboxSelected>>", self._table_selected)
        ttk.Button(target, text="刷新列表", command=self._choose_table_async).grid(
            row=2, column=2, padx=(8, 0)
        )
        self.table_status_var = tk.StringVar(value="")
        ttk.Label(target, textvariable=self.table_status_var, style="Hint.TLabel").grid(
            row=3, column=1, columnspan=2, sticky="w"
        )

        ttk.Label(target, text="API 后端").grid(
            row=4, column=0, sticky="w", padx=(0, 10), pady=4
        )
        backend = ttk.Combobox(
            target,
            textvariable=self.backend_var,
            values=("base_v3", "bitable_v1"),
            state="readonly",
            width=16,
        )
        backend.grid(row=4, column=1, sticky="w", pady=4)
        ttk.Label(
            target,
            text="base_v3 权限不足时可切换 bitable_v1。",
            style="Hint.TLabel",
        ).grid(row=4, column=2, columnspan=2, sticky="w", padx=(8, 0))

        actions = ttk.Frame(tab)
        actions.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        self.read_target_btn = ttk.Button(
            actions, text="读取目标表", command=self._read_target
        )
        self.read_target_btn.pack(side="left")
        ttk.Button(
            actions, text="从 config.yaml 载入", command=self._load_config_file
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            actions, text="下一步：选择 Excel →", command=self._next_to_source
        ).pack(side="right")
        self._action_buttons.append(self.read_target_btn)

        ttk.Label(tab, textvariable=self.target_summary_var).grid(
            row=4, column=0, sticky="w", pady=(0, 4)
        )
        self.fields_holder, self.fields_tree = self._make_tree(
            tab,
            columns=("name", "kind", "access", "note"),
            headings=("字段", "类型", "可写性", "说明"),
            widths=(260, 120, 90, 420),
            height=11,
        )
        self.fields_holder.grid(row=5, column=0, sticky="nsew")
        tab.rowconfigure(5, weight=1)
        self.fields_tree.tag_configure("readonly", foreground="#777777")

    # ------------------------------------------------------------------
    # Tab 2: source
    # ------------------------------------------------------------------

    def _build_tab_source(self) -> None:
        tab = ttk.Frame(self.notebook, padding=14)
        self.notebook.add(tab, text="2  选择 Excel")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(3, weight=1)
        tab.rowconfigure(5, weight=2)

        file_frame = ttk.LabelFrame(tab, text="数据文件", padding=12)
        file_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)

        ttk.Label(file_frame, text="文件路径").grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        ttk.Entry(file_frame, textvariable=self.source_path_var).grid(
            row=0, column=1, sticky="ew"
        )
        ttk.Button(file_frame, text="选择文件…", command=self._browse_source).grid(
            row=0, column=2, padx=(8, 0)
        )
        ttk.Label(file_frame, text="工作表").grid(
            row=1, column=0, sticky="w", padx=(0, 10), pady=(8, 0)
        )
        self.sheet_combo = ttk.Combobox(
            file_frame, textvariable=self.sheet_name_var, state="normal", width=28
        )
        self.sheet_combo.grid(row=1, column=1, sticky="w", pady=(8, 0))
        ttk.Button(file_frame, text="读取数据", command=self._read_source).grid(
            row=1, column=2, padx=(8, 0), pady=(8, 0)
        )
        ttk.Label(
            file_frame,
            text=".csv 没有工作表概念；.xlsx/.xls 可填写名称或序号。",
            style="Hint.TLabel",
        ).grid(row=2, column=1, sticky="w", pady=(6, 0))

        ttk.Label(tab, textvariable=self.source_summary_var).grid(
            row=1, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Label(tab, text="数据预览（前 8 行）", style="Hint.TLabel").grid(
            row=2, column=0, sticky="w"
        )
        self.preview_holder = ttk.Frame(tab)
        self.preview_holder.grid(row=3, column=0, sticky="nsew", pady=(4, 8))
        self.preview_holder.columnconfigure(0, weight=1)
        self.preview_holder.rowconfigure(0, weight=1)

        ttk.Label(
            tab,
            text="要同步的列：只读计算列会自动排除；取消勾选可只同步部分列。",
            style="Hint.TLabel",
        ).grid(row=4, column=0, sticky="w", pady=(4, 4))
        self.columns_frame = ScrollableFrame(tab, height=230)
        self.columns_frame.grid(row=5, column=0, sticky="nsew", pady=(0, 8))
        self.columns_frame.columnconfigure(0, weight=1)

        actions = ttk.Frame(tab)
        actions.grid(row=6, column=0, sticky="ew", pady=(4, 0))
        ttk.Button(actions, text="全选", command=lambda: self._select_all(True)).pack(
            side="left"
        )
        ttk.Button(
            actions, text="全不选", command=lambda: self._select_all(False)
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            actions, text="← 上一步", command=lambda: self.notebook.select(0)
        ).pack(side="left", padx=(24, 0))
        ttk.Button(
            actions, text="下一步：同步设置 →", command=self._next_to_settings
        ).pack(side="right")

    # ------------------------------------------------------------------
    # Tab 3: settings
    # ------------------------------------------------------------------

    def _build_tab_settings(self) -> None:
        tab = ttk.Frame(self.notebook, padding=14)
        self.notebook.add(tab, text="3  同步设置")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)

        # The settings form can exceed the window height (long mode summary,
        # long index hints), so it lives in a scrollable panel.  The action
        # row stays outside the panel and is always visible.
        panel = ScrollablePanel(tab)
        panel.grid(row=0, column=0, sticky="nsew")
        body = panel.inner
        body.columnconfigure(0, weight=1)

        mode_frame = ttk.LabelFrame(body, text="同步模式", padding=12)
        mode_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        mode_frame.columnconfigure(1, weight=1)

        self.mode_descriptions = {
            "append": "只追加，不读取已有记录，不更新、不删除。最安全。",
            "incremental": "只新增目标中不存在的 key；已存在的记录跳过。",
            "full": "已存在的 key 更新，缺失的新增；不删除目标记录。",
            "overwrite": "删除目标中与源 key 匹配的记录后重建；需要删除授权。",
            "clone": "清空整张目标表后重写；需要删除授权，最危险。",
        }
        self.mode_buttons: dict[str, ttk.Radiobutton] = {}
        for row, key in enumerate(
            ("append", "incremental", "full", "overwrite", "clone")
        ):
            radio = ttk.Radiobutton(
                mode_frame,
                text=gui_core.MODE_LABELS[key],
                value=key,
                variable=self.mode_var,
                command=self._mode_changed,
            )
            radio.grid(row=row, column=0, sticky="w", pady=3)
            self.mode_buttons[key] = radio

        # One dynamic description line instead of five static rows keeps the
        # form short and always describes the currently selected mode.
        self.mode_description_var.set(self.mode_descriptions["append"])
        self.mode_description_label = ttk.Label(
            mode_frame,
            textvariable=self.mode_description_var,
            style="Hint.TLabel",
            wraplength=900,
        )
        self.mode_description_label.grid(
            row=0, column=1, sticky="w", padx=(12, 0), pady=3
        )

        danger = ttk.Frame(mode_frame)
        danger.grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(
            danger,
            text="显示危险模式（覆盖 / 清空）",
            variable=self.show_dangerous_var,
            command=self._toggle_dangerous,
        ).pack(side="left")

        index_frame = ttk.LabelFrame(body, text="唯一索引列", padding=12)
        index_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        index_frame.columnconfigure(1, weight=1)
        ttk.Label(index_frame, text="索引列").grid(
            row=0, column=0, sticky="w", padx=(0, 10)
        )
        self.index_combo = ttk.Combobox(
            index_frame,
            textvariable=self.index_var,
            state="readonly",
            values=(),
            width=30,
        )
        self.index_combo.grid(row=0, column=1, sticky="w")
        self.index_hint = ttk.Label(
            index_frame,
            text="读取源文件和目标表后自动检查。",
            style="Hint.TLabel",
            wraplength=900,
        )
        self.index_hint.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
        ttk.Label(
            index_frame,
            text="by_key 模式要求该列在源数据和目标表中都唯一且非空。",
            style="Hint.TLabel",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))

        options = ttk.LabelFrame(body, text="选项", padding=12)
        options.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        options.columnconfigure(1, weight=1)

        ttk.Checkbutton(
            options,
            text="写后读回确认（推荐）",
            variable=self.verify_var,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=3)
        ttk.Checkbutton(
            options,
            text="自动创建目标表缺少的字段",
            variable=self.create_missing_var,
            command=self._rebuild_columns,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=3)
        ttk.Label(options, text="字段类型策略").grid(
            row=2, column=0, sticky="w", pady=3
        )
        ttk.Combobox(
            options,
            textvariable=self.field_strategy_var,
            values=("raw", "base", "auto", "intelligence"),
            state="readonly",
            width=18,
        ).grid(row=2, column=1, sticky="w", pady=3)
        ttk.Label(options, text="批大小").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Spinbox(
            options,
            from_=1,
            to=5000,
            textvariable=self.batch_size_var,
            width=10,
        ).grid(row=3, column=1, sticky="w", pady=3)

        ttk.Label(
            body,
            textvariable=self.mode_summary_var,
            wraplength=900,
            style="Hint.TLabel",
        ).grid(row=3, column=0, sticky="w", pady=(0, 4))
        self.mode_risk_label = ttk.Label(
            body,
            textvariable=self.mode_risk_var,
            wraplength=900,
            style="Danger.TLabel",
        )
        self.mode_risk_label.grid(row=4, column=0, sticky="w", pady=(0, 8))

        actions = ttk.Frame(tab)
        actions.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(
            actions, text="← 上一步", command=lambda: self.notebook.select(1)
        ).pack(side="left")
        self.preview_btn = ttk.Button(
            actions,
            text="生成预览（不会写入）",
            command=self._preview,
            style="Accent.TButton",
        )
        self.preview_btn.pack(side="right")
        self._action_buttons.append(self.preview_btn)

    # ------------------------------------------------------------------
    # Tab 4: preview and execution
    # ------------------------------------------------------------------

    def _build_tab_preview(self) -> None:
        tab = ttk.Frame(self.notebook, padding=14)
        self.notebook.add(tab, text="4  预览与执行")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)
        tab.rowconfigure(5, weight=1)

        ttk.Label(
            tab,
            text="预览使用与正式执行相同的 planner；预览不会写入飞书。",
            style="Hint.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Label(tab, textvariable=self.plan_risk_var, style="Danger.TLabel").grid(
            row=1, column=0, sticky="w"
        )

        plan_frame = ttk.LabelFrame(tab, text="计划", padding=8)
        plan_frame.grid(row=2, column=0, sticky="nsew", pady=(4, 8))
        plan_frame.columnconfigure(0, weight=1)
        plan_frame.rowconfigure(0, weight=1)
        self.plan_holder, self.plan_tree = self._make_tree(
            plan_frame,
            columns=("kind", "count", "unit", "risk"),
            headings=("动作", "数量", "单位", "风险"),
            widths=(300, 100, 100, 120),
            height=7,
        )
        self.plan_holder.grid(row=0, column=0, sticky="nsew")
        self.plan_tree.tag_configure("danger", foreground="#b00020")

        warnings_frame = ttk.LabelFrame(tab, text="提醒", padding=8)
        warnings_frame.grid(row=3, column=0, sticky="ew", pady=(0, 8))
        warnings_frame.columnconfigure(0, weight=1)
        self.warnings_text = tk.Text(
            warnings_frame, height=3, wrap="word", state="disabled"
        )
        self.warnings_text.grid(row=0, column=0, sticky="ew")

        actions = ttk.Frame(tab)
        actions.grid(row=4, column=0, sticky="ew", pady=(0, 6))
        self.repreview_btn = ttk.Button(
            actions, text="重新生成预览", command=self._preview
        )
        self.repreview_btn.pack(side="left")
        self.execute_btn = ttk.Button(
            actions,
            text="执行同步",
            command=self._execute,
            style="Accent.TButton",
            state="disabled",
        )
        self.execute_btn.pack(side="right")
        ttk.Button(actions, text="打开日志目录", command=self._open_logs).pack(
            side="right", padx=(0, 8)
        )
        self._action_buttons.extend([self.repreview_btn, self.execute_btn])

        log_frame = ttk.LabelFrame(tab, text="运行日志", padding=8)
        log_frame.grid(row=5, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = ScrolledText(
            log_frame, height=10, wrap="word", state="disabled"
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        ttk.Label(tab, textvariable=self.result_var, style="Success.TLabel").grid(
            row=6, column=0, sticky="w", pady=(6, 0)
        )

    def _make_tree(
        self,
        parent: tk.Misc,
        *,
        columns: tuple[str, ...],
        headings: tuple[str, ...],
        widths: tuple[int, ...],
        height: int,
    ) -> tuple[ttk.Frame, ttk.Treeview]:
        holder = ttk.Frame(parent)
        tree = ttk.Treeview(holder, columns=columns, show="headings", height=height)
        for column, heading, width in zip(columns, headings, widths):
            tree.heading(column, text=heading)
            tree.column(column, width=width, anchor="w", stretch=column == columns[-1])
        scrollbar = ttk.Scrollbar(holder, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        return holder, tree

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _report_callback_exception(self, exc, value, tb) -> None:  # noqa: ANN001
        detail = "".join(traceback.format_exception(exc, value, tb))
        self._append_log(detail)
        messagebox.showerror("程序错误", str(value), parent=self)

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message.rstrip() + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _set_busy(self, busy: bool, text: str = "") -> None:
        self._busy = busy
        for button in self._action_buttons:
            button.state(["disabled"] if busy else ["!disabled"])
        if busy and text:
            self._set_status(text)
        self.update_idletasks()

    def _secret(self) -> str:
        value = self.secret_var.get().strip()
        if value:
            return value
        if self._loaded_secret:
            return self._loaded_secret
        raise gui_core.GuiError(
            "请填写 App Secret。", hint="也可以在步骤 1 点击“从 config.yaml 载入”。"
        )

    def _toggle_secret(self) -> None:
        self.secret_entry.configure(show="" if self.show_secret_var.get() else "*")

    def _clear_loaded_secret(self) -> None:
        self._loaded_secret = None
        self.loaded_secret_var.set("")

    def _show_error(self, title: str, exc: Exception) -> None:
        if isinstance(exc, gui_core.GuiError):
            message = str(exc)
            if exc.hint:
                message += "\n\n" + exc.hint
        else:
            message = str(exc)
        self._append_log(f"{title}: {message}")
        messagebox.showerror(title, message, parent=self)

    # ------------------------------------------------------------------
    # Settings persistence
    # ------------------------------------------------------------------

    def _load_saved_settings(self) -> None:
        values = gui_core.load_settings()
        if not values:
            return
        mapping = {
            "app_id": self.app_id_var,
            "target_url": self.target_url_var,
            "source_path": self.source_path_var,
            "sheet_name": self.sheet_name_var,
            "mode": self.mode_var,
            "index_column": self.index_var,
            "field_strategy": self.field_strategy_var,
            "backend": self.backend_var,
            "batch_size": self.batch_size_var,
        }
        for key, var in mapping.items():
            if key in values and values[key] is not None:
                var.set(str(values[key]))
        if "verify_remote_writes" in values:
            self.verify_var.set(bool(values["verify_remote_writes"]))
        if "create_missing_fields" in values:
            self.create_missing_var.set(bool(values["create_missing_fields"]))
        if values.get("mode") in {"overwrite", "clone"}:
            self.show_dangerous_var.set(True)
            self._toggle_dangerous()

    def _save_settings(self) -> None:
        gui_core.save_settings(
            {
                "app_id": self.app_id_var.get().strip(),
                "target_url": self.target_url_var.get().strip(),
                "source_path": self.source_path_var.get().strip(),
                "sheet_name": self.sheet_name_var.get().strip(),
                "mode": self.mode_var.get(),
                "index_column": self.index_var.get(),
                "verify_remote_writes": bool(self.verify_var.get()),
                "create_missing_fields": bool(self.create_missing_var.get()),
                "field_strategy": self.field_strategy_var.get(),
                "backend": self.backend_var.get(),
                "batch_size": self.batch_size_var.get(),
            }
        )

    # ------------------------------------------------------------------
    # Tab 1 actions
    # ------------------------------------------------------------------

    def _load_config_file(self) -> None:
        if self._busy:
            return
        default = ROOT / "config.yaml"
        path = str(default) if default.is_file() else ""
        if not path:
            path = filedialog.askopenfilename(
                title="选择 XTF 配置",
                filetypes=[("YAML", "*.yaml *.yml"), ("All", "*.*")],
                parent=self,
            )
        if not path:
            return
        try:
            loaded = gui_core.load_config_file(path)
        except Exception as exc:  # noqa: BLE001
            self._show_error("载入配置失败", exc)
            return
        self.app_id_var.set(loaded.app_id)
        self.source_path_var.set(loaded.source_path)
        self.sheet_name_var.set(loaded.sheet_name or "")
        self.mode_var.set(loaded.mode)
        self.index_var.set(loaded.index_column or "")
        self.verify_var.set(loaded.verify_remote_writes)
        self.create_missing_var.set(loaded.create_missing_fields)
        self.field_strategy_var.set(loaded.field_strategy)
        self.backend_var.set(loaded.backend)
        self.batch_size_var.set(str(loaded.batch_size))
        self._loaded_secret = loaded.secret
        self.loaded_secret_var.set(
            "已从 config.yaml 载入 secret（不会显示或保存）"
            if loaded.secret
            else "config.yaml 中没有 secret"
        )
        if loaded.target and loaded.target.kind == "bitable":
            self._target_ref = loaded.target
            self._updating_url = True
            try:
                self.target_url_var.set(
                    f"https://feishu.cn/base/{loaded.target.app_token}"
                    f"?table={loaded.target.table_id}"
                )
            finally:
                self._updating_url = False
            self.parsed_target_var.set(loaded.target.label)
            # Populate the table dropdown from the loaded Base.
            self._choose_table_async()
        elif loaded.target and loaded.target.kind == "sheet":
            self._target_ref = loaded.target
            self._updating_url = True
            try:
                self.target_url_var.set(
                    f"https://feishu.cn/sheets/{loaded.target.spreadsheet_token}"
                )
            finally:
                self._updating_url = False
            messagebox.showinfo(
                "当前不支持 Sheet 界面",
                "已载入配置，但图形界面目前只支持多维表格；"
                "电子表格请使用命令行运行。",
                parent=self,
            )
        if loaded.mode in {"overwrite", "clone"}:
            self.show_dangerous_var.set(True)
            self._toggle_dangerous()
        self._refresh_sheet_names()
        self._set_status(f"已载入配置：{path}")

    def _target_url_changed(self) -> None:
        """Clear a stale table list when the Base link changes."""
        if self._updating_url:
            return
        raw = self.target_url_var.get().strip()
        ref = self._target_ref
        if ref is not None and ref.original_url == raw:
            return
        self._table_items = {}
        self._table_labels = {}
        self.table_combo.configure(values=())
        self.table_combo.set("")
        self.table_status_var.set("")
        self._target_ref = None

    def _parse_target(self) -> bool:
        try:
            ref = gui_core.parse_target_url(self.target_url_var.get())
        except Exception as exc:  # noqa: BLE001
            self._show_error("无法解析链接", exc)
            return False
        if ref.kind == "sheet":
            messagebox.showinfo(
                "当前不支持 Sheet 界面",
                "图形界面目前只支持多维表格（Bitable）。"
                "电子表格请使用命令行运行 XTF.py。",
                parent=self,
            )
            self._target_ref = None
            self.parsed_target_var.set("当前不支持 Sheet 界面")
            return False
        self._target_ref = ref
        if ref.table_id:
            self.parsed_target_var.set(ref.label)
        else:
            self.parsed_target_var.set(
                f"Base {ref.app_token}；请在下拉框中选择数据表。"
            )
        # Load the table list right away so the dropdown is usable.
        self._choose_table_async()
        return True

    def _current_target(self) -> gui_core.TargetRef:
        raw = self.target_url_var.get().strip()
        ref = self._target_ref
        if ref is None or ref.original_url != raw:
            ref = gui_core.parse_target_url(raw)
            self._target_ref = ref
        if ref.kind != "bitable":
            raise gui_core.GuiError("当前界面只支持多维表格。")
        if not ref.table_id:
            raise gui_core.GuiError(
                "链接中缺少 table 参数。",
                hint="请在“数据表”下拉框中选择，或粘贴带 ?table=... 的完整链接。",
            )
        return ref

    def _choose_table_async(self) -> None:
        if self._busy:
            return
        try:
            app_id = self.app_id_var.get().strip()
            secret = self._secret()
            ref = gui_core.parse_target_url(self.target_url_var.get())
            if ref.kind != "bitable":
                raise gui_core.GuiError("当前界面只支持多维表格。")
        except Exception as exc:  # noqa: BLE001
            self._show_error("无法选择数据表", exc)
            return
        self._set_busy(True, "正在读取数据表列表…")

        def worker() -> None:
            try:
                tables = gui_core.list_tables(
                    app_id, secret, ref.app_token, self.backend_var.get()
                )
            except Exception as exc:  # noqa: BLE001
                self.after(0, lambda exc=exc: self._choose_table_failed(exc))
                return
            self.after(0, lambda tables=tables: self._choose_table_done(ref, tables))

        threading.Thread(target=worker, daemon=True).start()

    def _choose_table_failed(self, exc: Exception) -> None:
        self._set_busy(False)
        self.table_status_var.set("读取数据表列表失败")
        self._show_error("读取数据表列表失败", exc)

    def _choose_table_done(
        self, ref: gui_core.TargetRef, tables: list[dict[str, str]]
    ) -> None:
        self._set_busy(False)
        if not tables:
            self.table_combo.configure(values=())
            self.table_status_var.set("没有读到任何数据表，请检查权限。")
            messagebox.showinfo(
                "没有可用数据表",
                "应用没有读到任何数据表。请检查权限和文档协作者设置。",
                parent=self,
            )
            return
        self._table_items = {item["table_id"]: item for item in tables}
        labels = {
            table_id: f"{item['name']}  ({table_id})"
            for table_id, item in self._table_items.items()
        }
        self._table_labels = labels
        self.table_combo.configure(values=list(labels.values()))
        self.table_status_var.set(f"共 {len(tables)} 张数据表")
        # Keep the table from the link when it is still accessible.
        current = ref.table_id
        if current and current in labels:
            self.table_combo.set(labels[current])
            self._apply_table_selection(current)
        else:
            self.table_combo.set("")
            self.parsed_target_var.set(
                f"Base {ref.app_token}；请在下拉框中选择数据表。"
            )

    def _table_selected(self, event: tk.Event | None = None) -> None:
        """Apply the table picked in the dropdown."""
        label = self.table_combo.get()
        table_id = ""
        for candidate, text in getattr(self, "_table_labels", {}).items():
            if text == label:
                table_id = candidate
                break
        if not table_id:
            return
        if not self._apply_table_selection(table_id):
            return
        # Picking a table is an explicit choice; read it right away.
        self._read_target()

    def _apply_table_selection(self, table_id: str) -> bool:
        ref = self._target_ref
        if ref is None:
            try:
                ref = gui_core.parse_target_url(self.target_url_var.get())
            except Exception:  # noqa: BLE001
                return False
            if ref.kind != "bitable":
                return False
        url = f"https://feishu.cn/base/{ref.app_token}?table={table_id}"
        selected = gui_core.TargetRef(
            kind="bitable",
            app_token=ref.app_token,
            table_id=table_id,
            original_url=url,
        )
        self._target_ref = selected
        self._updating_url = True
        try:
            self.target_url_var.set(url)
        finally:
            self._updating_url = False
        item = getattr(self, "_table_items", {}).get(table_id)
        if item:
            self.parsed_target_var.set(
                f"Base {ref.app_token} · 表 {item['name']} ({table_id})"
            )
        else:
            self.parsed_target_var.set(selected.label)
        return True

    def _read_target(self) -> None:
        if self._busy:
            return
        try:
            app_id = self.app_id_var.get().strip()
            secret = self._secret()
            ref = self._current_target()
        except Exception as exc:  # noqa: BLE001
            self._show_error("无法读取目标表", exc)
            return
        self._set_busy(True, "正在读取目标表字段和记录…")
        self._set_status("正在读取目标表…")

        def worker() -> None:
            try:
                info = gui_core.probe_bitable(
                    app_id, secret, ref, self.backend_var.get()
                )
            except Exception as exc:  # noqa: BLE001
                self.after(0, lambda exc=exc: self._read_target_failed(exc))
                return
            self.after(0, lambda info=info: self._read_target_done(info))

        threading.Thread(target=worker, daemon=True).start()

    def _read_target_failed(self, exc: Exception) -> None:
        self._set_busy(False)
        self._set_status("读取目标表失败")
        self._show_error("读取目标表失败", exc)

    def _read_target_done(self, info: gui_core.TargetInfo) -> None:
        self._set_busy(False)
        self._target_info = info
        self._set_status("目标表已读取")
        self.target_summary_var.set(
            f"表：{info.table_name}（{info.ref.table_id}） · "
            f"{info.record_count} 条记录 · {len(info.fields)} 个字段 · "
            f"后端 {info.backend}"
        )
        self._populate_fields(info)
        self._update_index_candidates()
        if self._source_info is not None:
            self._rebuild_columns()
        self._update_mode_summary()

    def _populate_fields(self, info: gui_core.TargetInfo) -> None:
        for item in self.fields_tree.get_children():
            self.fields_tree.delete(item)
        for field in info.fields:
            access = "可写" if field.writable else "只读"
            note = "" if field.writable else "计算列/系统字段，XTF 不会写入"
            self.fields_tree.insert(
                "",
                "end",
                values=(field.name, field.kind_label, access, note),
                tags=() if field.writable else ("readonly",),
            )

    def _next_to_source(self) -> None:
        if self._target_info is None:
            messagebox.showinfo(
                "请先读取目标表",
                "请先在步骤 1 填写凭据并点击“读取目标表”。",
                parent=self,
            )
            return
        self.notebook.select(1)

    # ------------------------------------------------------------------
    # Tab 2 actions
    # ------------------------------------------------------------------

    def _refresh_sheet_names(self) -> None:
        path = self.source_path_var.get().strip()
        if not path:
            self.sheet_combo.configure(values=())
            return
        names = gui_core.list_sheet_names(path)
        self.sheet_combo.configure(values=names)
        if names and self.sheet_name_var.get().strip() not in names:
            self.sheet_name_var.set(names[0])

    def _browse_source(self) -> None:
        if self._busy:
            return
        path = filedialog.askopenfilename(
            title="选择 Excel / CSV",
            filetypes=[
                ("Excel / CSV", "*.xlsx *.xls *.csv"),
                ("Excel", "*.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("All", "*.*"),
            ],
            parent=self,
        )
        if not path:
            return
        self.source_path_var.set(path)
        self._refresh_sheet_names()
        self._read_source()

    def _read_source(self) -> None:
        if self._busy:
            return
        path = self.source_path_var.get().strip()
        if not path:
            messagebox.showinfo(
                "请选择文件", "请先选择 Excel 或 CSV 文件。", parent=self
            )
            return
        sheet = self.sheet_name_var.get().strip() or None
        self._set_busy(True, "正在读取数据文件…")
        self._set_status("正在读取数据文件…")

        def worker() -> None:
            try:
                info = gui_core.read_source(path, sheet)
            except Exception as exc:  # noqa: BLE001
                self.after(0, lambda exc=exc: self._read_source_failed(exc))
                return
            self.after(0, lambda info=info: self._read_source_done(info))

        threading.Thread(target=worker, daemon=True).start()

    def _read_source_failed(self, exc: Exception) -> None:
        self._set_busy(False)
        self._set_status("读取数据文件失败")
        self._show_error("读取数据文件失败", exc)

    def _read_source_done(self, info: gui_core.SourceInfo) -> None:
        self._set_busy(False)
        self._source_info = info
        self.source_summary_var.set(
            f"{info.path.name} · {info.rows} 行 · {len(info.columns)} 列"
            + (f" · 工作表 {info.sheet_name}" if info.sheet_name else "")
        )
        self._populate_preview(info)
        self._rebuild_columns()
        self._update_index_candidates()
        self._update_mode_summary()
        self._set_status("数据文件已读取")

    def _populate_preview(self, info: gui_core.SourceInfo) -> None:
        for child in self.preview_holder.winfo_children():
            child.destroy()
        columns = info.columns[:12]
        tree = ttk.Treeview(
            self.preview_holder,
            columns=columns,
            show="headings",
            height=4,
        )
        for column in columns:
            tree.heading(column, text=column)
            tree.column(column, width=130, anchor="w", stretch=False)
        for row in info.preview:
            tree.insert("", "end", values=row[: len(columns)])
        yscroll = ttk.Scrollbar(
            self.preview_holder, orient="vertical", command=tree.yview
        )
        xscroll = ttk.Scrollbar(
            self.preview_holder, orient="horizontal", command=tree.xview
        )
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.preview_holder.columnconfigure(0, weight=1)
        self.preview_holder.rowconfigure(0, weight=1)
        if len(info.columns) > len(columns):
            ttk.Label(
                self.preview_holder,
                text=f"仅显示前 {len(columns)} 列，共 {len(info.columns)} 列。",
                style="Hint.TLabel",
            ).grid(row=2, column=0, sticky="w", pady=(4, 0))

    def _rebuild_columns(self) -> None:
        for child in self.columns_frame.inner.winfo_children():
            child.destroy()
        self._column_vars = {}
        self._column_checks = {}
        info = self._source_info
        if info is None:
            ttk.Label(
                self.columns_frame.inner,
                text="读取数据文件后会在这里列出列名。",
                style="Hint.TLabel",
            ).pack(anchor="w", padx=4, pady=4)
            return
        target_map = self._target_info.field_map if self._target_info else {}
        for row, column in enumerate(info.columns):
            frame = ttk.Frame(self.columns_frame.inner)
            frame.pack(fill="x", padx=4, pady=2)
            target_field = target_map.get(column)
            if target_field is None:
                enabled = bool(self.create_missing_var.get())
                mapping = (
                    "目标表无此字段；运行时将自动创建"
                    if enabled
                    else "目标表无此字段；不会写入（create_missing_fields=false）"
                )
                default = enabled
            elif target_field.writable:
                enabled = True
                mapping = (
                    f"→ 目标字段「{target_field.name}」（{target_field.kind_label}）"
                )
                default = True
            else:
                enabled = False
                mapping = "目标字段只读（公式/查找/系统字段），XTF 不会写入"
                default = False
            var = tk.BooleanVar(value=default)
            check = ttk.Checkbutton(
                frame,
                text=column,
                variable=var,
                state="normal" if enabled else "disabled",
            )
            check.pack(side="left")
            stat = info.stats.get(column)
            suffix = (
                f"  ·  源 {stat.nonempty} 非空 / {stat.unique} 唯一" if stat else ""
            )
            ttk.Label(
                frame,
                text=mapping + suffix,
                style="Hint.TLabel" if enabled else "Danger.TLabel",
            ).pack(side="left", padx=(10, 0))
            self._column_vars[column] = var
            self._column_checks[column] = check

    def _select_all(self, value: bool) -> None:
        for column, var in self._column_vars.items():
            check = self._column_checks.get(column)
            if check is not None and "disabled" not in check.state():
                var.set(value)
        self._update_mode_summary()

    def _effective_columns(self) -> list[str]:
        info = self._source_info
        if info is None:
            return []
        target_map = self._target_info.field_map if self._target_info else {}
        selected: list[str] = []
        for column in info.columns:
            var = self._column_vars.get(column)
            if var is None or not var.get():
                continue
            target_field = target_map.get(column)
            if target_field is not None and not target_field.writable:
                continue
            if target_field is None and not self.create_missing_var.get():
                continue
            selected.append(column)
        mode = self.mode_var.get()
        if mode in {"incremental", "full", "overwrite"}:
            index = self.index_var.get().strip()
            if index and index not in selected:
                selected.append(index)
        return selected

    def _columns_for_mode(self, mode: str) -> list[str]:
        info = self._source_info
        if info is None:
            return []
        if mode not in {"overwrite", "clone"}:
            return self._effective_columns()
        target_map = self._target_info.field_map if self._target_info else {}
        columns: list[str] = []
        for column in info.columns:
            target_field = target_map.get(column)
            if target_field is None:
                if self.create_missing_var.get():
                    columns.append(column)
            elif target_field.writable:
                columns.append(column)
        return columns

    def _next_to_settings(self) -> None:
        if self._source_info is None:
            messagebox.showinfo(
                "请先读取数据", "请先选择并读取 Excel/CSV 文件。", parent=self
            )
            return
        self._update_index_candidates()
        self.notebook.select(2)

    # ------------------------------------------------------------------
    # Tab 3 actions
    # ------------------------------------------------------------------

    def _toggle_dangerous(self) -> None:
        enabled = bool(self.show_dangerous_var.get())
        for key in ("overwrite", "clone"):
            button = self.mode_buttons.get(key)
            if button is not None:
                button.state(["!disabled"] if enabled else ["disabled"])
        if not enabled and self.mode_var.get() in {"overwrite", "clone"}:
            self.mode_var.set("append")
            self._mode_changed()
        if self._source_info is not None and self._target_info is not None:
            self._update_index_candidates()

    def _mode_changed(self) -> None:
        mode = self.mode_var.get()
        if mode in {"overwrite", "clone"} and not self.show_dangerous_var.get():
            self.show_dangerous_var.set(True)
            self._toggle_dangerous()
        if mode in {"incremental", "full", "overwrite"}:
            safe = [item for item in self._index_candidates if item.safe]
            if not safe:
                messagebox.showwarning(
                    "没有唯一索引列",
                    "当前源数据和目标表没有共同且唯一、非空的列。\n"
                    "请使用“只追加”，或先在目标表创建并回填唯一键。",
                    parent=self,
                )
                self.mode_var.set("append")
                mode = "append"
            elif not self.index_var.get():
                self.index_var.set(safe[0].name)
        self._update_mode_description()
        self._update_index_hint()
        self._update_mode_summary()

    def _update_mode_description(self) -> None:
        """Show only the selected mode's description, next to the radio."""
        mode = self.mode_var.get()
        description = self.mode_descriptions.get(mode, "")
        self.mode_description_var.set(description)
        style = "Danger.TLabel" if mode in {"overwrite", "clone"} else "Hint.TLabel"
        self.mode_description_label.configure(style=style)

    def _update_index_candidates(self) -> None:
        if self._source_info is None or self._target_info is None:
            self._index_candidates = []
            self.index_combo.configure(values=())
            self.index_hint.configure(text="读取源文件和目标表后自动检查。")
            return
        self._index_candidates = gui_core.index_candidates(
            self._source_info, self._target_info
        )
        safe = [item for item in self._index_candidates if item.safe]
        names = [item.name for item in safe]
        self.index_combo.configure(values=names)
        if self.index_var.get() not in names:
            self.index_var.set(names[0] if names else "")
        if names:
            self.index_hint.configure(text=f"可用唯一列：{', '.join(names)}")
            self.index_hint.configure(wraplength=900)
            for key in ("incremental", "full"):
                self.mode_buttons[key].state(["!disabled"])
            if self.show_dangerous_var.get():
                self.mode_buttons["overwrite"].state(["!disabled"])
            else:
                self.mode_buttons["overwrite"].state(["disabled"])
        else:
            details = "；".join(
                f"{item.name}：{item.reason}" for item in self._index_candidates[:3]
            )
            self.index_hint.configure(
                text="没有唯一且非空的列，只能使用“只追加”。"
                + (f"\n{details}" if details else "")
            )
            for key in ("incremental", "full", "overwrite"):
                self.mode_buttons[key].state(["disabled"])
            if self.mode_var.get() in {"incremental", "full", "overwrite"}:
                self.mode_var.set("append")
        self._update_index_hint()
        self._update_mode_summary()

    def _update_index_hint(self) -> None:
        mode = self.mode_var.get()
        if mode in {"append", "clone"}:
            self.index_combo.configure(state="disabled")
            return
        self.index_combo.configure(
            state="readonly" if self.index_combo["values"] else "disabled"
        )

    def _update_mode_summary(self) -> None:
        mode = self.mode_var.get()
        columns = self._columns_for_mode(mode)
        facts = [
            f"模式：{gui_core.MODE_LABELS.get(mode, mode)}",
            f"同步列：{len(columns)} 列",
        ]
        if mode in {"incremental", "full", "overwrite"}:
            facts.append(f"索引列：{self.index_var.get() or '未选择'}")
        if mode == "append":
            facts.append("不会更新或删除已有记录")
        if self._target_info is not None:
            readonly = sum(1 for item in self._target_info.fields if not item.writable)
            if readonly:
                facts.append(f"目标表 {readonly} 个只读字段已自动排除")
        self.mode_summary_var.set("；".join(facts))

        # Risk notes are kept apart so the factual line stays on one row.
        risks: list[str] = []
        if mode == "overwrite":
            risks.append("会删除匹配记录后重建；源文件没有的可写字段在新记录中为空")
        if mode == "clone":
            risks.append("会清空整张表；源文件没有的可写字段也会被清空")
        if mode in {"overwrite", "clone"}:
            risks.append("该模式会写入全部可写列，忽略部分列勾选")
            if self.backend_var.get() == "base_v3":
                risks.append(
                    "Base v3 删除回读曾有兼容问题；建议先备份，或改用 bitable_v1"
                )
        self.mode_risk_var.set("\n".join(risks))
        self.mode_risk_label.configure(
            style="Danger.TLabel" if risks else "Hint.TLabel"
        )

    def _settings_signature(self) -> str:
        mode = self.mode_var.get()
        source_path = Path(self.source_path_var.get().strip())
        file_state: dict[str, object] = {}
        try:
            stat = source_path.stat()
            file_state = {"mtime_ns": stat.st_mtime_ns, "size": stat.st_size}
        except OSError:
            pass
        payload = {
            "app_id": self.app_id_var.get().strip(),
            "target_url": self.target_url_var.get().strip(),
            "source_path": self.source_path_var.get().strip(),
            "sheet_name": self.sheet_name_var.get().strip(),
            "file_state": file_state,
            "mode": mode,
            "columns": self._columns_for_mode(mode),
            "index": self.index_var.get().strip(),
            "verify_remote_writes": bool(self.verify_var.get()),
            "create_missing_fields": bool(self.create_missing_var.get()),
            "field_strategy": self.field_strategy_var.get(),
            "backend": self.backend_var.get(),
            "batch_size": self.batch_size_var.get().strip(),
        }
        return json.dumps(payload, sort_keys=True, ensure_ascii=False)

    def _build_config_payload(self) -> tuple[dict, Path, list[str]]:
        if self._target_info is None:
            raise gui_core.GuiError("请先读取目标表。")
        if self._source_info is None:
            raise gui_core.GuiError("请先读取数据文件。")
        current_ref = self._current_target()
        if (current_ref.app_token, current_ref.table_id) != (
            self._target_info.ref.app_token,
            self._target_info.ref.table_id,
        ):
            raise gui_core.GuiError("目标表已改变，请重新点击“读取目标表”。")
        entered_path = Path(self.source_path_var.get().strip())
        if entered_path.resolve() != self._source_info.path.resolve():
            raise gui_core.GuiError("文件路径已改变，请先点击“读取数据”。")
        entered_sheet = self.sheet_name_var.get().strip() or None
        if (self._source_info.sheet_name or None) != entered_sheet:
            raise gui_core.GuiError("工作表已改变，请先点击“读取数据”。")
        mode = self.mode_var.get()
        columns = self._columns_for_mode(mode)
        if not columns:
            raise gui_core.GuiError("没有可同步的列。")
        index = (
            self.index_var.get().strip()
            if mode in {"incremental", "full", "overwrite"}
            else None
        )
        if mode in {"incremental", "full", "overwrite"} and not index:
            raise gui_core.GuiError("该模式需要选择一个唯一索引列。")
        source = self._source_info
        notes: list[str] = []
        temp_dir = Path(tempfile.mkdtemp(prefix="xtf_gui_"))
        self._temp_dirs.append(temp_dir)
        source_path = source.path
        sheet_name = source.sheet_name
        selective_columns: list[str] | None = None
        if mode in {"append", "clone"}:
            if set(columns) != set(source.columns):
                source_path = gui_core.sanitize_source(
                    source, columns, temp_dir / "source_filtered.xlsx"
                )
                sheet_name = None
                notes.append(f"已创建临时源文件，仅包含选中的 {len(columns)} 列。")
        else:
            if set(columns) != set(source.columns):
                selective_columns = columns
                notes.append(f"使用选择性同步，仅写入 {len(columns)} 个选中列。")
        try:
            batch_size = int(self.batch_size_var.get())
        except ValueError as exc:
            raise gui_core.GuiError("批大小必须是整数。") from exc
        config = gui_core.build_config(
            app_id=self.app_id_var.get().strip(),
            source_path=source_path,
            target=self._target_info.ref,
            mode=mode,
            index_column=index,
            selective_columns=selective_columns,
            verify_remote_writes=bool(self.verify_var.get()),
            create_missing_fields=bool(self.create_missing_var.get()),
            field_strategy=self.field_strategy_var.get(),
            backend=self.backend_var.get(),
            batch_size=batch_size,
            sheet_name=sheet_name,
        )
        config_path = gui_core.write_config(config, temp_dir)
        return config, config_path, notes

    def _preview(self) -> None:
        if self._busy:
            return
        try:
            config, config_path, notes = self._build_config_payload()
            self._secret()
        except Exception as exc:  # noqa: BLE001
            self._show_error("无法生成预览", exc)
            return
        self._preview_config = config
        self._preview_config_path = config_path
        self._preview_signature = self._settings_signature()
        self._plan_payload = None
        self.execute_btn.state(["disabled"])
        self._clear_plan()
        self._clear_log()
        for note in notes:
            self._append_log(note)
        self._set_status("正在生成预览…")
        self._run_cli(
            config_path,
            dry_run=True,
            allow_delete=False,
            on_done=self._preview_done,
        )

    def _preview_done(self, rc: int, stdout: str, error: str | None) -> None:
        if error:
            self._preview_signature = None
            self._show_error("预览失败", gui_core.GuiError(error))
            return
        payload = gui_core.parse_cli_json(stdout)
        if rc != 0 or not payload or not payload.get("ok"):
            self._preview_signature = None
            message = gui_core.payload_error_text(payload) or (
                f"退出码 {rc}：{gui_core.exit_label(rc)}"
            )
            self._append_log(message)
            messagebox.showerror("预览失败", message, parent=self)
            self._set_status("预览失败")
            return
        plan = payload.get("plan") or {}
        self._plan_payload = payload
        self._populate_plan(plan)
        self._set_status("预览完成，尚未写入")
        self.result_var.set("预览完成；确认计划后点击“执行同步”。")
        self.notebook.select(3)

    def _clear_plan(self) -> None:
        for item in self.plan_tree.get_children():
            self.plan_tree.delete(item)
        self.warnings_text.configure(state="normal")
        self.warnings_text.delete("1.0", "end")
        self.warnings_text.configure(state="disabled")
        self.plan_risk_var.set("")
        self.result_var.set("")

    def _populate_plan(self, plan: dict) -> None:
        rows = gui_core.plan_rows(plan)
        for row in rows:
            tags = ("danger",) if row["destructive"] else ()
            self.plan_tree.insert(
                "",
                "end",
                values=(row["kind"], row["count"], row["unit"], row["risk"]),
                tags=tags,
            )
        warnings = plan.get("warnings") or []
        destructive = gui_core.plan_is_destructive(plan)
        clears = gui_core.plan_clears_values(plan)
        risk_parts = []
        if destructive:
            risk_parts.append("危险：计划包含删除/清空动作，需要二次确认。")
        elif clears:
            risk_parts.append("注意：计划会清空部分字段值，但不会删除记录。")
        if not rows:
            risk_parts.append("没有需要写入的动作。")
        self.plan_risk_var.set(" ".join(risk_parts))
        self.warnings_text.configure(state="normal")
        for warning in warnings:
            self.warnings_text.insert("end", f"• {warning}\n")
        self.warnings_text.configure(state="disabled")
        self.execute_btn.state(["!disabled"] if rows else ["disabled"])

    def _execute(self) -> None:
        if self._busy or self._plan_payload is None:
            return
        plan = self._plan_payload.get("plan") or {}
        try:
            self._secret()
        except Exception as exc:  # noqa: BLE001
            self._show_error("无法执行", exc)
            return
        if (
            self._preview_config_path is None
            or not self._preview_config_path.is_file()
            or self._preview_signature is None
            or self._settings_signature() != self._preview_signature
        ):
            messagebox.showinfo(
                "设置已变化",
                "当前设置与上次预览不一致，请重新点击“重新生成预览”。",
                parent=self,
            )
            return
        config_path = self._preview_config_path
        destructive = gui_core.plan_is_destructive(plan) or self.mode_var.get() in {
            "overwrite",
            "clone",
        }
        if destructive:
            answer = simpledialog.askstring(
                "危险操作确认",
                "该计划会删除或清空目标数据。\n" "请输入“删除”两个字继续：",
                parent=self,
            )
            if answer != "删除":
                return
        else:
            risk = (
                "注意：计划会清空部分字段值。\n"
                if gui_core.plan_clears_values(plan)
                else ""
            )
            if not messagebox.askyesno(
                "确认执行",
                f"将执行以下计划：\n\n{gui_core.actions_summary(plan)}\n\n"
                f"{risk}已确认无误并继续？",
                parent=self,
            ):
                return
        self._clear_log()
        self._append_log("开始执行：" + gui_core.actions_summary(plan))
        self._set_status("正在执行同步…")
        self._run_cli(
            config_path,
            dry_run=False,
            allow_delete=destructive,
            on_done=self._execute_done,
        )

    def _execute_done(self, rc: int, stdout: str, error: str | None) -> None:
        self.execute_btn.state(["disabled"])
        self._preview_signature = None
        if error:
            self._show_error("执行失败", gui_core.GuiError(error))
            return
        payload = gui_core.parse_cli_json(stdout)
        text = gui_core.outcome_text(payload)
        if rc == 0 and payload and payload.get("ok"):
            self.result_var.set(text or "执行成功")
            self._set_status("执行完成")
            messagebox.showinfo("同步完成", text or "执行成功", parent=self)
            return
        message = gui_core.payload_error_text(payload) or (
            f"退出码 {rc}：{gui_core.exit_label(rc)}"
        )
        self.result_var.set(text or "执行失败")
        self._set_status("执行失败")
        self._append_log(message)
        messagebox.showerror("同步失败", message, parent=self)

    # ------------------------------------------------------------------
    # CLI process
    # ------------------------------------------------------------------

    def _run_cli(
        self,
        config_path: Path,
        *,
        dry_run: bool,
        allow_delete: bool,
        on_done,
    ) -> None:
        if self._busy:
            return
        try:
            secret = self._secret()
            command = gui_core.cli_command(
                config_path, dry_run=dry_run, allow_delete=allow_delete
            )
        except Exception as exc:  # noqa: BLE001
            self._show_error("无法启动同步", exc)
            return
        self._set_busy(True, "正在调用 XTF CLI…")
        env = gui_core.subprocess_env(secret)
        creationflags = (
            getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
        )

        def worker() -> None:
            proc: subprocess.Popen[str] | None = None
            try:
                proc = subprocess.Popen(
                    command,
                    cwd=str(ROOT),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    env=env,
                    creationflags=creationflags,
                )
                self._proc = proc

                def read_stderr() -> None:
                    assert proc is not None and proc.stderr is not None
                    for line in proc.stderr:
                        text = line.rstrip()
                        if text:
                            self.after(0, lambda text=text: self._append_log(text))

                reader = threading.Thread(target=read_stderr, daemon=True)
                reader.start()
                assert proc.stdout is not None
                stdout = proc.stdout.read()
                rc = proc.wait()
                reader.join(timeout=5)
                self.after(0, lambda: self._finish_cli(on_done, rc, stdout, None))
            except Exception as exc:  # noqa: BLE001
                self.after(
                    0, lambda exc=exc: self._finish_cli(on_done, -1, "", str(exc))
                )
            finally:
                self._proc = None

        threading.Thread(target=worker, daemon=True).start()

    def _finish_cli(self, on_done, rc: int, stdout: str, error: str | None) -> None:
        self._set_busy(False)
        on_done(rc, stdout, error)

    def _open_logs(self) -> None:
        logs = ROOT / "logs"
        logs.mkdir(exist_ok=True)
        try:
            if os.name == "nt":
                os.startfile(str(logs))  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", str(logs)])
        except Exception as exc:  # noqa: BLE001
            self._show_error("无法打开日志目录", exc)

    # ------------------------------------------------------------------
    # Shutdown
    # ------------------------------------------------------------------

    def _on_close(self) -> None:
        if self._busy:
            if not messagebox.askyesno(
                "任务仍在运行",
                "同步或读取仍在进行，确定要退出吗？",
                parent=self,
            ):
                return
            if self._proc is not None and self._proc.poll() is None:
                try:
                    self._proc.terminate()
                except Exception:  # noqa: BLE001
                    pass
        try:
            self._save_settings()
        except Exception:  # noqa: BLE001
            pass
        for folder in self._temp_dirs:
            shutil.rmtree(folder, ignore_errors=True)
        self.destroy()


def main() -> int:
    XTFGui._enable_dpi()
    app = XTFGui()
    if os.environ.get("XTF_GUI_SMOKE") == "1":
        app.after(1200, app.destroy)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
