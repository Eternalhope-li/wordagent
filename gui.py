"""WordAgent 桌面版（PC 端软件入口）— CustomTkinter 现代界面。

运行：python gui.py   （Windows 可直接双击「启动桌面版.bat」）

两种模式：
- 生成新文档：输入需求 -> 自动生成排版好的 docx
- 编辑已有文档：选择 .docx -> 输入修改要求 -> 预览修改计划 -> 确认后落笔（自动备份）
"""
from __future__ import annotations

import os
import queue
import sys
import threading
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

from agent import (
    Config, LLMClient, LLMError, Memory, fill_template, finalize_edit,
    prepare_edit, run_pipeline, __version__,
)

APP_TITLE = f"WordAgent AI 文档助手 v{__version__}"

PLACEHOLDER_GEN = "请输入文档需求，例如：\n写一份《2026 年市场推广方案》，商务风格，包含市场分析、目标人群、渠道策略、预算与效果预估，约 6 个章节。"
PLACEHOLDER_EDIT = "请输入修改要求，例如：\n把标题改成《2026 年度市场推广方案》\n把第二段改得更正式\n删除最后一段\n在“渠道策略”后新增一节：风险分析"

STD_OPTIONS = [
    "auto（AI 自动识别）",
    "gb_9704（国标公文）",
    "paper（学术论文 GB/T 7714）",
    "experiment（实验报告）",
    "business（商务文档）",
    "meeting（会议纪要）",
    "weekly（周报/月报）",
    "general（通用）",
]

STYLE_OPTIONS = [
    "auto（由 AI 自动选择）",
    "business（商务汇报）",
    "report（分析报告）",
    "academic（学术）",
    "creative（创意文案）",
    "default（通用）",
]

# 品牌配色（浅色主题下使用；深色主题组件自动适配）
BRAND = "#2B6DE8"
BRAND_DARK = "#1E3A8A"
BRAND_HOVER = "#1E5BD6"
# 亮/暗双色（(浅色, 深色)），组件随主题自动切换
SIDEBAR = ("#EAF0FA", "#1B2432")
SIDEBAR_TEXT = ("#5B6B8C", "#93A6C4")
SIDEBAR_BTN_TEXT = ("#2B3A55", "#DAE5F8")
SIDEBAR_BTN_HOVER = ("#D5E1F5", "#32466B")
SIDEBAR_BTN_BORDER = ("#B8C7E4", "#46597E")
BTN_SOLID = ("#5B6B8C", "#46536E")
BTN_SOLID_HOVER = ("#46536E", "#38445A")
MUTED_TEXT = ("#7A87A3", "#94A6C2")
LOG_BG = ("#F7F9FD", "#1A2330")
LOG_TEXT = ("#33415C", "#D9E4F8")


class WordAgentApp:
    def __init__(self, root: ctk.CTk):
        self.root = root
        self.config = Config.from_env()
        self.memory = Memory(self.config.memory_file)
        self.events: queue.Queue = queue.Queue()
        self.busy = False
        self.last_output: Path | None = None
        self.ref_files: list[str] = []

        # 恢复上次使用的主题（记忆里持久化）
        saved_theme = self.memory.get_setting("theme", "")
        if saved_theme in ("light", "dark"):
            ctk.set_appearance_mode("dark" if saved_theme == "dark" else "light")
        self._build_ui()
        self._theme_initialized = False
        self._apply_saved_theme()
        self._log(f"欢迎使用 {APP_TITLE}")
        if self.config.api_key:
            self._log(f"✓ 已读取 API Key（模型：{self.config.model}）")
        else:
            self._log("⚠ 未检测到 API Key，请在左侧填写，或创建 .env 文件（参考 .env.example）")
        self.root.after(120, self._poll_events)

    # ================= UI 构建 =================
    def _build_ui(self) -> None:
        self.root.title(APP_TITLE)
        self.root.geometry("1160x820")
        self.root.minsize(980, 700)

        # ---------- 顶部品牌栏 ----------
        header = ctk.CTkFrame(self.root, corner_radius=0, fg_color=BRAND_DARK, height=64)
        header.pack(fill="x")
        header.pack_propagate(False)

        brand = ctk.CTkLabel(header, text="📄 WordAgent", font=ctk.CTkFont("Microsoft YaHei UI", 22, "bold"),
                             text_color="#FFFFFF")
        brand.pack(side="left", padx=(20, 8), pady=12)
        ctk.CTkLabel(header, text="AI 文档生成与编辑助手", font=ctk.CTkFont("Microsoft YaHei UI", 12),
                     text_color="#BFD0F0").pack(side="left", pady=12)
        ctk.CTkLabel(header, text=f"v{__version__}", font=ctk.CTkFont("Consolas", 11),
                     text_color="#93A8D0").pack(side="right", padx=(0, 14))
        self.theme_menu = ctk.CTkOptionMenu(
            header, values=["浅色", "深色"], width=88, corner_radius=14,
            fg_color=BRAND, button_color=BRAND_HOVER, button_hover_color="#1748B8",
            text_color="#FFFFFF", command=self._on_theme_change,
        )
        self.theme_menu.set("浅色")
        self.theme_menu.pack(side="right", padx=(0, 16), pady=14)

        body = ctk.CTkFrame(self.root, fg_color="transparent", corner_radius=0)
        body.pack(fill="both", expand=True)

        # ---------- 左侧边栏 ----------
        self.sidebar = ctk.CTkFrame(body, width=216, corner_radius=0, fg_color=SIDEBAR)
        self.sidebar.pack(side="left", fill="y", padx=0, pady=0)
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(self.sidebar, text="工作模式", font=ctk.CTkFont("Microsoft YaHei UI", 12, "bold"),
                     text_color=SIDEBAR_TEXT).pack(anchor="w", padx=18, pady=(18, 8))

        self.mode_var = "generate"
        self.mode_gen_btn = self._mode_button("✨ 生成新文档", 0, self._switch_generate)
        self.mode_edit_btn = self._mode_button("✏️ 编辑已有文档", 1, self._switch_edit)

        ctk.CTkLabel(self.sidebar, text="模型配置", font=ctk.CTkFont("Microsoft YaHei UI", 12, "bold"),
                     text_color=SIDEBAR_TEXT).pack(anchor="w", padx=18, pady=(20, 8))

        self.key_var = ctk.StringVar(value=self.config.api_key)
        self.model_var = ctk.StringVar(value=self.config.model)

        self._sidebar_field("API Key")
        self.key_entry = ctk.CTkEntry(self.sidebar, textvariable=self.key_var, show="*",
                                      placeholder_text="sk-...", height=32)
        self.key_entry.pack(fill="x", padx=18, pady=(4, 8))

        self._sidebar_field("模型")
        ctk.CTkEntry(self.sidebar, textvariable=self.model_var, height=32).pack(fill="x", padx=18, pady=(4, 8))

        ctk.CTkButton(self.sidebar, text="⚙ API 设置（可保存）", height=32, corner_radius=12,
                      fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER, command=self._open_settings).pack(fill="x", padx=18, pady=(0, 10))

        self._sidebar_field("文档风格")
        self.style_var = ctk.StringVar(value=STYLE_OPTIONS[0])
        ctk.CTkComboBox(self.sidebar, values=STYLE_OPTIONS, variable=self.style_var,
                        state="readonly", height=32).pack(fill="x", padx=18, pady=(4, 8))

        self._sidebar_field("排版标准")
        self.std_var = ctk.StringVar(value=STD_OPTIONS[0])
        ctk.CTkComboBox(self.sidebar, values=STD_OPTIONS, variable=self.std_var,
                        state="readonly", height=32).pack(fill="x", padx=18, pady=(4, 8))

        self._sidebar_field("输出文件夹")
        self.outdir_var = ctk.StringVar(value=str(self.config.output_dir.resolve()))
        outdir_row = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        outdir_row.pack(fill="x", padx=18, pady=(4, 8))
        ctk.CTkEntry(outdir_row, textvariable=self.outdir_var, height=32, width=120).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(outdir_row, text="…", width=36, height=32, corner_radius=10,
                      command=self._browse_dir).pack(side="left", padx=(6, 0))

        # 编辑模式专属（初始隐藏）
        self.edit_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self._sidebar_field("要编辑的文件", parent=self.edit_frame)
        self.file_var = ctk.StringVar(value="")
        ctk.CTkEntry(self.edit_frame, textvariable=self.file_var, height=32,
                     placeholder_text="选择 .docx 文件").pack(fill="x", padx=18, pady=(4, 6))
        ctk.CTkButton(self.edit_frame, text="选择 Word 文档…", height=32, corner_radius=12,
                      fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER, command=self._browse_file).pack(fill="x", padx=18, pady=(0, 6))
        self.overwrite_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(self.edit_frame, text="覆盖原文件（仍会备份）", variable=self.overwrite_var,
                        font=ctk.CTkFont("Microsoft YaHei UI", 12)).pack(anchor="w", padx=18, pady=(2, 6))

        ctk.CTkLabel(self.sidebar, text="常用操作", font=ctk.CTkFont("Microsoft YaHei UI", 12, "bold"),
                     text_color=SIDEBAR_TEXT).pack(anchor="w", padx=18, pady=(14, 8))
        for text, cmd in (
            ("📂 打开输出文件夹", self._open_output_dir),
            ("🕘 最近文档", self._show_recent),
            ("📄 打开生成的文档", self._open_last_doc),
            ("🧠 查看记忆", self._show_memory),
            ("📤 导出日志", self._export_log),
            ("🗑 清空记忆", self._clear_memory),
        ):
            ctk.CTkButton(self.sidebar, text=text, height=34, corner_radius=12,
                          fg_color="transparent", text_color=SIDEBAR_BTN_TEXT, hover_color=SIDEBAR_BTN_HOVER,
                          anchor="w", command=cmd).pack(fill="x", padx=12, pady=2)

        # ---------- 右侧主区 ----------
        main = ctk.CTkFrame(body, fg_color="transparent")
        main.pack(side="left", fill="both", expand=True, padx=20, pady=18)

        # 指令卡片
        card = ctk.CTkFrame(main, corner_radius=16)
        card.pack(fill="both", expand=True, pady=(0, 14))
        self.prompt_title = ctk.CTkLabel(card, text="✍️ 输入指令", font=ctk.CTkFont("Microsoft YaHei UI", 15, "bold"))
        self.prompt_title.pack(anchor="w", padx=18, pady=(14, 4))
        ctk.CTkLabel(card, text="支持自然语言，Ctrl+Enter 快捷执行 · 编辑模式会先预览修改计划，确认后才写入",
                     font=ctk.CTkFont("Microsoft YaHei UI", 11), text_color=MUTED_TEXT).pack(anchor="w", padx=18)

        tpl_row = ctk.CTkFrame(card, fg_color="transparent")
        tpl_row.pack(anchor="w", padx=18, pady=(6, 0))
        ctk.CTkLabel(tpl_row, text="快捷模板：", font=ctk.CTkFont("Microsoft YaHei UI", 11),
                     text_color=MUTED_TEXT).pack(side="left", padx=(0, 6))
        for label in ("实验报告", "会议纪要", "需求文档", "周报", "数据分析"):
            ctk.CTkButton(tpl_row, text=label, width=78, height=26, corner_radius=10,
                          fg_color="transparent", border_width=1, border_color=SIDEBAR_BTN_BORDER,
                          text_color=SIDEBAR_BTN_TEXT, hover_color=SIDEBAR_BTN_HOVER,
                          font=ctk.CTkFont("Microsoft YaHei UI", 11),
                          command=lambda t=label: self._fill_template(t)).pack(side="left", padx=3)

        self.input_text = ctk.CTkTextbox(card, height=150, corner_radius=12,
                                         font=ctk.CTkFont("Microsoft YaHei UI", 13))
        self.input_text.pack(fill="both", expand=True, padx=18, pady=12)
        self.input_text.bind("<Control-Return>", lambda _e: self._run())
        self.input_text.bind("<FocusIn>", lambda _e: self._on_input_focus_in())
        self.input_text.bind("<FocusOut>", lambda _e: self._on_input_focus_out())
        self._set_placeholder(PLACEHOLDER_GEN)
        self._apply_caret_color()

        # 参考文件行（docx 会自动作为模板：格式与章节结构按它来）
        attach_row = ctk.CTkFrame(main, fg_color="transparent")
        attach_row.pack(fill="x", pady=(0, 10))
        self.fill_tpl_var = ctk.BooleanVar(value=False)
        self.fill_tpl_cb = ctk.CTkCheckBox(main, text="📝 按模板填写（在 docx 模板里填内容，格式完全继承模板；需添加 docx 作为参考文件）",
                                           variable=self.fill_tpl_var,
                                           font=ctk.CTkFont("Microsoft YaHei UI", 11),
                                           text_color=SIDEBAR_TEXT)
        self.fill_tpl_cb.pack(anchor="w", pady=(0, 8))
        self.attach_label = ctk.CTkLabel(attach_row, text="📎 参考文件：无（docx=按模板填写；图片/PDF=自动识别模板结构）",
                                         font=ctk.CTkFont("Microsoft YaHei UI", 11),
                                         text_color=SIDEBAR_TEXT, anchor="w")
        self.attach_label.pack(side="left", fill="x", expand=True)
        self.attach_add_btn = ctk.CTkButton(attach_row, text="添加文件…", width=104, height=30,
                                            corner_radius=12, fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER,
                                            font=ctk.CTkFont("Microsoft YaHei UI", 12), command=self._browse_refs)
        self.attach_add_btn.pack(side="right", padx=(6, 0))
        self.attach_clear_btn = ctk.CTkButton(attach_row, text="清空", width=64, height=30, corner_radius=12,
                                              fg_color="transparent", border_width=1, border_color=SIDEBAR_BTN_BORDER,
                                              text_color=SIDEBAR_BTN_TEXT, hover_color=SIDEBAR_BTN_HOVER,
                                              font=ctk.CTkFont("Microsoft YaHei UI", 12), command=self._clear_refs)
        self.attach_clear_btn.pack(side="right")

        # 执行按钮行
        act_row = ctk.CTkFrame(main, fg_color="transparent")
        act_row.pack(fill="x", pady=(0, 12))
        self.run_btn = ctk.CTkButton(act_row, text="🚀 生成 Word 文档", height=46, corner_radius=14,
                                     font=ctk.CTkFont("Microsoft YaHei UI", 15, "bold"),
                                     fg_color=BRAND, hover_color=BRAND_HOVER,
                                     command=self._run)
        self.run_btn.pack(side="left")
        ctk.CTkLabel(act_row, text="  重要文件自动备份，默认另存新文件，安全第一",
                     font=ctk.CTkFont("Microsoft YaHei UI", 12), text_color=SIDEBAR_TEXT).pack(side="left")

        # 状态 + 进度
        status_row = ctk.CTkFrame(main, fg_color="transparent")
        status_row.pack(fill="x", pady=(0, 6))
        self.status_var = ctk.StringVar(value="就绪。")
        ctk.CTkLabel(status_row, textvariable=self.status_var,
                     font=ctk.CTkFont("Microsoft YaHei UI", 12, "bold")).pack(side="left")
        self.progress = ctk.CTkProgressBar(main, height=8, corner_radius=4, progress_color=BRAND)
        self.progress.pack(fill="x", pady=(0, 14))
        self.progress.set(0)

        # 日志卡片
        log_card = ctk.CTkFrame(main, corner_radius=16)
        log_card.pack(fill="both", expand=True)
        ctk.CTkLabel(log_card, text="📋 运行日志", font=ctk.CTkFont("Microsoft YaHei UI", 13, "bold")).pack(anchor="w", padx=18, pady=(12, 6))
        self.log_text = ctk.CTkTextbox(log_card, corner_radius=12, height=170,
                                       font=ctk.CTkFont("Consolas", 11), state="disabled",
                                       fg_color=LOG_BG, text_color=LOG_TEXT)
        self.log_text.pack(fill="both", expand=True, padx=18, pady=(0, 14))
        self._apply_caret_color()

    def _mode_button(self, text: str, row: int, command) -> ctk.CTkButton:
        btn = ctk.CTkButton(self.sidebar, text=text, height=44, corner_radius=12,
                            font=ctk.CTkFont("Microsoft YaHei UI", 13, "bold"),
                            fg_color=BRAND, hover_color=BRAND_HOVER,
                            anchor="w", command=command)
        btn.pack(fill="x", padx=12, pady=3)
        return btn

    def _sidebar_field(self, text: str, parent=None) -> None:
        container = parent or self.sidebar
        ctk.CTkLabel(container, text=text, font=ctk.CTkFont("Microsoft YaHei UI", 11),
                     text_color=SIDEBAR_TEXT).pack(anchor="w", padx=18, pady=(2, 0))

    def _on_theme_change(self, value: str) -> None:
        ctk.set_appearance_mode("dark" if value == "深色" else "light")
        self.memory.set_setting("theme", "dark" if value == "深色" else "light")
        self._refresh_placeholder_color()
        self._apply_caret_color()

    def _apply_saved_theme(self) -> None:
        """把记忆中的主题同步到右上角下拉框（在 _build_ui 之后调用一次）。"""
        if getattr(self, "_theme_initialized", False):
            return
        self._theme_initialized = True
        saved = self.memory.get_setting("theme", "")
        mode = ctk.get_appearance_mode().lower()
        if saved == "dark" or (not saved and mode == "dark"):
            self.theme_menu.set("深色")
            self._refresh_placeholder_color()
            self._apply_caret_color()

    def _apply_caret_color(self) -> None:
        """设置文本光标颜色：浅色主题用深色光标，深色主题用浅色光标，覆盖全部文本框（含弹窗）。"""
        color = "#1F2937" if ctk.get_appearance_mode().lower() == "light" else "#FFFFFF"

        def walk(widget) -> None:
            if hasattr(widget, "_textbox"):
                try:
                    widget._textbox.configure(insertbackground=color, insertwidth=2)
                except Exception:
                    pass
            for child in widget.winfo_children():
                walk(child)

        walk(self.root)

    # ---------- 指令框占位符（CTkTextbox 不支持原生 placeholder） ----------
    def _placeholder_color(self) -> str:
        return "#9AA7BD" if ctk.get_appearance_mode().lower() == "light" else "#5F6B85"

    def _set_placeholder(self, text: str) -> None:
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", text)
        self.input_text.configure(text_color=self._placeholder_color())

    def _current_placeholder(self) -> str:
        return PLACEHOLDER_EDIT if self.mode_var == "edit" else PLACEHOLDER_GEN

    def _is_placeholder(self) -> bool:
        return self.input_text.get("1.0", "end").strip() == self._current_placeholder().strip()

    def _refresh_placeholder_color(self) -> None:
        if self._is_placeholder():
            self.input_text.configure(text_color=self._placeholder_color())

    def _on_input_focus_in(self) -> None:
        if self._is_placeholder():
            self.input_text.delete("1.0", "end")
            self.input_text.configure(text_color=("gray10" if ctk.get_appearance_mode().lower() == "light" else "white"))

    def _on_input_focus_out(self) -> None:
        if not self.input_text.get("1.0", "end").strip():
            self._set_placeholder(self._current_placeholder())

    # ================= 模式切换 =================
    def _set_mode(self, mode: str) -> None:
        self.mode_var = mode
        if mode == "edit":
            # 编辑模式也支持参考文件（作为格式/数据/要求依据）
            self.fill_tpl_cb.configure(state="disabled")
            self._refresh_attach_label_edit()
            self.mode_gen_btn.configure(fg_color="transparent", text_color=SIDEBAR_BTN_TEXT,
                                        hover_color=SIDEBAR_BTN_HOVER, border_width=1, border_color=SIDEBAR_BTN_BORDER)
            self.mode_edit_btn.configure(fg_color=BRAND, text_color="#FFFFFF",
                                         hover_color=BRAND_HOVER, border_width=0)
            self.edit_frame.pack(fill="x", pady=(4, 0))
            self.prompt_title.configure(text="✍️ 输入修改要求")
            self.run_btn.configure(text="📝 生成修改计划")
            if self._is_placeholder() or not self.input_text.get("1.0", "end").strip():
                self._set_placeholder(PLACEHOLDER_EDIT)
            self.status_var.set("编辑模式：选择文件并输入修改要求。")
        else:
            self.attach_add_btn.configure(state="normal")
            self.attach_clear_btn.configure(state="normal")
            self.fill_tpl_cb.configure(state="normal")
            self._refresh_attach_label()
            self.mode_edit_btn.configure(fg_color="transparent", text_color=SIDEBAR_BTN_TEXT,
                                         hover_color=SIDEBAR_BTN_HOVER, border_width=1, border_color=SIDEBAR_BTN_BORDER)
            self.mode_gen_btn.configure(fg_color=BRAND, text_color="#FFFFFF",
                                        hover_color=BRAND_HOVER, border_width=0)
            self.edit_frame.pack_forget()
            self.prompt_title.configure(text="✍️ 输入指令")
            self.run_btn.configure(text="🚀 生成 Word 文档")
            if self._is_placeholder() or not self.input_text.get("1.0", "end").strip():
                self._set_placeholder(PLACEHOLDER_GEN)
            self.status_var.set("生成模式：输入需求即可。")

    def _switch_generate(self) -> None:
        if not self.busy:
            self._set_mode("generate")

    def _switch_edit(self) -> None:
        if not self.busy:
            self._set_mode("edit")

    # ================= 动作 =================
    # ================= API 设置（持久化到 .env） =================
    def _open_settings(self) -> None:
        win = ctk.CTkToplevel(self.root)
        win.title("API 设置")
        win.geometry("440x640")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        ctk.CTkLabel(win, text="接口设置（保存后写入 .env，重启仍生效）",
                     font=ctk.CTkFont("Microsoft YaHei UI", 14, "bold"),
                     text_color=BRAND_DARK).pack(anchor="w", padx=24, pady=(18, 10))

        def field(label: str) -> ctk.CTkEntry:
            ctk.CTkLabel(win, text=label, font=ctk.CTkFont("Microsoft YaHei UI", 12),
                         text_color=MUTED_TEXT).pack(anchor="w", padx=24)
            e = ctk.CTkEntry(win, height=32)
            e.pack(fill="x", padx=24, pady=(3, 8))
            return e

        key_entry = field("API Key")
        key_entry.insert(0, self.config.api_key)
        url_entry = field("接口地址 Base URL")
        url_entry.insert(0, self.config.base_url)
        model_entry = field("模型名称")
        model_entry.insert(0, self.config.model)
        temp_entry = field("温度 (0-2，越高越有创造性)")
        temp_entry.insert(0, str(self.config.temperature))
        timeout_entry = field("请求超时（秒）")
        timeout_entry.insert(0, str(self.config.request_timeout))
        concurrency_entry = field("并发数（1-8）")
        concurrency_entry.insert(0, str(self.config.concurrency))
        max_tokens_entry = field("单次输出上限 max_tokens（推理模型建议 8192-32768）")
        max_tokens_entry.insert(0, str(self.config.max_tokens))

        status = ctk.CTkLabel(win, text="", font=ctk.CTkFont("Microsoft YaHei UI", 12), text_color="#C0392B")
        status.pack(anchor="w", padx=24)

        def save() -> None:
            try:
                key = key_entry.get().strip()
                url = url_entry.get().strip().rstrip("/")
                model = model_entry.get().strip()
                temperature = float(temp_entry.get().strip())
                timeout = int(timeout_entry.get().strip())
                concurrency = max(1, min(8, int(concurrency_entry.get().strip())))
                max_tokens = max(1024, min(32768, int(max_tokens_entry.get().strip())))
            except ValueError:
                status.configure(text="⚠ 数字格式不正确，请检查温度/超时/并发/max_tokens")
                return
            if not key or not url or not model:
                status.configure(text="⚠ API Key、接口地址、模型不能为空")
                return
            self._persist_env({
                "DEEPSEEK_API_KEY": key,
                "DEEPSEEK_BASE_URL": url,
                "DEEPSEEK_MODEL": model,
                "DEEPSEEK_TEMPERATURE": f"{temperature:g}",
                "DEEPSEEK_TIMEOUT": str(timeout),
                "DEEPSEEK_CONCURRENCY": str(concurrency),
                "DEEPSEEK_MAX_TOKENS": str(max_tokens),
            })
            self.config = Config.from_env()
            self.key_var.set(self.config.api_key)
            self.model_var.set(self.config.model)
            self._log("✓ API 设置已保存到 .env")
            win.destroy()

        row = ctk.CTkFrame(win, fg_color="transparent")
        row.pack(fill="x", padx=24, pady=(6, 18))
        ctk.CTkButton(row, text="保存设置", width=120, height=36, corner_radius=12,
                      fg_color=BRAND, hover_color=BRAND_HOVER, command=save).pack(side="left")
        ctk.CTkButton(row, text="取消", width=90, height=36, corner_radius=12,
                      fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER, command=win.destroy).pack(side="left", padx=(10, 0))

    def _persist_env(self, updates: dict[str, str]) -> None:
        # 安装版（PyInstaller 冻结）运行时 .env 位于程序目录；开发版在 cwd
        base = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path.cwd()
        path = base / ".env"
        lines = path.read_text(encoding="utf-8").splitlines() if path.exists() else []
        seen: set[str] = set()
        out: list[str] = []
        for line in lines:
            stripped = line.strip()
            if stripped and not stripped.startswith("#") and "=" in stripped:
                k = stripped.split("=", 1)[0].strip()
                if k in updates:
                    out.append(f"{k}={updates[k]}")
                    seen.add(k)
                    continue
            out.append(line)
        for k, v in updates.items():
            if k not in seen:
                out.append(f"{k}={v}")
        path.write_text("\n".join(out) + "\n", encoding="utf-8")

    def _browse_dir(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.outdir_var.get() or str(Path.cwd()))
        if chosen:
            self.outdir_var.set(chosen)

    def _browse_file(self) -> None:
        chosen = filedialog.askopenfilename(
            title="选择要编辑的 Word 文档",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")],
        )
        if chosen:
            self.file_var.set(chosen)

    def _browse_refs(self) -> None:
        chosen = filedialog.askopenfilenames(
            title="选择参考文件（图片/文本/csv/docx）",
            filetypes=[
                ("常用参考文件", "*.png;*.jpg;*.jpeg;*.bmp;*.webp;*.txt;*.md;*.csv;*.log;*.docx"),
                ("图片", "*.png;*.jpg;*.jpeg;*.bmp;*.webp;*.gif"),
                ("文本", "*.txt;*.md;*.csv;*.log;*.json"),
                ("Word 文档", "*.docx"),
                ("所有文件", "*.*"),
            ],
        )
        for f in chosen:
            if f not in self.ref_files:
                self.ref_files.append(f)
        if chosen:
            self._refresh_attach_label()
            self._log(f"📎 已添加 {len(chosen)} 个参考文件（共 {len(self.ref_files)} 个）")

    def _clear_refs(self) -> None:
        self.ref_files = []
        self._refresh_attach_label()
        self._log("📎 已清空参考文件。")

    def _refresh_attach_label(self) -> None:
        if not self.ref_files:
            self.attach_label.configure(text="📎 参考文件：无（docx=按模板填写；图片/PDF=自动识别模板结构）")
            return
        names = "、".join(Path(f).name for f in self.ref_files)
        if len(names) > 46:
            names = names[:43] + "…"
        self.attach_label.configure(text=f"📎 参考文件：{names}（{len(self.ref_files)} 个）")

    def _refresh_attach_label_edit(self) -> None:
        if not self.ref_files:
            self.attach_label.configure(text="📎 参考文件：可选（修改时结合参考格式/数据）")
            return
        names = "、".join(Path(f).name for f in self.ref_files)
        if len(names) > 46:
            names = names[:43] + "…"
        self.attach_label.configure(text=f"📎 参考文件：{names}（{len(self.ref_files)} 个，修改时结合）")

    def _fill_template(self, kind: str) -> None:
        templates = {
            "实验报告": "请写一份《XX 实验报告》：主题为 ______，包含实验目的、实验原理、实验环境与器材、"
                       "实验步骤、数据记录（用表格）、结果分析、误差分析、实验结论，数据要定量。",
            "会议纪要": "请写一份《XX 会议纪要》：包含会议基本信息（时间/地点/参会人）、议题与讨论要点、"
                       "决议事项、待办任务表格（事项/负责人/截止时间）。",
            "需求文档": "请写一份《XX 产品需求文档》，商务风格：背景与目标、用户与场景、功能需求（表格+优先级）、"
                       "非功能需求、里程碑与验收标准。",
            "周报": "请写一份本周工作周报：本周工作进展（分条）、关键数据（表格）、问题与风险、下周计划。",
            "数据分析": "请写一份《XX 数据分析报告》：数据概览（表格）、关键发现、分维度分析、结论与建议。",
        }
        text = templates.get(kind, "")
        if not text:
            return
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", text)
        self.input_text.configure(text_color=("gray10" if ctk.get_appearance_mode().lower() == "light" else "white"))
        self.status_var.set(f"已填入「{kind}」模板，把 XX/______ 替换成你的内容即可。")
        self.input_text.focus_set()

    def _open_output_dir(self) -> None:
        folder = Path(self.outdir_var.get())
        folder.mkdir(parents=True, exist_ok=True)
        os.startfile(str(folder))

    def _open_last_doc(self) -> None:
        if self.last_output and self.last_output.exists():
            os.startfile(str(self.last_output))
        else:
            messagebox.showinfo("提示", "还没有生成/修改过文档。")

    def _show_recent(self) -> None:
        """最近文档：列出记忆中的生成/编辑结果，一键打开或打开所在文件夹。"""
        items = [
            e for e in self.memory.entries
            if e.get("role") == "assistant" and e.get("file") and Path(str(e["file"])).exists()
        ][-12:][::-1]
        win = ctk.CTkToplevel(self.root)
        win.title("最近文档")
        win.geometry("760x520")
        win.transient(self.root)
        ctk.CTkLabel(win, text="🕘 最近文档（点击文件名打开，右侧按钮打开所在文件夹）",
                     font=ctk.CTkFont("Microsoft YaHei UI", 15, "bold")).pack(anchor="w", padx=18, pady=(16, 8))
        if not items:
            ctk.CTkLabel(win, text="（还没有生成或编辑过文档）",
                         font=ctk.CTkFont("Microsoft YaHei UI", 12), text_color=MUTED_TEXT).pack(anchor="w", padx=18, pady=20)
            return
        box = ctk.CTkScrollableFrame(win, corner_radius=12)
        box.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        for e in items:
            f = Path(str(e["file"]))
            row = ctk.CTkFrame(box, corner_radius=10)
            row.pack(fill="x", pady=4)
            meta = f"[{e.get('time', '')}] {e.get('title', f.name)}"
            ctk.CTkButton(row, text=meta, anchor="w", height=34, corner_radius=10,
                          fg_color="transparent", text_color=SIDEBAR_BTN_TEXT,
                          hover_color=SIDEBAR_BTN_HOVER,
                          command=lambda fp=f: self._open_path(fp)).pack(side="left", fill="x", expand=True, padx=(6, 0), pady=4)
            ctk.CTkButton(row, text="📂", width=44, height=32, corner_radius=10,
                          fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER,
                          command=lambda fp=f: self._open_folder_of(fp)).pack(side="right", padx=(4, 6), pady=4)

    @staticmethod
    def _open_path(path: Path) -> None:
        try:
            os.startfile(str(path))
        except Exception:
            pass

    @staticmethod
    def _open_folder_of(path: Path) -> None:
        try:
            os.startfile(str(path.parent))
        except Exception:
            pass

    def _export_log(self) -> None:
        """把当前运行日志导出为 .txt 文件。"""
        from datetime import datetime as _dt
        default = f"WordAgent日志_{_dt.now().strftime('%Y%m%d_%H%M%S')}.txt"
        target = filedialog.asksaveasfilename(
            title="导出运行日志", defaultextension=".txt", initialfile=default,
            filetypes=[("文本文件", "*.txt")],
        )
        if not target:
            return
        try:
            Path(target).write_text(self.log_text.get("1.0", "end").strip() + "\n", encoding="utf-8-sig")
            self._log(f"📤 日志已导出：{target}")
        except OSError as exc:
            messagebox.showerror("导出失败", f"无法写入日志文件：{exc}")

    def _show_memory(self) -> None:
        win = ctk.CTkToplevel(self.root)
        win.title("上下文记忆")
        win.geometry("720x460")
        win.transient(self.root)
        ctk.CTkLabel(win, text="🧠 上下文记忆（历史指令与结果）",
                     font=ctk.CTkFont("Microsoft YaHei UI", 15, "bold")).pack(anchor="w", padx=18, pady=(16, 6))
        box = ctk.CTkTextbox(win, corner_radius=12, font=ctk.CTkFont("Microsoft YaHei UI", 12))
        box.pack(fill="both", expand=True, padx=18, pady=12)
        box.insert("1.0", self.memory.history_text())
        box.configure(state="disabled")

    def _clear_memory(self) -> None:
        if messagebox.askyesno("确认", "确定清空全部上下文记忆吗？"):
            self.memory.clear()
            self._log("已清空上下文记忆。")

    # ================= 主入口 =================
    def _run(self) -> None:
        if self.busy:
            return
        command = self.input_text.get("1.0", "end").strip()
        if not command or self._is_placeholder():
            messagebox.showwarning("提示", "请先输入指令。")
            return
        key = self.key_var.get().strip() or self.config.api_key
        if not key:
            messagebox.showwarning(
                "缺少 API Key",
                "请填写 DeepSeek API Key。\n\n也可以在工作目录创建 .env 文件：\nDEEPSEEK_API_KEY=sk-xxx\nDEEPSEEK_MODEL=deepseek-v4-flash",
            )
            return
        self.config.api_key = key
        self.config.model = self.model_var.get().strip() or self.config.model

        if self.mode_var == "edit":
            src = Path(self.file_var.get().strip())
            if not src.exists():
                messagebox.showwarning("提示", "请选择要编辑的 .docx 文件。")
                return
            self._set_busy(True, "正在读取文档并生成修改计划……")
            threading.Thread(target=self._worker_edit_prepare, args=(src, command, list(self.ref_files)), daemon=True).start()
        else:
            style_override = None
            style_raw = self.style_var.get()
            if not style_raw.startswith("auto"):
                style_override = style_raw.split("（")[0]
            standard_override = None
            std_raw = self.std_var.get()
            if not std_raw.startswith("auto"):
                standard_override = std_raw.split("（")[0]
            fill_tpl = self.fill_tpl_var.get()
            output_dir = Path(self.outdir_var.get().strip()) or self.config.output_dir
            self._set_busy(True, "正在解析需求……")
            threading.Thread(target=self._worker_generate,
                             args=(command, style_override, standard_override, output_dir,
                                   list(self.ref_files), fill_tpl),
                             daemon=True).start()

    def _set_busy(self, busy: bool, status: str = "") -> None:
        self.busy = busy
        self.run_btn.configure(state="disabled" if busy else "normal")
        for btn in (self.mode_gen_btn, self.mode_edit_btn):
            btn.configure(state="disabled" if busy else "normal")
        if status:
            self.status_var.set(status)

    # ================= Worker（后台线程） =================
    def _worker_generate(self, command: str, style_override: str | None,
                         standard_override: str | None, output_dir: Path,
                         ref_files: list[str] | None = None, fill_tpl: bool = False) -> None:
        def log(msg: str) -> None:
            self.events.put(("log", msg))
        try:
            if fill_tpl:
                tpl = next((Path(f) for f in (ref_files or []) if str(f).lower().endswith(".docx")), None)
                if tpl is None:
                    self.events.put(("error", "「在模板中填写」需要先添加一个 .docx 模板文件作为参考文件。"))
                    return
                llm = LLMClient(self.config)
                self.config.output_dir_override = output_dir
                path = fill_template(tpl, command, self.config, self.memory, llm=llm, log=log)
                self.config.output_dir_override = None
                self.events.put(("done_generate", str(path)))
                return
            llm = LLMClient(self.config)
            path = run_pipeline(command, self.config, self.memory, log=log,
                                style_override=style_override, output_dir_override=output_dir,
                                reference_files=ref_files, llm=llm,
                                confirm_plan=self._request_plan_confirm,
                                standard_override=standard_override)
            if path is None:
                self.events.put(("cancelled",))
                return
            if hasattr(llm, "usage_text"):
                log(f"⚙ {llm.usage_text()}")
            self.events.put(("done_generate", str(path)))
        except (LLMError, Exception) as exc:  # noqa: BLE001
            self.events.put(("error", str(exc)))

    def _worker_edit_prepare(self, src: Path, instruction: str, ref_files: list[str] | None = None) -> None:
        def log(msg: str) -> None:
            self.events.put(("log", msg))
        try:
            llm = LLMClient(self.config)
            state = prepare_edit(src, instruction, self.config, self.memory, llm=llm, log=log,
                                 reference_files=ref_files)
            if hasattr(llm, "usage_text"):
                log(f"⚙ {llm.usage_text()}")
            self.events.put(("preview", state))
        except (LLMError, Exception) as exc:  # noqa: BLE001
            self.events.put(("error", str(exc)))

    def _worker_edit_finalize(self, state: dict) -> None:
        def log(msg: str) -> None:
            self.events.put(("log", msg))

        def resolve_ambiguous(op: dict, candidates: list[str]) -> str | None:
            """目标有歧义时弹窗让用户选择（主线程展示，阻塞等待）。"""
            result: list[str | None] = [None]
            evt = threading.Event()

            def _show() -> None:
                win = ctk.CTkToplevel(self.root)
                win.title("目标匹配到多个位置")
                win.geometry("640x460")
                win.transient(self.root)
                win.grab_set()
                ctk.CTkLabel(win, text=f"指令「{str(op.get('target', ''))[:40]}」匹配到多个段落，请选择：",
                             font=ctk.CTkFont("Microsoft YaHei UI", 13, "bold")).pack(anchor="w", padx=18, pady=(16, 8))
                box = ctk.CTkTextbox(win, corner_radius=12, font=ctk.CTkFont("Microsoft YaHei UI", 12))
                box.pack(fill="both", expand=True, padx=18, pady=10)
                for i, cand in enumerate(candidates[:6], 1):
                    box.insert("end", f"{i}. {cand}\n")
                box.configure(state="disabled")
                entry = ctk.CTkEntry(win, placeholder_text="输入序号后点确认（回车=取消）", height=34)
                entry.pack(fill="x", padx=18)
                entry.bind("<Return>", lambda e: _confirm())
                row = ctk.CTkFrame(win, fg_color="transparent")
                row.pack(fill="x", padx=18, pady=(10, 16))

                def _confirm() -> None:
                    ans = entry.get().strip()
                    if ans.isdigit() and 1 <= int(ans) <= len(candidates[:6]):
                        result[0] = candidates[:6][int(ans) - 1]
                    evt.set()
                    win.destroy()

                def _cancel() -> None:
                    result[0] = None
                    evt.set()
                    win.destroy()

                ctk.CTkButton(row, text="✓ 确认", width=110, height=36, corner_radius=12,
                              fg_color="#2E9E5B", hover_color="#238049",
                              font=ctk.CTkFont("Microsoft YaHei UI", 13), command=_confirm).pack(side="left")
                ctk.CTkButton(row, text="✕ 取消", width=100, height=36, corner_radius=12,
                              fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER,
                              font=ctk.CTkFont("Microsoft YaHei UI", 13), command=_cancel).pack(side="left", padx=10)

            self.root.after(0, _show)
            evt.wait()
            return result[0]

        def on_warnings(issues: list[str]) -> bool:
            """结构校验有风险时弹窗询问是否继续（阻塞等待用户选择）。"""
            result: list[bool] = [False]
            evt = threading.Event()

            def _show() -> None:
                win = ctk.CTkToplevel(self.root)
                win.title("结构校验风险")
                win.geometry("620x400")
                win.transient(self.root)
                win.grab_set()
                ctk.CTkLabel(win, text="⚠ 结构校验发现以下风险：",
                             font=ctk.CTkFont("Microsoft YaHei UI", 14, "bold")).pack(anchor="w", padx=18, pady=(16, 8))
                box = ctk.CTkTextbox(win, corner_radius=12, font=ctk.CTkFont("Microsoft YaHei UI", 12))
                box.pack(fill="both", expand=True, padx=18, pady=10)
                for issue in issues:
                    box.insert("end", f"· {issue}\n")
                box.insert("end", "\n仍要继续保存吗？原文件不会受影响（会自动备份）。")
                box.configure(state="disabled")
                row = ctk.CTkFrame(win, fg_color="transparent")
                row.pack(fill="x", padx=18, pady=(10, 16))

                def _yes() -> None:
                    result[0] = True
                    evt.set()
                    win.destroy()

                def _no() -> None:
                    result[0] = False
                    evt.set()
                    win.destroy()

                ctk.CTkButton(row, text="✓ 继续保存", width=120, height=36, corner_radius=12,
                              fg_color="#2E9E5B", hover_color="#238049",
                              font=ctk.CTkFont("Microsoft YaHei UI", 13), command=_yes).pack(side="left")
                ctk.CTkButton(row, text="✕ 取消", width=100, height=36, corner_radius=12,
                              fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER,
                              font=ctk.CTkFont("Microsoft YaHei UI", 13), command=_no).pack(side="left", padx=10)

            self.root.after(0, _show)
            evt.wait()
            return result[0]

        try:
            output_dir = Path(self.outdir_var.get().strip()) or self.config.output_dir
            path = finalize_edit(state, self.config, self.memory, output_dir=output_dir,
                                 save_as_new=not self.overwrite_var.get(), log=log,
                                 resolve_ambiguous=resolve_ambiguous, on_warnings=on_warnings)
            llm = state.get("llm")
            if hasattr(llm, "usage_text"):
                log(f"⚙ {llm.usage_text()}")
            self.events.put(("done_edit", str(path)))
        except (LLMError, Exception) as exc:  # noqa: BLE001
            self.events.put(("error", str(exc)))

    # ================= 事件轮询 =================
    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "plan_confirm":
                    self._show_plan_confirm(payload)
                elif kind == "cancelled":
                    self._set_busy(False)
                    self.status_var.set("已取消生成。")
                    self._log("已取消：大纲未确认，未生成文档。")
                elif kind == "done_generate":
                    self._finish_ok(payload, "生成")
                elif kind == "done_edit":
                    self._finish_ok(payload, "编辑")
                elif kind == "preview":
                    self._show_preview(payload)
                elif kind == "error":
                    self._finish_error(payload)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_events)

    def _request_plan_confirm(self, plan: dict) -> bool:
        """后台线程调用：把大纲发给主线程弹窗确认，阻塞等待结果。"""
        self._plan_evt = threading.Event()
        self._plan_ok = False
        self.events.put(("plan_confirm", plan))
        self._plan_evt.wait()
        return self._plan_ok

    def _show_plan_confirm(self, plan: dict) -> None:
        """大纲预览：满意才生成正文，避免不满意时整篇重来烧 token。"""
        win = ctk.CTkToplevel(self.root)
        win.title("大纲预览 — 确认后再生成正文")
        win.geometry("720x600")
        win.transient(self.root)
        win.grab_set()

        ctk.CTkLabel(win, text="📋 文档大纲预览", font=ctk.CTkFont("Microsoft YaHei UI", 16, "bold")).pack(anchor="w", padx=18, pady=(16, 6))
        style_label = {
            "business": "商务", "report": "报告", "academic": "学术",
            "creative": "创意文案", "default": "通用"}.get(plan.get("style"), "通用")
        ctk.CTkLabel(win, text=f"标题：{plan.get('title', '')}　|　风格：{style_label}　|　目录：{'是' if plan.get('toc') else '否'}",
                     font=ctk.CTkFont("Microsoft YaHei UI", 12), text_color=MUTED_TEXT).pack(anchor="w", padx=18)

        box = ctk.CTkTextbox(win, corner_radius=12, font=ctk.CTkFont("Microsoft YaHei UI", 12))
        box.pack(fill="both", expand=True, padx=18, pady=12)
        text = ""
        for sec in plan.get("sections", []):
            indent = "　　" * (max(1, int(sec.get("level", 1))) - 1)
            text += f"{indent}· {sec.get('heading', '')}\n"
        if not text:
            text = "（大纲为空）"
        box.insert("1.0", text)
        box.configure(state="disabled")

        ctk.CTkLabel(win, text="不满意？直接点取消，回到输入框补充章节/风格要求后重试，不消耗正文 token。",
                     font=ctk.CTkFont("Microsoft YaHei UI", 11), text_color=MUTED_TEXT).pack(anchor="w", padx=18)

        def on_ok() -> None:
            self._plan_ok = True
            self._plan_evt.set()
            win.destroy()

        def on_cancel() -> None:
            self._plan_ok = False
            self._plan_evt.set()
            win.destroy()

        row = ctk.CTkFrame(win, fg_color="transparent")
        row.pack(fill="x", padx=18, pady=(0, 16))
        ctk.CTkButton(row, text="✓ 按此大纲生成", width=150, height=38, corner_radius=12,
                      fg_color="#2E9E5B", hover_color="#238049",
                      font=ctk.CTkFont("Microsoft YaHei UI", 13, "bold"), command=on_ok).pack(side="left")
        ctk.CTkButton(row, text="✕ 取消修改指令", width=130, height=38, corner_radius=12,
                      fg_color="transparent", border_width=1,
                      font=ctk.CTkFont("Microsoft YaHei UI", 13), command=on_cancel).pack(side="left", padx=10)

    def _show_preview(self, state: dict) -> None:
        plan = state["plan"]
        win = ctk.CTkToplevel(self.root)
        win.title("修改计划预览 — 确认后才会写入")
        win.geometry("760x560")
        win.transient(self.root)
        win.grab_set()

        ctk.CTkLabel(win, text="🔎 修改计划预览", font=ctk.CTkFont("Microsoft YaHei UI", 16, "bold")).pack(anchor="w", padx=18, pady=(16, 4))
        ctk.CTkLabel(win, text=f"说明：{plan.get('summary', '')} · 共 {len(plan['operations'])} 项操作",
                     font=ctk.CTkFont("Microsoft YaHei UI", 12), text_color=MUTED_TEXT).pack(anchor="w", padx=18)

        text = ""
        for i, op in enumerate(plan["operations"], 1):
            text += f"{i}. [{op['op']}] 目标「{op['target']}」\n"
            if op.get("new_text"):
                text += f"   → 新内容：{op['new_text']}\n"
            if op.get("style"):
                text += f"   → 样式：{op['style']}\n"
            text += "\n"
        text += "⚠ 原文件将自动备份到 output/backups/，默认另存为新文件，原文件不受影响。"

        box = ctk.CTkTextbox(win, corner_radius=12, font=ctk.CTkFont("Microsoft YaHei UI", 12))
        box.pack(fill="both", expand=True, padx=18, pady=12)
        box.insert("1.0", text)
        box.configure(state="disabled")

        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill="x", padx=18, pady=(0, 16))

        def on_confirm() -> None:
            win.destroy()
            self._set_busy(True, "正在应用修改（备份 → 落笔 → 校验）……")
            self.progress.set(0.4)
            threading.Thread(target=self._worker_edit_finalize, args=(state,), daemon=True).start()

        def on_cancel() -> None:
            win.destroy()
            self._set_busy(False)
            self.status_var.set("已取消，未做任何修改。")
            self._log("已取消编辑，未做任何修改。")

        ctk.CTkButton(btn_row, text="✓ 确认修改", width=140, height=38, corner_radius=12,
                      fg_color="#2E9E5B", hover_color="#238049",
                      font=ctk.CTkFont("Microsoft YaHei UI", 13, "bold"), command=on_confirm).pack(side="left")
        ctk.CTkButton(btn_row, text="✕ 取消", width=110, height=38, corner_radius=12,
                      fg_color="transparent", border_width=1,
                      font=ctk.CTkFont("Microsoft YaHei UI", 13), command=on_cancel).pack(side="left", padx=10)

    # ================= 收尾 =================
    def _finish_ok(self, path: str, action: str) -> None:
        self._set_busy(False)
        self.progress.set(1.0)
        self.status_var.set(f"✔ {action}完成")
        self.last_output = Path(path)
        if messagebox.askyesno(f"{action}完成", f"文档已保存到：\n{path}\n\n是否立即打开？"):
            os.startfile(str(self.last_output))

    def _finish_error(self, message: str) -> None:
        self._set_busy(False)
        self.status_var.set("✖ 操作失败")
        self._log(f"✖ 错误：{message}")
        messagebox.showerror("操作失败", message)

    def _log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self._bump_progress(message)

    def _bump_progress(self, message: str) -> None:
        """根据日志关键字平滑推进进度条（生成/编辑通用）。"""
        pairs = [
            ("① 解析需求", 0.08), ("② 逐节生成", 0.12), ("③ 排版", 0.85),
            ("✔ 生成完成", 1.0),
            ("① 读取文档", 0.05), ("② 解析编辑指令", 0.12), ("③ 校验并应用", 0.4),
            ("④ AI 复核", 0.7), ("✔ 编辑完成", 1.0),
        ]
        if self.busy:
            import re
            m = re.search(r"\[(\d+)/(\d+)\] 撰写", message)
            if m:
                self.progress.set(0.12 + 0.68 * int(m.group(1)) / max(1, int(m.group(2))))
                return
            for keyword, value in pairs:
                if keyword in message:
                    self.progress.set(value)
                    return


def main() -> int:
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    # 单实例：重复启动时提示并退出
    try:
        import ctypes as _ct
        _mutex = _ct.windll.kernel32.CreateMutexW(None, False, "WordAgent_SingleInstance")
        if _ct.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            _ct.windll.user32.MessageBoxW(None, "WordAgent 已在运行，请勿重复启动。", "WordAgent", 0x40)
            return 1
    except Exception:
        pass
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    WordAgentApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())


