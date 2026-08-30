"""WordAgent 桌面版 v1.8.0 — 聊天式 AI 文档助手（PC 端软件）。

交互方式：
- 像聊天一样直接输入要求，AI 自动判断「生成新文档 / 按模板填写 / 编辑已有文档」。
- 点「📎 附件」选择 .docx（要编辑的文档或模板）以及参考文件（图片 / PDF / 文本）。
- 每次生成与修改的关键信息都会记入本地记忆，形成跨对话的长期上下文。

运行：python gui.py
"""
from __future__ import annotations

import os
import queue
import shutil
import sys
import threading
import time
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

from agent import (
    Config, LLMClient, LLMError, Memory, fill_template, finalize_edit,
    prepare_edit, run_pipeline, __version__,
)

APP_TITLE = f"WordAgent AI 文档助手 v{__version__}"
VERSION_LABEL = f"v{__version__}"

PLACEHOLDER = (
    "输入你的文档要求，例如：\n"
    "· 写一份《数据库课程实验报告》，含实验目的、实验环境、实验步骤、结果与分析\n"
    "· 或点「📎 附件」添加 .docx 后说：把标题改成《XXX》，第二节改得更正式"
)

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

CHIP_PROMPTS = [
    ("📊 实验报告", "写一份《科目：实验报告》，要求包含：实验目的、实验环境、实验原理、实验步骤、实验数据记录、结果分析与结论，数据需合理详实。"),
    ("📝 会议纪要", "写一份会议纪要，包含：会议主题、时间地点、参会人员、会议内容、决议事项、后续行动项与责任人。"),
    ("📦 需求文档", "写一份软件需求文档（PRD），包含：项目背景、目标用户、功能需求、非功能需求、交互流程、验收标准。"),
    ("🗓 周报", "写一份本周工作周报，包含：本周完成事项、关键进展、遇到的问题与解决、下周工作计划。"),
    ("📈 数据分析", "写一份数据分析报告，包含：数据来源、分析指标、趋势分析、异常说明、结论与建议。"),
    ("✏️ 编辑/套模板", "（先点 📎 添加 .docx 再发送）把标题改成……，第二节改得更正式……"),
]

# ---- 主题配色（亮色, 深色） ----
BRAND = "#2B6DE8"
BRAND_HOVER = "#1E5BD6"
USER_BUBBLE = ("#2B6DE8", "#2563EB")
ASSIST_BUBBLE = ("#F1F5FB", "#253047")
CHAT_BG = ("#F4F6FA", "#0F131B")
APP_BG = ("#F4F6FA", "#0F131B")
INPUT_BG = ("#FFFFFF", "#2B3A52")
SIDEBAR = ("#FFFFFF", "#151B26")
SIDEBAR_TEXT = ("#5B6B8C", "#93A6C4")
SIDEBAR_BTN_TEXT = ("#2B3A55", "#DAE5F8")
SIDEBAR_BTN_HOVER = ("#D5E1F5", "#32466B")
SIDEBAR_BTN_BORDER = ("#B8C7E4", "#46597E")
BTN_SOLID = ("#5B6B8C", "#46536E")
BTN_SOLID_HOVER = ("#46536E", "#38445A")
MUTED_TEXT = ("#7A87A3", "#94A6C2")
ERROR_TEXT = ("#C0392B", "#FF8A80")
GREEN = "#2E9E5B"
GREEN_HOVER = "#238049"

FILL_KEYWORDS = ("填写", "填充", "填一下", "模板", "模版", "套用", "照着", "补全", "填空")
# 强编辑词：没有附文档时也能判断“想编辑”
STRONG_EDIT = ("修改", "编辑", "删除", "删掉", "替换", "改成", "改为", "更正", "重写",
               "改一下", "改一改", "加一节", "加一段", "润色", "删去", "调整内容")
# 弱编辑词：只有附了 .docx 时才视为编辑意图，避免误伤生成类需求
WEAK_EDIT = ("调整", "更新", "修正", "补充", "完善", "排版", "格式", "新增", "添加",
             "去掉", "不要", "重排", "精修", "美化", "改")


def _base_stem(name: str) -> str:
    """去掉版本后缀，得到文档基础名（用于版本对命名）。"""
    for suf in ("_完成版", "_原始版", "_修改版"):
        if name.endswith(suf):
            return name[: -len(suf)]
    return name


class WordAgentApp:
    def __init__(self, root: ctk.CTk):
        self.root = root
        self.config = Config.from_env()
        self.memory = Memory(self.config.memory_file)
        self.events: queue.Queue = queue.Queue()
        self.busy = False
        self.last_output: Path | None = None
        self.attachments: list[str] = []

        # 聊天状态
        self.chat_history: list[dict] = []
        self._mid = 0
        self._bubbles: dict[int, list] = {}
        self._current_mid: int | None = None
        self._plan_evt: threading.Event | None = None
        self._plan_ok = False
        self._edit_evt: threading.Event | None = None
        self._edit_state: dict | None = None
        self._edit_ok = False
        self._started_at = 0.0

        saved_theme = self.memory.get_setting("theme", "")
        if saved_theme in ("light", "dark"):
            ctk.set_appearance_mode(saved_theme)
        self._build_ui()
        self._apply_caret_color()
        self._greet()
        if self.config.api_key:
            self._set_status(f"就绪 · 已配置 {self.config.model}")
        else:
            self._set_status("就绪 · 请先在左侧填写 API Key")
        self.root.after(120, self._poll_events)

    # ================= UI 构建 =================
    def _build_ui(self) -> None:
        self.root.title(APP_TITLE)
        self.root.geometry("1280x860")
        self.root.minsize(1040, 720)

        # 顶部品牌栏
        header = ctk.CTkFrame(self.root, corner_radius=0, fg_color=("#1B3F9E", "#0D1526"), height=48)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(header, text="📄 WordAgent", font=ctk.CTkFont("Microsoft YaHei UI", 17, "bold"),
                     text_color=("white", "#E2E8F0")).pack(side="left", padx=18)
        ctk.CTkLabel(header, text="AI 文档生成 · 模板填写 · 智能编辑",
                     font=ctk.CTkFont("Microsoft YaHei UI", 11),
                     text_color=("#DCE6FF", "#8AA0C4")).pack(side="left", padx=(0, 6))
        self.header_status = ctk.CTkLabel(header, text="就绪",
                                          font=ctk.CTkFont("Microsoft YaHei UI", 12),
                                          text_color=("#DCE6FF", "#93A6C4"))
        self.header_status.pack(side="right", padx=18)

        body = ctk.CTkFrame(self.root, fg_color="transparent")
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)
        self._build_sidebar(body)
        self._build_chat(body)

    # ---------- 左侧边栏 ----------
    def _build_sidebar(self, body) -> None:
        side = ctk.CTkFrame(body, width=240, corner_radius=0, fg_color=SIDEBAR)
        side.grid(row=0, column=0, sticky="nsw")
        side.grid_propagate(False)

        def section(title: str) -> None:
            ctk.CTkLabel(side, text=title, anchor="w",
                         font=ctk.CTkFont("Microsoft YaHei UI", 11, "bold"),
                         text_color=SIDEBAR_TEXT).pack(fill="x", padx=16, pady=(12, 4))

        # —— 模型配置 ——
        section("🔑 模型配置")
        self.key_var = ctk.StringVar(value=self.config.api_key)
        ctk.CTkEntry(side, textvariable=self.key_var, placeholder_text="DeepSeek API Key",
                     show="*", height=34, corner_radius=10,
                     border_color=SIDEBAR_BTN_BORDER).pack(fill="x", padx=16)
        self.model_var = ctk.StringVar(value=self.config.model)
        row = ctk.CTkFrame(side, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(8, 0))
        ctk.CTkEntry(row, textvariable=self.model_var, placeholder_text="模型名",
                     height=34, corner_radius=10, border_color=SIDEBAR_BTN_BORDER).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(row, text="保存", width=56, height=34, corner_radius=10,
                      fg_color=BRAND, hover_color=BRAND_HOVER, text_color="white",
                      font=ctk.CTkFont("Microsoft YaHei UI", 12), command=self._save_config).pack(side="left", padx=(8, 0))

        # —— 文档偏好 ——
        section("🎨 文档偏好")
        self.style_var = ctk.StringVar(value="auto（由 AI 自动选择）")
        ctk.CTkOptionMenu(side, variable=self.style_var, values=STYLE_OPTIONS,
                          height=32, corner_radius=10, fg_color=ASSIST_BUBBLE,
                          text_color=SIDEBAR_BTN_TEXT, button_color=BTN_SOLID,
                          button_hover_color=BTN_SOLID_HOVER,
                          font=ctk.CTkFont("Microsoft YaHei UI", 11)).pack(fill="x", padx=16)
        self.std_var = ctk.StringVar(value="auto（AI 自动识别）")
        ctk.CTkOptionMenu(side, variable=self.std_var, values=STD_OPTIONS,
                          height=32, corner_radius=10, fg_color=ASSIST_BUBBLE,
                          text_color=SIDEBAR_BTN_TEXT, button_color=BTN_SOLID,
                          button_hover_color=BTN_SOLID_HOVER,
                          font=ctk.CTkFont("Microsoft YaHei UI", 11)).pack(fill="x", padx=16, pady=(8, 0))
        self.outdir_var = ctk.StringVar(value=str(self.config.output_dir))
        orow = ctk.CTkFrame(side, fg_color="transparent")
        orow.pack(fill="x", padx=16, pady=(8, 0))
        ctk.CTkEntry(orow, textvariable=self.outdir_var, placeholder_text="输出目录",
                     height=32, corner_radius=10, border_color=SIDEBAR_BTN_BORDER).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(orow, text="📂", width=44, height=32, corner_radius=10,
                      fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER,
                      font=ctk.CTkFont("Microsoft YaHei UI", 12), command=self._browse_dir).pack(side="left", padx=(8, 0))
        ctk.CTkLabel(side, text="编辑自动保留两版：原始版 + 完成版（可回退）",
                     font=ctk.CTkFont("Microsoft YaHei UI", 10), wraplength=200, justify="left",
                     text_color=MUTED_TEXT).pack(anchor="w", padx=16, pady=(10, 0))

        # —— 常用操作 ——
        section("🛠 常用操作")
        for text, cmd in (
            ("📁 打开输出目录", self._open_output_dir),
            ("🕘 最近文档", self._show_recent),
            ("🗂 版本管理", self._show_versions),
            ("🧠 上下文记忆", self._show_memory),
            ("🧹 清空对话", self._clear_chat),
            ("🗑 清空记忆", self._clear_memory),
            ("📤 导出对话记录", self._export_chat),
        ):
            ctk.CTkButton(side, text=text, height=30, corner_radius=9,
                          fg_color="transparent", border_width=1,
                          border_color=SIDEBAR_BTN_BORDER, text_color=SIDEBAR_BTN_TEXT,
                          hover_color=SIDEBAR_BTN_HOVER, anchor="w",
                          font=ctk.CTkFont("Microsoft YaHei UI", 11), command=cmd).pack(fill="x", padx=16, pady=2)

        # —— 底部：主题切换 ——
        theme_row = ctk.CTkFrame(side, fg_color="transparent")
        theme_row.pack(side="bottom", fill="x", padx=16, pady=14)
        mode = ctk.get_appearance_mode().lower()
        self.theme_btn = ctk.CTkButton(
            theme_row, text="☀️ 浅色" if mode == "dark" else "🌙 深色",
            height=34, corner_radius=10, fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER,
            font=ctk.CTkFont("Microsoft YaHei UI", 12), command=self._toggle_theme)
        self.theme_btn.pack(fill="x")
        ctk.CTkLabel(theme_row, text=f"WordAgent {VERSION_LABEL}",
                     font=ctk.CTkFont("Microsoft YaHei UI", 10),
                     text_color=MUTED_TEXT).pack(pady=(8, 0))

    # ---------- 右侧聊天区 ----------
    def _build_chat(self, body) -> None:
        col = ctk.CTkFrame(body, fg_color=APP_BG, corner_radius=0)
        col.grid(row=0, column=1, sticky="nsew")
        col.grid_columnconfigure(0, weight=1)
        col.grid_rowconfigure(2, weight=1)

        # 对话标题栏
        chat_head = ctk.CTkFrame(col, fg_color="transparent")
        chat_head.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 0))
        ctk.CTkLabel(chat_head, text="💬 对话", font=ctk.CTkFont("Microsoft YaHei UI", 14, "bold"),
                     text_color=SIDEBAR_BTN_TEXT).pack(side="left")
        ctk.CTkLabel(chat_head, text="AI 自动识别：生成新文档 / 按模板填写 / 编辑已有文档",
                     font=ctk.CTkFont("Microsoft YaHei UI", 10),
                     text_color=MUTED_TEXT).pack(side="right")

        # 快捷模板 chips
        chips = ctk.CTkFrame(col, fg_color="transparent")
        chips.grid(row=1, column=0, sticky="ew", padx=14, pady=(6, 4))
        ctk.CTkLabel(chips, text="快捷模板：", font=ctk.CTkFont("Microsoft YaHei UI", 11),
                     text_color=MUTED_TEXT).pack(side="left", padx=(0, 6))
        for name, prompt in CHIP_PROMPTS:
            ctk.CTkButton(chips, text=name, height=26, corner_radius=13,
                          fg_color="transparent", border_width=1,
                          border_color=SIDEBAR_BTN_BORDER, text_color=SIDEBAR_BTN_TEXT,
                          hover_color=SIDEBAR_BTN_HOVER,
                          font=ctk.CTkFont("Microsoft YaHei UI", 11),
                          command=lambda p=prompt: self._use_chip(p)).pack(side="left", padx=4)

        # 聊天滚动区
        self.chat_scroll = ctk.CTkScrollableFrame(col, corner_radius=0, fg_color="transparent")
        self.chat_scroll.grid(row=2, column=0, sticky="nsew", padx=14, pady=(2, 4))
        self.chat_scroll.grid_columnconfigure(0, weight=1)
        # 底部常驻提示：让聊天区底部始终有信息感，不出现“空灰块”
        self._scroll_hint = ctk.CTkLabel(self.chat_scroll, text="💡 按 Enter 发送 · Shift+Enter 换行 · 可附 .docx / 图片 / PDF 参考文件",
                                         font=ctk.CTkFont("Microsoft YaHei UI", 10), text_color=MUTED_TEXT)
        self._scroll_hint.pack(side="bottom", fill="x", pady=(6, 8))

        # 附件栏
        self.attach_frame = ctk.CTkFrame(col, fg_color="transparent")
        self.attach_frame.grid(row=3, column=0, sticky="ew", padx=14, pady=(2, 0))

        # 输入区（现代聊条样式：高对比圆角容器 + 内部输入框）
        inp = ctk.CTkFrame(col, fg_color=INPUT_BG, corner_radius=18, border_width=1,
                           border_color=("#C9D6EE", "#51628A"))
        inp.grid(row=4, column=0, sticky="ew", padx=14, pady=(8, 12))
        inp.grid_columnconfigure(0, weight=1)
        self.input_text = ctk.CTkTextbox(inp, height=76, corner_radius=12, border_width=0,
                                         fg_color="transparent", wrap="word",
                                         font=ctk.CTkFont("Microsoft YaHei UI", 13))
        self.input_text.grid(row=0, column=0, columnspan=3, sticky="ew")
        self.input_text.bind("<Return>", self._on_enter)
        self.input_text.bind("<Shift-Return>", self._on_shift_enter)
        self.input_text.bind("<FocusIn>", self._on_input_focus_in)
        self.input_text.bind("<FocusOut>", self._on_input_focus_out)
        self._set_placeholder()

        brow = ctk.CTkFrame(inp, fg_color="transparent")
        brow.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ctk.CTkButton(brow, text="📎 附件", width=92, height=36, corner_radius=12,
                      fg_color="transparent", border_width=1, border_color=SIDEBAR_BTN_BORDER,
                      text_color=SIDEBAR_BTN_TEXT, hover_color=SIDEBAR_BTN_HOVER,
                      font=ctk.CTkFont("Microsoft YaHei UI", 12), command=self._pick_files).pack(side="left")
        ctk.CTkButton(brow, text="🧹 清空对话", width=110, height=36, corner_radius=12,
                      fg_color="transparent", border_width=1, border_color=SIDEBAR_BTN_BORDER,
                      text_color=SIDEBAR_BTN_TEXT, hover_color=SIDEBAR_BTN_HOVER,
                      font=ctk.CTkFont("Microsoft YaHei UI", 12), command=self._clear_chat).pack(side="left", padx=8)
        self.busy_label = ctk.CTkLabel(brow, text="", text_color=MUTED_TEXT,
                                       font=ctk.CTkFont("Microsoft YaHei UI", 11))
        self.busy_label.pack(side="right", padx=10)
        self.send_btn = ctk.CTkButton(brow, text="发送 ⏎", width=120, height=36, corner_radius=12,
                                      fg_color=BRAND, hover_color=BRAND_HOVER, text_color="white",
                                      font=ctk.CTkFont("Microsoft YaHei UI", 13, "bold"),
                                      command=self._send_message)
        self.send_btn.pack(side="right")

    # ---------- 输入框 / 附件 ----------
    def _placeholder_color(self) -> str:
        return "#9AA7BD" if ctk.get_appearance_mode().lower() == "light" else "#5F6B85"

    def _set_placeholder(self) -> None:
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", PLACEHOLDER)
        self.input_text.configure(text_color=self._placeholder_color())

    def _is_placeholder(self) -> bool:
        return self.input_text.get("1.0", "end").strip() == PLACEHOLDER.strip()

    def _on_input_focus_in(self, _event=None) -> None:
        if self._is_placeholder():
            self.input_text.delete("1.0", "end")
            self.input_text.configure(text_color=("gray10" if ctk.get_appearance_mode().lower() == "light" else "white"))

    def _on_input_focus_out(self, _event=None) -> None:
        if not self.input_text.get("1.0", "end").strip():
            self._set_placeholder()

    def _on_enter(self, _event=None) -> str:
        self._send_message()
        return "break"

    def _on_shift_enter(self, _event=None) -> str:
        self.input_text.insert("insert", "\n")
        return "break"

    def _use_chip(self, prompt: str) -> None:
        self.input_text.delete("1.0", "end")
        self.input_text.insert("1.0", prompt)
        self.input_text.configure(text_color=("gray10" if ctk.get_appearance_mode().lower() == "light" else "white"))
        self.input_text.focus_set()

    def _pick_files(self) -> None:
        if self.busy:
            return
        chosen = filedialog.askopenfilenames(
            title="选择 .docx（待编辑/模板）或参考文件（图片 / PDF / 文本）",
            filetypes=[
                ("Word / 图片 / PDF / 文本", "*.docx *.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff *.pdf *.txt *.csv"),
                ("Word 文档", "*.docx"),
                ("图片", "*.png *.jpg *.jpeg *.bmp *.webp"),
                ("PDF", "*.pdf"),
                ("所有文件", "*.*"),
            ],
        )
        for f in chosen:
            if f not in self.attachments:
                self.attachments.append(f)
        self._refresh_attach_row()

    def _refresh_attach_row(self) -> None:
        for child in self.attach_frame.winfo_children():
            child.destroy()
        if not self.attachments:
            return
        ctk.CTkLabel(self.attach_frame, text="📎 已选择：",
                     font=ctk.CTkFont("Microsoft YaHei UI", 11), text_color=MUTED_TEXT).pack(side="left", padx=(0, 6))
        for f in self.attachments:
            chip = ctk.CTkFrame(self.attach_frame, corner_radius=8, fg_color=ASSIST_BUBBLE)
            chip.pack(side="left", padx=3, pady=2)
            ctk.CTkLabel(chip, text=f" {Path(f).name}", font=ctk.CTkFont("Microsoft YaHei UI", 10),
                         text_color=SIDEBAR_BTN_TEXT).pack(side="left", padx=(8, 2), pady=3)
            ctk.CTkButton(chip, text="✕", width=22, height=22, corner_radius=6,
                          fg_color="transparent", hover_color=ERROR_TEXT,
                          text_color=MUTED_TEXT, font=ctk.CTkFont("Microsoft YaHei UI", 10),
                          command=lambda fp=f: self._remove_file(fp)).pack(side="left", padx=(2, 6))
        if len(self.attachments) > 1:
            ctk.CTkButton(self.attach_frame, text="全部清除", width=70, height=24, corner_radius=8,
                          fg_color="transparent", border_width=1, border_color=SIDEBAR_BTN_BORDER,
                          text_color=SIDEBAR_BTN_TEXT, hover_color=SIDEBAR_BTN_HOVER,
                          font=ctk.CTkFont("Microsoft YaHei UI", 10),
                          command=self._clear_files).pack(side="left", padx=6)

    def _remove_file(self, path: str) -> None:
        if path in self.attachments:
            self.attachments.remove(path)
        self._refresh_attach_row()

    def _clear_files(self) -> None:
        self.attachments = []
        self._refresh_attach_row()

    # ================= 意图识别与发送 =================
    def _detect_intent(self, text: str, files: list[str]):
        """自动判断需求类型：fill（按模板填写）/ edit（编辑已有文档）/ generate（生成新文档）。"""
        docx = next((f for f in files if str(f).lower().endswith(".docx")), None)
        others = [f for f in files if f != docx]
        if docx:
            if any(k in text for k in FILL_KEYWORDS):
                return "fill", docx, others, "📋 检测到模板文件，将按模板结构填写内容。"
            if any(k in text for k in STRONG_EDIT) or any(k in text for k in WEAK_EDIT):
                return "edit", docx, others, "📄 检测到 .docx 与修改要求，将先读取文档内容再精准修改。"
            return "generate", None, [docx] + others, "📄 已把 .docx 作为模板参考（如需修改/填写，请加上“修改/填写”等词）。"
        if any(k in text for k in STRONG_EDIT):
            if self.last_output and Path(self.last_output).exists():
                return "edit", str(self.last_output), files, f"📄 将基于当前文档继续修改：{Path(self.last_output).name}"
            return "need_doc", None, files, "看起来你想编辑文档：请先点「📎 附件」添加要编辑的 .docx，再发送这条要求。"
        return "generate", None, files, ""

    def _send_message(self) -> None:
        if self.busy:
            self._set_status("⏳ 正在处理上一条消息，请稍候…")
            return
        text = self.input_text.get("1.0", "end").strip()
        if self._is_placeholder():
            text = ""
        if not text:
            self._set_status("请先输入文档要求。")
            return
        key = self.key_var.get().strip() or self.config.api_key
        if not key:
            messagebox.showwarning(
                "缺少 API Key",
                "请先在左侧填写 DeepSeek API Key 并点「保存」。\n\n也可以在工作目录创建 .env 文件：\nDEEPSEEK_API_KEY=sk-xxx",
            )
            return
        self.config.api_key = key
        self.config.model = self.model_var.get().strip() or self.config.model

        self.input_text.delete("1.0", "end")
        files = list(self.attachments)
        self.attachments = []
        self._refresh_attach_row()
        self._add_user_bubble(text, files)
        self.chat_history.append({"role": "user", "text": text, "files": files})

        kind, target, refs, note = self._detect_intent(text, files)
        if kind == "need_doc":
            self._add_assist_bubble(note)
            return
        if note:
            self._add_system_bubble(note)
        mid = self._add_assist_bubble("⏳ 正在解析你的要求…")
        self._current_mid = mid
        self._started_at = time.time()
        self._set_busy(True)
        prev_output = str(self.last_output) if self.last_output else ""
        threading.Thread(target=self._worker_chat,
                         args=(kind, text, mid, target, refs, prev_output), daemon=True).start()
        self.input_text.focus_set()

    def _set_busy(self, busy: bool, status: str = "") -> None:
        self.busy = busy
        self.send_btn.configure(state="disabled" if busy else "normal")
        self.busy_label.configure(text="处理中…" if busy else "")
        if status:
            self._set_status(status)

    def _set_status(self, text: str) -> None:
        self.header_status.configure(text=text)

    # ================= Worker（后台线程） =================
    def _worker_chat(self, kind: str, text: str, mid: int, target: str | None, refs: list[str],
                     prev_output: str = "") -> None:
        def log(msg: str) -> None:
            self.events.put(("bubble_text", mid, msg))

        try:
            llm = LLMClient(self.config)
            if kind == "generate":
                style_override = None
                sr = self.style_var.get()
                if not sr.startswith("auto"):
                    style_override = sr.split("（")[0]
                standard_override = None
                sr2 = self.std_var.get()
                if not sr2.startswith("auto"):
                    standard_override = sr2.split("（")[0]
                output_dir = Path(self.outdir_var.get().strip()) or self.config.output_dir
                path = run_pipeline(text, self.config, self.memory, log=log,
                                    style_override=style_override, standard_override=standard_override,
                                    output_dir_override=output_dir, reference_files=refs,
                                    llm=llm, confirm_plan=self._request_plan_confirm)
                if path is None:
                    self.events.put(("cancelled", mid))
                    return
                usage = llm.usage_text() if hasattr(llm, "usage_text") else ""
                self.events.put(("done_generate", str(path), mid, usage))
            elif kind == "edit":
                output_dir = Path(self.outdir_var.get().strip()) or self.config.output_dir
                state = prepare_edit(Path(target), text, self.config, self.memory,
                                     llm=llm, log=log, reference_files=refs)
                self.events.put(("edit_plan", state, mid))
                if not self._wait_edit_confirm():
                    self.events.put(("cancelled", mid))
                    return
                path = finalize_edit(self._edit_state, self.config, self.memory,
                                     output_dir=output_dir, save_as_new=True,
                                     log=log, resolve_ambiguous=self._resolve_ambiguous,
                                     on_warnings=self._on_warnings, version_keep=True)
                # 双版本归档：若被编辑的是之前生成的“最早版本”，移除它（只留 原始版+完成版）
                try:
                    src_p = Path(target)
                    if prev_output and Path(prev_output) == src_p and src_p != Path(path) \
                            and "versions" not in src_p.parts and src_p.suffix.lower() == ".docx" \
                            and src_p.parent == output_dir:
                        src_p.unlink(missing_ok=True)
                        log(f"   已归档并移除最早版本：{src_p.name}")
                except OSError:
                    pass
                usage = llm.usage_text() if hasattr(llm, "usage_text") else ""
                base = _base_stem(Path(target).stem)
                orig = output_dir / "versions" / f"{base}_原始版.docx"
                self.events.put(("done_edit", str(path), mid, usage, str(orig)))
            elif kind == "fill":
                path = fill_template(Path(target), text, self.config, self.memory, llm=llm, log=log)
                usage = llm.usage_text() if hasattr(llm, "usage_text") else ""
                self.events.put(("done_generate", str(path), mid, usage))
        except (LLMError, Exception) as exc:  # noqa: BLE001
            self.events.put(("error", str(exc), mid))

    def _request_plan_confirm(self, plan: dict) -> bool:
        """后台线程调用：把大纲渲染进聊天区，阻塞等待用户点「确认 / 取消」。"""
        evt = threading.Event()
        self._plan_evt = evt
        self._plan_ok = False
        self.events.put(("plan_show", plan))
        evt.wait()
        return self._plan_ok

    def _wait_edit_confirm(self) -> bool:
        evt = threading.Event()
        self._edit_evt = evt
        self._edit_ok = False
        evt.wait()
        return self._edit_ok

    # ================= 聊天气泡 =================
    def _new_mid(self) -> int:
        self._mid += 1
        return self._mid

    def _scroll_bottom(self) -> None:
        try:
            self.chat_scroll.update_idletasks()
            self.chat_scroll._parent_canvas.yview_moveto(1.0)
        except Exception:
            pass

    def _add_user_bubble(self, text: str, files: list[str]) -> int:
        mid = self._new_mid()
        outer = ctk.CTkFrame(self.chat_scroll, fg_color="transparent")
        outer.pack(fill="x", pady=(10, 2), padx=10, before=getattr(self, "_scroll_hint", None))
        inner = ctk.CTkFrame(outer, fg_color=USER_BUBBLE, corner_radius=14)
        ctk.CTkLabel(inner, text=text, justify="left", anchor="w", wraplength=640,
                     font=ctk.CTkFont("Microsoft YaHei UI", 13),
                     text_color=("white", "white")).pack(anchor="e", padx=12, pady=(8, 4))
        for f in files[:5]:
            ctk.CTkLabel(inner, text=f"📎 {Path(f).name}", anchor="w", justify="left",
                         font=ctk.CTkFont("Microsoft YaHei UI", 10),
                         text_color=("#EAF1FF", "#EAF1FF")).pack(anchor="e", padx=12, pady=1)
        if len(files) > 5:
            ctk.CTkLabel(inner, text=f"… 共 {len(files)} 个文件", anchor="w",
                         font=ctk.CTkFont("Microsoft YaHei UI", 10),
                         text_color=("#D5E1F5", "#D5E1F5")).pack(anchor="e", padx=12, pady=1)
        ctk.CTkLabel(outer, text=time.strftime("%H:%M"), anchor="e",
                     font=ctk.CTkFont("Microsoft YaHei UI", 9), text_color=MUTED_TEXT).pack(anchor="e", padx=6, pady=(2, 0))
        inner.pack(anchor="e")
        self._bubbles[mid] = [outer, None, inner]
        self._scroll_bottom()
        return mid

    def _add_assist_bubble(self, text: str, mid: int | None = None) -> int:
        if mid is None:
            mid = self._new_mid()
        outer = ctk.CTkFrame(self.chat_scroll, fg_color="transparent")
        outer.pack(fill="x", pady=(10, 2), padx=10, before=getattr(self, "_scroll_hint", None))
        ctk.CTkLabel(outer, text=f"WordAgent · {time.strftime('%H:%M')}", anchor="w",
                     font=ctk.CTkFont("Microsoft YaHei UI", 9), text_color=MUTED_TEXT).pack(anchor="w", padx=6, pady=(0, 2))
        inner = ctk.CTkFrame(outer, fg_color=ASSIST_BUBBLE, corner_radius=14)
        label = ctk.CTkLabel(inner, text=text, justify="left", anchor="w", wraplength=680,
                             font=ctk.CTkFont("Microsoft YaHei UI", 13),
                             text_color=SIDEBAR_BTN_TEXT)
        label.pack(anchor="w", padx=12, pady=8)
        inner.pack(anchor="w")
        self._bubbles[mid] = [outer, label, inner]
        self._scroll_bottom()
        return mid

    def _add_system_bubble(self, text: str) -> None:
        ctk.CTkLabel(self.chat_scroll, text=text, justify="center", anchor="center",
                     font=ctk.CTkFont("Microsoft YaHei UI", 10), text_color=MUTED_TEXT,
                     wraplength=720).pack(fill="x", pady=4, padx=10, before=getattr(self, "_scroll_hint", None))

    def _set_bubble_text(self, mid: int, text: str) -> None:
        entry = self._bubbles.get(mid)
        if entry and entry[1] is not None:
            entry[1].configure(text=text)
            self._scroll_bottom()

    def _bubble_add_buttons(self, mid: int, buttons: list) -> None:
        """在气泡底部追加按钮行。buttons: [(text, fg, hover, command)]"""
        entry = self._bubbles.get(mid)
        if not entry:
            return
        inner = entry[2]
        row = ctk.CTkFrame(inner, fg_color="transparent")
        for text, fg, hover, cmd in buttons:
            ctk.CTkButton(row, text=text, height=32, corner_radius=10,
                          fg_color=fg, hover_color=hover, text_color="white",
                          font=ctk.CTkFont("Microsoft YaHei UI", 12),
                          command=cmd).pack(side="left", padx=4, pady=4)
        row.pack(anchor="w", padx=10, pady=(2, 8))
        self._scroll_bottom()

    # ================= 计划 / 结果渲染 =================
    def _render_plan_bubble(self, plan: dict) -> None:
        mid = self._current_mid or self._new_mid()
        style_label = {"business": "商务", "report": "报告", "academic": "学术",
                       "creative": "创意文案", "default": "通用"}.get(plan.get("style"), "通用")
        lines = [f"📋 大纲预览｜标题：{plan.get('title', '')}",
                 f"风格：{style_label}　|　目录：{'是' if plan.get('toc') else '否'}"]
        sections = plan.get("sections") or []
        if sections:
            lines.append("")
            for sec in sections:
                indent = "　　" * (max(0, int(sec.get("level", 1)) - 1))
                lines.append(f"{indent}· {sec.get('heading', '')}")
        lines.append("")
        lines.append("确认大纲后生成正文；不满意可直接取消，回到输入框补充要求后重试。")
        self._set_bubble_text(mid, "\n".join(lines))

        def ok() -> None:
            self._plan_ok = True
            if self._plan_evt:
                self._plan_evt.set()

        def cancel() -> None:
            self._plan_ok = False
            if self._plan_evt:
                self._plan_evt.set()

        self._bubble_add_buttons(mid, [
            ("✓ 按此大纲生成", GREEN, GREEN_HOVER, ok),
            ("✕ 取消修改指令", BTN_SOLID, BTN_SOLID_HOVER, cancel),
        ])

    def _render_edit_plan(self, state: dict, mid: int) -> None:
        plan = state.get("plan") or {}
        ops = plan.get("operations") or []
        lines = [f"📋 已读取《{Path(str(state['src'])).name}》，修改计划："]
        if plan.get("summary"):
            lines.append(f"　说明：{plan['summary']}")
        lines.append(f"　共 {len(ops)} 项操作：")
        for op in ops[:12]:
            target = str(op.get("target", ""))[:42]
            extra = op.get("new_text") or op.get("style") or ""
            tail = f" → {str(extra)[:30]}" if extra else ""
            lines.append(f"　· [{op.get('op', '?')}]「{target}」{tail}")
        if len(ops) > 12:
            lines.append(f"　… 其余 {len(ops) - 12} 项操作保存时一并应用")
        lines.append("")
        lines.append("确认后开始修改（原文件自动备份）；取消则不做任何改动。")
        self._set_bubble_text(mid, "\n".join(lines))

        def ok() -> None:
            self._edit_state = state
            self._edit_ok = True
            if self._edit_evt:
                self._edit_evt.set()

        def cancel() -> None:
            self._edit_ok = False
            if self._edit_evt:
                self._edit_evt.set()

        self._bubble_add_buttons(mid, [
            ("✓ 确认修改", GREEN, GREEN_HOVER, ok),
            ("✕ 取消", BTN_SOLID, BTN_SOLID_HOVER, cancel),
        ])

    def _finish_ok(self, path: str, mid: int, action: str, usage: str = "", orig: str = "") -> None:
        self.last_output = Path(path)
        used = f"\n\n⚙ {usage}" if usage else ""
        elapsed = time.time() - self._started_at if self._started_at else 0
        done_p = Path(path)
        if action == "修改" and orig and Path(orig).exists():
            self._set_bubble_text(
                mid,
                f"✅ 修改完成（用时 {elapsed:.0f} 秒）\n📄 完成版：{path}\n"
                f"🗂 原始版：{orig}{used}\n\n"
                "共保留两个版本，可随时回退到原始版。",
            )
            self._bubble_add_buttons(mid, [
                ("📄 打开完成版", BRAND, BRAND_HOVER, lambda: self._open_path(done_p)),
                ("🗂 打开原版", BTN_SOLID, BTN_SOLID_HOVER, lambda: self._open_path(Path(orig))),
                ("↩ 回退到原版", "#B45309", "#92400E", lambda: self._rollback(done_p, Path(orig))),
            ])
        else:
            self._set_bubble_text(
                mid,
                f"✅ {action}完成（用时 {elapsed:.0f} 秒），文档已保存：\n{path}{used}\n\n"
                "💡 继续对我说「把标题改成…」「再加一节结论…」即可接着修改。",
            )
            self._bubble_add_buttons(mid, [
                ("📄 打开文档", BRAND, BRAND_HOVER, lambda: self._open_path(done_p)),
                ("📂 所在文件夹", BTN_SOLID, BTN_SOLID_HOVER, lambda: self._open_folder_of(done_p)),
            ])
        self.chat_history.append({"role": "assistant", "text": f"[{action}完成] {path}"})
        self._set_busy(False)
        self._set_status(f"✔ {action}完成")

    def _rollback(self, done_p: Path, orig_p: Path) -> None:
        """回退到原始版：用原始版覆盖完成版并打开。"""
        try:
            shutil.copy2(orig_p, done_p)
            self.last_output = done_p
            self._add_system_bubble(f"↩ 已回退到原始版：{done_p.name}")
            self._set_status("✔ 已回退到原始版")
            self._open_path(done_p)
        except OSError as exc:
            messagebox.showerror("回退失败", f"无法恢复原始版本：{exc}")

    def _finish_error(self, message: str, mid: int) -> None:
        self._set_bubble_text(mid, f"❌ 操作失败：\n{message}")
        entry = self._bubbles.get(mid)
        if entry and entry[1] is not None:
            entry[1].configure(text_color=ERROR_TEXT)
        self._set_busy(False)
        self._set_status("✖ 操作失败")

    def _handle_cancelled(self, mid: int) -> None:
        self._set_bubble_text(mid, "已取消：未生成 / 未修改文档。")
        self._set_busy(False)
        self._set_status("已取消")

    # ================= 事件轮询 =================
    def _poll_events(self) -> None:
        try:
            while True:
                item = self.events.get_nowait()
                kind = item[0]
                if kind == "bubble_text":
                    self._set_bubble_text(item[1], item[2])
                elif kind == "plan_show":
                    self._render_plan_bubble(item[1])
                elif kind == "edit_plan":
                    self._render_edit_plan(item[1], item[2])
                elif kind == "cancelled":
                    self._handle_cancelled(item[1])
                elif kind == "done_generate":
                    self._finish_ok(item[1], item[2], "生成", item[3] if len(item) > 3 else "")
                elif kind == "done_edit":
                    self._finish_ok(item[1], item[2], "修改", item[3] if len(item) > 3 else "",
                                    item[4] if len(item) > 4 else "")
                elif kind == "error":
                    self._finish_error(item[1], item[2])
        except queue.Empty:
            pass
        self.root.after(120, self._poll_events)

    # ================= 编辑期阻塞弹窗 =================
    def _resolve_ambiguous(self, op: dict, candidates: list[str]) -> str | None:
        """目标有歧义时弹窗让用户选择（主线程展示，阻塞等待）。"""
        result: list[str | None] = [None]
        evt = threading.Event()

        def _show() -> None:
            win = ctk.CTkToplevel(self.root)
            win.title("目标匹配到多个位置")
            win.geometry("640x420")
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

            def _confirm() -> None:
                ans = entry.get().strip()
                if ans.isdigit() and 1 <= int(ans) <= len(candidates[:6]):
                    result[0] = candidates[:6][int(ans) - 1]
                evt.set()
                win.destroy()

            def _cancel() -> None:
                evt.set()
                win.destroy()

            row = ctk.CTkFrame(win, fg_color="transparent")
            row.pack(fill="x", padx=18, pady=(10, 16))
            ctk.CTkButton(row, text="✓ 确认", width=110, height=36, corner_radius=12,
                          fg_color=GREEN, hover_color=GREEN_HOVER,
                          font=ctk.CTkFont("Microsoft YaHei UI", 13), command=_confirm).pack(side="left")
            ctk.CTkButton(row, text="✕ 取消", width=100, height=36, corner_radius=12,
                          fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER,
                          font=ctk.CTkFont("Microsoft YaHei UI", 13), command=_cancel).pack(side="left", padx=10)

        self.root.after(0, _show)
        evt.wait()
        return result[0]

    def _on_warnings(self, issues: list[str]) -> bool:
        """结构校验有风险时弹窗询问是否继续（阻塞等待用户选择）。"""
        result: list[bool] = [False]
        evt = threading.Event()

        def _show() -> None:
            win = ctk.CTkToplevel(self.root)
            win.title("结构校验风险")
            win.geometry("620x380")
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
                evt.set()
                win.destroy()

            ctk.CTkButton(row, text="✓ 继续保存", width=120, height=36, corner_radius=12,
                          fg_color=GREEN, hover_color=GREEN_HOVER,
                          font=ctk.CTkFont("Microsoft YaHei UI", 13), command=_yes).pack(side="left")
            ctk.CTkButton(row, text="✕ 取消", width=100, height=36, corner_radius=12,
                          fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER,
                          font=ctk.CTkFont("Microsoft YaHei UI", 13), command=_no).pack(side="left", padx=10)

        self.root.after(0, _show)
        evt.wait()
        return result[0]

    # ================= 主题 / 配置 =================
    def _toggle_theme(self) -> None:
        mode = ctk.get_appearance_mode().lower()
        new = "light" if mode == "dark" else "dark"
        ctk.set_appearance_mode(new)
        self.memory.set_setting("theme", new)
        self.theme_btn.configure(text="☀️ 浅色" if new == "dark" else "🌙 深色")
        self._apply_caret_color()
        self._refresh_placeholder_color()

    def _refresh_placeholder_color(self) -> None:
        if self._is_placeholder():
            self.input_text.configure(text_color=self._placeholder_color())

    def _apply_caret_color(self) -> None:
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

    def _save_config(self) -> None:
        key = self.key_var.get().strip()
        model = self.model_var.get().strip()
        if not key:
            messagebox.showwarning("提示", "API Key 不能为空。")
            return
        self.config.api_key = key
        self.config.model = model or self.config.model
        self._persist_env({"DEEPSEEK_API_KEY": key, "DEEPSEEK_MODEL": self.config.model})
        self._set_status(f"✔ 已保存配置（{self.config.model}）")
        self._add_system_bubble(f"🔑 配置已保存：{self.config.model}")

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

    # ================= 常用操作 =================
    def _browse_dir(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.outdir_var.get() or str(Path.cwd()))
        if chosen:
            self.outdir_var.set(chosen)

    def _open_output_dir(self) -> None:
        try:
            os.startfile(self.outdir_var.get().strip() or str(self.config.output_dir))
        except Exception:
            messagebox.showinfo("提示", "输出目录不存在，请先设置。")

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

    def _show_recent(self) -> None:
        items = [
            e for e in self.memory.entries
            if e.get("role") == "assistant" and e.get("file") and Path(str(e["file"])).exists()
        ][-12:][::-1]
        win = ctk.CTkToplevel(self.root)
        win.title("最近文档")
        win.geometry("760x520")
        win.transient(self.root)
        ctk.CTkLabel(win, text="🕘 最近文档（点击打开，右侧按钮打开所在文件夹）",
                     font=ctk.CTkFont("Microsoft YaHei UI", 15, "bold")).pack(anchor="w", padx=18, pady=(16, 8))
        if not items:
            ctk.CTkLabel(win, text="（还没有生成或编辑过文档）",
                         font=ctk.CTkFont("Microsoft YaHei UI", 12),
                         text_color=MUTED_TEXT).pack(anchor="w", padx=18, pady=20)
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

    def _show_versions(self) -> None:
        """版本管理：列出输出目录里全部 原始版/完成版 版本对，支持打开与回退。"""
        out_dir = Path(self.outdir_var.get().strip()) or self.config.output_dir
        ver_dir = out_dir / "versions"
        pairs: dict[str, list] = {}
        if ver_dir.is_dir():
            for f in sorted(ver_dir.glob("*.docx"), key=lambda x: x.name):
                name = _base_stem(f.stem)
                if f.name.endswith("_原始版.docx"):
                    pairs.setdefault(name, [None, None])[0] = f
                elif f.name.endswith("_完成版.docx"):
                    pairs.setdefault(name, [None, None])[1] = f
        win = ctk.CTkToplevel(self.root)
        win.title("版本管理")
        win.geometry("720x480")
        win.transient(self.root)
        ctk.CTkLabel(win, text="🗂 版本管理（每份文档只保留 原始版 + 完成版）",
                     font=ctk.CTkFont("Microsoft YaHei UI", 15, "bold")).pack(anchor="w", padx=18, pady=(16, 8))
        if not pairs:
            ctk.CTkLabel(win, text="（还没有编辑过的文档，生成/修改后这里会出现版本对）",
                         font=ctk.CTkFont("Microsoft YaHei UI", 12),
                         text_color=MUTED_TEXT).pack(anchor="w", padx=18, pady=20)
            return
        box = ctk.CTkScrollableFrame(win, corner_radius=12)
        box.pack(fill="both", expand=True, padx=18, pady=(0, 16))
        for name in sorted(pairs):
            orig, done = pairs[name]
            row = ctk.CTkFrame(box, corner_radius=10)
            row.pack(fill="x", pady=4)
            ctk.CTkLabel(row, text=f"📄 {name}", anchor="w", width=200,
                         font=ctk.CTkFont("Microsoft YaHei UI", 12, "bold"),
                         text_color=SIDEBAR_BTN_TEXT).pack(side="left", padx=10, pady=6)
            if done:
                ctk.CTkButton(row, text="打开完成版", width=96, height=30, corner_radius=10,
                              fg_color=BRAND, hover_color=BRAND_HOVER, text_color="white",
                              font=ctk.CTkFont("Microsoft YaHei UI", 11),
                              command=lambda f=done: self._open_path(f)).pack(side="left", padx=4)
            if orig:
                ctk.CTkButton(row, text="打开原版", width=88, height=30, corner_radius=10,
                              fg_color=BTN_SOLID, hover_color=BTN_SOLID_HOVER,
                              font=ctk.CTkFont("Microsoft YaHei UI", 11),
                              command=lambda f=orig: self._open_path(f)).pack(side="left", padx=4)
            if orig and done:
                ctk.CTkButton(row, text="↩ 回退", width=72, height=30, corner_radius=10,
                              fg_color="#B45309", hover_color="#92400E", text_color="white",
                              font=ctk.CTkFont("Microsoft YaHei UI", 11),
                              command=lambda o=orig, d=done: self._rollback(d, o)).pack(side="left", padx=4)

    def _show_memory(self) -> None:
        win = ctk.CTkToplevel(self.root)
        win.title("上下文记忆")
        win.geometry("720x460")
        win.transient(self.root)
        ctk.CTkLabel(win, text="🧠 上下文记忆（历史指令与结果，AI 会结合这些上下文）",
                     font=ctk.CTkFont("Microsoft YaHei UI", 15, "bold")).pack(anchor="w", padx=18, pady=(16, 6))
        box = ctk.CTkTextbox(win, corner_radius=12, font=ctk.CTkFont("Microsoft YaHei UI", 12))
        box.pack(fill="both", expand=True, padx=18, pady=12)
        box.insert("1.0", self.memory.history_text())
        box.configure(state="disabled")

    def _clear_memory(self) -> None:
        if messagebox.askyesno("确认", "确定清空全部上下文记忆吗？"):
            self.memory.clear()
            self._add_system_bubble("🧠 已清空上下文记忆。")

    def _clear_chat(self) -> None:
        if messagebox.askyesno("确认", "清空当前聊天记录？（不影响上下文记忆）"):
            for child in self.chat_scroll.winfo_children():
                child.destroy()
            self._bubbles.clear()
            self.chat_history.clear()
            self._scroll_hint = ctk.CTkLabel(self.chat_scroll,
                                             text="💡 按 Enter 发送 · Shift+Enter 换行 · 可附 .docx / 图片 / PDF 参考文件",
                                             font=ctk.CTkFont("Microsoft YaHei UI", 10), text_color=MUTED_TEXT)
            self._scroll_hint.pack(side="bottom", fill="x", pady=(6, 8))
            self._greet()
            self._set_status("已清空对话")

    def _export_chat(self) -> None:
        from datetime import datetime as _dt
        default = f"WordAgent对话_{_dt.now().strftime('%Y%m%d_%H%M%S')}.txt"
        target = filedialog.asksaveasfilename(title="导出对话记录", defaultextension=".txt",
                                              initialfile=default, filetypes=[("文本文件", "*.txt")])
        if not target:
            return
        lines = [f"WordAgent 对话记录　{_dt.now().strftime('%Y-%m-%d %H:%M:%S')}", ""]
        for m in self.chat_history:
            role = "我" if m["role"] == "user" else "WordAgent"
            lines.append(f"【{role}】{m['text']}")
            for f in m.get("files", []):
                lines.append(f"　📎 {f}")
            lines.append("")
        try:
            Path(target).write_text("\n".join(lines), encoding="utf-8-sig")
            self._add_system_bubble(f"📤 对话已导出：{target}")
        except OSError as exc:
            messagebox.showerror("导出失败", f"无法写入文件：{exc}")

    # ================= 欢迎语 =================
    def _greet(self) -> None:
        self._add_assist_bubble(
            "👋 你好，我是 WordAgent！\n\n"
            "直接告诉我要什么文档即可，例如：\n"
            "· 写一份《数据库课程实验报告》\n"
            "· 写一份本周工作周报\n"
            "· 生成《2026 市场推广方案》\n\n"
            "要编辑或套模板时，先点「📎 附件」添加 .docx，然后说：\n"
            "· 把标题改成《XXX》，第二节改得更正式\n"
            "· 按模板填写这份实验报告\n\n"
            "我会自动判断需求类型，支持连续对话与上下文记忆。"
        )


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