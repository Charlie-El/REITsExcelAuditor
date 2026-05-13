from __future__ import annotations

import argparse
import ctypes
import multiprocessing
from pathlib import Path
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable

try:
    from reit_excel_auditor.annual_update import (
        DEFAULT_AI_BATCH_CHAR_LIMIT,
        DEFAULT_AI_MODEL,
        DEFAULT_AI_REQUEST_TIMEOUT_SECONDS,
        DEFAULT_AI_TOTAL_TIMEOUT_SECONDS,
        DEFAULT_DASHSCOPE_BASE_URL,
        DEFAULT_VISION_OCR_MODEL,
        AnnualUpdateError,
        AnnualUpdateOptions,
        AnnualUpdateResult,
        SUMMARY_OUTPUT_NAME,
        run_annual_update,
    )
    from reit_excel_auditor.transformer import (
        AUTO_TYPE,
        TABLE_TYPES,
        BatchResult,
        ConversionError,
        convert_input_path,
    )
except ImportError:  # Direct script execution fallback.
    from annual_update import (
        DEFAULT_AI_BATCH_CHAR_LIMIT,
        DEFAULT_AI_MODEL,
        DEFAULT_AI_REQUEST_TIMEOUT_SECONDS,
        DEFAULT_AI_TOTAL_TIMEOUT_SECONDS,
        DEFAULT_DASHSCOPE_BASE_URL,
        DEFAULT_VISION_OCR_MODEL,
        AnnualUpdateError,
        AnnualUpdateOptions,
        AnnualUpdateResult,
        SUMMARY_OUTPUT_NAME,
        run_annual_update,
    )
    from transformer import AUTO_TYPE, TABLE_TYPES, BatchResult, ConversionError, convert_input_path


APP_TITLE = "REITs Excel 自动审核转换工具"
WINDOW_SIZE = "1360x1040"
APP_ICON_RELATIVE = Path("reit_excel_auditor") / "assets" / "app_icon.ico"

TYPE_LABELS = [
    TABLE_TYPES[AUTO_TYPE],
    TABLE_TYPES["valuation"],
    TABLE_TYPES["traffic"],
    TABLE_TYPES["finance"],
    TABLE_TYPES["property"],
    TABLE_TYPES["energy"],
]
LABEL_TO_TYPE = {label: key for key, label in TABLE_TYPES.items()}

AI_PROVIDER_PROFILES = {
    "阿里云百炼 / 通义千问（推荐）": {
        "base_url": DEFAULT_DASHSCOPE_BASE_URL,
        "model": DEFAULT_AI_MODEL,
        "key_env": "",
        "note": "适合中文 OCR 文本整理；API Key 或环境变量名请由用户自行填写。",
    },
    "OpenAI": {
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-4o-mini",
        "key_env": "",
        "note": "需要 OpenAI API Key；如模型名称变化，可在模型框手动修改。",
    },
    "本地 Ollama（OpenAI兼容）": {
        "base_url": "http://localhost:11434/v1",
        "model": "qwen2.5:7b",
        "key_env": "",
        "note": "适合不希望发送 OCR 文本到云端的用户；需本机先启动 Ollama。",
    },
    "自定义 OpenAI 兼容接口": {
        "base_url": "",
        "model": "",
        "key_env": "",
        "note": "适合硅基流动、火山方舟、智谱、Moonshot 等兼容 Chat Completions 的服务。",
    },
}

OCR_API_PROVIDER_PROFILES = {
    "不使用云端 OCR（本地 OCR，无需 API）": {
        "base_url": "",
        "model": "",
        "key_env": "",
        "note": "本地 RapidOCR、PaddleOCR、Tesseract 和 PDF 文本抽取都不需要 API Key。",
    },
    "阿里云百炼 / 通义视觉 OCR": {
        "base_url": DEFAULT_DASHSCOPE_BASE_URL,
        "model": DEFAULT_VISION_OCR_MODEL,
        "key_env": "",
        "note": "仅当 OCR 引擎选择 vision_api 时使用；会上传需要 OCR 的截图或 PDF 渲染页。",
    },
    "OpenAI 视觉模型": {
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-4o-mini",
        "key_env": "",
        "note": "仅当 OCR 引擎选择 vision_api 时使用；可按实际账号改模型名称。",
    },
    "自定义 OpenAI 兼容视觉 OCR": {
        "base_url": "",
        "model": "",
        "key_env": "",
        "note": "适合其他兼容 Chat Completions 且支持 image_url 输入的视觉 OCR 服务。",
    },
}


class AuditorApp(tk.Tk):
    def __init__(self) -> None:
        enable_high_dpi()
        super().__init__()
        self.title(APP_TITLE)
        set_window_icon(self)
        self.geometry(WINDOW_SIZE)
        self.minsize(1040, 760)
        self.configure(bg="#F5F7FB")

        self.source_path = tk.StringVar()
        self.metadata_path = tk.StringVar()
        self.custom_template_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.table_type = tk.StringVar(value=TABLE_TYPES[AUTO_TYPE])
        self.generate_property_processed = tk.BooleanVar(value=False)
        self.template_mode_text = tk.StringVar(value="")
        self.folder_warning_text = tk.StringVar(value="")
        self.status_text = tk.StringVar(value="等待选择未审核 Excel 文件或文件夹")
        self.annual_workspace_path = tk.StringVar()
        self.annual_standard_input_path = tk.StringVar()
        self.annual_ocr_source_path = tk.StringVar()
        self.annual_report_source_path = tk.StringVar()
        self.annual_output_dir = tk.StringVar()
        self.annual_ocr_engine = tk.StringVar(value="auto")
        self.annual_ocr_api_provider = tk.StringVar(value="不使用云端 OCR（本地 OCR，无需 API）")
        self.annual_ocr_api_key = tk.StringVar()
        self.annual_ocr_api_key_env = tk.StringVar()
        self.annual_ocr_base_url = tk.StringVar()
        self.annual_ocr_model = tk.StringVar()
        self.annual_use_ai = tk.BooleanVar(value=False)
        self.annual_ai_provider = tk.StringVar(value="阿里云百炼 / 通义千问（推荐）")
        self.annual_api_key = tk.StringVar()
        self.annual_api_key_env = tk.StringVar()
        self.annual_base_url = tk.StringVar(value=DEFAULT_DASHSCOPE_BASE_URL)
        self.annual_model = tk.StringVar(value=DEFAULT_AI_MODEL)
        self.annual_output_start_year = tk.StringVar(value="2027")
        self.annual_max_pages = tk.StringVar(value="5")
        self.annual_strict_mode = tk.BooleanVar(value=False)
        self.annual_status_text = tk.StringVar(value="等待选择年报工作文件夹")
        self.annual_ocr_api_state_text = tk.StringVar(value="当前未启用：如需云端 OCR，请先在“基础设置”中把 OCR 引擎改为 vision_api。")

        self._configure_style()
        self._build_layout()
        self.custom_template_path.trace_add("write", lambda *_: self.update_template_mode())
        self.annual_ai_provider.trace_add("write", lambda *_: self.apply_annual_ai_provider())
        self.annual_ocr_engine.trace_add("write", lambda *_: self.update_annual_ocr_api_state())
        self.annual_ocr_api_provider.trace_add("write", lambda *_: self.apply_annual_ocr_api_provider())
        self.after(100, self.center_window)

    def center_window(self) -> None:
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        width = min(1360, max(1040, screen_width - 80))
        height = min(1040, max(760, screen_height - 40))
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        self.option_add("*Font", ("Microsoft YaHei UI", 10))

        style.configure("App.TFrame", background="#F5F7FB")
        style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        style.configure("Title.TLabel", background="#F5F7FB", foreground="#122033", font=("Microsoft YaHei UI", 20, "bold"))
        style.configure("Subtitle.TLabel", background="#F5F7FB", foreground="#5B667A", font=("Microsoft YaHei UI", 10))
        style.configure("CardTitle.TLabel", background="#FFFFFF", foreground="#122033", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Hint.TLabel", background="#FFFFFF", foreground="#697386", font=("Microsoft YaHei UI", 9))
        style.configure("Warning.TLabel", background="#FFFFFF", foreground="#C2410C", font=("Microsoft YaHei UI", 9, "bold"))
        style.configure("Status.TLabel", background="#EAF1FF", foreground="#174EA6", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TEntry", fieldbackground="#FFFFFF", bordercolor="#CBD5E1", lightcolor="#CBD5E1", darkcolor="#CBD5E1")
        style.configure("TCombobox", fieldbackground="#FFFFFF", bordercolor="#CBD5E1")
        style.configure("Option.TCheckbutton", background="#FFFFFF", foreground="#233044", font=("Microsoft YaHei UI", 10))
        style.configure("Primary.TButton", background="#1F6FEB", foreground="#FFFFFF", font=("Microsoft YaHei UI", 11, "bold"), padding=(18, 10), borderwidth=0)
        style.map("Primary.TButton", background=[("active", "#185ABC"), ("disabled", "#AAB6C8")])
        style.configure("Secondary.TButton", background="#EEF2F7", foreground="#233044", padding=(12, 8), borderwidth=0)
        style.map("Secondary.TButton", background=[("active", "#E2E8F0")])

    def _add_field_hint(self, parent: ttk.Frame, text: str, wraplength: int = 320) -> None:
        ttk.Label(
            parent,
            text=text,
            style="Hint.TLabel",
            justify="left",
            wraplength=wraplength,
        ).pack(anchor="w", pady=(4, 0))

    def _grid_form_cell(
        self,
        parent: ttk.Frame,
        row: int,
        column: int,
        title: str,
        hint: str,
        *,
        columnspan: int = 1,
        padx: tuple[int, int] = (0, 10),
        pady: tuple[int, int] = (0, 10),
        wraplength: int = 290,
    ) -> ttk.Frame:
        cell = ttk.Frame(parent, style="Card.TFrame")
        cell.grid(row=row, column=column, columnspan=columnspan, sticky="nsew", padx=padx, pady=pady)
        ttk.Label(cell, text=title, style="Hint.TLabel").pack(anchor="w")
        cell._reit_hint_text = hint  # type: ignore[attr-defined]
        cell._reit_hint_wraplength = wraplength  # type: ignore[attr-defined]
        return cell

    def _finish_grid_form_cell(self, cell: ttk.Frame) -> None:
        hint = getattr(cell, "_reit_hint_text", "")
        if hint:
            self._add_field_hint(cell, str(hint), wraplength=int(getattr(cell, "_reit_hint_wraplength", 290)))

    def _make_scrollable_tab(self, parent: ttk.Frame) -> ttk.Frame:
        outer = ttk.Frame(parent, style="App.TFrame")
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, bg="#F5F7FB", highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        container = ttk.Frame(canvas, style="App.TFrame", padding=(8, 8))
        window_id = canvas.create_window((0, 0), window=container, anchor="nw")

        def sync_scroll_region(_event: tk.Event) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def sync_container_width(event: tk.Event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        def on_mousewheel(event: tk.Event) -> None:
            if event.delta:
                canvas.yview_scroll(-int(event.delta / 120), "units")

        container.bind("<Configure>", sync_scroll_region)
        canvas.bind("<Configure>", sync_container_width)
        canvas.bind("<Enter>", lambda _event: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda _event: canvas.unbind_all("<MouseWheel>"))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return container

    def _build_layout(self) -> None:
        root = ttk.Frame(self, style="App.TFrame", padding=(18, 16))
        root.pack(fill="both", expand=True)
        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True)

        conversion_tab = ttk.Frame(notebook, style="App.TFrame")
        annual_tab = ttk.Frame(notebook, style="App.TFrame")
        notebook.add(conversion_tab, text="Excel 自动审核转换")
        notebook.add(annual_tab, text="年报现金流更新")

        container = self._make_scrollable_tab(conversion_tab)

        ttk.Label(container, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            container,
            text="支持单个 Excel 或整个文件夹批量转换；按表头字段自动识别格式，并输出问题与结果汇总表。",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 14))

        card = ttk.Frame(container, style="Card.TFrame", padding=(20, 14))
        card.pack(fill="x")

        self._add_source_row(card)

        action_frame = ttk.Frame(card, style="Card.TFrame")
        action_frame.pack(fill="x", pady=(14, 0))
        self.convert_button = ttk.Button(action_frame, text="开始转换", style="Primary.TButton", command=self.start_conversion)
        self.convert_button.pack(side="left")
        ttk.Label(action_frame, textvariable=self.status_text, style="Status.TLabel", padding=(14, 9)).pack(side="left", padx=(14, 0), fill="x", expand=True)

        ttk.Separator(card).pack(fill="x", pady=14)

        type_frame = ttk.Frame(card, style="Card.TFrame")
        type_frame.pack(fill="x")
        ttk.Label(type_frame, text="转换逻辑", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(type_frame, text="建议使用“自动识别”；批量文件夹会逐个读取表头识别。若一批文件全是同类，也可以手动指定。", style="Hint.TLabel").pack(anchor="w", pady=(2, 8))
        self.type_combo = ttk.Combobox(type_frame, textvariable=self.table_type, values=TYPE_LABELS, state="readonly", height=6)
        self.type_combo.pack(fill="x")
        self.property_processed_check = ttk.Checkbutton(
            type_frame,
            text="需要时额外生成第 4 类产权经营数据处理版（用于修复主配套资产面积字段）",
            variable=self.generate_property_processed,
            style="Option.TCheckbutton",
        )
        self.property_processed_check.pack(anchor="w", pady=(10, 0))
        ttk.Label(
            type_frame,
            text="默认不生成处理版；仅当输入识别为第 4 类且勾选本项时，才会复制原始明细表并修复/补全面积相关字段。",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        ttk.Separator(card).pack(fill="x", pady=14)

        self._add_file_row(
            card,
            title="自定义模板表（可选）",
            hint="上传后进入自定义模板输出模式：程序会按模板字段和格式提取来源表字段，不执行 1-5 类专属修复逻辑。",
            variable=self.custom_template_path,
            command=self.pick_custom_template_file,
            clear_command=self.clear_custom_template_file,
        )
        ttk.Label(card, textvariable=self.template_mode_text, style="Warning.TLabel").pack(anchor="w", pady=(4, 0))

        ttk.Separator(card).pack(fill="x", pady=14)

        self._add_file_row(
            card,
            title="补全信息表（可选）",
            hint="可上传含 REITs代码、REITs名称、上市日期、公告日期、开始日期、结束日期的表；字段不全也会自动使用已有字段。",
            variable=self.metadata_path,
            command=self.pick_metadata_file,
            clear_command=self.clear_metadata_file,
        )

        ttk.Separator(card).pack(fill="x", pady=14)

        self._add_file_row(
            card,
            title="输出文件夹（可选）",
            hint="默认输出到输入文件所在文件夹，或所选输入文件夹内；如需统一保存，可指定输出文件夹。",
            variable=self.output_dir,
            command=self.pick_output_dir,
            clear_command=self.clear_output_dir,
            file_button_text="选择文件夹",
        )

        log_card = ttk.Frame(container, style="Card.TFrame", padding=(18, 14))
        log_card.pack(fill="both", expand=True, pady=(14, 0))
        ttk.Label(log_card, text="转换日志", style="CardTitle.TLabel").pack(anchor="w")
        self.log = tk.Text(
            log_card,
            height=12,
            bg="#0F172A",
            fg="#E5EDF7",
            insertbackground="#FFFFFF",
            relief="flat",
            padx=14,
            pady=12,
            wrap="word",
            font=("Consolas", 10),
        )
        self.log.pack(fill="both", expand=True, pady=(8, 0))
        self.write_log("请选择单个 Excel 或一个文件夹。输出文件名会追加 _自动审核，并生成 自动审核_批量汇总.xlsx。")
        self._build_annual_layout(annual_tab)

    def _add_source_row(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.pack(fill="x")
        ttk.Label(frame, text="输入来源", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(frame, textvariable=self.folder_warning_text, style="Warning.TLabel").pack(anchor="w", pady=(0, 8))

        row = ttk.Frame(frame, style="Card.TFrame")
        row.pack(fill="x")
        ttk.Entry(row, textvariable=self.source_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="选择文件", style="Secondary.TButton", command=self.pick_input_file).pack(side="left", padx=(10, 0))
        ttk.Button(row, text="选择文件夹", style="Secondary.TButton", command=self.pick_input_folder).pack(side="left", padx=(8, 0))
        ttk.Button(row, text="清空", style="Secondary.TButton", command=self.clear_source_path).pack(side="left", padx=(8, 0))
        self._add_field_hint(frame, "操作提示：选择单个未审核 .xlsx，或选择一个文件夹批量处理；程序会逐个读取表头字段识别格式。", 1040)

    def _add_file_row(
        self,
        parent: ttk.Frame,
        title: str,
        hint: str,
        variable: tk.StringVar,
        command: Callable[[], None],
        clear_command: Callable[[], None] | None = None,
        file_button_text: str = "选择文件",
    ) -> None:
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.pack(fill="x")
        ttk.Label(frame, text=title, style="CardTitle.TLabel").pack(anchor="w")

        row = ttk.Frame(frame, style="Card.TFrame")
        row.pack(fill="x", pady=(6, 0))
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text=file_button_text, style="Secondary.TButton", command=command).pack(side="left", padx=(10, 0))
        if clear_command:
            ttk.Button(row, text="清空", style="Secondary.TButton", command=clear_command).pack(side="left", padx=(8, 0))
        self._add_field_hint(frame, f"操作提示：{hint}", 1040)

    def pick_input_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择未审核 Excel 文件",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.source_path.set(path)
            self.status_text.set("已选择单个输入文件")
            self.update_folder_warning()
            self.write_log(f"输入文件：{path}")

    def pick_input_folder(self) -> None:
        path = filedialog.askdirectory(title="选择包含未审核 Excel 的文件夹")
        if path:
            self.source_path.set(path)
            self.status_text.set("已选择输入文件夹")
            self.update_folder_warning()
            self.write_log(f"输入文件夹：{path}")

    def clear_source_path(self) -> None:
        self.source_path.set("")
        self.update_folder_warning()
        self.write_log("已清空输入来源。")

    def pick_custom_template_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择自定义模板表",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.custom_template_path.set(path)
            self.write_log(f"自定义模板表：{path}")

    def clear_custom_template_file(self) -> None:
        self.custom_template_path.set("")
        self.write_log("已清空自定义模板表，恢复普通转换模式。")

    def update_template_mode(self) -> None:
        if self.custom_template_path.get().strip():
            self.type_combo.configure(state="disabled")
            self.property_processed_check.configure(state="disabled")
            self.template_mode_text.set("当前为自定义模板模式：转换逻辑选择已停用，仅按上传模板字段和格式输出。")
        else:
            self.type_combo.configure(state="readonly")
            self.property_processed_check.configure(state="normal")
            self.template_mode_text.set("")
        self.update_folder_warning()

    def update_folder_warning(self) -> None:
        source = self.source_path.get().strip()
        custom_template = self.custom_template_path.get().strip()
        if source and custom_template and Path(source).is_dir():
            self.folder_warning_text.set("自定义模板文件夹模式：请确认该文件夹内 Excel 均为相同来源格式；程序会自动校验表头相似度。")
        else:
            self.folder_warning_text.set("")

    def pick_metadata_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择补全信息表",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.metadata_path.set(path)
            self.write_log(f"补全信息表：{path}")

    def clear_metadata_file(self) -> None:
        self.metadata_path.set("")
        self.write_log("已清空补全信息表。")

    def pick_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.output_dir.set(path)
            self.write_log(f"输出文件夹：{path}")

    def clear_output_dir(self) -> None:
        self.output_dir.set("")
        self.write_log("已恢复默认输出位置。")

    def _build_annual_layout(self, parent: ttk.Frame) -> None:
        container = self._make_scrollable_tab(parent)

        ttk.Label(container, text="年报现金流更新", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            container,
            text="按“选择材料 -> 基础设置 -> OCR/AI -> 输出运行”的顺序填写。默认本地处理，只有主动启用 AI 或 vision_api 才会调用外部接口。",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 14))

        material_card = ttk.Frame(container, style="Card.TFrame", padding=(20, 14))
        material_card.pack(fill="x")
        ttk.Label(material_card, text="1. 选择材料", style="CardTitle.TLabel").pack(anchor="w")
        self._add_field_hint(
            material_card,
            "必选：年报工作文件夹。可选：标准导入表、人工现金流资料、公募年报 PDF 文件夹。",
            1040,
        )

        self._add_file_row(
            material_card,
            title="年报工作文件夹（必选）",
            hint="放去年产权/特许经营权表、管理费率表、评估价值表等辅助材料。如需严格对齐审核格式，可在工作文件夹或项目根目录放已核/参考样表；软件只借用格式，未来现金流表由软件生成。",
            variable=self.annual_workspace_path,
            command=self.pick_annual_workspace,
            clear_command=self.clear_annual_workspace,
            file_button_text="选择文件夹",
        )

        ttk.Separator(material_card).pack(fill="x", pady=12)

        self._add_file_row(
            material_card,
            title="标准导入表（可选，最稳）",
            hint="如果已人工整理现金流、费率、评估值等标准表，可直接上传；未上传时，软件会使用 OCR/AI 识别结果。",
            variable=self.annual_standard_input_path,
            command=self.pick_annual_standard_input,
            clear_command=self.clear_annual_standard_input,
        )

        ttk.Separator(material_card).pack(fill="x", pady=12)

        self._add_file_row(
            material_card,
            title="人工现金流资料（可选）",
            hint="只放人工截取出的现金流 PDF、截图或 Word 图片；完整公募年报不要放这里，年报会单独直接提取文字。",
            variable=self.annual_ocr_source_path,
            command=self.pick_annual_ocr_source,
            clear_command=self.clear_annual_ocr_source,
            file_button_text="选择文件夹",
        )

        ttk.Separator(material_card).pack(fill="x", pady=12)

        self._add_file_row(
            material_card,
            title="公募年报 PDF 文件夹（可选）",
            hint="如果年报 PDF 与工作文件夹分开放置，请单独选择；用于提取基金净资产、折旧摊销等，不走截图 OCR。",
            variable=self.annual_report_source_path,
            command=self.pick_annual_report_source,
            clear_command=self.clear_annual_report_source,
            file_button_text="选择文件夹",
        )

        settings_card = ttk.Frame(container, style="Card.TFrame", padding=(20, 14))
        settings_card.pack(fill="x", pady=(14, 0))
        ttk.Label(settings_card, text="2. 基础设置", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            settings_card,
            text="建议先用默认设置跑一遍。本地 OCR 不需要 API；严格口径用于复核本轮材料是否足够完整。",
            style="Hint.TLabel",
            wraplength=1040,
            justify="left",
        ).pack(anchor="w", pady=(4, 10))

        settings_frame = ttk.Frame(settings_card, style="Card.TFrame")
        settings_frame.pack(fill="x")
        ocr_engine_cell = self._grid_form_cell(
            settings_frame,
            0,
            0,
            "OCR 引擎",
            "推荐 auto。本地 OCR 不需要 API；只有 vision_api 会调用云端 OCR。",
        )
        ttk.Combobox(
            ocr_engine_cell,
            textvariable=self.annual_ocr_engine,
            values=["auto", "rapidocr", "pdf_text", "paddleocr", "pytesseract", "vision_api"],
            state="readonly",
        ).pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(ocr_engine_cell)

        max_pages_cell = self._grid_form_cell(
            settings_frame,
            0,
            1,
            "每个 PDF 最多扫描页数",
            "建议 3-5。填 -1 表示跳过 PDF OCR，填 0 表示扫描全部页。",
        )
        ttk.Entry(max_pages_cell, textvariable=self.annual_max_pages, width=12).pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(max_pages_cell)

        start_year_cell = self._grid_form_cell(
            settings_frame,
            0,
            2,
            "正式表起始年份",
            "例如做 2027 年预测更新就填 2027；正式表只写入该年份及之后数据。",
        )
        ttk.Entry(start_year_cell, textvariable=self.annual_output_start_year, width=12).pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(start_year_cell)

        strict_cell = self._grid_form_cell(
            settings_frame,
            0,
            3,
            "严格口径复核模式",
            "勾选后只使用本轮材料，缺失字段留空并写入复核信息；不再用去年表补齐空字段。",
            padx=(0, 0),
        )
        ttk.Checkbutton(
            strict_cell,
            text="只用本轮材料，不用去年表补空",
            variable=self.annual_strict_mode,
            style="Option.TCheckbutton",
        ).pack(anchor="w", pady=(4, 0))
        self._finish_grid_form_cell(strict_cell)
        for column in range(4):
            settings_frame.columnconfigure(column, weight=1)

        ai_card = ttk.Frame(container, style="Card.TFrame", padding=(20, 14))
        ai_card.pack(fill="x", pady=(14, 0))
        ttk.Label(ai_card, text="3. AI 标准化（可选）", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            ai_card,
            text="默认不调用 AI。启用后只发送 OCR 提取出的文本，用于整理成标准导入表；不会上传 PDF、截图或 Excel 文件。",
            style="Warning.TLabel",
            wraplength=1040,
            justify="left",
        ).pack(anchor="w", pady=(4, 10))

        ai_frame = ttk.Frame(ai_card, style="Card.TFrame")
        ai_frame.pack(fill="x")
        ai_toggle_cell = self._grid_form_cell(
            ai_frame,
            0,
            0,
            "AI 标准化",
            "不勾选则不调用外部模型。勾选后只发送 OCR 文本，不上传 PDF、截图或 Excel。",
        )
        ttk.Checkbutton(
            ai_toggle_cell,
            text="启用 AI 标准化 OCR 文本",
            variable=self.annual_use_ai,
            style="Option.TCheckbutton",
        ).pack(anchor="w", pady=(4, 0))
        self._finish_grid_form_cell(ai_toggle_cell)

        ai_provider_cell = self._grid_form_cell(
            ai_frame,
            0,
            1,
            "AI 服务商预设",
            "选择后会自动填模型和 Base URL；API Key 或环境变量名需由用户自行填写。",
        )
        ttk.Combobox(
            ai_provider_cell,
            textvariable=self.annual_ai_provider,
            values=list(AI_PROVIDER_PROFILES),
            state="readonly",
            width=26,
        ).pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(ai_provider_cell)

        model_cell = self._grid_form_cell(ai_frame, 0, 2, "模型", "默认 qwen-flash。只做 OCR 文本整理，通常不需要大模型。")
        ttk.Entry(model_cell, textvariable=self.annual_model, width=22).pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(model_cell)

        base_url_cell = self._grid_form_cell(
            ai_frame,
            0,
            3,
            "API Base URL",
            "OpenAI 兼容接口地址。使用服务商预设后一般不用改。",
            padx=(0, 0),
        )
        ttk.Entry(base_url_cell, textvariable=self.annual_base_url).pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(base_url_cell)

        api_key_cell = self._grid_form_cell(
            ai_frame,
            1,
            0,
            "API Key（临时输入）",
            "可直接填本次运行的 Key；优先级高于环境变量，但软件不会保存到文件。",
        )
        ttk.Entry(api_key_cell, textvariable=self.annual_api_key, show="*").pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(api_key_cell)

        api_env_cell = self._grid_form_cell(
            ai_frame,
            1,
            1,
            "API Key 环境变量名（推荐）",
            "如使用系统环境变量，请填写变量名；直接在左侧临时输入 API Key 时这里可留空。",
            columnspan=3,
            padx=(0, 0),
            wraplength=820,
        )
        ttk.Entry(api_env_cell, textvariable=self.annual_api_key_env).pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(api_env_cell)

        ttk.Label(
            ai_frame,
            text=(
                "API Key 读取优先级：界面临时输入 > 用户填写的环境变量名。推荐使用环境变量，避免把 Key 留在截图或共享文件中。AI 默认按单个 OCR 材料片段逐批整理，"
                f"默认单批最多 {DEFAULT_AI_REQUEST_TIMEOUT_SECONDS} 秒、整轮最多 {DEFAULT_AI_TOTAL_TIMEOUT_SECONDS} 秒，超时会跳过并写入复核清单。"
            ),
            style="Hint.TLabel",
            wraplength=1120,
            justify="left",
        ).grid(row=2, column=0, columnspan=4, sticky="w", pady=(2, 0))
        for column in range(4):
            ai_frame.columnconfigure(column, weight=1)

        cloud_card = ttk.Frame(container, style="Card.TFrame", padding=(20, 14))
        cloud_card.pack(fill="x", pady=(14, 0))
        ttk.Label(cloud_card, text="4. 云端 OCR API（高级，可选）", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            cloud_card,
            text="只有 OCR 引擎选择 vision_api 时才会使用。普通本地 OCR 用户可以整块留空，软件会禁用这些输入框。",
            style="Warning.TLabel",
            wraplength=1040,
            justify="left",
        ).pack(anchor="w", pady=(4, 10))
        ttk.Label(
            cloud_card,
            textvariable=self.annual_ocr_api_state_text,
            style="Hint.TLabel",
            wraplength=1040,
            justify="left",
        ).pack(anchor="w", pady=(0, 10))

        cloud_frame = ttk.Frame(cloud_card, style="Card.TFrame")
        cloud_frame.pack(fill="x")
        ocr_provider_cell = self._grid_form_cell(
            cloud_frame,
            0,
            0,
            "OCR API 服务商预设",
            "仅 vision_api 启用。选择后自动填 OCR 模型和 Base URL；API Key 或环境变量名需由用户自行填写。",
        )
        ocr_provider_combo = ttk.Combobox(
            ocr_provider_cell,
            textvariable=self.annual_ocr_api_provider,
            values=list(OCR_API_PROVIDER_PROFILES),
            state="readonly",
            width=28,
        )
        ocr_provider_combo.pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(ocr_provider_cell)

        ocr_model_cell = self._grid_form_cell(cloud_frame, 0, 1, "OCR 模型", "默认 qwen-vl-ocr-latest；按服务商实际可用模型修改。")
        ocr_model_entry = ttk.Entry(ocr_model_cell, textvariable=self.annual_ocr_model)
        ocr_model_entry.pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(ocr_model_cell)

        ocr_base_cell = self._grid_form_cell(cloud_frame, 0, 2, "OCR API Base URL", "OpenAI 兼容视觉接口地址。使用预设后一般不用改。")
        ocr_base_entry = ttk.Entry(ocr_base_cell, textvariable=self.annual_ocr_base_url)
        ocr_base_entry.pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(ocr_base_cell)

        ocr_key_cell = self._grid_form_cell(
            cloud_frame,
            0,
            3,
            "OCR API Key（临时输入）",
            "可直接填本次运行的 Key；留空时读取右下方环境变量名。",
            padx=(0, 0),
        )
        ocr_key_entry = ttk.Entry(ocr_key_cell, textvariable=self.annual_ocr_api_key, show="*")
        ocr_key_entry.pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(ocr_key_cell)

        ocr_env_cell = self._grid_form_cell(
            cloud_frame,
            1,
            0,
            "OCR API Key 环境变量名（推荐）",
            "如使用系统环境变量，请填写变量名；直接在 OCR API Key 输入框填 Key 时这里可留空。",
            columnspan=4,
            padx=(0, 0),
            wraplength=1120,
        )
        ocr_key_env_entry = ttk.Entry(ocr_env_cell, textvariable=self.annual_ocr_api_key_env)
        ocr_key_env_entry.pack(fill="x", pady=(4, 0))
        self._finish_grid_form_cell(ocr_env_cell)
        self.ocr_api_widgets = [ocr_provider_combo, ocr_model_entry, ocr_base_entry, ocr_key_entry, ocr_key_env_entry]
        self.update_annual_ocr_api_state()
        for column in range(4):
            cloud_frame.columnconfigure(column, weight=1)

        output_card = ttk.Frame(container, style="Card.TFrame", padding=(20, 14))
        output_card.pack(fill="x", pady=(14, 0))
        ttk.Label(output_card, text="5. 输出与运行", style="CardTitle.TLabel").pack(anchor="w")
        self._add_field_hint(output_card, "输出会包含结果表、必要过程表和复核报告；正式表优先使用真实已核表或参考样表的格式，缺失字段会保留为空并写入复核信息。", 1040)

        self._add_file_row(
            output_card,
            title="输出文件夹（可选）",
            hint="默认输出到工作文件夹下的“年度更新_输出结果”。",
            variable=self.annual_output_dir,
            command=self.pick_annual_output_dir,
            clear_command=self.clear_annual_output_dir,
            file_button_text="选择文件夹",
        )

        action_frame = ttk.Frame(output_card, style="Card.TFrame")
        action_frame.pack(fill="x", pady=(14, 0))
        self.annual_button = ttk.Button(action_frame, text="开始年报更新", style="Primary.TButton", command=self.start_annual_update)
        self.annual_button.pack(side="left")
        ttk.Label(action_frame, textvariable=self.annual_status_text, style="Status.TLabel", padding=(14, 9)).pack(
            side="left", padx=(14, 0), fill="x", expand=True
        )

        log_card = ttk.Frame(container, style="Card.TFrame", padding=(18, 14))
        log_card.pack(fill="both", expand=True, pady=(14, 0))
        ttk.Label(log_card, text="运行日志与提示", style="CardTitle.TLabel").pack(anchor="w")
        self.annual_log = tk.Text(
            log_card,
            height=10,
            bg="#0F172A",
            fg="#E5EDF7",
            insertbackground="#FFFFFF",
            relief="flat",
            padx=14,
            pady=12,
            wrap="word",
            font=("Consolas", 10),
        )
        self.annual_log.pack(fill="both", expand=True, pady=(8, 0))
        self.write_annual_log("请选择年报工作文件夹。建议第一轮先不启用 AI，确认材料识别和过程表正常后，再按需打开 AI 标准化。")

    def apply_annual_ai_provider(self) -> None:
        profile = AI_PROVIDER_PROFILES.get(self.annual_ai_provider.get())
        if not profile:
            return
        base_url = profile.get("base_url", "")
        model = profile.get("model", "")
        if base_url:
            self.annual_base_url.set(base_url)
        if model:
            self.annual_model.set(model)
        key_env = profile.get("key_env", "")
        if key_env:
            self.annual_api_key_env.set(key_env)
        note = profile.get("note")
        if note and hasattr(self, "annual_log"):
            self.write_annual_log(f"AI 服务商预设：{self.annual_ai_provider.get()}。{note}")

    def apply_annual_ocr_api_provider(self) -> None:
        profile = OCR_API_PROVIDER_PROFILES.get(self.annual_ocr_api_provider.get())
        if not profile:
            return
        self.annual_ocr_base_url.set(profile.get("base_url", ""))
        self.annual_ocr_model.set(profile.get("model", ""))
        key_env = profile.get("key_env", "")
        if key_env:
            self.annual_ocr_api_key_env.set(key_env)
        note = profile.get("note")
        if note and hasattr(self, "annual_log"):
            self.write_annual_log(f"OCR API 预设：{self.annual_ocr_api_provider.get()}。{note}")

    def update_annual_ocr_api_state(self) -> None:
        widgets = getattr(self, "ocr_api_widgets", [])
        is_cloud_ocr = self.annual_ocr_engine.get().strip().lower() == "vision_api"
        state = "normal" if is_cloud_ocr else "disabled"
        for widget in widgets:
            try:
                if isinstance(widget, ttk.Combobox):
                    widget.configure(state="readonly" if is_cloud_ocr else "disabled")
                else:
                    widget.configure(state=state)
            except tk.TclError:
                pass
        if is_cloud_ocr:
            self.annual_ocr_api_state_text.set("当前已启用：云端 OCR API 输入框可以编辑；运行时会上传需要 OCR 的截图或 PDF 渲染页。")
            if self.annual_ocr_api_provider.get() == "不使用云端 OCR（本地 OCR，无需 API）":
                self.annual_ocr_api_provider.set("阿里云百炼 / 通义视觉 OCR")
            elif not self.annual_ocr_model.get().strip():
                self.apply_annual_ocr_api_provider()
        else:
            self.annual_ocr_api_state_text.set("当前未启用：如需云端 OCR，请先在“基础设置”中把 OCR 引擎改为 vision_api。")
            if self.annual_ocr_api_provider.get() != "不使用云端 OCR（本地 OCR，无需 API）":
                self.annual_ocr_api_provider.set("不使用云端 OCR（本地 OCR，无需 API）")

    def pick_annual_workspace(self) -> None:
        path = filedialog.askdirectory(title="选择年报现金流工作文件夹")
        if path:
            self.annual_workspace_path.set(path)
            self.annual_status_text.set("已选择年报工作文件夹")
            self.write_annual_log(f"年报工作文件夹：{path}")

    def clear_annual_workspace(self) -> None:
        self.annual_workspace_path.set("")
        self.write_annual_log("已清空年报工作文件夹。")

    def pick_annual_standard_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择标准导入表",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.annual_standard_input_path.set(path)
            self.write_annual_log(f"标准导入表：{path}")

    def clear_annual_standard_input(self) -> None:
        self.annual_standard_input_path.set("")
        self.write_annual_log("已清空标准导入表，将使用 OCR/AI 标准化结果；未来现金流表会作为输出重新生成。")

    def pick_annual_ocr_source(self) -> None:
        path = filedialog.askdirectory(title="选择人工截取的 OCR 来源文件夹")
        if path:
            self.annual_ocr_source_path.set(path)
            self.write_annual_log(f"OCR 来源文件夹：{path}")

    def clear_annual_ocr_source(self) -> None:
        self.annual_ocr_source_path.set("")
        self.write_annual_log("已清空 OCR 来源文件夹，将自动寻找人工现金流示例文件夹。")

    def pick_annual_report_source(self) -> None:
        path = filedialog.askdirectory(title="选择公募年报 PDF 文件夹")
        if path:
            self.annual_report_source_path.set(path)
            self.write_annual_log(f"公募年报文件夹：{path}")

    def clear_annual_report_source(self) -> None:
        self.annual_report_source_path.set("")
        self.write_annual_log("已清空公募年报文件夹；若工作文件夹中没有年报 PDF，则净资产/折旧摊销不会自动提取。")

    def pick_annual_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择年报更新输出文件夹")
        if path:
            self.annual_output_dir.set(path)
            self.write_annual_log(f"年报更新输出文件夹：{path}")

    def clear_annual_output_dir(self) -> None:
        self.annual_output_dir.set("")
        self.write_annual_log("已恢复年报更新默认输出位置。")

    def start_annual_update(self) -> None:
        workspace_path = self.annual_workspace_path.get().strip()
        if not workspace_path:
            messagebox.showwarning("缺少年报工作文件夹", "请先选择年报现金流工作文件夹。")
            return
        try:
            max_pages = int(self.annual_max_pages.get().strip() or "5")
        except ValueError:
            messagebox.showwarning("OCR 页数设置错误", "每个 PDF 最多扫描页数应为整数；-1 表示跳过，0 表示全部。")
            return
        try:
            output_start_year = int(self.annual_output_start_year.get().strip() or "2027")
        except ValueError:
            messagebox.showwarning("起始年份设置错误", "正式表起始年份应为 4 位年份，例如 2027。")
            return

        self.annual_button.state(["disabled"])
        self.annual_status_text.set("正在更新，请稍候...")
        self.write_annual_log("开始年报现金流更新...")
        if self.annual_use_ai.get():
            self.write_annual_log(
                "已启用 AI 标准化：只发送 OCR 文本，不上传 PDF、截图或目标 Excel；"
                f"单批最多 {DEFAULT_AI_REQUEST_TIMEOUT_SECONDS} 秒，整轮最多 {DEFAULT_AI_TOTAL_TIMEOUT_SECONDS} 秒。"
            )

        options = AnnualUpdateOptions(
            workspace_path=workspace_path,
            output_dir=self.annual_output_dir.get().strip() or None,
            standard_input_path=self.annual_standard_input_path.get().strip() or None,
            ocr_source_path=self.annual_ocr_source_path.get().strip() or None,
            annual_report_source_path=self.annual_report_source_path.get().strip() or None,
            ocr_engine=self.annual_ocr_engine.get(),
            use_ai=bool(self.annual_use_ai.get()),
            api_key=self.annual_api_key.get().strip() or None,
            api_key_env=self.annual_api_key_env.get().strip(),
            base_url=self.annual_base_url.get().strip() or DEFAULT_DASHSCOPE_BASE_URL,
            model=self.annual_model.get().strip() or DEFAULT_AI_MODEL,
            ocr_api_key=self.annual_ocr_api_key.get().strip() or None,
            ocr_api_key_env=self.annual_ocr_api_key_env.get().strip(),
            ocr_base_url=self.annual_ocr_base_url.get().strip() or DEFAULT_DASHSCOPE_BASE_URL,
            ocr_model=self.annual_ocr_model.get().strip() or DEFAULT_VISION_OCR_MODEL,
            output_start_year=output_start_year,
            max_ocr_pages_per_file=max_pages,
            allow_existing_context_fill=not bool(self.annual_strict_mode.get()),
            progress=lambda message: self.after(0, self.write_annual_log, message),
        )
        thread = threading.Thread(target=self._run_annual_update, args=(options,), daemon=True)
        thread.start()

    def _run_annual_update(self, options: AnnualUpdateOptions) -> None:
        try:
            result = run_annual_update(options)
        except AnnualUpdateError as exc:
            self.after(0, self._annual_update_failed, str(exc))
        except Exception as exc:
            self.after(0, self._annual_update_failed, f"年报更新失败：{exc}")
        else:
            self.after(0, self._annual_update_done, result)

    def _annual_update_done(self, result: AnnualUpdateResult) -> None:
        self.annual_button.state(["!disabled"])
        self.annual_status_text.set(f"更新完成：标准化 {result.standard_row_count} 行，OCR {result.ocr_item_count} 条")
        self.write_annual_log(f"汇总表：{result.summary_file}")
        if result.ocr_file == result.summary_file:
            self.write_annual_log(f"OCR原始识别、AI调用记录、更新计划、人工复核清单和输出对比已合并在 {SUMMARY_OUTPUT_NAME} 中。")
        else:
            self.write_annual_log(f"OCR结果：{result.ocr_file}")
        if result.standard_file == result.summary_file:
            self.write_annual_log("标准化导入表已合并在汇总表的“标准化导入表”工作表中。")
        else:
            self.write_annual_log(f"标准导入表：{result.standard_file}")
        if result.future_cashflow_file:
            self.write_annual_log(f"未来现金流表：{result.future_cashflow_file}")
        if result.annual_report_extract_file:
            self.write_annual_log(f"年报净资产/折旧摊销提取表：{result.annual_report_extract_file}")
        if result.property_file:
            self.write_annual_log(f"产权表：{result.property_file}")
        if result.concession_file:
            self.write_annual_log(f"特许经营权表：{result.concession_file}")
        if result.ai_call_file:
            self.write_annual_log(f"AI调用记录：{result.ai_call_file}")
        if result.plan_file != result.summary_file:
            self.write_annual_log(f"更新计划：{result.plan_file}")
            self.write_annual_log(f"人工复核清单：{result.review_file}")
            self.write_annual_log(f"输出对比检查：{result.comparison_file}")
        for warning in result.warnings:
            self.write_annual_log(f"提示：{warning}")
        messagebox.showinfo(
            "年报更新完成",
            f"年报现金流更新已完成。\n标准化行数：{result.standard_row_count}\nOCR记录数：{result.ocr_item_count}\n输出文件夹：{result.output_dir}",
        )

    def _annual_update_failed(self, message: str) -> None:
        self.annual_button.state(["!disabled"])
        self.annual_status_text.set("年报更新失败")
        self.write_annual_log(message)
        messagebox.showerror("年报更新失败", message)

    def write_annual_log(self, message: str) -> None:
        self.annual_log.insert("end", message + "\n")
        self.annual_log.see("end")

    def start_conversion(self) -> None:
        source_path = self.source_path.get().strip()
        if not source_path:
            messagebox.showwarning("缺少输入来源", "请先选择一个未审核 Excel 文件或文件夹。")
            return

        selected_label = self.table_type.get()
        selected_type = LABEL_TO_TYPE.get(selected_label, AUTO_TYPE)
        metadata_path = self.metadata_path.get().strip() or None
        custom_template_path = self.custom_template_path.get().strip() or None
        output_dir = self.output_dir.get().strip() or None
        generate_property_processed = bool(self.generate_property_processed.get())

        self.convert_button.state(["disabled"])
        self.status_text.set("正在转换，请稍候...")
        self.write_log("开始转换...")
        if custom_template_path:
            self.write_log("当前模式：自定义模板输出。")
        elif generate_property_processed:
            self.write_log("已启用：第 4 类产权经营数据处理版输出。")

        thread = threading.Thread(
            target=self._run_conversion,
            args=(source_path, selected_type, metadata_path, custom_template_path, output_dir, generate_property_processed),
            daemon=True,
        )
        thread.start()

    def _run_conversion(
        self,
        source_path: str,
        selected_type: str,
        metadata_path: str | None,
        custom_template_path: str | None,
        output_dir: str | None,
        generate_property_processed: bool,
    ) -> None:
        try:
            result = convert_input_path(
                input_path=source_path,
                selected_type=selected_type,
                metadata_path=metadata_path,
                custom_template_path=custom_template_path,
                output_dir=output_dir,
                generate_property_processed=generate_property_processed,
            )
        except ConversionError as exc:
            self.after(0, self._conversion_failed, str(exc))
        except Exception as exc:  # Last-resort guard for a desktop app.
            self.after(0, self._conversion_failed, f"转换失败：{exc}")
        else:
            self.after(0, self._conversion_done, result)

    def _conversion_done(self, result: BatchResult) -> None:
        self.convert_button.state(["!disabled"])
        self.status_text.set(f"转换完成：共 {result.total_count} 个，成功 {result.success_count} 个，失败 {result.failed_count} 个")
        self.write_log(f"汇总表：{result.summary_file}")

        for item in result.items:
            if item.status == "成功":
                self.write_log(f"成功：{item.input_file.name} | {TABLE_TYPES.get(item.detected_type or '', item.detected_type or '')} | {item.row_count} 行")
                for path in item.output_files:
                    self.write_log(f"  输出：{path}")
                for warning in item.warnings:
                    self.write_log(f"  提示：{warning}")
            else:
                self.write_log(f"失败：{item.input_file.name} | {item.error}")

        messagebox.showinfo(
            "转换完成",
            f"转换已完成。\n成功：{result.success_count}\n失败：{result.failed_count}\n汇总表：{result.summary_file}",
        )

    def _conversion_failed(self, message: str) -> None:
        self.convert_button.state(["!disabled"])
        self.status_text.set("转换失败")
        self.write_log(message)
        messagebox.showerror("转换失败", message)

    def write_log(self, message: str) -> None:
        self.log.insert("end", message + "\n")
        self.log.see("end")


def enable_high_dpi() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def resource_path(relative_path: Path) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
    return base_path / relative_path


def set_window_icon(window: tk.Tk) -> None:
    icon_path = resource_path(APP_ICON_RELATIVE)
    if not icon_path.exists():
        return
    try:
        window.iconbitmap(str(icon_path))
    except tk.TclError:
        pass


def run_cli(argv: list[str]) -> int | None:
    if not argv:
        return None
    configure_cli_streams()

    parser = argparse.ArgumentParser(description="REITs Excel 自动审核转换工具")
    parser.add_argument("--convert", help="要转换的未审核 Excel 文件或文件夹路径")
    parser.add_argument("--type", default=AUTO_TYPE, choices=list(TABLE_TYPES.keys()), help="转换类型，默认 auto")
    parser.add_argument("--metadata", default=None, help="可选补全信息表路径")
    parser.add_argument("--custom-template", default=None, help="可选自定义模板表路径；提供后将按模板字段和格式输出")
    parser.add_argument("--output-dir", default=None, help="可选输出目录")
    parser.add_argument("--property-processed", action="store_true", help="可选；第4类产权经营数据额外生成处理版，用于修复主配套资产面积字段")
    parser.add_argument("--annual-update", help="执行年报现金流更新的工作文件夹路径")
    parser.add_argument("--annual-standard-input", default=None, help="可选标准导入表路径；未提供时使用 OCR/AI 标准化结果，未来现金流表会作为输出重新生成")
    parser.add_argument("--annual-ocr-source", default=None, help="可选 OCR 来源文件夹；推荐使用人工截取出的现金流 PDF/截图文件夹")
    parser.add_argument("--annual-report-source", default=None, help="可选公募年报 PDF 文件夹；与工作文件夹分开放置时用于直接提取净资产和折旧摊销")
    parser.add_argument("--annual-ocr-engine", default="auto", choices=["auto", "rapidocr", "pdf_text", "paddleocr", "pytesseract", "vision_api"], help="OCR 引擎；vision_api 为云端视觉 OCR")
    parser.add_argument("--annual-ocr-api-key", default=None, help="可选；云端视觉 OCR API Key。为空时读取 --annual-ocr-api-key-env 指定的环境变量")
    parser.add_argument("--annual-ocr-api-key-env", default="DASHSCOPE_API_KEY", help="可选；云端视觉 OCR API Key 环境变量名")
    parser.add_argument("--annual-ocr-base-url", default=DEFAULT_DASHSCOPE_BASE_URL, help="云端视觉 OCR OpenAI 兼容 Base URL")
    parser.add_argument("--annual-ocr-model", default=DEFAULT_VISION_OCR_MODEL, help="云端视觉 OCR 模型名称")
    parser.add_argument("--annual-use-ai", action="store_true", help="启用 AI 将 OCR 文本整理为标准导入表")
    parser.add_argument("--annual-api-key", default=None, help="可选；AI API Key。为空时读取 --annual-api-key-env 指定的环境变量")
    parser.add_argument("--annual-api-key-env", default="DASHSCOPE_API_KEY", help="可选；AI API Key 环境变量名。OpenAI 可设为 OPENAI_API_KEY，本地 Ollama 可留空")
    parser.add_argument("--annual-base-url", default=DEFAULT_DASHSCOPE_BASE_URL, help="OpenAI 兼容接口 Base URL")
    parser.add_argument("--annual-model", default=DEFAULT_AI_MODEL, help="AI 标准化模型名称")
    parser.add_argument("--annual-ai-max-chars", type=int, default=None, help="AI 每批 OCR 文本最大字符数，默认使用内置安全值")
    parser.add_argument("--annual-ai-items-per-batch", type=int, default=1, help="AI 每批最多处理的 OCR 记录数，默认 1，适合逐基金/逐页面复核")
    parser.add_argument("--annual-ai-timeout", type=int, default=DEFAULT_AI_REQUEST_TIMEOUT_SECONDS, help="AI 单批请求超时秒数，默认 60")
    parser.add_argument("--annual-ai-total-timeout", type=int, default=DEFAULT_AI_TOTAL_TIMEOUT_SECONDS, help="AI 整轮总时长上限秒数，默认 300；0 表示不限制")
    parser.add_argument("--annual-output-start-year", type=int, default=2027, help="产权/特许经营权正式表写入起始年份，默认 2027")
    parser.add_argument("--annual-max-ocr-pages", type=int, default=5, help="每个 PDF 最多扫描页数；-1 跳过 OCR/文本抽取，0 扫描全部页面")
    parser.add_argument("--annual-skip-excel-open-check", action="store_true", help="跳过本机 Excel 打开校验和兼容修复")
    parser.add_argument("--annual-detailed-output-files", action="store_true", help="兼容旧流程：将 OCR、更新计划、复核清单、输出对比、AI调用记录拆成多个独立文件")
    parser.add_argument("--annual-strict-mode", action="store_true", help="严格口径：不使用去年源表补齐本轮未提供的空字段")
    args = parser.parse_args(argv)

    if args.annual_update:
        result = run_annual_update(
            AnnualUpdateOptions(
                workspace_path=args.annual_update,
                output_dir=args.output_dir,
                standard_input_path=args.annual_standard_input,
                ocr_source_path=args.annual_ocr_source,
                annual_report_source_path=args.annual_report_source,
                ocr_engine=args.annual_ocr_engine,
                use_ai=args.annual_use_ai,
                api_key=args.annual_api_key,
                api_key_env=args.annual_api_key_env,
                base_url=args.annual_base_url,
                model=args.annual_model,
                ocr_api_key=args.annual_ocr_api_key,
                ocr_api_key_env=args.annual_ocr_api_key_env,
                ocr_base_url=args.annual_ocr_base_url,
                ocr_model=args.annual_ocr_model,
                output_start_year=args.annual_output_start_year,
                max_ocr_pages_per_file=args.annual_max_ocr_pages,
                max_ai_chars=args.annual_ai_max_chars or DEFAULT_AI_BATCH_CHAR_LIMIT,
                ai_items_per_batch=args.annual_ai_items_per_batch,
                ai_request_timeout_seconds=args.annual_ai_timeout,
                ai_total_timeout_seconds=args.annual_ai_total_timeout,
                excel_open_check=not args.annual_skip_excel_open_check,
                compact_outputs=not args.annual_detailed_output_files,
                allow_existing_context_fill=not args.annual_strict_mode,
            )
        )
        cli_emit(f"summary={result.summary_file}")
        if result.comparison_file != result.summary_file:
            cli_emit(f"comparison={result.comparison_file}")
            cli_emit(f"ocr={result.ocr_file}")
        if result.standard_file == result.summary_file:
            cli_emit(f"standard={SUMMARY_OUTPUT_NAME}#标准化导入表")
        else:
            cli_emit(f"standard={result.standard_file}")
        if result.future_cashflow_file:
            cli_emit(f"future_cashflow={result.future_cashflow_file}")
        if result.annual_report_extract_file:
            cli_emit(f"annual_report_extract={result.annual_report_extract_file}")
        if result.property_file:
            cli_emit(f"property={result.property_file}")
        if result.concession_file:
            cli_emit(f"concession={result.concession_file}")
        if result.ai_call_file:
            cli_emit(f"ai_calls={result.ai_call_file}")
        cli_emit(f"standard_rows={result.standard_row_count}")
        cli_emit(f"ocr_items={result.ocr_item_count}")
        for warning in result.warnings:
            cli_emit(f"warning={warning}")
        return 0

    if not args.convert:
        return None

    result = convert_input_path(
        input_path=args.convert,
        selected_type=args.type,
        metadata_path=args.metadata,
        custom_template_path=args.custom_template,
        output_dir=args.output_dir,
        generate_property_processed=args.property_processed,
    )
    cli_emit(f"summary={result.summary_file}")
    cli_emit(f"total={result.total_count}")
    cli_emit(f"success={result.success_count}")
    cli_emit(f"failed={result.failed_count}")
    for item in result.items:
        cli_emit(f"item={item.status}|{item.detected_type or ''}|{item.input_file}")
        if item.error:
            cli_emit(f"error={item.error}")
        for output_file in item.output_files:
            cli_emit(f"output={output_file}")
        for warning in item.warnings:
            cli_emit(f"warning={warning}")
    return 0 if result.failed_count == 0 else 2


def cli_emit(message: str) -> None:
    stream = getattr(sys, "stdout", None)
    if stream is None:
        return
    try:
        print(message)
    except (OSError, ValueError):
        return


def configure_cli_streams() -> None:
    for stream in (getattr(sys, "stdout", None), getattr(sys, "stderr", None)):
        if stream is None or not hasattr(stream, "reconfigure"):
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def main() -> None:
    multiprocessing.freeze_support()
    cli_exit_code = run_cli(sys.argv[1:])
    if cli_exit_code is not None:
        raise SystemExit(cli_exit_code)

    app = AuditorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
