from __future__ import annotations

import argparse
import ctypes
from pathlib import Path
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable

try:
    from reit_excel_auditor.transformer import (
        AUTO_TYPE,
        TABLE_TYPES,
        BatchResult,
        ConversionError,
        convert_input_path,
    )
except ImportError:  # Direct script execution fallback.
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

        self._configure_style()
        self._build_layout()
        self.custom_template_path.trace_add("write", lambda *_: self.update_template_mode())
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

    def _build_layout(self) -> None:
        container = ttk.Frame(self, style="App.TFrame", padding=(26, 20))
        container.pack(fill="both", expand=True)

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

    def _add_source_row(self, parent: ttk.Frame) -> None:
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.pack(fill="x")
        ttk.Label(frame, text="输入来源", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(frame, text="可选择单个未审核 .xlsx，也可选择一个文件夹，程序会逐个识别其中的 Excel 文件。", style="Hint.TLabel").pack(anchor="w", pady=(2, 8))
        ttk.Label(frame, textvariable=self.folder_warning_text, style="Warning.TLabel").pack(anchor="w", pady=(0, 8))

        row = ttk.Frame(frame, style="Card.TFrame")
        row.pack(fill="x")
        ttk.Entry(row, textvariable=self.source_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="选择文件", style="Secondary.TButton", command=self.pick_input_file).pack(side="left", padx=(10, 0))
        ttk.Button(row, text="选择文件夹", style="Secondary.TButton", command=self.pick_input_folder).pack(side="left", padx=(8, 0))
        ttk.Button(row, text="清空", style="Secondary.TButton", command=self.clear_source_path).pack(side="left", padx=(8, 0))

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
        ttk.Label(frame, text=hint, style="Hint.TLabel").pack(anchor="w", pady=(2, 8))

        row = ttk.Frame(frame, style="Card.TFrame")
        row.pack(fill="x")
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text=file_button_text, style="Secondary.TButton", command=command).pack(side="left", padx=(10, 0))
        if clear_command:
            ttk.Button(row, text="清空", style="Secondary.TButton", command=clear_command).pack(side="left", padx=(8, 0))

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

    parser = argparse.ArgumentParser(description="REITs Excel 自动审核转换工具")
    parser.add_argument("--convert", help="要转换的未审核 Excel 文件或文件夹路径")
    parser.add_argument("--type", default=AUTO_TYPE, choices=list(TABLE_TYPES.keys()), help="转换类型，默认 auto")
    parser.add_argument("--metadata", default=None, help="可选补全信息表路径")
    parser.add_argument("--custom-template", default=None, help="可选自定义模板表路径；提供后将按模板字段和格式输出")
    parser.add_argument("--output-dir", default=None, help="可选输出目录")
    parser.add_argument("--property-processed", action="store_true", help="可选；第4类产权经营数据额外生成处理版，用于修复主配套资产面积字段")
    args = parser.parse_args(argv)

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
    print(f"summary={result.summary_file}")
    print(f"total={result.total_count}")
    print(f"success={result.success_count}")
    print(f"failed={result.failed_count}")
    for item in result.items:
        print(f"item={item.status}|{item.detected_type or ''}|{item.input_file}")
        if item.error:
            print(f"error={item.error}")
        for output_file in item.output_files:
            print(f"output={output_file}")
        for warning in item.warnings:
            print(f"warning={warning}")
    return 0 if result.failed_count == 0 else 2


def main() -> None:
    cli_exit_code = run_cli(sys.argv[1:])
    if cli_exit_code is not None:
        raise SystemExit(cli_exit_code)

    app = AuditorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
