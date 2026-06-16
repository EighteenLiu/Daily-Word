from __future__ import annotations

import contextlib
import io
import queue
import shutil
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


WORKSPACE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
INPUT_ROOT = WORKSPACE / "input"
OUTPUT_ROOT = WORKSPACE / "output"


def first_existing_input_file(*names: str) -> Path:
    for name in names:
        path = INPUT_ROOT / name
        if path.exists():
            return path
    return INPUT_ROOT / names[0]


DEFAULT_TEMPLATE = first_existing_input_file("daily_report_jinja_template.docx", "街道日报模板.docx", "日报模版.docx")
DEFAULT_GARBAGE_SUMMARY_TEMPLATE = first_existing_input_file("垃圾分类工作日报_jinja模板.docx", "5月17日垃圾分类工作日报.docx")
DEFAULT_DAILY_SUMMARY_TEMPLATE = first_existing_input_file("每日汇总情况_jinja模板.docx", "每日汇总情况(10).docx")


class DailyReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("日报生成工具")
        self.geometry("900x680")
        self.minsize(800, 620)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None

        self.ledger_var = tk.StringVar(value=str(self.find_newest_ledger() or ""))
        self.template_var = tk.StringVar(value=str(DEFAULT_TEMPLATE if DEFAULT_TEMPLATE.exists() else ""))
        self.summary_enabled_var = tk.BooleanVar(
            value=DEFAULT_GARBAGE_SUMMARY_TEMPLATE.exists() or DEFAULT_DAILY_SUMMARY_TEMPLATE.exists()
        )
        self.summary_from_existing_var = tk.BooleanVar(value=False)
        self.existing_reports_dir_var = tk.StringVar(value="")
        self.garbage_summary_template_var = tk.StringVar(
            value=str(DEFAULT_GARBAGE_SUMMARY_TEMPLATE if DEFAULT_GARBAGE_SUMMARY_TEMPLATE.exists() else "")
        )
        self.daily_summary_template_var = tk.StringVar(
            value=str(DEFAULT_DAILY_SUMMARY_TEMPLATE if DEFAULT_DAILY_SUMMARY_TEMPLATE.exists() else "")
        )
        self.outside_bucket_report_var = tk.StringVar(value="")
        self.compression_var = tk.StringVar(value="standard")

        self.configure_style()
        self.build_ui()
        self.after(100, self.drain_log_queue)

    def configure_style(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Hint.TLabel", foreground="#64748b")
        style.configure("Primary.TButton", padding=(18, 8))

    def build_ui(self) -> None:
        root = ttk.Frame(self, padding=20)
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)

        ttk.Label(root, text="日报生成工具", style="Title.TLabel").grid(row=0, column=0, sticky=tk.W)
        ttk.Label(root, text="选择台账和模板后，可同时生成街道日报、每日汇总和垃圾分类工作日报。", style="Hint.TLabel").grid(
            row=1,
            column=0,
            sticky=tk.W,
            pady=(4, 18),
        )

        form = ttk.Frame(root)
        form.grid(row=2, column=0, sticky=tk.EW)
        form.columnconfigure(1, weight=1)

        self.add_file_row(
            form,
            row=0,
            label="台账",
            variable=self.ledger_var,
            button_text="选择台账",
            command=self.choose_ledger,
        )
        self.add_file_row(
            form,
            row=1,
            label="桶外摆日报",
            variable=self.outside_bucket_report_var,
            button_text="选择文件",
            command=self.choose_outside_bucket_report,
        )
        ttk.Label(form, text="可选，用于把桶外摆内容写入对应街道日报", style="Hint.TLabel").grid(
            row=2,
            column=1,
            sticky=tk.W,
            padx=(12, 8),
        )
        self.add_file_row(
            form,
            row=3,
            label="街道日报模板",
            variable=self.template_var,
            button_text="选择模板",
            command=self.choose_template,
        )
        self.add_summary_toggle_row(form, row=4)
        self.existing_reports_widgets = self.add_existing_reports_row(form, row=5)
        self.garbage_summary_widgets = self.add_file_row(
            form,
            row=6,
            label="垃圾分类日报模板",
            variable=self.garbage_summary_template_var,
            button_text="选择模板",
            command=self.choose_garbage_summary_template,
        )
        self.daily_summary_widgets = self.add_file_row(
            form,
            row=7,
            label="每日汇总模板",
            variable=self.daily_summary_template_var,
            button_text="选择模板",
            command=self.choose_daily_summary_template,
        )
        self.add_compression_row(form, row=8)
        self.toggle_summary_rows()

        actions = ttk.Frame(root)
        actions.grid(row=3, column=0, sticky=tk.NSEW, pady=(18, 0))
        actions.columnconfigure(0, weight=1)
        actions.rowconfigure(2, weight=1)

        self.generate_button = ttk.Button(
            actions,
            text="生成日报",
            style="Primary.TButton",
            command=self.start_generation,
        )
        self.generate_button.grid(row=0, column=0, sticky=tk.W)

        ttk.Label(actions, text="日志", style="Hint.TLabel").grid(row=1, column=0, sticky=tk.W, pady=(16, 6))
        log_frame = ttk.Frame(actions)
        log_frame.grid(row=2, column=0, sticky=tk.NSEW)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_frame,
            height=12,
            wrap=tk.WORD,
            relief=tk.SOLID,
            borderwidth=1,
            font=("Consolas", 10),
        )
        self.log_text.grid(row=0, column=0, sticky=tk.NSEW)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def add_file_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        button_text: str,
        command,
    ):
        label_widget = ttk.Label(parent, text=label)
        entry = ttk.Entry(parent, textvariable=variable)
        button = ttk.Button(parent, text=button_text, command=command)
        label_widget.grid(row=row, column=0, sticky=tk.W, pady=8)
        entry.grid(row=row, column=1, sticky=tk.EW, padx=(12, 8), pady=8)
        button.grid(row=row, column=2, sticky=tk.EW, pady=8)
        return (label_widget, entry, button)

    def add_summary_toggle_row(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="汇总报告").grid(row=row, column=0, sticky=tk.W, pady=8)
        options = ttk.Frame(parent)
        options.grid(row=row, column=1, sticky=tk.W, padx=(12, 8), pady=8)
        ttk.Checkbutton(
            options,
            text="同时生成每日汇总和垃圾分类工作日报",
            variable=self.summary_enabled_var,
            command=self.toggle_summary_rows,
        ).grid(row=0, column=0, sticky=tk.W)
        ttk.Checkbutton(
            options,
            text="基于已有日报生成",
            variable=self.summary_from_existing_var,
            command=self.toggle_summary_rows,
        ).grid(row=0, column=1, sticky=tk.W, padx=(18, 0))
        ttk.Label(parent, text="不需要时可取消勾选", style="Hint.TLabel").grid(row=row, column=2, sticky=tk.W, pady=8)

    def add_existing_reports_row(self, parent: ttk.Frame, row: int):
        label = ttk.Label(parent, text="已有日报文件夹")
        entry = ttk.Entry(parent, textvariable=self.existing_reports_dir_var)
        button = ttk.Button(parent, text="选择文件夹", command=self.choose_existing_reports_dir)
        label.grid(row=row, column=0, sticky=tk.W, pady=8)
        entry.grid(row=row, column=1, sticky=tk.EW, padx=(12, 8), pady=8)
        button.grid(row=row, column=2, sticky=tk.EW, pady=8)
        return (label, entry, button)

    def toggle_summary_rows(self) -> None:
        summary_visible = self.summary_enabled_var.get()
        existing_visible = summary_visible and self.summary_from_existing_var.get()
        for widget in getattr(self, "garbage_summary_widgets", ()):
            if summary_visible:
                widget.grid()
            else:
                widget.grid_remove()
        for widget in getattr(self, "daily_summary_widgets", ()):
            if summary_visible:
                widget.grid()
            else:
                widget.grid_remove()
        for widget in getattr(self, "existing_reports_widgets", ()):
            if existing_visible:
                widget.grid()
            else:
                widget.grid_remove()

    def add_compression_row(self, parent: ttk.Frame, row: int) -> None:
        ttk.Label(parent, text="图片压缩").grid(row=row, column=0, sticky=tk.W, pady=8)
        options = ttk.Frame(parent)
        options.grid(row=row, column=1, sticky=tk.W, padx=(12, 8), pady=8)
        for index, (label, value) in enumerate(
            (
                ("标准", "standard"),
                ("轻度", "light"),
                ("强力", "strong"),
                ("不压缩", "none"),
            )
        ):
            ttk.Radiobutton(options, text=label, value=value, variable=self.compression_var).grid(
                row=0,
                column=index,
                padx=(0, 10),
                sticky=tk.W,
            )
        ttk.Label(
            parent,
            text="标准: 推荐；轻度: 更清晰；强力: 更小；none: 不压缩",
            style="Hint.TLabel",
        ).grid(row=row, column=2, sticky=tk.W, pady=8)

    def find_newest_ledger(self) -> Path | None:
        INPUT_ROOT.mkdir(parents=True, exist_ok=True)
        candidates = [
            path
            for suffix in ("*.xls", "*.xlsx")
            for path in INPUT_ROOT.glob(suffix)
            if not path.name.startswith("~$")
        ]
        return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None

    def choose_ledger(self) -> None:
        path = filedialog.askopenfilename(
            title="选择台账",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.ledger_var.set(str(Path(path)))

    def choose_template(self) -> None:
        path = filedialog.askopenfilename(
            title="选择日报模板",
            filetypes=[("Word templates", "*.docx"), ("All files", "*.*")],
        )
        if not path:
            return
        source = Path(path)
        self.template_var.set(str(source))
        self.append_log(f"[ok] 模板已选择: {source}")

    def choose_garbage_summary_template(self) -> None:
        path = filedialog.askopenfilename(
            title="选择垃圾分类工作日报模板",
            filetypes=[("Word templates", "*.docx"), ("All files", "*.*")],
        )
        if not path:
            return
        source = Path(path)
        self.garbage_summary_template_var.set(str(source))
        self.summary_enabled_var.set(True)
        self.toggle_summary_rows()
        self.append_log(f"[ok] 垃圾分类工作日报模板已选择: {source}")

    def choose_daily_summary_template(self) -> None:
        path = filedialog.askopenfilename(
            title="选择每日汇总模板",
            filetypes=[("Word templates", "*.docx"), ("All files", "*.*")],
        )
        if not path:
            return
        source = Path(path)
        self.daily_summary_template_var.set(str(source))
        self.summary_enabled_var.set(True)
        self.toggle_summary_rows()
        self.append_log(f"[ok] 每日汇总模板已选择: {source}")

    def choose_existing_reports_dir(self) -> None:
        path = filedialog.askdirectory(title="选择已有街道日报文件夹")
        if not path:
            return
        source = Path(path)
        self.existing_reports_dir_var.set(str(source))
        self.summary_enabled_var.set(True)
        self.summary_from_existing_var.set(True)
        self.toggle_summary_rows()
        self.append_log(f"[ok] 已有日报文件夹已选择: {source}")

    def choose_outside_bucket_report(self) -> None:
        path = filedialog.askopenfilename(
            title="选择桶外摆日报",
            filetypes=[("Documents", "*.docx *.doc *.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return
        source = Path(path)
        self.outside_bucket_report_var.set(str(source))
        self.append_log(f"[ok] 桶外摆日报已选择: {source}")

    def optional_path_from_var(self, variable: tk.StringVar) -> Path | None:
        value = variable.get().strip().strip('"')
        return Path(value) if value else None

    def start_generation(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("正在生成", "日报正在生成中，请等待完成。")
            return

        ledger = Path(self.ledger_var.get().strip().strip('"'))
        template = Path(self.template_var.get().strip().strip('"'))
        garbage_summary_template = self.optional_path_from_var(self.garbage_summary_template_var)
        daily_summary_template = self.optional_path_from_var(self.daily_summary_template_var)
        outside_bucket_report = self.optional_path_from_var(self.outside_bucket_report_var)
        existing_reports_dir = self.optional_path_from_var(self.existing_reports_dir_var)
        from_existing_reports = self.summary_enabled_var.get() and self.summary_from_existing_var.get()
        if not from_existing_reports:
            if not ledger.exists():
                messagebox.showerror("文件不存在", f"台账不存在:\n{ledger}")
                return
            if not template.exists():
                messagebox.showerror("文件不存在", f"模板不存在:\n{template}")
                return
        if self.summary_enabled_var.get():
            if not garbage_summary_template and not daily_summary_template:
                messagebox.showerror("汇总模板缺失", "请至少选择一个汇总模板，或取消勾选“同时生成每日汇总和垃圾分类工作日报”。")
                return
            if from_existing_reports and (not existing_reports_dir or not existing_reports_dir.exists() or not existing_reports_dir.is_dir()):
                messagebox.showerror("文件夹不存在", f"已有日报文件夹不存在:\n{existing_reports_dir or ''}")
                return
            for label, path in (
                ("垃圾分类工作日报模板", garbage_summary_template),
                ("每日汇总模板", daily_summary_template),
            ):
                if path and not path.exists():
                    messagebox.showerror("文件不存在", f"{label}不存在:\n{path}")
                    return
        else:
            garbage_summary_template = None
            daily_summary_template = None
        if outside_bucket_report and not outside_bucket_report.exists():
            messagebox.showerror("文件不存在", f"桶外摆日报不存在:\n{outside_bucket_report}")
            return

        self.generate_button.configure(state=tk.DISABLED, text="正在生成...")
        self.log_text.delete("1.0", tk.END)
        if from_existing_reports:
            self.append_log(f"[run] 基于已有日报文件夹: {existing_reports_dir}")
        else:
            self.append_log(f"[run] 台账: {ledger}")
            self.append_log(f"[run] 街道日报模板: {template}")
        if garbage_summary_template:
            self.append_log(f"[run] 垃圾分类工作日报模板: {garbage_summary_template}")
        if daily_summary_template:
            self.append_log(f"[run] 每日汇总模板: {daily_summary_template}")
        if outside_bucket_report and not from_existing_reports:
            self.append_log(f"[run] 桶外摆日报: {outside_bucket_report}")
        self.append_log(f"[run] 图片压缩: {self.compression_var.get()}")
        self.worker = threading.Thread(
            target=self.run_generation,
            args=(
                ledger,
                template,
                self.compression_var.get(),
                garbage_summary_template,
                daily_summary_template,
                outside_bucket_report,
                existing_reports_dir if from_existing_reports else None,
            ),
            daemon=True,
        )
        self.worker.start()

    def run_generation(
        self,
        ledger: Path,
        template: Path,
        image_compression: str,
        garbage_summary_template: Path | None,
        daily_summary_template: Path | None,
        outside_bucket_report: Path | None,
        existing_reports_dir: Path | None,
    ) -> None:
        try:
            import generate_daily_reports

            class QueueWriter(io.TextIOBase):
                def __init__(self, log_queue: queue.Queue[str]) -> None:
                    self.log_queue = log_queue
                    self.buffer = ""

                def write(self, text: str) -> int:
                    self.buffer += text
                    while "\n" in self.buffer:
                        line, self.buffer = self.buffer.split("\n", 1)
                        if line:
                            self.log_queue.put(line.rstrip())
                    return len(text)

                def flush(self) -> None:
                    if self.buffer:
                        self.log_queue.put(self.buffer.rstrip())
                        self.buffer = ""

            writer = QueueWriter(self.log_queue)
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                if existing_reports_dir:
                    import generate_daily_summaries

                    generate_daily_summaries.generate_summaries_from_existing_reports(
                        report_dir=existing_reports_dir,
                        garbage_template=garbage_summary_template,
                        daily_template=daily_summary_template,
                        output_dir=None,
                    )
                    return_code = 0
                else:
                    return_code = generate_daily_reports.generate_reports(
                        source=ledger,
                        overwrite=True,
                        template=template,
                        output_dir=OUTPUT_ROOT,
                        garbage_summary_template=garbage_summary_template,
                        daily_summary_template=daily_summary_template,
                        outside_bucket_path=outside_bucket_report,
                        image_compression=image_compression,
                    )
            writer.flush()
            self.log_queue.put(f"[done] exit code: {return_code}")
            if return_code == 0:
                self.log_queue.put(f"[ok] 输出目录: {OUTPUT_ROOT}")
        except Exception as exc:
            self.log_queue.put(f"[fail] {type(exc).__name__}: {exc}")
            for line in traceback.format_exception(type(exc), exc, exc.__traceback__):
                for part in line.rstrip().splitlines():
                    self.log_queue.put(part)
        finally:
            self.log_queue.put("__ENABLE_BUTTON__")

    def drain_log_queue(self) -> None:
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            if message == "__ENABLE_BUTTON__":
                self.generate_button.configure(state=tk.NORMAL, text="生成日报")
            else:
                self.append_log(message)
        self.after(100, self.drain_log_queue)

    def append_log(self, message: str) -> None:
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)


def main() -> int:
    app = DailyReportApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
