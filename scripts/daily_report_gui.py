from __future__ import annotations

import contextlib
import io
import queue
import shutil
import sys
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


WORKSPACE = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[1]
INPUT_ROOT = WORKSPACE / "input"
OUTPUT_ROOT = WORKSPACE / "output"
DEFAULT_TEMPLATE = INPUT_ROOT / "日报模版.docx"


class DailyReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("日报生成工具")
        self.geometry("760x500")
        self.minsize(680, 460)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None

        self.ledger_var = tk.StringVar(value=str(self.find_newest_ledger() or ""))
        self.template_var = tk.StringVar(value=str(DEFAULT_TEMPLATE if DEFAULT_TEMPLATE.exists() else ""))

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
        ttk.Label(root, text="选择台账和日报模板后，点击生成即可。", style="Hint.TLabel").grid(
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
            label="模板",
            variable=self.template_var,
            button_text="选择模板",
            command=self.choose_template,
        )

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
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=8)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky=tk.EW, padx=(12, 8), pady=8)
        ttk.Button(parent, text=button_text, command=command).grid(row=row, column=2, sticky=tk.EW, pady=8)

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

    def start_generation(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("正在生成", "日报正在生成中，请等待完成。")
            return

        ledger = Path(self.ledger_var.get().strip().strip('"'))
        template = Path(self.template_var.get().strip().strip('"'))
        if not ledger.exists():
            messagebox.showerror("文件不存在", f"台账不存在:\n{ledger}")
            return
        if not template.exists():
            messagebox.showerror("文件不存在", f"模板不存在:\n{template}")
            return

        self.generate_button.configure(state=tk.DISABLED, text="正在生成...")
        self.log_text.delete("1.0", tk.END)
        self.append_log(f"[run] 台账: {ledger}")
        self.append_log(f"[run] 模板: {template}")
        self.worker = threading.Thread(target=self.run_generation, args=(ledger, template), daemon=True)
        self.worker.start()

    def run_generation(self, ledger: Path, template: Path) -> None:
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
                return_code = generate_daily_reports.generate_reports(
                    source=ledger,
                    overwrite=True,
                    template=template,
                    output_dir=OUTPUT_ROOT,
                )
            writer.flush()
            self.log_queue.put(f"[done] exit code: {return_code}")
            if return_code == 0:
                self.log_queue.put(f"[ok] 输出目录: {OUTPUT_ROOT}")
        except Exception as exc:
            self.log_queue.put(f"[fail] {exc}")
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
