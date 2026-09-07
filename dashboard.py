"""BSSE CommValid MIS — real-time pipeline dashboard."""

import queue
import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, scrolledtext, ttk
from typing import Optional

from config import INPUT_CSV, MAIN_XLSB, WORK_FILE
from mail_panel import MailPanel
from pipeline_runner import PIPELINE_STEPS, StepStatus, run_pipeline


STATUS_COLORS = {
    StepStatus.PENDING: ("#9e9e9e", "○"),
    StepStatus.RUNNING: ("#1976d2", "●"),
    StepStatus.SUCCESS: ("#2e7d32", "✓"),
    StepStatus.FAILED: ("#c62828", "✗"),
    StepStatus.SKIPPED: ("#f57c00", "—"),
}


class DashboardApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BSSE CommValid MIS Dashboard")
        self.geometry("1280x760")
        self.minsize(1024, 640)

        self.ui_queue: queue.Queue = queue.Queue()
        self.stop_event = threading.Event()
        self.pipeline_thread: Optional[threading.Thread] = None
        self.step_widgets: dict[str, dict] = {}
        self.is_running = False

        self._configure_style()
        self._build_ui()
        self._poll_queue()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Status.TLabel", font=("Segoe UI", 9))
        style.configure("Run.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Stop.TButton", font=("Segoe UI", 10))

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # Header
        header = ttk.Frame(main)
        header.pack(fill=tk.X, pady=(0, 12))

        ttk.Label(header, text="BSSE CommValid MIS Dashboard", style="Title.TLabel").pack(
            side=tk.LEFT
        )

        self.status_label = ttk.Label(header, text="Idle", foreground="#666666")
        self.status_label.pack(side=tk.RIGHT)

        # Body: left (pipeline) | center (logs) | right (mail)
        body = ttk.Panedwindow(main, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(body, padding=(0, 0, 8, 0))
        center = ttk.Frame(body, padding=8)
        right = ttk.Frame(body, padding=(8, 0, 0, 0))
        body.add(left, weight=1)
        body.add(center, weight=2)
        body.add(right, weight=1)

        self._build_pipeline_panel(left)
        self._build_log_panel(center)
        self._build_mail_section(right)

        # Footer controls
        footer = ttk.Frame(main)
        footer.pack(fill=tk.X, pady=(12, 0))

        self.run_btn = ttk.Button(
            footer, text="▶  Run Pipeline", style="Run.TButton", command=self.start_pipeline
        )
        self.run_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = ttk.Button(
            footer,
            text="■  Stop",
            style="Stop.TButton",
            command=self.stop_pipeline,
            state=tk.DISABLED,
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(footer, text="Clear Logs", command=self.clear_logs).pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(footer, mode="indeterminate")
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(16, 0))

    def _build_pipeline_panel(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Pipeline Steps", font=("Segoe UI", 11, "bold")).pack(
            anchor=tk.W, pady=(0, 8)
        )

        paths = ttk.LabelFrame(parent, text="Paths", padding=8)
        paths.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(paths, text=f"Input CSV:\n{INPUT_CSV}", wraplength=280, font=("Segoe UI", 8)).pack(
            anchor=tk.W
        )
        ttk.Label(paths, text=f"Work File:\n{WORK_FILE}", wraplength=280, font=("Segoe UI", 8)).pack(
            anchor=tk.W, pady=(4, 0)
        )
        ttk.Label(
            paths, text=f"Main XLSB:\n{MAIN_XLSB}", wraplength=280, font=("Segoe UI", 8)
        ).pack(anchor=tk.W, pady=(4, 0))

        steps_frame = ttk.LabelFrame(parent, text="Live Status", padding=8)
        steps_frame.pack(fill=tk.BOTH, expand=True)

        for step in PIPELINE_STEPS:
            row = ttk.Frame(steps_frame)
            row.pack(fill=tk.X, pady=4)

            icon = ttk.Label(row, text="○", width=2, font=("Segoe UI", 11))
            icon.pack(side=tk.LEFT)

            text_frame = ttk.Frame(row)
            text_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

            title = ttk.Label(text_frame, text=step.title, font=("Segoe UI", 9, "bold"))
            title.pack(anchor=tk.W)

            desc = ttk.Label(
                text_frame, text=step.description, font=("Segoe UI", 8), foreground="#666666"
            )
            desc.pack(anchor=tk.W)

            msg = ttk.Label(text_frame, text="Waiting...", style="Status.TLabel", foreground="#9e9e9e")
            msg.pack(anchor=tk.W)

            self.step_widgets[step.id] = {
                "icon": icon,
                "msg": msg,
                "status": StepStatus.PENDING,
            }

    def _build_log_panel(self, parent: ttk.Frame) -> None:
        ttk.Label(parent, text="Real-Time Log", font=("Segoe UI", 11, "bold")).pack(
            anchor=tk.W, pady=(0, 8)
        )

        self.log_text = scrolledtext.ScrolledText(
            parent,
            wrap=tk.WORD,
            font=("Consolas", 9),
            state=tk.DISABLED,
            background="#1e1e1e",
            foreground="#d4d4d4",
            insertbackground="#ffffff",
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_text.tag_configure("timestamp", foreground="#6a9955")
        self.log_text.tag_configure("error", foreground="#f44747")
        self.log_text.tag_configure("success", foreground="#4ec9b0")

    def _build_mail_section(self, parent: ttk.Frame) -> None:
        self.mail_panel = MailPanel(parent)
        self.mail_panel.pack(fill=tk.BOTH, expand=True)

    def _poll_queue(self) -> None:
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                kind = msg.get("kind")

                if kind == "log":
                    self._append_log(msg["text"])
                elif kind == "step":
                    self._update_step(msg["step_id"], msg["status"], msg["message"])
                elif kind == "done":
                    self._on_pipeline_done(msg["success"], msg.get("error"))
        except queue.Empty:
            pass

        self.after(100, self._poll_queue)

    def _append_log(self, text: str, tag: Optional[str] = None) -> None:
        self.log_text.configure(state=tk.NORMAL)
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] ", "timestamp")
        self.log_text.insert(tk.END, text + "\n", tag or "")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def clear_logs(self) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _update_step(self, step_id: str, status: StepStatus, message: str) -> None:
        widget = self.step_widgets.get(step_id)
        if not widget:
            return

        color, symbol = STATUS_COLORS[status]
        widget["icon"].configure(text=symbol, foreground=color)
        widget["msg"].configure(text=message, foreground=color)
        widget["status"] = status

    def _reset_steps(self) -> None:
        for step_id in self.step_widgets:
            self._update_step(step_id, StepStatus.PENDING, "Waiting...")

    def start_pipeline(self) -> None:
        if self.is_running:
            return

        self.is_running = True
        self.stop_event.clear()
        self._reset_steps()
        self.clear_logs()
        self._append_log("Starting pipeline...")

        self.run_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self.status_label.configure(text="Running...", foreground="#1976d2")
        self.progress.start(12)

        def on_log(text: str) -> None:
            self.ui_queue.put({"kind": "log", "text": text})

        def on_step(step_id: str, status: StepStatus, message: str) -> None:
            self.ui_queue.put(
                {
                    "kind": "step",
                    "step_id": step_id,
                    "status": status,
                    "message": message,
                }
            )

        def on_done(success: bool, error: Optional[str]) -> None:
            self.ui_queue.put({"kind": "done", "success": success, "error": error})

        def worker() -> None:
            run_pipeline(on_log, on_step, on_done, self.stop_event)

        self.pipeline_thread = threading.Thread(target=worker, daemon=True)
        self.pipeline_thread.start()

    def stop_pipeline(self) -> None:
        if not self.is_running:
            return
        self.stop_event.set()
        self._append_log("Stop requested — waiting for current step to finish...")
        self.status_label.configure(text="Stopping...", foreground="#f57c00")

    def _on_pipeline_done(self, success: bool, error: Optional[str]) -> None:
        self.is_running = False
        self.progress.stop()
        self.run_btn.configure(state=tk.NORMAL)
        self.stop_btn.configure(state=tk.DISABLED)

        if success:
            self.status_label.configure(text="Completed", foreground="#2e7d32")
            self._append_log("Pipeline finished successfully.", "success")
        else:
            self.status_label.configure(text="Failed / Stopped", foreground="#c62828")
            if error:
                self._append_log(f"Pipeline ended: {error}", "error")


def main() -> None:
    app = DashboardApp()
    app.mainloop()


if __name__ == "__main__":
    main()
