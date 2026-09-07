"""Mail draft window and sent-mail panel (Outlook Classic hooks are placeholders)."""

import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk
from typing import Callable, List, Optional

from config import GROUP_MAIL_ID, MAIL_SIGNATURE


class MailDraftWindow(tk.Toplevel):
    """Separate window for composing a mail draft before Outlook send."""

    def __init__(
        self,
        parent: tk.Misc,
        on_send: Optional[Callable[[dict], None]] = None,
    ):
        super().__init__(parent)
        self.on_send = on_send
        self.title("Mail Draft — Outlook Classic")
        self.geometry("720x620")
        self.minsize(600, 500)

        self._build_ui()
        self._load_default_draft()

    def _build_ui(self) -> None:
        container = ttk.Frame(self, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            container,
            text="Compose mail (will send from group mailbox via Outlook Classic)",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor=tk.W, pady=(0, 10))

        form = ttk.Frame(container)
        form.pack(fill=tk.BOTH, expand=True)

        self._add_field(form, "From (Group)", "from_var", readonly=True)
        self._add_field(form, "To", "to_var")
        self._add_field(form, "CC", "cc_var")
        self._add_field(form, "Subject", "subject_var")

        ttk.Label(form, text="Body").grid(row=4, column=0, sticky=tk.NW, pady=(8, 4))
        body_frame = ttk.Frame(form)
        body_frame.grid(row=4, column=1, sticky=tk.NSEW, pady=(8, 4))

        self.body_text = tk.Text(body_frame, height=12, wrap=tk.WORD, font=("Segoe UI", 10))
        body_scroll = ttk.Scrollbar(body_frame, orient=tk.VERTICAL, command=self.body_text.yview)
        self.body_text.configure(yscrollcommand=body_scroll.set)
        self.body_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        body_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Label(form, text="Signature (read-only)").grid(row=5, column=0, sticky=tk.NW, pady=(8, 4))
        sig_frame = ttk.Frame(form)
        sig_frame.grid(row=5, column=1, sticky=tk.EW, pady=(8, 4))

        self.signature_text = tk.Text(
            sig_frame,
            height=4,
            wrap=tk.WORD,
            font=("Segoe UI", 9),
            state=tk.DISABLED,
            background="#f5f5f5",
        )
        self.signature_text.pack(fill=tk.X)

        form.columnconfigure(1, weight=1)
        form.rowconfigure(4, weight=1)

        btn_row = ttk.Frame(container)
        btn_row.pack(fill=tk.X, pady=(12, 0))

        ttk.Button(btn_row, text="Reset Draft", command=self._load_default_draft).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(btn_row, text="Save Draft", command=self._save_draft).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(btn_row, text="Send via Outlook", command=self._send_mail).pack(
            side=tk.RIGHT
        )

        hint = ttk.Label(
            container,
            text="TODO: Wire Send via Outlook to Outlook Classic COM (group mailbox + signature).",
            foreground="#666666",
            font=("Segoe UI", 8),
        )
        hint.pack(anchor=tk.W, pady=(8, 0))

    def _add_field(
        self,
        parent: ttk.Frame,
        label: str,
        var_name: str,
        readonly: bool = False,
    ) -> None:
        row = len(parent.grid_slaves()) // 2
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=4)

        var = tk.StringVar()
        setattr(self, var_name, var)

        entry = ttk.Entry(parent, textvariable=var)
        if readonly:
            entry.configure(state="readonly")
        entry.grid(row=row, column=1, sticky=tk.EW, pady=4, padx=(8, 0))

    def _load_default_draft(self) -> None:
        self.from_var.set(GROUP_MAIL_ID)
        self.to_var.set("")
        self.cc_var.set("")
        self.subject_var.set("BSSE CommValid MIS Report")
        self.body_text.delete("1.0", tk.END)
        self.body_text.insert(
            "1.0",
            "Dear Team,\n\n"
            "Please find attached the latest BSSE CommValid MIS report.\n\n",
        )
        self.signature_text.configure(state=tk.NORMAL)
        self.signature_text.delete("1.0", tk.END)
        self.signature_text.insert("1.0", MAIL_SIGNATURE)
        self.signature_text.configure(state=tk.DISABLED)

    def get_draft_data(self) -> dict:
        return {
            "from": self.from_var.get(),
            "to": self.to_var.get(),
            "cc": self.cc_var.get(),
            "subject": self.subject_var.get(),
            "body": self.body_text.get("1.0", tk.END).strip(),
            "signature": MAIL_SIGNATURE,
            "full_body": (
                self.body_text.get("1.0", tk.END).strip()
                + "\n\n"
                + MAIL_SIGNATURE
            ),
        }

    def _save_draft(self) -> None:
        messagebox.showinfo("Draft Saved", "Mail draft saved locally (in-memory).")

    def _send_mail(self) -> None:
        draft = self.get_draft_data()
        if not draft["to"].strip():
            messagebox.showwarning("Validation", "Please enter at least one recipient in To.")
            return

        if self.on_send:
            self.on_send(draft)
        else:
            messagebox.showinfo(
                "Outlook Integration",
                "Hook your Outlook Classic send logic in mail_outlook.py\n"
                "and connect it via on_send callback.",
            )


class SentMailPanel(ttk.Frame):
    """List of sent mails — populated when you wire Outlook send."""

    def __init__(self, parent: tk.Misc):
        super().__init__(parent)
        self.sent_records: List[dict] = []
        self._build_ui()

    def _build_ui(self) -> None:
        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(header, text="Sent Mail", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT)
        ttk.Button(header, text="Refresh", command=self._refresh).pack(side=tk.RIGHT)

        columns = ("sent_at", "to", "subject", "status")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=10)
        self.tree.heading("sent_at", text="Sent At")
        self.tree.heading("to", text="To")
        self.tree.heading("subject", text="Subject")
        self.tree.heading("status", text="Status")
        self.tree.column("sent_at", width=140, anchor=tk.W)
        self.tree.column("to", width=180, anchor=tk.W)
        self.tree.column("subject", width=220, anchor=tk.W)
        self.tree.column("status", width=80, anchor=tk.CENTER)

        scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Label(
            self,
            text="Sent mails appear here after Outlook send is implemented.",
            foreground="#666666",
            font=("Segoe UI", 8),
        ).pack(anchor=tk.W, pady=(6, 0))

    def add_sent_record(self, draft: dict, status: str = "Sent") -> None:
        record = {
            "sent_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "to": draft.get("to", ""),
            "subject": draft.get("subject", ""),
            "status": status,
        }
        self.sent_records.insert(0, record)
        self._render()

    def _render(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for rec in self.sent_records:
            self.tree.insert(
                "",
                tk.END,
                values=(rec["sent_at"], rec["to"], rec["subject"], rec["status"]),
            )

    def _refresh(self) -> None:
        self._render()


class MailPanel(ttk.Frame):
    """Embedded mail section on the main dashboard."""

    def __init__(self, parent: tk.Misc):
        super().__init__(parent)
        self.draft_window: Optional[MailDraftWindow] = None
        self._build_ui()

    def _build_ui(self) -> None:
        ttk.Label(self, text="Mail", font=("Segoe UI", 12, "bold")).pack(anchor=tk.W)

        info = ttk.LabelFrame(self, text="Outlook Classic (Group Mailbox)", padding=10)
        info.pack(fill=tk.X, pady=(8, 8))
        ttk.Label(info, text=f"From: {GROUP_MAIL_ID}").pack(anchor=tk.W)
        ttk.Label(
            info,
            text="Signature and Outlook send logic — implement in mail_outlook.py",
            foreground="#666666",
        ).pack(anchor=tk.W, pady=(4, 0))

        btn_row = ttk.Frame(self)
        btn_row.pack(fill=tk.X, pady=(0, 8))

        ttk.Button(btn_row, text="Open Mail Draft", command=self.open_draft_window).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(btn_row, text="Quick Send (placeholder)", command=self._quick_send).pack(
            side=tk.LEFT
        )

        self.sent_panel = SentMailPanel(self)
        self.sent_panel.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

    def open_draft_window(self) -> None:
        if self.draft_window is not None and self.draft_window.winfo_exists():
            self.draft_window.lift()
            self.draft_window.focus_force()
            return

        self.draft_window = MailDraftWindow(self, on_send=self._handle_send)

    def _handle_send(self, draft: dict) -> None:
        # Placeholder — replace with Outlook Classic COM send
        self.sent_panel.add_sent_record(draft, status="Queued")
        messagebox.showinfo(
            "Mail Queued",
            f"Mail to '{draft['to']}' queued.\n"
            "Implement send_outlook_mail() in mail_outlook.py for real delivery.",
        )

    def _quick_send(self) -> None:
        self.open_draft_window()
