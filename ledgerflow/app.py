from __future__ import annotations

import logging
import queue
import threading
import traceback
from pathlib import Path
from tkinter import (
    END, BOTH, DISABLED, LEFT, NORMAL, RIGHT, StringVar, Tk, Toplevel,
    filedialog, messagebox, ttk,
)
from tkinter.scrolledtext import ScrolledText

from ledgerflow.desktop_runner import DesktopRunSummary, run_desktop_flow
from ledgerflow.models import Category, Transaction
from ledgerflow.seed_updater import NUMERIC_CATEGORY_MAP, category_number, update_seed_file
from ledgerflow.utils.normalization import normalize_merchant

LOGGER = logging.getLogger(__name__)


class LedgerFlowApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("LedgerFlow")
        self.root.geometry("900x700")
        self.root.minsize(760, 620)

        self.bank_files: list[Path] = []
        self.seed_file_var = StringVar()
        self.yearly_workbook_var = StringVar()
        self.cache_file_var = StringVar()
        self.month_var = StringVar(value="1")

        self._event_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._worker: threading.Thread | None = None

        self._build_ui()
        self.root.after(100, self._process_events)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill=BOTH, expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(6, weight=1)
        frame.rowconfigure(8, weight=1)

        ttk.Label(frame, text="Bank transaction files").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Button(frame, text="Select files", command=self._select_bank_files).grid(row=0, column=2, sticky="e", padx=(8, 0), pady=(0, 8))
        self.bank_files_text = ScrolledText(frame, height=4, wrap="word")
        self.bank_files_text.grid(row=0, column=1, sticky="nsew", pady=(0, 8))
        self._configure_read_only_text(self.bank_files_text)

        ttk.Label(frame, text="Seed file").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=self.seed_file_var).grid(row=1, column=1, sticky="ew", pady=8)
        ttk.Button(frame, text="Browse", command=self._select_seed_file).grid(row=1, column=2, sticky="e", padx=(8, 0), pady=8)

        ttk.Label(frame, text="Yearly workbook").grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=self.yearly_workbook_var).grid(row=2, column=1, sticky="ew", pady=8)
        ttk.Button(frame, text="Browse", command=self._select_yearly_workbook).grid(row=2, column=2, sticky="e", padx=(8, 0), pady=8)

        ttk.Label(frame, text="Cache file (optional)").grid(row=3, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=self.cache_file_var).grid(row=3, column=1, sticky="ew", pady=8)
        ttk.Button(frame, text="Browse", command=self._select_cache_file).grid(row=3, column=2, sticky="e", padx=(8, 0), pady=8)

        ttk.Label(frame, text="Month").grid(row=4, column=0, sticky="w", pady=8)
        month_box = ttk.Combobox(
            frame,
            textvariable=self.month_var,
            values=[str(month) for month in range(1, 13)],
            state="readonly",
        )
        month_box.grid(row=4, column=1, sticky="w", pady=8)

        self.run_button = ttk.Button(frame, text="Run", command=self._start_run)
        self.run_button.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(12, 16))

        ttk.Label(frame, text="Status / progress").grid(row=6, column=0, columnspan=3, sticky="w")
        self.status_text = ScrolledText(frame, height=10, wrap="word")
        self.status_text.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(8, 16))
        self._configure_read_only_text(self.status_text)

        ttk.Label(frame, text="Final summary").grid(row=8, column=0, columnspan=3, sticky="w")
        self.summary_text = ScrolledText(frame, height=12, wrap="word")
        self.summary_text.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        self._configure_read_only_text(self.summary_text)

    def _select_bank_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select bank transaction files",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not paths:
            return
        self.bank_files = [Path(path) for path in paths]
        self._set_text(self.bank_files_text, "\n".join(str(path) for path in self.bank_files))

    def _select_seed_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select seed file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.seed_file_var.set(path)

    def _select_yearly_workbook(self) -> None:
        path = filedialog.askopenfilename(
            title="Select yearly workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.yearly_workbook_var.set(path)

    def _select_cache_file(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Select cache file",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if path:
            self.cache_file_var.set(path)

    def _start_run(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            return

        self._set_button_state(running=True)
        self._set_text(self.summary_text, "")
        self._append_status("Starting LedgerFlow run...")

        self._worker = threading.Thread(target=self._run_job, daemon=True)
        self._worker.start()

    def _run_job(self) -> None:
        try:
            summary = run_desktop_flow(
                bank_files=self.bank_files,
                seed_file=Path(self.seed_file_var.get()),
                yearly_workbook=Path(self.yearly_workbook_var.get()),
                month=int(self.month_var.get()),
                cache_file=Path(self.cache_file_var.get()) if self.cache_file_var.get().strip() else None,
                status_callback=lambda message: self._event_queue.put(("status", message)),
            )
            self._event_queue.put(("success", summary))
        except Exception as exc:  # noqa: BLE001
            LOGGER.exception("LedgerFlow desktop run failed.")
            self._event_queue.put(("error", self._friendly_error_message(exc)))
            self._event_queue.put(("debug_error", traceback.format_exc()))

    def _process_events(self) -> None:
        try:
            while True:
                event_type, payload = self._event_queue.get_nowait()
                if event_type == "status":
                    self._append_status(str(payload))
                elif event_type == "success":
                    summary: DesktopRunSummary = payload  # type: ignore[assignment]
                    self._render_summary(summary)
                    self._set_button_state(running=False)
                    if summary.gpt_transactions:
                        seed_path = Path(self.seed_file_var.get())
                        self._open_gpt_review(summary.gpt_transactions, seed_path)
                    else:
                        messagebox.showinfo("LedgerFlow", "Processing completed successfully.")
                elif event_type == "error":
                    self._append_status(str(payload))
                    self._set_button_state(running=False)
                    messagebox.showerror("LedgerFlow", str(payload))
                elif event_type == "debug_error":
                    self._append_status("Technical details were written to the log.")
                    LOGGER.debug("%s", payload)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._process_events)

    def _render_summary(self, summary: DesktopRunSummary) -> None:
        text = (
            f"Total extracted transactions: {summary.extracted_count}\n"
            f"Deduplicated count: {summary.deduplicated_count}\n"
            f"Filtered count: {summary.filtered_count}\n"
            f"Skipped due to invalid date: {summary.invalid_date_skips}\n"
            f"Seed hits: {summary.seed_hits}\n"
            f"Cache hits: {summary.cache_hits}\n"
            f"GPT hits: {summary.gpt_hits}\n"
            f"Fallback count: {summary.fallback_count}\n"
            f"Backup path: {summary.backup_path}\n"
            f"Raw CSV path: {summary.raw_csv_path}\n"
            f"Categorized CSV path: {summary.categorized_csv_path}\n"
            f"Final output workbook path: {summary.final_output_workbook_path}"
        )
        self._set_text(self.summary_text, text)

    def _open_gpt_review(self, gpt_txs: list[Transaction], seed_path: Path) -> None:
        GptReviewDialog(self.root, gpt_txs, seed_path)

    def _friendly_error_message(self, exc: Exception) -> str:
        if isinstance(exc, ValueError):
            return str(exc)
        if isinstance(exc, FileNotFoundError):
            return str(exc)
        return "The run failed. Check the selected files and try again."

    def _append_status(self, message: str) -> None:
        self.status_text.insert(END, f"{message}\n")
        self.status_text.see(END)

    def _set_text(self, widget: ScrolledText, value: str) -> None:
        widget.delete("1.0", END)
        widget.insert("1.0", value)

    def _set_button_state(self, running: bool) -> None:
        self.run_button.configure(state=DISABLED if running else NORMAL)

    def _configure_read_only_text(self, widget: ScrolledText) -> None:
        widget.bind("<Key>", self._block_text_edit)
        widget.bind("<Control-c>", self._copy_text)
        widget.bind("<Control-C>", self._copy_text)

    def _block_text_edit(self, event: object) -> str:
        return "break"

    def _copy_text(self, event: object) -> str:
        event.widget.event_generate("<<Copy>>")
        return "break"


class GptReviewDialog:
    """
    Compact review window for GPT-classified transactions.

    Shows each transaction with its current category assignment.  The user
    can confirm the GPT choice or replace it by pressing a number key (1-8).
    Pressing Enter or clicking Save writes all final decisions into the seed
    file so the seed grows over time.

    Key bindings:
        1 = Home          5 = Pharm
        2 = Leisure       6 = Business
        3 = Transportation  7 = Mortgage
        4 = Insurance     8 = Other
        Enter / Save      — commit decisions to seed file and close
        Escape            — close without saving
    """

    _CATEGORY_LABELS: str = (
        "1=Home  2=Leisure  3=Transport  4=Insurance  "
        "5=Pharm  6=Business  7=Mortgage  8=Other"
    )

    def __init__(
        self,
        parent: Tk,
        transactions: list[Transaction],
        seed_path: Path,
    ) -> None:
        self._transactions = list(transactions)
        self._seed_path = seed_path
        # decisions[i] = final Category for transactions[i]
        self._decisions: list[Category] = [
            tx.category or Category.OTHER for tx in transactions
        ]
        self._selected_index: int = 0

        self._win = Toplevel(parent)
        self._win.title(f"Review GPT results — {len(transactions)} transaction(s)")
        self._win.geometry("780x520")
        self._win.resizable(True, True)
        self._win.grab_set()  # modal

        self._build_ui()
        self._refresh_list()
        self._win.bind("<Escape>", lambda _e: self._win.destroy())
        self._win.bind("<Return>", lambda _e: self._save())
        for digit, cat in NUMERIC_CATEGORY_MAP.items():
            self._win.bind(str(digit), lambda _e, c=cat: self._apply_category(c))

    def _build_ui(self) -> None:
        frame = ttk.Frame(self._win, padding=12)
        frame.pack(fill=BOTH, expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        ttk.Label(frame, text=self._CATEGORY_LABELS, font=("TkFixedFont", 9)).grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        self._listbox_var = StringVar()
        self._listbox = ttk.Treeview(
            frame,
            columns=("merchant", "amount", "date", "category"),
            show="headings",
            selectmode="browse",
        )
        self._listbox.heading("merchant", text="Merchant")
        self._listbox.heading("amount", text="Amount")
        self._listbox.heading("date", text="Date")
        self._listbox.heading("category", text="Category (current)")
        self._listbox.column("merchant", width=260)
        self._listbox.column("amount", width=90, anchor="e")
        self._listbox.column("date", width=100)
        self._listbox.column("category", width=140)
        self._listbox.grid(row=1, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self._listbox.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self._listbox.configure(yscrollcommand=scrollbar.set)

        self._listbox.bind("<<TreeviewSelect>>", self._on_select)

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(btn_frame, text="Save to seed (Enter)", command=self._save).pack(side=RIGHT, padx=(8, 0))
        ttk.Button(btn_frame, text="Cancel (Esc)", command=self._win.destroy).pack(side=RIGHT)

    def _refresh_list(self) -> None:
        self._listbox.delete(*self._listbox.get_children())
        for i, (tx, cat) in enumerate(zip(self._transactions, self._decisions)):
            self._listbox.insert(
                "",
                "end",
                iid=str(i),
                values=(
                    tx.merchant,
                    f"{tx.amount:.2f}",
                    str(tx.date),
                    cat.value,
                ),
            )
        # Re-select current row
        if self._transactions:
            self._listbox.selection_set(str(self._selected_index))
            self._listbox.see(str(self._selected_index))

    def _on_select(self, _event: object) -> None:
        selected = self._listbox.selection()
        if selected:
            self._selected_index = int(selected[0])

    def _apply_category(self, cat: Category) -> None:
        self._decisions[self._selected_index] = cat
        # Advance to next row automatically for rapid review
        next_idx = min(self._selected_index + 1, len(self._transactions) - 1)
        self._selected_index = next_idx
        self._refresh_list()

    def _save(self) -> None:
        reviewed: dict[str, tuple[str, Category]] = {}
        for tx, cat in zip(self._transactions, self._decisions):
            norm = normalize_merchant(tx.merchant)
            if norm:
                reviewed[norm] = (tx.merchant, cat)
        try:
            update_seed_file(self._seed_path, reviewed)
            messagebox.showinfo(
                "Seed updated",
                f"Wrote {len(reviewed)} merchant(s) to seed:\n{self._seed_path}",
                parent=self._win,
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Seed update failed", str(exc), parent=self._win)
            return
        self._win.destroy()


def configure_logging() -> None:
    logging.basicConfig(
        filename="ledgerflow_gui.log",
        filemode="a",
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )


def main() -> None:
    configure_logging()
    root = Tk()
    ttk.Style(root).theme_use("clam")
    LedgerFlowApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
