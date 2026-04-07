from __future__ import annotations

import logging
import os
import queue
import sys
import threading
import traceback
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import (
    END, BOTH, DISABLED, LEFT, NORMAL, RIGHT, StringVar, Tk, Toplevel,
    filedialog, messagebox, ttk,
)
from tkinter.scrolledtext import ScrolledText

from ledgerflow.config import Config
from ledgerflow.desktop_runner import DesktopRunSummary, run_desktop_flow
from ledgerflow.models import Category, Transaction
from ledgerflow.seed_updater import NUMERIC_CATEGORY_MAP, update_seed_sheet
from ledgerflow.utils.normalization import normalize_merchant

LOGGER = logging.getLogger(__name__)
SUPPORTED_CURRENCY_CODES = ("USD", "EUR", "GBP")


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
        self.currency_rate_vars = {
            code: StringVar()
            for code in SUPPORTED_CURRENCY_CODES
        }
        self._loaded_config: Config | None = None

        self._event_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self._worker: threading.Thread | None = None

        self._build_ui()
        self._load_startup_config()
        self.root.after(100, self._process_events)

    def _build_ui(self) -> None:
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=BOTH, expand=True)

        frame = ttk.Frame(notebook, padding=16)
        notebook.add(frame, text="Run")
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(6, weight=1)
        frame.rowconfigure(8, weight=1)

        ttk.Label(frame, text="Bank transaction files").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Button(frame, text="Select files", command=self._select_bank_files).grid(row=0, column=2, sticky="e", padx=(8, 0), pady=(0, 8))
        self.bank_files_text = ScrolledText(frame, height=4, wrap="word")
        self.bank_files_text.grid(row=0, column=1, sticky="nsew", pady=(0, 8))
        self._configure_read_only_text(self.bank_files_text)

        ttk.Label(frame, text="Seed spreadsheet ID").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=self.seed_file_var).grid(row=1, column=1, columnspan=2, sticky="ew", pady=8)

        ttk.Label(frame, text="Year spreadsheet ID").grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=self.yearly_workbook_var).grid(row=2, column=1, columnspan=2, sticky="ew", pady=8)

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

        config_frame = ttk.Frame(notebook, padding=16)
        notebook.add(config_frame, text="Config")
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(
            config_frame,
            text="Manual currency rates: how many ILS one unit of currency is worth.",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        for row_idx, code in enumerate(SUPPORTED_CURRENCY_CODES, start=1):
            ttk.Label(config_frame, text=code).grid(row=row_idx, column=0, sticky="w", pady=6)
            ttk.Entry(config_frame, textvariable=self.currency_rate_vars[code]).grid(
                row=row_idx,
                column=1,
                sticky="ew",
                pady=6,
            )

        ttk.Button(config_frame, text="Save", command=self._save_currency_rates).grid(
            row=len(SUPPORTED_CURRENCY_CODES) + 1,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(12, 0),
        )

    def _select_bank_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select bank transaction files",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not paths:
            return
        self.bank_files = [Path(path) for path in paths]
        self._set_text(self.bank_files_text, "\n".join(str(path) for path in self.bank_files))

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
                seed_file=Path(self._resolved_seed_spreadsheet_id()),
                yearly_workbook=Path(self._resolved_year_spreadsheet_id()),
                month=int(self.month_var.get()),
                cache_file=Path(self.cache_file_var.get()) if self.cache_file_var.get().strip() else None,
                status_callback=lambda message: self._event_queue.put(("status", message)),
                base_config=self._loaded_config,
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
                        seed_spreadsheet_id = self._resolved_seed_spreadsheet_id()
                        self._open_gpt_review(summary.gpt_transactions, seed_spreadsheet_id)
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

    def _open_gpt_review(self, gpt_txs: list[Transaction], seed_spreadsheet_id: str) -> None:
        GptReviewDialog(self.root, gpt_txs, seed_spreadsheet_id)

    def _friendly_error_message(self, exc: Exception) -> str:
        if isinstance(exc, ValueError):
            return str(exc)
        if isinstance(exc, FileNotFoundError):
            return str(exc)
        return "The run failed. Check the selected files and try again."

    def _load_startup_config(self) -> None:
        loaded = load_default_config(DEFAULT_CONFIG_PATH)
        if loaded is None:
            return
        self._loaded_config = loaded
        if not self.seed_file_var.get().strip() and loaded.seed_spreadsheet_id:
            self.seed_file_var.set(loaded.seed_spreadsheet_id)
        if not self.yearly_workbook_var.get().strip() and loaded.year_spreadsheet_id:
            self.yearly_workbook_var.set(loaded.year_spreadsheet_id)
        self._load_currency_rate_vars(loaded)
        self._append_status(f"Loaded config from {DEFAULT_CONFIG_PATH.name}")

    def _load_currency_rate_vars(self, config: Config) -> None:
        for code, var in self.currency_rate_vars.items():
            rate = config.currency_rates.get(code)
            var.set("" if rate is None else format(rate, "f"))

    def _save_currency_rates(self) -> None:
        try:
            rates = validate_currency_rate_inputs(
                {code: var.get() for code, var in self.currency_rate_vars.items()}
            )
        except ValueError as exc:
            messagebox.showerror("LedgerFlow", str(exc))
            return

        config = self._loaded_config or Config()
        config.currency_rates = rates
        save_config(DEFAULT_CONFIG_PATH, config)
        self._loaded_config = config
        self._append_status(f"Saved currency rates to {DEFAULT_CONFIG_PATH.name}")
        messagebox.showinfo("LedgerFlow", "Currency rates saved.")

    def _resolved_seed_spreadsheet_id(self) -> str:
        fallback = self._loaded_config.seed_spreadsheet_id if self._loaded_config else None
        return resolve_spreadsheet_id(self.seed_file_var.get(), fallback)

    def _resolved_year_spreadsheet_id(self) -> str:
        fallback = self._loaded_config.year_spreadsheet_id if self._loaded_config else None
        return resolve_spreadsheet_id(self.yearly_workbook_var.get(), fallback)

    def _append_status(self, message: str) -> None:
        self.status_text.insert(END, f"{message}\n")
        self.status_text.see(END)

    def _set_text(self, widget: ScrolledText, value: str) -> None:
        widget.delete("1.0", END)
        widget.insert("1.0", value)

    def _set_button_state(self, running: bool) -> None:
        self.run_button.configure(state=DISABLED if running else NORMAL)

    def _configure_read_only_text(self, widget: ScrolledText) -> None:
        widget.configure(state=NORMAL)
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
        seed_spreadsheet_id: str,
    ) -> None:
        self._transactions = list(transactions)
        self._seed_spreadsheet_id = seed_spreadsheet_id
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
            seed_url = update_seed_sheet(
                self._seed_spreadsheet_id,
                os.getenv("LEDGERFLOW_SEED_WORKSHEET", "Seed"),
                reviewed,
            )
            messagebox.showinfo(
                "Seed updated",
                f"Wrote {len(reviewed)} merchant(s) to seed:\n{seed_url}",
                parent=self._win,
            )
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Seed update failed", str(exc), parent=self._win)
            return
        self._win.destroy()


def configure_logging() -> None:
    logging.basicConfig(
        filename=str(runtime_base_dir() / "ledgerflow_gui.log"),
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


def load_default_config(path: Path) -> Config | None:
    if not path.exists():
        return None
    return Config.from_json(path)


def resolve_spreadsheet_id(field_value: str, config_value: str | None) -> str:
    manual = field_value.strip()
    if manual:
        return manual
    return (config_value or "").strip()


def validate_currency_rate_inputs(values: dict[str, str]) -> dict[str, Decimal]:
    rates: dict[str, Decimal] = {}
    for code, raw_value in values.items():
        cleaned = raw_value.strip()
        if not cleaned:
            raise ValueError(f"Enter a positive numeric rate for {code}.")
        try:
            parsed = Decimal(cleaned)
        except InvalidOperation as exc:
            raise ValueError(f"Enter a positive numeric rate for {code}.") from exc
        if parsed <= 0:
            raise ValueError(f"Enter a positive numeric rate for {code}.")
        rates[code.upper()] = parsed
    return rates


def save_config(path: Path, config: Config) -> None:
    path.write_text(
        json_dumps_config(config),
        encoding="utf-8",
    )


def json_dumps_config(config: Config) -> str:
    import json

    return json.dumps(config.to_json_dict(), ensure_ascii=False, indent=2)


def runtime_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path.cwd()


DEFAULT_CONFIG_PATH = runtime_base_dir() / "config.json"
