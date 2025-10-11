"""Tkinter UI for the Spirit Lookup application."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import List

from .config import AppConfig
from .controller import LookupResult, SpiritLookupController
from .excel_helper import open_excel_helper
from .mail import MailClientError, open_mail_client
from .models import SpiritRecord
from .providers import DataProviderError, RecordNotFoundError


class SpiritLookupApp:
    """Tkinter based UI application."""

    def __init__(self, root: tk.Tk, controller: SpiritLookupController, config: AppConfig) -> None:
        self.root = root
        self.controller = controller
        self.config = config
        self.helper_config_path = config.fixture_path.parent / "excel_helper_config.json"

        self.current_query: str = ""
        self.current_page: int = 0
        self.cached_records: List[SpiritRecord] = []
        self.has_more: bool = False
        self._debounce_id: str | None = None

        self.status_var = tk.StringVar(value="Bereit.")
        self.search_var = tk.StringVar()
        self.spirit_entry_var = tk.StringVar()

        self._build_ui()
        self._load_initial()

    def _build_ui(self) -> None:
        self.root.title("Spirit Lookup")
        self.root.geometry("720x380")

        main_frame = ttk.Frame(self.root, padding="16")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        ttk.Label(main_frame, text="Spirit Code eingeben:").grid(row=0, column=0, sticky="w")
        self.spirit_entry = ttk.Entry(main_frame, textvariable=self.spirit_entry_var, width=30)
        self.spirit_entry.grid(row=1, column=0, sticky="we", padx=(0, 12))
        self.spirit_entry.bind("<Return>", lambda _event: self.on_search())

        ttk.Label(main_frame, text="oder Hotel auswählen:").grid(row=0, column=1, sticky="w")
        self.search_combo = ttk.Combobox(main_frame, textvariable=self.search_var, width=50)
        self.search_combo.grid(row=1, column=1, sticky="we")
        self.search_combo.bind("<KeyRelease>", self._on_search_var_change)
        self.search_combo.bind("<<ComboboxSelected>>", lambda _event: self.on_search())

        self.helper_button = ttk.Button(
            main_frame,
            text="Excel Helper",
            command=self._open_excel_helper,
        )
        self.helper_button.grid(row=2, column=0, sticky="w", pady=(8, 0))

        self.load_more_button = ttk.Button(
            main_frame,
            text="Weitere Ergebnisse laden",
            command=self._load_next_page,
        )
        self.load_more_button.grid(row=2, column=1, sticky="e", pady=(8, 0))
        self.load_more_button.grid_remove()

        search_button = ttk.Button(main_frame, text="Suchen", command=self.on_search)
        search_button.grid(row=1, column=2, padx=(12, 0))

        self.status_label = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor="w")
        self.status_label.grid(row=3, column=0, columnspan=3, sticky="we", pady=(16, 0))

        for col in range(3):
            main_frame.columnconfigure(col, weight=1 if col < 2 else 0)

    def _update_status(self, text: str) -> None:
        self.status_var.set(text)

    def _open_excel_helper(self) -> None:
        open_excel_helper(self.root, self.helper_config_path)

    def _load_initial(self) -> None:
        try:
            self._update_status("Lade Daten …")
            result = self.controller.list_records(page=0)
            self._apply_lookup_result(result)
            if not result.records:
                self._update_status("Keine Daten verfügbar. Nutzen Sie ggf. den Fixture-Modus.")
            else:
                self._update_status(f"{len(self.cached_records)} Hotels geladen.")
        except DataProviderError as exc:
            messagebox.showerror("Fehler beim Laden", str(exc))
            self._update_status("Laden fehlgeschlagen – bitte erneut versuchen.")

    def _apply_lookup_result(self, result: LookupResult) -> None:
        self.current_page = result.page
        self.has_more = result.has_more
        if result.page == 0:
            self.cached_records = list(result.records)
        else:
            known_codes = {record.spirit_code for record in self.cached_records}
            for record in result.records:
                if record.spirit_code not in known_codes:
                    self.cached_records.append(record)
                    known_codes.add(record.spirit_code)

        values = [record.display_label() for record in self.cached_records]
        if not values:
            self.search_combo["values"] = [""]
        else:
            self.search_combo["values"] = values
        if self.has_more:
            self.load_more_button.grid()
            self.load_more_button.configure(state=tk.NORMAL)
        else:
            self.load_more_button.grid_remove()

    def _schedule_search(self, query: str) -> None:
        if self._debounce_id:
            self.root.after_cancel(self._debounce_id)
        self._debounce_id = self.root.after(self.config.debounce_ms, lambda: self._perform_search(query))

    def _on_search_var_change(self, _event) -> None:
        query = self.search_var.get().strip()
        self.current_query = query
        self._schedule_search(query)

    def _perform_search(self, query: str, page: int = 0) -> None:
        try:
            self._update_status("Suche …")
            result = self.controller.list_records(query, page=page)
            self._apply_lookup_result(result)
            if not result.records:
                self._update_status("Keine Treffer – versuchen Sie eine andere Suche.")
            else:
                self._update_status(
                    f"{len(result.records)} Treffer auf Seite {result.page + 1}."
                    + (" Weitere Seiten verfügbar." if result.has_more else "")
                )
        except DataProviderError as exc:
            retry = messagebox.askretrycancel("Fehler", f"{exc}\nErneut versuchen?")
            if retry:
                self._perform_search(query, page=page)
            else:
                self._update_status("Suche abgebrochen.")

    def _load_next_page(self) -> None:
        if not self.has_more:
            return
        next_page = self.current_page + 1
        self._perform_search(self.current_query, page=next_page)

    def on_search(self) -> None:
        try:
            record = self.controller.search_by_input(
                spirit_code=self.spirit_entry_var.get().strip() or None,
                selected_label=self.search_var.get().strip() or None,
                cached_records=self.cached_records,
            )
            self._show_details(record)
        except RecordNotFoundError as exc:
            messagebox.showinfo("Keine Auswahl", str(exc))
        except DataProviderError as exc:
            retry = messagebox.askretrycancel("Fehler", f"{exc}\nErneut versuchen?")
            if retry:
                self.on_search()
        finally:
            self.spirit_entry.selection_clear()

    def _show_details(self, record: SpiritRecord) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title(f"{record.hotel_name} ({record.spirit_code})")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("640x480")
        dialog.resizable(True, True)

        dialog.columnconfigure(0, weight=1)
        dialog.rowconfigure(1, weight=1)

        ttk.Label(dialog, text="Schlüssel-Informationen", font=("Segoe UI", 12, "bold")).grid(
            row=0, column=0, sticky="nw", padx=16, pady=(16, 4)
        )

        info_frame = ttk.Frame(dialog, padding="16")
        info_frame.grid(row=1, column=0, sticky="nsew")
        info_frame.columnconfigure(1, weight=1)

        entries = [
            ("Spirit Code", record.spirit_code),
            ("Hotel", record.hotel_name),
            ("Region", record.region or "–"),
            ("Status", record.status or "–"),
            ("Ort", ", ".join(filter(None, [record.location_city, record.location_country])) or "–"),
            ("Adresse", record.address or "–"),
        ]
        for idx, (label, value) in enumerate(entries):
            ttk.Label(info_frame, text=f"{label}:", font=("Segoe UI", 10, "bold"), anchor="w").grid(
                row=idx, column=0, sticky="w", pady=2
            )
            ttk.Label(info_frame, text=value, anchor="w", wraplength=360).grid(
                row=idx, column=1, sticky="w", pady=2
            )

        contact_frame = ttk.LabelFrame(dialog, text="Kontaktinformationen", padding="16")
        contact_frame.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 16))
        contact_frame.columnconfigure(1, weight=1)

        if record.contacts:
            for idx, contact in enumerate(record.contacts):
                ttk.Label(contact_frame, text=contact.role or "Kontakt", font=("Segoe UI", 10, "bold")).grid(
                    row=idx * 2, column=0, sticky="nw"
                )
                ttk.Label(contact_frame, text=contact.name or "–").grid(row=idx * 2, column=1, sticky="w")
                buttons_frame = ttk.Frame(contact_frame)
                buttons_frame.grid(row=idx * 2 + 1, column=0, columnspan=2, sticky="w", pady=(0, 8))
                if contact.email:
                    ttk.Button(
                        buttons_frame,
                        text=f"E-Mail kopieren ({contact.email})",
                        command=lambda value=contact.email: self._copy_to_clipboard(value),
                    ).pack(side="left", padx=(0, 8))
                if contact.phone:
                    ttk.Button(
                        buttons_frame,
                        text=f"Telefon kopieren ({contact.phone})",
                        command=lambda value=contact.phone: self._copy_to_clipboard(value),
                    ).pack(side="left", padx=(0, 8))
        else:
            ttk.Label(contact_frame, text="Keine Kontakte vorhanden.").grid(row=0, column=0, sticky="w")

        action_frame = ttk.Frame(dialog, padding="16")
        action_frame.grid(row=3, column=0, sticky="ew")
        action_frame.columnconfigure(2, weight=1)

        confirm_var = tk.BooleanVar(value=False)
        confirm_check = ttk.Checkbutton(
            action_frame,
            text="Hinweise gelesen / Datenschutz akzeptiert",
            variable=confirm_var,
            command=lambda: draft_button.configure(state=(tk.NORMAL if confirm_var.get() else tk.DISABLED)),
        )
        confirm_check.grid(row=0, column=0, sticky="w")

        draft_button = ttk.Button(
            action_frame,
            text="Draft E-Mail",
            command=lambda: self._open_draft_email(dialog),
            state=tk.NORMAL if (self.config.draft_email_enabled and confirm_var.get()) else tk.DISABLED,
        )
        draft_button.grid(row=0, column=1, padx=(12, 0))

        if not self.config.draft_email_enabled:
            draft_button.configure(state=tk.DISABLED)

        close_button = ttk.Button(action_frame, text="Schließen", command=lambda: self._close_dialog(dialog))
        close_button.grid(row=0, column=2, sticky="e")

        def on_confirm_change(*_args):
            if self.config.draft_email_enabled:
                draft_button.configure(state=tk.NORMAL if confirm_var.get() else tk.DISABLED)

        confirm_var.trace_add("write", on_confirm_change)

        dialog.bind("<Escape>", lambda _event: self._close_dialog(dialog))
        dialog.focus_set()
        close_button.focus_set()

    def _copy_to_clipboard(self, value: str) -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self._update_status(f"'{value}' in Zwischenablage kopiert.")

    def _close_dialog(self, dialog: tk.Toplevel) -> None:
        dialog.grab_release()
        dialog.destroy()
        self.spirit_entry.focus_set()

    def _open_draft_email(self, dialog: tk.Toplevel) -> None:
        try:
            open_mail_client("mailto:")
        except MailClientError as exc:
            messagebox.showerror("E-Mail", str(exc))
        else:
            self._update_status("Mailclient wurde geöffnet.")
        finally:
            dialog.grab_release()
            dialog.destroy()
            self.spirit_entry.focus_set()


def run_app(config: AppConfig, controller: SpiritLookupController) -> None:
    root = tk.Tk()
    app = SpiritLookupApp(root, controller, config)
    root.mainloop()
