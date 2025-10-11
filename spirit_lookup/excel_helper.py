"""Tkinter GUI to manage Excel to JSON helper configuration."""

from __future__ import annotations

import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from typing import Dict, List

from .excel_helper_config import (
    ExcelHelperConfigEntry,
    ExcelHelperConfigStore,
    detect_email_headers,
)


class ExcelHelperWindow:
    """Modal dialog that allows configuring Excel column selections."""

    def __init__(self, parent: tk.Tk, config_path: Path):
        self.parent = parent
        self.config_path = config_path
        self.config_store = ExcelHelperConfigStore(config_path)

        self.window = tk.Toplevel(parent)
        self.window.title("Excel Helper")
        self.window.geometry("720x520")
        self.window.transient(parent)
        self.window.grab_set()

        self.selected_excel: Path | None = None
        self.column_vars: Dict[str, tk.BooleanVar] = {}
        self.email_vars: Dict[str, tk.BooleanVar] = {}

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.window, padding=16)
        container.grid(row=0, column=0, sticky="nsew")
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)

        ttk.Label(
            container,
            text="Wählen Sie eine Excel-Datei aus und markieren Sie die relevanten Spalten.",
            wraplength=640,
        ).grid(row=0, column=0, columnspan=3, sticky="w")

        select_button = ttk.Button(container, text="Excel-Datei auswählen", command=self._choose_excel_file)
        select_button.grid(row=1, column=0, sticky="w", pady=(12, 12))

        self.file_label = ttk.Label(container, text="Keine Datei gewählt", foreground="#555555")
        self.file_label.grid(row=1, column=1, columnspan=2, sticky="w", padx=(12, 0))

        self.columns_frame = ttk.LabelFrame(container, text="Spaltenauswahl", padding=12)
        self.columns_frame.grid(row=2, column=0, columnspan=3, sticky="nsew")
        self.columns_frame.columnconfigure(0, weight=1)

        self.contacts_frame = ttk.LabelFrame(container, text="Kontakt-E-Mails", padding=12)
        self.contacts_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(12, 0))
        self.contacts_frame.columnconfigure(0, weight=1)

        button_frame = ttk.Frame(container)
        button_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(12, 0))

        self.save_button = ttk.Button(button_frame, text="Konfiguration speichern", command=self._save_config)
        self.save_button.grid(row=0, column=0, sticky="w")
        self.save_button.configure(state=tk.DISABLED)

        preview_button = ttk.Button(button_frame, text="Konfiguration anzeigen", command=self._show_preview)
        preview_button.grid(row=0, column=1, sticky="w", padx=(12, 0))

        close_button = ttk.Button(button_frame, text="Schließen", command=self.window.destroy)
        close_button.grid(row=0, column=2, sticky="e")

        self.preview_text = tk.Text(container, height=8, width=80)
        self.preview_text.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(12, 0))
        self.preview_text.configure(state="disabled")

        container.rowconfigure(2, weight=1)
        container.rowconfigure(3, weight=1)
        container.rowconfigure(5, weight=1)

        self.window.bind("<Escape>", lambda _event: self.window.destroy())

    def _clear_frames(self) -> None:
        for frame in (self.columns_frame, self.contacts_frame):
            for child in frame.winfo_children():
                child.destroy()
        self.column_vars.clear()
        self.email_vars.clear()

    def _choose_excel_file(self) -> None:
        file_path = filedialog.askopenfilename(
            parent=self.window,
            title="Excel-Datei auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
        )
        if not file_path:
            return
        excel_path = Path(file_path)
        if not excel_path.exists():
            messagebox.showerror("Datei nicht gefunden", f"Die Datei '{excel_path}' konnte nicht geöffnet werden.")
            return
        try:
            headers = self._load_headers(excel_path)
        except ModuleNotFoundError:
            messagebox.showerror(
                "openpyxl nicht installiert",
                "Für den Excel-Import wird 'openpyxl' benötigt. Installiere das Paket z. B. mit `pip install openpyxl`.",
            )
            return
        except ValueError as exc:
            messagebox.showerror("Fehler beim Lesen", str(exc))
            return

        self.selected_excel = excel_path
        self.file_label.configure(text=str(excel_path))
        self._populate_headers(headers)
        self.save_button.configure(state=tk.NORMAL)
        self._show_preview()

    def _load_headers(self, excel_path: Path) -> List[str]:
        try:
            from openpyxl import load_workbook  # type: ignore
        except ModuleNotFoundError as exc:
            raise exc
        workbook = load_workbook(excel_path, data_only=True, read_only=True)
        worksheet = workbook.active
        first_row = next(worksheet.iter_rows(values_only=True), None)
        if not first_row:
            raise ValueError("Die Arbeitsmappe enthält keine Daten.")
        headers = [str(cell).strip() if cell is not None else "" for cell in first_row]
        return headers

    def _populate_headers(self, headers: List[str]) -> None:
        self._clear_frames()
        if not headers:
            ttk.Label(self.columns_frame, text="Keine Spaltenüberschriften gefunden.").grid(sticky="w")
            ttk.Label(self.contacts_frame, text="Keine E-Mail-Spalten gefunden.").grid(sticky="w")
            return

        stored = self.config_store.get_entry(self.selected_excel) if self.selected_excel else None
        stored_columns = set(stored.selected_columns) if stored else None
        stored_emails = set(stored.email_columns) if stored else None

        for idx, header in enumerate(headers):
            if not header:
                continue
            var = tk.BooleanVar(value=True if stored_columns is None else header in stored_columns)
            check = ttk.Checkbutton(self.columns_frame, text=header, variable=var)
            check.grid(row=idx, column=0, sticky="w", pady=2)
            self.column_vars[header] = var

        email_headers = detect_email_headers(headers)
        if not email_headers:
            ttk.Label(self.contacts_frame, text="Keine Spalten mit E-Mail-Adressen erkannt.").grid(sticky="w")
        else:
            for idx, header in enumerate(email_headers):
                var = tk.BooleanVar(value=True if stored_emails is None else header in stored_emails)
                check = ttk.Checkbutton(
                    self.contacts_frame,
                    text=header,
                    variable=var,
                )
                check.grid(row=idx, column=0, sticky="w", pady=2)
                self.email_vars[header] = var

    def _save_config(self) -> None:
        if not self.selected_excel:
            messagebox.showinfo("Keine Datei", "Bitte wählen Sie zuerst eine Excel-Datei aus.")
            return
        selected_columns = [header for header, var in self.column_vars.items() if var.get()]
        email_columns = [header for header, var in self.email_vars.items() if var.get()]
        if not selected_columns:
            messagebox.showwarning("Keine Spalten gewählt", "Bitte wählen Sie mindestens eine Spalte aus.")
            return
        entry = self.config_store.save_entry(self.selected_excel, selected_columns, email_columns)
        messagebox.showinfo(
            "Gespeichert",
            "Die Konfiguration wurde gespeichert. E-Mail-Spalten sind in der Draft-Konfiguration markiert.",
        )
        self._update_preview(entry)

    def _show_preview(self) -> None:
        if not self.selected_excel:
            self._set_preview_text("{}")
            return
        pretty = self.config_store.to_pretty_json(self.selected_excel)
        self._set_preview_text(pretty)

    def _update_preview(self, entry: ExcelHelperConfigEntry) -> None:
        if not self.selected_excel:
            return
        payload = {
            "excelPath": str(self.selected_excel),
            **entry.to_dict(),
        }
        self._set_preview_text(json.dumps(payload, indent=2, ensure_ascii=False))

    def _set_preview_text(self, content: str) -> None:
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert("1.0", content)
        self.preview_text.configure(state="disabled")


def open_excel_helper(parent: tk.Tk, config_path: Path) -> None:
    """Open the Excel helper dialog."""

    ExcelHelperWindow(parent, config_path)
