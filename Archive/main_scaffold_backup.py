#!/usr/bin/env python3
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date
import pandas as pd
from filters import apply_filters
from mail_utils import render_with_signature, get_outlook_app, save_forward
from ui_common import make_multiselect
from data import remember_config, load_recent_configs
from roles import ROLE_KEYS, get_role_map

TOOL_NAME = "Hyatt EAME Hotel Lookup and Multi E-Mail Tool"
VERSION = "7.0.0"
VERSION_DATE = date.today().strftime("%d.%m.%Y")

# Minimal scaffolding: status + notebook
root = tk.Tk()
root.title("Hotel Lookup v7")
root.geometry("1100x800")
style = ttk.Style()
style.map("TNotebook.Tab", background=[("selected", "#1f4fa3")], foreground=[("selected", "white")])
style.configure("TNotebook.Tab", padding=(8, 4))

status_var = tk.StringVar(value="Ready")
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

status_label = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor="w")
status_label.pack(fill="x", side="bottom")

# Placeholder tabs; full UI to be built by dedicated modules
lookup_frame = ttk.Frame(notebook, padding=10)
notebook.add(lookup_frame, text="Lookup")
multi_frame = ttk.Frame(notebook, padding=10)
notebook.add(multi_frame, text="Multi-Email")
excel_frame = ttk.Frame(notebook, padding=10)
notebook.add(excel_frame, text="Excel Emails")
setup_frame = ttk.Frame(notebook, padding=10)
notebook.add(setup_frame, text="Setup")

if __name__ == "__main__":
    root.mainloop()
