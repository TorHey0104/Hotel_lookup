#!/usr/bin/env python3
import os
from datetime import datetime

# Ensure pandas (and openpyxl) are available; try to import and attempt to install if missing
# Add `# type: ignore` to silence editors/linters that cannot resolve the package in the current environment.
try:
    import pandas as pd  # type: ignore
except Exception:
    import sys
    import subprocess
    try:
        # Try to install pandas and openpyxl (openpyxl is needed by pandas for .xlsx)
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "pandas", "openpyxl"])
        import importlib
        importlib.invalidate_caches()
        import pandas as pd  # type: ignore
    except Exception as e:
        # Provide a clear error if automatic installation fails
        raise ImportError("Could not import or install 'pandas' and/or 'openpyxl'. Please install them manually (e.g. pip install pandas openpyxl).") from e

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import urllib.parse # For URL encoding email addresses
import webbrowser

# ————————————————————————————————————
# ◉ CONFIGURE THIS
# ————————————————————————————————————
DATA_DIR  = r"C:\Users\4612135\OneDrive - Hyatt Hotels\___DATA"
FILE_NAME = "2a Hotels one line hotel.xlsx"
DEFAULT_FILE_PATH = os.path.join(DATA_DIR, FILE_NAME)

# Runtime data containers (populated by load_data)
df = pd.DataFrame()
hotel_names = []
data_file_path = ""


def format_timestamp(path: str) -> str:
    """Return a human friendly timestamp for the given file path."""
    try:
        mod_time = datetime.fromtimestamp(os.path.getmtime(path))
    except (FileNotFoundError, OSError):
        return "Unknown timestamp"
    return mod_time.strftime("%d.%m.%Y %H:%M")


def update_status():
    """Refresh status line and file label with current metadata."""
    if status_var is None:
        return

    if data_file_path and os.path.isfile(data_file_path):
        hotel_count = len(df) if not df.empty else 0
        status_var.set(
            f"Datei: {os.path.basename(data_file_path)} • Stand: {format_timestamp(data_file_path)} • Hotels geladen: {hotel_count}"
        )
    else:
        status_var.set("Keine Datendatei geladen")


def load_data(path: str):
    """Load Excel data and refresh UI widgets."""
    global df, hotel_names, data_file_path

    new_df = pd.read_excel(path, engine="openpyxl")
    # Ensure critical columns exist
    if 'Hotel' not in new_df.columns:
        raise ValueError("Die ausgewählte Datei enthält keine Spalte 'Hotel'.")

    df = new_df
    hotel_names = sorted(df['Hotel'].dropna().unique().tolist())
    data_file_path = path

    if hotel_combo is not None:
        hotel_combo['values'] = hotel_names

    update_status()


def prompt_for_file():
    """Ask user to select an Excel file and load it."""
    initial_dir = DATA_DIR if os.path.isdir(DATA_DIR) else os.getcwd()
    file_path = filedialog.askopenfilename(
        title="Excel-Datei auswählen",
        initialdir=initial_dir,
        filetypes=[("Excel-Dateien", "*.xlsx *.xlsm *.xls"), ("Alle Dateien", "*.*")],
    )
    if not file_path:
        return
    try:
        load_data(file_path)
    except Exception as exc:
        messagebox.showerror("Laden fehlgeschlagen", f"Die Datei konnte nicht geladen werden:\n{exc}")


def ensure_initial_data():
    """Load default data file if present, otherwise ask the user."""
    if os.path.isfile(DEFAULT_FILE_PATH):
        try:
            load_data(DEFAULT_FILE_PATH)
        except Exception as exc:
            messagebox.showerror("Startfehler", f"Konnte die Standarddatei nicht laden:\n{exc}")
    else:
        messagebox.showinfo(
            "Datendatei wählen",
            "Die Standarddatendatei wurde nicht gefunden. Bitte wählen Sie eine Datei über 'Datei → Datendatei öffnen …'.",
        )
        update_status()

# Build GUI
root = tk.Tk()
root.title("Hotel Lookup")

# Keep references for widgets configured after load_data
hotel_combo = None
status_var = tk.StringVar(value="Lade Daten …")


def open_data_file():
    prompt_for_file()


# Menu bar for data management
menubar = tk.Menu(root)
file_menu = tk.Menu(menubar, tearoff=0)
file_menu.add_command(label="Datendatei öffnen …", command=open_data_file)
file_menu.add_separator()
file_menu.add_command(label="Beenden", command=root.quit)
menubar.add_cascade(label="Datei", menu=file_menu)
root.config(menu=menubar)

# Spirit Code entry
tk.Label(root, text="Spirit Code:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
code_entry = tk.Entry(root, width=30)
code_entry.grid(row=0, column=1, padx=5, pady=5)

# Hotel combobox w/ autocomplete
tk.Label(root, text="Hotel:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
hotel_var   = tk.StringVar()
hotel_combo = ttk.Combobox(root, textvariable=hotel_var, values=hotel_names)
hotel_combo.grid(row=1, column=1, padx=5, pady=5)
hotel_combo.state(["!readonly"])

def on_hotel_keyrelease(event):
    val = hotel_var.get()
    hotel_combo['values'] = (
        hotel_names if not val
        else [h for h in hotel_names if val.lower() in h.lower()]
    )

hotel_combo.bind('<KeyRelease>', on_hotel_keyrelease)

def draft_email(checkbox_vars, hotel_name, details_window):
    """
    Drafts an Outlook email with selected recipients and closes the details window.
    """
    recipients = []
    # Iterate through the list of (BooleanVar, email) tuples
    for var, email in checkbox_vars:
        if var.get() and email: # If checkbox is checked and email exists
            recipients.append(email)
    
    if not recipients:
        messagebox.showinfo("No Recipients", "No email addresses selected.")
    else:
        # Mailto supports comma separators; Windows typically also accepts semicolons
        to_addresses = ",".join(recipients)
        subject = urllib.parse.quote(f"Hotel Information for {hotel_name}") # URL-encode subject

        # Construct mailto URI
        mailto_uri = f"mailto:{to_addresses}?subject={subject}"

        try:
            # Use Python's webbrowser module for cross-platform mailto handling
            opened = webbrowser.open(mailto_uri)
            if not opened:
                raise RuntimeError("webbrowser did not open the mail client")
        except Exception as e:
            # Fallback for Windows environments where webbrowser might return False
            if os.name == "nt":
                try:
                    os.startfile(mailto_uri)  # type: ignore[attr-defined]
                except Exception as win_err:
                    messagebox.showerror("Email Error", f"Could not open email client: {win_err}")
            else:
                messagebox.showerror("Email Error", f"Could not open email client: {e}")

    details_window.destroy() # Close the details window after attempting to draft email

def show_details_gui(row):
    """
    Displays hotel details in a new GUI window with checkboxes for roles.
    """
    win = tk.Toplevel(root)
    win.title(f"Details for {row.get('Hotel', 'N/A')}")
    win.geometry("700x610")
    win.minsize(500, 300)

    # Frame for general information
    info_frame = ttk.LabelFrame(win, text="Hotel Information", padding="10")
    info_frame.pack(padx=10, pady=10, fill="x")

    general_info_cols = [
        'Spirit Code', 'Hotel', 'City', 'Geographical Area', 'Relationship','Rooms', 'JV', 'JV Percent', 'Owner'
    ]
    
    for i, col in enumerate(general_info_cols):
        if col in row and pd.notna(row[col]):
            tk.Label(info_frame, text=f"{col}:", anchor="w", font=("Arial", 10, "bold")).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            tk.Label(info_frame, text=row[col], anchor="w", font=("Arial", 10)).grid(row=i, column=1, sticky="w", padx=5, pady=2)
    
    # Frame for roles with checkboxes
    roles_frame = ttk.LabelFrame(win, text="Key Personnel (Select for Email)", padding="10")
    roles_frame.pack(padx=10, pady=10, fill="both", expand=False) 

    # Define roles and their actual email column names from the original script
    # Corrected 'RVP of OPS' mapping to 'RVP of Ops' as per original script
    roles_to_checkbox = {
        'SVP': 'SVP',
        'RVP of OPS': 'RVP of Ops',
        'AVP of Ops': 'AVP of Ops',
        'AVP of Ops-managed': 'AVP of Ops-managed',
        'GM - Primary': 'GM - Primary',
        'GM': 'GM',
        'DOF': 'DOF',
        'Senior Director of Engineering': 'Senior Director of Engineering',
        'Engineering Director': 'Engineering Director',
        'Engineering Director / Chief Engineer': 'Engineering Director / Chief Engineer'
    }

    # Changed checkbox_vars to a list to store (BooleanVar, email_address) tuples
    checkbox_vars = [] 
    
    row_idx = 0
    for role, email_col in roles_to_checkbox.items():
        email_address = row.get(email_col)
        
        # Only create checkbox if the email column exists and has a value
        if email_col in row.index and pd.notna(email_address):
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(roles_frame, text=f"{role}: {email_address}", variable=var)
            chk.grid(row=row_idx, column=0, sticky="w", padx=5, pady=2)
            # Store the BooleanVar and the email string as a tuple in the list
            checkbox_vars.append((var, str(email_address))) 
            row_idx += 1
        else:
            # Optionally, display roles without email as plain text
            tk.Label(roles_frame, text=f"{role}: N/A (Email not found)", anchor="w", foreground="gray").grid(row=row_idx, column=0, sticky="w", padx=5, pady=2)
            row_idx += 1

    # Buttons
    button_frame = ttk.Frame(win)
    button_frame.pack(pady=10) # This should now be visible

    tk.Button(button_frame, text="Close", command=win.destroy).pack(side="left", padx=10)
    tk.Button(button_frame, text="Draft Email", command=lambda: draft_email(checkbox_vars, row.get('Hotel', 'N/A'), win)).pack(side="left", padx=10)

def lookup():
    if df.empty:
        messagebox.showwarning(
            "Keine Daten",
            "Es sind derzeit keine Daten geladen. Bitte laden Sie eine Excel-Datei über 'Datei → Datendatei öffnen …'.",
        )
        return

    spirit = code_entry.get().strip()
    hotel   = hotel_var.get().strip()
    if spirit:
        # Spirit Codes können numerisch sein: deshalb als String vergleichen
        result = df[df['Spirit Code'].astype(str).str.lower() == spirit.lower()]
    elif hotel:
        # Unscharfe Suche im Hotelnamen und in der Stadtspalte
        mask = df['Hotel'].astype(str).str.contains(hotel, case=False, na=False)
        if 'City' in df.columns:
            mask |= df['City'].astype(str).str.contains(hotel, case=False, na=False)
        result = df[mask]
    else:
        messagebox.showwarning("Whoops", "Enter Spirit Code or pick a hotel.")
        return

    if result.empty:
        messagebox.showinfo("Nada", "No matching hotel found.")
        return

    if len(result) == 1:
        row = result.iloc[0]
        show_details_gui(row)
    else:
        show_search_results(result)


def show_search_results(results_df):
    """Display a window with multiple matches for user selection."""
    win = tk.Toplevel(root)
    win.title("Suchergebnisse")
    win.geometry("500x300")

    tree = ttk.Treeview(win, columns=("Spirit Code", "Hotel", "City"), show="headings")
    tree.heading("Spirit Code", text="Spirit Code")
    tree.heading("Hotel", text="Hotel")
    tree.heading("City", text="Stadt")
    tree.column("Spirit Code", width=100)
    tree.column("Hotel", width=250)
    tree.column("City", width=130)
    tree.pack(fill="both", expand=True, padx=10, pady=10)

    for _, result_row in results_df.iterrows():
        tree.insert("", "end", values=(result_row.get('Spirit Code', ""), result_row.get('Hotel', ""), result_row.get('City', "")))

    def open_selected(event=None):
        selected = tree.focus()
        if not selected:
            messagebox.showinfo("Auswahl", "Bitte wählen Sie einen Eintrag aus.")
            return
        try:
            row_index = tree.index(selected)
        except tk.TclError:
            messagebox.showerror("Fehler", "Die Auswahl konnte nicht ermittelt werden.")
            return

        if row_index >= len(results_df):
            messagebox.showerror("Fehler", "Der ausgewählte Eintrag konnte nicht geladen werden.")
            return

        show_details_gui(results_df.iloc[row_index])
        win.destroy()

    btn_frame = ttk.Frame(win)
    btn_frame.pack(pady=(0, 10))
    ttk.Button(btn_frame, text="Details anzeigen", command=open_selected).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="Schließen", command=win.destroy).pack(side="left", padx=5)

    tree.bind("<Double-1>", open_selected)

tk.Button(root, text="Search", command=lookup).grid(row=2, column=0, columnspan=2, pady=10)

# Status bar
status_label = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor="w")
status_label.grid(row=3, column=0, columnspan=2, sticky="we", padx=0, pady=(5, 0))

# Load default data after UI widgets are ready
ensure_initial_data()

root.mainloop()
