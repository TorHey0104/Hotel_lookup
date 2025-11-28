#!/usr/bin/env python3
import os
from datetime import datetime

# Ensure pandas (and openpyxl) are available; try to import and attempt to install if missing
# Add "# type: ignore" to silence editors/linters that cannot resolve the package in the current environment.
try:
    import pandas as pd  # type: ignore
except Exception:
    import sys
    import subprocess
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "pandas", "openpyxl"])
        import importlib
        importlib.invalidate_caches()
        import pandas as pd  # type: ignore
    except Exception as e:  # pragma: no cover - auto-install fallback
        raise ImportError(
            "Could not import or install 'pandas' and/or 'openpyxl'. Please install them manually (e.g. pip install pandas openpyxl)."
        ) from e

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import importlib.util

# Cache Outlook availability and instance so email drafting is faster after the first use
WIN32COM_AVAILABLE = os.name == "nt" and importlib.util.find_spec("win32com.client") is not None
_outlook_app = None

# ---------------------------------------------------------------------------
# CONFIGURE THIS
# ---------------------------------------------------------------------------
DATA_DIR = r"C:\Users\4612135\OneDrive - Hyatt Hotels\___DATA"
FILE_NAME = "2a Hotels one line hotel.xlsx"
DEFAULT_FILE_PATH = os.path.join(DATA_DIR, FILE_NAME)

# Column names used for filtering
BRAND_COL = "Brand"
REGION_COL = "Region"
COUNTRY_COL = "Geography"  # Primary country/market column
COUNTRY_FALLBACK_COL = "Geographical Area"  # Fallback if Geography is missing

# Runtime data containers (populated by load_data)
df = pd.DataFrame()
hotel_names = []
data_file_path = ""
brand_values = []
region_values = []
country_values = []

# Tk widgets (assigned after root creation)
hotel_combo = None
status_var = None
brand_filter_var = None
region_filter_var = None
country_filter_var = None
filtered_tree = None
selected_tree = None
selected_rows = {}
current_filtered_indexes = []


def format_timestamp(path: str) -> str:
    """Return a human friendly timestamp for the given file path."""
    try:
        mod_time = datetime.fromtimestamp(os.path.getmtime(path))
    except (FileNotFoundError, OSError):
        return "Unknown timestamp"
    return mod_time.strftime("%d.%m.%Y %H:%M")


def get_country_value(row: pd.Series) -> str:
    """Return the country/area value from the configured columns."""
    if COUNTRY_COL in row and pd.notna(row[COUNTRY_COL]):
        return str(row[COUNTRY_COL])
    if COUNTRY_FALLBACK_COL in row and pd.notna(row[COUNTRY_FALLBACK_COL]):
        return str(row[COUNTRY_FALLBACK_COL])
    return ""


def update_status():
    """Refresh status line and file label with current metadata."""
    global status_var
    if status_var is None:
        return

    if data_file_path and os.path.isfile(data_file_path):
        hotel_count = len(df) if not df.empty else 0
        status_var.set(
            f"Datei: {os.path.basename(data_file_path)} | Stand: {format_timestamp(data_file_path)} | Hotels geladen: {hotel_count}"
        )
    else:
        status_var.set("Keine Datendatei geladen")


def update_filter_options():
    """Populate filter dropdowns based on loaded data."""
    global brand_values, region_values, country_values
    if df.empty:
        brand_values = []
        region_values = []
        country_values = []
    else:
        brand_values = sorted(df[BRAND_COL].dropna().astype(str).unique().tolist()) if BRAND_COL in df.columns else []
        region_values = sorted(df[REGION_COL].dropna().astype(str).unique().tolist()) if REGION_COL in df.columns else []
        if COUNTRY_COL in df.columns:
            country_values = sorted(df[COUNTRY_COL].dropna().astype(str).unique().tolist())
        elif COUNTRY_FALLBACK_COL in df.columns:
            country_values = sorted(df[COUNTRY_FALLBACK_COL].dropna().astype(str).unique().tolist())
        else:
            country_values = []

    if brand_filter_var is not None:
        brand_filter_var.set("Any")
    if region_filter_var is not None:
        region_filter_var.set("Any")
    if country_filter_var is not None:
        country_filter_var.set("Any")

    if filtered_tree is not None:
        refresh_filtered_hotels()


def load_data(path: str):
    """Load Excel data and refresh UI widgets."""
    global df, hotel_names, data_file_path

    new_df = pd.read_excel(path, engine="openpyxl")
    if "Hotel" not in new_df.columns:
        raise ValueError("Die ausgewaehlte Datei enthaelt keine Spalte 'Hotel'.")

    df = new_df
    hotel_names = sorted(df["Hotel"].dropna().unique().tolist())
    data_file_path = path

    if hotel_combo is not None:
        hotel_combo["values"] = hotel_names

    update_filter_options()
    update_status()


def prompt_for_file():
    """Ask user to select an Excel file and load it."""
    initial_dir = DATA_DIR if os.path.isdir(DATA_DIR) else os.getcwd()
    file_path = filedialog.askopenfilename(
        title="Excel-Datei auswaehlen",
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
            "Datendatei waehlen",
            "Die Standarddatendatei wurde nicht gefunden. Bitte waehlen Sie eine Datei ueber 'Datei -> Datendatei oeffnen'.",
        )
        update_status()


def get_outlook_app(force_refresh: bool = False):
    """Return a cached Outlook Application COM object (or create it on first use)."""
    global _outlook_app

    if _outlook_app is not None and not force_refresh:
        return _outlook_app

    import win32com.client  # type: ignore[import-untyped]

    try:
        _outlook_app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    except Exception:
        _outlook_app = win32com.client.Dispatch("Outlook.Application")

    return _outlook_app


def warm_outlook_app():
    """Preload Outlook once to speed up the first email draft."""
    if os.name == "nt" and WIN32COM_AVAILABLE:
        try:
            get_outlook_app()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Multi-select helpers
# ---------------------------------------------------------------------------

def filtered_dataframe():
    """Return dataframe filtered by current dropdown selections."""
    if df.empty:
        return pd.DataFrame()

    filt = df
    if brand_filter_var is not None:
        brand_val = brand_filter_var.get()
        if brand_val and brand_val != "Any" and BRAND_COL in filt.columns:
            filt = filt[filt[BRAND_COL].astype(str) == brand_val]
    if region_filter_var is not None:
        region_val = region_filter_var.get()
        if region_val and region_val != "Any" and REGION_COL in filt.columns:
            filt = filt[filt[REGION_COL].astype(str) == region_val]
    if country_filter_var is not None:
        country_val = country_filter_var.get()
        if country_val and country_val != "Any":
            if COUNTRY_COL in filt.columns:
                filt = filt[filt[COUNTRY_COL].astype(str) == country_val]
            elif COUNTRY_FALLBACK_COL in filt.columns:
                filt = filt[filt[COUNTRY_FALLBACK_COL].astype(str) == country_val]
    return filt


def refresh_filtered_hotels():
    """Refresh the filtered hotels list."""
    global current_filtered_indexes
    if filtered_tree is None:
        return

    for item in filtered_tree.get_children():
        filtered_tree.delete(item)

    filt_df = filtered_dataframe()
    current_filtered_indexes = []

    for idx, (_, row) in enumerate(filt_df.iterrows()):
        tree_id = str(row.name)
        current_filtered_indexes.append(row.name)
        filtered_tree.insert(
            "",
            "end",
            iid=tree_id,
            values=(
                row.get("Spirit Code", ""),
                row.get("Hotel", ""),
                row.get("City", ""),
                row.get(BRAND_COL, ""),
                row.get(REGION_COL, ""),
                get_country_value(row),
            ),
        )


def add_selected_hotels():
    """Add selected rows from the filtered tree to the selection list."""
    if filtered_tree is None:
        return
    selected = filtered_tree.selection()
    if not selected:
        messagebox.showinfo("Auswahl", "Bitte waehlen Sie mindestens ein Hotel aus der Filterliste aus.")
        return

    for tree_id in selected:
        try:
            row_idx = int(tree_id)
        except ValueError:
            continue
        if row_idx in selected_rows:
            continue
        row = df.loc[row_idx]
        selected_rows[row_idx] = row
    update_selected_tree()


def remove_selected_hotels():
    """Remove selected rows from the selection list."""
    if selected_tree is None:
        return
    chosen = selected_tree.selection()
    for tree_id in chosen:
        try:
            row_idx = int(tree_id)
        except ValueError:
            continue
        selected_rows.pop(row_idx, None)
    update_selected_tree()


def clear_selected_hotels():
    """Clear all selections."""
    selected_rows.clear()
    update_selected_tree()


def update_selected_tree():
    """Refresh the selected hotels list."""
    if selected_tree is None:
        return
    for item in selected_tree.get_children():
        selected_tree.delete(item)
    for row_idx, row in selected_rows.items():
        selected_tree.insert(
            "",
            "end",
            iid=str(row_idx),
            values=(
                row.get("Spirit Code", ""),
                row.get("Hotel", ""),
                row.get("City", ""),
                row.get(BRAND_COL, ""),
                row.get(REGION_COL, ""),
                get_country_value(row),
            ),
        )


def get_role_addresses(row: pd.Series, role_key: str):
    """Return a list of email addresses for the chosen role."""
    role_map = {
        "GM": ["GM - Primary", "GM"],
        "Engineering": ["Engineering Director / Chief Engineer", "Engineering Director"],
        "DOF": ["DOF"],
    }
    emails = []
    for col in role_map.get(role_key, []):
        if col in row and pd.notna(row[col]):
            emails.append(str(row[col]))
    return emails


def draft_emails_for_selection(gm_var, eng_var, dof_var):
    """Create Outlook draft emails for the selected hotels and roles."""
    if not selected_rows:
        messagebox.showinfo("Keine Hotels", "Bitte waehlen Sie mindestens ein Hotel aus der Auswahl aus.")
        return

    chosen_roles = []
    if gm_var.get():
        chosen_roles.append("GM")
    if eng_var.get():
        chosen_roles.append("Engineering")
    if dof_var.get():
        chosen_roles.append("DOF")

    if not chosen_roles:
        messagebox.showinfo("Keine Rollen", "Bitte waehlen Sie mindestens eine Empfaengerrolle (GM/Engineering/DOF).")
        return

    if os.name != "nt":
        messagebox.showerror("Unsupported Platform", "Outlook email drafting is only available on Windows.")
        return

    if not WIN32COM_AVAILABLE:
        messagebox.showerror(
            "Outlook Not Available",
            "This feature requires Microsoft Outlook and the 'pywin32' package (win32com.client).\nInstall with: pip install pywin32",
        )
        return

    try:
        outlook = get_outlook_app()
        mail_test = outlook.CreateItem(0)
    except Exception:
        try:
            outlook = get_outlook_app(force_refresh=True)
            mail_test = outlook.CreateItem(0)
        except Exception as exc:  # pragma: no cover - Outlook automation is Windows-specific
            messagebox.showerror("Email Error", f"Could not draft email in Outlook: {exc}")
            return

    missing_addresses = []

    for row_idx, row in selected_rows.items():
        recipients = []
        for role in chosen_roles:
            recipients.extend(get_role_addresses(row, role))
        recipients = [r for r in recipients if r]
        if not recipients:
            missing_addresses.append(row.get("Hotel", "N/A"))
            continue

        try:
            mail_item = outlook.CreateItem(0)
            mail_item.To = ";".join(recipients)
            hotel_name = row.get("Hotel", "Hotel")
            mail_item.Subject = f"Hotel Information for {hotel_name}"
            body_lines = [
                f"Hotel: {hotel_name}",
                f"City: {row.get('City', '')}",
                f"Brand: {row.get(BRAND_COL, '')}",
                f"Region: {row.get(REGION_COL, '')}",
                f"Country/Area: {get_country_value(row)}",
                "",
                "Please add your message here.",
            ]
            mail_item.Body = "\n".join(body_lines)
            mail_item.Display()
        except Exception as exc:
            messagebox.showerror("Email Error", f"Could not draft email for {row.get('Hotel', 'Hotel')}: {exc}")
            return

    if missing_addresses:
        messagebox.showinfo(
            "Keine Empfaenger", "Fuer folgende Hotels wurden keine E-Mail-Adressen in den gewaehlten Rollen gefunden:\n" + "\n".join(missing_addresses)
        )


# ---------------------------------------------------------------------------
# Lookup (single hotel) helpers
# ---------------------------------------------------------------------------

def lookup(spirit_entry, hotel_var_local):
    if df.empty:
        messagebox.showwarning(
            "Keine Daten",
            "Es sind derzeit keine Daten geladen. Bitte laden Sie eine Excel-Datei ueber 'Datei -> Datendatei oeffnen'.",
        )
        return

    spirit = spirit_entry.get().strip()
    hotel = hotel_var_local.get().strip()
    if spirit:
        result = df[df["Spirit Code"].astype(str).str.lower() == spirit.lower()]
    elif hotel:
        mask = df["Hotel"].astype(str).str.contains(hotel, case=False, na=False)
        if "City" in df.columns:
            mask |= df["City"].astype(str).str.contains(hotel, case=False, na=False)
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


def show_details_gui(row):
    win = tk.Toplevel(root)
    win.title(f"Details for {row.get('Hotel', 'N/A')}")
    win.geometry("700x760")
    win.minsize(500, 400)

    info_frame = ttk.LabelFrame(win, text="Hotel Information", padding="10")
    info_frame.pack(padx=10, pady=10, fill="x")

    general_info_cols = [
        "Spirit Code",
        "Hotel",
        "City",
        "Geographical Area",
        "Relationship",
        "Rooms",
        "JV",
        "JV Percent",
        "Owner",
    ]

    for i, col in enumerate(general_info_cols):
        if col in row and pd.notna(row[col]):
            tk.Label(info_frame, text=f"{col}:", anchor="w", font=("Arial", 10, "bold")).grid(row=i, column=0, sticky="w", padx=5, pady=2)
            tk.Label(info_frame, text=row[col], anchor="w", font=("Arial", 10)).grid(row=i, column=1, sticky="w", padx=5, pady=2)

    roles_frame = ttk.LabelFrame(win, text="Key Personnel (Select for Email)", padding="10")
    roles_frame.pack(padx=10, pady=10, fill="both", expand=False)

    roles_to_checkbox = {
        "SVP": "SVP",
        "RVP of OPS": "RVP of Ops",
        "AVP of Ops": "AVP of Ops",
        "AVP of Ops-managed": "AVP of Ops-managed",
        "GM - Primary": "GM - Primary",
        "GM": "GM",
        "DOF": "DOF",
        "Senior Director of Engineering": "Senior Director of Engineering",
        "Engineering Director": "Engineering Director",
        "Engineering Director / Chief Engineer": "Engineering Director / Chief Engineer",
    }

    checkbox_vars = []
    row_idx = 0
    for role, email_col in roles_to_checkbox.items():
        email_address = row.get(email_col)
        if email_col in row.index and pd.notna(email_address):
            var = tk.BooleanVar()
            chk = ttk.Checkbutton(roles_frame, text=f"{role}: {email_address}", variable=var)
            chk.grid(row=row_idx, column=0, sticky="w", padx=5, pady=2)
            checkbox_vars.append((var, str(email_address)))
            row_idx += 1
        else:
            tk.Label(roles_frame, text=f"{role}: N/A (Email not found)", anchor="w", foreground="gray").grid(
                row=row_idx, column=0, sticky="w", padx=5, pady=2
            )
            row_idx += 1

    button_frame = ttk.Frame(win)
    button_frame.pack(pady=10)

    tk.Button(button_frame, text="Close", command=win.destroy).pack(side="left", padx=10)
    tk.Button(
        button_frame,
        text="Start Email",
        command=lambda: draft_email_single(checkbox_vars, row.get("Hotel", "N/A"), win),
    ).pack(side="left", padx=10)


def draft_email_single(checkbox_vars, hotel_name, details_window):
    recipients = []
    for var, email in checkbox_vars:
        if var.get() and email:
            recipients.append(email)

    if not recipients:
        messagebox.showinfo("No Recipients", "No email addresses selected.")
        return

    if os.name != "nt":
        messagebox.showerror("Unsupported Platform", "Outlook email drafting is only available on Windows.")
        return

    if not WIN32COM_AVAILABLE:
        messagebox.showerror(
            "Outlook Not Available",
            "This feature requires Microsoft Outlook and the 'pywin32' package (win32com.client).\nInstall with: pip install pywin32",
        )
        return

    try:
        outlook = get_outlook_app()
        mail_item = outlook.CreateItem(0)
    except Exception:
        try:
            outlook = get_outlook_app(force_refresh=True)
            mail_item = outlook.CreateItem(0)
        except Exception as exc:  # pragma: no cover - Outlook automation is Windows-specific
            messagebox.showerror("Email Error", f"Could not draft email in Outlook: {exc}")
            _outlook_app = None
            return

    mail_item.To = ";".join(recipients)
    mail_item.Subject = f"Hotel Information for {hotel_name}"
    mail_item.Display()
    details_window.destroy()


def show_search_results(results_df):
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
        tree.insert("", "end", values=(result_row.get("Spirit Code", ""), result_row.get("Hotel", ""), result_row.get("City", "")))

    def open_selected(event=None):
        selected = tree.focus()
        if not selected:
            messagebox.showinfo("Auswahl", "Bitte waehlen Sie einen Eintrag aus.")
            return
        try:
            row_index = tree.index(selected)
        except tk.TclError:
            messagebox.showerror("Fehler", "Die Auswahl konnte nicht ermittelt werden.")
            return

        if row_index >= len(results_df):
            messagebox.showerror("Fehler", "Der ausgewaehlte Eintrag konnte nicht geladen werden.")
            return

        show_details_gui(results_df.iloc[row_index])
        win.destroy()

    btn_frame = ttk.Frame(win)
    btn_frame.pack(pady=(0, 10))
    ttk.Button(btn_frame, text="Details anzeigen", command=open_selected).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="Schliessen", command=win.destroy).pack(side="left", padx=5)

    tree.bind("<Double-1>", open_selected)


# ---------------------------------------------------------------------------
# GUI construction
# ---------------------------------------------------------------------------
root = tk.Tk()
root.title("Hotel Lookup")
root.geometry("1050x720")

status_var = tk.StringVar(value="Lade Daten ...")

# Menu bar
menubar = tk.Menu(root)
file_menu = tk.Menu(menubar, tearoff=0)
file_menu.add_command(label="Datendatei oeffnen", command=prompt_for_file)
file_menu.add_separator()
file_menu.add_command(label="Beenden", command=root.quit)
menubar.add_cascade(label="Datei", menu=file_menu)
root.config(menu=menubar)

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# ---------------------------------------------------------------------------
# Tab 1: Lookup
# ---------------------------------------------------------------------------
lookup_frame = ttk.Frame(notebook, padding=10)
notebook.add(lookup_frame, text="Lookup")

# Spirit Code entry
spirit_label = tk.Label(lookup_frame, text="Spirit Code:")
spirit_label.grid(row=0, column=0, sticky="e", padx=5, pady=5)
spirit_entry = tk.Entry(lookup_frame, width=30)
spirit_entry.grid(row=0, column=1, padx=5, pady=5)

# Hotel combobox w/ autocomplete
hotel_label = tk.Label(lookup_frame, text="Hotel:")
hotel_label.grid(row=1, column=0, sticky="e", padx=5, pady=5)
hotel_var = tk.StringVar()
hotel_combo = ttk.Combobox(lookup_frame, textvariable=hotel_var, values=hotel_names)
hotel_combo.grid(row=1, column=1, padx=5, pady=5)
hotel_combo.state(["!readonly"])


def on_hotel_keyrelease(event):
    val = hotel_var.get()
    hotel_combo["values"] = hotel_names if not val else [h for h in hotel_names if val.lower() in h.lower()]


hotel_combo.bind("<KeyRelease>", on_hotel_keyrelease)

search_button = tk.Button(lookup_frame, text="Search", command=lambda: lookup(spirit_entry, hotel_var))
search_button.grid(row=2, column=0, columnspan=2, pady=10)

# ---------------------------------------------------------------------------
# Tab 2: Multi-email
# ---------------------------------------------------------------------------
multi_frame = ttk.Frame(notebook, padding=10)
notebook.add(multi_frame, text="Multi-Email")

filters_frame = ttk.LabelFrame(multi_frame, text="Filter Hotels", padding=10)
filters_frame.pack(fill="x", padx=5, pady=5)

brand_filter_var = tk.StringVar(value="Any")
region_filter_var = tk.StringVar(value="Any")
country_filter_var = tk.StringVar(value="Any")

# Filter controls
row_f = 0
if True:
    ttk.Label(filters_frame, text="Brand").grid(row=row_f, column=0, sticky="w", padx=5, pady=2)
    brand_combo = ttk.Combobox(filters_frame, textvariable=brand_filter_var, values=["Any"], state="readonly")
    brand_combo.grid(row=row_f, column=1, sticky="ew", padx=5, pady=2)

    ttk.Label(filters_frame, text="Region").grid(row=row_f, column=2, sticky="w", padx=5, pady=2)
    region_combo = ttk.Combobox(filters_frame, textvariable=region_filter_var, values=["Any"], state="readonly")
    region_combo.grid(row=row_f, column=3, sticky="ew", padx=5, pady=2)

    ttk.Label(filters_frame, text="Country/Area").grid(row=row_f, column=4, sticky="w", padx=5, pady=2)
    country_combo = ttk.Combobox(filters_frame, textvariable=country_filter_var, values=["Any"], state="readonly")
    country_combo.grid(row=row_f, column=5, sticky="ew", padx=5, pady=2)

    filters_frame.columnconfigure(1, weight=1)
    filters_frame.columnconfigure(3, weight=1)
    filters_frame.columnconfigure(5, weight=1)

apply_filter_btn = ttk.Button(filters_frame, text="Apply Filter", command=refresh_filtered_hotels)
apply_filter_btn.grid(row=0, column=6, sticky="e", padx=8, pady=2)

# Trees
lists_frame = ttk.Frame(multi_frame)
lists_frame.pack(fill="both", expand=True, padx=5, pady=5)

filtered_frame = ttk.LabelFrame(lists_frame, text="Gefilterte Hotels", padding=5)
filtered_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

filtered_tree = ttk.Treeview(
    filtered_frame,
    columns=("Spirit", "Hotel", "City", "Brand", "Region", "Country"),
    show="headings",
    selectmode="extended",
)
for col, width in [
    ("Spirit", 80),
    ("Hotel", 220),
    ("City", 120),
    ("Brand", 120),
    ("Region", 120),
    ("Country", 140),
]:
    filtered_tree.heading(col, text=col)
    filtered_tree.column(col, width=width, stretch=True)
filtered_tree.pack(fill="both", expand=True)

buttons_frame = ttk.Frame(lists_frame)
buttons_frame.pack(side="left", fill="y")

ttk.Button(buttons_frame, text=">>>", command=add_selected_hotels).pack(pady=8)
ttk.Button(buttons_frame, text="Remove", command=remove_selected_hotels).pack(pady=8)
ttk.Button(buttons_frame, text="Clear All", command=clear_selected_hotels).pack(pady=8)

selected_frame = ttk.LabelFrame(lists_frame, text="Ausgewaehlte Hotels", padding=5)
selected_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))

selected_tree = ttk.Treeview(
    selected_frame,
    columns=("Spirit", "Hotel", "City", "Brand", "Region", "Country"),
    show="headings",
    selectmode="extended",
)
for col, width in [
    ("Spirit", 80),
    ("Hotel", 220),
    ("City", 120),
    ("Brand", 120),
    ("Region", 120),
    ("Country", 140),
]:
    selected_tree.heading(col, text=col)
    selected_tree.column(col, width=width, stretch=True)
selected_tree.pack(fill="both", expand=True)

# Role selection and draft
roles_frame = ttk.LabelFrame(multi_frame, text="Empfaengerrollen", padding=10)
roles_frame.pack(fill="x", padx=5, pady=5)

gm_var = tk.BooleanVar(value=True)
eng_var = tk.BooleanVar(value=False)
dof_var = tk.BooleanVar(value=False)

roles_label = ttk.Label(roles_frame, text="Wen anschreiben?")
roles_label.pack(side="left", padx=5)

ttk.Checkbutton(roles_frame, text="GM", variable=gm_var).pack(side="left", padx=5)

ttk.Checkbutton(roles_frame, text="Engineering", variable=eng_var).pack(side="left", padx=5)

ttk.Checkbutton(roles_frame, text="DOF", variable=dof_var).pack(side="left", padx=5)

draft_btn = ttk.Button(roles_frame, text="Draft Emails in Outlook", command=lambda: draft_emails_for_selection(gm_var, eng_var, dof_var))
draft_btn.pack(side="right", padx=10)

# Status bar
status_label = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor="w")
status_label.pack(fill="x", side="bottom")

# Load default data after UI widgets are ready
ensure_initial_data()

# Populate dropdown values after data load
brand_combo["values"] = ["Any"] + brand_values
region_combo["values"] = ["Any"] + region_values
country_combo["values"] = ["Any"] + country_values

# Initial filtered view
refresh_filtered_hotels()

# Warm Outlook in the background so the first email opens faster
root.after(200, warm_outlook_app)

root.mainloop()
