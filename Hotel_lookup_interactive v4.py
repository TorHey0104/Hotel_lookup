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

# Default column names (can be overridden in Setup tab)
DEFAULT_BRAND_COL = "Brand"
DEFAULT_REGION_COL = "Region"
DEFAULT_COUNTRY_COL = "Geography"
DEFAULT_COUNTRY_FALLBACK_COL = "Geographical Area"
DEFAULT_CITY_COL = "City"
DEFAULT_BRAND_BAND_COL = "Brand Band"
DEFAULT_RELATIONSHIP_COL = "Relationship"
DEFAULT_HYATT_DATE_COL = "Affiliation Date"
DEFAULT_GM_COL = "GM - Primary"
DEFAULT_ENG_COL = "Engineering Director / Chief Engineer"
DEFAULT_DOF_COL = "DOF"
DEFAULT_REG_ENG_SPEC_COL = ""  # optional

# Runtime data containers (populated by load_data)
df = pd.DataFrame()
hotel_names = []
data_file_path = ""
brand_values = []
region_values = []
country_values = []
brand_band_values = []
relationship_values = []

# Tk widgets / state
hotel_combo = None
status_var = None
brand_filter_var = None
region_filter_var = None
country_filter_var = None
brand_band_filter_var = None
relationship_filter_var = None
hyatt_year_var = None
hyatt_year_mode_var = None
filtered_tree = None
selected_tree = None
selected_rows = {}
current_filtered_indexes = []

# Column selection vars (set in setup tab)
brand_col_var = None
region_col_var = None
country_col_var = None
country_fallback_col_var = None
city_col_var = None
gm_col_var = None
eng_col_var = None
dof_col_var = None
reg_eng_spec_col_var = None
avp_col_var = None
md_col_var = None
brand_band_col_var = None
hyatt_date_col_var = None
relationship_col_var = None

brand_col_combo = None
region_col_combo = None
country_col_combo = None
country_fallback_combo = None
city_col_combo = None
gm_col_combo = None
eng_col_combo = None
dof_col_combo = None
reg_eng_spec_combo = None
avp_col_combo = None
md_col_combo = None
brand_band_col_combo = None
hyatt_date_col_combo = None
relationship_col_combo = None

# Lookup detail panel state
detail_info_vars = {}
detail_roles_frame = None
detail_checkbox_vars = []
detail_hotel_name = ""
detail_status_var = None
detail_start_email_btn = None


def format_timestamp(path: str) -> str:
    """Return a human friendly timestamp for the given file path."""
    try:
        mod_time = datetime.fromtimestamp(os.path.getmtime(path))
    except (FileNotFoundError, OSError):
        return "Unknown timestamp"
    return mod_time.strftime("%d.%m.%Y %H:%M")


def get_selected_col(var: tk.StringVar | None, allow_none: bool = False) -> str:
    if var is None:
        return ""
    val = var.get().strip()
    if val == "None":
        return ""
    return val


def get_brand_col():
    return get_selected_col(brand_col_var)


def get_region_col():
    return get_selected_col(region_col_var)


def get_city_col():
    return get_selected_col(city_col_var)


def get_country_col():
    return get_selected_col(country_col_var)


def get_country_fallback_col():
    return get_selected_col(country_fallback_col_var, allow_none=True)


def get_brand_band_col():
    return get_selected_col(brand_band_col_var, allow_none=True)

def get_relationship_col():
    return get_selected_col(relationship_col_var, allow_none=True)


def get_gm_col():
    return get_selected_col(gm_col_var)


def get_eng_col():
    return get_selected_col(eng_col_var)


def get_dof_col():
    return get_selected_col(dof_col_var)


def get_reg_eng_spec_col():
    return get_selected_col(reg_eng_spec_col_var, allow_none=True)


def get_avp_col():
    return get_selected_col(avp_col_var, allow_none=True)


def get_md_col():
    return get_selected_col(md_col_var, allow_none=True)


def get_hyatt_date_col():
    return get_selected_col(hyatt_date_col_var, allow_none=True)


def get_country_value(row: pd.Series) -> str:
    """Return the country/area value from the configured columns."""
    country_col = get_country_col()
    fallback_col = get_country_fallback_col()
    if country_col and country_col in row and pd.notna(row[country_col]):
        return str(row[country_col])
    if fallback_col and fallback_col in row and pd.notna(row[fallback_col]):
        return str(row[fallback_col])
    return ""


def get_city_value(row: pd.Series) -> str:
    city_col = get_city_col()
    if city_col and city_col in row and pd.notna(row[city_col]):
        return str(row[city_col])
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


def ensure_var_in_columns(var: tk.StringVar, preferred_order: list[str], allow_none: bool = False):
    """Ensure a column selection variable is set to an available column or None."""
    if var is None:
        return
    cols = list(df.columns)
    current = var.get()
    if current in cols or (allow_none and current == "None"):
        return
    for candidate in preferred_order:
        if candidate and candidate in cols:
            var.set(candidate)
            return
    if allow_none:
        var.set("None")
    elif cols:
        var.set(cols[0])
    else:
        var.set("")


def refresh_setup_tab_options():
    """Populate setup tab combos with current dataframe columns and keep selections valid."""
    cols = sorted([c for c in df.columns]) if not df.empty else []
    default_list = cols or [""]

    for combo in [
        brand_col_combo,
        region_col_combo,
        city_col_combo,
        gm_col_combo,
        eng_col_combo,
        dof_col_combo,
        avp_col_combo,
        md_col_combo,
        brand_band_col_combo,
        hyatt_date_col_combo,
        relationship_col_combo,
    ]:
        if combo is not None:
            combo["values"] = ["None"] + default_list

    if country_col_combo is not None:
        country_col_combo["values"] = ["None"] + default_list
    if country_fallback_combo is not None:
        country_fallback_combo["values"] = ["None"] + default_list
    if reg_eng_spec_combo is not None:
        reg_eng_spec_combo["values"] = ["None"] + default_list

    ensure_var_in_columns(brand_col_var, [DEFAULT_BRAND_COL], allow_none=True)
    ensure_var_in_columns(region_col_var, [DEFAULT_REGION_COL], allow_none=True)
    ensure_var_in_columns(city_col_var, [DEFAULT_CITY_COL], allow_none=True)
    ensure_var_in_columns(brand_band_col_var, [DEFAULT_BRAND_BAND_COL], allow_none=True)
    ensure_var_in_columns(relationship_col_var, [DEFAULT_RELATIONSHIP_COL], allow_none=True)
    ensure_var_in_columns(country_col_var, [DEFAULT_COUNTRY_COL, DEFAULT_COUNTRY_FALLBACK_COL], allow_none=True)
    ensure_var_in_columns(country_fallback_col_var, [DEFAULT_COUNTRY_FALLBACK_COL], allow_none=True)
    ensure_var_in_columns(hyatt_date_col_var, [DEFAULT_HYATT_DATE_COL], allow_none=True)
    ensure_var_in_columns(gm_col_var, [DEFAULT_GM_COL, "GM"], allow_none=True)  # include old fallback GM
    ensure_var_in_columns(eng_col_var, [DEFAULT_ENG_COL, "Engineering Director"], allow_none=True)  # include old fallback Engineering Director
    ensure_var_in_columns(dof_col_var, [DEFAULT_DOF_COL], allow_none=True)
    ensure_var_in_columns(avp_col_var, ["AVP of Ops", "AVP of Ops-managed"], allow_none=True)
    ensure_var_in_columns(md_col_var, ["SVP / Managing Director", "SVP"], allow_none=True)
    ensure_var_in_columns(reg_eng_spec_col_var, [DEFAULT_REG_ENG_SPEC_COL], allow_none=True)


def apply_column_settings():
    """Apply column selections to filters and refresh views."""
    update_filter_options()
    refresh_filtered_hotels()
    update_selected_tree()


def update_filter_options():
    """Populate filter dropdowns based on loaded data and chosen columns."""
    global brand_values, region_values, country_values, brand_band_values, relationship_values
    brand_col = get_brand_col()
    region_col = get_region_col()
    country_col = get_country_col() or get_country_fallback_col()
    brand_band_col = get_brand_band_col()
    relationship_col = get_relationship_col()

    if df.empty:
        brand_values = []
        region_values = []
        country_values = []
        brand_band_values = []
        relationship_values = []
    else:
        brand_values = sorted(df[brand_col].dropna().astype(str).unique().tolist()) if brand_col in df.columns else []
        region_values = sorted(df[region_col].dropna().astype(str).unique().tolist()) if region_col in df.columns else []
        if country_col and country_col in df.columns:
            country_values = sorted(df[country_col].dropna().astype(str).unique().tolist())
        else:
            country_values = []
        brand_band_values = (
            sorted(df[brand_band_col].dropna().astype(str).unique().tolist()) if brand_band_col in df.columns else []
        )
        relationship_values = (
            sorted(df[relationship_col].dropna().astype(str).unique().tolist()) if relationship_col in df.columns else []
        )

    if brand_filter_var is not None:
        brand_filter_var.set("Any")
    if region_filter_var is not None:
        region_filter_var.set("Any")
    if country_filter_var is not None:
        country_filter_var.set("Any")
    if brand_band_filter_var is not None:
        brand_band_filter_var.set("Any")
    if relationship_filter_var is not None:
        relationship_filter_var.set("Any")

    if 'brand_combo' in globals() and brand_combo is not None:
        brand_combo["values"] = ["Any"] + brand_values
    if 'region_combo' in globals() and region_combo is not None:
        region_combo["values"] = ["Any"] + region_values
    if 'country_combo' in globals() and country_combo is not None:
        country_combo["values"] = ["Any"] + country_values
    if 'brand_band_combo' in globals() and brand_band_combo is not None:
        brand_band_combo["values"] = ["Any"] + brand_band_values
    if 'relationship_combo' in globals() and relationship_combo is not None:
        relationship_combo["values"] = ["Any"] + relationship_values

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

    refresh_setup_tab_options()
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
# Lookup detail panel helpers
# ---------------------------------------------------------------------------
def init_detail_panel(parent):
    """Build the detail panel on the lookup tab."""
    global detail_info_vars, detail_roles_frame, detail_status_var, detail_start_email_btn

    detail_frame = ttk.LabelFrame(parent, text="Hotel Details", padding=10)
    detail_frame.pack(fill="both", expand=True)

    info_grid = ttk.Frame(detail_frame)
    info_grid.pack(fill="x", pady=(0, 8))

    fields = [
        "Spirit Code",
        "Hotel",
        "City",
        "Relationship",
        "Brand",
        "Brand Band",
        "Region",
        "Country/Area",
    ]
    detail_info_vars = {name: tk.StringVar(value="") for name in fields}

    for idx, name in enumerate(fields):
        ttk.Label(info_grid, text=f"{name}:").grid(row=idx, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(info_grid, textvariable=detail_info_vars[name], width=35).grid(
            row=idx, column=1, sticky="w", padx=4, pady=2
        )

    detail_roles_frame = ttk.LabelFrame(detail_frame, text="Recipients", padding=8)
    detail_roles_frame.pack(fill="both", expand=True, pady=(0, 8))

    actions = ttk.Frame(detail_frame)
    actions.pack(fill="x")

    detail_start_email_btn = ttk.Button(
        actions, text="Start Email", command=lambda: draft_email_single(detail_checkbox_vars, detail_hotel_name)
    )
    detail_start_email_btn.pack(side="right")

    detail_status_var = tk.StringVar(value="Select a hotel to view details.")
    ttk.Label(detail_frame, textvariable=detail_status_var, foreground="gray").pack(anchor="w", pady=(4, 0))


def clear_detail_panel(message: str = "Select a hotel to view details."):
    """Reset detail panel contents."""
    global detail_checkbox_vars, detail_hotel_name
    detail_checkbox_vars = []
    detail_hotel_name = ""
    for var in detail_info_vars.values():
        var.set("")
    for widget in detail_roles_frame.winfo_children():
        widget.destroy()
    ttk.Label(detail_roles_frame, text="No recipients available.", foreground="gray").pack(anchor="w")
    if detail_status_var is not None:
        detail_status_var.set(message)


def populate_detail_panel(row: pd.Series):
    """Fill detail panel with hotel info and role checkboxes."""
    global detail_checkbox_vars, detail_hotel_name
    detail_checkbox_vars = []
    detail_hotel_name = row.get("Hotel", "N/A")

    if detail_status_var is not None:
        detail_status_var.set(f"Details loaded for: {detail_hotel_name}")

    city_val = get_city_value(row)
    relationship_col = get_relationship_col()
    brand_col = get_brand_col()
    brand_band_col = get_brand_band_col()
    region_col = get_region_col()

    detail_info_vars["Spirit Code"].set(row.get("Spirit Code", ""))
    detail_info_vars["Hotel"].set(detail_hotel_name)
    detail_info_vars["City"].set(city_val)
    detail_info_vars["Relationship"].set(row.get(relationship_col, "") if relationship_col in row else "")
    detail_info_vars["Brand"].set(row.get(brand_col, "") if brand_col in row else "")
    detail_info_vars["Brand Band"].set(row.get(brand_band_col, "") if brand_band_col in row else "")
    detail_info_vars["Region"].set(row.get(region_col, "") if region_col in row else "")
    detail_info_vars["Country/Area"].set(get_country_value(row))

    for widget in detail_roles_frame.winfo_children():
        widget.destroy()

    roles_to_checkbox = {}
    if get_avp_col():
        roles_to_checkbox["AVP"] = get_avp_col()
    if get_md_col():
        roles_to_checkbox["MD"] = get_md_col()
    if get_gm_col():
        roles_to_checkbox["GM"] = get_gm_col()
    if get_eng_col():
        roles_to_checkbox["Engineering"] = get_eng_col()
    if get_dof_col():
        roles_to_checkbox["DOF"] = get_dof_col()
    if get_reg_eng_spec_col():
        roles_to_checkbox["Regional Eng Specialist"] = get_reg_eng_spec_col()

    if not roles_to_checkbox:
        ttk.Label(detail_roles_frame, text="No role columns configured.", foreground="gray").pack(anchor="w")
    else:
        for role, email_col in roles_to_checkbox.items():
            email_address = row.get(email_col)
            if email_col in row.index and pd.notna(email_address):
                var = tk.BooleanVar()
                chk = ttk.Checkbutton(detail_roles_frame, text=f"{role}: {email_address}", variable=var)
                chk.pack(anchor="w", pady=1)
                detail_checkbox_vars.append((var, str(email_address)))
            else:
                ttk.Label(detail_roles_frame, text=f"{role}: N/A (Email not found)", foreground="gray").pack(anchor="w")

# ---------------------------------------------------------------------------
# Multi-select helpers
# ---------------------------------------------------------------------------
def filtered_dataframe():
    """Return dataframe filtered by current dropdown selections."""
    if df.empty:
        return pd.DataFrame()

    filt = df
    brand_col = get_brand_col()
    region_col = get_region_col()
    country_col = get_country_col() or get_country_fallback_col()
    brand_band_col = get_brand_band_col()
    hyatt_col = get_hyatt_date_col()
    relationship_col = get_relationship_col()

    if brand_filter_var is not None:
        brand_val = brand_filter_var.get()
        if brand_val and brand_val != "Any" and brand_col in filt.columns:
            filt = filt[filt[brand_col].astype(str) == brand_val]
    if region_filter_var is not None:
        region_val = region_filter_var.get()
        if region_val and region_val != "Any" and region_col in filt.columns:
            filt = filt[filt[region_col].astype(str) == region_val]
    if country_filter_var is not None:
        country_val = country_filter_var.get()
        if country_val and country_val != "Any" and country_col in filt.columns:
            filt = filt[filt[country_col].astype(str) == country_val]
    if brand_band_filter_var is not None:
        brand_band_val = brand_band_filter_var.get()
        if brand_band_val and brand_band_val != "Any" and brand_band_col in filt.columns:
            filt = filt[filt[brand_band_col].astype(str) == brand_band_val]
    if relationship_filter_var is not None:
        rel_val = relationship_filter_var.get()
        if rel_val and rel_val != "Any" and relationship_col in filt.columns:
            filt = filt[filt[relationship_col].astype(str) == rel_val]

    # Hyatt date filter (year with before/after/on)
    if hyatt_col and hyatt_col in filt.columns and hyatt_year_mode_var is not None and hyatt_year_var is not None:
        mode = hyatt_year_mode_var.get()
        year_str = hyatt_year_var.get().strip()
        if mode and mode != "Any" and year_str.isdigit():
            target_year = int(year_str)
            years = pd.to_datetime(filt[hyatt_col], errors="coerce").dt.year
            if mode == "Before":
                filt = filt[years.notna() & (years < target_year)]
            elif mode == "Before/Equal":
                filt = filt[years.notna() & (years <= target_year)]
            elif mode == "Equal":
                filt = filt[years.notna() & (years == target_year)]
            elif mode == "After/Equal":
                filt = filt[years.notna() & (years >= target_year)]
            elif mode == "After":
                filt = filt[years.notna() & (years > target_year)]
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

    brand_col = get_brand_col()
    region_col = get_region_col()
    brand_band_col = get_brand_band_col()
    relationship_col = get_relationship_col()

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
                get_city_value(row),
                row.get(brand_col, "") if brand_col in row else "",
                row.get(brand_band_col, "") if brand_band_col in row else "",
                row.get(relationship_col, "") if relationship_col in row else "",
                row.get(region_col, "") if region_col in row else "",
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

    brand_col = get_brand_col()
    region_col = get_region_col()
    brand_band_col = get_brand_band_col()
    relationship_col = get_relationship_col()

    for row_idx, row in selected_rows.items():
        selected_tree.insert(
            "",
            "end",
            iid=str(row_idx),
            values=(
                row.get("Spirit Code", ""),
                row.get("Hotel", ""),
                get_city_value(row),
                row.get(brand_col, "") if brand_col in row else "",
                row.get(brand_band_col, "") if brand_band_col in row else "",
                row.get(relationship_col, "") if relationship_col in row else "",
                row.get(region_col, "") if region_col in row else "",
                get_country_value(row),
            ),
        )


def get_role_addresses(row: pd.Series, role_key: str):
    """Return a list of email addresses for the chosen role."""
    role_map = {
        "AVP": [get_avp_col()],
        "MD": [get_md_col()],
        "GM": [get_gm_col()],
        "Engineering": [get_eng_col()],
        "DOF": [get_dof_col()],
        "RegionalEngineeringSpecialist": [get_reg_eng_spec_col()],
    }
    emails = []
    for col in role_map.get(role_key, []):
        if col and col in row and pd.notna(row[col]):
            emails.append(str(row[col]))
    return emails


def draft_emails_for_selection(avp_var, md_var, gm_var, eng_var, dof_var, reg_eng_var):
    """Create Outlook draft emails for the selected hotels and roles."""
    if not selected_rows:
        messagebox.showinfo("Keine Hotels", "Bitte waehlen Sie mindestens ein Hotel aus der Auswahl aus.")
        return

    chosen_roles = []
    if avp_var.get():
        chosen_roles.append("AVP")
    if md_var.get():
        chosen_roles.append("MD")
    if gm_var.get():
        chosen_roles.append("GM")
    if eng_var.get():
        chosen_roles.append("Engineering")
    if dof_var.get():
        chosen_roles.append("DOF")
    if reg_eng_var.get():
        chosen_roles.append("RegionalEngineeringSpecialist")

    if not chosen_roles:
        messagebox.showinfo("Keine Rollen", "Bitte waehlen Sie mindestens eine Empfaengerrolle.")
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
    brand_col = get_brand_col()
    region_col = get_region_col()

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
                f"City: {get_city_value(row)}",
                f"Relationship: {row.get(get_relationship_col(), '') if get_relationship_col() in row else ''}",
                f"Brand: {row.get(brand_col, '') if brand_col in row else ''}",
                f"Brand Band: {row.get(get_brand_band_col(), '') if get_brand_band_col() in row else ''}",
                f"Region: {row.get(region_col, '') if region_col in row else ''}",
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
            "Keine Empfaenger",
            "Fuer folgende Hotels wurden keine E-Mail-Adressen in den gewaehlten Rollen gefunden:\n" + "\n".join(missing_addresses),
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
        clear_detail_panel("No data loaded.")
        return

    spirit = spirit_entry.get().strip()
    hotel = hotel_var_local.get().strip()
    city_col = get_city_col()

    if spirit:
        result = df[df["Spirit Code"].astype(str).str.lower() == spirit.lower()]
    elif hotel:
        mask = df["Hotel"].astype(str).str.contains(hotel, case=False, na=False)
        if city_col and city_col in df.columns:
            mask |= df[city_col].astype(str).str.contains(hotel, case=False, na=False)
        result = df[mask]
    else:
        messagebox.showwarning("Whoops", "Enter Spirit Code or pick a hotel.")
        clear_detail_panel("Enter Spirit Code or hotel to view details.")
        return

    if result.empty:
        messagebox.showinfo("Nada", "No matching hotel found.")
        clear_detail_panel("No matching hotel found.")
        return

    if len(result) == 1:
        row = result.iloc[0]
        populate_detail_panel(row)
    else:
        show_search_results(result)


def show_details_gui(row):
    win = tk.Toplevel(root)
    win.title(f"Details for {row.get('Hotel', 'N/A')}")
    win.geometry("700x760")
    win.minsize(500, 400)

    info_frame = ttk.LabelFrame(win, text="Hotel Information", padding="10")
    info_frame.pack(padx=10, pady=10, fill="x")

    general_info = [
        ("Spirit Code", "Spirit Code"),
        ("Hotel", "Hotel"),
        ("City", get_city_col()),
        ("Country/Area", get_country_col() or get_country_fallback_col()),
        ("Relationship", "Relationship"),
        ("Rooms", "Rooms"),
        ("JV", "JV"),
        ("JV Percent", "JV Percent"),
        ("Owner", "Owner"),
    ]

    row_idx_info = 0
    for label_text, col in general_info:
        if col and col in row and pd.notna(row[col]):
            tk.Label(info_frame, text=f"{label_text}:", anchor="w", font=("Arial", 10, "bold")).grid(
                row=row_idx_info, column=0, sticky="w", padx=5, pady=2
            )
            tk.Label(info_frame, text=row[col], anchor="w", font=("Arial", 10)).grid(
                row=row_idx_info, column=1, sticky="w", padx=5, pady=2
            )
            row_idx_info += 1

    roles_frame = ttk.LabelFrame(win, text="Key Personnel (Select for Email)", padding="10")
    roles_frame.pack(padx=10, pady=10, fill="both", expand=False)

    roles_to_checkbox = {}
    if get_avp_col():
        roles_to_checkbox["AVP"] = get_avp_col()
    if get_md_col():
        roles_to_checkbox["MD"] = get_md_col()
    if get_gm_col():
        roles_to_checkbox["GM"] = get_gm_col()
    if get_eng_col():
        roles_to_checkbox["Engineering"] = get_eng_col()
    if get_dof_col():
        roles_to_checkbox["DOF"] = get_dof_col()
    if get_reg_eng_spec_col():
        roles_to_checkbox["Regional Eng Specialist"] = get_reg_eng_spec_col()

    checkbox_vars = []
    row_idx = 0
    if not roles_to_checkbox:
        tk.Label(roles_frame, text="No role columns configured.", anchor="w", foreground="gray").grid(
            row=row_idx, column=0, sticky="w", padx=5, pady=2
        )
    else:
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


def draft_email_single(checkbox_vars, hotel_name, details_window=None):
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
    if details_window is not None:
        details_window.destroy()


def show_search_results(results_df):
    win = tk.Toplevel(root)
    win.title("Suchergebnisse")
    win.geometry("500x300")

    tree = ttk.Treeview(win, columns=("Spirit Code", "Hotel", "City"), show="headings")
    tree.heading("Spirit Code", text="Spirit Code")
    tree.heading("Hotel", text="Hotel")

    tree.heading("City", text="City")

    tree.column("Spirit Code", width=100)
    tree.column("Hotel", width=250)
    tree.column("City", width=130)
    tree.pack(fill="both", expand=True, padx=10, pady=10)

    for _, result_row in results_df.iterrows():
        tree.insert("", "end", values=(result_row.get("Spirit Code", ""), result_row.get("Hotel", ""), get_city_value(result_row)))

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

        populate_detail_panel(results_df.iloc[row_index])
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
root.geometry("1150x780")

status_var = tk.StringVar(value="Lade Daten ...")

# Initialize column selection vars
brand_col_var = tk.StringVar(value=DEFAULT_BRAND_COL)
region_col_var = tk.StringVar(value=DEFAULT_REGION_COL)
country_col_var = tk.StringVar(value=DEFAULT_COUNTRY_COL)
country_fallback_col_var = tk.StringVar(value=DEFAULT_COUNTRY_FALLBACK_COL)
city_col_var = tk.StringVar(value=DEFAULT_CITY_COL)
brand_band_col_var = tk.StringVar(value=DEFAULT_BRAND_BAND_COL)
relationship_col_var = tk.StringVar(value=DEFAULT_RELATIONSHIP_COL)
hyatt_date_col_var = tk.StringVar(value=DEFAULT_HYATT_DATE_COL)
gm_col_var = tk.StringVar(value=DEFAULT_GM_COL)
eng_col_var = tk.StringVar(value=DEFAULT_ENG_COL)
dof_col_var = tk.StringVar(value=DEFAULT_DOF_COL)
avp_col_var = tk.StringVar(value="AVP of Ops")
md_col_var = tk.StringVar(value="SVP / Managing Director")
reg_eng_spec_col_var = tk.StringVar(value="None")

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
lookup_frame.columnconfigure(1, weight=1)
lookup_frame.rowconfigure(0, weight=1)

lookup_form = ttk.Frame(lookup_frame)
lookup_form.grid(row=0, column=0, sticky="nw", padx=(0, 10))

spirit_label = tk.Label(lookup_form, text="Spirit Code:")
spirit_label.grid(row=0, column=0, sticky="e", padx=5, pady=5)
spirit_entry = tk.Entry(lookup_form, width=30)
spirit_entry.grid(row=0, column=1, padx=5, pady=5)

hotel_label = tk.Label(lookup_form, text="Hotel:")
hotel_label.grid(row=1, column=0, sticky="e", padx=5, pady=5)
hotel_var = tk.StringVar()
hotel_combo = ttk.Combobox(lookup_form, textvariable=hotel_var, values=hotel_names)
hotel_combo.grid(row=1, column=1, padx=5, pady=5)
hotel_combo.state(["!readonly"])


def on_hotel_keyrelease(event):
    val = hotel_var.get()
    hotel_combo["values"] = hotel_names if not val else [h for h in hotel_names if val.lower() in h.lower()]


hotel_combo.bind("<KeyRelease>", on_hotel_keyrelease)

search_button = tk.Button(lookup_form, text="Search", command=lambda: lookup(spirit_entry, hotel_var))
search_button.grid(row=2, column=0, columnspan=2, pady=10)

detail_container = ttk.Frame(lookup_frame)
detail_container.grid(row=0, column=1, sticky="nsew")
init_detail_panel(detail_container)
clear_detail_panel()

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
brand_band_filter_var = tk.StringVar(value="Any")
relationship_filter_var = tk.StringVar(value="Any")
hyatt_year_var = tk.StringVar(value="")
hyatt_year_mode_var = tk.StringVar(value="Any")

row_f = 0
if True:
    ttk.Label(filters_frame, text="Brand").grid(row=row_f, column=0, sticky="w", padx=5, pady=2)
    brand_combo = ttk.Combobox(filters_frame, textvariable=brand_filter_var, values=["Any"], state="readonly")
    brand_combo.grid(row=row_f, column=1, sticky="ew", padx=5, pady=2)

    ttk.Label(filters_frame, text="Brand Band").grid(row=row_f, column=2, sticky="w", padx=5, pady=2)
    brand_band_combo = ttk.Combobox(filters_frame, textvariable=brand_band_filter_var, values=["Any"], state="readonly")
    brand_band_combo.grid(row=row_f, column=3, sticky="ew", padx=5, pady=2)

    ttk.Label(filters_frame, text="Region").grid(row=row_f, column=4, sticky="w", padx=5, pady=2)
    region_combo = ttk.Combobox(filters_frame, textvariable=region_filter_var, values=["Any"], state="readonly")
    region_combo.grid(row=row_f, column=5, sticky="ew", padx=5, pady=2)

    ttk.Label(filters_frame, text="Relationship").grid(row=row_f, column=6, sticky="w", padx=5, pady=2)
    relationship_combo = ttk.Combobox(filters_frame, textvariable=relationship_filter_var, values=["Any"], state="readonly")
    relationship_combo.grid(row=row_f, column=7, sticky="ew", padx=5, pady=2)

    ttk.Label(filters_frame, text="Country/Area").grid(row=row_f, column=8, sticky="w", padx=5, pady=2)
    country_combo = ttk.Combobox(filters_frame, textvariable=country_filter_var, values=["Any"], state="readonly")
    country_combo.grid(row=row_f, column=9, sticky="ew", padx=5, pady=2)

    ttk.Label(filters_frame, text="Hyatt Date (year)").grid(row=row_f, column=10, sticky="w", padx=5, pady=2)
    hyatt_year_entry = ttk.Entry(filters_frame, textvariable=hyatt_year_var, width=8)
    hyatt_year_entry.grid(row=row_f, column=11, sticky="w", padx=5, pady=2)

    hyatt_mode_combo = ttk.Combobox(
        filters_frame,
        textvariable=hyatt_year_mode_var,
        values=["Any", "Before", "Before/Equal", "Equal", "After/Equal", "After"],
        state="readonly",
        width=12,
    )
    hyatt_mode_combo.grid(row=row_f, column=12, sticky="w", padx=5, pady=2)

    filters_frame.columnconfigure(1, weight=1)
    filters_frame.columnconfigure(3, weight=1)
    filters_frame.columnconfigure(5, weight=1)
    filters_frame.columnconfigure(7, weight=1)
    filters_frame.columnconfigure(9, weight=1)

apply_filter_btn = ttk.Button(filters_frame, text="Apply Filter", command=refresh_filtered_hotels)
apply_filter_btn.grid(row=0, column=13, sticky="e", padx=8, pady=2)

lists_frame = ttk.Frame(multi_frame)
lists_frame.pack(fill="both", expand=True, padx=5, pady=5)

filtered_frame = ttk.LabelFrame(lists_frame, text="Gefilterte Hotels", padding=5)
filtered_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

filtered_tree = ttk.Treeview(
    filtered_frame,
    columns=("Spirit", "Hotel", "City", "Brand", "Brand Band", "Relationship", "Region", "Country"),
    show="headings",
    selectmode="extended",
)
for col, width in [
    ("Spirit", 80),
    ("Hotel", 220),
    ("City", 120),
    ("Brand", 120),
    ("Brand Band", 120),
    ("Relationship", 120),
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
    columns=("Spirit", "Hotel", "City", "Brand", "Brand Band", "Relationship", "Region", "Country"),
    show="headings",
    selectmode="extended",
)
for col, width in [
    ("Spirit", 80),
    ("Hotel", 220),
    ("City", 120),
    ("Brand", 120),
    ("Brand Band", 120),
    ("Relationship", 120),
    ("Region", 120),
    ("Country", 140),
]:
    selected_tree.heading(col, text=col)
    selected_tree.column(col, width=width, stretch=True)
selected_tree.pack(fill="both", expand=True)

roles_frame = ttk.LabelFrame(multi_frame, text="Empfaengerrollen", padding=10)
roles_frame.pack(fill="x", padx=5, pady=5)

avp_var = tk.BooleanVar(value=False)
md_var = tk.BooleanVar(value=False)
gm_var = tk.BooleanVar(value=True)
eng_var = tk.BooleanVar(value=False)
dof_var = tk.BooleanVar(value=False)
reg_eng_var = tk.BooleanVar(value=False)

roles_label = ttk.Label(roles_frame, text="Wen anschreiben?")
roles_label.pack(side="left", padx=5)

ttk.Checkbutton(roles_frame, text="AVP", variable=avp_var).pack(side="left", padx=5)

ttk.Checkbutton(roles_frame, text="MD", variable=md_var).pack(side="left", padx=5)

ttk.Checkbutton(roles_frame, text="GM", variable=gm_var).pack(side="left", padx=5)

ttk.Checkbutton(roles_frame, text="Engineering", variable=eng_var).pack(side="left", padx=5)

ttk.Checkbutton(roles_frame, text="DOF", variable=dof_var).pack(side="left", padx=5)

ttk.Checkbutton(roles_frame, text="Regional Eng Specialist", variable=reg_eng_var).pack(side="left", padx=5)


draft_btn = ttk.Button(
    roles_frame,
    text="Draft Emails in Outlook",
    command=lambda: draft_emails_for_selection(avp_var, md_var, gm_var, eng_var, dof_var, reg_eng_var),
)
draft_btn.pack(side="right", padx=10)

# ---------------------------------------------------------------------------
# Tab 3: Setup
# ---------------------------------------------------------------------------
setup_frame = ttk.Frame(notebook, padding=10)
notebook.add(setup_frame, text="Setup")

setup_top = ttk.LabelFrame(setup_frame, text="Data Columns", padding=10)
setup_top.pack(fill="x", padx=5, pady=5)

brand_col_combo = ttk.Combobox(setup_top, textvariable=brand_col_var, state="readonly")
region_col_combo = ttk.Combobox(setup_top, textvariable=region_col_var, state="readonly")
country_col_combo = ttk.Combobox(setup_top, textvariable=country_col_var, state="readonly")
country_fallback_combo = ttk.Combobox(setup_top, textvariable=country_fallback_col_var, state="readonly")
city_col_combo = ttk.Combobox(setup_top, textvariable=city_col_var, state="readonly")
brand_band_col_combo = ttk.Combobox(setup_top, textvariable=brand_band_col_var, state="readonly")
relationship_col_combo = ttk.Combobox(setup_top, textvariable=relationship_col_var, state="readonly")
hyatt_date_col_combo = ttk.Combobox(setup_top, textvariable=hyatt_date_col_var, state="readonly")

row_setup = 0
labels = [
    ("Brand column", brand_col_combo),
    ("Brand Band column", brand_band_col_combo),
    ("Region column", region_col_combo),
    ("Country column", country_col_combo),
    ("Country fallback (optional)", country_fallback_combo),
    ("City column", city_col_combo),
    ("Relationship column", relationship_col_combo),
    ("Hyatt Date column (for year filter)", hyatt_date_col_combo),
]
for idx, (text, combo) in enumerate(labels):
    ttk.Label(setup_top, text=text).grid(row=row_setup + idx, column=0, sticky="w", padx=5, pady=2)
    combo.grid(row=row_setup + idx, column=1, sticky="ew", padx=5, pady=2)
setup_top.columnconfigure(1, weight=1)

roles_setup = ttk.LabelFrame(setup_frame, text="Recipient Columns", padding=10)
roles_setup.pack(fill="x", padx=5, pady=5)

avp_col_combo = ttk.Combobox(roles_setup, textvariable=avp_col_var, state="readonly")
md_col_combo = ttk.Combobox(roles_setup, textvariable=md_col_var, state="readonly")
gm_col_combo = ttk.Combobox(roles_setup, textvariable=gm_col_var, state="readonly")
eng_col_combo = ttk.Combobox(roles_setup, textvariable=eng_col_var, state="readonly")
dof_col_combo = ttk.Combobox(roles_setup, textvariable=dof_col_var, state="readonly")
reg_eng_spec_combo = ttk.Combobox(roles_setup, textvariable=reg_eng_spec_col_var, state="readonly")

labels_roles = [
    ("AVP column", avp_col_combo),
    ("Managing Director column", md_col_combo),
    ("GM column", gm_col_combo),
    ("Engineering column", eng_col_combo),
    ("DOF column", dof_col_combo),
    ("Regional Eng Specialist column (optional)", reg_eng_spec_combo),
]
for idx, (text, combo) in enumerate(labels_roles):
    ttk.Label(roles_setup, text=text).grid(row=idx, column=0, sticky="w", padx=5, pady=2)
    combo.grid(row=idx, column=1, sticky="ew", padx=5, pady=2)
roles_setup.columnconfigure(1, weight=1)

apply_columns_btn = ttk.Button(setup_frame, text="Apply column mapping", command=apply_column_settings)
apply_columns_btn.pack(anchor="e", padx=5, pady=10)

# Status bar
status_label = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor="w")
status_label.pack(fill="x", side="bottom")

# Load default data after UI widgets are ready
ensure_initial_data()

# Populate setup dropdown values after data load
refresh_setup_tab_options()

# Populate filter dropdown values after data load
brand_combo["values"] = ["Any"] + brand_values
region_combo["values"] = ["Any"] + region_values
country_combo["values"] = ["Any"] + country_values
brand_band_combo["values"] = ["Any"] + brand_band_values
relationship_combo["values"] = ["Any"] + relationship_values

# Initial filtered view
refresh_filtered_hotels()

# Warm Outlook in the background so the first email opens faster
root.after(200, warm_outlook_app)

root.mainloop()
