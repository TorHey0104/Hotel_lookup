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
import json
import importlib.util
import glob
import html
import math
from datetime import date

# Cache Outlook availability and instance so email drafting is faster after the first use
WIN32COM_AVAILABLE = os.name == "nt" and importlib.util.find_spec("win32com.client") is not None
_outlook_app = None

# ---------------------------------------------------------------------------
# CONFIGURE THIS
# ---------------------------------------------------------------------------
DATA_DIR = r"C:\Users\4612135\OneDrive - Hyatt Hotels\___DATA"
FILE_NAME = "2a Hotels one line hotel.xlsx"
DEFAULT_FILE_PATH = os.path.join(DATA_DIR, FILE_NAME)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "hyatt_logo.png")  # optional logo next to script

TOOL_NAME = "Hyatt EAME Hotel Lookup and Multi E-Mail Tool"
VERSION = "4.2.2"
VERSION_DATE = date.today().strftime("%d.%m.%Y")

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
brand_filter_var = None  # kept for legacy; not used with multiselect
region_filter_var = None  # kept for legacy; not used with multiselect
country_filter_var = None  # kept for legacy; not used with multiselect
brand_band_filter_var = None  # kept for legacy; not used with multiselect
relationship_filter_var = None  # kept for legacy; not used with multiselect
hyatt_year_var = None
hyatt_year_mode_var = None
brand_listbox = None
brand_band_listbox = None
region_listbox = None
relationship_listbox = None
country_listbox = None
filtered_tree = None
selected_tree = None
selected_rows = {}
current_filtered_indexes = []
role_send_vars = {}
ROLE_MODES = ["Skip", "To", "CC", "BCC"]
style = None

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
detail_row_current = None
signatures_cache = {}
splash_win = None
splash_status_var = None
splash_file_var = None
splash_logo_img = None

# Visible columns for filtered hotels
MANDATORY_FILTER_COLS = ["Spirit Code", "Hotel"]
visible_optional_filter_cols = ["City", "Brand", "Brand Band", "Relationship", "Region", "Country"]
filter_cols_listbox = None


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


def normalize_emails(raw: str):
    """Split a raw email string by common delimiters and drop placeholders like N/A."""
    parts = []
    for chunk in str(raw).replace(",", ";").split(";"):
        email = chunk.strip()
        if not email:
            continue
        low = email.lower()
        if low in {"n/a", "na", "none"}:
            continue
        parts.append(email)
    return parts


def render_template(row: pd.Series, template: str) -> str:
    """Replace placeholders in a template string using row values."""
    brand_col = get_brand_col()
    region_col = get_region_col()
    relationship_col = get_relationship_col()
    brand_band_col = get_brand_band_col()
    replacements = {
        "{hotel}": row.get("Hotel", ""),
        "{spirit_code}": row.get("Spirit Code", ""),
        "{city}": get_city_value(row),
        "{relationship}": row.get(relationship_col, "") if relationship_col in row else "",
        "{brand}": row.get(brand_col, "") if brand_col in row else "",
        "{brand_band}": row.get(brand_band_col, "") if brand_band_col in row else "",
        "{region}": row.get(region_col, "") if region_col in row else "",
        "{country}": get_country_value(row),
        "{owner}": row.get("Owner", ""),
        "{rooms}": row.get("Rooms", ""),
    }
    rendered = template
    for key, val in replacements.items():
        rendered = rendered.replace(key, str(val))
    return rendered


def load_signatures():
    """Load Outlook signature texts/html from the default signatures folder."""
    global signatures_cache
    if signatures_cache:
        return signatures_cache

    signatures_cache = {"None": {"html": "", "text": ""}}
    sig_dir = os.path.join(os.path.expandvars(r"%APPDATA%"), "Microsoft", "Signatures")
    if not os.path.isdir(sig_dir):
        return signatures_cache

    base_names = set()
    for ext in ("*.txt", "*.htm", "*.html"):
        for path in glob.glob(os.path.join(sig_dir, ext)):
            base_names.add(os.path.splitext(os.path.basename(path))[0])

    for name in base_names:
        txt_path = os.path.join(sig_dir, name + ".txt")
        htm_path = os.path.join(sig_dir, name + ".htm")
        html_path = os.path.join(sig_dir, name + ".html")
        sig_entry = {"html": "", "text": ""}

        if os.path.isfile(htm_path):
            try:
                with open(htm_path, "r", encoding="utf-8", errors="ignore") as fh:
                    sig_entry["html"] = fh.read()
            except Exception:
                sig_entry["html"] = ""
        elif os.path.isfile(html_path):
            try:
                with open(html_path, "r", encoding="utf-8", errors="ignore") as fh:
                    sig_entry["html"] = fh.read()
            except Exception:
                sig_entry["html"] = ""

        if os.path.isfile(txt_path):
            try:
                with open(txt_path, "r", encoding="utf-8", errors="ignore") as fh:
                    sig_entry["text"] = fh.read().strip()
            except Exception:
                sig_entry["text"] = ""

        signatures_cache[name] = sig_entry
    return signatures_cache


def render_with_signature(body_text: str, signature_entry: dict):
    """Return a dict with either html or text combined with signature."""
    if not signature_entry:
        return {"text": body_text}
    sig_html = signature_entry.get("html", "") if isinstance(signature_entry, dict) else ""
    sig_txt = signature_entry.get("text", "") if isinstance(signature_entry, dict) else ""

    if sig_html:
        body_html = html.escape(body_text).replace("\n", "<br>")
        return {"html": f"<div>{body_html}</div><br>{sig_html}"}
    elif sig_txt:
        return {"text": body_text + "\n\n" + sig_txt}
    else:
        return {"text": body_text}


def ensure_style():
    """Configure ttk style accents (e.g., active tab highlighting)."""
    global style
    if style is None:
        style = ttk.Style()
    current_theme = style.theme_use()
    style.theme_use(current_theme)
    # Blue accent for active tab
    style.map("TNotebook.Tab", background=[("selected", "#1f4fa3")], foreground=[("selected", "white")])
    style.configure("TNotebook.Tab", padding=(8, 4))


def show_splash():
    """Show a splash window while loading."""
    global splash_win, splash_status_var, splash_file_var, splash_logo_img
    if splash_win is not None:
        try:
            splash_win.destroy()
        except Exception:
            pass
    splash_win = tk.Toplevel()
    splash_win.overrideredirect(True)
    splash_win.attributes("-topmost", True)
    splash_win.transient(root)

    container = ttk.Frame(splash_win, padding=14, relief="raised", borderwidth=3)
    container.pack(fill="both", expand=True)

    logo_found = False
    logo_path_use = LOGO_PATH if os.path.isfile(LOGO_PATH) else os.path.join(DATA_DIR, "hyatt_logo.png")
    if os.path.isfile(logo_path_use):
        try:
            img_raw = tk.PhotoImage(file=logo_path_use)
            max_width = 400  # keep logo size reasonable
            if img_raw.width() > max_width:
                factor = math.ceil(img_raw.width() / max_width)
                splash_logo_img = img_raw.subsample(factor, factor)
            else:
                splash_logo_img = img_raw
            ttk.Label(container, image=splash_logo_img).pack(anchor="w", pady=(0, 6))
            logo_found = True
        except Exception:
            logo_found = False
    if not logo_found:
        ttk.Label(container, text="HYATT", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 6))

    ttk.Label(container, text=TOOL_NAME, font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(10, 3))
    ttk.Label(container, text=f"Version {VERSION} ({VERSION_DATE})", font=("Segoe UI", 11)).pack(anchor="w")
    ttk.Label(container, text="Author: Torsten Heyorth, Dir Engineering Operations", font=("Segoe UI", 10)).pack(anchor="w")
    ttk.Label(container, text="Created with OpenAI Codex & VS Code", font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 10))

    splash_file_var = tk.StringVar(value="Loading data file...")
    ttk.Label(container, textvariable=splash_file_var, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(4, 2))

    splash_status_var = tk.StringVar(value=f"{TOOL_NAME} v{VERSION} ({VERSION_DATE})")
    ttk.Label(container, textvariable=splash_status_var, foreground="gray", font=("Segoe UI", 10)).pack(anchor="w")

    ttk.Button(container, text="Understood...", command=close_splash).pack(anchor="e", pady=(12, 0), ipadx=8, ipady=3)

    splash_win.update_idletasks()
    desired_w = 720
    desired_h = 420
    ws = splash_win.winfo_screenwidth()
    hs = splash_win.winfo_screenheight()
    x = int((ws / 2) - (desired_w / 2))
    y = int((hs / 2) - (desired_h / 2))
    splash_win.geometry(f"{desired_w}x{desired_h}+{x}+{y}")
    splash_win.config(highlightbackground="#1f4fa3", highlightcolor="#1f4fa3", highlightthickness=2)
    try:
        splash_win.lift()
        splash_win.focus_force()
        splash_win.after(50, splash_win.lift)
        splash_win.after(50, lambda: splash_win.geometry(f"{desired_w}x{desired_h}+{x}+{y}"))
    except Exception:
        pass


def update_splash(file_path: str, status: str):
    if splash_file_var is not None:
        if file_path:
            splash_file_var.set(f"Loaded file: {os.path.basename(file_path)}")
        else:
            splash_file_var.set("Loading data file...")
    if splash_status_var is not None and status:
        splash_status_var.set(status)
    if splash_win is not None:
        splash_win.lift()
        splash_win.after(50, splash_win.lift)


def close_splash():
    global splash_win
    if splash_win is not None:
        try:
            splash_win.destroy()
        except Exception:
            pass
    splash_win = None


def refresh_filter_columns_list():
    """Refresh list of selectable columns for the filtered hotels table."""
    if filter_cols_listbox is None:
        return
    filter_cols_listbox.delete(0, tk.END)
    candidates = []
    if not df.empty:
        candidates = [c for c in sorted(df.columns) if c not in MANDATORY_FILTER_COLS]
    else:
        candidates = [c for c in visible_optional_filter_cols if c not in MANDATORY_FILTER_COLS]
    for col in candidates:
        filter_cols_listbox.insert(tk.END, col)
    # restore selections
    for idx, col in enumerate(candidates):
        if col in visible_optional_filter_cols:
            filter_cols_listbox.selection_set(idx)


def get_filtered_columns():
    """Return ordered list of columns to show in filtered tree (mandatory + selected optional)."""
    selected = []
    if filter_cols_listbox is not None:
        selected = [filter_cols_listbox.get(i) for i in filter_cols_listbox.curselection()]
    if not selected:
        selected = visible_optional_filter_cols
    return MANDATORY_FILTER_COLS + selected


def add_role_selector(parent, role_name, default_mode="Skip"):
    var = tk.StringVar(value=default_mode)
    cb = ttk.Combobox(parent, textvariable=var, values=ROLE_MODES, state="readonly", width=10)
    cb.bind("<<ComboboxSelected>>", lambda e: update_selected_tree())
    role_send_vars[role_name] = var
    row = len(parent.grid_slaves()) // 2
    ttk.Label(parent, text=role_name).grid(row=row, column=0, sticky="w", padx=4, pady=2)
    cb.grid(row=row, column=1, sticky="w", padx=4, pady=2)


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
    update_visible_optional_from_listbox()
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

    def reset_listbox(lb, values):
        if lb is None:
            return
        lb.delete(0, tk.END)
        for val in values:
            lb.insert(tk.END, val)

    reset_listbox(brand_listbox, brand_values)
    reset_listbox(brand_band_listbox, brand_band_values)
    reset_listbox(region_listbox, region_values)
    reset_listbox(relationship_listbox, relationship_values)
    reset_listbox(country_listbox, country_values)

    refresh_filter_columns_list()

    if filtered_tree is not None:
        refresh_filtered_hotels()


def update_visible_optional_from_listbox():
    """Capture selected optional columns for the filtered hotels view."""
    global visible_optional_filter_cols
    if filter_cols_listbox is None:
        return
    selected = [filter_cols_listbox.get(i) for i in filter_cols_listbox.curselection()]
    if selected:
        visible_optional_filter_cols = selected


def reset_filters():
    """Clear all filter selections."""
    for lb in [brand_listbox, brand_band_listbox, region_listbox, relationship_listbox, country_listbox]:
        if lb is not None:
            lb.selection_clear(0, tk.END)
    if hyatt_year_var is not None:
        hyatt_year_var.set("")
    if hyatt_year_mode_var is not None:
        hyatt_year_mode_var.set("Any")
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
    update_splash(path, "Data loaded.")


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
    update_splash("", f"Loading {os.path.basename(file_path)} ...")
    try:
        load_data(file_path)
        update_splash(file_path, "Data loaded.")
    except Exception as exc:
        messagebox.showerror("Laden fehlgeschlagen", f"Die Datei konnte nicht geladen werden:\n{exc}")


def load_config_file():
    """Load configuration JSON (data file path and role routing)."""
    config_path = filedialog.askopenfilename(
        title="Konfiguration laden",
        filetypes=[("JSON files", "*.json"), ("Alle Dateien", "*.*")],
    )
    if not config_path:
        return
    try:
        with open(config_path, "r", encoding="utf-8") as fh:
            cfg = json.load(fh)
    except Exception as exc:
        messagebox.showerror("Konfiguration", f"Konfigurationsdatei konnte nicht gelesen werden:\n{exc}")
        return

    data_path = cfg.get("data_file_path")
    if data_path:
        try:
            load_data(data_path)
        except Exception as exc:
            messagebox.showerror("Konfiguration", f"Datendatei aus Konfiguration konnte nicht geladen werden:\n{exc}")

    cols = cfg.get("columns", {})
    def set_if_present(var, key):
        if var is not None and key in cols and cols.get(key):
            var.set(cols[key])

    set_if_present(brand_col_var, "brand")
    set_if_present(brand_band_col_var, "brand_band")
    set_if_present(region_col_var, "region")
    set_if_present(relationship_col_var, "relationship")
    set_if_present(country_col_var, "country")
    set_if_present(country_fallback_col_var, "country_fallback")
    set_if_present(city_col_var, "city")
    set_if_present(hyatt_date_col_var, "hyatt_date")
    set_if_present(gm_col_var, "gm")
    set_if_present(eng_col_var, "eng")
    set_if_present(dof_col_var, "dof")
    set_if_present(avp_col_var, "avp")
    set_if_present(md_col_var, "md")
    set_if_present(reg_eng_spec_col_var, "reg_eng_spec")

    roles_cfg = cfg.get("roles", {})
    for role, val in roles_cfg.items():
        if role in role_send_vars and val in ROLE_MODES:
            role_send_vars[role].set(val)
    refresh_setup_tab_options()
    apply_column_settings()
    update_selected_tree()

    optional_cols = cfg.get("visible_filter_cols")
    if optional_cols:
        global visible_optional_filter_cols
        visible_optional_filter_cols = optional_cols
        refresh_filter_columns_list()


def save_config_file():
    """Save configuration (data file path, column mapping, role routing) to JSON."""
    update_visible_optional_from_listbox()
    config_path = filedialog.asksaveasfilename(
        title="Konfiguration speichern",
        defaultextension=".json",
        filetypes=[("JSON files", "*.json"), ("Alle Dateien", "*.*")],
    )
    if not config_path:
        return

    cfg = {
        "data_file_path": data_file_path,
        "columns": {
            "brand": get_brand_col(),
            "brand_band": get_brand_band_col(),
            "region": get_region_col(),
            "relationship": get_relationship_col(),
            "country": get_country_col(),
            "country_fallback": get_country_fallback_col(),
            "city": get_city_col(),
            "hyatt_date": get_hyatt_date_col(),
            "gm": get_gm_col(),
            "eng": get_eng_col(),
            "dof": get_dof_col(),
            "avp": get_avp_col(),
            "md": get_md_col(),
            "reg_eng_spec": get_reg_eng_spec_col(),
        },
        "roles": {role: var.get() for role, var in role_send_vars.items()},
        "visible_filter_cols": visible_optional_filter_cols,
    }
    try:
        with open(config_path, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, indent=2)
    except Exception as exc:
        messagebox.showerror("Konfiguration", f"Konfiguration konnte nicht gespeichert werden:\n{exc}")


def ensure_initial_data():
    """Load default data file if present, otherwise ask the user."""
    update_splash("", f"{TOOL_NAME} v{VERSION} ({VERSION_DATE}) - loading default data...")
    if os.path.isfile(DEFAULT_FILE_PATH):
        try:
            load_data(DEFAULT_FILE_PATH)
            update_splash(DEFAULT_FILE_PATH, "Default data loaded.")
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
            update_splash(data_file_path, "Warming Outlook...")
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
    global detail_checkbox_vars, detail_hotel_name, detail_row_current
    detail_checkbox_vars = []
    detail_hotel_name = row.get("Hotel", "N/A")
    detail_row_current = row

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
                canonical_role = "RegionalEngineeringSpecialist" if role.startswith("Regional") else role
                detail_checkbox_vars.append((var, str(email_address), canonical_role))
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

    def selected_values(lb):
        if lb is None:
            return []
        return [lb.get(i) for i in lb.curselection()]

    selected_brands = selected_values(brand_listbox)
    selected_regions = selected_values(region_listbox)
    selected_countries = selected_values(country_listbox)
    selected_bands = selected_values(brand_band_listbox)
    selected_relationships = selected_values(relationship_listbox)

    if selected_brands and brand_col in filt.columns:
        filt = filt[filt[brand_col].astype(str).isin(selected_brands)]
    if selected_regions and region_col in filt.columns:
        filt = filt[filt[region_col].astype(str).isin(selected_regions)]
    if selected_countries and country_col in filt.columns:
        filt = filt[filt[country_col].astype(str).isin(selected_countries)]
    if selected_bands and brand_band_col in filt.columns:
        filt = filt[filt[brand_band_col].astype(str).isin(selected_bands)]
    if selected_relationships and relationship_col in filt.columns:
        filt = filt[filt[relationship_col].astype(str).isin(selected_relationships)]

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

    # Reconfigure columns based on selection
    cols = get_filtered_columns()
    filtered_tree["columns"] = cols
    for col in cols:
        filtered_tree.heading(col, text=col)
        filtered_tree.column(col, width=120, stretch=True)

    for idx, (_, row) in enumerate(filt_df.iterrows()):
        tree_id = str(row.name)
        current_filtered_indexes.append(row.name)
        values = []
        for col in cols:
            if col == "Spirit Code":
                values.append(row.get("Spirit Code", ""))
            elif col == "Hotel":
                values.append(row.get("Hotel", ""))
            elif col == "Country" or col == "Country/Area":
                values.append(get_country_value(row))
            elif col == "City":
                values.append(get_city_value(row))
            else:
                values.append(row.get(col, ""))
        filtered_tree.insert("", "end", iid=tree_id, values=tuple(values))


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


def add_all_filtered_hotels():
    """Add all currently filtered hotels to the selection list."""
    filt_df = filtered_dataframe()
    for idx, row in filt_df.iterrows():
        if idx in selected_rows:
            continue
        selected_rows[idx] = row
    update_selected_tree()


def update_selected_tree():
    """Refresh the selected hotels list."""
    if selected_tree is None:
        return
    for item in selected_tree.get_children():
        selected_tree.delete(item)

    chosen_roles = []
    for role, var in role_send_vars.items():
        if var.get() != "Skip":
            chosen_roles.append(role)

    for row_idx, row in selected_rows.items():
        recips = []
        for role in chosen_roles or ["AVP", "MD", "GM", "Engineering", "DOF", "RegionalEngineeringSpecialist"]:
            recips.extend(get_role_addresses(row, role))
        recips = [r for r in recips if r]
        selected_tree.insert(
            "",
            "end",
            iid=str(row_idx),
            values=(
                row.get("Spirit Code", ""),
                row.get("Hotel", ""),
                "; ".join(recips),
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
            raw = str(row[col])
            for email in normalize_emails(raw):
                emails.append(email)
    return emails


def bind_autofit(tree: ttk.Treeview, min_width: int = 60):
    """Bind a resize handler to auto-distribute column widths."""
    if tree is None:
        return

    def _on_config(event):
        cols = tree["columns"]
        if not cols:
            return
        total = max(event.width - 20, len(cols) * min_width)
        per = total // len(cols)
        for col in cols:
            tree.column(col, width=per, stretch=True)

    tree.bind("<Configure>", _on_config)


def draft_emails_for_selection():
    """Create Outlook draft emails for the selected hotels and roles."""
    if not selected_rows:
        messagebox.showinfo("Keine Hotels", "Bitte waehlen Sie mindestens ein Hotel aus der Auswahl aus.")
        return

    chosen_roles = [role for role, var in role_send_vars.items() if var.get() != "Skip"]

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

    # Prompt for subject/body with placeholders
    def open_message_dialog():
        dialog = tk.Toplevel(root)
        dialog.title("Compose Email Template")
        dialog.geometry("480x360")
        ttk.Label(dialog, text="Subject (supports placeholders):").pack(anchor="w", padx=8, pady=(8, 2))
        subj_var = tk.StringVar(value="Hotel Information for {hotel}")
        subj_entry = ttk.Entry(dialog, textvariable=subj_var)
        subj_entry.pack(fill="x", padx=8)

        ttk.Label(dialog, text="Body (supports placeholders):").pack(anchor="w", padx=8, pady=(8, 2))
        body_text = tk.Text(dialog, height=10)
        body_text.pack(fill="both", expand=True, padx=8)
        body_text.insert("1.0", "Hotel: {hotel}\nSpirit: {spirit_code}\nCity: {city}\nBrand: {brand}\n\nYour message here.")

        placeholder_text = (
            "Placeholders:\n"
            "{hotel}, {spirit_code}, {city}, {relationship}, {brand}, {brand_band}, "
            "{region}, {country}, {owner}, {rooms}\n"
            "They will be replaced per hotel."
        )
        ttk.Label(dialog, text=placeholder_text, foreground="gray").pack(anchor="w", padx=8, pady=(4, 8))

        sigs = load_signatures()
        ttk.Label(dialog, text="Signature:").pack(anchor="w", padx=8, pady=(4, 2))
        sig_var = tk.StringVar(value="None")
        sig_combo = ttk.Combobox(dialog, textvariable=sig_var, values=list(sigs.keys()), state="readonly")
        sig_combo.pack(fill="x", padx=8, pady=(0, 6))

        def render_and_send():
            subject_template = subj_var.get()
            body_template = body_text.get("1.0", "end").rstrip("\n")
            dialog.destroy()
            send_emails(subject_template, body_template, sigs.get(sig_var.get(), {"html": "", "text": ""}))
        ttk.Button(dialog, text="Create Drafts", command=render_and_send).pack(pady=6)

    def send_emails(subject_template: str, body_template: str, signature_text: str):
        missing_addresses = []
        brand_col = get_brand_col()
        region_col = get_region_col()

        for row_idx, row in selected_rows.items():
            to_list = []
            cc_list = []
            bcc_list = []

            for role in chosen_roles:
                emails = get_role_addresses(row, role)
                mode = role_send_vars.get(role).get() if role in role_send_vars else "To"
                if mode == "To":
                    to_list.extend(emails)
                elif mode == "CC":
                    cc_list.extend(emails)
                elif mode == "BCC":
                    bcc_list.extend(emails)

            all_recips = [r for r in to_list + cc_list + bcc_list if r]
            if not all_recips:
                missing_addresses.append(row.get("Hotel", "N/A"))
                continue

            try:
                mail_item = outlook.CreateItem(0)
                mail_item.To = ";".join(to_list)
                mail_item.CC = ";".join(cc_list)
                mail_item.BCC = ";".join(bcc_list)
                hotel_name = row.get("Hotel", "Hotel")
                mail_item.Subject = render_template(row, subject_template)
                rendered = render_with_signature(render_template(row, body_template), signature_text)
                if rendered.get("html"):
                    mail_item.HTMLBody = rendered["html"]
                else:
                    mail_item.Body = rendered.get("text", "")
                mail_item.Display()
            except Exception as exc:
                messagebox.showerror("Email Error", f"Could not draft email for {row.get('Hotel', 'Hotel')}: {exc}")
                return

        if missing_addresses:
            messagebox.showinfo(
                "Keine Empfaenger",
                "Fuer folgende Hotels wurden keine E-Mail-Adressen in den gewaehlten Rollen gefunden:\n" + "\n".join(missing_addresses),
            )

    open_message_dialog()


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
    global detail_row_current
    detail_row_current = row
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
                canonical_role = "RegionalEngineeringSpecialist" if role.startswith("Regional") else role
                checkbox_vars.append((var, str(email_address), canonical_role))
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
    if os.name != "nt":
        messagebox.showerror("Unsupported Platform", "Outlook email drafting is only available on Windows.")
        return

    if not WIN32COM_AVAILABLE:
        messagebox.showerror(
            "Outlook Not Available",
            "This feature requires Microsoft Outlook and the 'pywin32' package (win32com.client).\nInstall with: pip install pywin32",
        )
        return

    # Compose dialog for single email
    def open_single_template():
        dialog = tk.Toplevel(root)
        dialog.title("Compose Email Template (Single Hotel)")
        dialog.geometry("480x320")
        ttk.Label(dialog, text="Subject (supports placeholders):").pack(anchor="w", padx=8, pady=(8, 2))
        subj_var = tk.StringVar(value="Hotel Information for {hotel}")
        subj_entry = ttk.Entry(dialog, textvariable=subj_var)
        subj_entry.pack(fill="x", padx=8)

        ttk.Label(dialog, text="Body (supports placeholders):").pack(anchor="w", padx=8, pady=(8, 2))
        body_text = tk.Text(dialog, height=8)
        body_text.pack(fill="both", expand=True, padx=8)
        body_text.insert(
            "1.0",
            "Hotel: {hotel}\nSpirit: {spirit_code}\nCity: {city}\nBrand: {brand}\n\nYour message here.",
        )

        placeholder_text = (
            "Placeholders: {hotel}, {spirit_code}, {city}, {relationship}, {brand}, {brand_band}, "
            "{region}, {country}, {owner}, {rooms}"
        )
        ttk.Label(dialog, text=placeholder_text, foreground="gray").pack(anchor="w", padx=8, pady=(4, 8))

        sigs = load_signatures()
        ttk.Label(dialog, text="Signature:").pack(anchor="w", padx=8, pady=(4, 2))
        sig_var = tk.StringVar(value="None")
        sig_combo = ttk.Combobox(dialog, textvariable=sig_var, values=list(sigs.keys()), state="readonly")
        sig_combo.pack(fill="x", padx=8, pady=(0, 6))

        def send_single():
            subject_template = subj_var.get()
            body_template = body_text.get("1.0", "end").rstrip("\n")
            dialog.destroy()

            try:
                outlook = get_outlook_app()
                mail_item = outlook.CreateItem(0)
            except Exception:
                try:
                    outlook = get_outlook_app(force_refresh=True)
                    mail_item = outlook.CreateItem(0)
                except Exception as exc:  # pragma: no cover - Outlook automation is Windows-specific
                    messagebox.showerror("Email Error", f"Could not draft email in Outlook: {exc}")
                    return

            to_list, cc_list, bcc_list = [], [], []
            for var, email, role_key in checkbox_vars:
                if var.get() and email:
                    emails = normalize_emails(email)
                    mode = role_send_vars.get(role_key).get() if role_key in role_send_vars else "To"
                    if mode == "To":
                        to_list.extend(emails)
                    elif mode == "CC":
                        cc_list.extend(emails)
                    elif mode == "BCC":
                        bcc_list.extend(emails)

            all_recips = [r for r in to_list + cc_list + bcc_list if r]
            if not all_recips:
                messagebox.showinfo("No Recipients", "No email addresses selected.")
                return

            mail_item.To = ";".join(to_list)
            mail_item.CC = ";".join(cc_list)
            mail_item.BCC = ";".join(bcc_list)

            mail_item.Subject = render_template(detail_row_current, subject_template)
            sig_entry = sigs.get(sig_var.get(), {"html": "", "text": ""})
            rendered = render_with_signature(render_template(detail_row_current, body_template), sig_entry)
            if rendered.get("html"):
                mail_item.HTMLBody = rendered["html"]
            else:
                mail_item.Body = rendered.get("text", "")
            mail_item.Display()
            if details_window is not None:
                details_window.destroy()

        ttk.Button(dialog, text="Create Draft", command=send_single).pack(pady=6)

    open_single_template()


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
ensure_style()
root.after(0, show_splash)
show_splash()

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
file_menu.add_command(label="Konfiguration laden", command=load_config_file)
file_menu.add_command(label="Konfiguration speichern", command=save_config_file)
file_menu.add_separator()
file_menu.add_command(label="Beenden", command=root.quit)
menubar.add_cascade(label="Datei", menu=file_menu)

def reopen_splash():
    show_splash()
    update_splash(data_file_path, "Status: Ready")

help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="About / Splash", command=reopen_splash)
menubar.add_cascade(label="About", menu=help_menu)

def show_readme():
    readme_path = os.path.join(BASE_DIR, "README.md")
    content = "README.md not found."
    if os.path.isfile(readme_path):
        try:
            with open(readme_path, "r", encoding="utf-8") as fh:
                content = fh.read()
        except Exception as exc:
            content = f"Could not read README.md:\n{exc}"
    win = tk.Toplevel(root)
    win.title("Instructions (README)")
    win.geometry("760x520")
    text = tk.Text(win, wrap="word")
    text.insert("1.0", content)
    text.config(state="disabled")
    text.pack(fill="both", expand=True, padx=6, pady=6)
    scroll = ttk.Scrollbar(win, command=text.yview)
    text.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")

help_menu2 = tk.Menu(menubar, tearoff=0)
help_menu2.add_command(label="Instructions", command=show_readme)
menubar.add_cascade(label="Help", menu=help_menu2)

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

hyatt_year_var = tk.StringVar(value="")
hyatt_year_mode_var = tk.StringVar(value="Any")

def make_multiselect(parent, label_text):
    wrap = ttk.Frame(parent)
    ttk.Label(wrap, text=label_text).pack(anchor="w")
    lb = tk.Listbox(wrap, selectmode="extended", height=6, exportselection=False)
    lb.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(wrap, orient="vertical", command=lb.yview)
    sb.pack(side="right", fill="y")
    lb.config(yscrollcommand=sb.set)
    return wrap, lb

row_f = 0
brand_wrap, brand_listbox = make_multiselect(filters_frame, "Brand (multi-select)")
brand_wrap.grid(row=row_f, column=0, sticky="nsew", padx=4, pady=2)

band_wrap, brand_band_listbox = make_multiselect(filters_frame, "Brand Band")
band_wrap.grid(row=row_f, column=1, sticky="nsew", padx=4, pady=2)

region_wrap, region_listbox = make_multiselect(filters_frame, "Region")
region_wrap.grid(row=row_f, column=2, sticky="nsew", padx=4, pady=2)

relationship_wrap, relationship_listbox = make_multiselect(filters_frame, "Relationship")
relationship_wrap.grid(row=row_f, column=3, sticky="nsew", padx=4, pady=2)

country_wrap, country_listbox = make_multiselect(filters_frame, "Country/Area")
country_wrap.grid(row=row_f, column=4, sticky="nsew", padx=4, pady=2)

hyatt_wrap = ttk.Frame(filters_frame)
hyatt_wrap.grid(row=row_f, column=5, sticky="nw", padx=4, pady=2)
ttk.Label(hyatt_wrap, text="Hyatt Date (year)").pack(anchor="w")
hyatt_year_entry = ttk.Entry(hyatt_wrap, textvariable=hyatt_year_var, width=10)
hyatt_year_entry.pack(anchor="w", pady=(0, 2))
hyatt_mode_combo = ttk.Combobox(
    hyatt_wrap,
    textvariable=hyatt_year_mode_var,
    values=["Any", "Before", "Before/Equal", "Equal", "After/Equal", "After"],
    state="readonly",
    width=12,
)
hyatt_mode_combo.pack(anchor="w")

for col in range(5):
    filters_frame.columnconfigure(col, weight=1)

apply_filter_btn = ttk.Button(filters_frame, text="Apply Filter", command=refresh_filtered_hotels)
apply_filter_btn.grid(row=0, column=6, sticky="e", padx=8, pady=2)

reset_filter_btn = ttk.Button(filters_frame, text="Reset Filters", command=reset_filters)
reset_filter_btn.grid(row=0, column=7, sticky="e", padx=8, pady=2)

lists_pane = ttk.Panedwindow(multi_frame, orient="horizontal")
lists_pane.pack(fill="both", expand=True, padx=5, pady=5)

# Buttons between filters and panes
buttons_bar = ttk.Frame(multi_frame)
buttons_bar.pack(fill="x", padx=5, pady=(0, 5))
ttk.Button(buttons_bar, text="Select", command=add_selected_hotels).pack(side="left", padx=4)
ttk.Button(buttons_bar, text="Select All", command=add_all_filtered_hotels).pack(side="left", padx=4)
ttk.Button(buttons_bar, text="Remove", command=remove_selected_hotels).pack(side="left", padx=4)
ttk.Button(buttons_bar, text="Remove All", command=clear_selected_hotels).pack(side="left", padx=4)

filtered_frame = ttk.LabelFrame(lists_pane, text="Gefilterte Hotels", padding=5)
lists_pane.add(filtered_frame, weight=1)

filtered_tree = ttk.Treeview(
    filtered_frame,
    columns=("Spirit", "Hotel", "City", "Brand", "Brand Band", "Relationship", "Region", "Country"),
    show="headings",
    selectmode="extended",
)
filtered_xscroll = ttk.Scrollbar(filtered_frame, orient="horizontal", command=filtered_tree.xview)
filtered_tree.configure(xscrollcommand=filtered_xscroll.set)
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
filtered_xscroll.pack(fill="x")
bind_autofit(filtered_tree)

selected_frame = ttk.LabelFrame(lists_pane, text="Ausgewaehlte Hotels", padding=5)
lists_pane.add(selected_frame, weight=1)

selected_tree = ttk.Treeview(
    selected_frame,
    columns=("Spirit", "Hotel", "Recipients"),
    show="headings",
    selectmode="extended",
)
selected_xscroll = ttk.Scrollbar(selected_frame, orient="horizontal", command=selected_tree.xview)
selected_tree.configure(xscrollcommand=selected_xscroll.set)
for col, width in [
    ("Spirit", 80),
    ("Hotel", 220),
    ("Recipients", 360),
]:
    selected_tree.heading(col, text=col)
    selected_tree.column(col, width=width, stretch=True)
selected_tree.pack(fill="both", expand=True)
selected_xscroll.pack(fill="x")
bind_autofit(selected_tree)

draft_btn = ttk.Button(multi_frame, text="Draft Emails in Outlook", command=draft_emails_for_selection)
draft_btn.pack(anchor="e", padx=8, pady=6)

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

role_delivery = ttk.LabelFrame(setup_frame, text="Role Delivery (To/CC/BCC)", padding=10)
role_delivery.pack(fill="x", padx=5, pady=5)
role_delivery.columnconfigure(1, weight=1)
add_role_selector(role_delivery, "AVP", "Skip")
add_role_selector(role_delivery, "MD", "Skip")
add_role_selector(role_delivery, "GM", "To")
add_role_selector(role_delivery, "Engineering", "CC")
add_role_selector(role_delivery, "DOF", "CC")
add_role_selector(role_delivery, "RegionalEngineeringSpecialist", "CC")

visible_cols_frame = ttk.LabelFrame(setup_frame, text='Columns shown in "Gefilterte Hotels"', padding=10)
visible_cols_frame.pack(fill="both", padx=5, pady=5)
filter_cols_listbox = tk.Listbox(visible_cols_frame, selectmode="extended", height=8, exportselection=False)
filter_cols_listbox.pack(fill="both", expand=True)

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
update_filter_options()

# Initial filtered view
refresh_filtered_hotels()

# Auto-close splash after 2 minutes
root.after(120000, close_splash)

# Warm Outlook in the background so the first email opens faster
root.after(200, warm_outlook_app)

root.mainloop()
