#!/usr/bin/env python3
import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText

# ————————————————————————————————————
# ◉ CONFIGURE THIS
# ————————————————————————————————————
DATA_DIR  = "/Users/torstenheyroth/Library/CloudStorage/OneDrive-HyattHotels/___DATA"
FILE_NAME = "2a Hotels one line hotel.xlsx"
FILE_PATH = os.path.join(DATA_DIR, FILE_NAME)
if not os.path.isfile(FILE_PATH):
    raise FileNotFoundError(f"No file at {FILE_PATH}")
# ————————————————————————————————————

# Load data
df = pd.read_excel(FILE_PATH, engine="openpyxl")

# Prepare hotel list
hotel_names = sorted(df['Hotel'].dropna().unique().tolist())

# Build GUI
root = tk.Tk()
root.title("Hotel Lookup")

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

def show_details(title, content):
    """Popup a scrollable window with the given content."""
    win = tk.Toplevel(root)
    win.title(title)
    # make it resizable
    win.geometry("600x400")
    win.minsize(400, 200)
    txt = ScrolledText(win, wrap=tk.WORD)
    txt.insert(tk.END, content)
    txt.configure(state='disabled')  # read-only
    txt.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

def lookup():
    spirit = code_entry.get().strip()
    hotel   = hotel_var.get().strip()
    if spirit:
        result = df[df['Spirit Code'] == spirit]
    elif hotel:
        result = df[df['Hotel'] == hotel]
    else:
        messagebox.showwarning("Whoops", "Enter Spirit Code or pick a hotel.")
        return

    if result.empty:
        messagebox.showinfo("Nada", "No matching hotel found.")
        return

    row = result.iloc[0]
    
    # Define the specific columns to display
    display_cols = [
        'Spirit Code', 'Hotel', 'Affiliation Date', 'Opening Date' 'Rooms','City', 'Geographical Area','Owner', 'JV', 'JV Percent','Relationship', 'SVP',
        'RVP of OPS', 'AVP of Ops', 'AVP of Ops-managed', 
        'GM - Primary', 'GM', 'DOF', 'Engineering Director', 'Engineering Director / Chief Engineer'
    ]
    
    # Filter for only the columns that exist in the DataFrame
    existing_display_cols = [col for col in display_cols if col in df.columns]

    details = "\n".join(f"{col}: {row[col]}" for col in existing_display_cols)
    
    # instead of messagebox, use scrollable window
    show_details(f"Info for {row['Hotel']}", details)

tk.Button(root, text="Search", command=lookup).grid(row=2, column=0, columnspan=2, pady=10)

root.mainloop()
