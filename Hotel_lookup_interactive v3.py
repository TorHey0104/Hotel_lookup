#!/usr/bin/env python3
import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import urllib.parse # For URL encoding email addresses

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
        to_addresses = ",".join(recipients)
        subject = urllib.parse.quote(f"Hotel Information for {hotel_name}") # URL-encode subject
        
        # Construct mailto URI
        mailto_uri = f"mailto:{to_addresses}?subject={subject}"
        
        try:
            # Use 'open' command on macOS to launch default mail client
            os.system(f'open "{mailto_uri}"')
        except Exception as e:
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
    
    # Call the new GUI function
    show_details_gui(row)

tk.Button(root, text="Search", command=lookup).grid(row=2, column=0, columnspan=2, pady=10)

root.mainloop()
