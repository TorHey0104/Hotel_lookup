import tkinter as tk
from tkinter import ttk

def make_multiselect(parent, label_text, height=6):
    wrap = ttk.Frame(parent)
    ttk.Label(wrap, text=label_text).pack(anchor="w")
    lb = tk.Listbox(wrap, selectmode="extended", height=height, exportselection=False)
    lb.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(wrap, orient="vertical", command=lb.yview)
    sb.pack(side="right", fill="y")
    lb.config(yscrollcommand=sb.set)
    return wrap, lb
