import tkinter as tk
from tkinter import ttk


def create_multi_tab(notebook, state, callbacks, make_multiselect, build_recipient_tree, enable_treeview_sort):
    """Build the multi-email tab. Mutates state dict with widgets/vars."""
    frame = ttk.Frame(notebook, padding=10)
    notebook.add(frame, text="Multi-Email")

    # Convenience getters from state
    attachments_enabled_var = state["attachments_enabled_var"]
    attachments_root_var = state["attachments_root_var"]
    hyatt_year_var = state["hyatt_year_var"]
    hyatt_year_mode_var = state["hyatt_year_mode_var"]

    # Top forward controls
    forward_bar = ttk.Frame(frame)
    forward_bar.pack(fill="x", padx=5, pady=(0, 6))
    state["forward_status_var"] = tk.StringVar(value="No source email captured.")
    ttk.Button(forward_bar, text="Browse Outlook...", command=callbacks["browse_outlook_email"]).pack(side="left", padx=4)
    ttk.Button(forward_bar, text="Clear Forward", command=callbacks["clear_forward_template"]).pack(side="left", padx=4)
    ttk.Label(forward_bar, textvariable=state["forward_status_var"], foreground="gray").pack(side="left", padx=8)

    # Quick filter
    quick_frame = ttk.Frame(frame)
    quick_frame.pack(fill="x", padx=5, pady=(0, 6))
    ttk.Label(quick_frame, text="Quick Spirit Codes (comma separated)").pack(side="left", padx=4)
    state["quick_spirit_var"] = tk.StringVar()
    ttk.Entry(quick_frame, textvariable=state["quick_spirit_var"]).pack(side="left", padx=4, fill="x", expand=True)
    ttk.Button(quick_frame, text="Apply Quick Filter", command=callbacks["refresh_filtered_hotels"]).pack(side="left", padx=4)
    state["filtered_count_var"] = tk.StringVar(value="Filtered: 0")
    ttk.Label(quick_frame, textvariable=state["filtered_count_var"], foreground="gray").pack(side="right", padx=4)

    # Attachments
    attachments_frame = ttk.LabelFrame(frame, text="Attachments (multi-email)", padding=6)
    attachments_frame.pack(fill="x", padx=5, pady=(0, 6))
    ttk.Checkbutton(attachments_frame, text="Enable attachments", variable=attachments_enabled_var).grid(row=0, column=0, sticky="w", padx=4, pady=2)
    ttk.Label(attachments_frame, text="Attachments root").grid(row=1, column=0, sticky="w", padx=4, pady=2)
    ttk.Entry(attachments_frame, textvariable=attachments_root_var).grid(row=1, column=1, sticky="ew", padx=4, pady=2)
    ttk.Button(attachments_frame, text="Browse", command=callbacks["browse_attachments_root"]).grid(row=1, column=2, sticky="e", padx=4, pady=2)
    attachments_frame.columnconfigure(1, weight=1)

    # Filters
    filters_frame = ttk.LabelFrame(frame, text="Filter Hotels", padding=10)
    filters_frame.pack(fill="x", padx=5, pady=5)
    hyatt_year_var.set("")
    hyatt_year_mode_var.set("Any")

    row_f = 0
    brand_wrap, state["brand_listbox"] = make_multiselect(filters_frame, "Brand (multi-select)")
    brand_wrap.grid(row=row_f, column=0, sticky="nsew", padx=4, pady=2)

    band_wrap, state["brand_band_listbox"] = make_multiselect(filters_frame, "Brand Band")
    band_wrap.grid(row=row_f, column=1, sticky="nsew", padx=4, pady=2)

    region_wrap, state["region_listbox"] = make_multiselect(filters_frame, "Region")
    region_wrap.grid(row=row_f, column=2, sticky="nsew", padx=4, pady=2)

    relationship_wrap, state["relationship_listbox"] = make_multiselect(filters_frame, "Relationship")
    relationship_wrap.grid(row=row_f, column=3, sticky="nsew", padx=4, pady=2)

    country_wrap, state["country_listbox"] = make_multiselect(filters_frame, "Country/Area")
    country_wrap.grid(row=row_f, column=4, sticky="nsew", padx=4, pady=2)

    hyatt_wrap = ttk.Frame(filters_frame)
    hyatt_wrap.grid(row=row_f, column=5, sticky="nw", padx=4, pady=2)
    ttk.Label(hyatt_wrap, text="Hyatt Date (year)").pack(anchor="w")
    ttk.Entry(hyatt_wrap, textvariable=hyatt_year_var, width=10).pack(anchor="w", pady=(0, 2))
    ttk.Combobox(
        hyatt_wrap,
        textvariable=hyatt_year_mode_var,
        values=["Any", "Before", "Before/Equal", "Equal", "After/Equal", "After"],
        state="readonly",
        width=12,
    ).pack(anchor="w")

    for col in range(5):
        filters_frame.columnconfigure(col, weight=1)

    ttk.Button(filters_frame, text="Apply Filter", command=callbacks["refresh_filtered_hotels"]).grid(row=0, column=6, sticky="e", padx=8, pady=2)
    ttk.Button(filters_frame, text="Reset Filters", command=callbacks["reset_filters"]).grid(row=0, column=7, sticky="e", padx=8, pady=2)

    lists_pane = ttk.Panedwindow(frame, orient="horizontal")
    lists_pane.pack(fill="both", expand=True, padx=5, pady=5)

    buttons_bar = ttk.Frame(frame)
    buttons_bar.pack(fill="x", padx=5, pady=(0, 5))
    ttk.Button(buttons_bar, text="Select", command=callbacks["add_selected_hotels"]).pack(side="left", padx=4)
    ttk.Button(buttons_bar, text="Select All", command=callbacks["add_all_filtered_hotels"]).pack(side="left", padx=4)
    ttk.Button(buttons_bar, text="Remove", command=callbacks["remove_selected_hotels"]).pack(side="left", padx=4)
    ttk.Button(buttons_bar, text="Remove All", command=callbacks["clear_selected_hotels"]).pack(side="left", padx=4)

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
        ("Spirit", 70),
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
    enable_treeview_sort(filtered_tree)

    selected_frame = ttk.LabelFrame(lists_pane, text="Ausgewaehlte Hotels", padding=5)
    lists_pane.add(selected_frame, weight=1)
    selected_widths = {
        "Spirit": 60,
        "Hotel": 260,
        "Recipients": 500,
        "AVP": 30,
        "MD": 30,
        "GM": 30,
        "ENG": 30,
        "DOF": 30,
        "RES": 30,
    }
    selected_tree, selected_xscroll = build_recipient_tree(selected_frame, widths=selected_widths, add_scroll=True)
    selected_tree.pack(fill="both", expand=True)
    if selected_xscroll:
        selected_xscroll.pack(fill="x")
    enable_treeview_sort(selected_tree)

    ttk.Button(frame, text="Draft Emails in Outlook", command=callbacks["draft_emails_for_selection"]).pack(anchor="e", padx=8, pady=6)
    ttk.Button(frame, text="Draft ONE collective email", command=callbacks["draft_collective_email"]).pack(anchor="e", padx=8, pady=(0, 6))

    state["filtered_tree"] = filtered_tree
    state["selected_tree"] = selected_tree
    return frame


def create_excel_tab(notebook, state, callbacks, build_recipient_tree, enable_treeview_sort):
    """Build the Excel-email tab. Mutates state dict with widgets/vars."""
    frame = ttk.Frame(notebook, padding=10)
    notebook.add(frame, text="Excel Emails")

    # State vars
    attachments_enabled_var = state["attachments_enabled_var"]
    attachments_root_var = state["attachments_root_var"]

    excel_top = ttk.Frame(frame)
    excel_top.pack(fill="x", pady=(0, 10))
    ttk.Label(excel_top, text="Excel-Datei fuer Emails laden:").pack(side="left", padx=4)
    ttk.Button(excel_top, text="Datei laden", command=callbacks["load_excel_email_file"]).pack(side="left", padx=4)
    ttk.Label(excel_top, textvariable=state["excel_file_label_var"], foreground="gray").pack(side="left", padx=8)

    excel_forward_bar = ttk.Frame(frame)
    excel_forward_bar.pack(fill="x", padx=5, pady=(0, 6))
    ttk.Label(excel_forward_bar, text="Forward-Email (optional):").pack(side="left", padx=4)
    ttk.Button(excel_forward_bar, text="Browse Outlook...", command=callbacks["browse_outlook_email"]).pack(side="left", padx=4)
    ttk.Button(excel_forward_bar, text="Clear Forward", command=callbacks["clear_forward_template"]).pack(side="left", padx=4)
    ttk.Label(excel_forward_bar, textvariable=state["forward_status_var"], foreground="gray").pack(side="left", padx=8)

    excel_attachments_frame = ttk.LabelFrame(frame, text="Attachments (Excel emails)", padding=6)
    excel_attachments_frame.pack(fill="x", padx=5, pady=(0, 6))
    ttk.Checkbutton(excel_attachments_frame, text="Enable attachments", variable=attachments_enabled_var).grid(row=0, column=0, sticky="w", padx=4, pady=2)
    ttk.Label(excel_attachments_frame, text="Attachments root").grid(row=1, column=0, sticky="w", padx=4, pady=2)
    ttk.Entry(excel_attachments_frame, textvariable=attachments_root_var).grid(row=1, column=1, sticky="ew", padx=4, pady=2)
    ttk.Button(excel_attachments_frame, text="Browse", command=callbacks["browse_attachments_root"]).grid(row=1, column=2, sticky="e", padx=4, pady=2)
    excel_attachments_frame.columnconfigure(1, weight=1)

    mapping_box = ttk.LabelFrame(frame, text="Spaltenzuordnung", padding=8)
    mapping_box.pack(fill="both", expand=True, pady=(0, 10))
    ttk.Label(mapping_box, text="Waehlen Sie je Spalte: Spirit Code / Include in Body / Skip").pack(anchor="w", pady=(0, 6))
    mapping_container = ttk.Frame(mapping_box)
    mapping_container.pack(fill="both", expand=True)
    headers_canvas = tk.Canvas(mapping_container, borderwidth=0, highlightthickness=0)
    headers_scroll = ttk.Scrollbar(mapping_container, orient="vertical", command=headers_canvas.yview)
    state["excel_headers_frame"] = ttk.Frame(headers_canvas)
    headers_window = headers_canvas.create_window((0, 0), window=state["excel_headers_frame"], anchor="nw")

    def _on_headers_config(event):
        headers_canvas.configure(scrollregion=headers_canvas.bbox("all"))

    state["excel_headers_frame"].bind("<Configure>", _on_headers_config)
    headers_canvas.configure(yscrollcommand=headers_scroll.set)
    headers_canvas.pack(side="left", fill="both", expand=True)
    headers_scroll.pack(side="right", fill="y")

    filter_bar = ttk.Frame(frame)
    filter_bar.pack(fill="x", pady=4)
    state["excel_filter_summary_var"] = tk.StringVar(value="")
    ttk.Button(filter_bar, text="Filter anwenden", command=callbacks["refresh_excel_filtered_tree"]).pack(side="left", padx=4)
    ttk.Button(
        filter_bar,
        text="Filter loeschen",
        command=lambda: [lb.selection_clear(0, tk.END) for _, _, lb in callbacks["excel_filter_controls"]()] or callbacks["refresh_excel_filtered_tree"](),
    ).pack(side="left", padx=4)
    ttk.Label(filter_bar, textvariable=state["excel_filter_summary_var"], foreground="gray").pack(side="left", padx=6)
    state["excel_filtered_count_var"] = tk.StringVar(value="Gefiltert: 0")
    ttk.Label(filter_bar, textvariable=state["excel_filtered_count_var"], foreground="gray").pack(side="left", padx=6)

    actions = ttk.Frame(frame)
    actions.pack(fill="x", pady=4)
    ttk.Label(actions, text="Versand-Modus:").pack(side="left", padx=10)
    ttk.Radiobutton(actions, text="Einzel pro Hotel", variable=state["excel_mode_var"], value="dedicated").pack(side="left")
    ttk.Radiobutton(actions, text="Eine Sammelmail", variable=state["excel_mode_var"], value="collective").pack(side="left", padx=6)
    ttk.Button(actions, text="Emails erstellen", command=callbacks["prompt_excel_compose"]).pack(side="right", padx=4)

    excel_lists = ttk.Panedwindow(frame, orient="horizontal")
    excel_lists.pack(fill="both", expand=True, pady=6)

    excel_filtered_frame = ttk.LabelFrame(excel_lists, text="Gefilterte Excel-Hotels", padding=5)
    excel_lists.add(excel_filtered_frame, weight=1)
    excel_widths = {
        "Spirit": 70,
        "Hotel": 220,
        "Recipients": 420,
        "AVP": 40,
        "MD": 40,
        "GM": 40,
        "ENG": 40,
        "DOF": 40,
        "RES": 40,
    }
    excel_filtered_tree, _ = build_recipient_tree(excel_filtered_frame, widths=excel_widths, add_scroll=False)
    excel_filtered_tree.pack(fill="both", expand=True)
    enable_treeview_sort(excel_filtered_tree)

    excel_selected_frame = ttk.LabelFrame(excel_lists, text="Ausgewaehlte Excel-Hotels", padding=5)
    excel_lists.add(excel_selected_frame, weight=1)
    excel_selected_tree, _ = build_recipient_tree(excel_selected_frame, widths=excel_widths, add_scroll=False)
    excel_selected_tree.pack(fill="both", expand=True)
    enable_treeview_sort(excel_selected_tree)

    excel_btns = ttk.Frame(frame)
    excel_btns.pack(fill="x", pady=4)
    ttk.Button(excel_btns, text="Auswahl hinzufuegen", command=callbacks["excel_add_selected"]).pack(side="left", padx=4)
    ttk.Button(excel_btns, text="Alle hinzufuegen", command=callbacks["excel_add_all"]).pack(side="left", padx=4)
    ttk.Button(excel_btns, text="Entfernen", command=callbacks["excel_remove_selected"]).pack(side="left", padx=4)
    ttk.Button(excel_btns, text="Alle entfernen", command=callbacks["excel_clear_selected"]).pack(side="left", padx=4)

    state["excel_filtered_tree"] = excel_filtered_tree
    state["excel_selected_tree"] = excel_selected_tree
    return frame
