import tkinter as tk
from tkinter import ttk, filedialog


def create_lookup_tab(
    notebook,
    hotel_names,
    lookup_handler,
    init_detail_panel_fn,
    clear_detail_panel_fn,
    init_single_compose_ui_fn,
    single_attachments_enabled_var,
    single_attachments_root_var,
):
    """Build and attach the single-hotel lookup tab."""
    frame = ttk.Frame(notebook, padding=10)
    notebook.add(frame, text="Lookup")
    frame.columnconfigure(1, weight=1)
    frame.rowconfigure(0, weight=1)
    frame.rowconfigure(1, weight=1)

    lookup_form = ttk.Frame(frame)
    lookup_form.grid(row=0, column=0, sticky="nw", padx=(0, 10))

    ttk.Label(lookup_form, text="Spirit Code:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    spirit_entry = tk.Entry(lookup_form, width=30)
    spirit_entry.grid(row=0, column=1, padx=5, pady=5)

    ttk.Label(lookup_form, text="Hotel:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    hotel_var = tk.StringVar()
    hotel_combo = ttk.Combobox(lookup_form, textvariable=hotel_var, values=hotel_names, state="normal")
    hotel_combo.grid(row=1, column=1, padx=5, pady=5)

    def on_hotel_keyrelease(event):
        val = hotel_var.get()
        hotel_combo["values"] = hotel_names if not val else [h for h in hotel_names if val.lower() in h.lower()]

    hotel_combo.bind("<KeyRelease>", on_hotel_keyrelease)

    tk.Button(lookup_form, text="Search", command=lambda: lookup_handler(spirit_entry, hotel_var)).grid(
        row=2, column=0, columnspan=2, pady=10
    )

    single_attach_frame = ttk.LabelFrame(lookup_form, text="Attachments (single email)", padding=8)
    single_attach_frame.grid(row=3, column=0, columnspan=2, sticky="we", padx=2, pady=6)
    ttk.Checkbutton(
        single_attach_frame,
        text="Enable attachments for single email",
        variable=single_attachments_enabled_var,
    ).grid(row=0, column=0, sticky="w", padx=4, pady=2)
    ttk.Label(single_attach_frame, text="Attachments root").grid(row=1, column=0, sticky="w", padx=4, pady=2)
    single_attach_entry = ttk.Entry(single_attach_frame, textvariable=single_attachments_root_var, width=40)
    single_attach_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=2)

    def browse_single_attach_root():
        sel = filedialog.askdirectory(title="Choose attachments root (single email)")
        if sel:
            single_attachments_root_var.set(sel)

    ttk.Button(single_attach_frame, text="Browse", command=browse_single_attach_root).grid(
        row=1, column=2, sticky="e", padx=4, pady=2
    )
    single_attach_frame.columnconfigure(1, weight=1)

    detail_container = ttk.Frame(frame)
    detail_container.grid(row=0, column=1, sticky="nsew")
    init_detail_panel_fn(detail_container)
    clear_detail_panel_fn()
    init_single_compose_ui_fn(frame)
    return frame, hotel_combo


def create_setup_tab(
    notebook,
    brand_col_var,
    region_col_var,
    country_col_var,
    country_fallback_col_var,
    city_col_var,
    brand_band_col_var,
    relationship_col_var,
    hyatt_date_col_var,
    avp_col_var,
    md_col_var,
    gm_col_var,
    eng_col_var,
    dof_col_var,
    reg_eng_spec_col_var,
    add_role_selector_fn,
    apply_column_settings_fn,
):
    """Build and attach the setup/config tab and return the widgets that need to stay globally reachable."""
    frame = ttk.Frame(notebook, padding=10)
    notebook.add(frame, text="Setup")

    setup_top = ttk.LabelFrame(frame, text="Data Columns", padding=10)
    setup_top.pack(fill="x", padx=5, pady=5)

    brand_col_combo = ttk.Combobox(setup_top, textvariable=brand_col_var, state="readonly")
    region_col_combo = ttk.Combobox(setup_top, textvariable=region_col_var, state="readonly")
    country_col_combo = ttk.Combobox(setup_top, textvariable=country_col_var, state="readonly")
    country_fallback_combo = ttk.Combobox(setup_top, textvariable=country_fallback_col_var, state="readonly")
    city_col_combo = ttk.Combobox(setup_top, textvariable=city_col_var, state="readonly")
    brand_band_col_combo = ttk.Combobox(setup_top, textvariable=brand_band_col_var, state="readonly")
    relationship_col_combo = ttk.Combobox(setup_top, textvariable=relationship_col_var, state="readonly")
    hyatt_date_col_combo = ttk.Combobox(setup_top, textvariable=hyatt_date_col_var, state="readonly")

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
        ttk.Label(setup_top, text=text).grid(row=idx, column=0, sticky="w", padx=5, pady=2)
        combo.grid(row=idx, column=1, sticky="ew", padx=5, pady=2)
    setup_top.columnconfigure(1, weight=1)

    roles_setup = ttk.LabelFrame(frame, text="Recipient Columns", padding=10)
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

    role_delivery = ttk.LabelFrame(frame, text="Role Delivery (To/CC/BCC)", padding=10)
    role_delivery.pack(fill="x", padx=5, pady=5)
    role_delivery.columnconfigure(1, weight=1)
    add_role_selector_fn(role_delivery, "AVP", "Skip")
    add_role_selector_fn(role_delivery, "MD", "Skip")
    add_role_selector_fn(role_delivery, "GM", "To")
    add_role_selector_fn(role_delivery, "Engineering", "CC")
    add_role_selector_fn(role_delivery, "DOF", "CC")
    add_role_selector_fn(role_delivery, "RegionalEngineeringSpecialist", "CC")

    visible_cols_frame = ttk.LabelFrame(frame, text='Columns shown in "Gefilterte Hotels"', padding=10)
    visible_cols_frame.pack(fill="both", padx=5, pady=5)
    filter_cols_listbox = tk.Listbox(visible_cols_frame, selectmode="extended", height=8, exportselection=False)
    filter_cols_listbox.pack(fill="both", expand=True)

    ttk.Button(frame, text="Apply column mapping", command=apply_column_settings_fn).pack(anchor="e", padx=5, pady=10)

    return frame, {
        "brand_col_combo": brand_col_combo,
        "region_col_combo": region_col_combo,
        "country_col_combo": country_col_combo,
        "country_fallback_combo": country_fallback_combo,
        "city_col_combo": city_col_combo,
        "brand_band_col_combo": brand_band_col_combo,
        "relationship_col_combo": relationship_col_combo,
        "hyatt_date_col_combo": hyatt_date_col_combo,
        "avp_col_combo": avp_col_combo,
        "md_col_combo": md_col_combo,
        "gm_col_combo": gm_col_combo,
        "eng_col_combo": eng_col_combo,
        "dof_col_combo": dof_col_combo,
        "reg_eng_spec_combo": reg_eng_spec_combo,
        "filter_cols_listbox": filter_cols_listbox,
    }
