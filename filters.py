import pandas as pd

MANDATORY_FILTER_COLS = ["Spirit Code", "Hotel"]

def apply_filters(
    df: pd.DataFrame,
    brand_col: str,
    region_col: str,
    country_col: str,
    brand_band_col: str,
    relationship_col: str,
    hyatt_col: str,
    quick_codes: list[str],
    selected_brands,
    selected_regions,
    selected_countries,
    selected_bands,
    selected_relationships,
    hyatt_year_var,
    hyatt_year_mode_var,
) -> pd.DataFrame:
    filt = df
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

    if quick_codes and "Spirit Code" in filt.columns:
        filt = filt[filt["Spirit Code"].astype(str).isin(quick_codes)]
    return filt
