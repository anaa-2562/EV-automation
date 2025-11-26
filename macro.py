import os
import pandas as pd
from typing import Tuple


def _read_raw(path: str) -> pd.DataFrame:
    """Load the Raw sheet (or CSV) exactly once."""
    if path.lower().endswith(".csv"):
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    for sheet in ("Raw", 0):
        try:
            return pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False, engine="openpyxl")
        except ValueError:
            continue
    raise ValueError("Could not load Raw sheet from input file")


def _read_help(path: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name="Help", dtype=str, keep_default_na=False, engine="openpyxl")


def audentes_verification_cleaned(raw_path: str, help_path: str, output_path: str) -> Tuple[pd.DataFrame, dict]:
    """Python translation of Audentes_Verification_Cleaned VBA macro."""

    print("Loading raw and help data...")
    df_raw = _read_raw(raw_path)
    df_help = _read_help(help_path)

    df_raw = df_raw.fillna("")
    df_help = df_help.fillna("")

    # Column headers from help sheet (matches VBA B1, D1, E1)
    provider_col = df_help.columns[0]  # Help column C1
    state_col = df_help.columns[1]     # Help column C2
    visit_type_col = df_help.columns[2]  # Help column C3
    workable_col_help = df_help.columns[3]  # Help column C4
    primary_ins_help = df_help.columns[4]  # Help column C5

    # Copy raw so we do not mutate original
    df = df_raw.copy()

    # Insert helper columns A / B / C analogous to Excel macro
    print("Creating helper columns (Appointment State, Workable Status, HelpCheck)...")
    provider_map = {
        str(k).strip().upper(): str(v).strip()
        for k, v in zip(df_help[provider_col], df_help[state_col])
        if str(k).strip()
    }
    visit_map = {
        str(k).strip().upper(): str(v).strip()
        for k, v in zip(df_help[visit_type_col], df_help[workable_col_help])
        if str(k).strip()
    }
    # HelpCheck simply mirrors the primary insurance names (used for exclusions)
    helpcheck_map = {
        str(k).strip().upper(): str(k).strip() for k in df_help[primary_ins_help] if str(k).strip()
    }

    df.insert(
        0,
        "Appointment State",
        df["Appointment Provider Name"]
        .astype(str)
        .map(lambda x: provider_map.get(x.strip().upper(), "")),
    )
    df.insert(
        1,
        "Workable Status",
        df.get("Visit Type", "")
        .astype(str)
        .map(lambda x: visit_map.get(x.strip().upper(), "")),
    )
    df.insert(
        2,
        "HelpCheck",
        df.get("Primary Insurance Name", "")
        .astype(str)
        .map(lambda x: helpcheck_map.get(x.strip().upper(), x.strip())),
    )

    # Track counts for logging
    initial_rows = len(df)

    print("Applying filter: remove rows with missing Appointment State (#N/A)...")
    mask_state = df["Appointment State"].astype(str).str.strip().isin(["", "#N/A"])
    removed_state = int(mask_state.sum())
    df = df[~mask_state].reset_index(drop=True)

    print("Applying filter: remove rows with Workable Status N or #N/A...")
    mask_workable = df["Workable Status"].astype(str).str.strip().str.upper().isin(["N", "#N/A"])
    removed_workable = int(mask_workable.sum())
    df = df[~mask_workable].reset_index(drop=True)

    print("Applying filter: remove HelpCheck codes L105/L107/L109C/L109Q/L109W...")
    exclude_codes = {"L105", "L107", "L109C", "L109Q", "L109W"}
    mask_codes = df["HelpCheck"].astype(str).str.strip().isin(exclude_codes)
    removed_codes = int(mask_codes.sum())
    df = df[~mask_codes].reset_index(drop=True)

    final_rows = len(df)
    print("Macro filtering summary:")
    print(f"  Initial rows: {initial_rows}")
    print(f"  Removed (Appointment State missing/#N/A): {removed_state}")
    print(f"  Removed (Workable Status N/#N/A): {removed_workable}")
    print(f"  Removed (HelpCheck exclusion codes): {removed_codes}")
    print(f"  Final remaining rows: {final_rows}")

    # Save cleaned output
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df.to_excel(output_path, index=False, engine="openpyxl")
    print(f"Cleaned file saved to: {output_path}")

    return df, {
        "initial": initial_rows,
        "removed_state": removed_state,
        "removed_workable": removed_workable,
        "removed_codes": removed_codes,
        "final": final_rows,
    }
