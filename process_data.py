"""
process_data.py — Final (integrated with macro + HX upload)

Pipeline:
1. Normalize headers.
2. Apply Help sheet mappings (provider -> appointment location, excluded insurances).
3. Filter workable = Y and exclude certain primary insurances.
4. Apply Allocation Priority (NP/FU logic).
5. Assign agents (round-robin).
6. Output HX CSV, warnings, and allocation debug.

Compatible with macro-cleaned Excel file.
"""
from typing import List, Optional
import pandas as pd
import re
import os
from datetime import datetime

# ---------- Helper utilities ----------

def _read_excel_auto(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Read CSV or Excel."""
    if path is None:
        raise ValueError("Path is None")
    if str(path).lower().endswith(".csv"):
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        sheet = 0 if sheet_name is None else sheet_name
        return pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False, engine="openpyxl")


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize headers (strip + title-case)."""
    mapping = {}
    for col in df.columns:
        if col is None:
            new = ""
        else:
            new = " ".join(str(col).strip().split())
            new = new.title()
        mapping[col] = new
    return df.rename(columns=mapping)


def get_hx_field_mapping() -> dict:
    """
    Final HX Template Field → Mapped Source Field (based on your provided table).
    """
    return {
        "Organization": "Audentes_Verification",  # static value
        "Primary Insurance Name": "Primary Insurance Name",
        "Provider Name": "Appointment Provider Name",
        "Primary Insurance ID": "Primary Insurance Subscriber No",
        "Patient Account Number": "Patient Acct No",
        "Patient Name": "Patient Name",
        "DOB": "Patient DOB",
        "Date of Service": "Appointment Date",
        "Visit Type": "Visit Type",
        "Appointment Location": "Appointment State",  # mapped from Help sheet
        "Appointment Time": "Appointment Start Time",
        "Physician NPI": "Appointment Provider NPI",
        "Secondary Insurance Name": "Secondary Insurance Name",
        "Secondary Insurance ID": "Secondary Insurance Subscriber No",
        "Tertiary Insurance Name": "Tertiary Insurance Name",
        "Tertiary Insurance ID": "Tertiary Insurance Subscriber No",
        "Allocation Priority": None,  # created dynamically
    }



# ---------- Help sheet loading ----------

def load_help_sheet(template_path: str):
    """
    Load Help sheet safely (even if Excel file is open or sheet name varies).
    Always returns: (help_df, provider_to_location, visit_type_to_workable, excluded_primary)
    """
    help_df = None

    try:
        # Load all sheet names first
        xl = pd.ExcelFile(template_path, engine="openpyxl")
        sheet_names = [s.strip().lower() for s in xl.sheet_names]

        # Find 'help' sheet by case-insensitive match
        help_sheet_name = next((s for s in xl.sheet_names if s.strip().lower() == "help"), None)
        if not help_sheet_name:
            # fallback: pick the first sheet if not found
            help_sheet_name = xl.sheet_names[0]

        help_df = xl.parse(help_sheet_name, dtype=str)
        help_df = _normalize_columns(help_df)
        print(f"Loaded Help sheet: {help_sheet_name} ({len(help_df)} rows)")

    except Exception as e:
        print(f"Warning: Could not read Help sheet from {template_path}: {e}")
        help_df = pd.DataFrame()

    # --- Extract mappings ---
    provider_col = next((c for c in help_df.columns if "provider" in c.lower()), None)
    loc_col = next((c for c in help_df.columns if "state" in c.lower() or "location" in c.lower()), None)

    provider_to_location = {}
    if provider_col and loc_col:
        for _, row in help_df[[provider_col, loc_col]].dropna().iterrows():
            provider = str(row[provider_col]).strip()
            location = str(row[loc_col]).strip()
            if provider:
                provider_to_location[provider.upper()] = location

    # Extract Visit Type -> Workable mapping from Help sheet
    visit_type_col = next((c for c in help_df.columns if "visit type" in c.lower()), None)
    workable_col = next((c for c in help_df.columns if "workable" in c.lower()), None)
    visit_type_to_workable = {}
    if visit_type_col and workable_col:
        for _, row in help_df[[visit_type_col, workable_col]].dropna().iterrows():
            visit_type = str(row[visit_type_col]).strip().upper()
            workable = str(row[workable_col]).strip().upper()
            visit_type_to_workable[visit_type] = workable
        print(f"Found {len(visit_type_to_workable)} Visit Type -> Workable mappings in Help sheet")
    else:
        print("Warning: Visit Type or Workable column not found in Help sheet - workable filtering will be skipped")

    # Extract all Primary Insurance Name values from Help sheet for exclusion
    prim_ins_col = next((c for c in help_df.columns if "primary insurance name" in c.lower()), None)
    excluded_primary = set()
    if prim_ins_col:
        excluded_primary = set(
            str(x).strip().upper()
            for x in help_df[prim_ins_col].dropna().tolist()
            if str(x).strip()
        )
        print(
            f"Found {len(excluded_primary)} Primary Insurance Name values in Help sheet to exclude"
        )
    else:
        print("Warning: Primary Insurance Name column not found in Help sheet - no exclusions will be applied")

    return help_df, provider_to_location, visit_type_to_workable, excluded_primary


# ============================================================
#   ESCALATION FILTERING (BEFORE ALLOCATION PRIORITY)
# ============================================================

def apply_escalation_filter(df, escalation_path):
    """
    Escalation filter:
    - Load escalation file (CSV, XLSX, or XLSM)
    - Find account column (tries: Acc#, Account Number, Account, Acc, Patient Account Number)
    - Compare to Patient Account Number column in main dataset (tries multiple variations)
    - Remove matching rows from main dataset
    """
    if not escalation_path or not os.path.exists(escalation_path):
        print("No escalation file provided – skipping escalation filter")
        return df
    
    try:
        # Load escalation file (CSV, XLSX, or XLSM)
        file_ext = os.path.splitext(escalation_path)[1].lower()
        if file_ext == ".csv":
            # Try multiple encodings for CSV files
            encodings = ["utf-8", "latin-1", "cp1252", "iso-8859-1"]
            esc = None
            for enc in encodings:
                try:
                    esc = pd.read_csv(escalation_path, dtype=str, keep_default_na=False, encoding=enc)
                    print(f"Successfully loaded escalation CSV with encoding: {enc}")
                    break
                except (UnicodeDecodeError, UnicodeError):
                    continue
            if esc is None:
                print(f"ERROR: Could not read CSV file with any encoding. Tried: {encodings}")
                return df
        elif file_ext in [".xlsx", ".xlsm"]:
            # Excel file - load first sheet
            esc = pd.read_excel(escalation_path, sheet_name=0, dtype=str, keep_default_na=False, engine="openpyxl")
        else:
            print(f"ERROR: Unsupported file format: {file_ext}. Expected .csv, .xlsx, or .xlsm")
            return df
        
        esc = _normalize_columns(esc)
        
        # Find account column - try multiple possible column names
        acc_col = None
        # Priority order: Acc#, Account Number, Account, Acc, Patient Account Number
        possible_names = ["acc#", "account number", "account", "acc", "patient account number"]
        for col in esc.columns:
            col_lower = col.strip().lower()
            if col_lower in possible_names:
                acc_col = col
                print(f"Found account column in escalation file: '{col}'")
                break
        
        if acc_col is None:
            print(f"ERROR: Account column not found in escalation file. Available columns: {list(esc.columns)}")
            print(f"Looking for one of: {possible_names}")
            return df
        
        # Collect escalation account numbers - normalize to strings and strip whitespace
        esc_accounts = esc[acc_col].astype(str).str.strip()
        # Remove empty values and invalid entries
        esc_accounts = set(a for a in esc_accounts if a and a.lower() not in ["", "nan", "none", "null"])
        
        if not esc_accounts:
            print("Escalation file has no valid account numbers – skipping")
            return df
        
        print(f"Loaded {len(esc_accounts)} account numbers from escalation file")
        # Debug: show first few account numbers
        sample_accounts = list(esc_accounts)[:5]
        print(f"Sample escalation account numbers: {sample_accounts}")
        
        # Find main DF account column - try multiple possible names
        main_acc_col = None
        possible_main_names = ["patient account number", "patient account", "account number", "account", "patient acct no"]
        for col in df.columns:
            col_lower = col.strip().lower()
            if any(name in col_lower for name in possible_main_names):
                main_acc_col = col
                print(f"Found account column in main dataset: '{col}'")
                break
        
        if main_acc_col is None:
            print(f"ERROR: Patient Account Number column not found in main dataset. Available columns: {list(df.columns)[:10]}...")
            return df
        
        # Filter out matches - normalize both sides for comparison
        before = len(df)
        df["_acc_temp"] = df[main_acc_col].astype(str).str.strip()
        # Count matches before filtering for debugging
        matches = df["_acc_temp"].isin(esc_accounts)
        match_count = matches.sum()
        print(f"Found {match_count} matching account numbers to filter out")
        
        if match_count > 0:
            # Show sample of accounts being filtered
            sample_matches = df[matches]["_acc_temp"].head(5).tolist()
            print(f"Sample accounts being filtered: {sample_matches}")
        
        df = df[~matches]
        df = df.drop(columns=["_acc_temp"])
        removed = before - len(df)
        print(f"Escalation filter removed {removed} rows (from {before} to {len(df)})")
        return df
        
    except Exception as e:
        print(f"Error applying escalation filter: {e} – skipping escalation filter.")
        import traceback
        print(traceback.format_exc())
        return df



# ---------- Workable + exclusion filters ----------

def check_workable_and_exclusions(df: pd.DataFrame, visit_type_to_workable: dict, excluded_primary_ins: set):
    df = df.copy()
    warnings = []

    # Workable - Look up Visit Type in Help sheet mapping
    visit_type_col = next((c for c in df.columns if "visit type" in c.lower()), None)
    if visit_type_col and visit_type_to_workable:
        # Map each row's Visit Type to its Workable status from Help sheet
        df_visit_normalized = df[visit_type_col].astype(str).str.strip().str.upper()
        df_workable_status = df_visit_normalized.map(visit_type_to_workable)
        
        # Exclude rows where Workable = "N" (only exclude if explicitly "N", keep if missing from mapping)
        mask_n = df_workable_status.str.upper() == "N"
        if mask_n.any():
            excluded_count = mask_n.sum()
            w = df[mask_n].copy()
            w["_warning_reason"] = "Workable = N (from Help sheet)"
            warnings.append(w)
            print(f"Excluding {excluded_count} rows with Visit Type having Workable = N in Help sheet")
        df = df[~mask_n].reset_index(drop=True)
    elif visit_type_col:
        print("Warning: Visit Type -> Workable mapping not available - skipping workable filter")

    # Excluded insurances - match against Primary Insurance Name values from Help sheet
    prim_col = next((c for c in df.columns if "primary insurance" in c.lower()), None)
    if prim_col and excluded_primary_ins:
        # Normalize values for comparison (strip whitespace, case-insensitive)
        df_prim_normalized = df[prim_col].astype(str).str.strip().str.upper()
        mask_excl = df_prim_normalized.isin(excluded_primary_ins)
        if mask_excl.any():
            excluded_count = mask_excl.sum()
            w2 = df[mask_excl].copy()
            w2["_warning_reason"] = "Excluded Primary Insurance"
            warnings.append(w2)
            print(f"Excluding {excluded_count} rows with Primary Insurance Name matching Help sheet values")
        df = df[~mask_excl].reset_index(drop=True)

    if warnings:
        warn_df = pd.concat(warnings, axis=0, ignore_index=True)
    else:
        warn_df = pd.DataFrame(columns=list(df.columns) + ["_warning_reason"])

    return df, warn_df


# ---------- Visit Status Filter (PEN/PR only) ----------

def apply_visit_status_filter(df: pd.DataFrame) -> pd.DataFrame:
    """Exclude rows where Visit Status is INS VER : Insurance Verified."""
    df = df.copy()
    visit_col = next((c for c in df.columns if "visit status" in c.lower()), None)
    
    if visit_col:
        df["_visit_status_u"] = df[visit_col].astype(str).str.strip().str.upper()
        before = len(df)
        df = df[df["_visit_status_u"] != "INS VER : INSURANCE VERIFIED"]
        df = df.drop(columns=["_visit_status_u"])
        print(f"Visit Status filter (exclude INS VER): {before} -> {len(df)} rows")
    else:
        print("WARNING: Visit Status column not found – skipping INS VER exclusion")
    
    return df


# ---------- Remove WC from Primary Insurance Name ----------

def remove_wc_visit_type(df: pd.DataFrame) -> pd.DataFrame:
    """Remove rows where Primary Insurance Name contains 'WC'."""
    df = df.copy()
    ins_col = next((c for c in df.columns if "primary insurance name" in c.lower()), None)
    
    if ins_col:
        df["_ins_u"] = df[ins_col].astype(str).str.upper()
        before = len(df)
        df = df[~df["_ins_u"].str.contains("WC", na=False)]
        df = df.drop(columns=["_ins_u"])
        print(f"WC Primary Insurance removal: {before} -> {len(df)} rows")
    else:
        print("WARNING: Primary Insurance Name column not found – cannot remove WC")
    
    return df


# ---------- Allocation Logic ----------

def _get_visit_type_series(df: pd.DataFrame) -> pd.Series:
    """Extract visit type column as a series for robust matching."""
    vt_col = next((c for c in df.columns if "visit type" in c.lower()), None)
    if vt_col:
        return df[vt_col].astype(str).fillna("")
    return pd.Series([""] * len(df))


def _assign_allocation_priority(df: pd.DataFrame, np_cycle: int = 8, fu_cycle: int = 8) -> pd.DataFrame:
    """
    Assign allocation priorities with corrected sorting:
      Sort order for all groups:
        1. Date of Service (oldest → newest)
        2. Provider Name (A → Z)
        3. Appointment Location (A → Z)

    NP (New Patients):
      - Global sort by DOS > Provider > Location
      - Assign 1..8 cycle globally
      - Reorder into 111122223333... blocks globally.

    FU (Follow-ups):
      - For each Appointment Location (state):
          - Sort by DOS > Provider > Location
          - Assign 1..8 cycle within that state
          - Reorder into 111122223333... per state
      - Concatenate states alphabetically.
    """
    df = df.copy().reset_index(drop=True)
    df["_orig_index"] = df.index

    # Identify Visit Type (determine NP/FU)
    visit_type_series = _get_visit_type_series(df)
    df["Allocation Group"] = visit_type_series.apply(lambda x: "NP" if "new" in str(x).lower() else "FU")

    # Parse DOS safely
    dos_col = next((c for c in df.columns if "date of service" in c.lower() or "appointment date" in c.lower()), None)
    if dos_col:
        df["_dos_parsed"] = pd.to_datetime(df[dos_col], errors="coerce")
    else:
        df["_dos_parsed"] = pd.NaT

    # Identify Provider and Location columns
    provider_col = next((c for c in df.columns if "provider" in c.lower()), None)
    location_col = next((c for c in df.columns if "appointment location" in c.lower() or "appointment state" in c.lower()), None)

    if not provider_col:
        df["Provider Name"] = ""
        provider_col = "Provider Name"

    if not location_col:
        df["Appointment Location"] = ""
        location_col = "Appointment Location"

    # ---------- NP Logic ----------
    df_np = df[df["Allocation Group"] == "NP"].copy()

    if not df_np.empty:
        # Detect correct provider column
        provider_candidates = [
            "Provider Name", "Appointment Provider Name", "Provider",
            "Appt Provider Name"
        ]
        np_provider_col = next((c for c in provider_candidates if c in df_np.columns), provider_col)

        # Detect correct location column (fallback allowed)
        location_candidates = [
            "Appointment Location", "Appointment State",
            "Location", "State"
        ]
        np_location_col = next((c for c in location_candidates if c in df_np.columns), location_col)

        # FALLBACK: if location column missing, create empty string
        if np_location_col not in df_np.columns:
            df_np[np_location_col] = ""

        # FALLBACK: provider must exist (never drop NP rows)
        if np_provider_col not in df_np.columns:
            df_np[np_provider_col] = ""

        # NP: Sort by Location -> Provider (A->Z for both)
        df_np = df_np.sort_values(
            by=[np_location_col, np_provider_col],
            ascending=[True, True],
            kind="stable"
        ).reset_index(drop=True)

        # Apply math logic safely
        total = len(df_np)
        base = total // np_cycle
        extra = total % np_cycle

        counts = [base + (1 if i < extra else 0) for i in range(np_cycle)]

        alloc_seq_list = []
        for seq, count in enumerate(counts, start=1):
            alloc_seq_list.extend([seq] * count)

        alloc_seq_list = alloc_seq_list[:total]

        df_np["_alloc_seq"] = alloc_seq_list
        df_np["_alloc_group"] = "NP"
    else:
        df_np = df_np.copy()


    # ---------- FU Logic ----------
    # -------- FOLLOW-UP (FU) LOGIC: Provider A→Z, Location A→Z, then Math Distribution --------

    ##############################
#   FOLLOW-UP (FU) LOGIC
##############################

    df_fu = df[df["Allocation Group"] == "FU"].copy()

    if not df_fu.empty:

        fu_cycle = 8

        fu_final = []

        # Determine provider and location columns within FU subset
        fu_provider_col = provider_col if provider_col in df_fu.columns else "Provider Name"
        fu_location_col = location_col if location_col in df_fu.columns else "Appointment Location"

        # Process state-wise
        for state in sorted(df_fu[fu_location_col].fillna("").unique()):

            state_df = df_fu[df_fu[fu_location_col] == state].copy()

            # FU: Sort by Provider -> Location (A->Z for both) within each state
            state_df = state_df.sort_values(by=[fu_provider_col, fu_location_col], ascending=[True, True], kind="stable").reset_index(drop=True)

            # ----- ALLOCATION MATH -----

            total = len(state_df)

            # counts per bucket (1..8)
            base = total // fu_cycle
            extra = total % fu_cycle

            # e.g. total=101 → [13,13,13,13,13,13,12,12]
            counts = [base + (1 if i < extra else 0) for i in range(fu_cycle)]

            # Build the allocation sequence
            seq = []
            for alloc_num, count in enumerate(counts, start=1):
                seq.extend([alloc_num] * count)

            # Assign final sequence
            state_df["_alloc_seq"] = seq
            state_df["_alloc_group"] = "FU"

            fu_final.append(state_df)

        df_fu = pd.concat(fu_final, ignore_index=True)
    else:
        df_fu = df_fu.copy()

    # ---------- Combine NP + FU ----------
    combined = pd.concat([df_np, df_fu], ignore_index=True, sort=False)

    # ---- BUILD ALLOCATION PRIORITY CODE ----
    def make_code(prefix, seq):
        if pd.isna(seq) or prefix not in ["NP", "FU"]:
            return ""
        return f"{prefix}{int(seq):03d}"

    combined["Allocation Priority"] = combined.apply(
        lambda r: make_code(r.get("_alloc_group"), r.get("_alloc_seq")),
        axis=1
    )

    return combined



# ---------- Agent Assignment ----------

def assign_agents(df: pd.DataFrame):
    agents = ["Agent-1","Agent-2","Agent-3","Agent-4","Agent-5","Agent-6","Agent-7","Agent-8"]
    prov_col = next((c for c in df.columns if "provider" in c.lower()), None)
    mapping = {}
    assigned = []
    rr = 0
    for prov in df[prov_col].fillna("").tolist():
        if prov in mapping:
            assigned.append(mapping[prov])
        else:
            mapping[prov] = agents[rr % len(agents)]
            assigned.append(mapping[prov])
            rr += 1
    df["Assigned Agent"] = assigned
    return df


# ---------- Build HX output ----------

def build_hx_csv(df: pd.DataFrame, out_dir: str, mapping: dict) -> str:
    """
    Build final HX CSV based on provided mapping.
    Ensures column renaming and inclusion match HX template exactly.
    """
    os.makedirs(out_dir, exist_ok=True)

    # Clean mapping keys (remove BOM, strip whitespace)
    mapping_clean = {}
    for hx_field, src_field in mapping.items():
        clean_key = hx_field.strip().replace('\ufeff', '').strip()
        mapping_clean[clean_key] = src_field
    out_cols = list(mapping_clean.keys())
    
    # Date fields that need mm/dd/yyyy formatting
    date_fields = {"DOB", "Date of Service"}
    
    def format_date(value, field_name):
        """Convert date to mm/dd/yyyy format."""
        if pd.isna(value) or str(value).strip() == "" or str(value).strip().lower() in ["nan", "none", "nat"]:
            return ""
        val_str = str(value).strip()
        # If already in mm/dd/yyyy format, return as-is
        if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', val_str):
            # Normalize to ensure 2-digit month/day
            parts = val_str.split('/')
            if len(parts) == 3:
                month, day, year = parts
                return f"{int(month):02d}/{int(day):02d}/{year}"
            return val_str
        try:
            # Try parsing as datetime with multiple formats
            dt = pd.to_datetime(value, errors='coerce', infer_datetime_format=True)
            if pd.isna(dt):
                # Try parsing common date string formats
                for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d', '%m-%d-%Y', '%d-%m-%Y']:
                    try:
                        dt = pd.to_datetime(value, format=fmt, errors='coerce')
                        if not pd.isna(dt):
                            break
                    except:
                        continue
            if pd.isna(dt):
                return val_str
            return dt.strftime("%m/%d/%Y")
        except Exception as e:
            return val_str
    
    output = {}
    for hx_field in out_cols:
        src_field = mapping_clean[hx_field]
        if src_field is None:
            output[hx_field] = df.get("Allocation Priority", "")
        elif src_field == "Audentes_Verification":
            output[hx_field] = ["Audentes_Verification"] * len(df)
        elif src_field in df.columns:
            if hx_field in date_fields:
                # Format dates to mm/dd/yyyy
                output[hx_field] = df[src_field].apply(lambda x: format_date(x, hx_field))
            else:
                output[hx_field] = df[src_field].fillna("").astype(str)
        else:
            alt = next((c for c in df.columns if c.strip().lower() == src_field.strip().lower()), None)
            if alt:
                if hx_field in date_fields:
                    # Format dates to mm/dd/yyyy
                    output[hx_field] = df[alt].apply(lambda x: format_date(x, hx_field))
                else:
                    output[hx_field] = df[alt].fillna("").astype(str)
            else:
                output[hx_field] = [""] * len(df)

    hx_df = pd.DataFrame(output)[out_cols]
    # Ensure all column headers are clean (no BOM, no leading/trailing spaces, no zero-width chars)
    hx_df.columns = [str(c).strip().replace('\ufeff', '').replace('\u200b', '').replace('\u200c', '').replace('\u200d', '').strip() for c in hx_df.columns]
    
    # Post-process date columns to ensure mm/dd/yyyy format
    for date_col in date_fields:
        if date_col in hx_df.columns:
            hx_df[date_col] = hx_df[date_col].apply(lambda x: format_date(x, date_col))
    
    # Verify Organization column exists and is first
    if 'Organization' not in hx_df.columns:
        # Try to find it with case-insensitive or spacing variations
        org_col = next((c for c in hx_df.columns if c.strip().lower() == 'organization'), None)
        if org_col:
            # Rename to exact match
            hx_df = hx_df.rename(columns={org_col: 'Organization'})
    
    # Ensure Organization is the first column
    if 'Organization' in hx_df.columns:
        cols = ['Organization'] + [c for c in hx_df.columns if c != 'Organization']
        hx_df = hx_df[cols]
        print(f"Verified 'Organization' column is first. Columns: {list(hx_df.columns)[:5]}...")
    else:
        print(f"WARNING: 'Organization' column not found! Available columns: {list(hx_df.columns)}")

    out_path = os.path.join(out_dir, f"HX_Final_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    # Use utf-8 instead of utf-8-sig to avoid BOM issues that HealthX might not handle
    hx_df.to_csv(out_path, index=False, encoding="utf-8")

    print(f"Final HX CSV created with {len(hx_df)} rows -> {out_path}")
    return out_path


# ---------- Post-Macro Filters ----------

def post_macro_filters(df: pd.DataFrame, escalation_path: Optional[str] = None) -> pd.DataFrame:
    """
    Apply post-macro business rules based on BRD:
    1. Keep Visit Status only 'PEN' or 'PR'
    2. Exclude payors containing 'WC'
    3. Exclude accounts from Escalation Tracker file
    4. Remove specific Status and Categorization entries
    """
    df = df.copy()
    initial_rows = len(df)
    print(f"Applying post-macro filters... (initial rows = {initial_rows})")

    # --- 1. Visit Status filter ---
    # Look for Visit Status column specifically (not just any "status" column)
    visit_col = next((c for c in df.columns if "visit status" in c.lower()), None)
    if visit_col:
        before = len(df)
        # Get unique values for debugging
        unique_vals = df[visit_col].astype(str).str.strip().str.upper().unique()
        print(f"Found Visit Status column: '{visit_col}' with values: {sorted(unique_vals[:10])}")
        # Filter for values that start with PEN or PR (handles "PEN : PENDING", "PR : PENDING REFERRAL", etc.)
        visit_status_upper = df[visit_col].astype(str).str.strip().str.upper()
        mask = visit_status_upper.str.startswith("PEN") | visit_status_upper.str.startswith("PR")
        df = df[mask]
        print(f"After Visit Status filter (PEN/PR only): {len(df)} rows (removed {before - len(df)})")
        if len(df) == 0:
            print(f"WARNING: All rows filtered out! Original values were: {sorted(unique_vals)}")
    else:
        print("Warning: Visit Status column not found, skipping filter")

    # --- 2. Exclude Primary Insurance containing 'WC' ---
    ins_col = next((c for c in df.columns if "primary insurance name" in c.lower()), None)
    if ins_col:
        before = len(df)
        df = df[~df[ins_col].astype(str).str.upper().str.contains("WC", na=False)]
        print(f"After excluding WC payors: {len(df)} rows (removed {before - len(df)})")
    else:
        print("Warning: Primary Insurance Name column not found, skipping WC filter")

    # --- 3. Escalation filtering is now handled by apply_escalation_filter() in run_pipeline() ---
    # (Removed old escalation lookup - now using proper status/DOS filtering in apply_escalation_filter)
    print("Escalation filtering will be applied later in the pipeline with proper status/DOS filtering")

    # --- 4. Remove certain 'Status' and 'Categorization' entries ---
    before = len(df)
    if before > 0:  # Only apply if we still have rows
        for c in df.columns:
            cl = c.lower()
            if "status" in cl and c != visit_col:  # Don't filter the Visit Status column we already filtered
                mask = df[c].astype(str).str.contains("Escalated on Smartsheet", case=False, na=False) | \
                       df[c].astype(str).str.contains("Escalated on Teams", case=False, na=False)
                df = df[~mask]
            if "categorization" in cl:
                df = df[~df[c].astype(str).str.contains("Phreesia", case=False, na=False)]
        print(f"After removing escalation & categorization statuses: {len(df)} rows (removed {before - len(df)})")
    else:
        print("Skipping status/categorization filter: no rows remaining")

    print(f"Post-macro filtration complete: {initial_rows} -> {len(df)} rows\n")
    return df


# ============================================================
#   6. ESCALATION FILTERING (BEFORE ALLOCATION PRIORITY)
# ============================================================

# def apply_escalation_filter(df, escalation_path):
#     """
#     Remove rows from df where Patient Account Number
#     matches any Acc# in the escalation file.
#     """
#     if not escalation_path or not os.path.exists(escalation_path):
#         print("No escalation file provided – skipping escalation filter")
#         return df
    
#     try:
#         # Load escalation sheet (CSV or Excel)
#         if str(escalation_path).lower().endswith(".csv"):
#             esc = pd.read_csv(escalation_path, dtype=str, keep_default_na=False)
#         else:
#             esc, _ = load_escalation_sheet(escalation_path)
#         esc = _normalize_columns(esc)
        
#         # Find Acc# column
#         acc_col = next((c for c in esc.columns if c.lower().strip() in ["acc#", "acc", "account", "patient account number"]), None)
#         if acc_col is None:
#             acc_col = esc.columns[0]  # fallback
#             print(f"WARNING: Could not find Acc# column. Using first column: '{acc_col}'")
        
#         # Collect escalation account numbers
#         esc_accounts = esc[acc_col].astype(str).str.strip()
#         esc_accounts = set(a for a in esc_accounts if a not in ["", "nan", "none", "null"])
        
#         if not esc_accounts:
#             print("Escalation file has no valid account numbers – skipping")
#             return df
        
#         # Find main DF account column
#         main_acc_col = next((c for c in df.columns if "patient account" in c.lower()), None)
#         if main_acc_col is None:
#             print("Main DF missing Patient Account Number column – skipping escalation filter")
#             return df
        
#         # Filter out matches
#         before = len(df)
#         df["_acc_temp"] = df[main_acc_col].astype(str).str.strip()
#         df = df[~df["_acc_temp"].isin(esc_accounts)]
#         df = df.drop(columns=["_acc_temp"])
#         print(f"Escalation filter removed {before - len(df)} rows")
#         return df
        
#     except Exception as e:
#         print(f"Error applying escalation filter: {e} – skipping escalation filter.")
#         return df


# ---------- Main Pipeline ----------
from macro import audentes_verification_cleaned

def run_pipeline(cleaned_file: str, template_wb: Optional[str], out_dir: str, escalation_file_path: Optional[str] = None, log_path: Optional[str] = None):
    """
    Run full pipeline after macro cleanup.
    Steps:
      1. Read cleaned macro output and normalize.
      2. Apply Visit Status filter (exclude INS VER : Insurance Verified).
      3. Remove WC from Visit Type.
      4. Load Help sheet mappings.
      5. Map Appointment Location (provider -> state).
      6. Apply workable + excluded primary insurance filters.
      7. Apply escalation filter (Acc# comparison) - BEFORE allocation priority.
      8. Perform allocation + agent assignment.
      9. Build final HX CSV + debug logs.
    """
    print("Loading cleaned macro output...")
    df = _read_excel_auto(cleaned_file)
    df = _normalize_columns(df)

    # --- Step 1: Visit Status Filter (exclude INS VER) ---
    print("Applying Visit Status filter (exclude INS VER)...")
    df = apply_visit_status_filter(df)

    # --- Step 2: Remove WC from Primary Insurance Name ---
    print("Removing WC from Primary Insurance Name...")
    df = remove_wc_visit_type(df)

    # --- Step 3: Load Help sheet mappings ---
    print("Loading Help sheet mappings...")
    help_df, provider_to_location, visit_type_to_workable, excluded_primary = load_help_sheet(template_wb)

    # --- Step 4: Map Appointment Location (from Help sheet) ---
    if "Appointment Provider Name" in df.columns and provider_to_location:
        df["Appointment Location"] = (
            df["Appointment Provider Name"]
            .astype(str)
            .str.strip()
            .str.upper()
            .map(provider_to_location)
        )
        print("Appointment Location mapped from Help sheet.")
    else:
        print("Could not map Appointment Location — provider names not found in Help sheet.")

    # --- Step 5: Apply Workable and Primary Insurance exclusions ---
    print("Checking Workable status and excluded Primary Insurance Names...")
    df_filtered, warnings = check_workable_and_exclusions(df, visit_type_to_workable, excluded_primary)

    # --- Step 6: Apply Escalation Filtering (BEFORE ALLOCATION PRIORITY) ---
    if escalation_file_path:
        print("Applying escalation filter (Acc# comparison)...")
        df_filtered = apply_escalation_filter(df_filtered, escalation_file_path)
    else:
        print("No escalation file provided, skipping escalation filter.")

    # --- Step 7: Allocation + Agent assignment ---
    df_alloc = _assign_allocation_priority(df_filtered)
    df_agents = assign_agents(df_alloc)

    # --- Step 8: Build final HX output ---
    print("Building final HX file...")
    out_path = build_hx_csv(df_agents, out_dir, get_hx_field_mapping())

    # --- Step 9: Save warnings and debug ---
    warnings_path = os.path.join(out_dir, "warnings.csv")
    if warnings is not None and not warnings.empty:
        warnings.to_csv(warnings_path, index=False)
    else:
        pd.DataFrame(columns=["_warning_reason"]).to_csv(warnings_path, index=False)

    debug_cols = [
        "Patient Account Number", "Patient Name", "Appointment Provider Name",
        "Appointment Location", "Appointment Date", "Visit Type",
        "Allocation Group", "Allocation Priority", "_alloc_seq", "Assigned Agent"
    ]
    existing_debug_cols = [c for c in debug_cols if c in df_agents.columns]
    debug_df = df_agents[existing_debug_cols] if existing_debug_cols else df_agents.head(0)
    debug_path = os.path.join(out_dir, "allocation_debug.csv")
    debug_df.to_csv(debug_path, index=False)

    print(f" HX file ready: {out_path}")
    print(f" Warnings: {warnings_path}")
    print(f" Debug: {debug_path}")

    return {
        "hx_csv": out_path,
        "warnings": warnings_path,
        "debug": debug_path,
        "processed_count": len(df_agents)
    }


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Run HX allocation process")
    parser.add_argument("--input", required=True, help="Cleaned macro Excel")
    parser.add_argument("--wb", required=True, help="Help workbook")
    parser.add_argument("--outdir", default="outputs", help="Output folder")
    args = parser.parse_args()
    print(run_pipeline(args.input, args.wb, args.outdir))
