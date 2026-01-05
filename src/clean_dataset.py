import pandas as pd
#import fuzzywuzzy
from fuzzywuzzy import process, fuzz
import numpy as np
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
import re


from dataclasses import dataclass, field
from typing import List, Optional

@dataclass
class MsgLog:
    info: List[str] = field(default_factory=list)
    warning: List[str] = field(default_factory=list)

    def add_info(self, msg: str):
        self.info.append(str(msg))

    def add_warn(self, msg: str):
        self.warning.append(str(msg))

## ---------------------------------------------------------------------------------
def get_and_standardize_uuid(df, log: Optional[MsgLog] = None):
    """
    Identify and standardize UUID column.

    Priority order:
    1. 'uuid'
    2. 'submission_uuid'
    3. '*_uuid' or '*_submission_uuid'
    4. any column containing 'uuid'

    Excludes:
    - 'edu_uuid'

    If found, renames the column to 'uuid'.
    """
    def info(m): 
        if log: log.add_info(m)
    def warn(m):
        if log: log.add_warn(m)
    cols = list(df.columns)
    cols_lower = {c.lower(): c for c in cols}

    # Helper to exclude edu_uuid explicitly
    def valid(col):
        return 'edu_uuid' not in col.lower()

    # Exact match: uuid
    if 'uuid' in cols_lower and valid(cols_lower['uuid']):
        return df.rename(columns={cols_lower['uuid']: 'uuid'})

    # Exact match: submission_uuid
    if 'submission_uuid' in cols_lower and valid(cols_lower['submission_uuid']):
        return df.rename(columns={cols_lower['submission_uuid']: 'uuid'})

    # Suffix matches
    suffix_matches = [
        c for c in cols
        if valid(c) and (
            c.lower().endswith('_uuid') or
            c.lower().endswith('_submission_uuid')
        )
    ]
    if suffix_matches:
        return df.rename(columns={suffix_matches[0]: 'uuid'})

    # Fallback: anything containing 'uuid'
    fallback_matches = [
        c for c in cols
        if 'uuid' in c.lower() and valid(c)
    ]
    if fallback_matches:
        return df.rename(columns={fallback_matches[0]: 'uuid'})

    # Hard stop if nothing is found
    raise ValueError(
        f"\n❌ No UUID column found in the education loop.\n\n"
        "This column is essential to match household/parents information "
        "with children education indicators.\n\n"
        "➡️ Please identify the column used for this matching in your dataset "
        "and rename it to 'uuid' in the MSNA file.\n"
    )
## ---------------------------------------------------------------------------------

##----------------------------------------------------------------------------------
def custom_to_datetime(date_str):
    try:
        # Try the default date parsing first
        return pd.to_datetime(date_str, errors='coerce')
    except:
        try:
            # Handle the 'Y-m-d H:M:S.f+TZ' format
            return pd.to_datetime(date_str, format='%Y-%m-%dT%H:%M:%S.%f%z', errors='coerce')
        except:
            try:
                # Handle the 'Y-m-d H:M:S.f' format without time zone
                return pd.to_datetime(date_str, format='%Y-%m-%d %H:%M:%S.%f', errors='coerce')
            except:
                try:
                    # Handle the 'dd/mm/yyyy' format
                    return pd.to_datetime(date_str, format='%d/%m/%Y', errors='coerce')
                except:
                    # Return NaT if all parsing attempts fail
                    return pd.NaT
##----------------------------------------------------------------------------------


## ---------------------------------------------------------------------------------
def standardize_uuid_household(df, log: Optional[MsgLog] = None):
    def info(m): 
        if log: log.add_info(m)
    def warn(m):
        if log: log.add_warn(m)
    cols = list(df.columns)
    cols_lower = {c.lower(): c for c in cols}

    # Priority: uuid (exact)
    if 'uuid' in cols_lower:
        return df.rename(columns={cols_lower['uuid']: 'uuid'})

    # Next: any column ending with _uuid
    suffix_matches = [c for c in cols if c.lower().endswith('_uuid')]
    if suffix_matches:
        return df.rename(columns={suffix_matches[0]: 'uuid'})

    # Fallback: anything containing uuid
    fallback_matches = [c for c in cols if 'uuid' in c.lower()]
    if fallback_matches:
        return df.rename(columns={fallback_matches[0]: 'uuid'})

    raise ValueError(
        "\n❌ No UUID column found in the household loop.\n\n"
        "This is required to identify households and merge household-level variables.\n\n"
        "➡️ Please rename the household identifier column in the MSNA file to 'uuid' (or '*_uuid').\n"
    )
## ---------------------------------------------------------------------------------


## ---------------------------------------------------------------------------------
## finding admin        
def extract_number(col_name: str) -> int|None:
    nums = re.findall(r'\d+', col_name or "")
    return int(nums[-1]) if nums else None

def looks_like_pcode(series, min_fraction=0.8):
    p = re.compile(r'^(?=.*[A-Za-z])(?=.*\d)[A-Za-z0-9]+$')
    vals = series.dropna().astype(str)
    if len(vals)==0: return False
    return (vals.str.match(p).sum() / len(vals)) >= min_fraction

def find_best_match(admin_target: str,
                    household_df,
                    similarity_threshold=70,
                    pcode_fraction=0.8,
                    fallback_threshold=50) -> str:
    cols = list(household_df.columns)
    target_num = extract_number(admin_target)

    # — Step 0: All columns that *share* that number in their name
    num_cols = [c for c in cols if extract_number(c)==target_num]

    # — Step 1: Among those, pick any with “code”
    code_cols = [c for c in num_cols if 'code' in c.lower()]
    if code_cols:
        candidates = code_cols
    elif num_cols:
        candidates = num_cols
    else:
        candidates = []

    # — Step 2: If we still have nothing, use your fuzzy + number + code logic:
    if not candidates:
        # fuzzy find “similar” by text
        base = re.sub(r'\d+','', admin_target).lower()
        similar = [c for c in cols
                   if fuzz.partial_ratio(base, re.sub(r'\d+','',c).lower())
                      >= similarity_threshold]

        # among “similar” pick same-number, then code, then rest
        same_num = [c for c in similar if extract_number(c)==target_num]
        code_in_same = [c for c in same_num if 'code' in c.lower()]

        if code_in_same:
            candidates = code_in_same
        elif same_num:
            candidates = same_num
        else:
            candidates = similar

    # — Step 3: fallback global fuzzy if we still have nothing
    if not candidates:
        matches = process.extract(admin_target, cols,
                                  scorer=fuzz.partial_ratio,
                                  limit=len(cols))
        candidates = [m[0] for m in matches if m[1]>=fallback_threshold]

    # — Step 4: If STILL nothing, just take the single best
    if not candidates:
        return process.extractOne(admin_target, cols)[0]

    # Debug print
    print("→ ordered candidates:", candidates)

    # — Step 5: pick first whose values *look* like P-codes
    for c in candidates:
        if looks_like_pcode(household_df[c], min_fraction=pcode_fraction):
            print(f"→ picking {c} (passed P-code check)")
            return c

    # — Step 6: fallback to single best fuzzy match
    print("→ none passed P-code check; falling back")
    return process.extractOne(admin_target, cols,
                              scorer=fuzz.partial_ratio)[0]
def list_pcode_like_columns(df, min_fraction=0.8):
    """
    Return a list of columns where values look like PCODEs
    (at least min_fraction of non-null values match the PCODE regex).
    """
    pcode_cols = []
    for c in df.columns:
        try:
            if looks_like_pcode(df[c], min_fraction=min_fraction):
                pcode_cols.append(c)
        except Exception:
            # In case a column type causes issues, just skip it
            continue
    return pcode_cols
## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def standardize_weights(household_df,log: Optional[MsgLog] = None):
    """
    Ensure a 'weights' column exists.
    - If a weight column is found (aliases or contains 'weight'), rename it to 'weights'
    - If not found, create weights=1 and print a clear warning message
    """
    def info(m): 
        if log: log.add_info(m)
    def warn(m):
        if log: log.add_warn(m)
    aliases = {"weights", "weight", "weight_final"}
    cols = list(household_df.columns)
    norm = {c: c.strip().lower() for c in cols}

    # 1) If a weights column already exists (any casing/spaces), normalize its name
    existing_weights = next((c for c in cols if norm[c] == "weights"), None)
    if existing_weights:
        if existing_weights != "weights":
            if "weights" in household_df.columns:
                # avoid duplicate column names
                household_df = household_df.drop(columns=[existing_weights])
            else:
                household_df = household_df.rename(columns={existing_weights: "weights"})
        info("--------------------------- 'weights' column already exists.")
        return household_df

    # 2) Try exact alias match (after normalization)
    found = next((c for c in cols if norm[c] in aliases), None)

    # 3) If not found, try any column that contains 'weight' (e.g. household_weight)
    if not found:
        found = next((c for c in cols if "weight" in norm[c]), None)

    # 4) If found, rename to weights (avoid duplicates)
    if found:
        if "weights" in household_df.columns and found != "weights":
            household_df = household_df.drop(columns=[found])
        else:
            household_df = household_df.rename(columns={found: "weights"})

        info(f"--------------------------- Weight column found ('{found}'), renamed to 'weights'.")
        return household_df

    # 5) If not found: create default + warn (do not stop)
    household_df["weights"] = 1
    warn(
        "⚠️ No weight column found. Creating default 'weights' = 1. "
        "Outputs will be UNWEIGHTED (each household has equal weight). "
        "If your MSNA has sampling weights, please rename the correct column to 'weights' and rerun."
    )

    return household_df
## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def _get_survey_name_and_label_cols(survey_df, log: Optional[MsgLog] = None):
    def info(m): 
        if log: log.add_info(m)
    def warn(m):
        if log: log.add_warn(m)
    # name column
    name_col = next((c for c in survey_df.columns if c.lower() == "name"), None)
    if not name_col:
        raise KeyError("survey_data must contain a 'name' column (XLSForm survey sheet).")

    # all label-like columns (label, label::English (en), label::French (fr), etc.)
    label_cols = [c for c in survey_df.columns if "label" in c.lower()]
    if not label_cols:
        raise KeyError("survey_data has no label column (expected at least 'label' or 'label::...').")

    return name_col, label_cols
## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def resolve_to_survey_name(var, survey_df, log: Optional[MsgLog] = None):

    def info(m): 
        if log: log.add_info(m)
    def warn(m):
        if log: log.add_warn(m)    
    if var in [None, "", "no_indicator"]:
        return var

    name_col, label_cols = _get_survey_name_and_label_cols(survey_df)

    # normalize for comparison
    var_norm = str(var).strip().lower()
    names_norm = survey_df[name_col].astype(str).str.strip().str.lower()

    # 1) Already a real name?
    if (names_norm == var_norm).any():
        return var  # keep original (already a name)

    # 2) Is it a label?
    for lc in label_cols:
        lab_norm = survey_df[lc].astype(str).str.strip().str.lower()
        hit = survey_df.loc[lab_norm == var_norm, name_col]
        if len(hit) > 0:
            resolved = hit.iloc[0]
            warn(
                f"⚠️ Variable '{var}' looks like a LABEL (found in survey_data['{lc}']) "
                f"not a NAME. Using the corresponding name: '{resolved}'. "
                f"Please export MSNA with question *names* (not labels)."
            )
            return resolved

    # 3) Not found in names or labels
    warn(
        f"⚠️ Variable '{var}' not found in survey_data names or labels edu loop). "
        "Check the XLSForm or the MSNA export format."
    )
    return var
## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def check_labeled(survey_data, access_var, barrier_var, name_col, label_col, log: Optional[MsgLog] = None):
    """
    Checks ONLY access_var and barrier_var against survey_data[name_col] and survey_data[label_col].
    - Never raises errors (returns False if columns missing)
    - Returns True if any var is found in label_col but not in name_col
    """
    def info(m): 
        if log: log.add_info(m)
    def warn(m):
        if log: log.add_warn(m)
    if name_col not in survey_data.columns or label_col not in survey_data.columns:
        print(f"⚠️ survey_data missing '{name_col}' or '{label_col}'. Skipping label/name check.")
        return False

    names = set(survey_data[name_col].dropna().astype(str).str.strip().str.lower())
    labels = set(survey_data[label_col].dropna().astype(str).str.strip().str.lower())

    labeled_dt = False

    def _check_one(var, var_name):
        nonlocal labeled_dt
        if var in [None, "", "no_indicator"]:
            return

        v = str(var).strip().lower()

        if v in labels and v not in names:
            labeled_dt = True
            info(f"⚠️ {var_name}='{var}' found in survey_data['{label_col}'] (not in '{name_col}').")
        elif v not in names and v not in labels:
            info(f"⚠️ {var_name}='{var}' NOT found in survey_data['{name_col}'] or ['{label_col}'].")

    _check_one(access_var, "access_var")
    _check_one(barrier_var, "barrier_var")

    return labeled_dt
## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def rename_edu_columns_from_survey_labels(
    edu_data,
    survey_data,
    vars_list,
    name_col="name",
    label_col=None, log: Optional[MsgLog] = None
):
    """
    For each variable in vars_list (strings):
      - if it matches survey_data[label_col], resolve to survey_data[name_col]
      - rename edu_data column from label -> name (case-insensitive safe)
    Returns: (edu_data_renamed, resolved_vars_dict)
    """
    def info(m): 
        if log: log.add_info(m)
    def warn(m):
        if log: log.add_warn(m)
    if label_col is None:
        raise ValueError("label_col must be provided (e.g. 'label::English (en)').")

    if name_col not in survey_data.columns or label_col not in survey_data.columns:
        print(f"⚠️ survey_data missing '{name_col}' or '{label_col}'. Cannot resolve labels -> names.")
        return edu_data, {v: v for v in vars_list}

    # Build mapping label -> name (normalized)
    label_series = survey_data[label_col].dropna().astype(str).str.strip()
    name_series  = survey_data[name_col].dropna().astype(str).str.strip()

    # align lengths safely (survey sheets usually align row-wise; this keeps it simple)
    tmp = survey_data[[name_col, label_col]].copy()
    tmp[name_col] = tmp[name_col].astype(str).str.strip()
    tmp[label_col] = tmp[label_col].astype(str).str.strip()

    # remove empty strings too (not just NaN)
    tmp = tmp[(tmp[name_col] != "") & (tmp[label_col] != "")]

    # check duplicates (case-insensitive)
    dup = tmp[tmp[label_col].str.lower().duplicated(keep=False)]
    if not dup.empty:
        info("⚠️ Duplicate labels found in survey_data; label->name mapping may be ambiguous.")
        # optional: print a few examples
        print(dup[[label_col, name_col]].head(10))

    label_to_name = dict(zip(tmp[label_col].str.lower(), tmp[name_col]))

    names_set = set(tmp[name_col].astype(str).str.strip().str.lower())

    # case-insensitive lookup for edu_data columns
    edu_cols_lower = {c.strip().lower(): c for c in edu_data.columns}

    resolved = {}

    for var in vars_list:
        if var in [None, "", "no_indicator"]:
            resolved[var] = var
            continue

        var_norm = str(var).strip().lower()

        # 1) Already a NAME?
        if var_norm in names_set:
            new_name = str(var).strip()
        # 2) Is it a LABEL?
        elif var_norm in label_to_name:
            new_name = label_to_name[var_norm]
        else:
            info(f"⚠️ '{var}' not found in survey_data['{name_col}'] nor in survey_data['{label_col}']. Keeping as-is.")
            new_name = var

        # Rename column in edu_data if needed
        old_col = edu_cols_lower.get(var_norm, None)
        if old_col is not None and new_name != old_col:
            if new_name in edu_data.columns:
                print(f"⚠️ Cannot rename '{old_col}' -> '{new_name}' because '{new_name}' already exists in edu_data.")
            else:
                edu_data = edu_data.rename(columns={old_col: new_name})
                # update lookup after renaming
                edu_cols_lower = {c.strip().lower(): c for c in edu_data.columns}

        resolved[var] = new_name

    return edu_data, resolved
## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def barrier_is_select_multiple(
    survey_data: pd.DataFrame,
    barrier_var: str,
    name_col: str = "name",
    label_col: str | None = None
) -> bool:
    """
    True if barrier_var corresponds to a select_multiple question in survey_data['type'].
    Matches by NAME first; if not found, optionally matches by LABEL.
    """
    if barrier_var in [None, "", "no_indicator"]:
        return False

    if name_col not in survey_data.columns or "type" not in survey_data.columns:
        print("⚠️ survey_data missing 'name' and/or 'type'. Cannot detect select_multiple.")
        return False

    b = str(barrier_var).strip().lower()

    s_name = survey_data[name_col].astype(str).str.strip().str.lower()
    hit = survey_data.loc[s_name == b]

    # optional label matching
    if hit.empty and label_col and label_col in survey_data.columns:
        s_lab = survey_data[label_col].astype(str).str.strip().str.lower()
        hit = survey_data.loc[s_lab == b]

    if hit.empty:
        print(f"⚠️ barrier_var='{barrier_var}' not found in survey_data['{name_col}']"
              + (f" or ['{label_col}']." if label_col else "."))
        return False

    t = str(hit.iloc[0]["type"]).strip().lower()
    return "select_multiple" in t

## ---------------------------------------------------------------------------------

##--------------------------------------------------------------------------------------------
def find_matching_choices(choices_df, barriers_list, label_var):
    # List to hold the results
    results = []
    
    # Iterate over each barrier in the list
    for barrier in barriers_list:
        # Filter choices where the label_var matches the current barrier
        matched_choices = choices_df[choices_df[label_var] == barrier]
        
        # If no matches are found, add a 'notfound' entry
        if matched_choices.empty:
            result_entry = {'name': 'notfound', 'label': barrier}
            results.append(result_entry)
        else:
            # For each matched choice, create an entry in the results list
            for _, choice in matched_choices.iterrows():
                result_entry = {'name': choice['name'], 'label': barrier}
                results.append(result_entry)
    
    return results
## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def _extract_code_from_name(name):
    """
    Extract the last integer from a choice name (e.g. 'barriers_education_21' -> 21, 'opt_5' -> 5).
    Returns None if no integer found.
    """
    if name is None:
        return None
    nums = re.findall(r"\d+", str(name))
    return int(nums[-1]) if nums else None
## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def build_severity_tbl_from_choices(choice_data, all_barriers_names, label_col,  log: Optional[MsgLog] = None):
    """
    Returns a list of dicts: [{'code': 21, 'label': 'Pregnancy', 'rank': 1}, ...]
    Rank is the position in all_barriers_names (1 = most severe).
    """
    def info(m): 
        if log: log.add_info(m)
    def warn(m):
        if log: log.add_warn(m)    
    # name -> label lookup (safe)
    tmp = choice_data.copy()
    if "name" not in tmp.columns:
        raise KeyError("choice_data must contain a 'name' column.")
    if label_col not in tmp.columns:
        raise KeyError(f"choice_data must contain the label column '{label_col}'.")

    name_to_label = dict(
        zip(
            tmp["name"].astype(str).str.strip(),
            tmp[label_col].astype(str).str.strip()
        )
    )

    severity_tbl = []
    for i, nm in enumerate(all_barriers_names, start=1):  # rank starts at 1
        if nm in [None, "", "notfound"]:
            continue
        code = _extract_code_from_name(nm)
        lab = name_to_label.get(str(nm).strip(), str(nm).strip())  # fallback to name if label missing
        severity_tbl.append({"code": code, "label": lab, "rank": i})

    # keep only rows where code exists (needed for code-extraction path)
    severity_tbl_codes = [r for r in severity_tbl if r["code"] is not None]

    # helpers like in R
    rank_by_code = {r["code"]: r["rank"] for r in severity_tbl_codes}
    label_by_code = {r["code"]: r["label"] for r in severity_tbl_codes}
    labels_in_rank_order = [r["label"] for r in severity_tbl]  # includes fallback labels too

    return severity_tbl, rank_by_code, label_by_code, labels_in_rank_order
## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def pick_most_severe(x, valid_codes_set, rank_by_code, label_by_code, labels_in_rank_order):
    """
    Implements the SAME logic as your R function:
    1) extract codes from x (numbers anywhere)
    2) pick best code by rank
    3) else match by label substrings in severity order
    """
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if not s:
        return np.nan

    s_low = s.lower()

    # 1) Extract codes (1..99 etc). We'll filter to valid_codes_set.
    nums = re.findall(r"(?<!\d)(?:\d+)(?!\d)", s_low)
    nums = [int(n) for n in nums if n.isdigit()]
    nums = [n for n in nums if n in valid_codes_set]

    if nums:
        best_code = min(nums, key=lambda c: rank_by_code.get(c, 10**9))
        return label_by_code.get(best_code, np.nan)

    # 2) No codes -> label substring scan (in severity order)
    for lab in labels_in_rank_order:
        if lab and str(lab).strip():
            if str(lab).lower() in s_low:
                return lab  # return canonical-cased label

    return np.nan
## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def detect_sm_values_are_names_or_labels(
    edu_data: pd.DataFrame,
    barrier_var: str,
    survey_data: pd.DataFrame,
    name_col: str = "name",
    label_col: str | None = None,
    sample_n: int = 200,
    min_hits: int = 5, log: Optional[MsgLog] = None
) -> dict:
    """
    Decide whether edu_data[barrier_var] contains select_multiple values as:
      - 'names' (tokens match survey_data[name_col])
      - 'labels' (substrings match survey_data[label_col])
      - 'unknown'

    Returns a dict with scores + examples.
    """
    def info(m): 
        if log: log.add_info(m)
    def warn(m):
        if log: log.add_warn(m)
    if barrier_var not in edu_data.columns:
        return {"format": "unknown", "reason": f"'{barrier_var}' not in edu_data"}

    if name_col not in survey_data.columns:
        return {"format": "unknown", "reason": f"survey_data missing '{name_col}'"}

    if not label_col or label_col not in survey_data.columns:
        return {"format": "unknown", "reason": f"survey_data missing label_col '{label_col}'"}

    # reference sets/lists
    names_set = set(
        survey_data[name_col].dropna().astype(str).str.strip().str.lower()
        .loc[lambda s: s != ""]
        .tolist()
    )
    labels_list = (
        survey_data[label_col].dropna().astype(str).str.strip().str.lower()
        .loc[lambda s: s != ""]
        .tolist()
    )

    # sample values
    s = edu_data[barrier_var].dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return {"format": "unknown", "reason": "No non-empty values to inspect"}

    sample = s.sample(min(sample_n, len(s)), random_state=0)

    # token split for SM exports
    split_pat = re.compile(r"[,\;\|\n\r\t]+|\s+")

    name_hits = 0
    label_hits = 0
    examples_name = []
    examples_label = []

    for val in sample:
        v = str(val).strip().lower()
        if not v:
            continue

        # --- names: token match (very strong for select_multiple)
        tokens = [t for t in split_pat.split(v) if t]
        if any(t in names_set for t in tokens):
            name_hits += 1
            if len(examples_name) < 3:
                examples_name.append(val)

        # --- labels: substring match (strong if export used labels)
        found_label = False
        for lab in labels_list:
            # ignore tiny labels to avoid accidental matches
            if len(lab) < 6:
                continue
            if lab in v:
                found_label = True
                break
        if found_label:
            label_hits += 1
            if len(examples_label) < 3:
                examples_label.append(val)

    n = len(sample)
    name_score = name_hits / n
    label_score = label_hits / n

    fmt = "unknown"
    # Decision rule: pick the clearer signal
    if name_hits >= min_hits and name_score >= label_score + 0.15:
        fmt = "names"
    elif label_hits >= min_hits and label_score >= name_score + 0.15:
        fmt = "labels"
    elif name_hits >= min_hits and label_hits < min_hits:
        fmt = "names"
    elif label_hits >= min_hits and name_hits < min_hits:
        fmt = "labels"

    return {
        "format": fmt,
        "name_score": name_score,
        "label_score": label_score,
        "name_hits": name_hits,
        "label_hits": label_hits,
        "n_sampled": n,
        "examples_name_like": examples_name,
        "examples_label_like": examples_label
    }

## ---------------------------------------------------------------------------------

## ---------------------------------------------------------------------------------
def build_barrier_so_select_multiple(
    edu_data: pd.DataFrame,
    barrier_var: str,
    names_severity_5: list[str],
    names_severity_4: list[str],
    default_value: str = "barrier_3",
    out_col: str = "barrier_so", log: Optional[MsgLog] = None
) -> pd.DataFrame:
    """
    For select_multiple barrier_var values:
      1) if any item from names_severity_5 is found in the value -> copy the matched item (priority to 5)
      2) else if any item from names_severity_4 is found -> copy matched item
      3) else default_value

    Notes:
    - If multiple matches exist within severity_5 (or 4), returns the first match in the list order.
    - Matching is case-insensitive, substring-based.
    """
    def info(m): 
        if log: log.add_info(m)
    def warn(m):
        if log: log.add_warn(m)
    # normalize candidate lists (remove empties, ensure string)
    sev5 = [str(x).strip() for x in (names_severity_5 or []) if str(x).strip() and str(x).strip() != "notfound"]
    sev4 = [str(x).strip() for x in (names_severity_4 or []) if str(x).strip() and str(x).strip() != "notfound"]

    # for case-insensitive search, compare lower versions but return original candidate string
    sev5_low = [(x, x.lower()) for x in sev5]
    sev4_low = [(x, x.lower()) for x in sev4]

    def pick_one(val):
        if pd.isna(val):
            return default_value
        s = str(val).strip()
        if not s:
            return default_value
        s_low = s.lower()

        # 1) priority severity 5
        for orig, low in sev5_low:
            if low and low in s_low:
                return orig

        # 2) then severity 4
        for orig, low in sev4_low:
            if low and low in s_low:
                return orig

        # 3) default
        return default_value

    edu_data[out_col] = edu_data[barrier_var].apply(pick_one)
    return edu_data
## ---------------------------------------------------------------------------------




########################################################### 
########################################################### 

def clean_make_dataset (country, edu_data, household_data, choice_data, survey_data, 
                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                natural_hazard_var,natural_hazard_var_sev,
                additional_last_var,additional_last_sev,
                additional_2_last_var,additional_2_last_sev,
                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                age_var, gender_var,
                label, 
                admin_var, vector_cycle, start_school, status_var,
                selected_language):

    messages = {"info": [], "warning": []}
    def add_info(msg): messages["info"].append(msg)
    def add_warn(msg): messages["warning"].append(msg)

    ##---------------- 1) rename the uuid columns with uuid  (Find the UUID columns, assuming they exist and taking only the first match for simplicity)
    edu_data = get_and_standardize_uuid(edu_data)
    household_data = standardize_uuid_household(household_data)
    edu_uuid_column = "uuid"
    household_uuid_column =  "uuid"

    ##---------------- 2)  Find or create the household collection date column as 'today' ---
    possible_columns = list(household_data.columns)
    possible_columns_lower = [c.lower() for c in possible_columns]
    today_candidates = [c for c in possible_columns if 'today' in c.lower() or c.lower() == 'today_date']
    start_candidates = [c for c in possible_columns if 'start' in c.lower()]
    household_start_column = None
    # Prioritize "today" (or today_date)
    if today_candidates:
        # Prefer an exact 'today' if present, else take first candidate
        exact_today = next((c for c in today_candidates if c.lower() == 'today'), None)
        household_start_column = exact_today if exact_today else today_candidates[0]
    # Fallback to "start"
    elif start_candidates:
        household_start_column = start_candidates[0]
    # If nothing found: create default 'today'
    else:
        print("No column found, assigning default value 01/06/2026 as data collection day.")
        household_data['today'] = pd.to_datetime("2026-06-01", dayfirst=False)
        household_start_column = 'today'
    # If we found a column, standardize its name to 'today'
    if household_start_column != 'today':
        # Avoid duplicate columns: if 'today' already exists, fill missing values then drop the old one
        if 'today' in household_data.columns:
            household_data['today'] = household_data['today'].combine_first(household_data[household_start_column])
            household_data = household_data.drop(columns=[household_start_column])
        else:
            household_data = household_data.rename(columns={household_start_column: 'today'})
    # Now parse/standardize 'today' and derive month
    household_data['today'] = household_data['today'].apply(custom_to_datetime)
    household_data['today'] = pd.to_datetime(household_data['today'], errors='coerce')
    household_data['month'] = household_data['today'].dt.month


    ##---------------- 3) find the admin column
    admin_target = admin_var
    admin_var = find_best_match(admin_target,  household_data)
    pcode_like_cols = list_pcode_like_columns(household_data, min_fraction=0.8)
    print("PCODE-like columns found:", pcode_like_cols)
    admin_var_found = find_best_match(admin_target, household_data)
    # Rename to a standard column name
    if admin_var_found != "admin_hno":
        if "admin_hno" in household_data.columns:
            # Avoid accidental overwrite / duplicates
            raise ValueError(
                "Column 'admin_hno' already exists in household_data. "
                f"Cannot rename '{admin_var_found}' to 'admin_hno' without overwriting."
            )
        household_data = household_data.rename(columns={admin_var_found: "admin_hno"})
    admin_var = "admin_hno"

    ##---------------- 4) find the weights column
    household_data = standardize_weights(household_data)

    ##---------------- 5) Standardize key variable names (status, age, gender)
    # --- status_var -> pop_status_group (household level) ---
    if status_var in household_data.columns:
        if "pop_status_group" in household_data.columns and status_var != "pop_status_group":
            raise ValueError(
                f"Cannot rename '{status_var}' to 'pop_status_group' because 'pop_status_group' already exists."
            )
        household_data = household_data.rename(columns={status_var: "pop_status_group"})
        status_var = "pop_status_group"
    else:
        raise KeyError(f"'{status_var}' not found in household_data columns.")

    # --- age_var -> ind_age (education loop) ---
    if age_var in edu_data.columns:
        if "ind_age" in edu_data.columns and age_var != "ind_age":
            raise ValueError(
                f"Cannot rename '{age_var}' to 'ind_age' because 'ind_age' already exists."
            )
        edu_data = edu_data.rename(columns={age_var: "ind_age"})
        age_var = "ind_age"
    else:
        raise KeyError(f"'{age_var}' not found in edu_data columns.")

    # --- gender_var -> ind_gender (education loop) ---
    if gender_var in edu_data.columns:
        if "ind_gender" in edu_data.columns and gender_var != "ind_gender":
            raise ValueError(
                f"Cannot rename '{gender_var}' to 'ind_gender' because 'ind_gender' already exists."
            )
        edu_data = edu_data.rename(columns={gender_var: "ind_gender"})
        gender_var = "ind_gender"
    else:
        raise KeyError(f"'{gender_var}' not found in edu_data columns.")

    ##---------------- 6) check if it is a labeled dataset
    labeled_dt = check_labeled(survey_data=survey_data,  access_var=access_var,  barrier_var=barrier_var, name_col="name", label_col=label)

    ##---------------- 7) fix select multiple
    barrier_sm_yes = barrier_is_select_multiple(
        survey_data=survey_data,
        barrier_var=barrier_var,
        name_col="name",
        label_col=label
    )
    print("barrier_sm_yes =", barrier_sm_yes)
    fmt_info = {"format": "unknown"}

    if barrier_sm_yes:
        fmt_info = detect_sm_values_are_names_or_labels(
            edu_data=edu_data,
            barrier_var=barrier_var,
            survey_data=survey_data,
            name_col="name",
            label_col=label,   # your label column like 'label::English (en)'
            sample_n=200,
            min_hits=5
        )
        print("Select_multiple value format:", fmt_info)




    # --- map selected severity LABELS -> choice NAMES
    severity_4_matches = find_matching_choices(choice_data, selected_severity_4_barriers, label_var=label)
    severity_5_matches = find_matching_choices(choice_data, selected_severity_5_barriers, label_var=label)

    names_severity_4 = [d["name"] for d in severity_4_matches if d["name"] != "notfound"]
    names_severity_5 = [d["name"] for d in severity_5_matches if d["name"] != "notfound"]

    # If edu values are label-based, match using label candidates instead of name candidates
    if barrier_sm_yes and fmt_info["format"] == "labels":
        names_severity_4 = selected_severity_4_barriers
        names_severity_5 = selected_severity_5_barriers

    # --- build final column
    if barrier_sm_yes:
        edu_data = build_barrier_so_select_multiple(
            edu_data=edu_data,
            barrier_var=barrier_var,
            names_severity_5=names_severity_5,
            names_severity_4=names_severity_4,
            default_value="barrier_3",
            out_col="edu_barrier_final"
        )
    else:
        if barrier_var not in edu_data.columns:
            raise KeyError(f"'{barrier_var}' not found in edu_data columns.")
        edu_data["edu_barrier_final"] = edu_data[barrier_var]



    return edu_data, household_data, survey_data, choice_data,messages