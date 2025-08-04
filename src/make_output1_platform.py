import os
import glob
import pandas as pd
import copy

## the file context_db will be looked for in the DATA_DIR_CONTEXT_DB according to the country code
## the pin2025 file is coming direclty from the previous step of the calculation in the platform: calculation_for_PiN_Dimension.py --> Tot_PiN_by_admin
## the OCHA base comes from the uploaded file from the user



###################################################################################################################################################

def merge_2025_contextDB (country , ocha_data , pin_2025 , context_db_folder ):


    country_code = country.split("--")[-1].strip()

    ## prepare the OCHA base: VERY IMPORTANT
    ocha_2025_base = ocha_data.iloc[:, :2]
    # Check if the IDP column exists and is not entirely empty
    idp_column_name = "IDP/PDI -- Children/Enfants (5-17)"
    if idp_column_name in ocha_data.columns:
        if not ocha_data[idp_column_name].isna().all():
            ocha_2025_base[idp_column_name] = ocha_data[idp_column_name]


    ## read context database for the country
    context_fname = f"{country_code}_context_db.xlsx"
    context_path  = os.path.join(context_db_folder, context_fname)
    if not os.path.exists(context_path):
        raise FileNotFoundError(f"No context file found at {context_path}")
    context_db = pd.read_excel(context_path)
    context_db = context_db.rename(columns={"admin": "Admin Pcode"})

    # 3) Inject “label of the Suggested reference area” right after that column
    ref_col   = "Suggested reference area for extrapolation"
    label_col = "Suggested reference area for extrapolation, LABEL"

    # Build a P‑code → Admin‐name mapping from the FULL ocha_data
    # (so we grab every unique PCODE + its Admin label)
    ocha_map = (
        ocha_data[["Admin Pcode", "Admin"]]
        .drop_duplicates()
        .set_index("Admin Pcode")["Admin"]
        .to_dict()
    )

    if ref_col in context_db.columns:
        insert_at = context_db.columns.get_loc(ref_col) + 1
        context_db.insert(
            insert_at,
            label_col,
            context_db[ref_col].map(ocha_map).fillna("No match in OCHA")
        )
    ## prepare pin 2025
    df_pin2025 = copy.deepcopy(pin_2025)
    first_col = df_pin2025.columns[0]
    df_pin2025 = df_pin2025.rename(columns={first_col: "Admin Pcode"})
    

    complete_output = ocha_2025_base.merge(df_pin2025, on="Admin Pcode", how="left")
    complete_output = complete_output.merge(context_db, on="Admin Pcode", how="left")


    def map_alternative_labels(codes_str):
        if pd.isna(codes_str) or codes_str == "":
            return ""
        codes = codes_str.split(";")
        labels = [ ocha_map.get(code, f"No match for {code}") for code in codes ]
        return ";".join(labels)

    # build the new column as a Series
    label_col2 = "Alternative reference areas, LABELS"
    labels_series = (
        complete_output["Alternative reference areas"]
        .apply(map_alternative_labels)
    )

    # pop it (so it doesn’t end up at the end)
    complete_output[label_col2] = labels_series
    # determine the insert position just after the original column
    orig = "Alternative reference areas"
    loc = complete_output.columns.get_loc(orig) + 1

    # pop and re‐insert at the correct spot
    col = complete_output.pop(label_col2)
    complete_output.insert(loc, label_col2, col)


    print("Final merged dataframe columns:")
    print(complete_output.columns)
    print("Number of rows:", len(complete_output))

    # Replace NaN with '-' only if ALL 'severity' columns in a row are NaN  --> this means that in that admin there was not any 2024 or 2025 msna
    severity_cols = [col for col in complete_output.columns if 'severity' in col]
    if severity_cols:
        mask_all_severity_nan = complete_output[severity_cols].isna().all(axis=1)
        complete_output.loc[mask_all_severity_nan, severity_cols] = "Missing MSNA data for 2024 AND 2025"


    idp_col = "totalIDP -- HPC 2026"
    totn_col = "TotN"

    if idp_column_name in complete_output.columns:
        complete_output = complete_output.rename(columns={
            idp_column_name: idp_col
        })
        # Move column to end
        cols = [col for col in complete_output.columns if col != "totalIDP -- HPC 2026"]
        complete_output = complete_output[cols + ["totalIDP -- HPC 2026"]]

    if idp_col in complete_output.columns and totn_col in complete_output.columns:
        # Ensure numeric types
        complete_output[idp_col] = pd.to_numeric(complete_output[idp_col], errors="coerce")
        complete_output[totn_col] = pd.to_numeric(complete_output[totn_col], errors="coerce")

        def compute_ratio(row):
            if pd.isna(row[idp_col]) or pd.isna(row[totn_col]):
                return pd.NA
            if row[idp_col] == 0:
                return 0
            if row[totn_col] == 0:
                return pd.NA
            return row[idp_col] / row[totn_col]

        complete_output["ratio IDP/ToTN - HPC2026"] = complete_output.apply(compute_ratio, axis=1)

        # Drop the intermediate IDP column
        complete_output.drop(columns=[idp_col], inplace=True)





    return complete_output


