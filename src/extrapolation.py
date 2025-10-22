import os
import glob
import pandas as pd
import copy

admin_var = 'Admin Pcode'
ref_user_admin = 'Reference area Admin Pcode (for extrapolation)'
ref_suggested_admin = 'Suggested reference area for extrapolation'
ref_suggested_other = 'Alternative reference areas'
severity_2025_cols = [
        '% severity levels 1-2', '# severity levels 1-2',
        '% severity level 3', '# severity level 3',
        '% severity level 4', '# severity level 4',
        '% severity level 5', '# severity level 5'
    ]
stable_cols = [
        "Education needs stable (tick 'X' if no change from last year)",
        "Les besoins en matière d'éducation sont stables (cochez « X » si aucun changement par rapport à l'année dernière)"
    ]
delta_cols = [
        "Education needs decreased (tick 'X' if improved from last year)",
        "Les besoins en matière d'éducation ont diminué (cochez « X » si la situation s'est améliorée par rapport à l'année dernière)",
        "Education needs worsened (tick 'X' if worsened from last year)",
        "Les besoins en matière d'éducation se sont aggravés (cochez « X » si la situation s'est aggravée par rapport à l'année dernière)"
    ]



## --------------------------------------------------------------------------------------------------------------
def find_valid_ref(row, valid_list):
    for candidate in [
        row[ref_user_admin],
        row[ref_suggested_admin]
    ]:
        if candidate in valid_list:
            return candidate
    # Check alternative list
    alternatives = str(row[ref_suggested_other]).split(';')
    for alt in alternatives:
        if alt in valid_list:
            return alt
    return ''  # None found
## --------------------------------------------------------------------------------------------------------------



## the pin 2024 will be looked for in the DATA_DIR_PIN2024 according to the country code
## the pin2025 file by population group has to be upload and it is stored: 
## the inputfile filled by the user is the output of the make_output1_platform.py filled by the user on those 4 specific columns

############################################################################################################

def extrapolate_df_2025_updated(base_2025, pin2025_by_status, pin2024_cat):
    
    print(base_2025.columns)

    # 1. ---- create lists: list_remove, list_2025, list_delta, list_same,
    #        - list_remove = list of admins where the values of the columns containing 'severity' is: 'Missing MSNA data for 2024 AND 2025'
    #        - list_2025 = list of admins where the values of the columns in place 4->11 are not all empty
    #        - list_same = list of admins where the colum called: either  'Education needs stable (tick 'X' if no change from last year)'
    #            or 'Les besoins en matière d'éducation sont stables (cochez « X » si aucun changement par rapport à l'année dernière)' contain 'X' or 'x'
    #        - list_delta = list of admins where the colum called: either  'Education needs decreased (tick 'X' if improved from last year)'
    #            or 'Les besoins en matière d'éducation ont diminué (cochez « X » si la situation s'est améliorée par rapport à l'année dernière)'or 
    #            Education needs worsened (tick 'X' if worsened from last year) or 'Les besoins en matière d'éducation se sont aggravés (cochez « X » si la situation s'est aggravée par rapport à l'année dernière)' contain 'X' or 'x'

        # Initialize empty lists
    # Output lists
    list_remove = []
    list_same = []
    list_delta = []

    pin_2025_delta = {}
    pin_2025_same = {}
    pin_2025_covered = {} ## the ones already covered by 2025 MSNA
    pin_2025_updated = {}

    list_2025_category = {}

    for category, df_2025 in pin2025_by_status.items():
        if df_2025.empty:
            continue
        first_col = df_2025.columns[0]
        df_2025 = df_2025.rename(columns={first_col: "Admin Pcode"})
        list_2025_cat = df_2025["Admin Pcode"].dropna().unique().tolist()
        list_2025_category[category] = {"list_2025": list_2025_cat}


        keep_cols = [
            "Admin Pcode",
            "% severity levels 1-2",
            "% severity level 3",
            "% severity level 4",
            "% severity level 5"
        ]

        # Filter columns safely
        existing_keep_cols = [col for col in keep_cols if col in df_2025.columns]
        df_2025_filtered = df_2025[existing_keep_cols].copy()

        # Save back
        pin_2025_covered[category] = df_2025_filtered

    print('===========================')    
    print(list_2025_category)    
    print('===========================')    

    # Columns
    severity_cols = [col for col in base_2025.columns if 'severity' in col]
    
    for _, row in base_2025.iterrows():
        admin_pcode = row[admin_var]

        # (1) Remove if ALL severity cols contain missing string
        if all(str(row[col]).strip().lower() == 'missing msna data for 2024 and 2025' for col in severity_cols):
            list_remove.append(admin_pcode)
            continue  # Skip further checks

    
        # (3) Check 'same' if all severity_2025_cols are empty
        if all(pd.isna(row[col]) or str(row[col]).strip() == '' for col in severity_2025_cols):
            for col in stable_cols:
                if col in row and str(row[col]).strip().lower() == 'x':
                    list_same.append(admin_pcode)
                    break

            # (4) If not stable, check for delta
            else:
                for col in delta_cols:
                    if col in row and str(row[col]).strip().lower() == 'x':
                        list_delta.append(admin_pcode)
                        break

    base_2025 = base_2025[~base_2025['Admin Pcode'].isin(list_remove)].copy()
    # Fallback: if NO delta admins at all, treat ALL remaining admins as 'same'
    if len(list_delta) == 0:
        remaining_admins = base_2025[admin_var].dropna().unique().tolist()
        # Make sure not to duplicate entries already tagged same
        already_same = set(list_same)
        list_same = list(already_same.union(set(remaining_admins)))


    pin_2024_missing_2025 = {}
    for category, df_2024 in pin2024_cat.items():
        first_col_2024 = df_2024.columns[0]
        df_2024 = df_2024.rename(columns={first_col_2024: "Admin Pcode 2024"})
        df_2024_filtered = df_2024[df_2024["Admin Pcode 2024"].isin(list_delta)].copy()
        df_2024_filtered.columns = df_2024_filtered.columns.str.replace(" ", "_")
        admin_col = "Admin_Pcode_2024"  # After replacing spaces
        df_2024_filtered = df_2024_filtered.rename(columns={
            col: f"{col}_2024_extrapolation_base"
            for col in df_2024_filtered.columns if col != admin_col
        })
        pin_2024_missing_2025[category] = df_2024_filtered

    print('------------------')
    print(list_remove)    
    print('------------------')
    print(list_same)    

    ## 2. ---- making the map between missing amdin and reference admin
    df_delta_category = {}

    for category, details in list_2025_category.items():
        valid_refs = details["list_2025"]

        # 🔧 If no delta admins, create an EMPTY df_delta with expected columns
        if len(list_delta) == 0:
            df_delta = pd.DataFrame(columns=[admin_var, 'admin_ref'])
            df_delta_category[category] = df_delta
            continue

        df_delta = base_2025[base_2025[admin_var].isin(list_delta)][[
            admin_var,
            ref_user_admin,
            ref_suggested_admin,
            ref_suggested_other
        ]].copy()

        df_delta['admin_ref'] = df_delta.apply(lambda row: find_valid_ref(row, valid_refs), axis=1)

        # Track non-matched refs to fall back to list_same
        no_match_admins = df_delta[df_delta['admin_ref'] == ''][admin_var].tolist()
        list_same.extend(no_match_admins)

        # Drop reference suggestion columns (cleanup)
        df_delta.drop(columns=[ref_user_admin, ref_suggested_admin, ref_suggested_other], inplace=True)

        # Store in dictionary by category
        df_delta_category[category] = df_delta
 

    merged_df_delta_all = {}  # Dictionary to store results per category

    ## 3. ---- delta preparation and merging
    for category, df_2025 in pin2025_by_status.items():
        if df_2025.empty:
            continue

        # Clean and prepare 2025 data
        first_col_2025 = df_2025.columns[0]
        percent_cols = [
            col for col in df_2025.columns
            if "%" in col and col != "% Tot PiN (severity levels 3-5)"
        ]
        keep_cols = [first_col_2025] + percent_cols
        df_2025 = df_2025.loc[:, keep_cols]
        df_2025 = df_2025.rename(columns={first_col_2025: "Admin Pcode REF"})

        # Clean and prepare 2024 data
        df_2024 = pin2024_cat[category]
        first_col_2024 = df_2024.columns[0]
        df_2024 = df_2024.rename(columns={first_col_2024: "Admin Pcode REF"})

        df_2024 = df_2024.rename(columns={
            col: col.replace(" ", "_") + "_24"
            for col in df_2024.columns if col != "Admin Pcode REF"
        })
        df_2025 = df_2025.rename(columns={
            col: col.replace(" ", "_") + "_25"
            for col in df_2025.columns if col != "Admin Pcode REF"
        })

        # Merge 2025 and 2024 reference data
        matrix_df = df_2025.merge(df_2024, on="Admin Pcode REF", how="left")
        
        df_delta = df_delta_category.get(category)
        # Merge with df_delta for this category
        if df_delta is None or df_delta.empty:
            merged_df_delta_all[category] = pd.DataFrame()
            continue

       
        merged_df_delta = df_delta.merge(
            matrix_df,
            left_on='admin_ref',
            right_on='Admin Pcode REF',
            how='left'
        )

        # Add delta % columns
        for level in ['1-2', '3', '4', '5']:
            if level == '1-2':
                col_24 = "%_severity_levels_1-2_24"
                col_25 = "%_severity_levels_1-2_25"
                delta_col = "delta_%_severity_levels_1-2"
            else:
                col_24 = f"%_severity_level_{level}_24"
                col_25 = f"%_severity_level_{level}_25"
                delta_col = f"delta_%_severity_level_{level}"

            if col_24 in merged_df_delta.columns and col_25 in merged_df_delta.columns:
                def compute_delta(row, col_24=col_24, col_25=col_25):
                    val_24 = row[col_24]
                    val_25 = row[col_25]

                    if pd.isna(val_24) or pd.isna(val_25):
                        return None
                    if val_24 == 0 and val_25 == 0:
                        return 1
                    if val_24 == 0:
                        return 0
                    return (val_24 - val_25) / val_24


                merged_df_delta[delta_col] = merged_df_delta.apply(compute_delta, axis=1)
        print(merged_df_delta.columns)
        merged_df_delta_all[category] = merged_df_delta
           
    ## 4. ---- merging the delta and the pin2024 of refernce  
    for category, df in merged_df_delta_all.items():
        if df is None or df.empty or ('Admin Pcode' not in df.columns):
            continue
        df_sev24 = pin_2024_missing_2025.get(category)
        if df_sev24 is not None and not df_sev24.empty:
            df_merged = df.merge(
                df_sev24,
                left_on='Admin Pcode',
                right_on='Admin_Pcode_2024',
                how='right'
            )
            # Drop duplicated Admin_Pcode_2024 if desired
            df_merged = df_merged.drop(columns=['Admin_Pcode_2024'])
            merged_df_delta_all[category] = df_merged


    ## 5. ---- calculate new severity after extrapolation    
    for category, df in merged_df_delta_all.items():
        if df is None or df.empty:
            pin_2025_delta[category] = pd.DataFrame()
            continue
        new_cols = []

        for level in ['1-2', '3', '4', '5']:
            base_col = f"%_severity_level_{level}_2024_extrapolation_base" if level != '1-2' else "%_severity_levels_1-2_2024_extrapolation_base"
            delta_col = f"delta_%_severity_level_{level}" if level != '1-2' else "delta_%_severity_levels_1-2"
            new_col = f"new_%_severity_level_{level}_extrapolated" if level != '1-2' else "new_%_severity_levels_1-2_extrapolated"

            if base_col in df.columns and delta_col in df.columns:
                df[new_col] = df[base_col] * (1 - df[delta_col])
                new_cols.append(new_col)

        # Normalize so that the sum of new extrapolated columns = 100
        if new_cols:
            sum_col = df[new_cols].sum(axis=1)
            for col in new_cols:
                norm_col = col.replace("new_", "new_normalized_")
                df[norm_col] = df[col] / sum_col * 100

        # Select only first column + last 4 normalized severity columns
        first_col = df.columns[0]
        final_cols = [
            first_col,
            "new_normalized_%_severity_levels_1-2_extrapolated",
            "new_normalized_%_severity_level_3_extrapolated",
            "new_normalized_%_severity_level_4_extrapolated",
            "new_normalized_%_severity_level_5_extrapolated"
        ]
        df_only_2025 = df[final_cols].copy()

        # Rename columns to the final display names
        rename_map = {
            "new_normalized_%_severity_levels_1-2_extrapolated": "% severity levels 1-2",
            "new_normalized_%_severity_level_3_extrapolated": "% severity level 3",
            "new_normalized_%_severity_level_4_extrapolated": "% severity level 4",
            "new_normalized_%_severity_level_5_extrapolated": "% severity level 5"
        }
        df_only_2025 = df_only_2025.rename(columns=rename_map)

        # Save outputs
        pin_2025_delta[category] = df_only_2025
        # Save back
        merged_df_delta_all[category] = df

    ## 6. ---- get values from 2024 pin of the admin in the category 'same'   
    for category, df_2024 in pin2024_cat.items():
        first_col_2024 = df_2024.columns[0]
        df_2024 = df_2024.rename(columns={first_col_2024: "Admin Pcode"})
        df_2024_filtered = df_2024[df_2024["Admin Pcode"].isin(list_same)].copy()
        admin_col = "Admin Pcode"  # After replacing spaces
        pin_2025_same[category] = df_2024_filtered

    print('------------------')
    #print(pin_2025_covered['host'])    
    print('------------------')


    ## 7. ---- Combine all DataFrames into one
    for category, df_1 in pin_2025_covered.items():
        df_delta = pin_2025_delta.get(category, pd.DataFrame())
        df_same = pin_2025_same.get(category, pd.DataFrame())

        df_combined = pd.concat([df_delta, df_same, df_1], axis=0, ignore_index=True)
        pin_2025_updated [category] = df_combined






    return merged_df_delta_all ,pin_2025_updated


