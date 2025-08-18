import pandas as pd
#import fuzzywuzzy
from fuzzywuzzy import process, fuzz
import numpy as np
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
import re
from collections import defaultdict

int_2 = '2.0'
int_3 = '3.0'
int_4 = '4.0'
int_5 = '5.0'
label_perc2 = '% severity levels 1-2'
label_perc3 = '% severity level 3'
label_perc4 = '% severity level 4'
label_perc5 = '% severity level 5'
label_tot2 = '# severity levels 1-2'
label_tot3 = '# severity level 3'
label_tot4 = '# severity level 4'
label_tot5 = '# severity level 5'
label_perc_tot = '% Tot PiN (severity levels 3-5)'
label_tot = '# Tot PiN (severity levels 3-5)'
label_admin_severity = 'Area severity'
label_tot_population = 'TotN'
label_tot_enrolled = 'E'
label_tot_OoS = 'OoS'
label_tot_inschool = 'inschool'

int_acc = 'access'
int_agg= 'aggravating circumstances'
int_lc = 'learning condition'
int_penv = 'protected environment'
int_out = 'Not in need'
label_perc_acc = '% Access'
label_perc_agg= '% Aggravating circumstances'
label_perc_lc = '% Learning conditions'
label_perc_penv = '% Protected environment'
label_perc_out = '% Not in need'
label_tot_acc = '# Access'
label_tot_agg= '# Aggravating circumstances'
label_tot_lc = '# Learning conditions'
label_tot_penv = '# Protected environment'
label_tot_out = '# Not in need'
label_dimension_perc_tot = '% Tot in PiN Dimensions'
label_dimension_tot = '# Tot in PiN Dimensions'
label_dimension_tot_population = 'TotN'

label_tot_sev3_indicator_access= 'severity level 3: (ToT # children) indicator Access'
label_tot_sev3_indicator_teacher = 'severity level 3: (ToT # children) indicator Teacher Absence Disruption'
label_tot_sev3_indicator_hazard = 'severity level 3: (ToT # children) indicator Natural Hazard Disruption'
label_tot_sev4_indicator_idp = 'severity level 4: (ToT # children) indicator School Used As Shelter Disruption'
label_tot_sev5_indicator_occupation = 'severity level 5: (ToT # children) indicator School Occupation Disruption'
label_tot_sev4_aggravating_circumstances = 'severity level 4: (ToT # children) indicator aggravating circumstances (cumulative of all Level 4 aggravating circumstances)'
label_tot_sev5_aggravating_circumstances = 'severity level 5: (ToT # children) indicator aggravating circumstances (cumulative of all Level 5 aggravating circumstances)'

label_perc_sev3_indicator_access= 'severity level 3: (% of children) indicator Access'
label_perc_sev3_indicator_teacher = 'severity level 3: (% of children) indicator Teacher Absence Disruption'
label_perc_sev3_indicator_hazard = 'severity level 3: (% of children) indicator Natural Hazard Disruption'
label_perc_sev4_indicator_idp = 'severity level 4: (% of children) indicator School Used As Shelter Disruption'
label_perc_sev5_indicator_occupation = 'severity level 5: (% of children) indicator School Occupation Disruption'
label_perc_sev4_aggravating_circumstances = 'severity level 4: (% of children) indicator aggravating circumstances (cumulative of all Level 4 aggravating circumstances)'
label_perc_sev5_aggravating_circumstances = 'severity level 5: (% of children) indicator aggravating circumstances (cumulative of all Level 5 aggravating circumstances)'

##--------------------------------------------------------------------------------------------
##--------------------------------------------------------------------------------------------
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
##--------------------------------------------------------------------------------------------
# Step 1: Categorize codes by length
def categorize_levels_dynamic(prefix_list):
    # Dictionary to hold codes grouped by their length
    length_dict = defaultdict(list)

    # Loop through each code and categorize by length
    for code in prefix_list:
        code_length = len(code)
        length_dict[code_length].append(code)
    
    return length_dict
##--------------------------------------------------------------------------------------------
# Step 2: Modify the logic to find the appropriate columns in `edu_data`
# Helper function to find matching columns for each length level
def find_matching_columns_for_admin_levels(edu_data, household_data, prefix_list, admin_var):
    # Categorize codes based on length
    length_dict = categorize_levels_dynamic(prefix_list)
    admin_columns_representative = {}

    # Get the available columns from the `edu_data` and `household_data` dataframes
    edu_columns = edu_data.columns

    # Use the new best-match finder that checks P-codes etc.
    best_match_for_admin_var = find_best_match(admin_var, household_data)
    print(f"Best match for admin_var ({admin_var}) is: {best_match_for_admin_var}")

    # Iterate through each column in the edu_data dataframe
    for col in edu_columns:
        # Convert the column to strings to ensure type consistency
        column_data = edu_data[col].astype(str)

        # For each length group in the `length_dict`, check for matches
        for length, codes in length_dict.items():
            matching_values = column_data.isin(codes)

            # If there are any matches, add the column to the admin_columns_representative dictionary for that length
            if matching_values.any():
                if length not in admin_columns_representative:
                    admin_columns_representative[length] = []
                admin_columns_representative[length].append(col)
                print(f"Matching column found: {col} for length {length}")

    # Helper to pick the column with the most non-empty values
    def prioritize_non_empty_columns(columns):
        non_empty_counts = {col: edu_data[col].notna().sum() for col in columns}
        sorted_columns = sorted(non_empty_counts, key=non_empty_counts.get, reverse=True)
        return sorted_columns[0] if sorted_columns else None

    # Handle multiple or single admin levels
    if len(length_dict) > 1:
        print("Multiple levels detected:")
        for length, codes in length_dict.items():
            print(f"Level {length}: {codes}")

        best_columns = {}
        for length, columns in admin_columns_representative.items():
            best_columns[length] = prioritize_non_empty_columns(columns)
        admin_columns_representative = best_columns
    else:
        if length_dict:
            single_level = next(iter(length_dict.keys()))
            columns_for_single_level = admin_columns_representative.get(single_level, [])
            if columns_for_single_level:
                admin_columns_representative[single_level] = prioritize_non_empty_columns(columns_for_single_level)
            else:
                admin_columns_representative = {}

    return admin_columns_representative

##--------------------------------------------------------------------------------------------
# Function to translate labels
def translate_labels(data, translation_dict):
    """
    Translates column names and text values inside a DataFrame or dictionary of DataFrames.
    Supports partial matches in column names.

    Args:
        data (pd.DataFrame or dict): DataFrame or dictionary of DataFrames to be translated.
        translation_dict (dict): Dictionary where keys are words/phrases to translate and values are translations.

    Returns:
        Translated DataFrame(s).
    """
    
    def translate_column_names(column_name):
        """Apply translation to parts of column names based on dictionary."""
        for eng, fr in translation_dict.items():
            if eng in column_name:  # Check if part of the column matches
                column_name = column_name.replace(eng, fr)  # Replace only that part
        return column_name
    
    # Check if 'data' is a DataFrame or a dictionary of DataFrames
    if isinstance(data, pd.DataFrame):
        # Translate column names with partial replacements
        data.columns = [translate_column_names(col) for col in data.columns]

        # Translate values inside the DataFrame (only for string columns)
        for col in data.columns:
            if data[col].dtype == 'object':  # Apply replacement only for string columns
                for eng, fr in translation_dict.items():
                    data[col] = data[col].str.replace(eng, fr, regex=True)

        return data

    elif isinstance(data, dict):
        # If 'data' is a dictionary, apply translation to each DataFrame
        for key in data:
            data[key] = translate_labels(data[key], translation_dict)
        return data

    else:
        raise TypeError("Input must be a pandas DataFrame or a dictionary of DataFrames.")

##--------------------------------------------------------------------------------------------
def calculate_prop(df, admin_var, pop_group_var, target_var, agg_var='weights'):

    df_results = df.groupby([admin_var, pop_group_var, target_var]).agg(
            total_weight=(agg_var, 'sum')
        ).groupby(level=[0, 1]).apply(
            lambda x: x / x.sum()
        ).unstack(fill_value=0)

    return df_results 
##--------------------------------------------------------------------------------------------
def reduce_index(df, level, pop_group_var):
    df.columns = df.columns.get_level_values(1)
    df=df.droplevel(0, axis=0) 
    df=df.droplevel(0, axis=0) 
    if level == 0: df = df.reset_index( level = [0 , 1] ) 
    if level == 1: df = df.reset_index( level = [0 , 1, 2] ) 

    # Splitting the DataFrame based on pop_group_var
    groups = df.groupby(pop_group_var)
    df_list = {name: group for name, group in groups}

    return df_list
##--------------------------------------------------------------------------------------------
def run_mismatch_admin_analysis(df, admin_var, admin_column_rapresentative, pop_group_var, analysis_variable, 
                                admin_low_ok_list, prefix_list, grouped_dict):
    all_expanded_results_admin_up = {}  # Collect results from both levels by category
    admin_var_dummy = 'admin_var_dummy'


    # Check if the `admin_var` column is empty
    if df[admin_var].notna().any():
        # 1. Run the analysis grouped by 'admin_var' (Analysis A)
        results_analysis_admin_low = calculate_prop (df=df, admin_var=admin_var, pop_group_var=pop_group_var, target_var= analysis_variable)
        results_analysis_admin_low = reduce_index(results_analysis_admin_low, 0, pop_group_var)

        # 3. Filter results_analysis_admin_low to only include rows where 'admin_var' is in 'admin_low_ok_list'
        if admin_low_ok_list:
            for category, pop_group_df in results_analysis_admin_low.items():
                # Apply filtering to the 'admin_var' column
                pop_group_df = pop_group_df[pop_group_df[admin_var].isin(admin_low_ok_list)]
                results_analysis_admin_low[category] = pop_group_df
        else:
            print("admin_low_ok_list is empty, skipping filtering for Analysis A.")
            results_analysis_admin_low = {}  # Or set to None if you prefer
    else:
        print(f"admin_var column ({admin_var}) is empty, skipping Analysis A.")
        results_analysis_admin_low = {}  # Or set to None if you prefer



    # Case where 'admin_column_rapresentative' is a dictionary, even with one level
    if isinstance(admin_column_rapresentative, dict):
        # Check if it's a single-level case (only one key in the dictionary)
        if len(admin_column_rapresentative) == 1:
            # Extract the single value from the dictionary
            length, admin_col = list(admin_column_rapresentative.items())[0]
            
            # Run the analysis grouped by this single admin column
            results_analysis_admin_up = calculate_prop (df=df, admin_var=admin_col, pop_group_var=pop_group_var, target_var= analysis_variable)
            results_analysis_admin_up = reduce_index(results_analysis_admin_up, 0, pop_group_var)

            admin_var_dummy = 'admin_var_dummy'
            for category, pop_group_df in results_analysis_admin_up.items():
                    pop_group_df.rename(columns={admin_col: admin_var_dummy}, inplace=True)
                    results_analysis_admin_up[category] = pop_group_df

            # 4. Filter results_analysis_admin_up to only include rows where 'admin_col_level1' is in 'prefix_list'
            if prefix_list:
                for category, pop_group_df in results_analysis_admin_up.items():
                    pop_group_df = pop_group_df[pop_group_df[admin_var_dummy].isin(prefix_list)]
                    results_analysis_admin_up[category] = pop_group_df


            # Expand results based on 'grouped_dict'
            for category, pop_group_df in results_analysis_admin_up.items():
                for admin_column_value in grouped_dict.keys():
                    if admin_column_value in pop_group_df[admin_var_dummy].values:
                        matching_rows = pop_group_df[pop_group_df[admin_var_dummy] == admin_column_value]

                        # Duplicate rows for each detailed admin
                        for detailed_admin in grouped_dict[admin_column_value]:
                            expanded_row = matching_rows.copy()
                            expanded_row[admin_var_dummy] = detailed_admin
                            all_expanded_results_admin_up.setdefault(category, []).append(expanded_row)

        else:
            # Case where there are multiple levels (e.g., {4: 'i_admin1', 6: 'i_admin2'})
            results_analysis_admin_up = {}
            all_expanded_results_admin_up_single = {}

            for idx, (length, admin_col) in enumerate(admin_column_rapresentative.items()):
                #print(f"Running analysis for length {length} with column {admin_col} (index {idx})")

                # Perform analysis grouped by each admin column
                results_analysis_admin_up[idx] = calculate_prop (df=df, admin_var=admin_col, pop_group_var=pop_group_var, target_var= analysis_variable)
                results_analysis_admin_up[idx] = reduce_index(results_analysis_admin_up[idx], 0, pop_group_var)
                
                admin_var_dummy = 'admin_var_dummy'
                for category, pop_group_df in results_analysis_admin_up[idx].items():
                        pop_group_df.rename(columns={admin_col: admin_var_dummy}, inplace=True)
                        results_analysis_admin_up[idx][category] = pop_group_df


                #print('results_analysis_admin_up')
                #print(results_analysis_admin_up[idx])

                # 4. Filter the results based on the prefix list
                if prefix_list:
                    for category, pop_group_df in results_analysis_admin_up[idx].items():
                        pop_group_df = pop_group_df[pop_group_df[admin_var_dummy].isin(prefix_list)]
                        results_analysis_admin_up[idx][category] = pop_group_df

                #print('results_analysis_admin_up FILTERED')
                #print(results_analysis_admin_up[idx])

                # Initialize the expanded results for this index
                all_expanded_results_admin_up_single[idx] = {}

                # Expand results based on 'grouped_dict'
                for category, pop_group_df in results_analysis_admin_up[idx].items():
                    for admin_column_value in grouped_dict.keys():
                        if admin_column_value in pop_group_df[admin_var_dummy].values:
                            matching_rows = pop_group_df[pop_group_df[admin_var_dummy] == admin_column_value]

                            # Duplicate rows for each detailed admin
                            for detailed_admin in grouped_dict[admin_column_value]:
                                expanded_row = matching_rows.copy()
                                expanded_row[admin_var_dummy] = detailed_admin
                                all_expanded_results_admin_up_single[idx].setdefault(category, []).append(expanded_row)
                #print(f'results_analysis_admin_up EXPANDED for index {idx}')
                #print(all_expanded_results_admin_up_single[idx])

            # Concatenate the two levels of results into a single DataFrame
            for category in all_expanded_results_admin_up_single[0].keys():
                if category in all_expanded_results_admin_up_single[1]:
                    #print(category)
                    pop_group_df_0 = pd.concat(all_expanded_results_admin_up_single[0][category], ignore_index=True)
                    pop_group_df_1 = pd.concat(all_expanded_results_admin_up_single[1][category], ignore_index=True)

                    final_concat = pd.concat([pop_group_df_0, pop_group_df_1], ignore_index=True)
                    all_expanded_results_admin_up[category] = final_concat
                else:
                    # Handle cases where category only exists in one of the levels
                    all_expanded_results_admin_up[category] = pd.concat(all_expanded_results_admin_up_single[0][category], ignore_index=True)

    else:
        raise ValueError("admin_column_rapresentative should always be a dictionary in this case.")

    if all_expanded_results_admin_up:
        # Convert lists of DataFrames to DataFrames by concatenating them first
        for category in all_expanded_results_admin_up.keys():
            if isinstance(all_expanded_results_admin_up[category], list):
                all_expanded_results_admin_up[category] = pd.concat(all_expanded_results_admin_up[category], ignore_index=True)

        # Now concatenate all the DataFrames into a single DataFrame
        expanded_results_admin_up_df = pd.concat(all_expanded_results_admin_up.values(), ignore_index=True)
    else:
        expanded_results_admin_up_df = pd.DataFrame()

    results_analysis_admin_up_duplicated = expanded_results_admin_up_df

    # 5. Merge with results from Analysis A (if Analysis A was run)
    results_analysis_complete = {}

    if results_analysis_admin_low:
        print('If Analysis A results exist, merge Analysis A (admin_low) with Analysis B (admin_up)')
        # If Analysis A results exist, merge Analysis A (admin_low) with Analysis B (admin_up)
        for category, admin_low in results_analysis_admin_low.items():
            if category in results_analysis_admin_up_duplicated[pop_group_var].unique():
                admin_up = results_analysis_admin_up_duplicated[results_analysis_admin_up_duplicated[pop_group_var] == category].copy() 
                admin_up.rename(columns={admin_var_dummy: admin_var}, inplace=True)

                # Combine admin_low and admin_up
                all_admin = pd.concat([admin_low, admin_up], ignore_index=True)
                results_analysis_complete[category] = all_admin
            else:
                results_analysis_complete[category] = admin_low
    else:
        print('# Process only Analysis B (admin_up) results')
        # Process only Analysis B (admin_up) results
        for category in results_analysis_admin_up_duplicated[pop_group_var].unique():
            admin_up = results_analysis_admin_up_duplicated[results_analysis_admin_up_duplicated[pop_group_var] == category].copy() 
            admin_up.rename(columns={admin_var_dummy: admin_var}, inplace=True)

            # Use only admin_up for this case
            results_analysis_complete[category] = admin_up

    # Return final results

    return results_analysis_complete
##--------------------------------------------------------------------------------------------
def clean_indicator_columns(pin_list, dataframe_name):
    # Remove column named `0` if it exists

    for category, grouped_df in pin_list.items():

        if 0 in grouped_df.columns:
            grouped_df = grouped_df.drop(columns=[0])
        # Rename column `1` to match the DataFrame name without `_list`
        new_column_name = dataframe_name.replace('_list', '')
        if 1 in grouped_df.columns:
            grouped_df = grouped_df.rename(columns={1: new_column_name})
    
        pin_list[category] = grouped_df
    return pin_list

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
##--------------------------------------------------------------------------------------------
def process_indicator_dataframes(indicator_access_list, indicator_dataframes, choice_data, 
                                 selected_severity_4_barriers, selected_severity_5_barriers, 
                                 label_var, admin_var, pop_group_var):

    pin_by_indicator_status = {}

    # Step 1: Merge Indicator DataFrames
    for category, grouped_ind_df in indicator_access_list.items():
        # Start with the base DataFrame for this category
        pop_group_ind_df = grouped_ind_df.copy()
        
        # Merge with other indicator DataFrames
        for indicator_df in indicator_dataframes:
            if category in indicator_df:  # Check if the category exists in the current indicator DataFrame
                pop_group_ind_df = pd.merge(
                    pop_group_ind_df,
                    indicator_df[category],
                    on=[admin_var, pop_group_var],
                    how='left',  # Preserve rows from the base DataFrame
                    suffixes=('', '_dup')  # Add a suffix to duplicate columns
                )
        for col in pop_group_ind_df.columns:
            if col.endswith('_dup'):
                original_col = col.replace('_dup', '')
                if original_col in pop_group_ind_df.columns:  # If original exists, remove duplicate
                    pop_group_ind_df.drop(columns=[col], inplace=True)
                else:
                    pop_group_ind_df.rename(columns={col: original_col}, inplace=True)

        pin_by_indicator_status[category] = pop_group_ind_df 

    # Step 2: Identify Severity Names and Prepare Column Renaming
    severity_4_matches = find_matching_choices(choice_data, selected_severity_4_barriers, label_var=label_var)
    severity_5_matches = find_matching_choices(choice_data, selected_severity_5_barriers, label_var=label_var)

    names_severity_4 = [entry['name'] for entry in severity_4_matches]
    names_severity_5 = [entry['name'] for entry in severity_5_matches]

    # Step 3: Ensure only relevant columns are kept
    if pin_by_indicator_status:
        sample_category = next(iter(pin_by_indicator_status))  # Get sample category
        sample_df = pin_by_indicator_status[sample_category]  # Get DataFrame

        # Essential columns that must be kept
        essential_columns = [admin_var, pop_group_var]
        optional_columns = ['indicator_access', 
            'indicator_teacher', 'indicator_hazard',
            'indicator_idp', 'indicator_occupation', 'indicator_barrier4', 'indicator_barrier5'
        ]
        essential_columns += [col for col in optional_columns if col in sample_df.columns]

        # Step 4: Prepare Column Renaming
        # Step 4: Prepare Column Renaming
        essential_column_rename = {
            col: new_name for col, new_name in {
                'indicator_access': 'rate_indicator_access',
                'indicator_teacher': 'subsetInSchool_sev3_indicator_teacher',
                'indicator_hazard': 'subsetInSchool_sev3_indicator_hazard',
                'indicator_idp': 'subsetInSchool_sev4_indicator_idp',
                'indicator_occupation': 'subsetInSchool_sev5_indicator_occupation',
                'indicator_barrier4': 'subsetOoS_sev4_aggravating_circumstances',
                'indicator_barrier5': 'subsetOoS_sev5_aggravating_circumstances'
            }.items() if col in essential_columns
        }

        # Severity 4 and 5 column renaming
        severity_4_rename = {entry['name']: f"subsetOoS_sev4: {entry['label']}" for entry in severity_4_matches}
        severity_5_rename = {entry['name']: f"subsetOoS_sev5: {entry['label']}" for entry in severity_5_matches}
        # Merge all renaming mappings
        rename_mapping = {**essential_column_rename, **severity_4_rename, **severity_5_rename}

        # Step 5: Apply filtering & renaming to each category DataFrame
        for category, grouped_ind_df in pin_by_indicator_status.items():
            pop_group_ind_df = grouped_ind_df.copy()

            # Identify severity columns that exist
            severity_columns = set(names_severity_4 + names_severity_5)
            available_severity_columns = [col for col in pop_group_ind_df.columns if col in severity_columns]

            # Keep only necessary columns
            final_columns = [col for col in essential_columns + available_severity_columns if col in pop_group_ind_df.columns]
            pop_group_ind_df = pop_group_ind_df[final_columns]

            # Rename columns based on severity mapping
            pop_group_ind_df.rename(columns=rename_mapping, inplace=True)

            # Store processed DataFrame
            pin_by_indicator_status[category] = pop_group_ind_df

    return pin_by_indicator_status       
##--------------------------------------------------------------------------------------------        
def map_template_to_status(template_values, suggestions_mapping, status_values):
    results = {}
    for template in template_values:
        suggestions = suggestions_mapping.get(template, [])
        # Search for the first matching status with direct comparisons
        match = next((status for status in status_values if status in suggestions), None)
        if match:
            results[template] = match
        else:
            results[template] = 'No match found'
    return results

##--------------------------------------------------------------------------------------------        
def extract_status_data(ocha_data, mapped_statuses, pop_group_var):
    # Data frames dictionary to store each category's DataFrame
    data_frames = {}
    
    for category, status in mapped_statuses.items():
        if status != 'No match found':
            # Use the status as the category name for clarity and direct mapping
            category_name = status  # This changes the category name to the matched status value
            
            # Prepare the column names to extract based on the matched status
            children_col = f"{category} -- Children/Enfants (5-17)"
            
            # Check if these columns exist in the DataFrame
            if all(col in ocha_data.columns for col in [children_col]):
                # Create a new DataFrame for this category using the status as the category name
                category_df = ocha_data[['Admin', children_col]].copy()
                category_df.rename(columns={
                    children_col: 'TotN'
                }, inplace=True)
                category_df['Category'] = category_name  # Set the category name to the matched status
                category_df[pop_group_var] = status
                data_frames[category_name] = category_df
            else:
                print(f"Columns for {category} not found in OCHA data.")
        else:
            print(f"No match found for the category: {category}, skipping data extraction for this category.")

    return data_frames
##--------------------------------------------------------------------------------------------
def cap_and_redistribute(enrollment_df, valid_mappings, max_iterations=10):
    """
    For each pop group (e.g., 'host', 'idp', 'ret'), if the column 
    "{pop_group} -- E" exceeds the cap given in "{pop_group} -- TotN",
    cap it and redistribute the excess proportionally to the TotN of the other groups.
    
    Parameters:
      enrollment_df : pandas DataFrame with columns like:
          "{pop_group} -- E" and "{pop_group} -- TotN"
      valid_mappings: dictionary mapping original labels to pop group names, e.g.,
          {'Host/Hôte': 'host', 'IDP/PDI': 'idp', 'Returnees/Retournés': 'ret'}
      max_iterations: maximum number of iterations to perform
    Returns:
      Modified enrollment_df with reallocated E values.
    """
    
    iteration = 0
    # Loop until no changes occur or until maximum iterations are reached.
    while iteration < max_iterations:
        any_adjustment = False  # To track if any group needed adjustment in this iteration
        
        # For each pop group, check if its E value is above its TotN
        for pop_group in valid_mappings.values():
            e_col = f"{pop_group} -- E"
            tot_col = f"{pop_group} -- TotN"
            
            # Calculate the excess amount for rows where E > TotN
            excess = enrollment_df[e_col] - enrollment_df[tot_col]
            mask = excess > 0
            
            if mask.any():
                any_adjustment = True
                # For rows where the value exceeds the cap, store the excess
                excess_amount = excess[mask]
                
                # Cap the group's E at TotN for those rows
                enrollment_df.loc[mask, e_col] = enrollment_df.loc[mask, tot_col]
                
                # Identify the other groups to which we will redistribute the excess
                other_groups = [pg for pg in valid_mappings.values() if pg != pop_group]
                # For these rows, compute the sum of TotN for the other groups
                tot_sum = enrollment_df.loc[mask, [f"{pg} -- TotN" for pg in other_groups]].sum(axis=1)
                
                # For each other group, add a share of the excess proportional to its TotN
                for other in other_groups:
                    other_e_col = f"{other} -- E"
                    other_tot_col = f"{other} -- TotN"
                    
                    # The allocated excess for this group is:
                    # excess * (TotN_other / sum_{other groups} TotN)
                    allocation = excess_amount * enrollment_df.loc[mask, other_tot_col] / tot_sum
                    
                    # Increase the current E value by the allocated excess
                    enrollment_df.loc[mask, other_e_col] = enrollment_df.loc[mask, other_e_col] + allocation
        
        # If no adjustment was made in this iteration, we are done.
        if not any_adjustment:
            break
        
        iteration += 1
    
    return enrollment_df

##--------------------------------------------------------------------------------------------
def add_figures_columns(df, enrolled_col="E", oos_col="OoS"):
    for col in df.columns:
        # Check for in-school severity columns
        if "subsetInSchool_" in col:
            df["N___" +col] = df[col] * df[enrolled_col]
        # Check for out-of-school severity columns
        elif "subsetOoS_" in col:
            df["N___"+ col ] = df[col] * df[oos_col]
    return df
##--------------------------------------------------------------------------------------------
def add_additional_severity_columns(df, enrolled_col="E", oos_col="OoS"):
    # For in-school: sum all weighted columns with prefix "N___subsetInSchool_"
    in_school_cols = [col for col in df.columns if col.startswith("N___subsetInSchool_")]
    df["N__subsetInSchool_sev2"] = df[enrolled_col] - df[in_school_cols].sum(axis=1)
    
    # For out-of-school: subtract the sum of the two specific weighted columns
    oos_subset_cols = ["N___subsetOoS_sev4_aggravating_circumstances", "N___subsetOoS_sev5_aggravating_circumstances"]
    # It is assumed these columns exist; if not, you might want to check or handle errors.
    df["N___subsetOoS_sev3"] = df[oos_col] - df[oos_subset_cols].sum(axis=1)
    
    return df
##--------------------------------------------------------------------------------------------
def final_severity_columns(df,  admin_var = 'admin', pop_group_var = 'pop_group'):

    df = df.rename(columns={'E': 'In-School Children'})
    df = df.rename(columns={'OoS': 'Out-of-School Children'})

    df = df.rename(columns={'N__subsetInSchool_sev2': label_tot2})

    in_school_cols_sev3 = [col for col in df.columns if col.startswith("N___subsetInSchool_sev3")]
    in_school_cols_sev4 = [col for col in df.columns if col.startswith("N___subsetInSchool_sev4")]
    in_school_cols_sev5 = [col for col in df.columns if col.startswith("N___subsetInSchool_sev5")]

    df[label_tot3] = df['N___subsetOoS_sev3'] + df[in_school_cols_sev3].sum(axis=1)
    df[label_tot4] = df['N___subsetOoS_sev4_aggravating_circumstances'] + df[in_school_cols_sev4].sum(axis=1)
    df[label_tot5] = df['N___subsetOoS_sev5_aggravating_circumstances'] + df[in_school_cols_sev5].sum(axis=1)
    
    allowed_cols = [admin_var, label_tot_population,   "In-School Children", "Out-of-School Children", pop_group_var,
                    label_tot2, label_tot3, label_tot4, label_tot5]
    df = df[allowed_cols]

    return df

##--------------------------------------------------------------------------------------------
def final_dimension_columns(df,  admin_var = 'admin', pop_group_var = 'pop_group'):

    df = df.rename(columns={'E': 'In-School Children'})
    df = df.rename(columns={'OoS': 'Out-of-School Children'})

    in_school_cols_sev3 = [col for col in df.columns if col.startswith("N___subsetInSchool_sev3")]
    in_school_cols_sev4 = [col for col in df.columns if col.startswith("N___subsetInSchool_sev4")]
    in_school_cols_sev5 = [col for col in df.columns if col.startswith("N___subsetInSchool_sev5")]


    df[label_tot_acc] = df ['N___subsetOoS_sev3']
    df[label_tot_lc] = df[in_school_cols_sev3].sum(axis=1)
    df[label_tot_penv] = df[in_school_cols_sev4].sum(axis=1) + df[in_school_cols_sev5].sum(axis=1)
    df[label_tot_agg] = df['N___subsetOoS_sev4_aggravating_circumstances'] + df['N___subsetOoS_sev5_aggravating_circumstances']

    df[label_dimension_tot] = df[label_tot_acc] + df[label_tot_lc] + df[label_tot_penv] + df[label_tot_agg]
    
    allowed_cols = [admin_var, label_tot_population,   "In-School Children", "Out-of-School Children", pop_group_var,
                    label_tot_acc, label_tot_lc, label_tot_penv, label_tot_agg, label_dimension_tot]
    df = df[allowed_cols]
    return df

##--------------------------------------------------------------------------------------------
def calculate_category_factors(df, total_col, category_col, category_name):
    """
    Calculate ratios for specific categories and filter the DataFrame to include only necessary columns.
    Returns a dictionary of DataFrames.
    """
    result_df = df.copy()
    result_df[category_name] = result_df[category_col] / result_df[total_col].replace(0, pd.NA)
    result_df['Category'] = category_name
    columns_to_keep = ['Admin',  category_name, 'Category']
    
    # Wrap the result in a dictionary using category_name as the key
    return {category_name: result_df[columns_to_keep]}

##--------------------------------------------------------------------------------------------
def calculate_cycle_factors(df, factor_cycle, primary_start, secondary_end, vector_cycle, single_cycle):

    if single_cycle:
        factor_cycle[0] = (vector_cycle[0] - primary_start +1) / (secondary_end - primary_start + 2)
        factor_cycle[1] =0
        factor_cycle[2] =  (secondary_end - vector_cycle[0]) / (secondary_end - primary_start + 2)
    else:
        factor_cycle[0] = (vector_cycle[0] - primary_start +1) / (secondary_end - primary_start + 2)
        factor_cycle[1] = (vector_cycle[1] - vector_cycle[0]) / (secondary_end - primary_start + 2)
        factor_cycle[2] = (secondary_end - vector_cycle[1]) / (secondary_end - primary_start + 2)

    # Create dictionaries to hold the categories and their respective factors
    categories = {
        'primary': factor_cycle[0],
        'intermediate level': factor_cycle[1],
        'secondary': factor_cycle[2]
    }
    # Create DataFrames for each category
    result = {}
    for category, factor in categories.items():
        temp_df = df.copy()
        temp_df[category] = factor
        temp_df['Category'] = category
        columns_to_keep = ['Admin',  category, 'Category']
        result[category] = temp_df[columns_to_keep]
    return result

##--------------------------------------------------------------------------------------------
def final_indicator_columns(df,  admin_var = 'admin', pop_group_var = 'pop_group'):

    df = df.rename(columns={'E': 'In-School Children',
                            'OoS': 'Out-of-School Children'})
    
    rename_dict = {
        'N___subsetOoS_sev3': label_tot_sev3_indicator_access,
        'N___subsetInSchool_sev3_indicator_teacher': label_tot_sev3_indicator_teacher,
        'N___subsetInSchool_sev3_indicator_hazard': label_tot_sev3_indicator_hazard,
        'N___subsetInSchool_sev4_indicator_idp': label_tot_sev4_indicator_idp,
        'N___subsetOoS_sev4_aggravating_circumstances': label_tot_sev4_aggravating_circumstances,
        'N___subsetInSchool_sev5_indicator_occupation': label_tot_sev5_indicator_occupation,
        'N___subsetOoS_sev5_aggravating_circumstances': label_tot_sev5_aggravating_circumstances
    }

    # Only keep keys that exist in the dataframe:
    rename_dict = {k: v for k, v in rename_dict.items() if k in df.columns}

    df = df.rename(columns=rename_dict)

    new_cols = {}
    for col in df.columns:
        if col.startswith("N___subsetOoS_sev5: "):
            new_cols[col] = col.replace("N___subsetOoS_sev5: ", "severity level 5: (ToT # children) ")
        elif col.startswith("N___subsetOoS_sev4: "):
            new_cols[col] = col.replace("N___subsetOoS_sev4: ", "severity level 4: (ToT # children) ")
    df = df.rename(columns=new_cols)

    
    basic_allowed  = [admin_var, label_tot_population,   "In-School Children", "Out-of-School Children", pop_group_var,
                    label_tot_sev3_indicator_access, label_tot_sev3_indicator_teacher, label_tot_sev3_indicator_hazard, label_tot_sev4_indicator_idp, label_tot_sev5_indicator_occupation,
                    label_tot_sev4_aggravating_circumstances,label_tot_sev5_aggravating_circumstances ]
    
    basic_allowed = [col for col in basic_allowed if col in df.columns]

    extra_cols = [col for col in df.columns 
                  if col.startswith("severity level 4: (ToT # children)") 
                  or col.startswith("severity level 5: (ToT # children)")]
    
    # Combine the basic and extra columns (preserving the order in basic_allowed first)
    allowed_cols = basic_allowed + extra_cols
    df = df[allowed_cols]
    return df
##--------------------------------------------------------------------------------------------        
# Function to handle numeric and percentage columns
def rounding_dataframe(df, figures_round, percentage_round):
    for col in df.columns:
        if col.startswith('#'):
            # Convert to numeric and round
            df[col] = pd.to_numeric(df[col], errors='coerce').round(figures_round)
        elif col.startswith('%'):
            # Convert to numeric, multiply by 100, and round
            df[col] = pd.to_numeric(df[col], errors='coerce').apply(lambda x: round(x * 100, percentage_round))        
##--------------------------------------------------------------------------------------------
# preparation for overview--> SUM all the admin per population group and per strata 
def collapse_and_summarize(pin_per_admin_status_strata, category_str, admin_var):
    collapsed_results = {}
    
    # Iterate over the input dictionary
    for category, df in pin_per_admin_status_strata.items():
        # Create a copy of the first row to preserve the structure
        summed_df = df.iloc[0:1].copy()

        # Identify columns to skip from summation and columns to set to zero
        columns_to_skip = [col for col in df.columns if col.startswith('%') or col == admin_var  or col == 'Population group' or col == 'Category' or col == 'Area severity']
        columns_to_zero = [col for col in df.columns if col.startswith('%')]

        # Sum all numerical columns except the skipped ones
        for col in df.columns:
            if col not in columns_to_skip:
                summed_df[col] = df[col].sum()

        # Set non-sum columns with fixed values
        summed_df[admin_var] = 'whole country'
        summed_df['Population group'] = category
        if 'Area severity' in summed_df.columns:
            del summed_df['Area severity']

        # Set percentage columns to zero
        for col in columns_to_zero:
            summed_df[col] = 0

        # Add the modified DataFrame to the results dictionary
        summed_df = summed_df.iloc[:1]
        collapsed_results[category] = summed_df


    # Initialize the overview DataFrame with the first entry
    first_key = next(iter(collapsed_results))  # Get the first key from the dictionary
    overview_strata = collapsed_results[first_key].copy()

    # Iterate through all DataFrames in the dictionary and add their values to the overview DataFrame
    for category, df in collapsed_results.items():
        if category != first_key:  # Skip the initial DataFrame used for initialization
            overview_strata += df

    # Set final summary values
    overview_strata[admin_var] = 'Whole country'
    overview_strata['Population group'] = 'All population groups'
    overview_strata['Category'] = category_str

    overview_strata[label_perc_tot] = 0
    overview_strata[label_tot] = (overview_strata[label_tot3] +
                               overview_strata[label_tot4] +
                               overview_strata[label_tot5])

    cols = list(overview_strata.columns)
    cols.insert(cols.index(label_tot) + 1, cols.pop(cols.index('Category')))
    overview_strata = overview_strata[cols]

    return overview_strata
##--------------------------------------------------------------------------------------------
def reorder_severity_columns(df, admin_var = 'admin1', pop_gorup_var = 'pop_group'):
    # Define the fixed order for admin-related columns
    admin_cols = [admin_var, label_tot_population, 'In-School Children', 'Out-of-School Children', pop_gorup_var]

    # Extract all columns
    all_columns = list(df.columns)

    # Dictionary to hold severity levels
    severity_groups = {3: [], 4: [], 5: []}
    total_columns_map = {}

    # Categorize columns into % of children and ToT # children groups
    for col in all_columns:
        if "severity level 3" in col and "(ToT # children)" not in col:
            severity_groups[3].append(col)
        elif "severity level 4" in col and "(ToT # children)" not in col:
            severity_groups[4].append(col)
        elif "severity level 5" in col and "(ToT # children)" not in col:
            severity_groups[5].append(col)

        if "(ToT # children)" in col:
            base_col = col.replace(" (ToT # children)", "")
            total_columns_map[base_col] = col  # Map the base name to its total column

    # Construct the final column order
    final_columns = admin_cols  # Start with admin columns

    for severity in [3, 4, 5]:  # Order by severity level
        for perc_col in severity_groups[severity]:
            final_columns.append(perc_col)  # Add % of children column first
            base_col = perc_col.replace(" (% of children)", "")  # Get base column name

            if base_col in total_columns_map:  # If a total column exists, add it right after
                final_columns.append(total_columns_map[base_col])

    # Ensure missing total columns are still included at the end
    remaining_tot_cols = [col for col in df.columns if col not in final_columns and "(ToT # children)" in col]
    final_columns.extend(remaining_tot_cols)

    # Reorder the DataFrame
    return df[final_columns]


##--------------------------------------------------------------------------------------------
def add_disability_factor(df,factor=0.1, category = 'Disability'):
    # Copy the dataframe to avoid altering the original data
    result_df = df.copy()
    
    # Create the disability factor column
    category_name = category
    result_df[category_name] = factor
    
    # Add a category label
    result_df['Category'] = category_name
    
    # Define columns to keep
    columns_to_keep = ['Admin', category_name, 'Category']
    
    # Return the filtered DataFrame
    return {category_name: result_df[columns_to_keep]}



##--------------------------------------------------------------------------------------------
# %PiN AND #PiN PER ADMIN AND POPULATION GROUP for the strata: GENDER, SCHOOL-CYCLE 
def adjust_pin_by_strata_factor(pin_df, factor_df, category_label, tot_column, admin_var):
    # Merge the pin DataFrame with the factor DataFrame on the 'Admin_2' column
    factorized_df = pd.merge(
        pin_df, factor_df, 
        left_on=[admin_var], 
        right_on=["Admin"], 
        how='left'
    )
   # Columns that need to be adjusted by the factor
    columns_to_adjust = [col for col in factorized_df.columns if col.startswith('#') or col == tot_column ]
    del factorized_df['Admin']

    # Apply the multiplication for each column that needs adjustment
    for col in columns_to_adjust:
        factorized_df[col] *= factorized_df[category_label]
 
    if (category_label == "Girl"): 
        print('-----------------------------------------------------    factorized_df')
        print(factorized_df)
        print(factor_df)
        print(pin_df)

    # Drop the now unneeded factor column
    factorized_df.drop(columns=[category_label], inplace=True)
    return factorized_df



#--------------------------------------------------------------------------------------------
def aggregate_pin_per_admin_status(pin_per_admin_status, admin_var):
    # build combined_df WITHOUT changing the original dict
    cleaned = []
    for df in pin_per_admin_status.values():
        tmp = df.copy()
        # zero % cols if you need to
        pct_cols = [c for c in tmp.columns if c.startswith('%')]
        for c in pct_cols:
            tmp[c] = 0
        # drop only on the copy
        to_drop = [c for c in [label_perc_tot, label_tot, label_admin_severity] if c in tmp.columns]
        tmp = tmp.drop(columns=to_drop, errors='ignore')
        cleaned.append(tmp)

    combined_df = pd.concat(cleaned, ignore_index=True)

    grouped_df = combined_df.groupby([admin_var]).agg({
        label_tot_population: 'sum',
        label_perc2: 'sum',
        label_tot2: 'sum',
        label_perc3: 'sum',
        label_tot3: 'sum',
        label_perc4: 'sum',
        label_tot4: 'sum',
        label_perc5: 'sum',
        label_tot5: 'sum',
    }).reset_index()

    return grouped_df


########################################################################################################################################
########################################################################################################################################
##############################################    PIN CALCULATION FUNCTION    ##########################################################
########################################################################################################################################
########################################################################################################################################
def calculatePIN_with_JENA (data_combination, country, edu_data, household_data, choice_data, survey_data, ocha_data,mismatch_ocha_data,jena_data,
                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                age_var, gender_var,
                label, 
                admin_var, vector_cycle, start_school, status_var,
                mismatch_admin,
                selected_language):
    

    ## essential variables --------------------------------------------------------------------------------------------
    single_cycle = (vector_cycle[1] == 0)
    primary_start = 6
    if country == 'Afghanistan -- AFG': 
        primary_start = 7

    secondary_end = 17

    
    host_suggestion = ["Urban","PND","always_lived","general_pop",'non_deplace','Host Community',"Host community members",'host_communi', "always_lived","non_displaced_vulnerable",'host',"non_pdi","hote","menage_n_deplace","resident","lebanese","Populationnondéplacée","ocap","non_deplacee","Residents","yes","4"]
    IDP_suggestion = ['host_family','PDI',"Rural","displaced","IDP", 'New IDPs','pdi', 'idp', 'site','idp_host' ,"menage_deplace_interne", 'Out-of-camp','no',  'pdi_fam', '2', '1' ]
    returnee_suggestion = ['displaced_previously' ,'cb_returnee','retourne','ret','Returnee HH','returnee' ,'ukrainian moldovan','Returnees','5']
    refugee_suggestion = ['refugees','REF', 'refugee','refugie', 'refugie','prl', 'refugiee', '3']
    ndsp_suggestion = ['ndsp','Protracted IDPs', "hote affected by IDP",'displaced_camp', 'idp_site','pdi_site', "In-camp"]
    status_to_be_excluded = ['dnk', 'other', 'pnta', 'dont_know', 'no_answer', 'prefer_not_to_answer', 'pnpr', 'nsp', 'autre', 'do_not_know', 'decline']
    template_values = ['Host/Hôte',	'IDP/PDI',	'Returnees/Retournés', 'Refugees/Refugiees', 'Other']
    suggestions_mapping = {
        'Host/Hôte': host_suggestion,
        'IDP/PDI': IDP_suggestion,
        'Returnees/Retournés': returnee_suggestion,
        'Refugees/Refugiees': refugee_suggestion,
        'Other': ndsp_suggestion
    }

    ## admin level finding for the MSNA part. 
    admin_target = admin_var
    pop_group_var = status_var
    ocha_pop_data = ocha_data

    ocha_pop_data = ocha_pop_data.rename(columns={'Admin': 'Admin_label'})
    ocha_pop_data = ocha_pop_data.rename(columns={'Admin Pcode': 'Admin'})
    ocha_pop_data = ocha_pop_data.drop(columns=['Admin_label'])

    print(ocha_pop_data)

    admin_var = find_best_match(admin_target,  household_data)

    admin_column_rapresentative = []
    grouped_dict = {}
    if mismatch_admin:
        ocha_mismatch_list = mismatch_ocha_data
        # Create a defaultdict to store grouped data
        detailed_list = ocha_mismatch_list.iloc[:, 1].astype(str).tolist()  # Converting to string
        prefix_list = ocha_mismatch_list.iloc[:, 2].dropna().astype(str).tolist()  # Drop NaN and convert to string
        admin_low_ok_list = ocha_mismatch_list.iloc[:, 0].dropna().astype(str).tolist()  # Drop NaN and convert to string

        grouped_dict = defaultdict(list)
        # Iterate over each prefix in the prefix_list
        for prefix in prefix_list:
            # Find all detailed entries that start with the current prefix
            for detailed_entry in detailed_list:
                if detailed_entry.startswith(prefix):
                    grouped_dict[prefix].append(detailed_entry)
        # Convert defaultdict to a regular dictionary for better readability
        grouped_dict = dict(grouped_dict)
        # Print the resulting dictionary
        for key, value in grouped_dict.items():
            print(f"{key}: {value}")

        length_dict = categorize_levels_dynamic(prefix_list)

        admin_column_rapresentative = find_matching_columns_for_admin_levels(edu_data, household_data, prefix_list, admin_var)
        print('admin_column_rapresentative')
        print(admin_column_rapresentative)

    ####### ** 1 **       ------------------------------ manipulation and join between H and edu data   ------------------------------------------     #######
    ####### ** 2 **       ------------------------------ severity definition and calculation ------------------------------------------     #######
    # in the function add_severity

    print (edu_data.columns)
    
    ####### ** 3 **       ------------------------------ Analysis per ADMIN AND POPULATION GROUP ------------------------------------------     #######
    edu_data = edu_data[edu_data[access_var].notna()]
    edu_data = edu_data[edu_data['severity_category'].notna()]

    df = pd.DataFrame(edu_data)

    ## in-school and OoS children subset
    df_in_school = edu_data[edu_data['var.access'].isin([1])]
    df_oos = edu_data[edu_data['var.access'].isin([0])]



    analysis_config_subset = {
        'var.access': {'df': df, 'target_var': 'var.access'},
        'var.teacher': {'df': df_in_school, 'target_var': 'var.teacher'},
        'var.hazard': {'df': df_in_school, 'target_var': 'var.hazard'},
        'var.idp': {'df': df_in_school, 'target_var': 'var.idp'},
        'var.occupation': {'df': df_in_school, 'target_var': 'var.occupation'},
        'var.barrier4': {'df': df_oos, 'target_var': 'var.barrier4'},
        'var.barrier5': {'df': df_oos, 'target_var': 'var.barrier5'},
        barrier_var: {'df': df_oos, 'target_var': barrier_var}
    }


    results_dict = {} 
    
    if mismatch_admin:
        detailed_list = ocha_mismatch_list.iloc[:, 1].astype(str).tolist()  # Converting to string
        admin_up_msna = ocha_mismatch_list.iloc[:, 2].dropna().astype(str).tolist()  # Drop NaN and convert to string
        admin_low_ok_list = ocha_mismatch_list.iloc[:, 0].dropna().astype(str).tolist()  # Drop NaN and convert to string

        #print(admin_up_msna)

        for analysis_var, config in analysis_config_subset.items():
            source_df = config['df']
            target_var = config['target_var']
            results_dict[analysis_var] = run_mismatch_admin_analysis(
                source_df,
                admin_var,
                admin_column_rapresentative,
                pop_group_var,
                analysis_variable=target_var,
                admin_low_ok_list=admin_low_ok_list,
                prefix_list=admin_up_msna,
                grouped_dict=grouped_dict
            )
                   
    else: ## no mistmach on admin and unit of analysis
        for analysis_var, config in analysis_config_subset.items():
            source_df = config['df']
            target_var = config['target_var']
            results_dict[analysis_var] = calculate_prop(
                df=source_df,
                admin_var=admin_var,
                pop_group_var=pop_group_var,
                target_var=target_var
            )    
       
        # Reduce the index for all results
        for key in results_dict:
            results_dict[key] = reduce_index(results_dict[key], 0, pop_group_var)


    # Extract results into individual variables if needed
    indicator_access_list = results_dict.get('var.access')
    indicator_teacher_list = results_dict.get('var.teacher')
    indicator_hazard_list = results_dict.get('var.hazard')
    indicator_idp_list = results_dict.get('var.idp')
    indicator_occupation_list = results_dict.get('var.occupation')
    indicator_barrier4_list = results_dict.get('var.barrier4')
    indicator_barrier5_list = results_dict.get('var.barrier5')
    indicator_barrier_list = results_dict.get(barrier_var)



    # Clean indicator columns
    indicator_access_list = clean_indicator_columns(indicator_access_list, 'indicator_access_list')
    indicator_teacher_list = clean_indicator_columns(indicator_teacher_list, 'indicator_teacher_list')
    indicator_hazard_list = clean_indicator_columns(indicator_hazard_list, 'indicator_hazard_list')
    indicator_idp_list = clean_indicator_columns(indicator_idp_list, 'indicator_idp_list')
    indicator_occupation_list = clean_indicator_columns(indicator_occupation_list, 'indicator_occupation_list')
    indicator_barrier4_list = clean_indicator_columns(indicator_barrier4_list, 'indicator_barrier4_list')
    indicator_barrier5_list = clean_indicator_columns(indicator_barrier5_list, 'indicator_barrier5_list')

    pin_by_indicator_status = {}
    # List of all indicator DataFrames grouped by category
    indicator_dataframes = [
        indicator_access_list,
        indicator_teacher_list,
        indicator_hazard_list,
        indicator_idp_list,
        indicator_occupation_list,
        indicator_barrier4_list,
        indicator_barrier5_list,
        indicator_barrier_list
    ]
    pin_by_indicator_status_list = process_indicator_dataframes(
        indicator_access_list=indicator_access_list,
        indicator_dataframes=indicator_dataframes,
        choice_data=choice_data,
        selected_severity_4_barriers=selected_severity_4_barriers,
        selected_severity_5_barriers=selected_severity_5_barriers,
        label_var=label,
        admin_var=admin_var,
        pop_group_var=pop_group_var
    )
    





    ################### use msna acceess rate to get the subsets of in-school and oos
    ####### ** 1 **       ------------------------------ step 1: extract access rate from MSNA by pop group
    access_rate_df = None
    access_rate_list={}

    for category, rate_pop_df in pin_by_indicator_status_list.items():
        # Start with the base DataFrame for this category
        print(category)
        rate_pop_df = rate_pop_df.copy()

        new_col_name = f"rate_indicator_access"
        rate_pop_df = rate_pop_df.iloc[:, [0, 2]]  # Keep admin1 and rate_indicator_access
        rate_pop_df.columns = [admin_var, new_col_name]  # Rename columns  

        access_rate_list [category] = rate_pop_df


    
    for category, df in access_rate_list.items():
        print("-------------==========================================   "+category)
        print(df)

    ####### ** 3 **       ------------------------------ step 3: matching between the admin and the ocha population data
    ## finding the match between the OCHA status cathegory and the country status. 
    status_values = [status for status in edu_data[pop_group_var].unique() if status not in status_to_be_excluded]# Retrieve unique values directly without converting to lowercase
    for key, suggestions in suggestions_mapping.items():
        suggestions_mapping[key] = suggestions  # keeping original case

    mapped_statuses = map_template_to_status(template_values, suggestions_mapping, status_values)
    print (mapped_statuses)
    category_data_frames = extract_status_data(ocha_pop_data, mapped_statuses, pop_group_var)# Extract population figures based on mapped statuses without modifying the case

    for category, df in category_data_frames.items():
        df.rename(columns={'Admin': admin_var}, inplace=True)
        print("--------------------------   "+category)
        print(df)

    ocha_inschool_oos_subset = {}
    for category, ocha_df in category_data_frames.items():     
        msna_access_rate = access_rate_list[category]

        ocha_inschool_oos_df = ocha_df.merge(msna_access_rate, on=admin_var)

        ocha_inschool_oos_df.drop(columns=[pop_group_var], errors="ignore", inplace=True)

        ocha_inschool_oos_df[label_tot_inschool] = ocha_inschool_oos_df[label_tot_population] * ocha_inschool_oos_df['rate_indicator_access']
        ocha_inschool_oos_df[label_tot_OoS] = ocha_inschool_oos_df[label_tot_population] * (1- ocha_inschool_oos_df['rate_indicator_access'])

        ocha_inschool_oos_subset[category] = ocha_inschool_oos_df

    for category, df in ocha_inschool_oos_subset.items():
        print("--------------------------   "+category)
        print(df)    

    ## ------------------------------------------------------------------------ 
    ## ---------------------------  jena manipulation ------------------------ 
    ## ------------------------------------------------------------------------ 

    print(jena_data.columns)
    # 1) Rename first column to admin_var
    jena_data = jena_data.rename(columns={jena_data.columns[0]: admin_var})

    # 2) Collapse empty strings to NaN and drop all-NaN columns
    jena_data = jena_data.replace(r'^\s*$', np.nan, regex=True)
    jena_data = jena_data.dropna(axis=1, how='all')

    # 3) Rename only columns that still exist
    jena_column_rename = {
        col: new_name
        for col, new_name in {
            'Enrolled students -- Children/Enfants (5-17)': 'school_children',
            'school PTR': 'sev3_indicator',
            'threshold PTR (Change it according to the context average PTR)': 'sev3_th_indicator',
            'sev4 - Protection indicator value -- continuous / discrete numerical variable': 'sev4_indicator_number',
            'sev4 - threshold for numerical variable --> if above the child is in need': 'sev4_th_indicator_number',
            'sev5 - Protection indicator value -- continuous / discrete numerical variable': 'sev5_indicator_number',
            'sev5 - threshold for numerical variable --> if above the child is in need': 'sev5_th_indicator_number',
            'sev4 - Protection indicator value -- categorical variable': 'sev4_indicator_cat',
            'sev4 - category for categorical variable --> if selected the child is in need': 'sev4_th_indicator_cat',
            'sev5 - Protection indicator value -- categorical variable': 'sev5_indicator_cat',
            'sev5 - category for categorical variable --> if selected the child is in need': 'sev5_th_indicator_cat'
        }.items()
        if col in jena_data.columns
    }
    jena_data = jena_data.rename(columns=jena_column_rename)

    print(jena_data.columns)

    # 4) Create output columns 
    jena_data['sev3_children'] = 0
    jena_data['sev4_children'] = 0
    jena_data['sev5_children'] = 0

    # 5) Safe numeric coercion (only for cols that exist)
    num_cols = [c for c in [
        'sev3_indicator','sev3_th_indicator',
        'sev4_indicator_number','sev4_th_indicator_number',
        'sev5_indicator_number','sev5_th_indicator_number',
        'school_children'
    ] if c in jena_data.columns]
    if num_cols:
        jena_data[num_cols] = jena_data[num_cols].apply(pd.to_numeric, errors='coerce')

    # 6) Normalized categorical series (only if present)
    def _norm(s):
        return s.astype(str).str.strip().str.lower()

    sev4_cat = _norm(jena_data['sev4_indicator_cat']) if 'sev4_indicator_cat' in jena_data else None
    sev4_th  = _norm(jena_data['sev4_th_indicator_cat']) if 'sev4_th_indicator_cat' in jena_data else None
    sev5_cat = _norm(jena_data['sev5_indicator_cat']) if 'sev5_indicator_cat' in jena_data else None
    sev5_th  = _norm(jena_data['sev5_th_indicator_cat']) if 'sev5_th_indicator_cat' in jena_data else None

    # 7) Build masks with strict priority: sev5 → sev4 → sev3
    # sev5: prefer categorical equality if cat cols exist; else numeric >
    m5_cat = (sev5_cat is not None) and (sev5_th is not None) and (sev5_cat.notna() & sev5_th.notna() & (sev5_cat == sev5_th))
    m5_num = all(c in jena_data.columns for c in ['sev5_indicator_number','sev5_th_indicator_number']) and \
            (jena_data['sev5_indicator_number'].notna() & jena_data['sev5_th_indicator_number'].notna() &
            (jena_data['sev5_indicator_number'] > jena_data['sev5_th_indicator_number']))
    m5 = m5_cat if isinstance(m5_cat, pd.Series) else False
    m5 = m5 | (m5_num if isinstance(m5_num, pd.Series) else False)

    # sev4: only where not sev5; prefer cat, else numeric >
    m4_cat = (sev4_cat is not None) and (sev4_th is not None) and (sev4_cat.notna() & sev4_th.notna() & (sev4_cat == sev4_th))
    m4_num = all(c in jena_data.columns for c in ['sev4_indicator_number','sev4_th_indicator_number']) and \
            (jena_data['sev4_indicator_number'].notna() & jena_data['sev4_th_indicator_number'].notna() &
            (jena_data['sev4_indicator_number'] > jena_data['sev4_th_indicator_number']))
    m4 = m4_cat if isinstance(m4_cat, pd.Series) else False
    m4 = m4 | (m4_num if isinstance(m4_num, pd.Series) else False)
    m4 = (~m5) & (m4 if isinstance(m4, pd.Series) else False)

    # sev3: only where neither sev5 nor sev4, numeric >
    m3 = all(c in jena_data.columns for c in ['sev3_indicator','sev3_th_indicator']) and \
        (jena_data['sev3_indicator'].notna() & jena_data['sev3_th_indicator'].notna() &
        (jena_data['sev3_indicator'] > jena_data['sev3_th_indicator']))
    m3 = (~m5) & (~m4) & (m3 if isinstance(m3, pd.Series) else False)

    # 8) Assign children counts (keep zeros otherwise)
    if 'school_children' in jena_data.columns:
        jena_data['sev5_children'] = np.where(m5, jena_data['school_children'], 0)
        jena_data['sev4_children'] = np.where(m4, jena_data['school_children'], 0)
        jena_data['sev3_children'] = np.where(m3, jena_data['school_children'], 0)
    else:
        # Optional: warn if school_children is missing
        print("Warning: 'school_children' column not present; sev*_children remain 0.")


    children_cols = ['school_children', 'sev3_children', 'sev4_children', 'sev5_children']

    # Group by and sum
    jena_severity = jena_data.groupby(admin_var)[children_cols].sum().reset_index()

    # 4) Create % severity columns 
    jena_severity['sev3_inschool'] =  jena_severity['sev3_children']/jena_severity['school_children'] 
    jena_severity['sev4_inschool'] =  jena_severity['sev4_children']/jena_severity['school_children'] 
    jena_severity['sev5_inschool'] =  jena_severity['sev5_children']/jena_severity['school_children'] 


    # Keep only the rate columns from JENA
    jena_rates = jena_severity[[admin_var, 'sev3_inschool', 'sev4_inschool', 'sev5_inschool']].copy()


    # Merge per category
    merged_ocha_jena = {}
    for category, ocha_df in ocha_inschool_oos_subset.items():
        merged_ocha_jena[category] = (
            ocha_df.merge(jena_rates, on=admin_var, how='left')
        )



    ####### ** 4 **       ------------------------------ step 1: extract aggreaavting circumns rate from MSNA by pop group
    agg_rate_list = {}

    # Columns you want to keep (plus admin_var)
    target_cols = [
        'subsetOoS_sev4_aggravating_circumstances',
        'subsetOoS_sev5_aggravating_circumstances'
    ]

    for category, rate_pop_df in pin_by_indicator_status_list.items():
        rate_pop_df = rate_pop_df.copy()
        print(rate_pop_df.columns)
        # Keep admin_var and the target columns if they exist
        cols_to_keep = [admin_var] + [c for c in target_cols if c in rate_pop_df.columns]
        rate_pop_df = rate_pop_df[cols_to_keep]

        agg_rate_list[category] = rate_pop_df





    merged_ocha_jena_msna = {}

    # Merge per category
    for category, merge1 in merged_ocha_jena.items():
        agg_cat = agg_rate_list[category]
        merged_ocha_jena_msna[category] = (
            merge1.merge(agg_cat, on=admin_var, how='left')
        )


    for category, df in merged_ocha_jena_msna.items():
        for c in ['sev3_inschool','sev4_inschool','sev5_inschool',
                'subsetOoS_sev4_aggravating_circumstances',
                'subsetOoS_sev5_aggravating_circumstances',
                label_tot_inschool, label_tot_OoS, label_tot_population]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
        # keep aggravating rates within [0,1] and compute residual safely
        df['subsetOoS_sev4_aggravating_circumstances'] = df.get('subsetOoS_sev4_aggravating_circumstances', 0).clip(0,1)
        df['subsetOoS_sev5_aggravating_circumstances'] = df.get('subsetOoS_sev5_aggravating_circumstances', 0).clip(0,1)
        merged_ocha_jena_msna[category] = df


    pin_jena_msna = {}
    for category, pin_cat in merged_ocha_jena_msna.items():
        for label in [label_perc2,label_tot2, label_perc3,label_tot3, label_perc4, label_tot4,label_perc5,label_tot5, 
                     label_perc_out,label_tot_out, label_perc_acc, label_tot_acc, label_perc_lc, label_tot_lc, label_perc_penv,label_tot_penv, label_perc_agg, label_tot_agg, 
                     label_perc_tot,label_tot,label_admin_severity]:
            pin_cat[label] = 0

        pin_cat[label_tot5] = pin_cat[label_tot_inschool]*pin_cat['sev5_inschool'] + pin_cat[label_tot_OoS]*pin_cat['subsetOoS_sev5_aggravating_circumstances']
        
        pin_cat[label_tot4] = pin_cat[label_tot_inschool]*pin_cat['sev4_inschool'] +  pin_cat[label_tot_OoS]*pin_cat['subsetOoS_sev4_aggravating_circumstances']
        
        pin_cat[label_tot3] = pin_cat[label_tot_inschool]*pin_cat['sev3_inschool'] +  (pin_cat[label_tot_OoS]*(1-pin_cat['subsetOoS_sev5_aggravating_circumstances'] - pin_cat['subsetOoS_sev4_aggravating_circumstances']))

        pin_cat[label_tot2] = (pin_cat[label_tot_population] - (pin_cat[label_tot3] + pin_cat[label_tot4] + pin_cat[label_tot5])).clip(lower=0)


        pin_cat[label_tot_lc] = pin_cat[label_tot_inschool]*pin_cat['sev3_inschool']
        pin_cat[label_tot_penv] = pin_cat[label_tot_inschool]*pin_cat['sev4_inschool'] + pin_cat[label_tot_inschool]*pin_cat['sev5_inschool']
        pin_cat[label_tot_agg] = pin_cat[label_tot_OoS]*pin_cat['subsetOoS_sev5_aggravating_circumstances'] +  pin_cat[label_tot_OoS]*pin_cat['subsetOoS_sev4_aggravating_circumstances']
        pin_cat[label_tot_acc] = pin_cat[label_tot_OoS]*(1-pin_cat['subsetOoS_sev5_aggravating_circumstances'] - pin_cat['subsetOoS_sev4_aggravating_circumstances'])
        pin_cat[label_tot_out] = (pin_cat[label_tot_population] - (pin_cat[label_tot_lc] + pin_cat[label_tot_penv] + pin_cat[label_tot_agg] + pin_cat[label_tot_acc] ) ).clip(lower=0)

        # percentages (guard divide-by-zero)
        denom = pin_cat[label_tot_population].replace(0, np.nan)
        pin_cat[label_perc3] = (pin_cat[label_tot3] / denom).fillna(0.0)
        pin_cat[label_perc4] = (pin_cat[label_tot4] / denom).fillna(0.0)
        pin_cat[label_perc5] = (pin_cat[label_tot5] / denom).fillna(0.0)
        pin_cat[label_perc2] = (pin_cat[label_tot2] / denom).fillna(0.0)
        pin_cat[label_tot]   = pin_cat[label_tot3] + pin_cat[label_tot4] + pin_cat[label_tot5]
        pin_cat[label_perc_tot] = (pin_cat[label_tot] / denom).fillna(0.0)

        pin_cat[label_perc_acc] = (pin_cat[label_tot_acc] / denom).fillna(0.0)
        pin_cat[label_perc_lc] = (pin_cat[label_tot_lc] / denom).fillna(0.0)
        pin_cat[label_perc_penv] = (pin_cat[label_tot_penv] / denom).fillna(0.0)
        pin_cat[label_perc_agg] = (pin_cat[label_tot_agg] / denom).fillna(0.0)
        pin_cat[label_perc_out] = (pin_cat[label_tot_out] / denom).fillna(0.0)


                # Define conditions based on specified logic
        conditions = [
            pin_cat[label_perc5] > 0.2,
            (pin_cat[label_perc5] + pin_cat[label_perc4]) > 0.2,
            (pin_cat[label_perc5] + pin_cat[label_perc4] + pin_cat[label_perc3]) > 0.2,
            (pin_cat[label_perc5] + pin_cat[label_perc4] + pin_cat[label_perc3] + pin_cat[label_perc2]) > 0.2
        ]
        # Corresponding values for each condition
        choices = ['5', '4', '3', '1-2']
        # Apply the conditions to determine admin severity
        pin_cat[label_admin_severity] = np.select(conditions, choices, default='0')


        pin_jena_msna[category] =pin_cat

    
    
    # Remove intermediate columns from each category DataFrame
    cols_to_remove_pin = [
        'subsetOoS_sev4_aggravating_circumstances',
        'subsetOoS_sev5_aggravating_circumstances',
        'sev3_inschool', 'sev4_inschool', 'sev5_inschool',
        label_tot_inschool, label_tot_OoS,
        'rate_indicator_access', label_perc_out,label_tot_out, label_perc_acc, label_tot_acc, label_perc_lc, label_tot_lc, label_perc_penv,label_tot_penv, label_perc_agg, label_tot_agg
    ]
    cols_to_remove_dimension = [
        'subsetOoS_sev4_aggravating_circumstances',
        'subsetOoS_sev5_aggravating_circumstances',
        'sev3_inschool', 'sev4_inschool', 'sev5_inschool',
        label_tot_inschool, label_tot_OoS,label_perc2,label_tot2, label_perc3,label_tot3, label_perc4, label_tot4,label_perc5,label_tot5,
        'rate_indicator_access'
    ]

    Tot_PiN_JIAF = {}
    for category, df in pin_jena_msna.items():
        to_drop = [c for c in cols_to_remove_pin if c in df.columns]
        Tot_PiN_JIAF[category] = df.drop(columns=to_drop, errors='ignore')


    Tot_Dimension_JIAF = {}
    for category, df in pin_jena_msna.items():
        to_drop = [c for c in cols_to_remove_dimension if c in df.columns]
        Tot_Dimension_JIAF[category] = df.drop(columns=to_drop, errors='ignore')
        
    #Tot_Dimension_JIAF


      ####### ** 5 **       ------------------------------ creating tables with factors for the gender and school-cycle categories ------------------------------------------     #######
    ## calculate the difference population group per school-cycle according to the country and the tot-children population 
    factor_cycle = [0.5,0.5,0]
    factor_disability = 0.1
    ## create table per strata
    category_tot = 'All'
    category_girl = 'Girl'
    category_boy = 'Boy'
    category_ece= 'ECE'
    category_primary= 'primary'
    category_upper_primary= 'intermediate level'
    category_secondary= 'secondary'
    category_disability = 'Disability'
    children_tot_col = 'ToT -- Children/Enfants (5-17)'
    girls_tot_col = 'ToT -- Girls/Filles (5-17)'
    boys_tot_col = 'ToT -- Boys/Garcons (5-17)'
    ece_tot_col = '5yo -- Children/Enfants'

    # Calculate category factors
    category_factors = {
        **calculate_category_factors(ocha_pop_data, children_tot_col, girls_tot_col, category_girl),
        **calculate_category_factors(ocha_pop_data, children_tot_col, boys_tot_col, category_boy),
        **calculate_category_factors(ocha_pop_data, children_tot_col, ece_tot_col, category_ece)
    }
    # Calculate factors for each school cycle
    school_cycle_factors = calculate_cycle_factors(ocha_pop_data, factor_cycle, primary_start, secondary_end, vector_cycle, single_cycle)
    # Combine all factors into one dictionary
    factor_category = {**category_factors, **school_cycle_factors}
    disability_factors = add_disability_factor(ocha_pop_data, factor_disability,category_disability )
    # Update the factor_category dictionary with the new disability category
    factor_category.update(disability_factors)



    
    ####### ** strata 6.A **       ------------------------------ %PiN AND #PiN PER ADMIN AND POPULATION GROUP using ocha figures ------------------------------------------     #######
    factor_girl_df = factor_category['Girl']
    factor_girl_df.drop('Category', axis=1, inplace=True)
    factor_boy_df = factor_category['Boy']
    factor_boy_df.drop('Category', axis=1, inplace=True)
    factor_ece_df = factor_category['ECE']
    factor_ece_df.drop('Category', axis=1, inplace=True)
    factor_primary_df = factor_category['primary']
    factor_primary_df.drop('Category', axis=1, inplace=True)
    factor_secondary_df = factor_category['secondary']
    factor_secondary_df.drop('Category', axis=1, inplace=True)
    factor_upper_primary_df = factor_category['intermediate level']
    factor_upper_primary_df.drop('Category', axis=1, inplace=True)



 ####### ** 7 **       ------------------------------ %PiN AND #PiN PER ADMIN AND POPULATION GROUP for the strata: GENDER, SCHOOL-CYCLE ------------------------------------------     #######
    ## PiN
    pin_per_admin_status_girl = {}
    pin_per_admin_status_boy = {}
    pin_per_admin_status_ece = {}
    pin_per_admin_status_primary = {}
    pin_per_admin_status_upper_primary = {}
    pin_per_admin_status_secondary = {}
    pin_per_admin_status_disabilty = {}

    for category, df in Tot_PiN_JIAF.items():
        pin_per_admin_status_girl[category] = adjust_pin_by_strata_factor(df, factor_category[category_girl], category_girl, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_boy[category] = adjust_pin_by_strata_factor(df, factor_category[category_boy], category_boy, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_ece[category] = adjust_pin_by_strata_factor(df, factor_category[category_ece], category_ece, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_primary[category] = adjust_pin_by_strata_factor(df, factor_category[category_primary], category_primary, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_upper_primary[category] = adjust_pin_by_strata_factor(df, factor_category[category_upper_primary], category_upper_primary, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_secondary[category] = adjust_pin_by_strata_factor(df, factor_category[category_secondary], category_secondary, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_disabilty[category] = adjust_pin_by_strata_factor(df, factor_category[category_disability], category_disability, tot_column= label_tot_population, admin_var=admin_var)



    overall_pin_per_admin_df = aggregate_pin_per_admin_status(Tot_PiN_JIAF, admin_var)

    ####### ** 8.C **       ------------------------------ calculate tot PiN --> 3+ and admin severity for overall_pin_per_admin_df ------------------------------------------     #######
    Tot_PiN_by_admin = overall_pin_per_admin_df
    # Iterate over the pin_per_admin_status dictionary to apply the new operations
        # Initialize new columns for percentage total, total PiN, and admin severity
    Tot_PiN_by_admin[label_perc_tot] = 0
    Tot_PiN_by_admin[label_tot] = 0
    Tot_PiN_by_admin[label_admin_severity] = 0

    Tot_PiN_by_admin[label_perc2] = Tot_PiN_by_admin[label_tot2]/Tot_PiN_by_admin[label_tot_population]
    Tot_PiN_by_admin[label_perc3] = Tot_PiN_by_admin[label_tot3]/Tot_PiN_by_admin[label_tot_population]
    Tot_PiN_by_admin[label_perc4] = Tot_PiN_by_admin[label_tot4]/Tot_PiN_by_admin[label_tot_population]
    Tot_PiN_by_admin[label_perc5] = Tot_PiN_by_admin[label_tot5]/Tot_PiN_by_admin[label_tot_population]

    # Reorder columns to place new columns at desired positions
    cols = list(Tot_PiN_by_admin.columns)
    cols.insert(cols.index(label_tot5) + 1, cols.pop(cols.index(label_perc_tot)))
    cols.insert(cols.index(label_perc_tot) + 1, cols.pop(cols.index(label_tot)))
    cols.insert(cols.index(label_tot) + 1, cols.pop(cols.index(label_admin_severity)))
    Tot_PiN_by_admin = Tot_PiN_by_admin[cols]

    # Calculate the total percentage and total PiN for severity levels 3+
    Tot_PiN_by_admin[label_perc_tot] = (Tot_PiN_by_admin[label_perc3] +
                                    Tot_PiN_by_admin[label_perc4] +
                                    Tot_PiN_by_admin[label_perc5])

    Tot_PiN_by_admin[label_tot] = (Tot_PiN_by_admin[label_tot3] +
                            Tot_PiN_by_admin[label_tot4] +
                            Tot_PiN_by_admin[label_tot5])

    # Define conditions based on specified logic
    conditions = [
        Tot_PiN_by_admin[label_perc5] > 0.2,
        (Tot_PiN_by_admin[label_perc5] + Tot_PiN_by_admin[label_perc4]) > 0.2,
        (Tot_PiN_by_admin[label_perc5] + Tot_PiN_by_admin[label_perc4] + Tot_PiN_by_admin[label_perc3]) > 0.2,
        (Tot_PiN_by_admin[label_perc5] + Tot_PiN_by_admin[label_perc4] + Tot_PiN_by_admin[label_perc3] + Tot_PiN_by_admin[label_perc2]) > 0.2
    ]
    # Corresponding values for each condition
    choices = ['5', '4', '3', '1-2']
    # Apply the conditions to determine admin severity
    Tot_PiN_by_admin[label_admin_severity] = np.select(conditions, choices, default='0')

    tot_5_17_label = 'TOTAL (5-17 y.o.)'
    girl_5_17_label = 'Girls (5-17 y.o.)'
    boy_5_17_label = 'Boys (5-17 y.o.)'
    ece_5yo_label = 'ECE (5 y.o.)'





    ####### ** 9 **       ------------------------------  preparation for overview--> SUM all the admin per population group and per strata ------------------------------------------     #######
    overview_ToT = collapse_and_summarize(Tot_PiN_JIAF, tot_5_17_label, admin_var=admin_var)
    overview_girl = collapse_and_summarize(pin_per_admin_status_girl, girl_5_17_label, admin_var=admin_var)
    overview_boy = collapse_and_summarize(pin_per_admin_status_boy, boy_5_17_label, admin_var=admin_var)
    overview_ece = collapse_and_summarize(pin_per_admin_status_ece, ece_5yo_label, admin_var=admin_var)
    overview_primary = collapse_and_summarize(pin_per_admin_status_primary, 'Primary school', admin_var=admin_var)
    overview_upper_primary = collapse_and_summarize(pin_per_admin_status_upper_primary, 'Intermediate school-level', admin_var=admin_var)
    overview_secondary = collapse_and_summarize(pin_per_admin_status_secondary, 'Secondary school', admin_var=admin_var)
    overview_disabilty = collapse_and_summarize(pin_per_admin_status_disabilty, 'Children with disability', admin_var=admin_var)

    for category, df in Tot_PiN_JIAF.items():
        print(df.columns)

    

    collapsed_results_pop = {}
    for category, df in Tot_PiN_JIAF.items():
            # Create a copy of the first row to preserve the structure
            summed_df = df.iloc[0:1].copy()

            # Identify columns to skip from summation and columns to set to zero
            columns_to_skip = [col for col in df.columns if col.startswith('%') or col == admin_var  or col == 'Population group' or col == 'Category' or col== label_admin_severity]
            columns_to_zero = [col for col in df.columns if col.startswith('%')]

            # Sum all numerical columns except the skipped ones
            for col in df.columns:
                if col not in columns_to_skip:
                    summed_df[col] = df[col].sum()

            # Set non-sum columns with fixed values
            summed_df[admin_var] = 'whole country'
            summed_df['Population group'] = category
            del summed_df[label_admin_severity]
            # Set percentage columns to zero
            for col in columns_to_zero:
                summed_df[col] = 0

            # Add the modified DataFrame to the results dictionary
            summed_df = summed_df.iloc[:1]
            collapsed_results_pop[category] = summed_df


    

    ####### ** 10.A **       ------------------------------  Creating OVERVIEW file ------------------------------------------     #######
    dfs_overview_ToT = []
    dfs_overview_ToT.append(overview_ToT)

    # Add the single-row entries from collapsed_results_pop
    for category, df in collapsed_results_pop.items():
        if not df.empty:  # Ensure the DataFrame is not empty
            single_row = df.iloc[0].copy()
            if country != 'Afghanistan -- AFG':
                single_row['Category'] = f"{category} (5-17 y.o.)"
            else:
                single_row['Category'] = f"{category} (6-17 y.o.)"

            # Convert the Series to a DataFrame with a single row and append it to the list
            single_row_df = single_row.to_frame().T
            dfs_overview_ToT.append(single_row_df)
        else:
            print(f"Warning: DataFrame for category {category} is empty, creating a dummy row.")

    # important overview for the total and population figures. It has all the percentages and total numbers
    final_overview_df = pd.concat(dfs_overview_ToT, ignore_index=True) ## table with all severities and tot and pop_group

    # table for the overview with total number to feed to OCHA and for the first sheet of the output
    strata_summarized_data_OCHA = [
        overview_girl,
        overview_boy,
        overview_ece,
        overview_disabilty
    ]
    small_overview = dfs_overview_ToT
    small_overview.extend(strata_summarized_data_OCHA)
    # Concatenate all DataFrames in the list into a single DataFrame
    final_overview_df_OCHA = pd.concat(small_overview, ignore_index=True) ## table to reduce with all the population figures numbers

    ## organization and manipulation 
    cols = list(final_overview_df.columns)
    cols.insert(cols.index(admin_var) + 1, cols.pop(cols.index('Category')))
    final_overview_df = final_overview_df[cols]
    del final_overview_df[admin_var]
    final_overview_df = final_overview_df.rename(columns={'Category': 'Strata'})

    cols_ocha = list(final_overview_df_OCHA.columns)
    cols_ocha.insert(cols_ocha.index(admin_var) + 1, cols_ocha.pop(cols_ocha.index('Category')))
    final_overview_df_OCHA = final_overview_df_OCHA[cols_ocha]
    del final_overview_df_OCHA[admin_var]
    final_overview_df_OCHA = final_overview_df_OCHA.rename(columns={'Category': 'Strata'})

    final_overview_df[label_perc2] = final_overview_df[label_tot2]/final_overview_df[label_tot_population]
    final_overview_df[label_perc3] = final_overview_df[label_tot3]/final_overview_df[label_tot_population]
    final_overview_df[label_perc4] = final_overview_df[label_tot4]/final_overview_df[label_tot_population]
    final_overview_df[label_perc5] = final_overview_df[label_tot5]/final_overview_df[label_tot_population]
    final_overview_df[label_perc_tot] = final_overview_df[label_tot]/final_overview_df[label_tot_population]



    
    ####### ** 11 ** ------------------------------  Rounding and Saving the JIAF AND OCHA OUTPUT ------------------------------------------ #######
    # Define rounding parameters
    percentage_round = 1
    figures_round = 0

    # Process Tot_PiN_JIAF DataFrames
    for category, df in Tot_PiN_JIAF.items():
        rounding_dataframe(df, figures_round, percentage_round)
        df[label_tot_population] = pd.to_numeric(df[label_tot_population], errors='coerce').round(figures_round)


    rounding_dataframe(Tot_PiN_by_admin, figures_round, percentage_round)
    Tot_PiN_by_admin[label_tot_population] = pd.to_numeric(Tot_PiN_by_admin[label_tot_population], errors='coerce').round(figures_round)
    
    # Process Tot_Dimension_JIAF DataFrames
    for category, df in Tot_Dimension_JIAF.items():
        rounding_dataframe(df, figures_round, percentage_round)
        df[label_dimension_tot_population] = pd.to_numeric(df[label_dimension_tot_population], errors='coerce').round(figures_round)


    # Process final_overview_df
    rounding_dataframe(final_overview_df, figures_round, percentage_round)
    final_overview_df[label_tot_population] = pd.to_numeric(final_overview_df[label_tot_population], errors='coerce').round(figures_round)

    rounding_dataframe(final_overview_df_OCHA, figures_round, percentage_round)
    final_overview_df_OCHA[label_tot_population] = pd.to_numeric(final_overview_df_OCHA[label_tot_population], errors='coerce').round(figures_round)
    final_overview_df_OCHA = final_overview_df_OCHA[['Strata', 'Population group',label_tot]]



    country_label = country.replace(" ", "_").replace("--", "_").replace("/", "_")

    translation_dict = {
        label_perc2: '% niveaux de sévérité 1-2',
        label_perc3: '% niveau de sévérité 3',
        label_perc4: '% niveau de sévérité 4',
        label_perc5: '% niveau de sévérité 5',
        label_tot2: '# niveaux de sévérité 1-2',
        label_tot3: '# niveau de sévérité 3',
        label_tot4: '# niveau de sévérité 4',
        label_tot5: '# niveau de sévérité 5',
        label_perc_tot: '% Tot PiN (niveaux de sévérité 3-5)',
        label_tot: '# Tot PiN (niveaux de sévérité 3-5)',
        label_admin_severity: 'Sévérité de la zone',
        label_tot_population: 'Population totale',
        '5-17 y.o.': '5-17 ans',
        'Girls': 'Filles',
        'Boys': 'Garcons',
        '5 y.o.': "5 ans",
        'ECE': 'Éducation préscolaire',
        'All population groups': 'Tous les groupes de population',
        'Population group': 'Groupe de population',
        'Children with disability': 'Enfants en situation de handicap',
        "Primary school": "École primaire",
        "Intermediate school-level": "Niveau scolaire intermédiaire",
        "Secondary school":"École secondaire",
        "severity level 3 -- OoS children -- % of children not accessing education who do not face any aggravating circumstances": "Niveau de sévérité 3 -- enfants non scolarisés -- % d'enfants n'ayant pas accès à l'éducation et ne souffrant d'aucune circonstance aggravante",
        "severity level 3 -- in-school children -- % of children whose education was disrupted by teacher absence" : "Niveau de sévérité 3 -- enfants scolarisés -- % d'enfants dont l'éducation a été perturbée par l'absence d'un enseignant",
        "severity level 3 -- in-school children -- % of children whose education was disrupted by natural hazard" : "Niveau de sévérité 3 -- enfants scolarisés -- % d'enfants dont l'éducation a été perturbée par un risque naturel",
        "severity level 4 -- in-school children -- % of children whose education was disrupted by the school being used as shelter" : "Niveau de sévérité 4 -- enfants scolarisés -- % d'enfants dont l'éducation a été perturbée par l'utilisation de l'école comme abri",
        "severity level 5 -- in-school children -- % of children whose education was disrupted by the school being occupied by armed groups" : "Niveau de sévérité 5 -- enfants scolarisés -- % d'enfants dont l'éducation a été perturbée par l'occupation de l'école par des groupes armés",
        "severity level 3 -- OoS children -- ToT # of children not accessing education who do not face any aggravating circumstances": "Niveau de sévérité 3 -- enfants non scolarisés -- # d'enfants n'ayant pas accès à l'éducation et ne souffrant d'aucune circonstance aggravante",
        "severity level 3 -- in-school children -- ToT # of children whose education was disrupted by teacher absence" : "Niveau de sévérité 3 -- enfants scolarisés -- # d'enfants dont l'éducation a été perturbée par l'absence d'un enseignant",
        "severity level 3 -- in-school children -- ToT # of children whose education was disrupted by natural hazard" : "Niveau de sévérité 3 -- enfants scolarisés -- # d'enfants dont l'éducation a été perturbée par un risque naturel",
        "severity level 4 -- in-school children -- ToT # of children whose education was disrupted by the school being used as shelter" : "Niveau de sévérité 4 -- enfants scolarisés -- # d'enfants dont l'éducation a été perturbée par l'utilisation de l'école comme abri",
        "severity level 5 -- in-school children -- ToT # of children whose education was disrupted by the school being occupied by armed groups" : "Niveau de sévérité 5 -- enfants scolarisés -- # d'enfants dont l'éducation a été perturbée par l'occupation de l'école par des groupes armés",
        "severity level 4 -- OoS children -- % of children not accessing education due to the aggravating circumstance": "niveau de sévérité 4 -- enfants non scolarisés -- % d'enfants n'ayant pas accès à l'éducation en raison de la circonstance aggravante ",
        "severity level 5 -- OoS children -- % of children not accessing education due to the aggravating circumstance": "niveau de sévérité 5 -- enfants non scolarisés -- % d'enfants n'ayant pas accès à l'éducation en raison de la circonstance aggravante ",
        "severity level 4 -- OoS children -- ToT # of children not accessing education due to the aggravating circumstance": "niveau de sévérité 4 -- enfants non scolarisés -- # d'enfants n'ayant pas accès à l'éducation en raison de la circonstance aggravante ",
        "severity level 5 -- OoS children -- ToT # of children not accessing education due to the aggravating circumstance": "niveau de sévérité 5 -- enfants non scolarisés -- # d'enfants n'ayant pas accès à l'éducation en raison de la circonstance aggravante "
        }

    
    print(final_overview_df['Strata'].unique())


    if selected_language == 'French':
        final_overview_df = translate_labels(final_overview_df, translation_dict)
        final_overview_df_OCHA = translate_labels(final_overview_df_OCHA, translation_dict)
        Tot_PiN_by_admin = translate_labels(Tot_PiN_by_admin, translation_dict)
        Tot_PiN_JIAF = translate_labels(Tot_PiN_JIAF,translation_dict)


    






    return jena_severity, merged_ocha_jena, merged_ocha_jena_msna,pin_jena_msna, Tot_PiN_JIAF,Tot_Dimension_JIAF, final_overview_df_OCHA, final_overview_df, Tot_PiN_by_admin