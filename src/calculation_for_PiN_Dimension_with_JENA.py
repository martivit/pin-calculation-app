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
        optional_columns = [
            'indicator_access', 'indicator_teacher', 'indicator_hazard',
            'indicator_idp', 'indicator_occupation', 'indicator_barrier4', 'indicator_barrier5'
        ]
        essential_columns += [col for col in optional_columns if col in sample_df.columns]

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
