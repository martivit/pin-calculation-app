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
def apply_scenario_drops(df: pd.DataFrame, data_combination: str = "emmm", *, inplace: bool = False) -> pd.DataFrame:
    """
    Remove columns according to data_combination:
      - 'emmm': drop all subsetInSchool_sev3/4/5_emis (+ N___ versions)
      - 'eemm': drop subsetInSchool_sev3_indicator* (+ N___) AND subsetInSchool_sev4/5_emis (+ N___)
                (keep subsetInSchool_sev3_emis)
      - 'eeem': drop subsetInSchool_sev3/4/5_indicator* (+ N___) ; keep all *_emis
    """
    target = df if inplace else df.copy()
    dc = (data_combination or "").lower()

    patterns = []
    if dc == "emmm":
        patterns = [r"^(?:N___)?subsetInSchool_sev[345]_emis$"]
    elif dc == "eemm":
        patterns = [
            r"^(?:N___)?subsetInSchool_sev3_indicator_.*",
            r"^(?:N___)?subsetInSchool_sev[45]_emis$",
        ]
    elif dc == "eeem":
        patterns = [
            r"^(?:N___)?subsetInSchool_sev3_indicator_.*",
            r"^(?:N___)?subsetInSchool_sev4_indicator_.*",
            r"^(?:N___)?subsetInSchool_sev5_indicator_.*",
        ]

    if patterns:
        to_drop = [c for c in target.columns if any(re.match(p, c) for p in patterns)]
        if to_drop:
            target.drop(columns=to_drop, inplace=True, errors="ignore")

    return target

    # i have to remove some columns according to the data_combination
    #if i have emmm we remove the sev3/sev4/sev5_emis
    #if i have eemm we remove sev3 teacher and hazard
    # if i have eeem  we remove sev3 teacher and hazard and sev 4 displaced and sev4 occupation
    # then we can do add_additional_severity_column which calculated the severity2 for in school as difference between the pin and total kids in school
    
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
def recompute_sev3_emis_if_eemm(df: pd.DataFrame, data_combination: str = "eemm", enrolled_col: str = "E") -> pd.DataFrame:
    """
    Only for scenario 'eemm':
      If residual = E - N___subsetInSchool_sev3_emis - (sev4_indicators + sev5_indicators) < 0
      then set:
        N___subsetInSchool_sev3_emis = max(E - (sev4_indicators + sev5_indicators), 0)
        subsetInSchool_sev3_emis     = N___subsetInSchool_sev3_emis / E
      Otherwise leave the row unchanged.
    """
    if (data_combination or "").lower() != "eemm":
        return df

    out = df.copy()
    sev3N_col = "N___subsetInSchool_sev3_emis"

    # If the numerator column doesn't exist, nothing to fix
    if sev3N_col not in out.columns or enrolled_col not in out.columns:
        return out

    # Sum weighted in-school indicator pieces for sev4/5
    sev4_cols = [c for c in out.columns if c.startswith("N___subsetInSchool_sev4_indicator_")]
    sev5_cols = [c for c in out.columns if c.startswith("N___subsetInSchool_sev5_indicator_")]
    sev4_sum = out[sev4_cols].sum(axis=1) if sev4_cols else 0.0
    sev5_sum = out[sev5_cols].sum(axis=1) if sev5_cols else 0.0
    other_inschool = sev4_sum + sev5_sum

    E = out[enrolled_col].astype(float)
    cur_sev3N = out[sev3N_col].astype(float)

    residual = E.fillna(0) - cur_sev3N.fillna(0) - other_inschool.fillna(0)
    mask = residual < 0  # only fix rows that go negative

    if mask.any():
        new_sev3N = (E.fillna(0) - other_inschool.fillna(0)).clip(lower=0)
        out.loc[mask, sev3N_col] = new_sev3N[mask]

        # Recompute the rate only on those rows; create the rate col if missing
        rate_col = "subsetInSchool_sev3_emis"
        if rate_col not in out.columns:
            out[rate_col] = np.nan

        denom = E.replace(0, np.nan)
        new_rate = (out[sev3N_col] / denom).clip(lower=0, upper=1)
        out.loc[mask, rate_col] = new_rate[mask].fillna(0)

    return out


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
# Function to handle numeric and percentage columns
def rounding_dataframe(df, figures_round, percentage_round):
    for col in df.columns:
        if col.startswith('#'):
            # Convert to numeric and round
            df[col] = pd.to_numeric(df[col], errors='coerce').round(figures_round)
        elif col.startswith('%'):
            # Convert to numeric, multiply by 100, and round
            df[col] = pd.to_numeric(df[col], errors='coerce').apply(lambda x: round(x * 100, percentage_round))     

########################################################################################################################################
########################################################################################################################################
##############################################    PIN CALCULATION FUNCTION    ##########################################################
########################################################################################################################################
########################################################################################################################################
def calculatePIN_with_EMIS (data_combination, country, edu_data, household_data, choice_data, survey_data, ocha_data,mismatch_ocha_data,emis_data,
                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                age_var, gender_var,
                label, 
                admin_var, vector_cycle, start_school, status_var,host_value ,idp_value ,returnee_value ,refugee_value, other_value ,
                mismatch_admin,
                selected_language):

    if "m" not in data_combination:
        raise ValueError("'m' is required in data_combination")
    

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

    #admin_var = find_best_match(admin_target,  household_data)
    admin_var = 'admin_hno'

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

    print(indicator_access_list)

    # Clean indicator columns
    indicator_access_list = clean_indicator_columns(indicator_access_list, 'indicator_access_list')
    indicator_teacher_list = clean_indicator_columns(indicator_teacher_list, 'indicator_teacher_list')
    indicator_hazard_list = clean_indicator_columns(indicator_hazard_list, 'indicator_hazard_list')
    indicator_idp_list = clean_indicator_columns(indicator_idp_list, 'indicator_idp_list')
    indicator_occupation_list = clean_indicator_columns(indicator_occupation_list, 'indicator_occupation_list')
    indicator_barrier4_list = clean_indicator_columns(indicator_barrier4_list, 'indicator_barrier4_list')
    indicator_barrier5_list = clean_indicator_columns(indicator_barrier5_list, 'indicator_barrier5_list')

    print(indicator_access_list)

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


    ################### use enrolled numbers in emis to calculate the tot number of enrolled kids per pop_grop
    ####### ** 1 **       ------------------------------ step 1: extract access rate from MSNA by pop group
    access_rate_df = None

    for category, rate_pop_df in pin_by_indicator_status_list.items():
        # Start with the base DataFrame for this category
        print(category)
        rate_pop_df = rate_pop_df.copy()

        new_col_name = f"{category} -- rate_indicator_access"
        rate_pop_df = rate_pop_df.iloc[:, [0, 2]]  # Keep admin1 and rate_indicator_access
        rate_pop_df.columns = [admin_var, new_col_name]  # Rename columns  

        # Merge with previous categories
        if access_rate_df is None:
            access_rate_df = rate_pop_df
        else:
            access_rate_df = access_rate_df.merge(rate_pop_df, on=admin_var, how="outer")

    access_rate_list={}
    for category, rate_pop_df in pin_by_indicator_status_list.items():
        # Start with the base DataFrame for this category
        print(category)
        rate_pop_df = rate_pop_df.copy()

        new_col_name = f"rate_indicator_access"
        rate_pop_df = rate_pop_df.iloc[:, [0, 2]]  # Keep admin1 and rate_indicator_access
        rate_pop_df.columns = [admin_var, new_col_name]  # Rename columns  

        access_rate_list [category] = rate_pop_df


    print(access_rate_df)
    print(access_rate_list)

    ####### ** 2 **       ------------------------------ step 2: group by admin the emis data
    emis_df = emis_data.groupby("Admin Pcode")["Enrolled students -- Children/Enfants (5-17)"].sum().reset_index()
    emis_df = emis_df.rename(columns={'Admin Pcode': admin_var})
    emis_df = emis_df.rename(columns={'Enrolled students -- Children/Enfants (5-17)': 'enrolled_emis'})

    print('--------------------------------------------------')
    print(emis_df)

    ## ----- step 3.1: organize and label ocha data
    ####### ** 3 **       ------------------------------ step 3: matching between the admin and the ocha population data
    ## finding the match between the OCHA status cathegory and the country status. 
    status_values = [status for status in edu_data[pop_group_var].unique() if status not in status_to_be_excluded]# Retrieve unique values directly without converting to lowercase
    for key, suggestions in suggestions_mapping.items():
        suggestions_mapping[key] = suggestions  # keeping original case

    mapped_statuses = map_template_to_status(template_values, suggestions_mapping, status_values)
    print (mapped_statuses)
    mapped_statuses = {
        "Host/Hôte": host_value,
        "IDP/PDI": idp_value,
        "Returnees/Retournés": returnee_value,
        'Refugees/Refugiees': refugee_value,
        "Other": other_value    }
    print (mapped_statuses)
    category_data_frames = extract_status_data(ocha_pop_data, mapped_statuses, pop_group_var)# Extract population figures based on mapped statuses without modifying the case

    for category, df in category_data_frames.items():
        df.rename(columns={'Admin': admin_var}, inplace=True)

    ocha_data_frames_for_emis = category_data_frames
    for category, df in ocha_data_frames_for_emis.items():
        df.rename(columns={'Admin': admin_var}, inplace=True)
        df.rename(columns={'TotN': f"{category} -- TotN"}, inplace=True)
        
        # Remove 'Category' and 'pop_group' columns if they exist
        df.drop(columns=['Category', pop_group_var], errors='ignore', inplace=True)

    ocha_number_df = None
    for category, ocha_pop_df in ocha_data_frames_for_emis.items():
        # Start with the base DataFrame for this category
        ocha_pop_df = ocha_pop_df.copy()
        # Merge with previous categories
        if ocha_number_df is None:
            ocha_number_df = ocha_pop_df
        else:
            ocha_number_df = ocha_number_df.merge(ocha_pop_df, on=admin_var, how="outer")


    ####### ** 4 **       ------------------------------ step 4: merge OCHA, EMiS, rate MSNA
    enrollment_df = ocha_number_df
    enrollment_df = enrollment_df.merge(emis_df, on=admin_var, how="outer")
    enrollment_df = enrollment_df.merge(access_rate_df, on=admin_var, how="outer")

    ####### ** 4.1 **       ------------------------------ step 4.1: sanity check EMIS < OCHA figures
    ocha_totn_cols = [col for col in enrollment_df.columns if " -- TotN" in col]
    enrollment_df['ocha_sum'] = enrollment_df[ocha_totn_cols].sum(axis=1)
    enrollment_df['enrolled_emis'] = enrollment_df.apply(
        lambda row: row['ocha_sum'] if row['enrolled_emis'] > row['ocha_sum'] else row['enrolled_emis'],
        axis=1
    )
    enrollment_df.drop(columns=['ocha_sum'], inplace=True)
    print(enrollment_df)

    ####### ** 5 **       ------------------------------ step 5: CALCULATION
    valid_mappings = {k: v for k, v in mapped_statuses.items() if v != 'No match found'}

    # ------ step 5.1: calculate preliminary pop group enrolled
    for label, pop_group in valid_mappings.items():
        tot_col = f"{pop_group} -- TotN"
        rate_col = f"{pop_group} -- rate_indicator_access"
        einitial_col = f"{pop_group} -- E_initial"
        # Calculate Einitial as TotN * rate_indicator_access
        enrollment_df[einitial_col] = enrollment_df[tot_col] * enrollment_df[rate_col]

    einitial_cols = [f"{pop_group} -- E_initial" for pop_group in valid_mappings.values()]

    # ------ step 5.2: extract k factor
    enrollment_df['k_factor'] = enrollment_df['enrolled_emis'] / enrollment_df[einitial_cols].sum(axis=1)

    # ------ step 5.3: re-calcualte the correct enrolled by pop group
    for pop_group in valid_mappings.values():
        einitial_col = f"{pop_group} -- E_initial"
        e_col = f"{pop_group} -- E"
        enrollment_df[e_col] = enrollment_df['k_factor'] * enrollment_df[einitial_col]

    # ------ step 5.4: check E > OCHA and cap to 100% and redistribute the rest
    enrollment_df = cap_and_redistribute(enrollment_df, valid_mappings)
    for label, pop_group in valid_mappings.items():
        tot_col = f"{pop_group} -- TotN"
        e_col = f"{pop_group} -- E"
        oos_col = f"{pop_group} -- OoS"

        # Calculate Einitial as TotN * rate_indicator_access
        enrollment_df[oos_col] = np.where(
            np.abs(enrollment_df[tot_col] - enrollment_df[e_col]) < 1,
            0,
            enrollment_df[tot_col] - enrollment_df[e_col]
        )


    print(enrollment_df)
    print(enrollment_df.columns)
    
    ####### ** 6 **       ------------------------------ step 6: assigning in-school and OoS in the correct severity category

    # ------ step 6.1: re-organize the enrollment_df by pop group
    pop_figures_E_OoS_by_pop_group = {}

    for label, pop_group in valid_mappings.items():
        # Define the columns to keep: the admin identifier plus the three columns for the current pop group.
        cols = [
            admin_var, 
            f"{pop_group} -- TotN", 
            f"{pop_group} -- E", 
            f"{pop_group} -- OoS"
        ]
        pop_group_df = enrollment_df[cols].copy()
        pop_group_df.columns = [admin_var, label_tot_population, label_tot_enrolled, label_tot_OoS]
        pop_group_df = pop_group_df.dropna(subset=[label_tot_enrolled, label_tot_OoS], how='all').reset_index(drop=True)

        # Store the resulting dataframe in the dictionary, keyed by the pop_group value.
        pop_figures_E_OoS_by_pop_group[pop_group] = pop_group_df

    # Optionally, print out the first few rows of each dataframe to verify:
    for pop_group, df in pop_figures_E_OoS_by_pop_group.items():
        print(f"Data for pop group '{pop_group}':")
        print(df, "\n")


    # ------ step 6.2: merge with pin_by_indicator_status_list
    severity_by_pop_group = {}

    for pop_group in pop_figures_E_OoS_by_pop_group:
        # Check if there is a corresponding indicator dataframe for this pop group
        if pop_group in pin_by_indicator_status_list:
            df_indicators = pin_by_indicator_status_list[pop_group]
            df_pop = pop_figures_E_OoS_by_pop_group[pop_group]
            
            # Merge using the admin variable (for example, 'admin1')
            merged_df = pd.merge(df_pop, df_indicators, on=admin_var, how="outer")
            print(merged_df.columns)
            severity_by_pop_group[pop_group] = merged_df
        else:
            print(f"Warning: No indicator dataframe found for pop group '{pop_group}'.")

    print ('===============================___________________________________________==============================')
    print(severity_by_pop_group)

    emis_ptr = emis_data

    ## ----------------------------- PTR ----------------------------------------------
    if data_combination == 'eemm' or data_combination == 'eeem':
        print(emis_ptr.columns)
        # 1) Rename first column to admin_var
        emis_ptr = emis_ptr.rename(columns={emis_ptr.columns[0]: admin_var})

        # 2) Collapse empty strings to NaN and drop all-NaN columns
        emis_ptr = emis_ptr.replace(r'^\s*$', np.nan, regex=True)
        emis_ptr = emis_ptr.dropna(axis=1, how='all')

        # 3) Rename only columns that still exist
        emis_ptr_column_rename = {
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
            if col in emis_ptr.columns
        }
        emis_ptr = emis_ptr.rename(columns=emis_ptr_column_rename)

        print(emis_ptr.columns)
        print(emis_ptr)

            # 4) Create output columns 
        emis_ptr['sev3_emis_children'] = 0
        emis_ptr['sev4_emis_children'] = 0
        emis_ptr['sev5_emis_children'] = 0

        # 5) Safe numeric coercion (only for cols that exist)
        num_cols = [c for c in [
            'sev3_indicator','sev3_th_indicator',
            'sev4_indicator_number','sev4_th_indicator_number',
            'sev5_indicator_number','sev5_th_indicator_number',
            'school_children'
        ] if c in emis_ptr.columns]
        if num_cols:
            emis_ptr[num_cols] = emis_ptr[num_cols].apply(pd.to_numeric, errors='coerce')

        # 6) Normalized categorical series (only if present)
        def _norm(s):
            return s.astype(str).str.strip().str.lower()

        sev4_cat = _norm(emis_ptr['sev4_indicator_cat']) if 'sev4_indicator_cat' in emis_ptr else None
        sev4_th  = _norm(emis_ptr['sev4_th_indicator_cat']) if 'sev4_th_indicator_cat' in emis_ptr else None
        sev5_cat = _norm(emis_ptr['sev5_indicator_cat']) if 'sev5_indicator_cat' in emis_ptr else None
        sev5_th  = _norm(emis_ptr['sev5_th_indicator_cat']) if 'sev5_th_indicator_cat' in emis_ptr else None

        # 7) Build masks with strict priority: sev5 → sev4 → sev3
        # sev5: prefer categorical equality if cat cols exist; else numeric >
        m5_cat = (sev5_cat is not None) and (sev5_th is not None) and (sev5_cat.notna() & sev5_th.notna() & (sev5_cat == sev5_th))
        m5_num = all(c in emis_ptr.columns for c in ['sev5_indicator_number','sev5_th_indicator_number']) and \
                (emis_ptr['sev5_indicator_number'].notna() & emis_ptr['sev5_th_indicator_number'].notna() &
                (emis_ptr['sev5_indicator_number'] > emis_ptr['sev5_th_indicator_number']))
        m5 = m5_cat if isinstance(m5_cat, pd.Series) else False
        m5 = m5 | (m5_num if isinstance(m5_num, pd.Series) else False)

        # sev4: only where not sev5; prefer cat, else numeric >
        m4_cat = (sev4_cat is not None) and (sev4_th is not None) and (sev4_cat.notna() & sev4_th.notna() & (sev4_cat == sev4_th))
        m4_num = all(c in emis_ptr.columns for c in ['sev4_indicator_number','sev4_th_indicator_number']) and \
                (emis_ptr['sev4_indicator_number'].notna() & emis_ptr['sev4_th_indicator_number'].notna() &
                (emis_ptr['sev4_indicator_number'] > emis_ptr['sev4_th_indicator_number']))
        m4 = m4_cat if isinstance(m4_cat, pd.Series) else False
        m4 = m4 | (m4_num if isinstance(m4_num, pd.Series) else False)
        m4 = (~m5) & (m4 if isinstance(m4, pd.Series) else False)

        # sev3: only where neither sev5 nor sev4, numeric >
        m3 = all(c in emis_ptr.columns for c in ['sev3_indicator','sev3_th_indicator']) and \
            (emis_ptr['sev3_indicator'].notna() & emis_ptr['sev3_th_indicator'].notna() &
            (emis_ptr['sev3_indicator'] > emis_ptr['sev3_th_indicator']))
        m3 = (~m5) & (~m4) & (m3 if isinstance(m3, pd.Series) else False)

        # 8) Assign children counts (keep zeros otherwise)
        if 'school_children' in emis_ptr.columns:
            emis_ptr['sev5_emis_children'] = np.where(m5, emis_ptr['school_children'], 0)
            emis_ptr['sev4_emis_children'] = np.where(m4, emis_ptr['school_children'], 0)
            emis_ptr['sev3_emis_children'] = np.where(m3, emis_ptr['school_children'], 0)
        else:
            # Optional: warn if school_children is missing
            print("Warning: 'school_children' column not present; sev*_children remain 0.")


        children_cols = ['school_children', 'sev3_emis_children', 'sev4_emis_children', 'sev5_emis_children']

        # Group by and sum
        emis_ptr_severity = emis_ptr.groupby(admin_var)[children_cols].sum().reset_index()

        # 4) Create % severity columns 
        emis_ptr_severity['subsetInSchool_sev3_emis'] =  emis_ptr_severity['sev3_emis_children']/emis_ptr_severity['school_children'] 
        emis_ptr_severity['subsetInSchool_sev4_emis'] =  emis_ptr_severity['sev4_emis_children']/emis_ptr_severity['school_children'] 
        emis_ptr_severity['subsetInSchool_sev5_emis'] =  emis_ptr_severity['sev5_emis_children']/emis_ptr_severity['school_children'] 


        # Keep only the rate columns from JENA
        emis_ptr_rates = emis_ptr_severity[[admin_var, 'subsetInSchool_sev3_emis', 'subsetInSchool_sev4_emis', 'subsetInSchool_sev5_emis']].copy()


        print('gggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj')
        print(emis_ptr_severity)
        print(emis_ptr_rates)



        for category, ocha_df in severity_by_pop_group.items():
            severity_by_pop_group[category] = (
                ocha_df.merge(emis_ptr_rates, on=admin_var, how='left')
            )






    # ------ step 6.3: calculate the ToT# for each indicator
    severity_by_pop_group = { 
        pop_group: add_figures_columns(df.copy()) 
        for pop_group, df in severity_by_pop_group.items() 
    }


    for pop_group, df in severity_by_pop_group.items():
        print(f"after add_figures_columns '{pop_group}':")
        print(df, "\n")
        print(df.columns)

    # i have to remove some columns according to the data_combination
    #if i have emmm we remove the sev3/sev4/sev5_emis
    #if i have eemm we remove sev3 teacher and hazard
    # if i have eeem  we remove sev3 teacher and hazard and sev 4 displaced and sev4 occupation
    # then we can do add_additional_severity_column which calculated the severity2 for in school as difference between the pin and total kids in school
    
    test_intermediate_step = severity_by_pop_group


    severity_by_pop_group = {
        pop_group: apply_scenario_drops(df.copy(),data_combination=data_combination)
        for pop_group, df in severity_by_pop_group.items()
    }


    severity_by_pop_group = {
        pop_group: recompute_sev3_emis_if_eemm(df.copy(), data_combination=data_combination)
        for pop_group, df in severity_by_pop_group.items()
    }


    # ------ step 6.4: calculate the severity 2 as difference between tot E and inschool-sev3 and calculate OoS sev3 as difference between OoS and sev4 and sev5
    severity_by_pop_group = {
        pop_group: add_additional_severity_columns(df.copy())
        for pop_group, df in severity_by_pop_group.items()
    }

    for pop_group, df in severity_by_pop_group.items():
        print(f"after apply_scenario_drops '{pop_group}':")
        print(df, "\n")
        print(df.columns)






    pin_by_dimension = severity_by_pop_group
    pin_by_indicator = severity_by_pop_group

    # ------ step 6.5: aggregate same severity # for OoS and in-school
    pin_by_pop_group = {
        pop_group: final_severity_columns(df.copy(), admin_var=admin_var, pop_group_var=pop_group_var)
        for pop_group, df in severity_by_pop_group.items()
    }

    # ------ step 6.6: calculate the % for each severity
    for pop_group, pop_group_df in pin_by_pop_group.items():
        pop_group_df.columns = [str(col) for col in pop_group_df.columns]

        ## Arranging columns 
        cols = list(pop_group_df.columns)
        pop_group_df = pop_group_df[cols]

        # Initialize total columns with zeros
        for label in [label_perc2, label_perc3, label_perc4, label_perc5]:
            pop_group_df[label] = 0


        cols = list(pop_group_df.columns)
        for perc, tot in [(label_perc2, label_tot2), 
                        (label_perc3, label_tot3), 
                        (label_perc4, label_tot4), 
                        (label_perc5, label_tot5)]:
            # Get the current index of the tot column.
            insert_idx = cols.index(tot)
            perc_col = cols.pop(cols.index(perc))
            cols.insert(insert_idx, perc_col)

        pop_group_df = pop_group_df[cols]     

        # Calculate total PiN for each severity level
        for perc_label, total_label in [(label_perc2, label_tot2), 
                                        (label_perc3, label_tot3), 
                                        (label_perc4, label_tot4), 
                                        (label_perc5, label_tot5)]:
            pop_group_df[perc_label] = pop_group_df[total_label] / pop_group_df[label_tot_population]

        pop_group_df[label_perc_tot] = (pop_group_df[label_tot3] +pop_group_df[label_tot4] +pop_group_df[label_tot5]) / pop_group_df[label_tot_population]


        pop_group_df[label_tot] = (pop_group_df[label_tot3] +
                        pop_group_df[label_tot4] +
                        pop_group_df[label_tot5])
        

                   # Define conditions based on specified logic
        conditions = [
            pop_group_df[label_perc5] > 0.2,
            (pop_group_df[label_perc5] + pop_group_df[label_perc4]) > 0.2,
            (pop_group_df[label_perc5] + pop_group_df[label_perc4] + pop_group_df[label_perc3]) > 0.2,
            (pop_group_df[label_perc5] + pop_group_df[label_perc4] + pop_group_df[label_perc3] + pop_group_df[label_perc2]) > 0.2
        ]
        # Corresponding values for each condition
        choices = ['5', '4', '3', '1-2']
        # Apply the conditions to determine admin severity
        pop_group_df[label_admin_severity] = np.select(conditions, choices, default='0')

        pop_group_df = pop_group_df.drop(["In-School Children", "Out-of-School Children"], axis=1)


        print('################################################################################') 
        print(pop_group_df.columns)
        print('################################################################################') 
        pin_by_pop_group[pop_group] = pop_group_df

    Tot_PiN_JIAF = pin_by_pop_group



    # ------ step 6.7: aggregate same severity # for OoS and in-school
    pin_by_dimension_pop_group = {
        pop_group: final_dimension_columns(df.copy(), admin_var=admin_var, pop_group_var = pop_group_var)
        for pop_group, df in pin_by_dimension.items()
    }
    # ------ step 6.8: PiN by dimension in need!!! for the word document 
    pin_by_dimension_in_need_pop_group = pin_by_dimension_pop_group
    # ------ step 6.9: calculate the % for each severity
    for pop_group, pop_group_df in pin_by_dimension_in_need_pop_group.items():
        pop_group_df.columns = [str(col) for col in pop_group_df.columns]

        ## Arranging columns 
        cols = list(pop_group_df.columns)
        pop_group_df = pop_group_df[cols]
        # Initialize total columns with zeros
        for label in [label_perc_acc, label_perc_agg, label_perc_lc, label_perc_penv]:
            pop_group_df[label] = 0


        cols = list(pop_group_df.columns)
        for perc, tot in [(label_perc_acc, label_tot_acc), 
                        (label_perc_agg, label_tot_agg), 
                        (label_perc_lc, label_tot_lc), 
                        (label_perc_penv, label_tot_penv)]:
            # Get the current index of the tot column.
            insert_idx = cols.index(tot)
            perc_col = cols.pop(cols.index(perc))
            cols.insert(insert_idx, perc_col)

        pop_group_df = pop_group_df[cols]     

        # Calculate total PiN for each severity level
        for perc_label, total_label in [(label_perc_acc, label_tot_acc), 
                                        (label_perc_agg, label_tot_agg), 
                                        (label_perc_lc, label_tot_lc), 
                                        (label_perc_penv, label_tot_penv)]:
            pop_group_df[perc_label] = pop_group_df[total_label] / pop_group_df[label_dimension_tot]

        pop_group_df[label_dimension_perc_tot] = 100
        pop_group_df[label_dimension_tot] = pop_group_df[label_tot_population]

        pop_group_df = pop_group_df.drop(["In-School Children", "Out-of-School Children"], axis=1)

        pin_by_dimension_in_need_pop_group[pop_group] = pop_group_df




    # ------ step 6.10: PiN by indicators
    pin_by_indicator_pop_group = {
        pop_group: final_indicator_columns(df.copy(), admin_var=admin_var, pop_group_var = pop_group_var)
        for pop_group, df in pin_by_indicator.items()
    }
    print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~') 

    # ------ step 6.11: calculate the % for each severity
    # ------ step 6.11: calculate the % for each severity
    # Ensure column names are unique before processing
    # Ensure column names are unique before processing
    for pop_group, pop_group_df in pin_by_indicator_pop_group.items():
        pop_group_df.columns = [str(col) for col in pop_group_df.columns]

        # Save original column list before deduplication
        original_columns = pop_group_df.columns.tolist()

        # Remove duplicate column names before proceeding
        pop_group_df = pop_group_df.loc[:, ~pop_group_df.columns.duplicated()]

        # Save deduplicated column list
        deduplicated_columns = pop_group_df.columns.tolist()

        # Find removed columns
        removed_columns = [col for col in original_columns if col not in deduplicated_columns]

        if removed_columns:
            print(f"⚠️ Removed duplicate columns in pop_group '{pop_group}':\n {removed_columns}\n")
        
        ## Arranging columns
        cols = list(pop_group_df.columns)
        pop_group_df = pop_group_df[cols]

        if label_tot_population not in pop_group_df.columns:
            print(f"Warning: Missing '{label_tot_population}' in pop_group '{pop_group}'. Skipping calculations.")
            continue

        # Avoid division by zero by replacing 0 with NaN
        pop_group_df.loc[:, label_tot_population] = pop_group_df[label_tot_population].replace(0, np.nan)

        predefined_tot_to_perc = {
            label_tot_sev3_indicator_access: label_perc_sev3_indicator_access,
            label_tot_sev3_indicator_teacher: label_perc_sev3_indicator_teacher,
            label_tot_sev3_indicator_hazard: label_perc_sev3_indicator_hazard,
            label_tot_sev4_indicator_idp: label_perc_sev4_indicator_idp,
            label_tot_sev5_indicator_occupation: label_perc_sev5_indicator_occupation,
            label_tot_sev4_aggravating_circumstances: label_perc_sev4_aggravating_circumstances,
            label_tot_sev5_aggravating_circumstances: label_perc_sev5_aggravating_circumstances
        }

        # **Only keep pairs where the total column exists in the DataFrame**
        valid_tot_to_perc = {
            tot_col: perc_col for tot_col, perc_col in predefined_tot_to_perc.items() if tot_col in pop_group_df.columns
        }

        # Compute % columns for valid pairs only using `.loc`
        for tot_col, perc_col in valid_tot_to_perc.items():
            if perc_col not in pop_group_df.columns:  # Prevent duplicate creation
                pop_group_df.loc[:, perc_col] = pop_group_df[tot_col] / pop_group_df[label_tot_population]

        # Dynamically generate % columns for any "ToT # children" indicators
        for col in list(pop_group_df.columns):  # Convert to list to avoid runtime errors during iteration
            match = re.match(r"(severity level \d+:) \(ToT # children\) (.+)", col)
            if match:
                severity_level, description = match.groups()
                perc_col = f"{severity_level} (% of children) {description}"  # Create new column name

                if perc_col not in pop_group_df.columns:
                    pop_group_df.loc[:, perc_col] = pop_group_df[col] / pop_group_df[label_tot_population]        

        pop_group_df = reorder_severity_columns(pop_group_df)


        # Save the cleaned DataFrame
        pin_by_indicator_pop_group[pop_group] = pop_group_df




    print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~') 


    for pop_group, df in pin_by_dimension_pop_group.items():
        print(f"pin by dimension 2 '{pop_group}':")
        print(df, "\n")
        print(df.columns)
        
    for pop_group, df in pin_by_indicator_pop_group.items():
        print(f"pin by indicartor 2 '{pop_group}':")
        print(df, "\n")
        print(df.columns)




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


    




    return pin_by_indicator_status_list, enrollment_df, pop_figures_E_OoS_by_pop_group, severity_by_pop_group, pin_by_pop_group, pin_by_dimension_in_need_pop_group,pin_by_indicator_pop_group, test_intermediate_step, Tot_PiN_JIAF, final_overview_df_OCHA, final_overview_df, Tot_PiN_by_admin