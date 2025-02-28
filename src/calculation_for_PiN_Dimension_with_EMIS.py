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

##--------------------------------------------------------------------------------------------
## finding admin        
def extract_number(s):
    match = re.search(r'\d+', s)
    return int(match.group()) if match else None
def find_similar_columns(admin_target, columns):
    # Extract the base target without numbers for string comparison
    base_target = re.sub(r'\d+', '', admin_target).lower()

    # Find columns that have high string similarity with the base target
    similar_columns = []
    for col in columns:
        base_col = re.sub(r'\d+', '', col).lower()
        similarity_score = fuzz.partial_ratio(base_target, base_col)
        if similarity_score > 70:  # Set a threshold for similarity
            similar_columns.append(col)
    
    return similar_columns
def find_best_match(admin_target, columns):
    # Extract the target number
    target_number = extract_number(admin_target)

    # Step 1: Find columns similar in text content
    similar_columns = find_similar_columns(admin_target, columns)

    if not similar_columns:
        # Fallback to fuzzy matching across all columns if no similar columns are found
        return process.extractOne(admin_target, columns)[0]

    # Step 2: Among similar columns, prioritize those with matching numbers
    candidates_with_same_number = [col for col in similar_columns if extract_number(col) == target_number]

    if candidates_with_same_number:
        # Further prioritize candidates that include the word 'code'
        candidates_with_code = [col for col in candidates_with_same_number if 'code' in col.lower()]

        if candidates_with_code:
            # If there are candidates with 'code', return the best match among them
            return process.extractOne(admin_target, candidates_with_code)[0]
        else:
            # If no candidates with 'code', return the best match among all candidates with the same number
            return process.extractOne(admin_target, candidates_with_same_number)[0]
    else:
        # Fallback to fuzzy matching among the similar columns
        return process.extractOne(admin_target, similar_columns)[0]
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
    household_columns = household_data.columns

    # Find the best match for `admin_var` in `household_data`
    best_match_for_admin_var = find_similar_columns(admin_var, household_columns)
    print(f"Best match for admin_var ({admin_var}) is: {best_match_for_admin_var[0]}")

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

    # Prioritize columns based on the number of non-empty values
    def prioritize_non_empty_columns(columns):
        non_empty_counts = {col: edu_data[col].notna().sum() for col in columns}
        sorted_columns = sorted(non_empty_counts, key=non_empty_counts.get, reverse=True)
        return sorted_columns[0] if sorted_columns else None

    # Handle the case where multiple levels (lengths) are detected
    if len(length_dict) > 1:
        print("Multiple levels detected:")
        for length, codes in length_dict.items():
            print(f"Level {length}: {codes}")

        # Match columns based on length and prioritize based on the number of non-empty values
        best_columns = {}
        for length, columns in admin_columns_representative.items():
            # Prioritize based on the number of non-empty values
            best_columns_for_level = prioritize_non_empty_columns(columns)
            best_columns[length] = best_columns_for_level

        admin_columns_representative = best_columns
    else:
        # For single level case, directly prioritize the column with non-empty values
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
                'indicator_teacher': 'subsetInSchool_indicator_teacher',
                'indicator_hazard': 'subsetInSchool_indicator_hazard',
                'indicator_idp': 'subsetInSchool_indicator_idp',
                'indicator_occupation': 'subsetInSchool_indicator_occupation',
                'indicator_barrier4': 'subsetOoS_aggravating_circumstances',
                'indicator_barrier5': 'subsetOoS_aggravating_circumstances'
            }.items() if col in essential_columns
        }

        # Severity 4 and 5 column renaming
        severity_4_rename = {entry['name']: f"subsetOoS: {entry['label']}" for entry in severity_4_matches}
        severity_5_rename = {entry['name']: f"subsetOoS: {entry['label']}" for entry in severity_5_matches}

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

    admin_var = find_best_match(admin_target,  household_data.columns)

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


    print(access_rate_df)

    ####### ** 2 **       ------------------------------ step 2: group by admin the emis data
    emis_df = emis_data.groupby("Admin Pcode")["Enrolled students -- Children/Enfants (5-17)"].sum().reset_index()
    emis_df = emis_df.rename(columns={'Admin Pcode': admin_var})
    emis_df = emis_df.rename(columns={'Enrolled students -- Children/Enfants (5-17)': 'enrolled_emis'})

    print(emis_df)

    ## ----- step 3.1: organize and label ocha data
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

    print(enrollment_df)

    ####### ** 5 **       ------------------------------ step 5: CALCULATION
    valid_mappings = {k: v for k, v in mapped_statuses.items() if v != 'No match found'}

    for label, pop_group in valid_mappings.items():
        tot_col = f"{pop_group} -- TotN"
        rate_col = f"{pop_group} -- rate_indicator_access"
        einitial_col = f"{pop_group} -- E_initial"
        # Calculate Einitial as TotN * rate_indicator_access
        enrollment_df[einitial_col] = enrollment_df[tot_col] * enrollment_df[rate_col]

    einitial_cols = [f"{pop_group} -- E_initial" for pop_group in valid_mappings.values()]
    enrollment_df['k_factor'] = enrollment_df['enrolled_emis'] / enrollment_df[einitial_cols].sum(axis=1)

    for pop_group in valid_mappings.values():
        einitial_col = f"{pop_group} -- E_initial"
        e_col = f"{pop_group} -- E"
        enrollment_df[e_col] = enrollment_df['k_factor'] * enrollment_df[einitial_col]



    # To check the result:
    print(enrollment_df)



    enrollment_df = cap_and_redistribute(enrollment_df, valid_mappings)

    for label, pop_group in valid_mappings.items():
        tot_col = f"{pop_group} -- TotN"
        e_col = f"{pop_group} -- E"
        oos_col = f"{pop_group} -- OoS"

        # Calculate Einitial as TotN * rate_indicator_access
        enrollment_df[oos_col] = enrollment_df[tot_col] - enrollment_df[e_col]

    print(enrollment_df)
    
    
    return pin_by_indicator_status_list, enrollment_df