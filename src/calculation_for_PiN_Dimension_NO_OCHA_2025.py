import pandas as pd
#import fuzzywuzzy
from fuzzywuzzy import process, fuzz
import numpy as np
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
import re

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
def calculate_age_correction(start_month, collection_month):
    # Create a dictionary to map the first three letters of month names to their numeric equivalents
    month_lookup = {datetime.date(2000, i, 1).strftime('%b').lower(): i for i in range(1, 13)}
    
    # Convert month names to their numeric equivalents using the predefined lookup
    start_month_num = month_lookup[start_month.strip()[:3].lower()]
    
    # Adjust the start month number for a school year starting in the previous calendar year
    adjusted_start_month_num = start_month_num - 12 if start_month_num > 6 else start_month_num
    
    # Determine if the age correction should be applied based on the month difference
    age_correction = (collection_month - adjusted_start_month_num) > 6
    return age_correction

##--------------------------------------------------------------------------------------------
def find_matching_choices(choices_df, barriers_list, label_var):
    # List to hold the results
    results = []
    
    # Iterate over each barrier in the list
    for barrier in barriers_list:
        # Filter choices where 'label::english' matches the current barrier
        matched_choices = choices_df[choices_df[label_var] == barrier]
        
        # For each matched choice, create an entry in the results list
        for _, choice in matched_choices.iterrows():
            result_entry = {'name': choice['name'], 'label': barrier}
            results.append(result_entry)
    
    return results
##--------------------------------------------------------------------------------------------
def assign_school_cycle(edu_age_corrected, single_cycle=False, lower_primary_start_var=6, lower_primary_end_var=13, upper_primary_end_var=None):
    if single_cycle:
        # If single cycle is True, handle as a primary to secondary without upper primary
        if lower_primary_start_var <= edu_age_corrected <= lower_primary_end_var:
            return 'primary'
        elif lower_primary_end_var + 1 <= edu_age_corrected <= 18:
            return 'secondary'
        elif edu_age_corrected == 5: 
            return 'ECE'
        else:
            return 'out of school range'
    else:
        # If single cycle is False, handle lower primary, upper primary, and secondary phases
        if lower_primary_start_var <= edu_age_corrected <= lower_primary_end_var:
            return 'lower primary'
        elif upper_primary_end_var and lower_primary_end_var + 1 <= edu_age_corrected <= upper_primary_end_var:
            return 'intermediate level'
        elif upper_primary_end_var and upper_primary_end_var + 1 <= edu_age_corrected <= 18:
            return 'secondary'
        elif edu_age_corrected == 5: 
            return 'ECE'
        else:
            return 'out of school range'
        

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
def extract_subtables(df, pop_group_var):
    # Ensure the DataFrame has a MultiIndex and includes the population group variable
    if not isinstance(df.index, pd.MultiIndex) or pop_group_var not in df.index.names:
        raise ValueError("DataFrame must have a MultiIndex and include the specified population group variable.")

    # Get unique population groups from the specified level of the index
    unique_groups = df.index.get_level_values(pop_group_var).unique()

    # Dictionary to store each sub-DataFrame
    subtables_dict = {}

    # Extract subtables for each unique group
    for group in unique_groups:
        # Extract data for the current group
        sub_df = df.xs(group, level=pop_group_var)
        
        # Reset the index to turn MultiIndex into regular columns
        sub_df = sub_df.reset_index()
        
        # Set new DataFrame with simplified headers
        subtables_dict[group] = sub_df.rename(columns=lambda x: x if isinstance(x, str) else str(x))

    return subtables_dict

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
    columns_to_adjust = [col for col in factorized_df.columns if col.startswith('#') or col == tot_column]
    del factorized_df['Admin']

    # Apply the multiplication for each column that needs adjustment
    for col in columns_to_adjust:
        factorized_df[col] *= factorized_df[category_label]

    # Drop the now unneeded factor column
    factorized_df.drop(columns=[category_label], inplace=True)
    return factorized_df



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
# preparation for overview--> SUM all the admin per population group and per strata 
def collapse_and_summarize_dimension(pin_per_admin_status_strata, category_str, admin_var):
    collapsed_results = {}
    
    # Iterate over the input dictionary
    for category, df in pin_per_admin_status_strata.items():
        # Create a copy of the first row to preserve the structure
        summed_df = df.iloc[0:1].copy()

        # Identify columns to skip from summation and columns to set to zero
        columns_to_skip = [col for col in df.columns if col.startswith('%') or col == admin_var  or col == 'Population group' or col == 'Category' or col == 'Area severity' or col == 'Strata_category']
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
        if 'Strata_category' in summed_df.columns:
            del summed_df['Strata_category']

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

    overview_strata[label_dimension_perc_tot] = 0
    overview_strata[label_dimension_tot] = (overview_strata[label_tot_acc] +
                               overview_strata[label_tot_agg] +
                               overview_strata[label_tot_penv] + overview_strata[label_tot_lc])

    cols = list(overview_strata.columns)
    cols.insert(cols.index(label_dimension_tot) + 1, cols.pop(cols.index('Category')))
    overview_strata = overview_strata[cols]

    return overview_strata
##--------------------------------------------------------------------------------------------
def aggregate_pin_per_admin_status(pin_per_admin_status, admin_var):
    # Concatenate all DataFrames from the pin_per_admin_status dictionary into a single DataFrame
    combined_df = pd.concat(pin_per_admin_status.values(), ignore_index=True)
    columns_to_zero = [col for col in combined_df.columns if col.startswith('%')]
    for col in columns_to_zero:
        combined_df[col] = 0

    # Group by 'admin_var'  summing the numeric columns
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
    
    # After summing, calculate the total PiN (3+) across severity levels
    #grouped_df[label_tot] = grouped_df[label_tot3] + grouped_df[label_tot4] + grouped_df[label_tot5]
    #grouped_df[label_perc_tot] = 0

    return grouped_df

##--------------------------------------------------------------------------------------------        
def merge_pin_ocha_with_strata (strata_pin_per_admin_status, category_data_frames,factor_strata_df, severity_strata_list, admin_var, pop_group_var,strata_label = 'Girl'):

    # Assume category_data_frames is a dictionary of DataFrames, indexed by category
    for category, df in category_data_frames.items():
        # Ensure both DataFrames are ready to merge
        if category in severity_strata_list:
            # Fetch the corresponding DataFrame from the grouped data
            grouped_df = severity_strata_list[category] 
            factorized_df = pd.merge(
                df, factor_strata_df, 
                left_on=[admin_var], 
                right_on=["Admin"], 
                how='left'
            )

            factorized_df['TotN'] *= factorized_df[strata_label]
            del factorized_df[strata_label]
            del factorized_df['Admin']

            pop_group_df = pd.merge(grouped_df, factorized_df, on=[admin_var, pop_group_var])
            pop_group_df.columns = [str(col) for col in pop_group_df.columns]

            ## Arranging columns 
            cols = list(pop_group_df.columns)
            pop_group_df = pop_group_df[cols]

            ## Calculation of the total PiN and admin severity
            pop_group_df = pop_group_df.rename(columns={
                pop_group_var: 'Population group',
                int_2: label_perc2,
                int_3: label_perc3,
                int_4: label_perc4,
                int_5: label_perc5
            })

            # Initialize total columns with zeros
            for label in [label_tot2, label_tot3, label_tot4, label_tot5]:
                pop_group_df[label] = 0

            cols = list(pop_group_df.columns)
            # Move the newly added column to the desired position
            cols.insert(cols.index(label_perc2) + 1, cols.pop(cols.index(label_tot2)))
            cols.insert(cols.index(label_perc3) + 1, cols.pop(cols.index(label_tot3)))
            cols.insert(cols.index(label_perc4) + 1, cols.pop(cols.index(label_tot4)))
            cols.insert(cols.index(label_perc5) + 1, cols.pop(cols.index(label_tot5)))
            pop_group_df = pop_group_df[cols]     

            # Calculate total PiN for each severity level
            for perc_label, total_label in [(label_perc2, label_tot2), 
                                            (label_perc3, label_tot3), 
                                            (label_perc4, label_tot4), 
                                            (label_perc5, label_tot5)]:
                pop_group_df[total_label] = pop_group_df[perc_label] * pop_group_df[label_tot_population]

            # Reorder columns as needed
            cols = list(pop_group_df.columns)
            cols.insert(cols.index('Population group') + 1, cols.pop(cols.index(label_tot_population)))
            pop_group_df = pop_group_df[cols]

            cols.remove(label_tot_population)
            cols.insert(cols.index('Population group') + 1, label_tot_population)
            pop_group_df = pop_group_df[cols]

            # Save modified DataFrame back into the dictionary under the category key
            strata_pin_per_admin_status[category] = pop_group_df

    return strata_pin_per_admin_status

##--------------------------------------------------------------------------------------------
def merge_dimension_ocha_with_strata (strata_dimension_per_admin_status, category_data_frames,factor_strata_df, dimension_strata_list, admin_var, pop_group_var,strata_label = 'Girl'):

    # Assume category_data_frames is a dictionary of DataFrames, indexed by category
    for category, df in category_data_frames.items():
        # Ensure both DataFrames are ready to merge
        if category in dimension_strata_list:
            # Fetch the corresponding DataFrame from the grouped data
            grouped_df = dimension_strata_list[category] 

            factorized_df = pd.merge(
                df, factor_strata_df, 
                left_on=[admin_var], 
                right_on=["Admin"], 
                how='left'
            )

            factorized_df['TotN'] *= factorized_df[strata_label]
            del factorized_df[strata_label]
            del factorized_df['Admin']
            # Merge on specified columns
            pop_group_df = pd.merge(grouped_df, factorized_df, on=[admin_var, pop_group_var])
            pop_group_df.columns = [str(col) for col in pop_group_df.columns]

            ## arranging columns 
            cols = list(pop_group_df.columns)
            pop_group_df = pop_group_df[cols]


            ## !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!   calculation of the tot Pin and admin severity -->
            # Step 1: Create the new column with initial zeros
            pop_group_df = pop_group_df.rename(columns={
                pop_group_var: 'Population group',
                int_acc: label_perc_acc,
                int_agg: label_perc_agg,
                int_lc: label_perc_lc,
                int_penv: label_perc_penv,
                int_out: label_perc_out
            })
            del pop_group_df['Category']

            # Initialize total columns with zeros
            for label in [label_tot_acc, label_tot_agg, label_tot_lc, label_tot_penv, label_tot_out]:
                pop_group_df[label] = 0


            cols = list(pop_group_df.columns)
            cols.remove(label_perc_out)
            cols.insert( cols.index(label_perc_penv) + 1, label_perc_out)
            pop_group_df = pop_group_df[cols]

        
            cols = list(pop_group_df.columns)
            # Move the newly added column to the desired position
            
            cols.insert(cols.index(label_perc_acc) + 1, cols.pop(cols.index(label_tot_acc)))
            cols.insert(cols.index(label_perc_agg) + 1, cols.pop(cols.index(label_tot_agg)))
            cols.insert(cols.index(label_perc_lc) + 1, cols.pop(cols.index(label_tot_lc)))
            cols.insert(cols.index(label_perc_penv) + 1, cols.pop(cols.index(label_tot_penv)))
            cols.insert(cols.index(label_perc_out) + 1, cols.pop(cols.index(label_tot_out)))
            pop_group_df = pop_group_df[cols]


            for perc_label, total_label in [(label_perc_acc, label_tot_acc), 
                                            (label_perc_agg, label_tot_agg), 
                                            (label_perc_lc, label_tot_lc), 
                                            (label_perc_penv, label_tot_penv),
                                            (label_perc_out, label_tot_out)]:
                pop_group_df[total_label] = pop_group_df[perc_label] * pop_group_df[label_dimension_tot_population]
            # Save modified DataFrame back into the dictionary under the category key

                    
            # Reorder columns as needed
            cols = list(pop_group_df.columns)
            cols.insert(cols.index('Population group') + 1, cols.pop(cols.index(label_dimension_tot_population)))
            pop_group_df = pop_group_df[cols]     


            cols.remove(label_dimension_tot_population)
            cols.insert( cols.index('Population group') + 1, label_dimension_tot_population)
            pop_group_df = pop_group_df[cols]
            strata_dimension_per_admin_status[category] = pop_group_df
    return strata_dimension_per_admin_status


##--------------------------------------------------------------------------------------------
def ensure_columns(pin_list, needed_columns):
    for category, grouped_df in pin_list.items():
        # Check if columns are missing and add them
        missing_columns = [col for col in needed_columns if col not in grouped_df.columns]
        #print(grouped_df.columns)
        #print(missing_columns)
        # Add missing columns only, without duplicating any existing ones
        for column in missing_columns:
            grouped_df[column] = 0  # Add the missing column with a default value of zero
        
        # Update the DataFrame in the dictionary
        pin_list[category] = grouped_df

    return pin_list
##--------------------------------------------------------------------------------------------
def custom_to_datetime(date_str):
    try:
        return pd.to_datetime(date_str, errors='coerce')
    except:
        try:
            return pd.to_datetime(date_str, format='%Y-%m-%d %H:%M:%S.%f', errors='coerce')
        except:
            return pd.NaT
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
########################################################################################################################################
########################################################################################################################################
##############################################    PIN CALCULATION FUNCTION    ##########################################################
########################################################################################################################################
########################################################################################################################################
def calculatePIN_NO_OCHA_2025 (country, edu_data, household_data, choice_data, survey_data, 
                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                age_var, gender_var,
                label, 
                admin_var, vector_cycle, start_school, status_var):

    admin_target = admin_var
    pop_group_var = status_var




    ## essential variables --------------------------------------------------------------------------------------------
    single_cycle = (vector_cycle[1] == 0)
    primary_start = 6
    secondary_end = 17

    host_suggestion = ["always_lived",'Host Community','host_communi', "always_lived","non_displaced_vulnerable",'host',"non_pdi","hote","menage_n_deplace","menage_n_deplace","resident","lebanese","Populationnondéplacée","ocap","non_deplacee","Residents","yes","4"]
    IDP_suggestion = ["displaced", 'New IDPs','pdi', 'idp', 'site', 'camp', 'migrant', 'Out-of-camp', 'In-camp','no', 'pdi_site', 'pdi_fam', '2', '1' ]
    returnee_suggestion = ['displaced_previously' ,'cb_returnee','ret','Returnee HH','returnee' ,'ukrainian moldovan','Returnees','5']
    refugee_suggestion = ['refugees', 'refugee', 'prl', 'refugiee', '3']
    ndsp_suggestion = ['ndsp','Protracted IDPs']
    status_to_be_excluded = ['dnk', 'other', 'pnta', 'dont_know', 'no_answer', 'prefer_not_to_answer', 'pnpr', 'nsp', 'autre', 'do_not_know', 'decline']
    template_values = ['Host/Hôte',	'IDP/PDI',	'Returnees/Retournés', 'Refugees/Refugiee', 'Other']
    suggestions_mapping = {
        'Host/Hôte': host_suggestion,
        'IDP/PDI': IDP_suggestion,
        'Returnees/Retournés': returnee_suggestion,
        'Refugees/Refugiee': refugee_suggestion,
        'Other': ndsp_suggestion
    }
    # --------------------------------------------------------------------------------------------


    ####### ** 1 **       ------------------------------ manipulation and join between H and edu data   ------------------------------------------     #######
    ####### ** 2 **       ------------------------------ severity definition and calculation ------------------------------------------     #######
    # in the function add_severity

    ####### ** 3 **       ------------------------------ Analysis per ADMIN AND POPULATION GROUP ------------------------------------------     #######
    admin_var = find_best_match(admin_target,  household_data.columns)
    edu_data = edu_data[edu_data[access_var].notna()]
    edu_data = edu_data[edu_data['severity_category'].notna()]


    df = pd.DataFrame(edu_data)
    # Filtering data based on gender
    female_df = edu_data[edu_data[gender_var].isin(['female', 'femme', 'woman_girl'])]
    male_df = edu_data[edu_data[gender_var].isin(['male', 'homme', 'man_boy'])]
    # Filtering data based on school cycle
    ece_df = edu_data[edu_data['school_cycle'].isin(['ECE'])]
    primary_df = edu_data[edu_data['school_cycle'].isin(['primary'])]
    secondary_df = edu_data[edu_data['school_cycle'].isin(['secondary'])]
    if not single_cycle:
        intermediate_df = edu_data[edu_data['school_cycle'].isin(['intermediate level'])]
    # filtering only kids in need == 3+
    in_need_df = edu_data[edu_data['severity_category'].isin([3, 4, 5])]
    #------    CORRECT PIN    -------            
    severity_admin_status = df.groupby([admin_var, pop_group_var, 'severity_category']).agg(
        total_weight=('weights', 'sum')
    ).groupby(level=[0, 1]).apply(
        lambda x: x / x.sum()
    ).unstack(fill_value=0)
    #-------    CORRECT TARGETTING    -------          
    dimension_admin_status = df.groupby([admin_var, pop_group_var, 'dimension_pin']).agg(
        total_weight=('weights', 'sum')
    ).groupby(level=[0, 1]).apply(
        lambda x: x / x.sum()
    ).unstack(fill_value=0)
    ## subset in need
    dimension_admin_status_in_need = in_need_df.groupby([admin_var, pop_group_var, 'dimension_pin']).agg(
        total_weight=('weights', 'sum')
    ).groupby(level=[0, 1]).apply(
        lambda x: x / x.sum()
    ).unstack(fill_value=0)
    # -------- GENDER DISAGGREGATION  ---------    
    severity_female = female_df.groupby([admin_var, pop_group_var, 'severity_category']).agg(
        total_weight=('weights', 'sum')
    ).groupby(level=[0, 1]).apply(
        lambda x: x / x.sum()
    ).unstack(fill_value=0)
    severity_male = male_df.groupby([admin_var, pop_group_var, 'severity_category']).agg(
        total_weight=('weights', 'sum')
    ).groupby(level=[0, 1]).apply(
        lambda x: x / x.sum()
    ).unstack(fill_value=0)
    dimension_female = female_df.groupby([admin_var, pop_group_var, 'dimension_pin']).agg(
        total_weight=('weights', 'sum')
    ).groupby(level=[0, 1]).apply(
        lambda x: x / x.sum()
    ).unstack(fill_value=0)
    dimension_male = male_df.groupby([admin_var, pop_group_var, 'dimension_pin']).agg(
        total_weight=('weights', 'sum')
    ).groupby(level=[0, 1]).apply(
        lambda x: x / x.sum()
    ).unstack(fill_value=0)
    # -------- SCHOOL-CYCLE DISAGGREGATION  ---------    
    dimension_ece = ece_df.groupby([admin_var, pop_group_var, 'dimension_pin']).agg(
        total_weight=('weights', 'sum')
    ).groupby(level=[0, 1]).apply(
        lambda x: x / x.sum()
    ).unstack(fill_value=0)
    dimension_primary = primary_df.groupby([admin_var, pop_group_var, 'dimension_pin']).agg(
        total_weight=('weights', 'sum')
    ).groupby(level=[0, 1]).apply(
        lambda x: x / x.sum()
    ).unstack(fill_value=0)
    dimension_secondary = secondary_df.groupby([admin_var, pop_group_var, 'dimension_pin']).agg(
        total_weight=('weights', 'sum')
    ).groupby(level=[0, 1]).apply(
        lambda x: x / x.sum()
    ).unstack(fill_value=0)
    if not single_cycle:
        dimension_intermediate = intermediate_df.groupby([admin_var, pop_group_var, 'dimension_pin']).agg(
        total_weight=('weights', 'sum')
        ).groupby(level=[0, 1]).apply(
            lambda x: x / x.sum()
        ).unstack(fill_value=0)



    ## reducing the multiindex of the panda dataframe
    severity_admin_status_list = reduce_index(severity_admin_status, 0, pop_group_var)
    dimension_admin_status_list = reduce_index(dimension_admin_status, 0, pop_group_var)
    dimension_admin_status_in_need_list = reduce_index(dimension_admin_status_in_need,  0, pop_group_var) ## only who is in need we check the distriburion of need

    severity_female_list = reduce_index(severity_female, 0, pop_group_var)
    severity_male_list = reduce_index(severity_male, 0, pop_group_var)
    dimension_female_list = reduce_index(dimension_female, 0, pop_group_var)
    dimension_male_list = reduce_index(dimension_male, 0, pop_group_var)
    dimension_ece_list = reduce_index(dimension_ece, 0, pop_group_var)
    dimension_primary_list = reduce_index(dimension_primary, 0, pop_group_var)
    dimension_secondary_list = reduce_index(dimension_secondary, 0, pop_group_var)
    if not single_cycle: dimension_intermediate_list = reduce_index(dimension_intermediate, 0, pop_group_var)


    severity_needed_columns = [2.0, 3.0, 4.0, 5.0]
    dimension_needed_columns = ['access','aggravating circumstances', 'learning condition', 'protected environment']
    severity_admin_status_list = ensure_columns(severity_admin_status_list, severity_needed_columns)
    severity_female_list = ensure_columns(severity_female_list, severity_needed_columns)
    severity_male_list = ensure_columns(severity_male_list, severity_needed_columns)

    dimension_admin_status_list = ensure_columns(dimension_admin_status_list, dimension_needed_columns)
    dimension_admin_status_in_need_list = ensure_columns(dimension_admin_status_in_need_list, dimension_needed_columns)
    dimension_female_list = ensure_columns(dimension_female_list, dimension_needed_columns)
    dimension_male_list = ensure_columns(dimension_male_list, dimension_needed_columns)
    dimension_ece_list = ensure_columns(dimension_ece_list, dimension_needed_columns)
    dimension_primary_list = ensure_columns(dimension_primary_list, dimension_needed_columns)
    dimension_secondary_list = ensure_columns(dimension_secondary_list, dimension_needed_columns)
    if not single_cycle:    dimension_intermediate_list = ensure_columns(dimension_intermediate_list, dimension_needed_columns)

    pin_per_admin_status = {}
    for category, df in severity_admin_status_list.items():
        # Fetch the corresponding DataFrame from the grouped data
        pop_group_df = df
        pop_group_df = pop_group_df.rename(columns={
            pop_group_var: 'Population group',
            2.0: label_perc2,
            3.0: label_perc3,
            4.0: label_perc4,
            5.0: label_perc5
        })

        # Update the dictionary with the renamed DataFrame
        pin_per_admin_status[category] = pop_group_df


    country_label = country.replace(" ", "_").replace("--", "_").replace("/", "_")







    return pin_per_admin_status, dimension_admin_status_list, severity_female_list, severity_male_list ,country_label