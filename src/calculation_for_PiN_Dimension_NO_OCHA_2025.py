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


label_perc_sev3_indicator_access= 'severity level 3, indicator Access to education -- % of children'
label_perc_sev3_indicator_teacher = 'severity level 3, indicator: Education was disrupted by teacher absence -- % of children'
label_perc_sev3_indicator_hazard = 'severity level 3, indicator: Education was disrupted by natural hazard -- % of children'
label_perc_sev4_indicator_idp = 'severity level 4, indicator: Education was disrupted by the school being used as shelter -- % of children'
label_perc_sev5_indicator_occupation = 'severity level 5, indicator: Education was disrupted by school being occupied by armed groups -- % of children'
label_perc_sev4_aggravating_circumstances = 'severity level 4, indicator: individual aggravating circumstances (cumulative of all Level 4 aggravating circumstances) -- % of children'
label_perc_sev5_aggravating_circumstances = 'severity level 5, indicator: individual aggravating circumstances (cumulative of all Level 5 aggravating circumstances) -- % of children'


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
        # Add missing columns only, without duplicating any existing ones
        for column in missing_columns:
            grouped_df[column] = 0  # Add the missing column with a default value of zero
        
        # Update the DataFrame in the dictionary
        pin_list[category] = grouped_df

    return pin_list
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
def calculate_prop(df, admin_var, pop_group_var, target_var, agg_var='weights'):

    df_results = df.groupby([admin_var, pop_group_var, target_var]).agg(
            total_weight=(agg_var, 'sum')
        ).groupby(level=[0, 1]).apply(
            lambda x: x / x.sum()
        ).unstack(fill_value=0)

    return df_results

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
                'indicator_access': 'sev3_indicator_access',
                'indicator_teacher': 'sev3_indicator_teacher',
                'indicator_hazard': 'sev3_indicator_hazard',
                'indicator_idp': 'sev4_indicator_idp',
                'indicator_occupation': 'sev5_indicator_occupation',
                'indicator_barrier4': 'sev4_aggravating_circumstances',
                'indicator_barrier5': 'sev5_aggravating_circumstances'
            }.items() if col in essential_columns
        }

        # Severity 4 and 5 column renaming
        severity_4_rename = {entry['name']: f"severity level 4, aggravating circumnstance: {entry['label']}" for entry in severity_4_matches}
        severity_5_rename = {entry['name']: f"severity level 5, aggravating circumnstance: {entry['label']}" for entry in severity_5_matches}

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

# Function to extract the severity level
def extract_severity_level(col):
    match = re.search(r'severity level (\d+)', col)
    return int(match.group(1)) if match else None
########################################################################################################################################
########################################################################################################################################
##############################################    PIN CALCULATION FUNCTION    ##########################################################
########################################################################################################################################
########################################################################################################################################
def calculatePIN_NO_OCHA_2025 (country, edu_data, household_data, choice_data, survey_data, mismatch_ocha_data,
                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                    age_var, gender_var,
                                                                                    label, 
                                                                                    admin_var, vector_cycle, start_school, status_var,
                                                                                    mismatch_admin,
                                                                                    selected_language):

    admin_target = admin_var
    pop_group_var = status_var

    admin_var = find_best_match(admin_target,  household_data.columns)

    admin_column_rapresentative = []
    grouped_dict = {}
    if mismatch_admin:
        ocha_mismatch_list = mismatch_ocha_data
        # Create a defaultdict to store grouped data
        detailed_list = ocha_mismatch_list.iloc[:, 1].astype(str).tolist()  # Converting to string
        prefix_list = ocha_mismatch_list.iloc[:, 2].dropna().astype(str).tolist()  # Drop NaN and convert to string
        admin_low_ok_list = ocha_mismatch_list.iloc[:, 0].dropna().astype(str).tolist()  # Drop NaN and convert to string

        #print(detailed_list)
        #print(prefix_list)

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
        #print("Codes grouped by length:")
        #for length, codes in length_dict.items():
            #print(f"Length {length}: {codes}")



        admin_column_rapresentative = find_matching_columns_for_admin_levels(edu_data, household_data, prefix_list, admin_var)
        print('admin_column_rapresentative')
        print(admin_column_rapresentative)

    ## essential variables --------------------------------------------------------------------------------------------
    single_cycle = (vector_cycle[1] == 0)
    primary_start = 6
    if country == 'Afghanistan -- AFG': 
        primary_start = 7

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
    edu_data = edu_data[edu_data[access_var].notna()]
    edu_data = edu_data[edu_data['severity_category'].notna()]


    df = pd.DataFrame(edu_data)
    # Filtering data based on gender
    female_df = edu_data[edu_data[gender_var].isin(['female', 'femme', 'woman_girl', 'feminin'])]
    male_df = edu_data[edu_data[gender_var].isin(['male', 'homme', 'man_boy', 'masculin'])]
    # Filtering data based on school cycle
    ece_df = edu_data[edu_data['school_cycle'].isin(['ECE'])]
    primary_df = edu_data[edu_data['school_cycle'].isin(['primary'])]
    secondary_df = edu_data[edu_data['school_cycle'].isin(['secondary'])]
    if not single_cycle:
        intermediate_df = edu_data[edu_data['school_cycle'].isin(['intermediate level'])]
    # filtering only kids in need == 3+
    in_need_df = edu_data[edu_data['severity_category'].isin([3, 4, 5])]


    analysis_config = {
        'severity_category': {'df': df, 'target_var': 'severity_category'},
        'dimension_pin': {'df': df, 'target_var': 'dimension_pin'},
        'dimension_pin_in_need': {'df': in_need_df, 'target_var': 'dimension_pin'},
        'severity_female': {'df': female_df, 'target_var': 'severity_category'},
        'severity_male': {'df': male_df, 'target_var': 'severity_category'},
        'dimension_female': {'df': female_df, 'target_var': 'dimension_pin'},
        'dimension_male': {'df': male_df, 'target_var': 'dimension_pin'},
        'dimension_ece': {'df': ece_df, 'target_var': 'dimension_pin'},
        'dimension_primary': {'df': primary_df, 'target_var': 'dimension_pin'},
        'dimension_secondary': {'df': secondary_df, 'target_var': 'dimension_pin'},
        'indicator.access': {'df': df, 'target_var': 'indicator.access'},
        'indicator.teacher': {'df': df, 'target_var': 'indicator.teacher'},
        'indicator.hazard': {'df': df, 'target_var': 'indicator.hazard'},
        'indicator.idp': {'df': df, 'target_var': 'indicator.idp'},
        'indicator.occupation': {'df': df, 'target_var': 'indicator.occupation'},
        'indicator.barrier4': {'df': df, 'target_var': 'indicator.barrier4'},
        'indicator.barrier5': {'df': df, 'target_var': 'indicator.barrier5'},
        barrier_var: {'df': df, 'target_var': barrier_var}
    }

    if not single_cycle:
        analysis_config['dimension_intermediate'] = {'df': intermediate_df, 'target_var': 'dimension_pin'}

    results_dict = {} 

    if mismatch_admin:
        detailed_list = ocha_mismatch_list.iloc[:, 1].astype(str).tolist()  # Converting to string
        admin_up_msna = ocha_mismatch_list.iloc[:, 2].dropna().astype(str).tolist()  # Drop NaN and convert to string
        admin_low_ok_list = ocha_mismatch_list.iloc[:, 0].dropna().astype(str).tolist()  # Drop NaN and convert to string

        #print(admin_up_msna)

        for analysis_var, config in analysis_config.items():
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
        for analysis_var, config in analysis_config.items():
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
    severity_admin_status_list = results_dict.get('severity_category')
    dimension_admin_status_list = results_dict.get('dimension_pin')
    dimension_admin_status_in_need_list = results_dict.get('dimension_pin_in_need')
    severity_female_list = results_dict.get('severity_female')
    severity_male_list = results_dict.get('severity_male')
    dimension_female_list = results_dict.get('dimension_female')
    dimension_male_list = results_dict.get('dimension_male')
    dimension_ece_list = results_dict.get('dimension_ece')
    dimension_primary_list = results_dict.get('dimension_primary')
    dimension_secondary_list = results_dict.get('dimension_secondary')
    dimension_intermediate_list = results_dict.get('dimension_intermediate') if not single_cycle else None
    indicator_access_list = results_dict.get('indicator.access')
    indicator_teacher_list = results_dict.get('indicator.teacher')
    indicator_hazard_list = results_dict.get('indicator.hazard')
    indicator_idp_list = results_dict.get('indicator.idp')
    indicator_occupation_list = results_dict.get('indicator.occupation')
    indicator_barrier4_list = results_dict.get('indicator.barrier4')
    indicator_barrier5_list = results_dict.get('indicator.barrier5')
    indicator_barrier_list = results_dict.get(barrier_var)

    ## checking number of columns
    # Ensure columns for severity
    severity_needed_columns = [2.0, 3.0, 4.0, 5.0]
    severity_admin_status_list = ensure_columns(severity_admin_status_list, severity_needed_columns)
    severity_female_list = ensure_columns(severity_female_list, severity_needed_columns)
    severity_male_list = ensure_columns(severity_male_list, severity_needed_columns)

    # Ensure columns for dimension
    dimension_needed_columns = ['access', 'aggravating circumstances', 'learning condition', 'protected environment']
    dimension_admin_status_list = ensure_columns(dimension_admin_status_list, dimension_needed_columns)
    dimension_admin_status_in_need_list = ensure_columns(dimension_admin_status_in_need_list, dimension_needed_columns)
    dimension_female_list = ensure_columns(dimension_female_list, dimension_needed_columns)
    dimension_male_list = ensure_columns(dimension_male_list, dimension_needed_columns)
    dimension_ece_list = ensure_columns(dimension_ece_list, dimension_needed_columns)
    dimension_primary_list = ensure_columns(dimension_primary_list, dimension_needed_columns)
    dimension_secondary_list = ensure_columns(dimension_secondary_list, dimension_needed_columns)
    if not single_cycle:
        dimension_intermediate_list = ensure_columns(dimension_intermediate_list, dimension_needed_columns)

    # Clean indicator columns
    indicator_access_list = clean_indicator_columns(indicator_access_list, 'indicator_access_list')
    indicator_teacher_list = clean_indicator_columns(indicator_teacher_list, 'indicator_teacher_list')
    indicator_hazard_list = clean_indicator_columns(indicator_hazard_list, 'indicator_hazard_list')
    indicator_idp_list = clean_indicator_columns(indicator_idp_list, 'indicator_idp_list')
    indicator_occupation_list = clean_indicator_columns(indicator_occupation_list, 'indicator_occupation_list')
    indicator_barrier4_list = clean_indicator_columns(indicator_barrier4_list, 'indicator_barrier4_list')
    indicator_barrier5_list = clean_indicator_columns(indicator_barrier5_list, 'indicator_barrier5_list')
    #indicator_barrier_list = clean_indicator_columns(indicator_barrier_list, 'indicator_barrier_list')


    
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


    for pop_group, df in severity_admin_status_list.items():
        print(f"severity1  '{pop_group}':")
        print(df, "\n")
        print(df.columns)

    for pop_group, df in pin_by_indicator_status_list.items():
        print(f"indicator1  '{pop_group}':")
        print(df, "\n")
        print(df.columns)    

##--------------------------------------------------------------------------------------------
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


    for pop_group, df in pin_per_admin_status.items():
        print(f"severity2  '{pop_group}':")
        print(df, "\n")
        print(df.columns)

    for category, pop_group_df in pin_per_admin_status.items():
        
        pop_group_df[label_admin_severity] = 0

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
        # Save the updated DataFrame back to the dictionary
        pin_per_admin_status[category] = pop_group_df



    dimension_per_admin_status = {}
    for category, df in dimension_admin_status_list.items():
        # Fetch the corresponding DataFrame from the grouped data
        pop_group_df = df
        pop_group_df = pop_group_df.rename(columns={
            pop_group_var: 'Population group',
            int_acc: label_perc_acc,
            int_agg: label_perc_agg,
            int_lc: label_perc_lc,
            int_penv: label_perc_penv,
            int_out: label_perc_out
        })
        # Update the dictionary with the renamed DataFrame
        dimension_per_admin_status[category] = pop_group_df

    indicator_per_admin_status = {}
    for category, df in pin_by_indicator_status_list.items():
        # Fetch the corresponding DataFrame from the grouped data
        pop_group_df = df
        pop_group_df = pop_group_df.rename(columns={
            pop_group_var: 'Population group',
            'sev3_indicator_access': label_perc_sev3_indicator_access,
            'sev3_indicator_teacher': label_perc_sev3_indicator_teacher,
            'sev3_indicator_hazard': label_perc_sev3_indicator_hazard,
            'sev4_indicator_idp': label_perc_sev4_indicator_idp,
            'sev5_indicator_occupation': label_perc_sev5_indicator_occupation,
            'sev4_aggravating_circumstances': label_perc_sev4_aggravating_circumstances,
            'sev5_aggravating_circumstances': label_perc_sev5_aggravating_circumstances
        })
        # Update the dictionary with the renamed DataFrame
        indicator_per_admin_status[category] = pop_group_df



    for category, df in indicator_per_admin_status.items():
        # Fetch the corresponding DataFrame from the pin_per_admin_status
        pin_df = pin_per_admin_status.get(category)

        if pin_df is not None:
            # Select only the necessary columns for merging
            pin_df_subset = pin_df[[admin_var, label_admin_severity]]

            # Merge the severity label into the indicator DataFrame
            df = df.merge(pin_df_subset, on=admin_var, how='left')

        # Function to extract severity level from column names
        def extract_severity_level(col):
            match = re.search(r'severity level (\d+)', col)
            return int(match.group(1)) if match else None

        # Separate columns into severity and non-severity groups
        non_severity_columns = [col for col in df.columns if extract_severity_level(col) is None]
        severity_columns = [col for col in df.columns if extract_severity_level(col) is not None]

        # Sort severity columns numerically
        severity_columns_sorted = sorted(severity_columns, key=extract_severity_level)

        # Maintain the original order of non-severity columns, while sorting severity columns
        final_columns = non_severity_columns[:2] + severity_columns_sorted + non_severity_columns[2:]

        # Reorder DataFrame columns
        indicator_per_admin_status[category] = df[final_columns]  # Now using updated df







    for pop_group, df in indicator_per_admin_status.items():
        print(f"pin by indicartor  '{pop_group}':")
        print(df, "\n")
        print(df.columns)



    country_label = country.replace(" ", "_").replace("--", "_").replace("/", "_")





    percentage_round = 1
    figures_round = 0
    for category, df in indicator_per_admin_status.items():

            for col in df.columns:
                if "(ToT # children)" in col:
                    # Convert to numeric and round (total numbers)
                    df[col] = pd.to_numeric(df[col], errors='coerce').round(figures_round)
                elif "severity level" in col and "indicator" in col and "(ToT # children)" not in col:
                    # Convert to numeric, multiply by 100, and round as percentage
                    df[col] = pd.to_numeric(df[col], errors='coerce').multiply(100).round(2)
                elif "severity level" in col and "(ToT # children)" not in col:
                    # Convert to numeric and round normally (for other severity-level values)
                    df[col] = pd.to_numeric(df[col], errors='coerce').round(2)

            # Ensure no NaNs remain
            df.fillna(0, inplace=True)

            # Save modified DataFrame back into the dictionary
            indicator_per_admin_status[category] = df




    return pin_per_admin_status, dimension_admin_status_list, indicator_per_admin_status ,country_label