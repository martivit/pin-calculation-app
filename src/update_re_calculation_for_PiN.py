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
import os

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

label_perc_sev3_indicator_access= 'severity level 3 -- OoS children -- % of children not accessing education who dot not face any aggravating circumstances'
label_perc_sev3_indicator_teacher = 'severity level 3 -- in-school children -- % of children whose ducation was disrupted by teacher absence'
label_perc_sev3_indicator_hazard = 'severity level 3 -- in-school children -- % of children whose education was disrupted by natural hazard'
label_perc_sev4_indicator_idp = 'severity level 4 -- in-school children -- % of children whose education was disrupted by the school being used as shelter'
label_perc_sev5_indicator_occupation = 'severity level 5 -- in-school children -- % of children whose education was disrupted by the school being occupied by armed groups'
label_perc_sev4_aggravating_circumstances = 'severity level 4, indicator: individual aggravating circumstances (cumulative of all Level 4 aggravating circumstances) -- % of children'
label_perc_sev5_aggravating_circumstances = 'severity level 5, indicator: individual aggravating circumstances (cumulative of all Level 5 aggravating circumstances) -- % of children'

label_tot_sev3_indicator_access= 'severity level 3 -- OoS children -- # of children not accessing education who dot not face any aggravating circumstances'
label_tot_sev3_indicator_teacher = 'severity level 3 -- in-school children -- # of children whose ducation was disrupted by teacher absence'
label_tot_sev3_indicator_hazard = 'severity level 3 -- in-school children -- # of children whose education was disrupted by natural hazard'
label_tot_sev4_indicator_idp = 'severity level 4 -- in-school children -- # of children whose education was disrupted by the school being used as shelter'
label_tot_sev5_indicator_occupation = 'severity level 5 -- in-school children -- # of children whose education was disrupted by the school being occupied by armed groups'
label_tot_sev4_aggravating_circumstances = 'severity level 4, indicator: individual aggravating circumstances (cumulative of all Level 4 aggravating circumstances) -- % of children'
label_tot_sev5_aggravating_circumstances = 'severity level 5, indicator: individual aggravating circumstances (cumulative of all Level 5 aggravating circumstances) -- % of children'

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
def build_ocha_by_category(ocha_data, mapped_statuses):
    ocha_by_category = {}

    for template_label, status_label in mapped_statuses.items():
        if status_label == "No match found":
            continue  # Skip

        # Try to find the matching column in ocha_data
        matching_col = next(
            (col for col in ocha_data.columns if template_label in col and 'children' in col.lower()),
            None
        )

        if matching_col:
            df = ocha_data[['Admin Pcode', matching_col]].copy()
            df = df.rename(columns={matching_col: 'TotN'})
            df['Category'] = status_label
            df['Population Group'] = status_label
            ocha_by_category[status_label] = df
            print(f"✅ Extracted for {status_label} from column: '{matching_col}'")
        else:
            print(f"❗ Could not find column for {template_label} ({status_label})")

    return ocha_by_category


##--------------------------------------------------------------------------------------------
def calculate_category_factors(df, total_col, category_col, category_name):
    """
    Calculate ratios for specific categories and filter the DataFrame to include only necessary columns.
    Returns a dictionary of DataFrames.
    """
    result_df = df.copy()
    result_df[category_name] = result_df[category_col] / result_df[total_col].replace(0, pd.NA)
    result_df['Category'] = category_name
    columns_to_keep = ['Admin Pcode',  category_name, 'Category']
    
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
        columns_to_keep = ['Admin Pcode',  category, 'Category']
        result[category] = temp_df[columns_to_keep]
    return result

##--------------------------------------------------------------------------------------------
# %PiN AND #PiN PER ADMIN AND POPULATION GROUP for the strata: GENDER, SCHOOL-CYCLE 
def adjust_pin_by_strata_factor(pin_df, factor_df, category_label, tot_column, admin_var):
    # Merge the pin DataFrame with the factor DataFrame on the 'Admin_2' column
    factorized_df = pd.merge(
        pin_df, factor_df, 
        left_on=[admin_var], 
        right_on=["Admin Pcode"], 
        how='left'
    )
   # Columns that need to be adjusted by the factor
    columns_to_adjust = [col for col in factorized_df.columns if col.startswith('#') or col == tot_column]
    del factorized_df['Admin Pcode']

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
    columns_to_keep = ['Admin Pcode', category_name, 'Category']
    
    # Return the filtered DataFrame
    return {category_name: result_df[columns_to_keep]}
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
##--------------------------------------------------------------------------------------------        



COUNTRY_CYCLE_MAP = [
    'Central African Republic -- CAR':    [8, 12],
    'Burkina Faso -- BFA':    [12, 0],
    'Ethiopia -- ETH':    [12, 0],
    'Democratic Republic of the Congo -- DRC':    [12, 0],
    'Mali -- MLI':    [12, 0],
    'Lebanon -- LBN':    [12, 0],
    'Somalia -- SOM':    [15, 0]
]

########################################################################################################################################
########################################################################################################################################
##############################################    PIN CALCULATION FUNCTION    ##########################################################
########################################################################################################################################
########################################################################################################################################
def UPDATE_calculatePIN (country , pin_2025_updated_perc, ocha_data,
                label ,  
                selected_language ):
    
    pin_2025_updated_perc = {
        key.replace('_pin_updated', ''): value
        for key, value in pin_2025_updated_perc.items()
    }
    category_list = list(pin_2025_updated_perc.keys())


    ocha_data = ocha_data.rename(columns={'Admin': 'Admin_label'})
    ocha_data = ocha_data.drop(columns=['Admin_label'])
    host_suggestion = ["Urban","PND",'host_community',"always_lived","general_pop",'non_deplace','Host Community',"Host community members",'host_communi', "always_lived","non_displaced_vulnerable",'host',"non_pdi","hote","menage_n_deplace","resident","lebanese","Populationnondéplacée","ocap","non_deplacee","Residents","yes","4"]
    IDP_suggestion = ['host_family','idp_host', 'PDI',"Rural","displaced","IDP", 'pdi_famille','New IDPs','pdi', 'idp', 'idp_host' ,"menage_deplace_interne", 'Out-of-camp','no',  'pdi_fam', '2', '1' ]
    returnee_suggestion = ['displaced_previously' ,'retournee','cb_returnee','retourne','ret','Returnee HH','returnee' ,'ukrainian moldovan','Returnees','5']
    refugee_suggestion = ['refugees','REF', 'refugee','refugie', 'refugie','prl', 'refugiee', '3']
    ndsp_suggestion = ['ndsp','Protracted IDPs', "hote affected by IDP",'displaced_camp','idp_site','pdi_site' "In-camp"]
    status_to_be_excluded = ['dnk', 'other', 'pnta', 'dont_know', 'no_answer', 'prefer_not_to_answer', 'pnpr', 'nsp', 'autre', 'do_not_know', 'decline']
    template_values = ['Host/Hôte',	'IDP/PDI',	'Returnees/Retournés', 'Refugees/Refugiees', 'Other']  
    suggestions_mapping = {
        'Host/Hôte': host_suggestion,        
        'IDP/PDI': IDP_suggestion,
        'Returnees/Retournés': returnee_suggestion,
        'Refugees/Refugiees': refugee_suggestion,
        'Other': ndsp_suggestion
    }

    
    ####### ** 4 **       ------------------------------ matching between the admin and the ocha population data ------------------------------------------     #######
    ## finding the match between the OCHA status cathegory and the country status. 
    satus_values = [status for status in category_list if status not in status_to_be_excluded]# Retrieve unique values directly without converting to lowercase
    for key, suggestions in suggestions_mapping.items():
        print(suggestions)
        suggestions_mapping[key] = suggestions  # keeping original case

    
    pop_group_var = "Population Group"

    mapped_statuses = map_template_to_status(template_values, suggestions_mapping, category_list)
    print("---------------------")
    print(mapped_statuses)
    print("---------------------")

    ocha_by_category = build_ocha_by_category(ocha_data, mapped_statuses)
    ####### ** 5 **       ------------------------------ creating tables with factors for the gender and school-cycle categories ------------------------------------------     #######
    try:
        vector_cycle = COUNTRY_CYCLE_MAP[country]
    except KeyError:
        raise ValueError(f"No entry for '{country}' in COUNTRY_CYCLE_MAP")
    single_cycle = (vector_cycle[1] == 0)
    primary_start = 6
    if country == 'Afghanistan -- AFG': 
        primary_start = 7

    secondary_end = 17
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
        **calculate_category_factors(ocha_data, children_tot_col, girls_tot_col, category_girl),
        **calculate_category_factors(ocha_data, children_tot_col, boys_tot_col, category_boy),
        **calculate_category_factors(ocha_data, children_tot_col, ece_tot_col, category_ece)
    }
    # Calculate factors for each school cycle
    school_cycle_factors = calculate_cycle_factors(ocha_data, factor_cycle, primary_start, secondary_end, vector_cycle, single_cycle)
    # Combine all factors into one dictionary
    factor_category = {**category_factors, **school_cycle_factors}
    disability_factors = add_disability_factor(ocha_data, factor_disability,category_disability )
    factor_category.update(disability_factors)

    for category, df in pin_2025_updated_perc.items():
        # Keep first column as-is, divide all others by 100
        df.iloc[:, 1:] = df.iloc[:, 1:] / 100
        pin_2025_updated_perc[category] = df

    admin_var = 'Admin Pcode'
    ####### ** 6.A **       ------------------------------ %PiN AND #PiN PER ADMIN AND POPULATION GROUP using ocha figures ------------------------------------------     #######
    pin_per_admin_status = {}
    # Assume category_data_frames is a dictionary of DataFrames, indexed by category
    for category, df in ocha_by_category.items():
        # Ensure both DataFrames are ready to merge
        if category in pin_2025_updated_perc:
            # Fetch the corresponding DataFrame from the grouped data
            grouped_df = pin_2025_updated_perc[category]     
            # Merge on specified columns
            pop_group_df = pd.merge(grouped_df, df, on=[admin_var])
            first_col = pop_group_df.columns[0]
            pop_group_df ['Population group'] = category
            cols = list(pop_group_df.columns)
            cols.insert(1, cols.pop(cols.index('Population group')))
            pop_group_df = pop_group_df[cols]         
            pop_group_df.columns = [str(col) for col in pop_group_df.columns]

            ## arranging columns 
            cols = list(pop_group_df.columns)
            pop_group_df = pop_group_df[cols]


            ## !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!   calculation of the tot Pin and admin severity -->
            # Step 1: Create the new column with initial zeros
            del pop_group_df['Category']

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
            cols.insert( cols.index('Population group') + 1, label_tot_population)
            pop_group_df = pop_group_df[cols]

            # Save modified DataFrame back into the dictionary under the category key
            pin_per_admin_status[category] = pop_group_df

    for category, df in pin_per_admin_status.items():
        print(f"pin by cat  '{category}':")
        print(df)


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


    for category, df in pin_per_admin_status.items():
        pin_per_admin_status_girl[category] = adjust_pin_by_strata_factor(df, factor_category[category_girl], category_girl, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_boy[category] = adjust_pin_by_strata_factor(df, factor_category[category_boy], category_boy, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_ece[category] = adjust_pin_by_strata_factor(df, factor_category[category_ece], category_ece, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_primary[category] = adjust_pin_by_strata_factor(df, factor_category[category_primary], category_primary, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_upper_primary[category] = adjust_pin_by_strata_factor(df, factor_category[category_upper_primary], category_upper_primary, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_secondary[category] = adjust_pin_by_strata_factor(df, factor_category[category_secondary], category_secondary, tot_column= label_tot_population, admin_var=admin_var)
        pin_per_admin_status_disabilty[category] = adjust_pin_by_strata_factor(df, factor_category[category_disability], category_disability, tot_column= label_tot_population, admin_var=admin_var)


    ####### ** 8.0 **       ------------------------------ aggregagte the popupulation group pin in 1 output by admin ------------------------------------------     #######
    overall_pin_per_admin_df = aggregate_pin_per_admin_status(pin_per_admin_status, admin_var)

    ####### ** 8.A **       ------------------------------ calculate tot PiN --> 3+ and admin severity for pin_per_admin_status ------------------------------------------     #######
    Tot_PiN_JIAF = pin_per_admin_status
    # Iterate over the pin_per_admin_status dictionary to apply the new operations
    for category, pop_group_df in Tot_PiN_JIAF.items():
        # Initialize new columns for percentage total, total PiN, and admin severity
        pop_group_df[label_perc_tot] = 0
        pop_group_df[label_tot] = 0
        pop_group_df[label_admin_severity] = 0

        # Reorder columns to place new columns at desired positions
        cols = list(pop_group_df.columns)
        cols.insert(cols.index(label_tot5) + 1, cols.pop(cols.index(label_perc_tot)))
        cols.insert(cols.index(label_perc_tot) + 1, cols.pop(cols.index(label_tot)))
        cols.insert(cols.index(label_tot) + 1, cols.pop(cols.index(label_admin_severity)))
        pop_group_df = pop_group_df[cols]

        # Calculate the total percentage and total PiN for severity levels 3+
        pop_group_df[label_perc_tot] = (pop_group_df[label_perc3] +
                                        pop_group_df[label_perc4] +
                                        pop_group_df[label_perc5])

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
        
        pop_group_df = pop_group_df.iloc[:, :-1]

        # Save the updated DataFrame back to the dictionary
        Tot_PiN_JIAF[category] = pop_group_df

    for category, df in Tot_PiN_JIAF.items():
        print(f"Tot_PiN_JIAF by cat  '{category}':")
        print(df)

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

    if country == 'Afghanistan -- AFG':
        tot_5_17_label = 'TOTAL (6-17 y.o.)'
        girl_5_17_label = 'Girls (6-17 y.o.)'
        boy_5_17_label = 'Boys (6-17 y.o.)'
        ece_5yo_label = 'ECE (6 y.o.)'

    ####### ** 9 **       ------------------------------  preparation for overview--> SUM all the admin per population group and per strata ------------------------------------------     #######
    overview_ToT = collapse_and_summarize(pin_per_admin_status, tot_5_17_label, admin_var=admin_var)
    overview_girl = collapse_and_summarize(pin_per_admin_status_girl, girl_5_17_label, admin_var=admin_var)
    overview_boy = collapse_and_summarize(pin_per_admin_status_boy, boy_5_17_label, admin_var=admin_var)
    overview_ece = collapse_and_summarize(pin_per_admin_status_ece, ece_5yo_label, admin_var=admin_var)
    overview_primary = collapse_and_summarize(pin_per_admin_status_primary, 'Primary school', admin_var=admin_var)
    overview_upper_primary = collapse_and_summarize(pin_per_admin_status_upper_primary, 'Intermediate school-level', admin_var=admin_var)
    overview_secondary = collapse_and_summarize(pin_per_admin_status_secondary, 'Secondary school', admin_var=admin_var)
    overview_disabilty = collapse_and_summarize(pin_per_admin_status_disabilty, 'Children with disability', admin_var=admin_var)


    print('----------------------------           overview_girl')
    print(overview_girl)

    collapsed_results_pop = {}
    for category, df in pin_per_admin_status.items():
            # Create a copy of the first row to preserve the structure
            summed_df = df.iloc[0:1].copy()

            # Identify columns to skip from summation and columns to set to zero
            columns_to_skip = [col for col in df.columns if col.startswith('%') or col == admin_var  or col == 'Population group' or col == 'Category' or col== 'Area severity']
            columns_to_zero = [col for col in df.columns if col.startswith('%')]

            # Sum all numerical columns except the skipped ones
            for col in df.columns:
                if col not in columns_to_skip:
                    summed_df[col] = df[col].sum()

            # Set non-sum columns with fixed values
            summed_df[admin_var] = 'whole country'
            summed_df['Population group'] = category
            del summed_df['Area severity']
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
        pin_per_admin_status = translate_labels(pin_per_admin_status, translation_dict)

    









    return Tot_PiN_JIAF,Tot_PiN_by_admin,final_overview_df_OCHA,final_overview_df , pin_per_admin_status






