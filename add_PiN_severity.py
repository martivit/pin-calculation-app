import pandas as pd
#import fuzzywuzzy
from fuzzywuzzy import process
import numpy as np
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell


int_2 = '2.0'
int_3 = '3.0'
int_4 = '4.0'
int_5 = '5.0'
label_perc2 = '% 1-2'
label_perc3 = '% 3'
label_perc4 = '% 4'
label_perc5 = '% 5'
label_tot2 = '# 1-2'
label_tot3 = '# 3'
label_tot4 = '# 4'
label_tot5 = '# 5'
label_perc_tot = '% Tot PiN (3+)'
label_tot = '# Tot PiN (3+)'
label_admin_severity = 'Area severity'
label_tot_population = 'TotN'

int_acc = 'access'
int_agg= 'aggravating circumstances'
int_lc = 'learning condition'
int_penv = 'protected environment'
int_out = 'not falling within the PiN dimensions'
label_perc_acc = '% Access'
label_perc_agg= '% Aggravating circumstances'
label_perc_lc = '% Learning conditions'
label_perc_penv = '% Protected environment'
label_perc_out = '% Not falling within the PiN dimensions'
label_tot_acc = '# Access'
label_tot_agg= '# Aggravating circumstances'
label_tot_lc = '# Learning conditions'
label_tot_penv = '# Protected environment'
label_tot_out = '# Not falling within the PiN dimensions'
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
def calculate_severity(access, barrier, armed_disruption, idp_disruption, teacher_disruption, names_severity_4, names_severity_5):
    # Helper function to safely normalize string inputs
    def normalize(input_value):
        if isinstance(input_value, str):
            return input_value.lower()
        elif isinstance(input_value, (int, float)):  # Handle numeric values directly
            return input_value
        return ""  # Default to empty string if input is not a string or number
    
    # Normalize the input to handle different cases and languages
    normalized_access = normalize(access)
    normalized_armed_disruption = normalize(armed_disruption)
    normalized_idp_disruption = normalize(idp_disruption)
    normalized_teacher_disruption = normalize(teacher_disruption)

    # Normalize to handle English and French variations of "yes" and "no"
    yes_answers = ['yes', 'oui', '1', 1]
    no_answers = ['no', 'non', '0', 0]
    

    if normalized_access in no_answers:
        if barrier in names_severity_5:
            return 5
        elif barrier in names_severity_4:
            return 4
        else:
            return 3
    elif normalized_access in yes_answers:
        if normalized_armed_disruption in yes_answers:
            return 5
        elif normalized_idp_disruption in yes_answers:
            return 4
        elif normalized_teacher_disruption in yes_answers:
            return 3
        else:
            return 2
    return None  # Default fallback in case none of the conditions are met


##--------------------------------------------------------------------------------------------
def assign_dimension_pin(access, severity):
    # Normalize access status
    def normalize(input_string):
        if isinstance(input_string, str):
            return input_string.lower()
        return ""  # Default to empty string if input is not a string
    
    # Normalize the input to handle different cases and languages
    normalized_access = normalize(access)

    # Normalize to handle English and French variations of "yes" and "no"
    yes_answers = ['yes', 'oui']
    no_answers = ['no', 'non']

    # Mapping severity to dimension labels
    if normalized_access in no_answers:
        if severity in [4, 5]: return 'aggravating circumstances'
        elif severity == 3: return 'access'
    elif normalized_access in yes_answers:
        if severity == 3: return 'learning condition'
        if severity in [4, 5]: return 'protected environment'    
        if severity == 2: return 'not falling within the PiN dimensions'   
    
    return None  # Default fallback in case none of the conditions are met         

##--------------------------------------------------------------------------------------------
def print_subtables(severity_admin_status, pop_group_var):
    # Get the level number for pop_group_var
    level_number = severity_admin_status.index.names.index(pop_group_var)
    
    # Get unique groups
    unique_groups = severity_admin_status.index.get_level_values(level_number).unique()
    
    # Iterate and print subtables
    for group in unique_groups:
        subtable = severity_admin_status.xs(group, level=level_number)
        print(f"\nSubtable for {pop_group_var} = {group}")
        print(subtable)
        print("\n" + "-"*50 + "\n")

##--------------------------------------------------------------------------------------------
def save_subtables_to_excel(severity_admin_status, pop_group_var, file_path):
    # Get the level number for pop_group_var
    level_number = severity_admin_status.index.names.index(pop_group_var)
    
    # Get unique groups
    unique_groups = severity_admin_status.index.get_level_values(level_number).unique()
    
    # Create a Pandas Excel writer using XlsxWriter as the engine
    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        # Iterate and save subtables
        for group in unique_groups:
            subtable = severity_admin_status.xs(group, level=level_number)
            subtable.to_excel(writer, sheet_name=f"{pop_group_var}_{group}")
            print(f"-------------------- Subtable for {pop_group_var} = {group} saved to sheet {pop_group_var}_{group}")


##--------------------------------------------------------------------------------------------
def custom_to_datetime(date_str):
    try:
        return pd.to_datetime(date_str, errors='coerce')
    except:
        try:
            return pd.to_datetime(date_str, format='%Y-%m-%d %H:%M:%S.%f', errors='coerce')
        except:
            return pd.NaT

########################################################################################################################################
########################################################################################################################################
##############################################    PIN CALCULATION FUNCTION    ##########################################################
########################################################################################################################################
########################################################################################################################################
def add_severity (country, edu_data, household_data, choice_data, survey_data, ocha_data,
                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                age_var, gender_var,
                label, 
                admin_var, vector_cycle, start_school, status_var):

    admin_target = admin_var
    pop_group_var = status_var
    ocha_pop_data = ocha_data

    ## essential variables --------------------------------------------------------------------------------------------

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
        
    household_data['weight'] = 1
    # Find the UUID columns, assuming they exist and taking only the first match for simplicity
    edu_uuid_column = [col for col in edu_data.columns if 'uuid' in col.lower()][0]  # Take the first item directly
    household_uuid_column = [col for col in household_data.columns if 'uuid' in col.lower()][0]  # Take the first item directly

    household_start_column = [col for col in household_data.columns if 'start' in col.lower()][0]  # Take the first item directly


    # Extract the month from the 'start_time' column
    household_data[household_start_column] = household_data[household_start_column].apply(custom_to_datetime)
    household_data['month'] = household_data[household_start_column].dt.month

    # Find the most similar column to "Admin2" in household_data
    admin_var = process.extractOne(admin_target, household_data.columns.tolist())[0]  # Take the string directly
    # Columns to include in the merge
    columns_to_include = [household_uuid_column, admin_var, pop_group_var, 'month', 'weights', 'weight']


    columns_to_drop = [col for col in columns_to_include if col in edu_data.columns and col != edu_uuid_column and col != household_uuid_column]
    edu_data = edu_data.drop(columns=columns_to_drop, errors='ignore')

    # ----> Perform the joint_by
    edu_data = pd.merge(edu_data, household_data[columns_to_include], left_on=edu_uuid_column, right_on=household_uuid_column, how='left')
    ##refining for school age-children
    #edu_data = edu_data[(edu_data[age_var] >= 5) & (edu_data[age_var] <= 18)]

    edu_data['edu_age_corrected'] = edu_data.apply(lambda row: row[age_var] - 1 if calculate_age_correction(start_school, row['month']) else row[age_var], axis=1)
   
    edu_data = edu_data[(edu_data['edu_age_corrected'] >= 5) & (edu_data['edu_age_corrected'] <= 17)]


    ####### ** 2 **       ------------------------------ severity definition and calculation ------------------------------------------     #######
    severity_4_matches = find_matching_choices(choice_data, selected_severity_4_barriers, label_var=label)
    severity_5_matches = find_matching_choices(choice_data, selected_severity_5_barriers, label_var=label)
    names_severity_4 = [entry['name'] for entry in severity_4_matches]
    names_severity_5 = [entry['name'] for entry in severity_5_matches]

    # Apply the conditions and choices to create the new 'severity_category' column
    edu_data['severity_category'] = edu_data.apply(lambda row: calculate_severity(
        access=row[access_var], 
        barrier=row[barrier_var], 
        armed_disruption=row[armed_disruption_var], 
        idp_disruption=row[idp_disruption_var], 
        teacher_disruption=row[teacher_disruption_var], 
        names_severity_4=names_severity_4, 
        names_severity_5=names_severity_5
    ), axis=1)

    # Add the new column 'dimension_pin' to edu_data
    edu_data['dimension_pin'] = edu_data.apply(lambda row: assign_dimension_pin(
        access=row[access_var],
        severity= row['severity_category']
        ), axis=1)

    return edu_data


