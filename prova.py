import pandas as pd
import fuzzywuzzy
from fuzzywuzzy import process
import numpy as np
import datetime
from pprint import pprint
import samplics
from samplics.categorical import Tabulation, CrossTabulation
from samplics.utils.types import PopParam, RepMethod


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
def find_matching_choices(choices_df, barriers_list):
    # List to hold the results
    results = []
    
    # Iterate over each barrier in the list
    for barrier in barriers_list:
        # Filter choices where 'label::english' matches the current barrier
        matched_choices = choices_df[choices_df['label::english'] == barrier]
        
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
        else:
            return 'out of school range'
    else:
        # If single cycle is False, handle lower primary, upper primary, and secondary phases
        if lower_primary_start_var <= edu_age_corrected <= lower_primary_end_var:
            return 'lower primary'
        elif upper_primary_end_var and lower_primary_end_var + 1 <= edu_age_corrected <= upper_primary_end_var:
            return 'upper primary'
        elif upper_primary_end_var and upper_primary_end_var + 1 <= edu_age_corrected <= 18:
            return 'secondary'
        else:
            return 'out of school range'
        
##--------------------------------------------------------------------------------------------
def calculate_severity(access, barrier, armed_disruption, idp_disruption, teacher_disruption, names_severity_4, names_severity_5):
    # Helper function to safely normalize string inputs
    def normalize(input_string):
        if isinstance(input_string, str):
            return input_string.lower()
        return ""  # Default to empty string if input is not a string
    
    # Normalize the input to handle different cases and languages
    normalized_access = normalize(access)
    normalized_armed_disruption = normalize(armed_disruption)
    normalized_idp_disruption = normalize(idp_disruption)
    normalized_teacher_disruption = normalize(teacher_disruption)

    # Normalize to handle English and French variations of "yes" and "no"
    yes_answers = ['yes', 'oui']
    no_answers = ['no', 'non']
    

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
    
    return 'no value'  # Default fallback in case none of the conditions are met         

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
            girls_col = f"{category} -- Girls/Filles (5-17)"
            boys_col = f"{category} -- Boys/Garcons (5-17)"
            
            # Check if these columns exist in the DataFrame
            if all(col in ocha_data.columns for col in [children_col, girls_col, boys_col]):
                # Create a new DataFrame for this category using the status as the category name
                category_df = ocha_data[['Admin', 'Admin Pcode', children_col, girls_col, boys_col]].copy()
                category_df.rename(columns={
                    children_col: 'Children (5-17)',
                    girls_col: 'Girls (5-17)',
                    boys_col: 'Boys (5-17)'
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
# what should arrive from the user selection
admin_target = 'Admin_2: Regions'
pop_group_var = 'place_of_origin'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
armed_disruption_var = 'edu_disrupted_occupation'
barrier_var = 'edu_barrier'
selected_severity_4_barriers = [
    "Protection risks whilst at the school ",
    "Protection risks whilst travelling to the school ",
    "Child needs to work at home or on the household's own farm (i.e. is not earning an income for these activities, but may allow other family members to earn an income) ",
    "Child participating in income generating activities outside of the home",
    "Marriage, engagement and/or pregnancy",
    "Unable to enroll in school due to lack of documentation",
    "Discrimination or stigmatization of the child for any reason"]
selected_severity_5_barriers = ["Child is associated with armed forces or armed groups "]
age_var = 'edu_ind_age'
gender_var = 'edu_ind_gender'
start_month = 'september'
single_cycle = False
lower_primary_start = 5
lower_primary_end = 10
upper_primary_end = 14


##--------------------------------------------------------------------------------------------
## status definition/suggestion:
host_suggestion = ["always_lived",'Host Community','host_communi', "always_lived","non_displaced_vulnerable",'host',"non_pdi","hote","menage_n_deplace","menage_n_deplace","resident","lebanese","Populationnondéplacée","ocap","non_deplacee","Residents","yes","4"]
IDP_suggestion = ["displaced", 'New IDPs','pdi', 'idp', 'site', 'camp', 'migrant', 'Out-of-camp', 'In-camp','no', 'pdi_site', 'pdi_fam', '2', '1' ]
returnee_suggestion = ['displaced_previously' ,'Protracted IDPs','cb_returnee','ret','Returnee HH','returnee' ,'ukrainian moldovan','Returnees','5']
refugee_suggestion = ['refugees', 'refugee', 'prl', 'refugiee', '3']
ndsp_suggestion = ['ndsp']
status_to_be_excluded = ['dnk', 'other', 'pnta', 'dont_know', 'no_answer', 'prefer_not_to_answer', 'pnpr', 'nsp', 'autre', 'do_not_know', 'decline']
template_values = ['Host/Hôte',	'IDP/PDI',	'Returnees/Retournés', 'Refugees/Refugiee', 'Other']
suggestions_mapping = {
    'Host/Hôte': host_suggestion,
    'IDP/PDI': IDP_suggestion,
    'Returnees/Retournés': returnee_suggestion,
    'Refugees/Refugiee': refugee_suggestion,
    'Other': ndsp_suggestion
}
##--------------------------------------------------------------------------------------------
##--------------------------------------------------------------------------------------------
##--------------------------------------------------------------------------------------------


# Path to your Excel file
excel_path = 'input/REACH_MSNA_2024_clean dataset_template_final.xlsx'
excel_path_ocha = 'input/ocha_pop.xlsx'

# Load the Excel file
xls = pd.ExcelFile(excel_path, engine='openpyxl')
# Print all sheet names (optional)
print(xls.sheet_names)
# Dictionary to hold your dataframes
dfs = {}
# Read each sheet into a dataframe
for sheet_name in xls.sheet_names:
    dfs[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)

# Access specific dataframes
edu_data = dfs['edu_ind data']
household_data = dfs['SOM2404_MSNA_Tool Data ']
survey = dfs['survey']
choices = dfs['choices']

ocha_pop_data = pd.read_excel(pd.ExcelFile(excel_path_ocha, engine='openpyxl') )


#######   ------ manipulation and join between H and edu data   ------   #######
household_data['weight'] = 1

# Find the UUID columns, assuming they exist and taking only the first match for simplicity
edu_uuid_column = [col for col in edu_data.columns if 'uuid' in col.lower()][0]  # Take the first item directly
household_uuid_column = [col for col in household_data.columns if 'uuid' in col.lower()][0]  # Take the first item directly


# Extract the month from the 'start_time' column
household_data['start'] = pd.to_datetime(household_data['start'])
household_data['month'] = household_data['start'].dt.month

# Find the most similar column to "Admin2" in household_data
admin_var = process.extractOne(admin_target, household_data.columns.tolist())[0]  # Take the string directly

# Columns to include in the merge
columns_to_include = [household_uuid_column, admin_var, pop_group_var, 'month', 'weights', 'weight']
edu_data = edu_data.drop(columns=[col for col in columns_to_include if col in edu_data.columns], errors='ignore')

# ----> Perform the joint_by
edu_data = pd.merge(edu_data, household_data[columns_to_include], left_on=edu_uuid_column, right_on=household_uuid_column, how='left')

##refining for school age-children
#edu_data = edu_data[(edu_data[age_var] >= 5) & (edu_data[age_var] <= 18)]
edu_data['edu_age_corrected'] = edu_data.apply(lambda row: row[age_var] - 1 if calculate_age_correction(start_month, row['month']) else row[age_var], axis=1)
edu_data['school_cycle'] = edu_data['edu_age_corrected'].apply(
    lambda x: assign_school_cycle(
        x, 
        single_cycle=single_cycle, 
        lower_primary_start_var=lower_primary_start, 
        lower_primary_end_var=lower_primary_end, 
        upper_primary_end_var=upper_primary_end if not single_cycle else None
    )
)
edu_data = edu_data[(edu_data['edu_age_corrected'] >= 5) & (edu_data['edu_age_corrected'] <= 17)]

severity_4_matches = find_matching_choices(choices, selected_severity_4_barriers)
severity_5_matches = find_matching_choices(choices, selected_severity_5_barriers)
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




## finding the match between the OCHA status cathegory and the country status. 
status_allvalues = edu_data[pop_group_var].unique()
status_values = [status for status in edu_data[pop_group_var].unique() if status not in status_to_be_excluded]# Retrieve unique values directly without converting to lowercase
for key, suggestions in suggestions_mapping.items():
    suggestions_mapping[key] = suggestions  # keeping original case

mapped_statuses = map_template_to_status(template_values, suggestions_mapping, status_values)
category_data_frames = extract_status_data(ocha_pop_data, mapped_statuses, pop_group_var)# Extract population figures based on mapped statuses without modifying the case

# Debugging and data inspection
for key, value in mapped_statuses.items():
    print(f"{key}: {value}")
for category, df in category_data_frames.items():
    df.rename(columns={'Admin': admin_var}, inplace=True)
    print(f"Category: {category}")
    print(df.head())  # Display the first few rows of the DataFrame
    print("\n" + "-"*50 + "\n")  # Print a separator for better readability between categories

#############################################################################################################
################################################## ANALYSIS #################################################
#############################################################################################################
pippo= edu_data.groupby(gender_var)
weight = edu_data["weights"]
#severity_cat = edu_data["severity_category"].to_numpy()
#dimension_cat = edu_data["dimension_pin"].to_numpy()
#gender_cat = edu_data[gender_var].to_numpy()

startum_gender = edu_data[gender_var]
startum_school_cycle = edu_data['school_cycle']




df = pd.DataFrame(edu_data)

print('------===================================================-------')



# Calculate weighted proportions for each category within each stratum_gender
severity_by_admin = edu_data.groupby([admin_var, 'severity_category']).agg(
    total_weight=('weights', 'sum')
).groupby(level=0).apply(
    lambda x: x / x.sum()
).unstack(fill_value=0)


pin_dimension_by_admin = df.groupby([admin_var, 'dimension_pin']).agg(
    total_weight=('weights', 'sum')
).groupby(level=0).apply(
    lambda x: x / x.sum()
).unstack(fill_value=0)

weighted_by_admingender = df.groupby([admin_var, gender_var]).agg(
    total_weight=('weights', 'sum')
).groupby(level=0).apply(
    lambda x: x / x.sum()
).unstack(fill_value=0)




print("\nseverity_by_admin")
print(severity_by_admin)

print("\npin_dimension_by_admin")
print(pin_dimension_by_admin)

print("\nknowing the demographic")
print(weighted_by_admingender)

print('                             ')
print('                             ')
print('               -------- GENDER DISAGGREGATION  ---------           ')
severity_by_gender = df.groupby([admin_var, pop_group_var,gender_var, 'severity_category']).agg(
    total_weight=('weights', 'sum')
).groupby(level=[0, 1]).apply(
    lambda x: x / x.sum()
).unstack(fill_value=0)

print(severity_by_gender)

print('                             ')
print('               -------- school-cycle DISAGGREGATION  ---------           ')
severity_by_cycle = df.groupby([admin_var, pop_group_var,startum_school_cycle, 'severity_category']).agg(
    total_weight=('weights', 'sum')
).groupby(level=[0, 1]).apply(
    lambda x: x / x.sum()
).unstack(fill_value=0)

print(severity_by_cycle)

print('                             ')
print('            -------    CORRECT PIN    -------             ')
severity_admin_status = df.groupby([admin_var, pop_group_var, 'severity_category']).agg(
    total_weight=('weights', 'sum')
).groupby(level=[0, 1]).apply(
    lambda x: x / x.sum()
).unstack(fill_value=0)

print(severity_admin_status)


print('***********************************************************')
def reduce_index(df, level):
    df.columns = df.columns.get_level_values(1)
    df=df.droplevel(0, axis=0) 
    df=df.droplevel(0, axis=0) 
    if level == 0: df = df.reset_index( level = [0 , 1] ) 
    if level == 1: df = df.reset_index( level = [0 , 1, 2] ) 

    # Splitting the DataFrame based on pop_group_var
    groups = df.groupby(pop_group_var)
    df_list = {name: group for name, group in groups}

    return df_list

severity_admin_status_list = reduce_index(severity_admin_status, 0)
severity_by_gender_list = reduce_index(severity_by_gender, 1)
severity_by_cycle_list = reduce_index(severity_by_cycle, 1)

output_file_path_test_strata_gender = 'output/severity_with_additional_strata_gender.xlsx'
output_file_path_test_strata_cycle = 'output/severity_with_additional_strata_cycle.xlsx'

with pd.ExcelWriter(output_file_path_test_strata_gender, engine='openpyxl') as writer:
    for group_name, df_group in severity_by_gender_list.items():
        # Each DataFrame is written to a separate sheet named after the group
        df_group.to_excel(writer, sheet_name=str(group_name), index=False)
with pd.ExcelWriter(output_file_path_test_strata_cycle, engine='openpyxl') as writer:
    for group_name, df_group in severity_by_cycle_list.items():
        # Each DataFrame is written to a separate sheet named after the group
        df_group.to_excel(writer, sheet_name=str(group_name), index=False)

print('=========================================')
# Constants for labels
LABELS_PERC = {2: '% 1-2', 3: '% 3', 4: '% 4', 5: '% 5'}
LABELS_TOT = {2: '# 1-2', 3: '# 3', 4: '# 4', 5: '# 5'}
LABEL_PERC_TOT = '% Tot PiN'
LABEL_TOT = '# Tot PiN'
LABEL_ADMIN_SEVERITY = 'Area severity'
excel_path = 'output_prova/severity_by_admin_and_pop_group.xlsx'


def rename_and_reorder(df, pop_group_var, admin_var):
    # Renaming numeric severity columns to string labels
    severity_columns = {2.0: LABELS_PERC[2], 3.0: LABELS_PERC[3], 4.0: LABELS_PERC[4], 5.0: LABELS_PERC[5]}
    df = df.rename(columns=severity_columns)
    df = df.rename(columns={pop_group_var: 'Population group'})

    # Explicitly reorder columns to the specified order
    desired_order = [
        'admin_2', 'Admin Pcode', 'Population group',
        LABELS_PERC[2], LABELS_TOT[2], LABELS_PERC[3], LABELS_TOT[3],
        LABELS_PERC[4], LABELS_TOT[4], LABELS_PERC[5], LABELS_TOT[5],
        LABEL_PERC_TOT, LABEL_TOT, LABEL_ADMIN_SEVERITY,
        'Children (5-17)', 'Girls (5-17)', 'Boys (5-17)'
    ]

    # Ensure only columns that exist in the DataFrame are reordered; this prevents KeyErrors
    df = df[[col for col in desired_order if col in df.columns]]
    return df


def calculate_totals(df):
    for severity in [2, 3, 4, 5]:
        perc_column = LABELS_PERC[severity]
        tot_column = LABELS_TOT[severity]
        if perc_column in df.columns:
            df[tot_column] = df[perc_column] * df['Children (5-17)']
        else:
            print(f"Column {perc_column} not found in DataFrame.")
            df[tot_column] = 0  # Default to 0 if the column isn't found
        
    df[LABEL_PERC_TOT] = df[[LABELS_PERC[3], LABELS_PERC[4], LABELS_PERC[5]]].sum(axis=1, skipna=True)
    df[LABEL_TOT] = df[[LABELS_TOT[3], LABELS_TOT[4], LABELS_TOT[5]]].sum(axis=1, skipna=True)
    
    conditions = [df[LABELS_PERC[5]] > 0.2, (df[LABELS_PERC[5]] + df[LABELS_PERC[4]]) > 0.2, (df[LABELS_PERC[5]] + df[LABELS_PERC[4]]+ df[LABELS_PERC[3]]) > 0.2, (df[LABELS_PERC[5]] + df[LABELS_PERC[4]]+ df[LABELS_PERC[3]]+  df[LABELS_PERC[2]]) > 0.2]
    choices = ['5', '4','3', '1-2']
    df[LABEL_ADMIN_SEVERITY] = np.select(conditions, choices, default='0')
    
    return df

def save_excel(category_data_frames, severity_admin_status_list, admin_var, pop_group_var, excel_path):
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        for category, df in category_data_frames.items():
            if category in severity_admin_status_list:
                grouped_df = severity_admin_status_list[category]
                merged_df = pd.merge(grouped_df, df, on=[admin_var, pop_group_var])
                merged_df = rename_and_reorder(merged_df, pop_group_var, admin_var)
                merged_df = calculate_totals(merged_df)
                merged_df.to_excel(writer, sheet_name=category)
                print(f"Processed and saved {category} to Excel.")




excel_path = 'output_prova/severity_by_admin_and_pop_group.xlsx'
save_excel(category_data_frames, severity_admin_status_list, admin_var, pop_group_var, excel_path)


# Call the function to print subtables
#print_subtables(severity_admin_status, pop_group_var)
        
# Define the output file path
output_file_path = 'output_prova/severity_admin_status.xlsx'
severity_admin_status.to_excel(output_file_path, index=False, engine='openpyxl')

# Call the function to save subtables to Excel
#save_subtables_to_excel(severity_admin_status, pop_group_var, output_file_path)



# Path for the Excel file


print(f"All categories have been saved to {excel_path}")




file_path = 'output_prova/edu_data_filtered.xlsx'

# Save the DataFrame to an Excel file
edu_data.to_excel(file_path, index=False, engine='openpyxl')
 






# Check if these columns exist in the DataFrame
category_df = ocha_pop_data[['Admin', 'Admin Pcode', children_tot_col]].copy()
category_df.rename(columns={children_tot_col: 'TotN'}, inplace=True)
category_df['Category'] = category_tot  # Set the category name to the matched status
category_df[pop_group_var] = 'All pop group'
category_data_frames[category_tot] = category_df

category_df_girl = ocha_pop_data[['Admin', 'Admin Pcode', girls_tot_col]].copy()
category_df_girl.rename(columns={girls_tot_col: 'TotN'}, inplace=True)
category_df_girl['Category'] = category_girl  # Set the category name to the matched status
category_df_girl[pop_group_var] = 'All pop group -- GIRL'
category_data_frames[category_girl] = category_df_girl

category_df_boy = ocha_pop_data[['Admin', 'Admin Pcode', boys_tot_col]].copy()
category_df_boy.rename(columns={boys_tot_col: 'TotN'}, inplace=True)
category_df_boy['Category'] = category_boy  # Set the category name to the matched status
category_df_boy[pop_group_var] = 'All pop group -- BOY'
category_data_frames[category_boy] = category_df_boy

category_df_ece = ocha_pop_data[['Admin', 'Admin Pcode', ece_tot_col]].copy()
category_df_ece.rename(columns={ece_tot_col: 'TotN'}, inplace=True)
category_df_ece['Category'] = category_ECE  # Set the category name to the matched status
category_df_ece[pop_group_var] = 'All pop group -- ECE'
category_data_frames[category_ECE] = category_df_ece





factor_girl = ocha_pop_data
factor_girl[category_girl] = ocha_pop_data[girls_tot_col]/ocha_pop_data[children_tot_col]
factor_girl['Category'] = category_girl
columns_to_keep = ['Admin', 'Admin Pcode',category_girl, 'Category']
factor_girl = factor_girl[columns_to_keep]

factor_boy = ocha_pop_data
factor_boy[category_boy] = ocha_pop_data[boys_tot_col]/ocha_pop_data[children_tot_col]
factor_boy['Category'] = category_boy
columns_to_keep = ['Admin', 'Admin Pcode',category_boy, 'Category']
factor_boy = factor_boy[columns_to_keep]

factor_ece = ocha_pop_data
factor_ece[category_ece] = ocha_pop_data[ece_tot_col]/ocha_pop_data[children_tot_col]
factor_ece['Category'] = category_ece
columns_to_keep = ['Admin', 'Admin Pcode',category_ece, 'Category']
factor_ece = factor_ece[columns_to_keep]



if single_cycle:
    factor_cycle[0] = (vector_cycle[0]-primary_start +1) / (secondary_end-primary_start+1)
    factor_cycle[1] = (secondary_end - vector_cycle[0]) / (secondary_end-primary_start+1)
else: 
    factor_cycle[0] = (vector_cycle[0]-primary_start +1) / (secondary_end-primary_start+1)
    factor_cycle[1] = (vector_cycle[1]-vector_cycle[0] ) / (secondary_end-primary_start+1)
    factor_cycle[2] = (secondary_end - vector_cycle[1]) / (secondary_end-primary_start+1)
print(factor_cycle)


if single_cycle:
    factor_primary = ocha_pop_data
    factor_primary[category_primary] = factor_cycle[0]
    factor_primary['Category'] = category_primary
    columns_to_keep = ['Admin', 'Admin Pcode',category_primary, 'Category']
    factor_primary = factor_primary[columns_to_keep]

    factor_secondary = ocha_pop_data
    factor_secondary[category_secondary] = factor_cycle[1]
    factor_secondary['Category'] = category_secondary
    columns_to_keep = ['Admin', 'Admin Pcode',category_secondary, 'Category']
    factor_secondary = factor_secondary[columns_to_keep]
    factor_upper_primary = ocha_pop_data
    factor_upper_primary[category_upper_primary] = 0
    factor_upper_primary['Category'] = category_upper_primary
    columns_to_keep = ['Admin', 'Admin Pcode',category_upper_primary, 'Category']
    factor_upper_primary = factor_upper_primary[columns_to_keep]
else:
    factor_primary = ocha_pop_data
    factor_primary[category_primary] = factor_cycle[0]
    factor_primary['Category'] = category_primary
    columns_to_keep = ['Admin', 'Admin Pcode',category_primary, 'Category']
    factor_primary = factor_primary[columns_to_keep]

    factor_secondary = ocha_pop_data
    factor_secondary[category_secondary] = factor_cycle[2]
    factor_secondary['Category'] = category_secondary
    columns_to_keep = ['Admin', 'Admin Pcode',category_secondary, 'Category']
    factor_secondary = factor_secondary[columns_to_keep]

    factor_upper_primary = ocha_pop_data
    factor_upper_primary[category_upper_primary] = factor_cycle[1]
    factor_upper_primary['Category'] = category_upper_primary
    columns_to_keep = ['Admin', 'Admin Pcode',category_upper_primary, 'Category']
    factor_upper_primary = factor_upper_primary[columns_to_keep]


print(factor_girl)
print(factor_boy)
print(factor_ece)
print(factor_primary)
print(factor_upper_primary)
print(factor_secondary)














## sum for PiN == 3+
        pop_group_df[label_perc_tot] = 0
        pop_group_df[label_tot] = 0
        pop_group_df[label_admin_severity] = 0
        cols = list(pop_group_df.columns)
        cols.insert(cols.index(label_tot5) + 1, cols.pop(cols.index(label_perc_tot)))
        cols.insert(cols.index(label_perc_tot) + 1, cols.pop(cols.index(label_tot)))
        cols.insert(cols.index(label_tot) + 1, cols.pop(cols.index(label_admin_severity)))

        # 3+
        pop_group_df[label_perc_tot] = pop_group_df[label_perc3] + pop_group_df[label_perc4] + pop_group_df[label_perc5] # 3+
        pop_group_df[label_tot] = pop_group_df[label_tot3] + pop_group_df[label_tot4] + pop_group_df[label_tot5] # 3+

        # Conditions based on your specifications
        conditions = [
            pop_group_df[label_perc5] > 0.2,
            (pop_group_df[label_perc5] + pop_group_df[label_perc4]) > 0.2,
            (pop_group_df[label_perc5] + pop_group_df[label_perc4] + pop_group_df[label_perc3]) > 0.2,
            (pop_group_df[label_perc5] + pop_group_df[label_perc4] + pop_group_df[label_perc3] + pop_group_df[label_perc2]) > 0.2
        ]

        # Corresponding values for each condition as strings
        choices = ['5', '4', '3', '1-2']

        # Applying the conditions and choices to the DataFrame with a string default
        pop_group_df[label_admin_severity] = np.select(conditions, choices, default='0')  




        #for col in pop_group_df.columns:
            #if col.startswith('#'):
                #pop_group_df[col] = pop_group_df[col].round(0)  # One digit after the decimal
            #elif col.startswith('%'):
                #pop_group_df[col] = pop_group_df[col].round(2)  # Two digits after the decimal

        #pop_group_df[label_tot_population] = pop_group_df[label_tot_population].round(0)






        def password_entered():
        """Checks whether a password entered by the user is correct."""
        h = sha256()
        pasw = st.session_state["password_user"]
        h.update(pasw.encode('utf-8'))
        hash = h.hexdigest()
        print(pasw)
        print(hash)
        
        if st.session_state["username"] in list(st.secrets.passwords.keys()) and hash in st.secrets.passwords[st.session_state["username"]]:
        # and hmac.compare_digest(   
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store the username or password.
            del st.session_state["username"]
        else:
            st.session_state["password_correct"] = False
 


        if st.session_state["username"]:
            user = st.session_state["username"]
            print(user)
            if user in st.secrets.passwords_assigned

    # Return True if the username + password is validated.
    if st.session_state.get("password_correct", False):
        return True
 
    # Show inputs for username + password.
    login_form()
    if "password_correct" in st.session_state:
        st.error("😕 User not known or password incorrect")
    return False





        cols = list(pop_group_df.columns)
        # Move the newly added column to the desired position
        cols.insert( cols.index(int_perc_penv) + 1, int_perc_out)

        cols.insert(cols.index(int_perc_acc) + 1, cols.pop(cols.index(int_tot_acc)))
        cols.insert(cols.index(int_perc_agg) + 1, cols.pop(cols.index(int_tot_agg)))
        cols.insert(cols.index(int_perc_lc) + 1, cols.pop(cols.index(int_tot_lc)))
        cols.insert(cols.index(int_perc_penv) + 1, cols.pop(cols.index(int_tot_penv)))
        cols.insert(cols.index(int_perc_out) + 1, cols.pop(cols.index(int_tot_out)))

        pop_group_df = pop_group_df[cols]     

        # Calculate total PiN for each severity level
        for perc_label, total_label in [(int_perc_acc, int_tot_acc), 
                                        (int_perc_agg, int_tot_agg), 
                                        (int_perc_lc, int_tot_lc), 
                                        (int_perc_penv, int_tot_penv),
                                        (int_perc_out, int_tot_out)]:
            pop_group_df[total_label] = pop_group_df[perc_label] * pop_group_df[label_dimension_tot_population]

        
        # Reorder columns as needed
        cols = list(pop_group_df.columns)
        cols.insert(cols.index('Population group') + 1, cols.pop(cols.index(label_dimension_tot_population)))
        pop_group_df = pop_group_df[cols]     


        cols.remove(label_dimension_tot_population)
        cols.insert( cols.index('Population group') + 1, label_dimension_tot_population)
        pop_group_df = pop_group_df[cols]