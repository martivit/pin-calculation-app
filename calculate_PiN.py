import pandas as pd
from fuzzywuzzy import process
import numpy as np


# Path to your Excel file
excel_path = 'input/SOM2404_MSNA_Tool_-_all_versions_-_False_-_2024-06-04-16-23-53 (1).xlsx'

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
edu_data = dfs['edu_ind']
household_data = dfs['SOM2404_MSNA_Tool']
survey = dfs['survey']
choices = dfs['choices']

# Find the UUID columns, assuming they exist and taking only the first match for simplicity
edu_uuid_column = [col for col in edu_data.columns if 'uuid' in col.lower()][0]  # Take the first item directly
household_uuid_column = [col for col in household_data.columns if 'uuid' in col.lower()][0]  # Take the first item directly

# Find the most similar column to "Admin2" in household_data
column_admin = process.extractOne('Admin2', household_data.columns.tolist())[0]  # Take the string directly

# Additional column you wish to include
pop_group_column = 'place_of_origin'
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

# Use the function to find matches for 'selected_severity_4_barriers'
severity_4_matches = find_matching_choices(choices, selected_severity_4_barriers)
severity_5_matches = find_matching_choices(choices, selected_severity_5_barriers)
names_severity_4 = [entry['name'] for entry in severity_4_matches]
names_severity_5 = [entry['name'] for entry in severity_5_matches]

print("Names for Severity 4 Barriers:", names_severity_4)
print("Names for Severity 5 Barriers:", names_severity_5)

names_for_target_label = [entry['name'] for entry in severity_4_matches if entry['label'] == 'Unable to enroll in school due to lack of documentation']



# Columns to include in the merge
columns_to_include = [household_uuid_column, column_admin, pop_group_column]

# Perform the merge
edu_data = pd.merge(edu_data, household_data[columns_to_include], left_on=edu_uuid_column, right_on=household_uuid_column, how='left')
edu_data = edu_data[(edu_data[age_var] >= 5) & (edu_data[age_var] <= 18)]

# Define conditions based on the decision criteria you provided
conditions = [
    # Condition group for "access_var == no"
    (edu_data[access_var] == 'no') & (edu_data[barrier_var].isin(names_severity_5)),  # Severity 5
    (edu_data[access_var] == 'no') & (edu_data[barrier_var].isin(names_severity_4)),  # Severity 4
    (edu_data[access_var] == 'no'),  # Severity 3, default for access_var == no

    # Condition group for "access_var == yes"
    (edu_data[access_var] == 'yes') & (edu_data[armed_disruption_var] == 'yes'),  # Severity 5
    (edu_data[access_var] == 'yes') & (edu_data[idp_disruption_var] == 'yes'),  # Severity 4
    (edu_data[access_var] == 'yes') & (edu_data[teacher_disruption_var] == 'yes'),  # Severity 3
    (edu_data[access_var] == 'yes')  # Severity 2, default for access_var == yes
]

# Define the severity categories corresponding to each condition
choices = [
    5,  # Severity 5 for access_var == no and barrier in severity 5 barriers
    4,  # Severity 4 for access_var == no and barrier in severity 4 barriers
    3,  # Default Severity 3 for all other cases when access_var == no
    5,  # Severity 5 for access_var == yes and armed disruption
    4,  # Severity 4 for access_var == yes and IDP disruption
    3,  # Severity 3 for access_var == yes and teacher disruption
    2   # Default Severity 2 for all other cases when access_var == yes
]

# Apply the conditions and choices to create the new 'severity_category' column
edu_data['severity_category'] = np.select(conditions, choices, default=0)  # Use default=1 for cases not covered explicitly

# Print the first few rows to verify the new 'severity_category'
print(edu_data[['severity_category']])

file_path = 'output/edu_data_filtered.xlsx'

# Save the DataFrame to an Excel file
edu_data.to_excel(file_path, index=False, engine='openpyxl')