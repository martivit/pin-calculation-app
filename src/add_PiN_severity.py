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

def calculate_severity(country, gender, age, access, barrier, armed_disruption, natural_hazard,idp_disruption, teacher_disruption,names_severity_4, names_severity_5):

    # Helper function to safely normalize string inputs
    def normalize(input_value):
        if isinstance(input_value, str):
            return input_value.lower()
        elif isinstance(input_value, (int, float)):  # Handle numeric values directly
            return input_value
        return ""  # Default to empty string if input is not a string or number

    # Normalize the input to handle different cases and languages
    normalized_age = normalize(age)
    normalized_gender = normalize(gender)
    normalized_access = normalize(access)
    normalized_armed_disruption = normalize(armed_disruption) if armed_disruption is not None else None
    normalized_natural_hazard = normalize(natural_hazard) if natural_hazard is not None else None
    normalized_idp_disruption = normalize(idp_disruption)
    normalized_teacher_disruption = normalize(teacher_disruption)
    #normalized_protection_at_school = normalize(protection_at_school) if protection_at_school is not None else None
    #normalized_protection_to_school = normalize(protection_to_school) if protection_to_school is not None else None

    # Normalize to handle English and French variations of "yes" and "no"
    yes_answers = ['yes', 'oui', '1', 1]
    no_answers = ['no', 'non', '0', 0]

    if country != 'Afghanistan -- AFG':
    # Main severity calculation logic
        if normalized_access in no_answers:
            if barrier in names_severity_5:
                return 5
            elif barrier in names_severity_4:
                return 4
            else:
                return 3
        elif normalized_access in yes_answers:
            # Check if 'armed_disruption' is valid and not None
            if normalized_armed_disruption is not None and normalized_armed_disruption in yes_answers:
                return 5
            elif normalized_idp_disruption in yes_answers:
                return 4
            elif normalized_teacher_disruption in yes_answers:
                return 3
            elif normalized_natural_hazard is not None and normalized_natural_hazard in yes_answers:
                return 3
            else:
                return 2
        
        return None  # Default fallback in case none of the conditions are met

    else: 
         # Main severity calculation logic
        if normalized_access in no_answers:
            if barrier in names_severity_5:
                return 5
            elif gender == 'female' and age > 12:
                return 5
            elif barrier in names_severity_4:
                return 4
            else:
                return 3
        elif normalized_access in yes_answers:
            # Check if 'armed_disruption' is valid and not None
            if normalized_armed_disruption is not None and normalized_armed_disruption in yes_answers:
                return 5
            elif normalized_idp_disruption in yes_answers:
                return 4
            elif normalized_teacher_disruption in yes_answers:
                return 3
            elif normalized_natural_hazard is not None and normalized_natural_hazard in yes_answers:
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
    yes_answers = ['yes', 'oui', 1, '1']
    no_answers = ['no', 'non', 0, '0']

    # Mapping severity to dimension labels
    if normalized_access in no_answers:
        if severity in [4, 5]: return 'aggravating circumstances'
        elif severity == 3: return 'access'
    elif normalized_access in yes_answers:
        if severity == 3: return 'learning condition'
        if severity in [4, 5]: return 'protected environment'    
        if severity == 2: return 'Not in need'   
    
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
        # Try the default date parsing first
        return pd.to_datetime(date_str, errors='coerce')
    except:
        try:
            # Handle the 'Y-m-d H:M:S.f+TZ' format
            return pd.to_datetime(date_str, format='%Y-%m-%dT%H:%M:%S.%f%z', errors='coerce')
        except:
            try:
                # Handle the 'Y-m-d H:M:S.f' format without time zone
                return pd.to_datetime(date_str, format='%Y-%m-%d %H:%M:%S.%f', errors='coerce')
            except:
                try:
                    # Handle the 'dd/mm/yyyy' format
                    return pd.to_datetime(date_str, format='%d/%m/%Y', errors='coerce')
                except:
                    # Return NaT if all parsing attempts fail
                    return pd.NaT

##--------------------------------------------------------------------------------------------
def assign_school_cycle(edu_age_corrected, single_cycle=False, lower_primary_start_var=6, lower_primary_end_var=13, upper_primary_end_var=None):

    if lower_primary_start_var == 6: primary_minus_one = 5
    else: primary_minus_one = 6
    if single_cycle:
        # If single cycle is True, handle as a primary to secondary without upper primary
        if lower_primary_start_var <= edu_age_corrected <= lower_primary_end_var:
            return 'primary'
        elif lower_primary_end_var + 1 <= edu_age_corrected <= 18:
            return 'secondary'
        elif edu_age_corrected == primary_minus_one: 
            return 'ECE'
        else:
            return 'out of school range'
    else:
        # If single cycle is False, handle lower primary, upper primary, and secondary phases
        if lower_primary_start_var <= edu_age_corrected <= lower_primary_end_var:
            return 'primary'
        elif upper_primary_end_var and lower_primary_end_var + 1 <= edu_age_corrected <= upper_primary_end_var:
            return 'intermediate level'
        elif upper_primary_end_var and upper_primary_end_var + 1 <= edu_age_corrected <= 18:
            return 'secondary'
        elif edu_age_corrected == primary_minus_one: 
            return 'ECE'
        else:
            return 'out of school range'
        
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
def add_severity (country, edu_data, household_data, choice_data, survey_data, 
                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                age_var, gender_var,
                label, 
                admin_var, vector_cycle, start_school, status_var,
                selected_language):

    admin_target = admin_var
    pop_group_var = status_var


    ## essential variables --------------------------------------------------------------------------------------------

    host_suggestion = ["Urban","always_lived",'Host Community','host_communi', "always_lived","non_displaced_vulnerable",'host',"non_pdi","hote","menage_n_deplace","menage_n_deplace","resident","lebanese","Populationnondéplacée","ocap","non_deplacee","Residents","yes","4"]
    IDP_suggestion = ["Rural","displaced", 'New IDPs','pdi', 'idp', 'site', 'camp', 'migrant', 'Out-of-camp', 'In-camp','no', 'pdi_site', 'pdi_fam', '2', '1' ]
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
    admin_levels_per_country = {
        'Afghanistan -- AFG': ['Admin_1: Region', 'Admin_2: Province', 'Admin_3: Districts'],
        'Burkina Faso -- BFA': ['Admin_1: Regions (Région)', 'Admin_2: Province', 'Admin_3: Department (Département)'],
        'Cameroon -- CMR': ['Admin_1', 'Admin_2', 'Admin_3'],
        'Central African Republic -- CAR': ['Admin_1: Prefectures (préfectures)', 'Admin_2: Sub-prefectures (sous-préfectures)', 'Admin_3: Communes'],
        'Democratic Republic of the Congo -- DRC': ['Admin_1: Provinces', 'Admin_2: Territories', 'Admin_3: Sectors/chiefdoms/communes'],
        'Haiti -- HTI': ['Admin_1: Departments (départements)', 'Admin_2: Arrondissements', 'Admin_3: Communes'],
        'Iraq -- IRQ': ['Admin_1: Governorates', 'Admin_2: Districts (aqḍyat)', 'Admin_3: Sub-districts (naḥiyat)'],
        'Kenya -- KEN': ['Admin_1: Counties', 'Admin_2: Sub-counties (kaunti ndogo)', 'Admin_3: Wards (mtaa)'],
        'Bangladesh -- BGD': ['Admin_1: Divisions (bibhag)', 'Admin_2: Districts (zila)', 'Admin_3: Upazilas'],
        'Lebanon -- LBN': ['Admin_1: Governorates', 'Admin_2: Districts (qaḍya)', 'Admin_3: Municipalities'],
        'Moldova -- MDA': ['Admin_1: Districts', 'Admin_2: Cities', 'Admin_3: Communes'],
        'Mali -- MLI': ['Admin_1: Régions', 'Admin_2: Cercles', 'Admin_3: Arrondissements'],
        'Mozambique -- MOZ': ['Admin_1: Provinces (provincias)', 'Admin_2: Districts (distritos)', 'Admin_3: Postos'],
        'Myanmar -- MMR': ['Admin_1: States/Regions', 'Admin_2: Districts', 'Admin_3: Townships'],
        'Niger -- NER': ['Admin_1: Régions ', 'Admin_2: Départements', 'Admin_3: Communes'],
        'Syria -- SYR': ['Admin_1: Governorates', 'Admin_2: Districts (mintaqah)', 'Admin_3: Subdistricts (nawaḥi)'],
        'Ukraine -- UKR': ['Admin_1: Oblasts', 'Admin_2: Raions', 'Admin_3: Hromadas'],
        'Somalia -- SOM': ['Admin_1: States', 'Admin_2: Regions', 'Admin_3: Districts']
    }



    ####### ** 1 **       ------------------------------ manipulation and join between H and edu data   ------------------------------------------     #######
        
    # Find the UUID columns, assuming they exist and taking only the first match for simplicity
    edu_uuid_column = [col for col in edu_data.columns if 'uuid' in col.lower()][0]  # Take the first item directly
    household_uuid_column = [col for col in household_data.columns if 'uuid' in col.lower()][0]  # Take the first item directly
    print(household_uuid_column)
    print(edu_uuid_column)



    if country != 'Afghanistan -- AFG':
        # Safely get the first column that contains 'start' in its name
        household_start_column = [col for col in household_data.columns if 'start' in col.lower()]
        if household_start_column:
            household_start_column = household_start_column[0]  # Take the first item directly
        else:
            raise KeyError("No column containing 'start' found in household_data.")
    else:
        # Assign the 'today' column if the country is Afghanistan
        household_start_column = 'today'
        if household_start_column not in household_data.columns:
            raise KeyError(f"'today' column is missing in household_data for Afghanistan.")

    # Convert the date column to datetime and extract the month
    
    household_data[household_start_column] = household_data[household_start_column].apply(custom_to_datetime)
    household_data[household_start_column] = pd.to_datetime(household_data[household_start_column], errors='coerce')

    household_data['month'] = household_data[household_start_column].dt.month
    #household_data['month'] = 7



    admin_var = find_best_match(admin_target,  household_data.columns)
    print(admin_var)


    weight_column = None

    # Check if 'weights' column exists, if not, find and rename the correct weight column
    if 'weights' not in household_data.columns:
        weight_column = [col for col in household_data.columns if 'weight' in col.lower()][0]  # Take the first matching weight column
        household_data = household_data.rename(columns={weight_column: 'weights'})
    else:
        print("--------------------------- Weights column already exists.")

    # Get the admin levels for the specified country
    admin_levels = admin_levels_per_country.get(country, [])
    # Flatten the admin levels to extract just the terms (like 'Region', 'District', etc.)
    admin_keywords = [term.split(": ")[-1].lower() for term in admin_levels]
    # Create a list of all household data columns that contain 'admin' or any of the admin keywords
    admin_columns_from_household = [col for col in household_data.columns if 'admin' in col.lower() or any(keyword in col.lower() for keyword in admin_keywords)]
    # Make sure admin_var is not duplicated
    if admin_var in admin_columns_from_household:
        admin_columns_from_household.remove(admin_var)
    # Now add the admin columns to the columns to include, without duplicating
    columns_to_include = [household_uuid_column, admin_var, pop_group_var, 'month', 'weights'] + admin_columns_from_household
    # Ensure there are no duplicate column names in columns_to_include
    columns_to_include = list(set(columns_to_include))

    print(columns_to_include)

    columns_to_drop = [col for col in columns_to_include if col in edu_data.columns and col != edu_uuid_column and col != household_uuid_column]
    edu_data = edu_data.drop(columns=columns_to_drop, errors='ignore')

    # ----> Perform the joint_by
    edu_data = pd.merge(edu_data, household_data[columns_to_include], left_on=edu_uuid_column, right_on=household_uuid_column, how='left')
    ##refining for school age-children
    #edu_data = edu_data[(edu_data[age_var] >= 5) & (edu_data[age_var] <= 18)]

    edu_data['edu_age_corrected'] = edu_data.apply(lambda row: row[age_var] - 1 if calculate_age_correction(start_school, row['month']) else row[age_var], axis=1)

    single_cycle = (vector_cycle[1] == 0)
    if country != 'Afghanistan -- AFG': primary_start = 6
    else: primary_start = 7
    edu_data['school_cycle'] = edu_data['edu_age_corrected'].apply(
        lambda x: assign_school_cycle(
            x, 
            single_cycle=single_cycle, 
            lower_primary_start_var=primary_start, 
            lower_primary_end_var=vector_cycle[0], 
            upper_primary_end_var=vector_cycle[1] if not single_cycle else None
        )
    )
   
    if country != 'Afghanistan -- AFG':
        edu_data = edu_data[(edu_data['edu_age_corrected'] >= 5) & (edu_data['edu_age_corrected'] <= 17)]
    else:
        edu_data = edu_data[(edu_data['edu_age_corrected'] >= 6) & (edu_data['edu_age_corrected'] <= 17)]


    ####### ** 2 **       ------------------------------ severity definition and calculation ------------------------------------------     #######
    severity_4_matches = find_matching_choices(choice_data, selected_severity_4_barriers, label_var=label)
    severity_5_matches = find_matching_choices(choice_data, selected_severity_5_barriers, label_var=label)
    names_severity_4 = [entry['name'] for entry in severity_4_matches]
    names_severity_5 = [entry['name'] for entry in severity_5_matches]

    edu_data['severity_category'] = edu_data.apply(lambda row: calculate_severity(
        country = country,
        gender = row[gender_var],
        age = row ['edu_age_corrected'],
        access=row[access_var], 
        barrier=row[barrier_var], 
        armed_disruption=row[armed_disruption_var] if armed_disruption_var != 'no_indicator' else None, 
        natural_hazard=row[natural_hazard_var] if natural_hazard_var != 'no_indicator' else None, 
        idp_disruption=row[idp_disruption_var], 
        teacher_disruption=row[teacher_disruption_var], 
        #protection_at_school=row['e_incident_ecol'] if country == 'Burkina Faso -- BFA'  else None,
        #protection_to_school=row['e_incident_trajet'] if country == 'Burkina Faso -- BFA'  else None,
        names_severity_4=names_severity_4, 
        names_severity_5=names_severity_5
    ), axis=1)

    # Add the new column 'dimension_pin' to edu_data
    edu_data['dimension_pin'] = edu_data.apply(lambda row: assign_dimension_pin(
        access=row[access_var],
        severity= row['severity_category']
        ), axis=1)




    return edu_data


