## MMR

status_var = 'pop_group'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
armed_disruption_var = 'edu_disrupted_occupation'#'edu_disrupted_occupation'no_indicator
barrier_var = 'edu_barrier'
selected_severity_4_barriers = [
    "Protection/safety risks while commuting to school",
    "Protection/safety risks while at school",
    "Child needs to work at home or on the household's own farm (i.e. is not earning an income for these activities, but may allow other family members to earn an income)",
    "Child participating in income generating activities outside of the home",
    "Child marriage, engagement or pregnancies",
    "Discrimination or stigmatization of the child for any reason",
    "Unable to enroll in school due to lack of documentation"]
selected_severity_5_barriers = ["Child is associated with armed forces or armed groups "]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'ind_age'
gender_var = 'ind_gender'
start_school = 'September'
country= 'Myanmar -- MMR'

#admin_var = 'Admin_3: Townships'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
admin_var = 'Admin_1: States/Regions'#'Admin_2: Regions' 

vector_cycle = [10,14]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label::English'

# Path to your Excel file
excel_path = 'input/REACH_MMR_MMR2402_MSNA_Dataset_VALIDATED.xlsx'
excel_path_ocha = 'input/ocha_pop_MMR.xlsx'
#excel_path_ocha = 'input/test_ocha.xlsx'

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
edu_data = dfs['02_clean_data_indiv']
household_data = dfs['01_clean_data_main']
survey_data = dfs['survey']
choice_data = dfs['choices']

ocha_xls = pd.ExcelFile(excel_path_ocha, engine='openpyxl')

# Read specific sheets into separate dataframes
ocha_data = pd.read_excel(ocha_xls, sheet_name='ocha')  # 'ocha' sheet
mismatch_ocha_data = pd.read_excel(ocha_xls, sheet_name='scope-fix')  # 'scope-fix' sheet
mismatch_admin = False



## BFA


status_var = 'i_type_pop'
access_var = 'e_enfant_scolarise_formel'
teacher_disruption_var = 'e_absence_enseignant'
idp_disruption_var = 'e_ecole_abris'
armed_disruption_var = 'no_indicator'#'edu_disrupted_occupation'no_indicator
barrier_var = 'e_raison_pas_educ_formel'
selected_severity_4_barriers = [
    "Risques de protection à l’école (tels que le harcèlement physique et verbal, risque de viol, les attaques contre les écoles ou d’autres incidents de protection)",
"Risques de protection pendant le trajet vers l’école (tels que les incidents de harcèlement physique et verbal, risque de viol ou d’autres incidents de protection)"
]
selected_severity_5_barriers = ["L'enfant est associé à des forces armées ou à des groupes armés"]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'sne_enfant_ind_age'
gender_var = 'sne_enfant_ind_genre'
start_school = 'September'
country= 'Burkina Faso -- BFA'

#admin_var = 'Admin_3: Townships'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
admin_var = 'Admin_3: Department (Département)'#'Admin_2: Regions' 

vector_cycle = [10,14]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label'

# Path to your Excel file
excel_path = 'input/BFA2402_MSNA_2024_DATA_CLEANED_VF.xlsx'
excel_path_ocha = 'input/ocha_pop_BFA.xlsx'
#excel_path_ocha = 'input/test_ocha.xlsx'

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
edu_data = dfs['loop_sne_cleaned']
household_data = dfs['main_cleaned']
survey_data = dfs['survey']
choice_data = dfs['choices']

ocha_xls = pd.ExcelFile(excel_path_ocha, engine='openpyxl')

# Read specific sheets into separate dataframes
ocha_data = pd.read_excel(ocha_xls, sheet_name='ocha')  # 'ocha' sheet
mismatch_ocha_data = pd.read_excel(ocha_xls, sheet_name='scope-fix')  # 'scope-fix' sheet
mismatch_admin = True




## AFG
status_var = 'urbanity'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
armed_disruption_var = 'no_indicator'#'edu_disrupted_occupation'no_indicator
natural_hazard_var = 'edu_disrupted_hazards'
barrier_var = 'resn_no_access'
selected_severity_4_barriers = [
 "Protection risks whilst at the school " ,
"Protection risks whilst travelling to the school ",
"Child needs to work at home or on the household's own farm (i.e. is not earning an income for these activities, but may allow other family members to earn an income) ",
"Child participating in income generating activities outside of the home"

]
selected_severity_5_barriers = ["Child is associated with armed forces or armed groups "]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'ind_age'
gender_var = 'edu_ind_gender'
start_school = 'November'
country= 'Afghanistan -- AFG'

#admin_var = 'Admin_3: Townships'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
admin_var = 'Admin_3: Districts'#'Admin_2: Regions' 

vector_cycle = [14,16]
single_cycle = (vector_cycle[1] == 0)
primary_start = 7
secondary_end = 17
label = 'label::English'

# Path to your Excel file
excel_path = 'input/AFG_WoAA_2024_data.xlsx'
excel_path_ocha = 'input/AFG_ocha.xlsx'
#excel_path_ocha = 'input/test_ocha.xlsx'

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
household_data = dfs['AFG_WoAA_2024_data_main_recoded']
edu_data = dfs['AFG_WoAA_2024_edu_loop']
survey_data = dfs['survey']
choice_data = dfs['choices']

ocha_xls = pd.ExcelFile(excel_path_ocha, engine='openpyxl')

# Read specific sheets into separate dataframes
ocha_data = pd.read_excel(ocha_xls, sheet_name='ocha')  # 'ocha' sheet
mismatch_ocha_data = pd.read_excel(ocha_xls, sheet_name='scope-fix')  # 'scope-fix' sheet
mismatch_admin = True

selected_language = "French"



## SOM
status_var = 'population_group'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
armed_disruption_var = 'edu_disrupted_hazards'#'edu_disrupted_occupation'no_indicator
natural_hazard_var = 'edu_disrupted_hazards'
barrier_var = 'edu_barrier'
selected_severity_4_barriers = [
"Protection risks whilst at the school " ,
"Protection risks whilst travelling to the school "

]
selected_severity_5_barriers = ["Child is associated with armed forces or armed groups "]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'edu_ind_age'
gender_var = 'edu_ind_gender'
start_school = 'September'
country= 'Somalia -- SOM'

#admin_var = 'Admin_3: Townships'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
admin_var = 'Admin_2: Districts'#'Admin_2: Regions' 

vector_cycle = [12,16]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label::english'

# Path to your Excel file
excel_path = 'input/REACH_MSNA_2024_FINAL_Cleaned_Weights.xlsx'
excel_path_ocha = 'input/ocha.xlsx'
#excel_path_ocha = 'input/test_ocha.xlsx'

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
household_data = dfs['main']
edu_data = dfs['edu_ind']
survey_data = dfs['survey']
choice_data = dfs['choices']

ocha_xls = pd.ExcelFile(excel_path_ocha, engine='openpyxl')

# Read specific sheets into separate dataframes
ocha_data = pd.read_excel(ocha_xls, sheet_name='ocha')  # 'ocha' sheet
mismatch_ocha_data = pd.read_excel(ocha_xls, sheet_name='scope-fix')  # 'scope-fix' sheet
mismatch_admin = False

selected_language = "English"




## NER

status_var = 'd_statut_deplacement'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
armed_disruption_var = 'edu_disrupted_hazards'#'edu_disrupted_occupation'no_indicator
natural_hazard_var = 'edu_disrupted_hazards'
barrier_var = 'edu_barrier'
selected_severity_4_barriers = [
    "Risques de protection à l’école (tels que le harcèlement physique et verbal, risque de viol, les attaques contre les écoles ou d’autres incidents de protection)",
"Risques de protection pendant le trajet vers l’école (tels que les incidents de harcèlement physique et verbal, risque de viol ou d’autres incidents de protection)"
]
selected_severity_5_barriers = ["L'enfant est associé à des forces armées ou à des groupes armés"]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'edu_age'
gender_var = 'edu_gender'
start_school = 'September'
country= 'Niger -- NER'

#admin_var = 'Admin_3: Townships'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
admin_var = 'Admin_2: Départements'#'Admin_2: Regions' 

vector_cycle = [12,16]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label::french'

# Path to your Excel file
excel_path = 'input/ner_msna_clean_data_FINAL.xlsx'
excel_path_ocha = 'input/ocha_NER_update.xlsx'
#excel_path_ocha = 'input/test_ocha.xlsx'

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
household_data = dfs['raw_data_clean']
edu_data = dfs['loop_data_clean']
survey_data = dfs['kobo_survey']
choice_data = dfs['kobo_choices']

ocha_xls = pd.ExcelFile(excel_path_ocha, engine='openpyxl')

# Read specific sheets into separate dataframes
ocha_data = pd.read_excel(ocha_xls, sheet_name='ocha')  # 'ocha' sheet
mismatch_ocha_data = pd.read_excel(ocha_xls, sheet_name='scope-fix')  # 'scope-fix' sheet
mismatch_admin = False

selected_language = "French"




## DRC

status_var = 'hoh_dis'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
armed_disruption_var = 'edu_disrupted_hazards'#'edu_disrupted_occupation'no_indicator
natural_hazard_var = 'edu_disrupted_hazards'
barrier_var = 'edu_barrier'
selected_severity_4_barriers = [
    "Risques de protection à l’école (tels que le harcèlement physique et verbal, risque de viol, les attaques contre les écoles ou d’autres incidents de protection)",
"Risques de protection pendant le trajet vers l’école (tels que les incidents de harcèlement physique et verbal, risque de viol ou d’autres incidents de protection)"
]
selected_severity_5_barriers = ["L'enfant est associé à des forces armées ou à des groupes armés"]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'edu_ind_age'
gender_var = 'edu_ind_gender'
start_school = 'September'
country= 'Democratic Republic of the Congo -- DRC'

#admin_var = 'Admin_3: Townships'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
admin_var = 'Admin_3: Sectors/chiefdoms/communes'#'Admin_2: Regions' 

vector_cycle = [12,16]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label::french'

# Path to your Excel file
excel_path = 'input/REACH_DRC2404_MSNA2024_Clean-Data.xlsx'
excel_path_ocha = 'input/DRC_ocha.xlsx'
#excel_path_ocha = 'input/test_ocha.xlsx'

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
household_data = dfs['hh_data']
edu_data = dfs['edu_data']
survey_data = dfs['survey']
choice_data = dfs['choices']

ocha_xls = pd.ExcelFile(excel_path_ocha, engine='openpyxl')

# Read specific sheets into separate dataframes
ocha_data = pd.read_excel(ocha_xls, sheet_name='ocha')  # 'ocha' sheet
mismatch_ocha_data = pd.read_excel(ocha_xls, sheet_name='scope-fix')  # 'scope-fix' sheet
mismatch_admin = False

selected_language = "French"



## CAR
status_var = 'type_population'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
armed_disruption_var = 'edu_disrupted_hazards'#'edu_disrupted_occupation'no_indicator
natural_hazard_var = 'edu_disrupted_hazards'
barrier_var = 'edu_barrier'
selected_severity_4_barriers = [
"Absence d'école appropriée et accessible"

]
selected_severity_5_barriers = ["Le handicap ou les problèmes de santé de l'enfant l'empêchent d'aller à l'école"]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'edu_ind_age'
gender_var = 'edu_ind_gender'
start_school = 'September'
country= 'Central African Republic -- CAR'

#admin_var = 'Admin_3: Townships'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
admin_var = 'Admin_2: Sub-prefectures (sous-préfectures)'#'Admin_2: Regions' 

vector_cycle = [12,16]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label::french'

# Path to your Excel file
excel_path = 'input/CAR2402_REACH_MSNA_Base-de-donnees-nettoyees_septembre-2024-1.xlsx'
excel_path_ocha = 'input/Ocha_pop_CAR.xlsx'
#excel_path_ocha = 'input/test_ocha.xlsx'

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
household_data = dfs['menage']
edu_data = dfs['Education']
survey_data = dfs['survey']
choice_data = dfs['choices']

ocha_xls = pd.ExcelFile(excel_path_ocha, engine='openpyxl')

# Read specific sheets into separate dataframes
ocha_data = pd.read_excel(ocha_xls, sheet_name='ocha')  # 'ocha' sheet
mismatch_ocha_data = pd.read_excel(ocha_xls, sheet_name='scope-fix')  # 'scope-fix' sheet
mismatch_admin = False
