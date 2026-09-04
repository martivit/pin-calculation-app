import pandas as pd
#import fuzzywuzzy
from fuzzywuzzy import process
import numpy as np
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
from io import BytesIO
from src.add_PiN_severity import add_severity
from src.clean_dataset import clean_make_dataset
from src.calculation_for_PiN_Dimension import calculatePIN
from src.calculation_for_PiN_Dimension_NO_OCHA import calculatePIN_NO_OCHA
from src.calculation_for_PiN_Dimension_NO_OCHA_2025 import calculatePIN_NO_OCHA_2025
from src.vizualize_PiN import create_output
from src.vizualize_PiN import create_indicator_output
from src.vizualize_PiN import create_indicator_output_no_ocha
from src.vizualize_PiN import create_pin_raw_output
from src.snapshot_PiN import create_snapshot_PiN
from src.snapshot_PiN_FR import create_snapshot_PiN_FR
from src.save_parameter import generate_word_document
from src.save_parameter import generate_parameters
from src.save_parameter_FR import generate_word_document_FR
from src.save_parameter_FR import generate_parameters_FR
from src.create_map_severity import make_map_severity

from docx import Document
from docx.shared import Pt, RGBColor
import matplotlib.pyplot as plt
from docx.shared import Inches
import os, glob




################################################
##           input from thee user             ##
################################################



hybrid_country= True
step_2_hpc= False


## MMR

status_var = 'pop_group'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
natural_hazard_var ='no_indicator'
natural_hazard_var_sev = None
additional_last_var = 'no_indicator'
additional_last_sev = None
additional_2_last_var = 'no_indicator'
additional_2_last_sev = None
armed_disruption_var = 'edu_disrupted_attack'#'edu_disrupted_occupation'no_indicator
barrier_var = 'edu_barrier'
selected_severity_4_barriers = [
    "Protection/safety risks while commuting to school",
    "Protection/safety risks while at school",
    "Child needs to work at home or on the household's own farm (i.e. is not earning an income for these activities, but may allow other family members to earn an income)",
    "Child participating in income generating activities outside of the home",
    "Child marriage, engagement or pregnancies",
    "Discrimination or stigmatization of the child for any reason",
    "Unable to enroll in school due to lack of documentation"]
selected_severity_5_barriers = ["Child is associated with armed forces or armed groups", "Pregnancy"]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'ind_age'
gender_var = 'ind_gender'
start_school = 'September'
country= 'Myanmar -- MMR'

selected_language = 'label::english (en)'

#admin_var = 'Admin_3: Townships'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
admin_var = 'admin1'#'Admin_2: Regions' 

vector_cycle = [10,14]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label::english (en)'

# Path to your Excel file
excel_path = 'input/MMR/REACH_MSNA 2026_Dataset.xlsx'
excel_path_ocha = 'input/MMR/Template_Population_figures.xlsx'
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
edu_data = dfs['indiv_clean_data']
household_data = dfs['main_clean_data']
survey_data = dfs['survey']
choice_data = dfs['choices']

ocha_xls = pd.ExcelFile(excel_path_ocha, engine='openpyxl')

# Read specific sheets into separate dataframes
ocha_data = pd.read_excel(ocha_xls, sheet_name='ocha')  # 'ocha' sheet
mismatch_ocha_data = pd.read_excel(ocha_xls, sheet_name='scope-fix')  # 'scope-fix' sheet
mismatch_admin = True




##################################################################################################################################################################################################################
##################################################################################################################################################################################################################
#############################################################################        CALCULATION PIN              ################################################################################################
##################################################################################################################################################################################################################
##################################################################################################################################################################################################################
##################################################################################################################################################################################################################

edu_data, household_data, survey_data, choice_data, messages = clean_make_dataset(
    country, edu_data, household_data, choice_data, survey_data,
    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
    natural_hazard_var, natural_hazard_var_sev,
    additional_last_var, additional_last_sev,
    additional_2_last_var, additional_2_last_sev,
    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
    age_var, gender_var,
    label,
    admin_var, vector_cycle, start_school, status_var,
    selected_language)

# status_var = "pop_status_group"
# age_var = "ind_age"
# gender_var = "ind_gender"
# barrier_var = "edu_barrier_final"

edu_data_severity = add_severity(country,
                                edu_data,
                                household_data,
                                choice_data,
                                survey_data,
                                access_var,
                                teacher_disruption_var,
                                idp_disruption_var,
                                armed_disruption_var,
                                natural_hazard_var,
                                natural_hazard_var_sev,
                                additional_last_var,
                                additional_last_sev,
                                additional_2_last_var,
                                additional_2_last_sev,
                                barrier_var,
                                selected_severity_4_barriers,
                                selected_severity_5_barriers,
                                age_var,
                                gender_var,
                                label,
                                admin_var,
                                vector_cycle,
                                start_school,
                                status_var,
                                selected_language=selected_language)



file_path = 'output_validation/00_edu_data_with_severity_MMR.xlsx'
# Save the DataFrame to an Excel file

if(type(edu_data_severity) is tuple):
    edu_data_severity_df=pd.DataFrame(edu_data_severity[0]).copy()
    edu_data_severity_df.to_excel(file_path, index=False, engine='openpyxl')
else:
    edu_data_severity[0].to_excel(file_path, index=False, engine='openpyxl')
