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






## DRC

status_var = 'pop_group'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disruption_teacher'
idp_disruption_var = 'edu_disruption_displaced'
armed_disruption_var = 'no_indicator'#'edu_disrupted_occupation'no_indicator
natural_hazard_var = 'no_indicator'

barrier_var = 'edu_barrier'
selected_severity_4_barriers = ["Risques de protection pendant le trajet vers l'école",
"Risques de protection à l'école",
"Enfant aidant à la maison / à la ferme", "Impossibilité d'enregistrer ou d'inscrire l'enfant à l'école" ,  "Mariage et/ou grossesse"                                                            
]
selected_severity_5_barriers = ["Les enfants rejoignent ou sont recrutés par des groupes armés"]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'age_years'
gender_var = 'ind_gender'
start_school = 'September'
country= 'Democratic Republic of the Congo -- DRC'

admin_var = 'Admin_3'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
#admin_var = 'Admin_1: States/Regions'#'Admin_2: Regions' 

vector_cycle = [11,0]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label'

# Path to your Excel file
excel_path = 'input/REACH_MSNA_2023_DRC_clean dataset_v2.xlsx'
excel_path_ocha = 'input/DRC_ocha_2025.xlsx'
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
edu_data = dfs['HH roster']
household_data = dfs['BDD nettoyée']
survey_data = dfs['Questionnaire']
choice_data = dfs['Options']


ocha_xls = pd.ExcelFile(excel_path_ocha, engine='openpyxl')

# Read specific sheets into separate dataframes
ocha_data = pd.read_excel(ocha_xls, sheet_name='ocha')  # 'ocha' sheet
mismatch_ocha_data = pd.read_excel(ocha_xls, sheet_name='scope-fix')  # 'scope-fix' sheet
mismatch_admin = False
no_ocha_data = False

selected_language = "French"




##################################################################################################################################################################################################################
##################################################################################################################################################################################################################
#############################################################################        CALCULATION PIN              ################################################################################################
##################################################################################################################################################################################################################
##################################################################################################################################################################################################################
##################################################################################################################################################################################################################

edu_data_severity = add_severity (country, edu_data, household_data, choice_data, survey_data,
                                                                                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                                                                                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                age_var, gender_var,
                                                                                label, 
                                                                                admin_var, vector_cycle, start_school, status_var,
                                                                                selected_language= selected_language)



file_path = 'output_validation/00_edu_data_with_severity.xlsx'
# Save the DataFrame to an Excel file
edu_data_severity.to_excel(file_path, index=False, engine='openpyxl')
