import pandas as pd
#import fuzzywuzzy
from fuzzywuzzy import process
import numpy as np
import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
from io import BytesIO
from add_PiN_severity import add_severity
from calculation_for_PiN_Dimension import calculatePIN
from vizualize_PiN import create_output
from docx import Document
from docx.shared import Pt, RGBColor
import matplotlib.pyplot as plt
from docx.shared import Inches
from snapshot_PiN import create_snapshot_PiN





################################################
##           input from thee user             ##
################################################


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
admin_var = 'Admin_2: Province'#'Admin_2: Regions' 

vector_cycle = [14,16]
single_cycle = (vector_cycle[1] == 0)
primary_start = 7
secondary_end = 17
label = 'label::English'

# Path to your Excel file
excel_path = 'input/AFG_WoAA_2024_data.xlsx'
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
household_data = dfs['AFG_WoAA_2024_data_main_recoded']
edu_data = dfs['AFG_WoAA_2024_edu_loop']
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

edu_data_severity = add_severity (country, edu_data, household_data, choice_data, survey_data,
                                                                                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                                                                                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                age_var, gender_var,
                                                                                label, 
                                                                                admin_var, vector_cycle, start_school, status_var)



file_path = 'output_validation/00_edu_data_with_severity.xlsx'
# Save the DataFrame to an Excel file
edu_data_severity.to_excel(file_path, index=False, engine='openpyxl')

