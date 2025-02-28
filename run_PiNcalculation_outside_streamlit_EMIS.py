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
from src.calculation_for_PiN_Dimension_with_EMIS import calculatePIN_with_EMIS
from src.calculation_for_PiN_Dimension_NO_OCHA import calculatePIN_NO_OCHA
from src.vizualize_PiN import create_output
from src.vizualize_PiN import create_indicator_output
from src.snapshot_PiN import create_snapshot_PiN
from src.snapshot_PiN_FR import create_snapshot_PiN_FR
from src.save_parameter import generate_word_document
from src.save_parameter import generate_parameters
from src.save_parameter_FR import generate_word_document_FR
from src.save_parameter_FR import generate_parameters_FR
from docx import Document
from docx.shared import Pt, RGBColor
import matplotlib.pyplot as plt
from docx.shared import Inches




################################################
##           input from thee user             ##
################################################

data_combination = 'emmm'


## Lemuria
status_var = 'pop_group'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
armed_disruption_var = 'edu_disrupted_occupation'#'edu_disrupted_occupation'no_indicator
natural_hazard_var = 'no_indicator'
barrier_var = 'edu_barrier'
selected_severity_4_barriers = ['Cannot afford education-related costs (e.g. tuition, supplies, transportation)', 'There is a lack of interest/Education is not a priority either for the child or the household']#"L'école a été fermée en raison de dommages, d'une catastrophe naturelle ou d'un conflit.",, "Discrimination ou stigmatisation de l'enfant pour quelque raison que ce soit"
selected_severity_5_barriers = ['School has been closed due to natural disaster', 'School has been closed due to conflict', 'Lack of or poor quality of teachers', 'Protection/safety risks while commuting to school', 'Protection/safety risks while at school', 'Child marriage, engagement or pregnancies']
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'ind_age'
gender_var = 'ind_gender'
start_school = 'September'
country= 'Lemuria -- LMR'

#admin_var = 'Admin_3: Townships'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
admin_var = 'Admin_2: District'#'Admin_2: Regions' 

vector_cycle = [12,16]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label::English'

# Path to your Excel file
excel_path = 'input/Lemuria_MSNA_2022.xlsx'
excel_path_ocha = 'input/OCHA_pop_LMR.xlsx'
excel_path_emis = 'input/emis_LMR.xlsx'

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
household_data = dfs['01_clean_data_main']
edu_data = dfs['02_clean_data_indiv']
survey_data = dfs['survey']
choice_data = dfs['choices']

ocha_xls = pd.ExcelFile(excel_path_ocha, engine='openpyxl')

# Read specific sheets into separate dataframes
ocha_data = pd.read_excel(ocha_xls, sheet_name='ocha')  # 'ocha' sheet
mismatch_ocha_data = pd.read_excel(ocha_xls, sheet_name='scope-fix')  # 'scope-fix' sheet
mismatch_admin = False

emis_xls = pd.ExcelFile(excel_path_emis, engine='openpyxl')
emis_data = pd.read_excel(emis_xls)  # 'ocha' sheet


selected_language = "English"

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


if ocha_data is not None:
    (pin_by_indicator_status_list, enrollment_df) = calculatePIN_with_EMIS (data_combination,country, edu_data_severity, household_data, choice_data, survey_data, ocha_data,mismatch_ocha_data,emis_data,
                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                    age_var, gender_var,
                                                                                    label, 
                                                                                    admin_var, vector_cycle, start_school, status_var,
                                                                                    mismatch_admin,
                                                                                    selected_language= selected_language)




    file_path_E_1 = 'output_validation/E_indicator_results_withright_subset.xlsx'
    file_path_E_2 = 'output_validation/E_enrolment_by_pop_group.xlsx'


    # Create an Excel writer object
    with pd.ExcelWriter(file_path_E_1) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in pin_by_indicator_status_list.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)


    enrollment_df.to_excel(file_path_E_2, index=False, engine='openpyxl')
