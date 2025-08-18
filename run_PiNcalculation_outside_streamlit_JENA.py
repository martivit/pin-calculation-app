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
from src.calculation_for_PiN_Dimension_with_JENA import calculatePIN_with_JENA
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

data_combination = 'mjjm'



## NER

status_var = 'd_statut_deplacement'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
armed_disruption_var = 'edu_disrupted_occupation'#'edu_disrupted_occupation'no_indicator
natural_hazard_var = 'edu_disrupted_hazards'
barrier_var = 'edu_barrier'
selected_severity_4_barriers = [
    "Risques de protection à l'école",
"Risques de protection pendant le trajet vers l'école",
"L'enfant doit travailler à la maison ou dans la ferme du ménage (c'est-à-dire qu'il ne gagne pas de revenu pour ces activités, mais peut permettre à d'autres membres de la famille de gagner un revenu)",
"L'enfant participe à des activités génératrices de revenus en dehors du foyer"

]
selected_severity_5_barriers = ["L'enfant est associé à des forces armées ou à des groupes armés","Impossibilité de payer les coûts directs de l'éducation (par exemple, les frais de scolarité, les fournitures, le transport)" ]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'ind_age'
gender_var = 'ind_gender'
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
excel_path_jena = 'input/Niger_JENA_PTR_protection.xlsx'
jena_exls = pd.ExcelFile(excel_path_jena, engine='openpyxl')
jena_data = pd.read_excel(jena_exls)  # 'ocha' sheet


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
country_label = country.replace(" ", "_").replace("--", "_").replace("/", "_")


if ocha_data is not None:
    (jena_df, merged_ocha_jena, merged_ocha_jena_msna,pin_jena_msna, Tot_PiN_JIAF,Tot_Dimension_JIAF, final_overview_df_OCHA, final_overview_df, Tot_PiN_by_admin)=  calculatePIN_with_JENA (data_combination, country, edu_data_severity, household_data, choice_data, survey_data, ocha_data,mismatch_ocha_data,jena_data,
                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                age_var, gender_var,
                label, 
                admin_var, vector_cycle, start_school, status_var,
                mismatch_admin,
                selected_language)




    file_path_E_1 = 'output_validation/J_indicator_results_withright_subset.xlsx'
    file_path_E_2 = 'output_validation/J_indicator_results_withright_subset_masna.xlsx'
    file_path_E_3 = 'output_validation/J_sev_children.xlsx'
    file_path_E_4 = 'output_validation/J_severity_pop_group.xlsx'
    file_path_E_5 = 'output_validation/J_PiN_pop_group_working.xlsx'
    file_path_E_6 = 'output_validation/J_PiN_pop_group.xlsx'
    file_path_E_7 = 'output_validation/J_PiN_by_dimension.xlsx'


    jena_df.to_excel(file_path_E_3, index=False, engine='openpyxl')
    final_overview_df.to_excel('output_validation/J_test.xlsx', index=False, engine='openpyxl')


    # Create an Excel writer object
    with pd.ExcelWriter(file_path_E_1) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in merged_ocha_jena.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

                # Create an Excel writer object
    with pd.ExcelWriter(file_path_E_2) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in merged_ocha_jena_msna.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    with pd.ExcelWriter(file_path_E_5) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in pin_jena_msna.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    with pd.ExcelWriter(file_path_E_6) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in Tot_PiN_JIAF.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    with pd.ExcelWriter(file_path_E_7) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in Tot_Dimension_JIAF.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    print(' jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj') 
    print(final_overview_df)
    print(' jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj') 

    if selected_language == "English":
        doc_output = create_snapshot_PiN(country_label, final_overview_df, final_overview_df_OCHA, selected_language=selected_language)
    if selected_language == "French":
        doc_output = create_snapshot_PiN_FR(country_label, final_overview_df, final_overview_df_OCHA,selected_language=selected_language)
        