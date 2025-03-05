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
    (severity_admin_status_list, dimension_admin_status_list, severity_female_list, severity_male_list, factor_category,  pin_per_admin_status, dimension_per_admin_status,indicator_per_admin_status,
    female_pin_per_admin_status, male_pin_per_admin_status, 
    pin_per_admin_status_girl, pin_per_admin_status_boy,pin_per_admin_status_ece, pin_per_admin_status_primary, pin_per_admin_status_upper_primary, pin_per_admin_status_secondary, 
    Tot_PiN_JIAF, Tot_Dimension_JIAF, final_overview_df,final_overview_df_OCHA, 
    final_overview_dimension_df,final_overview_dimension_df_in_need,
    Tot_PiN_by_admin,
    country_label) = calculatePIN (country, edu_data_severity, household_data, choice_data, survey_data, ocha_data,mismatch_ocha_data,
                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                    age_var, gender_var,
                                                                                    label, 
                                                                                    admin_var, vector_cycle, start_school, status_var,
                                                                                    mismatch_admin,
                                                                                    selected_language= selected_language)





    # Create the Excel files
    label_total_pin_sheet = "PiN TOTAL"


    ocha_excel = create_output(country_label,Tot_PiN_JIAF, final_overview_df, final_overview_df_OCHA, label_total_pin_sheet,  admin_var,  ocha= True, tot_severity=Tot_PiN_by_admin, selected_language=selected_language)

    indicator_output = create_indicator_output(country_label, indicator_per_admin_status, admin_var=admin_var)

    #dimension_jiaf_excel = create_output(Tot_Dimension_JIAF, final_overview_dimension_df, "By dimension TOTAL",   admin_var, dimension= True, ocha= False)
    #dimension_ocha_excel = create_output(Tot_Dimension_JIAF, final_overview_dimension_df, "By dimension TOTAL",  admin_var, dimension= True, ocha= True)
    if selected_language == 'English':
        doc_output = create_snapshot_PiN(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need,selected_language=selected_language)

    if selected_language == 'French':
        doc_output = create_snapshot_PiN_FR(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need,selected_language=selected_language)

    ##   ***********************************    save for intermediate check:
    file_path_pin_1 = 'output_validation/01_pin_percentage.xlsx'
    file_path_dimension_1 = 'output_validation/01_dimension_percentage.xlsx'
    file_path_pin_female_1 = 'output_validation/0a_pin_female_percentage.xlsx'
    file_path_pin_male_1 = 'output_validation/0a_pin_male_percentage.xlsx'
    file_path_factor = 'output_validation/02_factor_strata.xlsx'
    file_path_pin_2 = 'output_validation/03_pin_percentage_total_OCHA.xlsx'
    file_path_pin_female_2a = 'output_validation/0b_female_pin_percentage_total_OCHA.xlsx'
    file_path_pin_male_2a = 'output_validation/0b_male_pin_percentage_total_OCHA.xlsx'
    file_path_dimension_2 = 'output_validation/03_dimension_percentage_total_OCHA.xlsx'
    file_path_indicator_2 = 'output_validation/03_indicator_percentage_total_OCHA.xlsx'

    file_path_factor_girl3= 'output_validation/04_pin_factor_girl.xlsx'
    file_path_factor_boy3= 'output_validation/04_pin_factor_boy.xlsx'
    file_path_factor_ece3= 'output_validation/04_pin_factor_ECE.xlsx'
    file_path_factor_primary3= 'output_validation/04_pin_factor_primary.xlsx'
    file_path_factor_uprimary3= 'output_validation/04_pin_factor_upperprimary.xlsx'
    file_path_factor_secondary3= 'output_validation/04_pin_factor_secondary.xlsx'

    file_path_overview= 'output_validation/05_pin_overview.xlsx'
    file_path_overview_OCHA= 'output_validation/05_pin_overview_OCHA.xlsx'

    file_path_dimension_overview= 'output_validation/05_dimension_overview.xlsx'
    file_path_dimension_overview_in_need= 'output_validation/05_dimension_overview_in_need.xlsx'


    file_path_pin_tot_by_admin = 'output_validation/06_pin_tot_by_admin_area_severity.xlsx'


    # Create an Excel writer object
    with pd.ExcelWriter(file_path_pin_1) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in severity_admin_status_list.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)
    with pd.ExcelWriter(file_path_dimension_1) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in dimension_admin_status_list.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    with pd.ExcelWriter(file_path_pin_female_1) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in severity_female_list.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)
    with pd.ExcelWriter(file_path_pin_male_1) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in severity_male_list.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)
    with pd.ExcelWriter(file_path_factor) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in factor_category.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    with pd.ExcelWriter(file_path_pin_2) as writer:
    # Iterate over each category and DataFrame in the dictionary
        for category, df in pin_per_admin_status.items():
        # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)
    with pd.ExcelWriter(file_path_dimension_2) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in dimension_per_admin_status.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    with pd.ExcelWriter(file_path_indicator_2) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in indicator_per_admin_status.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)



    with pd.ExcelWriter(file_path_pin_female_2a) as writer:
    # Iterate over each category and DataFrame in the dictionary
        for category, df in female_pin_per_admin_status.items():
        # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)
    with pd.ExcelWriter(file_path_pin_male_2a) as writer:
    # Iterate over each category and DataFrame in the dictionary
        for category, df in male_pin_per_admin_status.items():
        # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)


    with pd.ExcelWriter(file_path_factor_girl3) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in pin_per_admin_status_girl.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)
    with pd.ExcelWriter(file_path_factor_boy3) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in pin_per_admin_status_boy.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)
    with pd.ExcelWriter(file_path_factor_ece3) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in pin_per_admin_status_ece.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    with pd.ExcelWriter(file_path_factor_primary3) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in pin_per_admin_status_primary.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)
    with pd.ExcelWriter(file_path_factor_uprimary3) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in pin_per_admin_status_upper_primary.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    with pd.ExcelWriter(file_path_factor_secondary3) as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in pin_per_admin_status_secondary.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)


    final_overview_df.to_excel(file_path_overview, index=False, engine='openpyxl')
    final_overview_df_OCHA.to_excel(file_path_overview_OCHA, index=False, engine='openpyxl')
    final_overview_dimension_df.to_excel(file_path_dimension_overview, index=False, engine='openpyxl')
    final_overview_dimension_df_in_need.to_excel(file_path_dimension_overview_in_need, index=False, engine='openpyxl')

    Tot_PiN_by_admin.to_excel(file_path_pin_tot_by_admin, index=False, engine='openpyxl')

    # Save the BytesIO objects to Excel files

    print('before saving')

    # Save ocha_excel
    with open("output_validation/final__OCHA__platform_output.xlsx", "wb") as f:
        f.write(ocha_excel.getbuffer())

    with open("output_validation/final__indicator__platform_output.xlsx", "wb") as f:
        f.write(indicator_output.getbuffer())    

    # Save dimension_jiaf_excel
    #with open("output_validation/final__dimension_JIAF__platform_output.xlsx", "wb") as f:
        #f.write(dimension_jiaf_excel.getbuffer())

    # Save dimension_ocha_excel
    #with open("output_validation/final__dimension_OCHA__platform_output.xlsx", "wb") as f:
        #f.write(dimension_ocha_excel.getbuffer())



    # Save the Word document to a file
    file_path = "output_validation/pin_snapshot_with_charts_and_text2.docx"
    with open(file_path, "wb") as f:
        f.write(doc_output.getvalue())