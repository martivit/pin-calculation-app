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

admin_var = 'Admin_1: States/Regions'#'Admin_2: Regions' 


vector_cycle = [10,14]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label::English'

# Path to your Excel file
excel_path = 'input/REACH_MMR_MMR2402_MSNA_Dataset_VALIDATED.xlsx'
excel_path_ocha = 'input/ocha_pop_MMR.xlsx'

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

ocha_data = pd.read_excel(pd.ExcelFile(excel_path_ocha, engine='openpyxl') )

##################################################################################################################################################################################################################
##################################################################################################################################################################################################################
#############################################################################        CALCULATION PIN              ################################################################################################
##################################################################################################################################################################################################################
##################################################################################################################################################################################################################
##################################################################################################################################################################################################################

edu_data_severity = add_severity (country, edu_data, household_data, choice_data, survey_data, ocha_data,
                                                                                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                age_var, gender_var,
                                                                                label, 
                                                                                admin_var, vector_cycle, start_school, status_var)



file_path = 'output_validation/00_edu_data_with_severity.xlsx'
# Save the DataFrame to an Excel file
edu_data_severity.to_excel(file_path, index=False, engine='openpyxl')



(severity_admin_status_list, dimension_admin_status_list, severity_female_list, severity_male_list, factor_category,  pin_per_admin_status, dimension_per_admin_status,
 female_pin_per_admin_status, male_pin_per_admin_status, 
 pin_per_admin_status_girl, pin_per_admin_status_boy,pin_per_admin_status_ece, pin_per_admin_status_primary, pin_per_admin_status_upper_primary, pin_per_admin_status_secondary, 
 Tot_PiN_JIAF, Tot_Dimension_JIAF, final_overview_df,final_overview_df_OCHA, 
   final_overview_dimension_df,final_overview_dimension_df_in_need,
 Tot_PiN_by_admin,
   country_label) = calculatePIN (country, edu_data_severity, household_data, choice_data, survey_data, ocha_data,
                                                                                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                age_var, gender_var,
                                                                                label, 
                                                                                admin_var, vector_cycle, start_school, status_var)





# Create the Excel files
jiaf_excel = create_output(Tot_PiN_JIAF, final_overview_df, final_overview_df_OCHA,"PiN TOTAL",  admin_var,  ocha= False)
ocha_excel = create_output(Tot_PiN_JIAF, final_overview_df, final_overview_df_OCHA, "PiN TOTAL",  admin_var,  ocha= True, tot_severity=Tot_PiN_by_admin)
#dimension_jiaf_excel = create_output(Tot_Dimension_JIAF, final_overview_dimension_df, "By dimension TOTAL",   admin_var, dimension= True, ocha= False)
#dimension_ocha_excel = create_output(Tot_Dimension_JIAF, final_overview_dimension_df, "By dimension TOTAL",  admin_var, dimension= True, ocha= True)
doc_output = create_snapshot_PiN(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need)



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

# Save jiaf_excel
with open("output_validation/final__JIAF__platform_output.xlsx", "wb") as f:
    f.write(jiaf_excel.getbuffer())

# Save ocha_excel
with open("output_validation/final__OCHA__platform_output.xlsx", "wb") as f:
    f.write(ocha_excel.getbuffer())

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