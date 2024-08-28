import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
from add_PiN_severity import add_severity
from calculation_for_PiN_Dimension import calculatePIN
from vizualize_PiN import create_output
from snapshot_PiN import create_snapshot_PiN
from shared_utils import language_selector


st.logo('pics/logos.png')

st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico',  layout='wide')

# Call the language selector function
language_selector()

# Access the translations
translations = st.session_state.translations

if 'password_correct' not in st.session_state:
    st.error(translations["no_user"])
    st.stop()

if 'uploaded_data' not in st.session_state:
    st.warning(translations["no_data"])  
    st.stop()

## ====================================================================================================
## ===================================== calculate and download the PiN
## ====================================================================================================
# Streamlit app layout
st.title(translations["pin_calculation_results_title"])
st.write(translations["pin_calculation_message"])



start_school =  st.session_state.get('start_school')
vector_cycle =  st.session_state.get('vector_cycle')
country =  st.session_state.get('country')
edu_data =  st.session_state.get('edu_data')  # Convert DataFrame to dict
household_data =  st.session_state.get('household_data')  # Convert DataFrame to dict
status_var =  st.session_state.get('status_var')
survey_data =  st.session_state.get('survey_data')  # Convert DataFrame to dict
choice_data =  st.session_state.get('choice_data') # Convert DataFrame to dict
label =  st.session_state.get('label')
age_var =  st.session_state.get('age_var')
gender_var =  st.session_state.get('gender_var')
access_var =  st.session_state.get('access_var')
teacher_disruption_var =  st.session_state.get('teacher_disruption_var')
idp_disruption_var =  st.session_state.get('idp_disruption_var')
armed_disruption_var =  st.session_state.get('armed_disruption_var')
barrier_var =  st.session_state.get('barrier_var')
selected_severity_4_barriers =  st.session_state.get('selected_severity_4_barriers', [])
selected_severity_5_barriers =  st.session_state.get('selected_severity_5_barriers', [])
admin_var =  st.session_state.get('admin_var')
ocha_data = st.session_state.get('uploaded_ocha_data')



## add indicator ---> severity
edu_data_severity = add_severity (country, edu_data, household_data, choice_data, survey_data, ocha_data,
                                                                                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                age_var, gender_var,
                                                                                label, 
                                                                                admin_var, vector_cycle, start_school, status_var)





## calculate PiN
(severity_admin_status_list, dimension_admin_status_list, severity_female_list, severity_male_list, factor_category,  pin_per_admin_status, dimension_per_admin_status,
 female_pin_per_admin_status, male_pin_per_admin_status, 
 pin_per_admin_status_girl, pin_per_admin_status_boy,pin_per_admin_status_ece, pin_per_admin_status_primary, pin_per_admin_status_upper_primary, pin_per_admin_status_secondary, 
 Tot_PiN_JIAF, Tot_Dimension_JIAF, final_overview_df, final_overview_df_OCHA, final_overview_df_MSNA,
 final_overview_dimension_df,
 Tot_PiN_by_admin,
   country_label) = calculatePIN (country, edu_data_severity, household_data, choice_data, survey_data, ocha_data,
                                                                                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                age_var, gender_var,
                                                                                label, 
                                                                                admin_var, vector_cycle, start_school, status_var)






# Create the Excel files
#jiaf_excel = create_output(Tot_PiN_JIAF, final_overview_df, "PiN TOTAL",   admin_var, dimension= False, ocha= False, tot_severity=Tot_PiN_by_admin)
ocha_excel = create_output(Tot_PiN_JIAF, final_overview_df, final_overview_df_OCHA, "PiN TOTAL",  admin_var,  ocha= True, tot_severity=Tot_PiN_by_admin)
#dimension_jiaf_excel = create_output(Tot_Dimension_JIAF, final_overview_dimension_df, "By dimension TOTAL",   admin_var, dimension= True, ocha= False)
#dimension_ocha_excel = create_output(Tot_Dimension_JIAF, final_overview_dimension_df, "By dimension TOTAL",  admin_var, dimension= True, ocha= True)

#doc_snapshot = create_snapshot_PiN(country_label, final_overview_df, final_overview_dimension_df)




#st.download_button(
#    label="Download PiN JIAF Excel",
#    data=jiaf_excel.getvalue(),
#    file_name=f"PiN_JIAF_{country_label}.xlsx",
#    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#)

st.download_button(
    label=translations["download_pin"],
    data=ocha_excel.getvalue(),
    file_name=f"PiN_overview_OCHA_{country_label}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

#st.download_button(
#    label="Download Dimension JIAF Excel",
#    data=dimension_jiaf_excel.getvalue(),
#    file_name=f"Dimension_JIAF_{country_label}.xlsx",
#    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#)

#st.download_button(
#    label="Download Dimension OCHA Excel",
#    data=dimension_ocha_excel.getvalue(),
#    file_name=f"Dimension_overview_OCHA_{country_label}.xlsx",
#    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#)

#st.download_button(
#    label=translations["download_word"],
#    data=doc_snapshot.getvalue(),
#    file_name="pin_snapshot.docx",
#    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
#)

st.subheader(translations["hno_guidelines_subheader"])
st.markdown(translations["hno_guidelines_message"])