import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
from add_PiN_severity import add_severity
from calculation_for_PiN_Dimension import calculatePIN
from calculation_for_PiN_Dimension_NO_OCHA import calculatePIN_NO_OCHA
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
# Access the OCHA data if it was uploaded
ocha_data = st.session_state.get('uploaded_ocha_data')
mismatch_ocha_data = st.session_state.get('ocha_mismatch_data')


# Check if the user indicated that they do not have OCHA data
no_ocha_data = st.session_state.get('no_upload_ocha_data', False)
mismatch_admin = st.session_state.get('mismatch_admin', False)



## add indicator ---> severity
edu_data_severity = add_severity (country, edu_data, household_data, choice_data, survey_data, 
                                                                                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                age_var, gender_var,
                                                                                label, 
                                                                                admin_var, vector_cycle, start_school, status_var)





## calculate PiN
if ocha_data is not None:
    (severity_admin_status_list, dimension_admin_status_list, severity_female_list, severity_male_list, factor_category,  pin_per_admin_status, dimension_per_admin_status,
    female_pin_per_admin_status, male_pin_per_admin_status, 
    pin_per_admin_status_girl, pin_per_admin_status_boy,pin_per_admin_status_ece, pin_per_admin_status_primary, pin_per_admin_status_upper_primary, pin_per_admin_status_secondary, 
    Tot_PiN_JIAF, Tot_Dimension_JIAF, final_overview_df,final_overview_df_OCHA, 
    final_overview_dimension_df,final_overview_dimension_df_in_need,
    Tot_PiN_by_admin,
    country_label) = calculatePIN (country, edu_data_severity, household_data, choice_data, survey_data, ocha_data,mismatch_ocha_data,
                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                    age_var, gender_var,
                                                                                    label, 
                                                                                    admin_var, vector_cycle, start_school, status_var,
                                                                                    mismatch_admin)



    ocha_excel = create_output(Tot_PiN_JIAF, final_overview_df, final_overview_df_OCHA, "PiN TOTAL",  admin_var,  ocha= True, tot_severity=Tot_PiN_by_admin)
    doc_output = create_snapshot_PiN(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need)

    st.download_button(
        label=translations["download_pin"],
        data=ocha_excel.getvalue(),
        file_name=f"PiN_results_{country_label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.download_button(
        label=translations["download_word"],
        data=doc_output.getvalue(),
        file_name=f"PiN_snapshot_{country_label}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    st.subheader(translations["hno_guidelines_subheader"])
    st.markdown(translations["hno_guidelines_message"])






######################################################################### no ocha data
if no_ocha_data:
    (severity_admin_status_list, dimension_admin_status_list,
    severity_female_list, severity_male_list ,
    country_label) = calculatePIN_NO_OCHA (country, edu_data_severity, household_data, choice_data, survey_data, 
                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                    age_var, gender_var,
                                                                                    label, 
                                                                                    admin_var, vector_cycle, start_school, status_var)

    # Create an in-memory BytesIO buffer to hold the Excel file
    excel_pin = BytesIO()

    # Create an Excel writer object and write the DataFrames to it
    with pd.ExcelWriter(excel_pin, engine='xlsxwriter') as writer:
        # Iterate over each category and DataFrame in the dictionary
        for category, df in severity_admin_status_list.items():
            # Write the DataFrame to a sheet named after the category
            df.to_excel(writer, sheet_name=category, index=False)

    # Set the buffer position to the start
    excel_pin.seek(0)

    # Create a download button for the Excel file in Streamlit
    st.download_button(
        label="Download PiN percentages by admin and by population group",
        data=excel_pin,
        file_name=f"PiN_percentages_{country_label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


