import streamlit as st
import numpy as np
import pandas as pd
from backup import calculatePIN
import tempfile


st.logo('pics/logos.png')

st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico',  layout='wide')

if 'password_correct' not in st.session_state:
    st.error('Please Login from the Home page and try again.')
    st.stop()


## ====================================================================================================
## ===================================== calculate and download the PiN
## ====================================================================================================


st.write ('test session state')

st.write("Start School:", st.session_state.get('start_school'))
st.write("Vector Cycle:", st.session_state.get('vector_cycle'))
st.write("Country:", st.session_state.get('country'))
#st.write("Education Data (as dict):", st.session_state.get('edu_data').to_dict())
#st.write("Household Data (as dict):", st.session_state.get('household_data').to_dict())
st.write("Status Variable:", st.session_state.get('status_var'))
#st.write("Survey Data (as dict):", st.session_state.get('survey_data').to_dict())
#st.write("Choice Data (as dict):", st.session_state.get('choice_data').to_dict())
st.write("Label:", st.session_state.get('label'))
st.write("Age Variable:", st.session_state.get('age_var'))
st.write("Gender Variable:", st.session_state.get('gender_var'))
st.write("Access Variable:", st.session_state.get('access_var'))
st.write("Teacher Disruption Variable:", st.session_state.get('teacher_disruption_var'))
st.write("IDP Disruption Variable:", st.session_state.get('idp_disruption_var'))
st.write("Armed Disruption Variable:", st.session_state.get('armed_disruption_var'))
st.write("Barrier Variable:", st.session_state.get('barrier_var'))
st.write("Selected Severity 4 Barriers:", st.session_state.get('selected_severity_4_barriers', []))
st.write("Selected Severity 5 Barriers:", st.session_state.get('selected_severity_5_barriers', []))
st.write("Admin Variable:", st.session_state.get('admin_var'))


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












(Tot_PiN_JIAF, Tot_Dimension_JIAF, 
 final_overview_df, final_overview_dimension_df, country_label) = calculatePIN (country, edu_data, household_data, choice_data, survey_data, ocha_data,
                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                age_var, gender_var,
                label, 
                admin_var, vector_cycle, start_school, status_var)


# Function to create an Excel file and return the file path
def create_excel_file(dataframes, overview_df, overview_sheet_name):
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    with pd.ExcelWriter(temp_file.name, engine='xlsxwriter') as writer:
        overview_df.to_excel(writer, sheet_name=overview_sheet_name, index=False)
        for category, df in dataframes.items():
            sheet_name = f"{overview_sheet_name.split()[0]} -- {category}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return temp_file.name

# Create the Excel files
jiaf_excel_path = create_excel_file(Tot_PiN_JIAF, final_overview_df, "PiN TOTAL")
ocha_excel_path = create_excel_file(Tot_PiN_JIAF, final_overview_df, "PiN TOTAL")
dimension_jiaf_excel_path = create_excel_file(Tot_Dimension_JIAF, final_overview_dimension_df, "By dimension TOTAL")
dimension_ocha_excel_path = create_excel_file(Tot_Dimension_JIAF, final_overview_dimension_df, "By dimension TOTAL")

# Streamlit app layout
st.title("PiN Calculation Results")

st.download_button(
    label="Download PiN JIAF Excel",
    data=open(jiaf_excel_path, 'rb'),
    file_name=f"PiN_JIAF_{country_label}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.download_button(
    label="Download PiN OCHA Excel",
    data=open(ocha_excel_path, 'rb'),
    file_name=f"PiN_overview_OCHA_{country_label}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.download_button(
    label="Download Dimension JIAF Excel",
    data=open(dimension_jiaf_excel_path, 'rb'),
    file_name=f"Dimension_JIAF_{country_label}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.download_button(
    label="Download Dimension OCHA Excel",
    data=open(dimension_ocha_excel_path, 'rb'),
    file_name=f"Dimension_overview_OCHA_{country_label}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)