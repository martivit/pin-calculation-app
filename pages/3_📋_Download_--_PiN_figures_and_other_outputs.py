import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
from src.add_PiN_severity import add_severity
from src.calculation_for_PiN_Dimension import calculatePIN
from src.calculation_for_PiN_Dimension_NO_OCHA import calculatePIN_NO_OCHA
from src.vizualize_PiN import create_output
from src.snapshot_PiN import create_snapshot_PiN
from src.snapshot_PiN_FR import create_snapshot_PiN_FR
from src.save_parameter import generate_word_document
from src.save_parameter import generate_parameters
from src.save_parameter_FR import generate_word_document_FR
from src.save_parameter_FR import generate_parameters_FR
from shared_utils import language_selector
#from github import Github
import requests
import base64
from datetime import datetime
import zipfile





#from translate_PiN import translate_excel_sheets_with_formatting


st.logo('pics/GEC Global English logo_Colour_JPEG.jpg')

st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico',  layout='wide')

# Call the language selector function
language_selector()

# Access the translations
translations = st.session_state.translations
selected_language = st.session_state.get('selected_language', 'English')

#if 'password_correct' not in st.session_state:
    #st.error(translations["no_user"])
    #st.stop()

if 'uploaded_data' not in st.session_state:
    st.warning(translations["no_data"])  
    st.stop()


github_token = st.secrets["github"]["token"]

def upload_to_github(file_content, file_name, repo_name, branch_name, commit_message, token):
    """
    Uploads a file to a GitHub repository using the GitHub REST API.

    :param file_content: The binary content of the file to be uploaded.
    :param file_name: The path in the repository where the file should be uploaded.
    :param repo_name: The full name of the repository (e.g., "username/repo").
    :param branch_name: The branch to push changes to.
    :param commit_message: The commit message for the file upload.
    :param token: GitHub Personal Access Token.
    """
    # GitHub API base URL
    api_url = f"https://api.github.com/repos/{repo_name}/contents/{file_name}"

    # Encode the file content to Base64
    encoded_content = base64.b64encode(file_content).decode('utf-8')

    # Headers with the GitHub token
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    # Check if the file already exists
    response = requests.get(api_url, headers=headers)
    if response.status_code == 200:
        # File exists, update it
        sha = response.json()["sha"]
        data = {
            "message": commit_message,
            "content": encoded_content,
            "sha": sha,
            "branch": branch_name
        }
        response = requests.put(api_url, headers=headers, json=data)
    elif response.status_code == 404:
        # File does not exist, create it
        data = {
            "message": commit_message,
            "content": encoded_content,
            "branch": branch_name
        }
        response = requests.put(api_url, headers=headers, json=data)
    else:
        # Some other error
        raise Exception(f"Failed to check file existence: {response.status_code} {response.text}")

    # Handle response
    if response.status_code in [200, 201]:
        # Successful creation or update
        return response.json()["html_url"]
    else:
        raise Exception(f"Failed to upload file: {response.status_code} {response.text}")



def create_zip_file(country_label, excel_file, word_snapshot, word_parameters):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")  # Current timestamp
    zip_buffer = BytesIO()  # Create an in-memory ZIP file
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        # Add the Excel file with timestamp
        zip_file.writestr(f"PiN_results_{country_label}_{timestamp}.xlsx", excel_file.getvalue())
        # Add the Word Snapshot with timestamp
        zip_file.writestr(f"PiN_snapshot_{country_label}_{timestamp}.docx", word_snapshot.getvalue())
        # Add the Parameters Word Document with timestamp
        zip_file.writestr(f"Parameters_Input_Document_{timestamp}.docx", word_parameters.getvalue())
    zip_buffer.seek(0)  # Reset the buffer to the beginning
    return zip_buffer









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
natural_hazard_var =  st.session_state.get('natural_hazard_disruption_var')
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


parameters = generate_parameters(st.session_state)
parameters_FR = generate_parameters_FR(st.session_state)



## add indicator ---> severity
edu_data_severity = add_severity (country, edu_data, household_data, choice_data, survey_data, 
                                                                                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                                                                                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                age_var, gender_var,
                                                                                label, 
                                                                                admin_var, vector_cycle, start_school, status_var,
                                                                                selected_language)





## calculate PiN
if ocha_data is not None:
    (severity_admin_status_list, dimension_admin_status_list, severity_female_list, severity_male_list, factor_category,  pin_per_admin_status, dimension_per_admin_status,
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
                                                                                    selected_language)


    label_total_pin_sheet = "PiN TOTAL"

    #ocha_excel = create_output(country_label,Tot_PiN_JIAF, final_overview_df, final_overview_df_OCHA, label_total_pin_sheet,  admin_var,  ocha= True, tot_severity=Tot_PiN_by_admin, selected_language=selected_language, parameters=parameters)
    if selected_language == "French":
        ocha_excel = create_output(
            country_label,
            Tot_PiN_JIAF,
            final_overview_df,
            final_overview_df_OCHA,
            label_total_pin_sheet,
            admin_var,
            ocha=True,
            tot_severity=Tot_PiN_by_admin,
            selected_language=selected_language,
            parameters=parameters_FR  
        )
    else:
        ocha_excel = create_output(
            country_label,
            Tot_PiN_JIAF,
            final_overview_df,
            final_overview_df_OCHA,
            label_total_pin_sheet,
            admin_var,
            ocha=True,
            tot_severity=Tot_PiN_by_admin,
            selected_language=selected_language,
            parameters=parameters  
        )


    if selected_language == "English":
        doc_output = create_snapshot_PiN(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need, selected_language=selected_language)
        doc_parameter_output = generate_word_document(parameters)

    if selected_language == "French":
        doc_output = create_snapshot_PiN_FR(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need, selected_language=selected_language)
        doc_parameter_output = generate_word_document_FR(parameters_FR)

    zip_file_name = f"PiN_Documents_{country_label}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
    zip_file = create_zip_file(country_label, ocha_excel, doc_output, doc_parameter_output)

    # Create a single download button for the ZIP file
    if st.download_button(
        label=translations["download_all"],
        data=zip_file,
        file_name=zip_file_name,
        mime="application/zip"
    ):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            repo_name = "martivit/pin-calculation-app"
            branch_name = "develop_2025"

            # File paths in the repository
            file_path_in_repo_excel = f"platform_PiN_output/{country}/PiN_results_{country}_{timestamp}.xlsx"
            file_path_in_repo_doc = f"platform_PiN_output/{country}/PiN_snapshot_{country}_{timestamp}.docx"

            github_token = st.secrets["github"]["token"]

            # Initialize success messages for both uploads
            pr_url_excel = None
            pr_url_doc = None

            # Try uploading both files
            try:
                pr_url_excel = upload_to_github(
                    file_content=ocha_excel.getvalue(),
                    file_name=file_path_in_repo_excel,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add PiN results (Excel) for {country_label}",
                    token=github_token
                )
            except Exception as e:
                st.error(f"Failed to upload Excel file to GitHub: {e}")

            try:
                pr_url_doc = upload_to_github(
                    file_content=doc_output.getvalue(),
                    file_name=file_path_in_repo_doc,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add PiN snapshot (Word) for {country_label}",
                    token=github_token
                )
            except Exception as e:
                st.error(f"Failed to upload Word document to GitHub: {e}")

            # Display success messages only if files were successfully uploaded
            if pr_url_excel:
                st.success(f"Excel file uploaded to GitHub successfully! [View File]({pr_url_excel})")
            if pr_url_doc:
                st.success(f"Word document uploaded to GitHub successfully! [View File]({pr_url_doc})")

        except Exception:
            #st.error(f"Unexpected error during GitHub upload: {e}")
            pass
 
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




