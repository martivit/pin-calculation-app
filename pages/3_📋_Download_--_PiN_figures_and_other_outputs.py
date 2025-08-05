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
from src.make_output1_platform import merge_2025_contextDB
from src.create_output1_excel import create_output1_user
from src.extrapolation import extrapolate_df_2025_updated
from src.update_re_calculation_for_PiN import UPDATE_calculatePIN

from shared_utils import language_selector
#from github import Github
import requests
import base64
from datetime import datetime
import zipfile

import os
import glob



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

if 'uploaded_data' not in st.session_state and 'uploaded_other_data' not in st.session_state and 'step_2_hpc' not in st.session_state :
    st.warning(translations["no_data"])  
    st.stop()


github_token = st.secrets["github"]["token"]


###########################################################################################################
##--------------------------------------------------------------------------------------------------------------------
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
    #if response.status_code in [200, 201]:
        # Successful creation or update
        #st.write("✅ Upload successful!")
        #return response.json()["html_url"]
    #else:
        #st.error(f"⚠️ Upload failed: {response.status_code} - {response.text}")
        #return None

##--------------------------------------------------------------------------------------------------------------------
def create_zip_file(country_label, excel_file,indicator_output,word_snapshot, word_parameters):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")  # Current timestamp
    zip_buffer = BytesIO()  # Create an in-memory ZIP file
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        # Add the Excel file with timestamp
        zip_file.writestr(f"PiN_results_{country_label}_{timestamp}.xlsx", excel_file.getvalue())
        zip_file.writestr(f"PiN_by_indicator_{country_label}_{timestamp}.xlsx", indicator_output.getvalue())
        # Add the Word Snapshot with timestamp
        zip_file.writestr(f"PiN_snapshot_{country_label}_{timestamp}.docx", word_snapshot.getvalue())
        # Add the Parameters Word Document with timestamp
        zip_file.writestr(f"Parameters_Input_Document_{timestamp}.docx", word_parameters.getvalue())
    zip_buffer.seek(0)  # Reset the buffer to the beginning
    return zip_buffer

def create_zip_file_FR(country_label, excel_file,indicator_output, word_parameters):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")  # Current timestamp
    zip_buffer = BytesIO()  # Create an in-memory ZIP file
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        # Add the Excel file with timestamp
        zip_file.writestr(f"PiN_results_{country_label}_{timestamp}.xlsx", excel_file.getvalue())
        zip_file.writestr(f"PiN_by_indicator_{country_label}_{timestamp}.xlsx", indicator_output.getvalue())
        # Add the Word Snapshot with timestamp
        #zip_file.writestr(f"PiN_snapshot_{country_label}_{timestamp}.docx", word_snapshot.getvalue())
        # Add the Parameters Word Document with timestamp
        zip_file.writestr(f"Parameters_Input_Document_{timestamp}.docx", word_parameters.getvalue())
    zip_buffer.seek(0)  # Reset the buffer to the beginning
    return zip_buffer
##--------------------------------------------------------------------------------------------------------------------
def create_zip_file_no_ocha(country_label, pin_percentage, indicator_output,word_parameters):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")  # Current timestamp
    zip_buffer = BytesIO()  # Create an in-memory ZIP file
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        zip_file.writestr(f"PiN_percentage_{country_label}_{timestamp}.xlsx", pin_percentage.getvalue())
        # Add the Excel file with timestamp
        zip_file.writestr(f"PiN_by_indicator_{country_label}_{timestamp}.xlsx", indicator_output.getvalue())
        # Add the Parameters Word Document with timestamp
        zip_file.writestr(f"Parameters_Input_Document_{timestamp}.docx", word_parameters.getvalue())
    zip_buffer.seek(0)  # Reset the buffer to the beginning
    return zip_buffer

##--------------------------------------------------------------------------------------------------------------------
def dict_of_dfs_to_bytesio_excel(dfs: dict[str, pd.DataFrame]) -> BytesIO:
    """
    Write each DataFrame in `dfs` to its own sheet in an in-memory Excel file.
    Sheet names are the dict keys, truncated to 31 chars.
    Returns a BytesIO you can feed directly to st.download_button.
    """
    output = BytesIO()
    # use openpyxl engine so you get a true .xlsx
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            safe_name = str(sheet_name)[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
    output.seek(0)
    return output






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
updated_2025_pin_file = st.session_state.get('updated_2025_pin_file') 
uploaded_covered_2025_pin = st.session_state.get('uploaded_covered_2025_pin')


# Check if the user indicated that they do not have OCHA data
no_ocha_data = st.session_state.get('no_upload_ocha_data', False)
mismatch_admin = st.session_state.get('mismatch_admin', False)


parameters = generate_parameters(st.session_state)
parameters_FR = generate_parameters_FR(st.session_state)

step_2_hpc = st.session_state.get('step_2_hpc') 



# 0.                                                          Scenario/step flags
###################################################################################################################################################
hybrid_scenario_countries = [
    'Central African Republic -- CAR',
    'Burkina Faso -- BFA',
    'Ethiopia -- ETH',
    'Democratic Republic of the Congo -- DRC',
    'Mali -- MLI',
    'Lebanon -- LBN',
    'Somalia -- SOM'
]

hybrid_country= False
if country in hybrid_scenario_countries: hybrid_country= True
step_2_hpc = st.session_state.get('step_2_hpc') 

jena_countries = ['Niger -- NER', 'Nigeria -- NRA']
jena_country = False
if country in jena_countries: jena_country = True

DATA_DIR_CONTEXT_DB = "context_DB"
DATA_DIR_PIN2024 = "pin2024_cat"


###################################################################################################################################################
###################################################################################################################################################
# 1.                                                   PiN calculation first time using MSNA
###################################################################################################################################################
if not step_2_hpc and not jena_country:

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
        (indicator_barrier4_list,indicator_barrier_list,severity_admin_status_list, dimension_admin_status_list, severity_female_list, severity_male_list, factor_category,  pin_per_admin_status, dimension_per_admin_status,indicator_per_admin_status,
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
        
        

timestamp = datetime.now().strftime("%m%d_%I%p")




###################################################################################################################################################
###################################################################################################################################################
# 2.                                                   Creation of output for FULL MSNA countries and donwnload
###################################################################################################################################################
if ocha_data is not None and not step_2_hpc and not jena_country and not hybrid_country:

    label_total_pin_sheet = "PiN TOTAL"

    # ------------------------ A. create excel PiN classic file
    if selected_language == "French":
        ocha_excel = create_output(country_label,Tot_PiN_JIAF,final_overview_df,final_overview_df_OCHA,label_total_pin_sheet,admin_var,ocha=True,tot_severity=Tot_PiN_by_admin,selected_language=selected_language,parameters=parameters_FR  )
    else:
        ocha_excel = create_output(country_label,Tot_PiN_JIAF,final_overview_df,final_overview_df_OCHA,label_total_pin_sheet,admin_var,ocha=True,tot_severity=Tot_PiN_by_admin,selected_language=selected_language,parameters=parameters  )

    # ------------------------ B. create excel PiN by indicator file
    indicator_output = create_indicator_output(country_label, indicator_per_admin_status, admin_var=admin_var)

    # ------------------------ C. create word PiN snapshot
    if selected_language == "English":
        doc_output = create_snapshot_PiN(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need, selected_language=selected_language)
        doc_parameter_output = generate_word_document(parameters)
    if selected_language == "French":
        doc_parameter_output = generate_word_document_FR(parameters_FR)
        doc_output = create_snapshot_PiN_FR(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need,selected_language=selected_language)

    # ------------------------ D. create Zip file with all important documents
    zip_file_name = f"PiN_Documents_{country_label}_{timestamp}.zip"

    if selected_language == "English":
        zip_file = create_zip_file(country_label, ocha_excel,indicator_output, doc_output, doc_parameter_output)
    if selected_language == "French":
        #zip_file = create_zip_file_FR(country_label, ocha_excel,indicator_output,  doc_parameter_output)
        zip_file = create_zip_file(country_label, ocha_excel,indicator_output, doc_output, doc_parameter_output)

    # ------------------------ E. download zip file
    st.download_button(
        label=translations["download_all"],
        data=zip_file,
        file_name=zip_file_name,
        mime="application/zip")

    # ------------------------ F. save in github --> gitpush
    if st.download_button(
        label=translations["download_all"],
        data=zip_file,
        file_name=zip_file_name,
        mime="application/zip"
    ):


        #if "github" in st.secrets and "token" in st.secrets["github"]:
            #st.write("✅ GitHub token found in secrets.")
        #else:
            #st.error("❌ GitHub token not found in secrets. Check your Streamlit configuration.")
        try:
            repo_name = "martivit/pin-calculation-app"
            branch_name = "develop_2025"

            # File paths in the repository
            file_path_in_repo_excel = f"platform_PiN_output/{country}/PiN_results_{country}_{timestamp}.xlsx"
            if selected_language == "English":
                file_path_in_repo_doc = f"platform_PiN_output/{country}/PiN_snapshot_{country}_{timestamp}.docx"
            if selected_language == "French":
                file_path_in_repo_doc = f"platform_PiN_output/{country}/PiN_parameter_{country}_{timestamp}.docx"

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
            except Exception :
                pass
                #st.error(f"Failed to upload Excel file to GitHub: {e}")

            try:
                pr_url_doc = upload_to_github(
                    file_content=doc_output.getvalue(),
                    file_name=file_path_in_repo_doc,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add PiN snapshot (Word) for {country_label}",
                    token=github_token
                )
            except Exception :
                pass
                #st.error(f"Failed to upload Word document to GitHub: {e}")

            # Display success messages only if files were successfully uploaded
            if pr_url_excel:
                st.success(f"Excel file uploaded to GitHub successfully! [View File]({pr_url_excel})")
            if pr_url_doc:
                st.success(f"Word document uploaded to GitHub successfully! [View File]({pr_url_doc})")

        except Exception :
            #st.error(f"Unexpected error during GitHub upload: {e}")
            pass

    st.subheader(translations["hno_guidelines_subheader"])
    st.markdown(translations["hno_guidelines_message"])


###################################################################################################################################################
###################################################################################################################################################
# 3.                                   First step of temporary PiN file for HYBRID MSNA countries and donwnload
###################################################################################################################################################
if ocha_data is not None and not step_2_hpc and not jena_country and hybrid_country:

    ## merge the PiN 2025 calculated for targeted areas with the secondary data (II, ACLED, clustering, additional empy columns)
    output_1_2025 = merge_2025_contextDB (country,  ocha_data, Tot_PiN_by_admin, DATA_DIR_CONTEXT_DB)   
    ## format with color and headers the output_1_2025
    formatted_output_1_2025 = create_output1_user(output_1_2025)
    output1_file_name = f"PiN_temporary_to_fill_{country_label}_{timestamp}.xlsx"
    pin_by_status_file_name = f"{country_label}_PiN_targeted_MSNA_2025_{timestamp}.xlsx"

    raw_excel = dict_of_dfs_to_bytesio_excel(Tot_PiN_JIAF)


    ## donwload
    st.download_button(
        label=translations["download_output1"],
        data=formatted_output_1_2025,
        file_name=   output1_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.download_button(
        label=translations["download_covered_area"],
        data=raw_excel,
        file_name=   pin_by_status_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

###################################################################################################################################################
###################################################################################################################################################
# 4.      step_2_hpc = TRUE         HYBRID countries, second and final download with updated PiN and extrapolation
###################################################################################################################################################

if step_2_hpc and hybrid_country:

    # --------- Locate and load the correct 2024 file from pin2024_cat ----------
    country_code = country.split("--")[-1].strip()
    pattern_2024 = os.path.join(DATA_DIR_PIN2024, f"{country_code}_PIN2024.xlsx")
    matching_files = glob.glob(pattern_2024)
    if not matching_files:
        raise FileNotFoundError(f"No 2024 PiN file found matching: {pattern_2024}")
    else:
        pin2024_file = matching_files[0]
        pin2024_cat = pd.read_excel(pin2024_file, sheet_name=None)


    st.subheader(translations["extrapolation_summary"])


    ## extrapolate using the delta method
    (merged_df_delta_all,pin_2025_updated) = extrapolate_df_2025_updated(updated_2025_pin_file, uploaded_covered_2025_pin, pin2024_cat)

    ## put together and recalculate the pin BY SEVERITY ONLY
    (Tot_PiN_JIAF,Tot_PiN_by_admin,final_overview_df_OCHA,final_overview_df, pin_per_admin_status) = UPDATE_calculatePIN (country , pin_2025_updated, ocha_data ,label,vector_cycle, selected_language )

###################################################################################################################################################
###################################################################################################################################################
# 5.                                          Creation of output for HYBRID MSNA countries and donwnload
###################################################################################################################################################

    label_total_pin_sheet = "PiN TOTAL"

    ## here --> fix creation output to not have parameters

    # ------------------------ A. create excel PiN classic file
    if selected_language == "French":
        ocha_excel = create_output(country_label,Tot_PiN_JIAF,final_overview_df,final_overview_df_OCHA,label_total_pin_sheet,admin_var,ocha=True,tot_severity=Tot_PiN_by_admin,selected_language=selected_language  )
    else:
        ocha_excel = create_output(country_label,Tot_PiN_JIAF,final_overview_df,final_overview_df_OCHA,label_total_pin_sheet,admin_var,ocha=True,tot_severity=Tot_PiN_by_admin,selected_language=selected_language )



    file_path_updated_pin = f"PiN_results_{country}_{timestamp}.xlsx"

    st.download_button(
        label=translations["download_pin"],
        data=ocha_excel,
        file_name=   file_path_updated_pin)








######################################################################### no ocha data

if no_ocha_data:
    (severity_admin_status_list, dimension_admin_status_list,
    indicator_per_admin_status,
    country_label) = calculatePIN_NO_OCHA_2025 (country, edu_data_severity, household_data, choice_data, survey_data,mismatch_ocha_data,
                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                    age_var, gender_var,
                                                                                    label, 
                                                                                    admin_var, vector_cycle, start_school, status_var,
                                                                                    mismatch_admin,
                                                                                    selected_language= selected_language)

    indicator_output = create_indicator_output_no_ocha(country_label, indicator_per_admin_status, admin_var=admin_var, selected_language=selected_language)
    pin_percentage_output = create_pin_raw_output(country_label, severity_admin_status_list, admin_var=admin_var, selected_language=selected_language)

    
    if selected_language == "English":
        doc_parameter_output = generate_word_document(parameters)

    if selected_language == "French":
        doc_parameter_output = generate_word_document_FR(parameters_FR)

    zip_file_name = f"PiN_by_indicator_Documents_{country_label}_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
    zip_file = create_zip_file_no_ocha(country_label, pin_percentage_output, indicator_output,  doc_parameter_output)

    

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
            file_path_in_repo_excel = f"platform_PiN_output/{country}/PiN_by_indicator_results_{country}_{timestamp}.xlsx"
            file_path_in_repo_excel2 = f"platform_PiN_output/{country}/PiN_percentage_{country}_{timestamp}.xlsx"

            github_token = st.secrets["github"]["token"]

            # Initialize success messages for both uploads
            pr_url_excel = None
            pr_url_doc = None

            # Try uploading both files
            try:
                pr_url_excel = upload_to_github(
                    file_content=indicator_output.getvalue(),
                    file_name=file_path_in_repo_excel,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add PiN results no ocha (Excel) for {country_label}",
                    token=github_token
                )
            except Exception :
                pass
                #st.error(f"Failed to upload Excel file to GitHub: {e}")

            try:
                pr_url_doc = upload_to_github(
                    file_content=pin_percentage_output.getvalue(),
                    file_name=file_path_in_repo_excel2,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add PiN pergentage for {country_label}",
                    token=github_token
                )
            except Exception :
                pass
                #st.error(f"Failed to upload Word document to GitHub: {e}")

            # Display success messages only if files were successfully uploaded
            if pr_url_excel:
                st.success(f"Excel file uploaded to GitHub successfully! [View File]({pr_url_excel})")
            if pr_url_doc:
                st.success(f"Word document uploaded to GitHub successfully! [View File]({pr_url_doc})")

        except Exception :
            #st.error(f"Unexpected error during GitHub upload: {e}")
            pass
 
    st.subheader(translations["hno_guidelines_subheader"])
    st.markdown(translations["hno_guidelines_message"])

