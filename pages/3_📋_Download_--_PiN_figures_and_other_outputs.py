import streamlit as st
import numpy as np
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.cell.cell import MergedCell  # Import MergedCell
from src.clean_dataset import clean_make_dataset
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
from src.create_map_severity import make_map_severity
from src.calculation_for_PiN_Dimension_with_JENA import calculatePIN_with_JENA
from src.calculation_for_PiN_Dimension_with_EMIS import calculatePIN_with_EMIS
from urllib.parse import quote

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

if 'password_correct' not in st.session_state:
    st.error(translations["no_user"])
    st.stop()

if 'uploaded_data' not in st.session_state and 'uploaded_other_data' not in st.session_state and 'step_2_hpc' not in st.session_state :
    st.warning(translations["no_data"])  
    st.stop()


github_token = st.secrets["github"]["token"]
timestamp = datetime.now().strftime("%m%d_%I%p")


###########################################################################################################
##--------------------------------------------------------------------------------------------------------------------
def upload_to_github(file_content, file_name, repo_name, branch_name, commit_message, token):
    api_url = f"https://api.github.com/repos/{repo_name}/contents/{quote(file_name)}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    }
    encoded_content = base64.b64encode(file_content).decode("utf-8")

    # Check existence ON THE TARGET BRANCH
    r = requests.get(api_url, headers=headers, params={"ref": branch_name})
    if r.status_code not in (200, 404):
        raise Exception(f"Existence check failed: {r.status_code} {r.text}")

    data = {
        "message": commit_message,
        "content": encoded_content,
        "branch": branch_name,
    }
    if r.status_code == 200:
        data["sha"] = r.json().get("sha")

    put = requests.put(api_url, headers=headers, json=data)
    if put.status_code not in (200, 201):
        raise Exception(f"Upload failed: {put.status_code} {put.text}")

    j = put.json()
    return (j.get("content") or {}).get("html_url") or f"https://github.com/{repo_name}/blob/{branch_name}/{file_name}"

##--------------------------------------------------------------------------------------------------------------------
def create_zip_file(country_label, excel_file,indicator_output,word_snapshot, word_parameters, maps):
    zip_buffer = BytesIO()  # Create an in-memory ZIP file
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        # Add the Excel file with timestamp
        zip_file.writestr(f"PiN_results_{country_label}_{timestamp}.xlsx", excel_file.getvalue())
        zip_file.writestr(f"PiN_by_indicator_{country_label}_{timestamp}.xlsx", indicator_output.getvalue())
        # Add the Word Snapshot with timestamp
        zip_file.writestr(f"PiN_snapshot_{country_label}_{timestamp}.docx", word_snapshot.getvalue())
        # Add the Parameters Word Document with timestamp
        zip_file.writestr(f"Parameters_Input_Document_{timestamp}.docx", word_parameters.getvalue())
        for field, buf in maps.items():
            filename = f"{country_label}_{field.replace(' ', '_')}.png"
            zip_file.writestr(filename, buf.getvalue())

    zip_buffer.seek(0)  # Reset the buffer to the beginning
    return zip_buffer

def create_zip_file_FR(country_label, excel_file,indicator_output, word_parameters, maps):
    zip_buffer = BytesIO()  # Create an in-memory ZIP file
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        # Add the Excel file with timestamp
        zip_file.writestr(f"PiN_results_{country_label}_{timestamp}.xlsx", excel_file.getvalue())
        zip_file.writestr(f"PiN_by_indicator_{country_label}_{timestamp}.xlsx", indicator_output.getvalue())
        # Add the Word Snapshot with timestamp
        #zip_file.writestr(f"PiN_snapshot_{country_label}_{timestamp}.docx", word_snapshot.getvalue())
        # Add the Parameters Word Document with timestamp
        zip_file.writestr(f"Parameters_Input_Document_{timestamp}.docx", word_parameters.getvalue())
        for field, buf in maps.items():
            filename = f"{country_label}_{field.replace(' ', '_')}.png"
            zip_file.writestr(filename, buf.getvalue())
    zip_buffer.seek(0)  # Reset the buffer to the beginning
    return zip_buffer

def create_zip_file_step1_hybrid(country_label, formatted_output_1_2025, raw_excel,word_snapshot, doc_parameter_output,indicator_output, maps=None, timestamp=None):
    zip_buffer = BytesIO()  # Create an in-memory ZIP file
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        # Add the Excel file with timestamp
        zip_file.writestr(f"PiN_temporary_to_fill_{country_label}_{timestamp}.xlsx",
                          formatted_output_1_2025.getvalue())
        zip_file.writestr(f"PiN_by_indicator_{country_label}_{timestamp}.xlsx", indicator_output.getvalue())

        # Add the raw Excel file
        zip_file.writestr(f"{country_label}_PiN_targeted_MSNA_2025_{timestamp}.xlsx",
                          raw_excel.getvalue())
        # Add the Word Snapshot with timestamp
        zip_file.writestr(f"PiN_snapshot_{country_label}_{timestamp}.docx", word_snapshot.getvalue())
        # Add the Parameters Word Document
        zip_file.writestr(f"Parameters_Input_Document_{timestamp}.docx",
                          doc_parameter_output.getvalue())

        # Only add maps if provided and not empty
        if maps:
            for field, buf in maps.items():
                if buf:  # Make sure buf is not None
                    filename = f"{country_label}_{field.replace(' ', '_')}.png"
                    zip_file.writestr(filename, buf.getvalue())

    zip_buffer.seek(0)  # Reset the buffer to the beginning
    return zip_buffer

def create_zip_file_step2_hybrid(country_label, excel_file, word_snapshot, maps):
    zip_buffer = BytesIO()  # Create an in-memory ZIP file
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        # Add the Excel file with timestamp
        zip_file.writestr(f"PiN_results_{country_label}_{timestamp}.xlsx", excel_file.getvalue())
        # Add the Word Snapshot with timestamp
        zip_file.writestr(f"PiN_snapshot_{country_label}_{timestamp}.docx", word_snapshot.getvalue())
        for field, buf in maps.items():
            filename = f"{country_label}_{field.replace(' ', '_')}.png"
            zip_file.writestr(filename, buf.getvalue())

    zip_buffer.seek(0)  # Reset the buffer to the beginning
    return zip_buffer

##--------------------------------------------------------------------------------------------------------------------
def create_zip_file_no_ocha(country_label, pin_percentage, indicator_output,word_parameters):
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


def build_ocha_excel_for_repo(ocha_df: pd.DataFrame,
                              mismatch_df: pd.DataFrame | None = None,
                              mismatch_admin: bool = False) -> BytesIO:
    """
    Create an in-memory Excel with sheet 'ocha_data' and,
    if mismatch_admin is True and mismatch_df provided, a sheet 'mismatch_ocha_data'.
    """
    dfs = {"ocha_data": ocha_df}
    if mismatch_admin and mismatch_df is not None and not mismatch_df.empty:
        dfs["mismatch_ocha_data"] = mismatch_df
    return dict_of_dfs_to_bytesio_excel(dfs)




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
#natural_hazard_var =  st.session_state.get('natural_hazard_disruption_var')
natural_hazard_var = st.session_state.get('selected_disruption_natural_hazard_column')  
natural_hazard_var_sev = st.session_state.get('natural_hazard_disruption_severity')
# NOTE: this is the *column name*. For convenience you may also read:
natural_hazard_return_var = st.session_state.get('natural_hazard_disruption_var')  # same value if user confirmed
additional_indicators = st.session_state.get("additional_indicators", [])
if not st.session_state.get("additional_indicator_enable", False):
    additional_indicators = []  # force empty if toggle is OFFcustom_map = st.session_state.get('custom_indicator_mappings', {})
additional_last_var = st.session_state.get('additional_indicator_last_var')
additional_last_sev = st.session_state.get('additional_indicator_last_severity')
additional_last_dim = st.session_state.get('additional_indicator_last_dimension')
additional_2_indicators = st.session_state.get("additional_2_indicators", [])
if not st.session_state.get("additional_2_indicator_enable", False):
    additional_2_indicators = []  # force empty if toggle is OFFcustom_map = st.session_state.get('custom_indicator_mappings', {})
additional_2_last_var = st.session_state.get('additional_2_indicator_last_var')
additional_2_last_sev = st.session_state.get('additional_2_indicator_last_severity')
additional_2_last_dim = st.session_state.get('additional_2_indicator_last_dimension')

barrier_var =  st.session_state.get('barrier_var')
selected_severity_4_barriers =  st.session_state.get('selected_severity_4_barriers', [])
selected_severity_5_barriers =  st.session_state.get('selected_severity_5_barriers', [])
admin_var =  st.session_state.get('admin_var')
# Access the OCHA data if it was uploaded
ocha_data = st.session_state.get('uploaded_ocha_data')
mismatch_ocha_data = st.session_state.get('ocha_mismatch_data')
updated_2025_pin_file = st.session_state.get('updated_2025_pin_file') 
uploaded_covered_2025_pin = st.session_state.get('uploaded_covered_2025_pin')
pop_map = st.session_state.get("pop_group_value_map", {})
pop_map_ok = st.session_state.get("pop_group_value_map_confirmed", False)

host_value     = pop_map.get("host")       # REQUIRED (string)
idp_value      = pop_map.get("idp")        # optional (string or None)
returnee_value = pop_map.get("returnee")   # optional (string or None)
other_value    = pop_map.get("other")      # optional (string or None)
refugee_value = pop_map.get("refugee")      # optional (string or None)
#st.write(host_value)
#st.write(idp_value)
#st.write(returnee_value)
#st.write(other_value)

#st.write(additional_last_var)
#st.write(additional_last_sev)

#st.write(natural_hazard_var)
#st.write(natural_hazard_var_sev)
#st.write(natural_hazard_return_var)


# Check if the user indicated that they do not have OCHA data
no_ocha_data = st.session_state.get('no_upload_ocha_data', False)
mismatch_admin = st.session_state.get('mismatch_admin', False)


parameters = generate_parameters(st.session_state)
parameters_FR = generate_parameters_FR(st.session_state)

step_2_hpc = st.session_state.get('step_2_hpc') 

#jena
data_combination = st.session_state.get('data_combination') 
other_data = st.session_state.get('uploaded_other_data')


# 0.                                                          Scenario/step flags
###################################################################################################################################################
hybrid_scenario_countries = [
    'Central African Republic -- CAR',
    'Ethiopia -- ETH',
    'Democratic Republic of the Congo -- DRC',
    #'Mali -- MLI',
    'Lebanon -- LBN',
    'Somalia -- SOM',
    'South Sudan -- SSD'
]

hybrid_country= False
if country in hybrid_scenario_countries: hybrid_country= True
step_2_hpc = st.session_state.get('step_2_hpc') 

alternative_countries = ['Niger -- NER', 'Nigeria -- NRA','Mozambique -- MOZ' ]
alternative_country = False
jena_country= False
emis_country=False

dc = (data_combination or "")

alternative_country = (country in alternative_countries) and (dc != "mmmm")

if alternative_country and data_combination == 'mjjm': jena_country = True
if alternative_country and (data_combination == 'emmm' or data_combination == 'eemm' or data_combination == 'eeem'): emis_country = True


DATA_DIR_CONTEXT_DB = "context_DB"
DATA_DIR_PIN2024 = "pin2024_cat"


###################################################################################################################################################
###################################################################################################################################################
# 1.                                                   PiN calculation first time using MSNA
###################################################################################################################################################
if not step_2_hpc and not alternative_country:

    try:
        edu_data, household_data, survey_data, choice_data, messages = clean_make_dataset (country, edu_data, household_data, choice_data, survey_data, 
                                                                                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                natural_hazard_var,natural_hazard_var_sev,
                                                                                additional_last_var,additional_last_sev,
                                                                                additional_2_last_var,additional_2_last_sev,
                                                                                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                age_var, gender_var,
                                                                                label, 
                                                                                admin_var, vector_cycle, start_school, status_var,
                                                                                selected_language)

    except Exception as e:
        st.error(str(e))
        st.stop()
    warnings = getattr(messages, "warning", None) or messages.get("warning", []) if isinstance(messages, dict) else []
    infos    = getattr(messages, "info", None)    or messages.get("info", [])    if isinstance(messages, dict) else []


    for w in messages.warning:
        st.warning(w)

    if messages.info:
        with st.expander("Processing log"):
            for i in messages.info:
                st.info(i)
    
    status_var =  "pop_status_group"
    age_var = "ind_age"
    gender_var =  "ind_gender"
    barrier_var = "edu_barrier_final"

    ## add indicator ---> severity
    edu_data_severity, drop_msg = add_severity (country, edu_data, household_data, choice_data, survey_data, 
                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                    natural_hazard_var,natural_hazard_var_sev,
                                                                                    additional_last_var,additional_last_sev,
                                                                                    additional_2_last_var,additional_2_last_sev,
                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                    age_var, gender_var,
                                                                                    label, 
                                                                                    admin_var, vector_cycle, start_school, status_var,
                                                                                    selected_language)

    if (drop_msg): 
        st.warning(drop_msg)   
    #st.dataframe(edu_data_severity)
    #st.dataframe(household_data)
    #st.write (status_var)
    #st.write (age_var)
    #st.write (gender_var)
    #st.write (access_var)
    #st.write (teacher_disruption_var)
    #st.write (idp_disruption_var)
    #st.write (armed_disruption_var)
    #st.write (natural_hazard_var)
    #st.write (natural_hazard_var_sev)
    #st.write (barrier_var)
    #st.write (selected_severity_4_barriers)
    #st.write (selected_severity_5_barriers)
    #st.write (admin_var)
    #st.write (label)



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
                                                                                        admin_var, vector_cycle, start_school, status_var,host_value ,idp_value ,returnee_value ,refugee_value, other_value ,
                                                                                        mismatch_admin,
                                                                                        selected_language= selected_language, hybrid_country= hybrid_country)
        
        





###################################################################################################################################################
###################################################################################################################################################
# 2.                                                   Creation of output for FULL MSNA countries and donwnload
###################################################################################################################################################
if ocha_data is not None and not step_2_hpc and not alternative_country and not hybrid_country:

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
        doc_output = create_snapshot_PiN_FR(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need,selected_language=selected_language, step1=False)

    maps = make_map_severity(country, pin_data=Tot_PiN_by_admin,hpc_df=ocha_data)



    # ------------------------ D. create Zip file with all important documents
    zip_file_name = f"PiN_Documents_{country_label}_{timestamp}.zip"

    if selected_language == "English":
        zip_file = create_zip_file(country_label, ocha_excel,indicator_output, doc_output, doc_parameter_output, maps)
    if selected_language == "French":
        #zip_file = create_zip_file_FR(country_label, ocha_excel,indicator_output,  doc_parameter_output)
        zip_file = create_zip_file(country_label, ocha_excel,indicator_output, doc_output, doc_parameter_output, maps)


    ## re-create ocha file to save on github 
    ocha_excel_for_repo = build_ocha_excel_for_repo(
        ocha_df=ocha_data,
        mismatch_df=mismatch_ocha_data,
        mismatch_admin=mismatch_admin
    )
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
        country_slug = country.replace(" ", "_").replace("--", "_").replace("/", "_")
        file_path_in_repo_excel = f"platform_PiN_output/{country_slug}/PiN_results_{country_slug}_{timestamp}.xlsx"
        file_path_in_repo_ocha = f"platform_PiN_output/{country_slug}/ocha_figures_{country_slug}_{timestamp}.xlsx"


        try:
            repo_name = "Global-Education-Cluster-PiN/pin-calculation-app"
            branch_name = "develop_2025"

            github_token = st.secrets["github"]["token"]

            # Initialize success messages for both uploads
            pr_url_excel = None
            pr_url_ocha = None
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
                #st.error(f"Failed to upload Word document to GitHub: {e}")   
                # 
            try: 
                pr_url_ocha = upload_to_github(
                    file_content=ocha_excel_for_repo.getvalue(),  # bytes of the Excel file we just built
                    file_name=file_path_in_repo_ocha,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add uploaded OCHA figures (Excel) for {country_label}",
                    token=github_token
                ) 
            except Exception :
                pass
                #st.error(f"Failed to upload Word document to GitHub: {e}")   
                # 
            #st.success(f"Excel file uploaded to GitHub successfully! [View File]({pr_url_excel})")

        except Exception :
            #st.error(f"Unexpected error during GitHub upload: {e}")
            pass

    st.subheader(translations["hno_guidelines_subheader"])
    st.markdown(translations["hno_guidelines_message"])


###################################################################################################################################################
###################################################################################################################################################
# 3.                                   First step of temporary PiN file for HYBRID MSNA countries and donwnload
###################################################################################################################################################
if ocha_data is not None and not step_2_hpc and not alternative_country and hybrid_country:



    ## merge the PiN 2025 calculated for targeted areas with the secondary data (II, ACLED, clustering, additional empy columns)
    output_1_2025 = merge_2025_contextDB (country,  ocha_data, Tot_PiN_by_admin, DATA_DIR_CONTEXT_DB)  
    #st.dataframe(output_1_2025) 
    ## format with color and headers the output_1_2025
    formatted_output_1_2025 = create_output1_user(output_1_2025)
    output1_file_name = f"PiN_temporary_to_fill_{country_label}_{timestamp}.xlsx"
    pin_by_status_file_name = f"{country_label}_PiN_targeted_MSNA_2025_{timestamp}.xlsx"

    raw_excel = dict_of_dfs_to_bytesio_excel(Tot_PiN_JIAF)

   # ------------------------ C. create word PiN snapshot
    if selected_language == "English":
        doc_parameter_output = generate_word_document(parameters)
        doc_output = create_snapshot_PiN(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need, selected_language=selected_language)

    if selected_language == "French":
        doc_parameter_output = generate_word_document_FR(parameters_FR)
        doc_output = create_snapshot_PiN_FR(country_label, final_overview_df, final_overview_df_OCHA,final_overview_dimension_df, final_overview_dimension_df_in_need,selected_language=selected_language,step1=True)

    # ------------------------ B. create excel PiN by indicator file
    indicator_output = create_indicator_output(country_label, indicator_per_admin_status, admin_var=admin_var)

    maps_1step = make_map_severity(country, pin_data=Tot_PiN_by_admin,hpc_df=ocha_data)


    # ------------------------ D. create Zip file with all important documents
    zip_file_name = f"PiN_Temporary_{country_label}_{timestamp}.zip"

    if selected_language == "English":
        zip_file = create_zip_file_step1_hybrid(country_label,formatted_output_1_2025, raw_excel,doc_output,  doc_parameter_output,indicator_output, maps_1step, timestamp)
    if selected_language == "French":
        #zip_file = create_zip_file_FR(country_label, ocha_excel,indicator_output,  doc_parameter_output)
        zip_file = create_zip_file_step1_hybrid(country_label,formatted_output_1_2025, raw_excel,doc_output,  doc_parameter_output , indicator_output, maps_1step, timestamp)

    # ------------------------ E. download zip file
    if st.download_button(
        label=translations["download_all_temporary"],
        data=zip_file,
        file_name=zip_file_name,
        mime="application/zip", key = 'second'):

        #if "github" in st.secrets and "token" in st.secrets["github"]:
            #st.write("✅ GitHub token found in secrets.    ")
        #else:
            #st.error("❌ GitHub token not found in secrets. Check your Streamlit configuration.")
        country_slug = country.replace(" ", "_").replace("--", "_").replace("/", "_")
        file_path_in_repo_excel = f"platform_PiN_output/{country_slug}/PiN_step1_{country_slug}_{timestamp}.xlsx"
        file_path_in_repo_doc = f"platform_PiN_output/{country_slug}/Param_step1_{country_slug}_{timestamp}.docx"
        file_path_in_repo_pop = f"platform_PiN_output/{country_slug}/PiN_pop_step1_{country_slug}_{timestamp}.xlsx"
        file_path_in_repo_ocha = f"platform_PiN_output/{country_slug}/ocha_figures_{country_slug}_{timestamp}.xlsx"

            ## re-create ocha file to save on github 
        ocha_excel_for_repo = build_ocha_excel_for_repo(
            ocha_df=ocha_data,
            mismatch_df=mismatch_ocha_data,
            mismatch_admin=mismatch_admin
        )

        try:
            repo_name = "Global-Education-Cluster-PiN/pin-calculation-app"
            branch_name = "develop_2025"

            github_token = st.secrets["github"]["token"]

            # Initialize success messages for both uploads
            pr_url_excel = None
            pr_url_doc = None

            try:
                pr_url_excel = upload_to_github(
                    file_content=formatted_output_1_2025.getvalue(),
                    file_name=file_path_in_repo_excel,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add PiN step 1 hybrid for {country_label}",
                    token=github_token
                )
            except Exception :
                pass

            try: 
                pr_url_doc = upload_to_github(
                    file_content=doc_parameter_output.getvalue(),
                    file_name=file_path_in_repo_doc,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add PiN parameters step 1 hybrid for {country_label}",
                    token=github_token
                )
            except Exception :
                pass
                #st.error(f"Failed to upload Word document to GitHub: {e}")   
                # 
            try: 
                pr_url_ocha = upload_to_github(
                    file_content=ocha_excel_for_repo.getvalue(),  # bytes of the Excel file we just built
                    file_name=file_path_in_repo_ocha,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add uploaded OCHA figures (Excel) for {country_label}",
                    token=github_token
                ) 
            except Exception :
                pass
                #st.error(f"Failed to upload Word document to GitHub: {e}")   
                # 
            try: 
                pr_url_doc = upload_to_github(
                    file_content=raw_excel.getvalue(),
                    file_name=file_path_in_repo_pop,
                    repo_name=repo_name,
                    branch_name=branch_name,
                    commit_message=f"Add PiN by pop, step 1 hybrid  for {country_label}",
                    token=github_token
                )
            except Exception :
                pass
                #st.error(f"Failed to upload Word document to GitHub: {e}")       

        except Exception :
            #st.error(f"Unexpected error during GitHub upload: {e}")
            pass


###################################################################################################################################################
###################################################################################################################################################
# 4.      step_2_hpc = TRUE         HYBRID countries, second and final download with updated PiN and extrapolation
###################################################################################################################################################

if step_2_hpc and hybrid_country:
    country_label = country.replace(" ", "_").replace("--", "_").replace("/", "_")

    # --------- Locate and load the correct 2024 file from pin2024_cat ----------
    country_code = country.split("--")[-1].strip()
    pattern_2024 = os.path.join(DATA_DIR_PIN2024, f"{country_code}_PIN2024.xlsx")
    matching_files = glob.glob(pattern_2024)
    if not matching_files:
        raise FileNotFoundError(f"No 2024 PiN file found matching: {pattern_2024}")
    else:
        pin2024_file = matching_files[0]
        pin2024_cat = pd.read_excel(pin2024_file, sheet_name=None)


    with st.container(border=True):

        st.markdown(translations["extrapolation_summary"])


    ## extrapolate using the delta method
    (merged_df_delta_all,pin_2025_updated) = extrapolate_df_2025_updated(updated_2025_pin_file, uploaded_covered_2025_pin, pin2024_cat)

    ## put together and recalculate the pin BY SEVERITY ONLY
    (Tot_PiN_JIAF,Tot_PiN_by_admin,final_overview_df_OCHA,final_overview_df, pin_per_admin_status) = UPDATE_calculatePIN (country , pin_2025_updated, ocha_data ,label, selected_language )

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

    if selected_language == "English":
        doc_output = create_snapshot_PiN(country_label, final_overview_df, final_overview_df_OCHA, selected_language=selected_language)
    if selected_language == "French":
        doc_output = create_snapshot_PiN_FR(country_label, final_overview_df, final_overview_df_OCHA,selected_language=selected_language, step1=False)




    maps_2step = make_map_severity(country, pin_data=Tot_PiN_by_admin, hpc_df=ocha_data)



    # ------------------------ D. create Zip file with all important documents
    zip_file_name = f"PiN_Documents_{country_label}_{timestamp}.zip"

    if selected_language == "English":
        zip_file = create_zip_file_step2_hybrid(country_label, ocha_excel, doc_output, maps_2step)
    if selected_language == "French":
        #zip_file = create_zip_file_FR(country_label, ocha_excel,indicator_output,  doc_parameter_output)
        zip_file = create_zip_file_step2_hybrid(country_label, ocha_excel, doc_output, maps_2step)

    # ------------------------ E. download zip file
    if st.download_button(
        label=translations["download_all"],
        data=zip_file,
        file_name=zip_file_name,
        mime="application/zip", key = 'third'):

        #if "github" in st.secrets and "token" in st.secrets["github"]:
            #st.write("✅ GitHub token found in secrets.")
        #else:
            #st.error("❌ GitHub token not found in secrets. Check your Streamlit configuration.")
        country_slug = country.replace(" ", "_").replace("--", "_").replace("/", "_")
        file_path_in_repo_excel = f"platform_PiN_output/{country_slug}/PiN_step2_{country_slug}_{timestamp}.xlsx"


        try:
            repo_name = "Global-Education-Cluster-PiN/pin-calculation-app"
            branch_name = "develop_2025"

            github_token = st.secrets["github"]["token"]

            # Initialize success messages for both uploads
            pr_url_excel = None
            pr_url_doc = None

            pr_url_excel = upload_to_github(
                file_content=ocha_excel.getvalue(),
                file_name=file_path_in_repo_excel,
                repo_name=repo_name,
                branch_name=branch_name,
                commit_message=f"Add PiN results step2 hybrid for {country_label}",
                token=github_token
            )
            #st.success(f"Excel file uploaded to GitHub successfully! [View File]({pr_url_excel})")

        except Exception :
            #st.error(f"Unexpected error during GitHub upload: {e}")
            pass



    st.subheader(translations["hno_guidelines_subheader"])
    st.markdown(translations["hno_guidelines_message"])


###################################################################################################################################################
###################################################################################################################################################
# 6.                                                   PiN calculation  and output with JENA countries
###################################################################################################################################################
if jena_country and ocha_data is not None:
    country_label = country.replace(" ", "_").replace("--", "_").replace("/", "_")

    if 'm' in data_combination:

        try:
            edu_data, household_data, survey_data, choice_data, messages = clean_make_dataset (country, edu_data, household_data, choice_data, survey_data, 
                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                    natural_hazard_var,natural_hazard_var_sev,
                                                                                    additional_last_var,additional_last_sev,
                                                                                    additional_2_last_var,additional_2_last_sev,
                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                    age_var, gender_var,
                                                                                    label, 
                                                                                    admin_var, vector_cycle, start_school, status_var,
                                                                                    selected_language)

        except Exception as e:
            st.error(str(e))
            st.stop()
        warnings = getattr(messages, "warning", None) or messages.get("warning", []) if isinstance(messages, dict) else []
        infos    = getattr(messages, "info", None)    or messages.get("info", [])    if isinstance(messages, dict) else []


        for w in messages.warning:
            st.warning(w)

        if messages.info:
            with st.expander("Processing log"):
                for i in messages.info:
                    st.info(i)
        
        status_var =  "pop_status_group"
        age_var = "ind_age"
        gender_var =  "ind_gender"
        barrier_var = "edu_barrier_final"

        ## add indicator ---> severity
        edu_data_severity, drop_msg = add_severity (country, edu_data, household_data, choice_data, survey_data, 
                                                                                        access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                        natural_hazard_var,natural_hazard_var_sev,
                                                                                        additional_last_var,additional_last_sev,
                                                                                        additional_2_last_var,additional_2_last_sev,
                                                                                        barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                        age_var, gender_var,
                                                                                        label, 
                                                                                        admin_var, vector_cycle, start_school, status_var,
                                                                                        selected_language)

        if (drop_msg): 
            st.warning(drop_msg)   


        (jena_df, merged_ocha_jena, merged_ocha_jena_msna,pin_jena_msna, Tot_PiN_JIAF,Tot_Dimension_JIAF,
          final_overview_df_OCHA, final_overview_df, Tot_PiN_by_admin)=  calculatePIN_with_JENA (data_combination,
                                                                                                  country, edu_data_severity, household_data, choice_data, survey_data, ocha_data,mismatch_ocha_data,  other_data,
                                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                                    age_var, gender_var,
                                                                                                    label, 
                                                                                                    admin_var, vector_cycle, start_school, status_var,host_value ,idp_value ,returnee_value ,refugee_value, other_value ,
                                                                                                    mismatch_admin,
                                                                                                    selected_language)




    label_total_pin_sheet = "PiN TOTAL"

    ## here --> fix creation output to not have parameters

    # ------------------------ A. create excel PiN classic file
    if selected_language == "French":
        ocha_excel = create_output(country_label,Tot_PiN_JIAF,final_overview_df,final_overview_df_OCHA,label_total_pin_sheet,admin_var,ocha=True,tot_severity=Tot_PiN_by_admin,selected_language=selected_language  )
    else:
        ocha_excel = create_output(country_label,Tot_PiN_JIAF,final_overview_df,final_overview_df_OCHA,label_total_pin_sheet,admin_var,ocha=True,tot_severity=Tot_PiN_by_admin,selected_language=selected_language )


    #st.dataframe(final_overview_df)

    if selected_language == "English":
        doc_output = create_snapshot_PiN(country_label, final_overview_df, final_overview_df_OCHA, selected_language=selected_language)
    if selected_language == "French":
        doc_output = create_snapshot_PiN_FR(country_label, final_overview_df, final_overview_df_OCHA,selected_language=selected_language,step1=False)

    
    maps_jena = make_map_severity(country, pin_data=Tot_PiN_by_admin, hpc_df=ocha_data)



    # ------------------------ D. create Zip file with all important documents
    zip_file_name_jena = f"PiN_Documents_{country_label}_{timestamp}.zip"

    if selected_language == "English":
        zip_file_jena = create_zip_file_step2_hybrid(country_label, ocha_excel, doc_output, maps_jena)
    if selected_language == "French":
        #zip_file = create_zip_file_FR(country_label, ocha_excel,indicator_output,  doc_parameter_output)
        zip_file_jena = create_zip_file_step2_hybrid(country_label, ocha_excel, doc_output, maps_jena)

    # ------------------------ E. download zip file
    if st.download_button(
        label=translations["download_all"],
        data=zip_file_jena,
        file_name=zip_file_name_jena,
        mime="application/zip", key = 'third'):

        
        #if "github" in st.secrets and "token" in st.secrets["github"]:
            #st.write("✅ GitHub token found in secrets.")
        #else:
            #st.error("❌ GitHub token not found in secrets. Check your Streamlit configuration.")
        country_slug = country.replace(" ", "_").replace("--", "_").replace("/", "_")
        file_path_in_repo_excel = f"platform_PiN_output/{country_slug}/PiN_results_{country_slug}_{timestamp}.xlsx"


        try:
            repo_name = "Global-Education-Cluster-PiN/pin-calculation-app"
            branch_name = "develop_2025"

            github_token = st.secrets["github"]["token"]

            # Initialize success messages for both uploads
            pr_url_excel = None

            pr_url_excel = upload_to_github(
                file_content=ocha_excel.getvalue(),
                file_name=file_path_in_repo_excel,
                repo_name=repo_name,
                branch_name=branch_name,
                commit_message=f"Add PiN results jena for {country_label}",
                token=github_token
            )
            #st.success(f"Excel file uploaded to GitHub successfully! [View File]({pr_url_excel})")

        except Exception :
            #st.error(f"Unexpected error during GitHub upload: {e}")
            pass




###################################################################################################################################################
###################################################################################################################################################
# 7.                                                   PiN calculation  and output with EMIS countries
###################################################################################################################################################
if emis_country and ocha_data is not None:
    country_label = country.replace(" ", "_").replace("--", "_").replace("/", "_")

    if 'm' in data_combination:  

        try:
            edu_data, household_data, survey_data, choice_data, messages = clean_make_dataset (country, edu_data, household_data, choice_data, survey_data, 
                                                                                    access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                    natural_hazard_var,natural_hazard_var_sev,
                                                                                    additional_last_var,additional_last_sev,
                                                                                    additional_2_last_var,additional_2_last_sev,
                                                                                    barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                    age_var, gender_var,
                                                                                    label, 
                                                                                    admin_var, vector_cycle, start_school, status_var,
                                                                                    selected_language)

        except Exception as e:
            st.error(str(e))
            st.stop()
        warnings = getattr(messages, "warning", None) or messages.get("warning", []) if isinstance(messages, dict) else []
        infos    = getattr(messages, "info", None)    or messages.get("info", [])    if isinstance(messages, dict) else []


        for w in messages.warning:
            st.warning(w)

        if messages.info:
            with st.expander("Processing log"):
                for i in messages.info:
                    st.info(i)
        
        status_var =  "pop_status_group"
        age_var = "ind_age"
        gender_var =  "ind_gender"
        barrier_var = "edu_barrier_final"

        ## add indicator ---> severity
        edu_data_severity, drop_msg = add_severity (country, edu_data, household_data, choice_data, survey_data, 
                                                                                        access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,
                                                                                        natural_hazard_var,natural_hazard_var_sev,
                                                                                        additional_last_var,additional_last_sev,
                                                                                        additional_2_last_var,additional_2_last_sev,
                                                                                        barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                        age_var, gender_var,
                                                                                        label, 
                                                                                        admin_var, vector_cycle, start_school, status_var,
                                                                                        selected_language)

        if (drop_msg): 
            st.warning(drop_msg)   



        (pin_by_indicator_status_list, enrollment_df, pop_figures_E_OoS_by_pop_group, severity_by_pop_group, 
            pin_by_pop_group, 
            pin_by_dimension_in_need_pop_group,pin_by_indicator_pop_group, test_intermediate_step,
            Tot_PiN_JIAF, final_overview_df_OCHA, final_overview_df, Tot_PiN_by_admin)=  calculatePIN_with_EMIS (data_combination, country, edu_data, household_data, choice_data, survey_data, ocha_data,mismatch_ocha_data,other_data,
                                                                                                access_var, teacher_disruption_var, idp_disruption_var, armed_disruption_var,natural_hazard_var,
                                                                                                barrier_var, selected_severity_4_barriers, selected_severity_5_barriers,
                                                                                                age_var, gender_var,
                                                                                                label, 
                                                                                                admin_var, vector_cycle, start_school, status_var,host_value ,idp_value ,returnee_value ,refugee_value, other_value ,
                                                                                                mismatch_admin,
                                                                                                selected_language)




    label_total_pin_sheet = "PiN TOTAL"

    ## here --> fix creation output to not have parameters

    # ------------------------ A. create excel PiN classic file
    if selected_language == "French":
        ocha_excel = create_output(country_label,Tot_PiN_JIAF,final_overview_df,final_overview_df_OCHA,label_total_pin_sheet,admin_var,ocha=True,tot_severity=Tot_PiN_by_admin,selected_language=selected_language  )
    else:
        ocha_excel = create_output(country_label,Tot_PiN_JIAF,final_overview_df,final_overview_df_OCHA,label_total_pin_sheet,admin_var,ocha=True,tot_severity=Tot_PiN_by_admin,selected_language=selected_language )


    #st.dataframe(final_overview_df)

    if selected_language == "English":
        doc_output = create_snapshot_PiN(country_label, final_overview_df, final_overview_df_OCHA, selected_language=selected_language)
    if selected_language == "French":
        doc_output = create_snapshot_PiN_FR(country_label, final_overview_df, final_overview_df_OCHA,selected_language=selected_language,step1=False)

    
    maps_emis = make_map_severity(country, pin_data=Tot_PiN_by_admin, hpc_df=ocha_data)



    # ------------------------ D. create Zip file with all important documents
    zip_file_name_emis = f"PiN_Documents_{country_label}_{timestamp}.zip"

    if selected_language == "English":
        zip_file_emis = create_zip_file_step2_hybrid(country_label, ocha_excel, doc_output, maps_emis)
    if selected_language == "French":
        #zip_file = create_zip_file_FR(country_label, ocha_excel,indicator_output,  doc_parameter_output)
        zip_file_emis = create_zip_file_step2_hybrid(country_label, ocha_excel, doc_output, maps_emis)

    # ------------------------ E. download zip file
    if st.download_button(
        label=translations["download_all"],
        data=zip_file_emis,
        file_name=zip_file_name_emis,
        mime="application/zip", key = 'third'):

        
        #if "github" in st.secrets and "token" in st.secrets["github"]:
            #st.write("✅ GitHub token found in secrets.")
        #else:
            #st.error("❌ GitHub token not found in secrets. Check your Streamlit configuration.")
        country_slug = country.replace(" ", "_").replace("--", "_").replace("/", "_")
        file_path_in_repo_excel = f"platform_PiN_output/{country_slug}/PiN_results_{country_slug}_{timestamp}.xlsx"


        try:
            repo_name = "Global-Education-Cluster-PiN/pin-calculation-app"
            branch_name = "develop_2025"

            github_token = st.secrets["github"]["token"]

            # Initialize success messages for both uploads
            pr_url_excel = None
            pr_url_doc = None

            pr_url_excel = upload_to_github(
                file_content=ocha_excel.getvalue(),
                file_name=file_path_in_repo_excel,
                repo_name=repo_name,
                branch_name=branch_name,
                commit_message=f"Add PiN results EMIS for {country_label}",
                token=github_token
            )
            #st.success(f"Excel file uploaded to GitHub successfully! [View File]({pr_url_excel})")

        except Exception :
            #st.error(f"Unexpected error during GitHub upload: {e}")
            pass







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

