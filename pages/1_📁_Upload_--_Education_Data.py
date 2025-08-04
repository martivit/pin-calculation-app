import streamlit as st
import pandas as pd
import time
from shared_utils import language_selector
from fuzzywuzzy import process, fuzz
import numpy as np
import unicodedata


st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico',  layout='wide')
#if 'password_correct' not in st.session_state:
    #st.error(translations["no_user"])
    #st.stop()

# Call the language selector function
language_selector()
# Access the translations
translations = st.session_state.translations
st.title(translations["title_page1"])


##-------------------------------------   Variables   -----------------------------------------------------

countries = ['no selection',
    'Afghanistan -- AFG', 'Burkina Faso -- BFA', 'Cameroon -- CMR', 'Central African Republic -- CAR', 
    'Democratic Republic of the Congo -- DRC', 'Ethiopia -- ETH', 'Haiti -- HTI', 'Iraq -- IRQ', 'Lemuria -- LMR','Kenya -- KEN', 
    'Bangladesh -- BGD', 'Lebanon -- LBN', 'Moldova -- MDA', 'Mali -- MLI', 'Mozambique -- MOZ', 
    'Myanmar -- MMR', 'Niger -- NER','Nigeria -- NRA', 'Syria -- SYR', 'Ukraine -- UKR', 'Somalia -- SOM', 'South Sudan -- SSD','Sparkea -- SPR'
]
REQUIRED_COLUMNS = {
    'uuid': {'uuid', '_uuid', 'uuid_X'},
    'individual gender': {'ind_gender', 'edu_gender', 'ind_sex', 'sne_enfant_ind_genre', 'sex','edu_sex', 'edu_ind_sex', 'gender_member', 'genre'},
    'individual age': {'ind_age', 'age', 'edu_age', 'edu_ind_age', 'age'},
    'admin': {'admin1', 'admin2', 'admin3', 'camp', 'state', 'county', 'district'},
    'edu access': {'edu_access', 'enrolled_school', 'e_enfant_scolarise_formel'},
    'distruption teacher': {'edu_disrupted_teacher', 'teacher', 'e_absence_enseignant'},
    'distruption hazard': {'edu_disrupted_hazards', 'hazard', 'e_alea'},
    'distruption displaced': {'edu_disrupted_displaced', 'distrupted_idp', 'e_ecole_abris'},
    'edu barrier': {'edu_barrier', 'resn_no_access', 'e_raison_pas_educ_formel'},
    'survey start': {'start', 'date', 'start_time', 'Start_datetime', 'survey_start_date', 'today'}
}
REQUIRED_FILLED_COLUMN = [
    "Education needs stable (tick 'X' if no change from last year)",
    "Education needs decreased (tick 'X' if improved from last year)",
    "Education needs worsened (tick 'X' if worsened from last year)",
    "Reference area Admin Pcode (for extrapolation)"
]
REQUIRED_COLUMN_UPDATED_PIN = REQUIRED_FILLED_COLUMN  # as you said

FUZZY_THRESHOLD = 90  # Match similarity percentage (higher = stricter)
pin_dimensions = [
    ("a) **Access to education**", "Access to education"),
    ("b) **Learning conditions**", "Learning conditions"),
    ("c) **Protected environment**", "Protected environment"),
    ("d) **Individual protected circumstances**", "Individual protected circumstances")
]
data_sources = ["MSNA", "EMIS", "JENA"]
data_sources_individual_circumstances = ["MSNA", "JENA", "no-data"]



##-------------------------------------   functions   -----------------------------------------------------
##---------------------------------------------------------------------------------------------------------
def check_conditions_and_proceed():
    if selected_country != 'no selection' and 'uploaded_data' and 'uploaded_other_data' in st.session_state:
        if 'no_upload_ocha_data' in st.session_state or 'uploaded_ocha_data' in st.session_state:
            st.session_state.ready_to_proceed = True
        else:
            st.session_state.ready_to_proceed = False
            st.warning(translations["warning_ocha"])#Please upload the OCHA data to proceed."
    else:
        st.session_state.ready_to_proceed = False
        if selected_country == 'no selection':
            st.warning(translations["warning_missing_country"])#Please select a valid country to proceed.
        else:
            st.warning(translations["warning_MSNA_other"])#Please upload the MSNA data to proceed

    # Display success message if ready to proceed
    if st.session_state.get('ready_to_proceed', False):
        st.success("You have completed all necessary steps!")
##---------------------------------------------------------------------------------------------------------
def check_conditions_and_proceed_onlymsna():
    if selected_country != 'no selection' and 'uploaded_data'  in st.session_state:
        if 'no_upload_ocha_data' in st.session_state or 'uploaded_ocha_data' in st.session_state:
            st.session_state.ready_to_proceed = True
        else:
            st.session_state.ready_to_proceed = False
            st.warning(translations["warning_ocha"])#Please upload the OCHA data to proceed."
    else:
        st.session_state.ready_to_proceed = False
        if selected_country == 'no selection':
            st.warning(translations["warning_missing_country"])#Please select a valid country to proceed.
        else:
            st.warning(translations["warning_MSNA"])#Please upload the MSNA data to proceed

    # Display success message if ready to proceed
    if st.session_state.get('ready_to_proceed', False):
        st.success("You have completed all necessary steps!")
##---------------------------------------------------------------------------------------------------------
def validate_columns_across_sheets(all_sheets):
    """Validate mandatory columns across all sheets while excluding 'end'."""
    column_matches = {key: None for key in REQUIRED_COLUMNS}  # Track matches per key
    unmatched_columns = set(REQUIRED_COLUMNS.keys())  # Track remaining unmatched keys
    
    for sheet_name, sheet_df in all_sheets.items():
        if sheet_name.lower() in ['survey', 'choices']:
            continue  # Skip unwanted sheets

        valid_columns = [col for col in sheet_df.columns if col.lower() != 'end']

        for key, alternatives in REQUIRED_COLUMNS.items():
            if column_matches[key]:  # Skip if already matched
                continue

            alternatives_lower = {alt.lower() for alt in alternatives}

            # Step 1: Substring Matching (more flexible)
            found_column = next(
                (col for col in valid_columns if any(alt in col.lower() or col.lower() in alt for alt in alternatives_lower)),
                None
            )

            # Step 2: Fuzzy Matching (with a lower threshold for flexibility)
            if not found_column:
                best_match, score = process.extractOne(
                    key, valid_columns, scorer=fuzz.partial_ratio
                )
                if best_match and score >= 70:  # Adjusted threshold for flexibility
                    found_column = best_match
            
            if found_column:
                column_matches[key] = (sheet_name, found_column)
                unmatched_columns.discard(key)

    return column_matches, unmatched_columns
##---------------------------------------------------------------------------------------------------------
def validate_updated_pin_simple(df, required_columns):
    """
    Assumes `df` is the uploaded updated PiN DataFrame.
    Drops the first row, then:
      1. Checks that all required_columns are present in df.columns (exact match).
      2. Checks that none of those required columns are entirely empty (after dropping row 0).
    Returns: (is_valid: bool, message: str)
    """
    # Body without the header row
    df_body = df.iloc[1:].reset_index(drop=True)

    # 1. Presence check
    missing_cols = [col for col in required_columns if col not in df_body.columns]
    if missing_cols:
        template = translations.get(
            "missing_pin_columns",
            "Missing required columns: {cols}"
        )
        try:
            msg = template.format(cols=", ".join(missing_cols))
        except Exception:  # in case template is malformed
            msg = f"Missing required columns: {', '.join(missing_cols)}"
        return False, msg

    # 2. Content check: ensure required columns are not entirely empty
    empty_cols = []
    for col in required_columns:
        series = (
            df_body[col]
            .astype(str)
            .str.strip()
            .replace({"": np.nan, "nan": np.nan, "NaN": np.nan})
        )
        if series.isna().all():
            empty_cols.append(col)

    if empty_cols:
        template = translations.get(
            "empty_pin_column",
            "The following required filled columns are entirely empty: {cols}"
        )
        try:
            msg = template.format(cols=", ".join(empty_cols))
        except Exception:
            msg = f"The following required filled columns are entirely empty: {', '.join(empty_cols)}"
        return False, msg

    return True, ""

##---------------------------------------------------------------------------------------------------------
# Function to load the existing template from the file system
def load_template():
    with open('input/Template_Population_figures.xlsx', 'rb') as f:
        template = f.read()
    return template

##---------------------------------------------------------------------------------------------------------
def perform_ocha_data_checks(ocha_data):
    # Check if all mandatory columns are present
    mandatory_columns = [
        'Admin', 'ToT -- Children/Enfants (5-17)', 'ToT -- Girls/Filles (5-17)',
        'ToT -- Boys/Garcons (5-17)', '5yo -- Children/Enfants',
        'Host/Hôte -- Children/Enfants (5-17)'
    ]
    missing_columns = [col for col in mandatory_columns if col not in ocha_data.columns]
    if missing_columns:
        return f"Missing mandatory columns: {', '.join(missing_columns)}"



    # Prepare to collect error messages
    errors = []

     # Check if columns contain non-empty values
    for col in mandatory_columns:
        if ocha_data[col].isnull().all():  # If all values are NaN
            errors.append(f"Column '{col}' is empty.")
        elif ocha_data[col].isnull().sum() > 0:  # If some values are NaN
            errors.append(f"Column '{col}' contains missing values.")

    # Prepare the data by replacing NaN with 0 and rounding to the nearest whole number
    filled_ocha_data = ocha_data.fillna(0)

    # Check if the sum of values matches for specified columns
    for index, row in filled_ocha_data.iterrows():
        children_total = row['ToT -- Children/Enfants (5-17)']
        children_sum = row[['Host/Hôte -- Children/Enfants (5-17)', 'IDP/PDI -- Children/Enfants (5-17)',
                            'Returnees/Retournés -- Children/Enfants (5-17)', 'Refugees/Refugiees -- Children/Enfants (5-17)',
                            'Other -- Children/Enfants (5-17)']].sum()
        girls_boys_total = row['ToT -- Girls/Filles (5-17)'] + row['ToT -- Boys/Garcons (5-17)']

        if abs(children_total - children_sum) > 0.5:
            errors.append(f"Row {index} (Admin: {row['Admin']}): The sum of the individual population-group categories does not match 'ToT -- Children/Enfants (5-17)'")

        if abs(children_total - girls_boys_total) > 0.5:
            errors.append(f"Row {index} (Admin: {row['Admin']}): Sum of 'Girls' and 'Boys' does not match 'ToT -- Children/Enfants (5-17)'")

    if errors:
        return "\n".join(errors)  # Return all errors at once

    # If all checks pass
    return "Data is valid"



#######################################################################################################################

#------ Step 1: Country Selection
st.subheader(translations["country_section"])

selected_country = st.selectbox(
    translations["page1_country"],
    countries,
    index=countries.index(st.session_state.get('country', 'no selection'))
)
if selected_country != st.session_state.get('country'):
    st.session_state['country'] = selected_country

#------ Step 2: OCHA Data Upload
st.subheader(translations["ocha_data_section"])
no_ocha_data_checkbox = st.checkbox(f"**{translations['no_ocha_data']}**")

st.markdown(st.session_state.translations["download_template_message"], unsafe_allow_html=True) # Provide a message and download button for the OCHA data template with a highlighted section
st.download_button(label=translations["template"],# Add a download button for the existing template
                   data=load_template(),
                   file_name='Template_Population_figures.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if no_ocha_data_checkbox:
    st.session_state['no_upload_ocha_data'] = True
    upload_scope_fix_checkbox = st.checkbox("Upload Scope-Fix Sheet")
    if upload_scope_fix_checkbox:
        uploaded_scope_fix_file = st.file_uploader(translations["upload_scope_fix"], type=["xlsx"])
        if uploaded_scope_fix_file is not None:
            try:
                ocha_mismatch_data = pd.read_excel(uploaded_scope_fix_file, sheet_name='scope-fix', engine='openpyxl')
                st.session_state['ocha_mismatch_data'] = ocha_mismatch_data
                st.success("Scope-Fix sheet uploaded successfully!")
                # Work on a copy to avoid modifying the original dataset
                df_copy = ocha_mismatch_data.copy()

                # Check the first three columns of the first row
                first_row_check = (
                    df_copy.iloc[0, :3]
                    .astype(str)
                    .str.strip()
                    # Replace *all* these possible placeholders with NA
                    .replace(["", "nan", "NaN", "None"], pd.NA)
                    .notna()
                    .sum()
                )   
                if first_row_check == 0:  # Only modify the copy for the check
                    df_copy.iloc[0, :] = 1  # Modify the copy, not the original
                else:
                    df_copy.iloc[0, :] = np.nan  # Modify the copy, not the original

                # Check if scope_fix should be True
                scope_fix = first_row_check >= 2
                if scope_fix:
                    st.session_state['scope_fix'] = True

                st.success(translations["ok_upload"])
            except Exception as e:
                st.error(f"Error loading sheets: {str(e)}")  # Handle any errors, like missing sheets        
else:
    if 'no_upload_ocha_data' in st.session_state:
        del st.session_state['no_upload_ocha_data']
    
    if 'uploaded_ocha_data' in st.session_state and 'ocha_mismatch_data' in st.session_state:
        ocha_data = st.session_state['uploaded_ocha_data']
        ocha_mismatch_data = st.session_state['ocha_mismatch_data']

        # Work on a copy to avoid modifying the original data
        df_copy = ocha_mismatch_data.copy()

        second_row = df_copy.iloc[1]
        non_empty_count = second_row.iloc[:3].astype(str).str.strip().replace("", pd.NA).notna().sum()

        scope_fix = non_empty_count >= 2           
        st.write(translations["ok_upload"])
    else:
        # OCHA data uploader
        uploaded_ocha_file = st.file_uploader(translations["upload_ocha"], type=["xlsx"])
        if uploaded_ocha_file is not None:
            # Load both the 'ocha' and 'scope-fix' sheets
            try:
                ocha_data = pd.read_excel(uploaded_ocha_file, sheet_name='ocha', engine='openpyxl')
                ocha_mismatch_data = pd.read_excel(uploaded_ocha_file, sheet_name='scope-fix', engine='openpyxl')
                
                # Perform checks on the 'ocha' data
                check_message_ocha = perform_ocha_data_checks(ocha_data)
                
                if check_message_ocha == "Data is valid":
                    # Store the sheets in session state
                    st.session_state['uploaded_ocha_data'] = ocha_data
                    st.session_state['ocha_mismatch_data'] = ocha_mismatch_data

                    # Work on a copy to avoid modifying the original dataset
                    df_copy = ocha_mismatch_data.copy()
                    #st.write(ocha_mismatch_data.shape)
                    #st.write(ocha_mismatch_data.head(5))
                    #st.dataframe(df_copy)
                    # Check the first three columns of the first row
                    first_row_check = (
                        df_copy.iloc[0, :3]
                        .astype(str)
                        .str.strip()
                        # Replace *all* these possible placeholders with NA
                        .replace(["", "nan", "NaN", "None"], pd.NA)
                        .notna()
                        .sum()
                    )   
                    #st.write(first_row_check)
                    if first_row_check == 0:  # Only modify the copy for the check
                        df_copy.iloc[1, :] = 1  # Modify the copy, not the original
                    else:
                        df_copy.iloc[1, :] = np.nan  # Modify the copy, not the original

                    #st.dataframe(df_copy)

                    # Check if scope_fix should be True
                    scope_fix = first_row_check >= 2
                    if scope_fix:
                        st.session_state['scope_fix'] = True
                        st.write (scope_fix)
                    st.success(translations["ok_upload"])
                else:
                    st.error(check_message_ocha)  # Display the error message for 'ocha' sheet if checks fail
                    
            except Exception as e:
                st.error(f"Error loading sheets: {str(e)}")  # Handle any errors, like missing sheets

hybrid_scenario_countries = [
    'Central African Republic -- CAR',
    'Burkina Faso -- BFA',
    'Ethiopia -- ETH',
    'Democratic Republic of the Congo -- DRC',
    'Mali -- MLI',
    'Lebanon -- LBN',
    'Somalia -- SOM'
]
# Define the two scenario labels clearly
SCENARIO_1_LABEL = translations["SCENARIO_1"]
SCENARIO_2_LABEL = translations["SCENARIO_2"]



# Determine scenario for this country
if selected_country in hybrid_scenario_countries:
    with st.container(border=True):
        st.markdown(f"""
            <div style="margin-bottom:6px;">
                <div style="font-size:18px; font-weight:600; margin:0;">{translations["step_hpc"]}</div>
            </div>
        """, unsafe_allow_html=True)
        scenario_choice = st.radio(
            "",
            options=[SCENARIO_1_LABEL, SCENARIO_2_LABEL],
            key = 'scenario_choice'

        )
else:
    # everyone else defaults to scenario 1
    scenario_choice = SCENARIO_1_LABEL

is_scenario_2 = scenario_choice == SCENARIO_2_LABEL
st.session_state['step_2_hpc'] = False 
if is_scenario_2: st.session_state['step_2_hpc'] = True 


#----- Step 3: Select Available Data Sources
special_countries = ['Niger -- NER', 'Nigeria -- NRA']
use_full_selection = (selected_country in special_countries) and not is_scenario_2
#st.subheader(translations["select_data_section_2"])
user_selection = ""

if is_scenario_2:
    # Scenario 2: skip MSNA/dimension pills; require extrapolation input file instead
    st.subheader(translations["title_pin_upload"])
    with st.container(border=True):
        # Example: expect a structured Excel/CSV input with the extrapolation results
        uploaded_updated_2025_pin_file = st.file_uploader(translations["istruction_upload_pin"], type=["xlsx"])

    if uploaded_updated_2025_pin_file is not None:
        updated_2025_pin_file_checks = pd.read_excel(uploaded_updated_2025_pin_file, engine='openpyxl',skiprows=1)
        updated_2025_pin_file = pd.read_excel(uploaded_updated_2025_pin_file, engine='openpyxl')

        st.session_state['updated_2025_pin_file'] = updated_2025_pin_file
        try:
            df_updated = updated_2025_pin_file_checks

            valid, msg = validate_updated_pin_simple(df_updated, REQUIRED_COLUMN_UPDATED_PIN)
            if valid:
                st.session_state['updated_2025_pin_file'] = updated_2025_pin_file
                st.success(translations["extrapolation_uploaded"])
            else:
                st.error(msg)
        except Exception as e:
            st.error(f"Failed to read/validate updated PiN file: {e}")
    else:
        st.warning(translations["extrapolation_required"])

    user_selection = 'mmmm'  # keeps downstream logic consistent: treat as MSNA-only combination




elif  use_full_selection:

    st.subheader(translations["select_data_section"])

    st.markdown(
    f"""
    <span style="font-size: 18px; font-weight: bold;">
        {translations['explaination_data_dimension']}

    </span>
    """, unsafe_allow_html=True
    )

#----- Step 3.a: Select combinantion according to dimension

# Store user selections
    selections = {}
    for label, dimension in pin_dimensions:
        # Restrict options for "Individual protected circumstances" to exclude EMIS
        options = data_sources if dimension != "Individual protected circumstances" else data_sources_individual_circumstances
        
        selected_source = st.pills(
            label=f"{label} - {translations['dimension_selection']}",
            options=options,
            selection_mode="single",
            key=f"{dimension}_source"
        )
        selections[dimension] = selected_source if selected_source else "o"

    #Convert selections to a string in the correct order
    user_selection = "".join([
        "m" if selections[dim] == "MSNA" else
        "e" if selections[dim] == "EMIS" else
        "j" if selections[dim] == "JENA" else
        "n" if selections[dim] == "no-data" else "o"
        for _, dim in pin_dimensions
    ])

else:
    st.subheader(translations["select_data_section_2"])
    user_selection = 'mmmm'

st.session_state['data_combination'] = user_selection 


# Ensure all selections are made
if "o" in user_selection:
    st.warning(translations["warning_upload_data"])
else:
    # Define template mapping
    template_mapping = {
        "emmm": "Template_EMIS_Access.xlsx",
        "eemm": "Template_EMIS_Access_PTR.xlsx",
        "eeen": "Template_EMIS_All.xlsx",
        "mmem": "Template_EMIS_Access_protection.xlsx",
        "eeem": "Template_EMIS_Access_PTR_protection.xlsx",
        "meem": "Template_EMIS_PTR_protection.xlsx"

    }

    if not is_scenario_2:
        if user_selection not in template_mapping and user_selection != "mmmm"and not is_scenario_2:
            st.error(translations["error_combination"])
        else:
            template_file = template_mapping.get(user_selection, "Default_Template.xlsx")

            with st.container(border=True):
                
                if use_full_selection:
                    explanation_message = translations['explaination_mmmm'] if user_selection == "mmmm" else translations['explaination_emmm']
                    st.markdown(
                    f"""
                    <div style="background-color: #e6f7ff; padding: 1px; border-radius: 5px; border-left: 5px solid #00529B;">
                        <p style="color: #00529B; font-weight: bold; font-size: 16px;">
                            {explanation_message}
                        </p>
                    </div>
                    """, unsafe_allow_html=True
                    )
                    if user_selection != "mmmm":
                        template_file = template_mapping.get(user_selection, "Default_Template.xlsx")
                        with open(f"input/{template_file}", "rb") as file:
                            st.download_button(
                                label=translations["download_template"],
                                data=file,
                                file_name=template_file,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )



                    if user_selection == "mmmm":
                        st.subheader(translations["msna_only"])
                    
                    else:
                        st.subheader(translations["msna_other"] if "m" in user_selection else translations["other_only"])


                if "m" in user_selection and  not is_scenario_2:
                    if 'uploaded_data' in st.session_state:
                        data = st.session_state['uploaded_data']
                        st.write(translations["refresh_msna"])#MSNA Data already uploaded. If you want to change the data, just refresh 🔄 the page
                    else:
                        # MSNA data uploader
                        uploaded_file = st.file_uploader(translations["upload_msna"], type=["csv", "xlsx"])
                        if uploaded_file is not None:
                            st.write(translations["wait"])
                            bar = st.progress(0)
                            try:
                                # Load all sheets
                                all_sheets = pd.read_excel(uploaded_file, sheet_name=None, engine='openpyxl')
                                st.session_state['uploaded_data'] = all_sheets
                                bar.progress(30)

                                # Validate columns across sheets
                                column_matches, unmatched_columns = validate_columns_across_sheets(all_sheets)
                                bar.progress(60)
                                #st.success(f"✅ {translations['all_mandatory_columns_found']}") 
                                if unmatched_columns:
                                    st.error(f"### ⚠️ **{translations['missing_mandatory_columns']}**")  
                                    for col in unmatched_columns:
                                        st.write(f"- **{col}** {translations['not_found_in_sheet']}") 
                                else:
                                    st.success(f"✅ {translations['all_mandatory_columns_found']}") 
                                bar.progress(100)
                            except Exception as e:
                                st.error(f"Failed to process the uploaded file: {e}")
                                bar.progress(0)

                
                if user_selection != "mmmm" and not is_scenario_2 :
                
                    uploaded_template_file = st.file_uploader(translations["upload_other"], type=["xlsx"])
                    st.session_state['uploaded_other_data'] = uploaded_template_file

                    if uploaded_template_file is not None:
                        st.success("Processed template uploaded successfully!")




# Check conditions to allow proceeding
if is_scenario_2:
    st.session_state.ready_to_proceed = False
    # Scenario 2 requires: OCHA uploaded + updated 2025 PiN file
    has_ocha = 'uploaded_ocha_data' in st.session_state
    has_updated_pin = 'updated_2025_pin_file' in st.session_state
    tick_no_ocha = 'no_upload_ocha_data'  in st.session_state
    # Scenario 2 requires: OCHA (or explicit no-OCHA) + extrapolation input
    if has_ocha and has_updated_pin:
        st.session_state.ready_to_proceed = True
        st.success(translations["success_scenario2"])
    else:
        # Only show warnings for what's missing
        if not has_ocha and tick_no_ocha:
            st.error(translations["warning_ocha2"])
        if not has_ocha and not tick_no_ocha:
            st.warning(translations["warning_ocha"])
        if not has_updated_pin:
            st.warning(translations["extrapolation_required"])
else:
    if user_selection == 'mmmm':
        check_conditions_and_proceed_onlymsna()
    else:
        check_conditions_and_proceed()


col1, col2 = st.columns([0.60, 0.40])
label_text = st.session_state.translations["proceed_to_calculation_label"]

scope_test = st.session_state.get('scope_fix')

#if scope_test:
    #st.write("Scope-Fix sheet contains data!")
#else:
    #st.write("Scope-Fix sheet is empty!")

if not is_scenario_2:
    with col2: 
        st.page_link("pages/2_📊_Calculation_--_PiN.py", label=translations["proceed_to_calculation_label"], icon='📊')
if is_scenario_2:
    with col2: 
        st.page_link("pages/3_🧮_Update_and_Download_--_PiN_figures_and_other_outputs.py",  label=translations['to_page3_2'], icon='🧮')

