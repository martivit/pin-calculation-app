import streamlit as st
import pandas as pd
import time
from shared_utils import language_selector
from fuzzywuzzy import process, fuzz
import numpy as np


st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico',  layout='wide')


# Call the language selector function
language_selector()

# Access the translations
translations = st.session_state.translations


st.title(translations["title_page1"])


def check_conditions_and_proceed():
    if selected_country != 'no selection' and 'uploaded_data' in st.session_state:
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

#if 'password_correct' not in st.session_state:
    #st.error(translations["no_user"])
    #st.stop()


# Country selection setup
countries = ['no selection',
    'Afghanistan -- AFG', 'Burkina Faso -- BFA', 'Cameroon -- CMR', 'Central African Republic -- CAR', 
    'Democratic Republic of the Congo -- DRC', 'Haiti -- HTI', 'Iraq -- IRQ', 'Lemuria -- LMR','Kenya -- KEN', 
    'Bangladesh -- BGD', 'Lebanon -- LBN', 'Moldova -- MDA', 'Mali -- MLI', 'Mozambique -- MOZ', 
    'Myanmar -- MMR', 'Niger -- NER', 'Syria -- SYR', 'Ukraine -- UKR', 'Somalia -- SOM', 'South Sudan -- SSD'
]
selected_country = st.selectbox(
    translations["page1_country"],
    countries,
    index=countries.index(st.session_state.get('country', 'no selection'))
)
if selected_country != st.session_state.get('country'):
    st.session_state['country'] = selected_country



REQUIRED_COLUMNS = {
    'uuid': {'uuid', '_uuid', 'uuid_X'},
    'individual gender': {'ind_gender', 'edu_gender', 'ind_sex', 'sne_enfant_ind_genre', 'sex','edu_sex', 'edu_ind_sex', 'gender_member', 'genre'},
    'individual age': {'ind_age', 'age', 'edu_age', 'edu_ind_age', 'age'},
    'admin': {'admin1', 'admin2', 'admin3', 'camp', 'state', 'county', 'district'},
    'edu access': {'edu_access', 'enrolled_school'},
    'distruption teacher':{'edu_disrupted_teacher', 'teacher'},
    'distruption hazard':{'edu_disrupted_hazards', 'hazard'},
    'distruption displaced':{'edu_disrupted_displaced', 'distrupted_idp'},
    'edu barrier': {'edu_barrier', 'resn_no_access', 'e_raison_pas_educ_formel'},
    'survey start': {'start', 'date'}
}
FUZZY_THRESHOLD = 90  # Match similarity percentage (higher = stricter)

def validate_columns_across_sheets(all_sheets):
    """Validate mandatory columns across all sheets while excluding 'end'."""
    column_matches = {key: None for key in REQUIRED_COLUMNS}  # Track matches per key
    unmatched_columns = set(REQUIRED_COLUMNS.keys())  # Track remaining unmatched keys
    
    # Iterate through sheets
    for sheet_name, sheet_df in all_sheets.items():
        if sheet_name.lower() in ['survey', 'choices']:
            continue  # Skip unwanted sheets

        # Exclude 'end' column from matching
        valid_columns = [col for col in sheet_df.columns if col.lower() != 'end']


        for key, alternatives in REQUIRED_COLUMNS.items():
            if column_matches[key]:  # Skip if already matched
                continue
            
            # Step 1: Case-insensitive Exact Match
            found_column = next(
                (col for col in valid_columns if col.lower() in {alt.lower() for alt in alternatives}),
                None
            )
            
            
            # Step 2: Partial Substring Matching (Manually Check Substrings)
            if not found_column:
                found_column = next(
                    (col for col in valid_columns if any(alt in col.lower() for alt in {alt.lower() for alt in alternatives})),
                    None
                )
            
            # Step 3: Fuzzy Match as Backup
            if not found_column:
                best_match, score = process.extractOne(
                    key, valid_columns, scorer=fuzz.partial_ratio
                )
                if best_match and score >= FUZZY_THRESHOLD:
                    found_column = best_match
            
            # Record the match
            if found_column:
                column_matches[key] = (sheet_name, found_column)
                unmatched_columns.discard(key)
    
    return column_matches, unmatched_columns




#######################################################################################################################
# Check if data already uploaded and preserved in session state
if 'uploaded_data' in st.session_state:
    data = st.session_state['uploaded_data']
    st.write(translations["refresh"])#MSNA Data already uploaded. If you want to change the data, just refresh 🔄 the page
    
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

            # Display results
            #st.write("### ✅ **Column Mapping Across Sheets:**")
            #for key, match in column_matches.items():
                #if match:
                    #st.write(f"- **{key}** → Found in sheet: **{match[0]}** as column: `{match[1]}`")

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


st.markdown("---")  # Markdown horizontal rule

# Provide a message and download button for the OCHA data template with a highlighted section
st.markdown(st.session_state.translations["download_template_message"], unsafe_allow_html=True)


# Function to load the existing template from the file system
def load_template():
    with open('input/Template_Population_figures.xlsx', 'rb') as f:
        template = f.read()
    return template

# Add a download button for the existing template
st.download_button(label=translations["template"],
                   data=load_template(),
                   file_name='Template_Population_figures.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

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


# Add or modify the section where OCHA data is uploaded
if 'uploaded_data' in st.session_state:
    # Checkbox for indicating no OCHA data
    no_ocha_data_checkbox = st.checkbox(f"**{translations['no_ocha_data']}**")
    
    if no_ocha_data_checkbox:
        st.session_state['no_upload_ocha_data'] = True
    else:
        if 'no_upload_ocha_data' in st.session_state:
            del st.session_state['no_upload_ocha_data']
        
        if 'uploaded_ocha_data' in st.session_state and 'ocha_mismatch_data' in st.session_state:
            ocha_data = st.session_state['uploaded_ocha_data']
            ocha_mismatch_data = st.session_state['ocha_mismatch_data']

            second_row = ocha_mismatch_data.iloc[1]
            non_empty_count = second_row.iloc[:3].astype(str).str.strip().replace("", pd.NA).notna().sum()

            scope_fix = non_empty_count >= 2           
            st.write(translations["ok_upload"])
            st.write("OCHA Data Preview:")
            st.dataframe(ocha_data.head())  # Show a preview of the 'ocha' sheet data
            

            st.write("Scope-Fix Data Preview:")
            #st.dataframe(ocha_mismatch_data.head())  # Show a preview of the 'scope-fix' sheet data
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

                        df = pd.DataFrame(ocha_mismatch_data)
                        # Replace all non-NaN/non-None values in the second row with 1
                        for col in df.columns:
                            if pd.notna(df.at[0, col]) and df.at[0, col] != '':
                                df.at[0, col] = 1
                            else:
                                df.at[0, col] = np.nan

                        st.dataframe(df) 
                        second_row = df.iloc[0]

                        non_empty_count = second_row.notna().sum()
                        scope_fix = non_empty_count >= 2
                        if scope_fix:
                            st.session_state['scope_fix'] = True

                        
                        st.success(translations["ok_upload"])
                    else:
                        st.error(check_message_ocha)  # Display the error message for 'ocha' sheet if checks fail
                        
                except Exception as e:
                    st.error(f"Error loading sheets: {str(e)}")  # Handle any errors, like missing sheets



# Check conditions to allow proceeding
check_conditions_and_proceed()

col1, col2 = st.columns([0.60, 0.40])
label_text = st.session_state.translations["proceed_to_calculation_label"]

scope_test = st.session_state.get('scope_fix')

if scope_test:
    st.write("Scope-Fix sheet contains data!")
else:
    st.write("Scope-Fix sheet is empty!")

with col2: 
    st.page_link("pages/2_📊_Calculation_--_PiN.py", label=translations["proceed_to_calculation_label"], icon='📊')
