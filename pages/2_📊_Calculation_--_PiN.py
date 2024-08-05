import streamlit as st
import pandas as pd
import extra_streamlit_components as stx

st.logo('pics/logos.png')

st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico', layout='wide')
st.title('Indicator and Severity categorization')


if 'password_correct' not in st.session_state:
    st.error('Please Login from the Home page and try again.')
    st.stop()

if 'current_step' not in st.session_state:
    st.session_state['current_step'] = 0

# Define the steps for the stepper bar
steps = ["Select correct data frame and language", "Select education indicators", "Define aggravating circumstances", "Define disaggregation variables"]
current_step = st.session_state['current_step']
new_step = stx.stepper_bar(steps=steps)

if new_step is not None and new_step != st.session_state['current_step']:
    st.session_state['current_step'] = new_step


if 'init' not in st.session_state:
    st.session_state.update({
        'init': True,
        'data_selections_confirmed': False,
        'label_selected': False,
        'age_column_confirmed': False,
        'gender_column_confirmed': False,
        'education_access_column_confirmed': False,
        'disruption_teacher_column_confirmed': False,
        'disruption_idp_column_confirmed': False,
        'disruption_armed_column_confirmed': False,
        'barriers_column_confirmed': False,
        'indicator_confirmed': False,
        'severity_4_confirmed': False,
        'severity_5_confirmed': False,
        'selected_barriers': [],  # List to store user-selected barriers
        'admin_level_confirmed': False,
        'school_start_month_confirmed': False,
        'school_cycle_confirmed': False,
        'other_parameters_confirmed': False
        
    })
if 'lower_primary_end' not in st.session_state:
    st.session_state['lower_primary_end'] = 11  # Default end age for lower primary
if 'upper_primary_end' not in st.session_state:
    st.session_state['upper_primary_end'] = 16  # Default end age for upper primary
if 'vector_cycle' not in st.session_state:
    st.session_state['vector_cycle'] = [11,16]  # vector


## INFO ADMIN PER COUNTRY ##
admin_levels_per_country = {
    'Afghanistan -- AFG': ['Admin_1: Province', 'Admin_2: District', 'Admin_3: Subdistrict'],
    'Burkina Faso -- BFA': ['Admin_1: Regions (Région)', 'Admin_2: Province', 'Admin_3: Department (Département)'],
    'Central African Republic -- CAR': ['Admin_1: Prefectures (préfectures)', 'Admin_2: Sub-prefectures (sous-préfectures)', 'Admin_3: Communes'],
    'Democratic Republic of the Congo -- DRC': ['Admin_1: Provinces', 'Admin_2: Territories', 'Admin_3: Sectors/chiefdoms/communes'],
    'Haiti -- HTI': ['Admin_1: Departments (départements)', 'Admin_2: Arrondissements', 'Admin_3: Communes'],
    'Iraq -- IRQ': ['Admin_1: Governorates', 'Admin_2: Districts (aqḍyat)', 'Admin_3: Sub-districts (naḥiyat)'],
    'Kenya -- KEN': ['Admin_1: Counties', 'Admin_2: Sub-counties (kaunti ndogo)', 'Admin_3: Wards (mtaa)'],
    'Bangladesh -- BGD': ['Admin_1: Divisions (bibhag)', 'Admin_2: Districts (zila)', 'Admin_3: Upazilas'],
    'Lebanon -- LBN': ['Admin_1: Governorates', 'Admin_2: Districts (qaḍya)', 'Admin_3: Municipalities'],
    'Moldova -- MDA': ['Admin_1: Districts', 'Admin_2: Cities', 'Admin_3: Communes'],
    'Mali -- MLI': ['Admin_1: Régions', 'Admin_2: Cercles', 'Admin_3: Arrondissements'],
    'Mozambique -- MOZ': ['Admin_1: Provinces (provincias)', 'Admin_2: Districts (distritos)', 'Admin_3: Postos'],
    'Myanmar -- MMR': ['Admin_1: States/Regions', 'Admin_2: Districts', 'Admin_3: Townships'],
    'Niger -- NER': ['Admin_1: Régions ', 'Admin_2: Départements', 'Admin_3: Communes'],
    'Syria -- SYR': ['Admin_1: Governorates', 'Admin_2: Districts (mintaqah)', 'Admin_3: Subdistricts (nawaḥi)'],
    'Ukraine -- UKR': ['Admin_1: Oblasts', 'Admin_2: Raions', 'Admin_3: Hromadas'],
    'Somalia -- SOM': ['Admin_1: States', 'Admin_2: Regions', 'Admin_3: Districts']
}



##---------------------------------------------------------------------------------------------------------
# Function to display status indicators
def display_status(description, status):
    color = 'green' if status else 'gray'
    st.markdown(f"<span style='color: {color}; font-size: 20px; margin-right: 5px;'>●</span> {description}", unsafe_allow_html=True)
##---------------------------------------------------------------------------------------------------------
def handle_full_selection(suggestions, column_type, custom_message):
    # Construct the full message with custom formatting
    full_message = f"Please select the variable required for measuring the <span style='color: #014bb4;'><strong>{custom_message}</strong></span>"
    st.markdown(full_message, unsafe_allow_html=True)

    # Use a key that ensures the selectbox is unique and not accidentally re-used
    select_key = f'{column_type}_selectbox'
    selected_column = st.selectbox(
        "Choose one:",
        ['No selection'] + suggestions,
        key=select_key
    )

    # Add a confirmation button
    if st.button(f"Confirm {column_type.replace('_', ' ').capitalize()}", key=f'confirm_{column_type}'):
        if selected_column != 'No selection':
            st.session_state[f'selected_{column_type}_column'] = selected_column
            st.session_state[f'{column_type}_column_confirmed'] = True
            st.success(f"{column_type.capitalize()} column '{selected_column}' has been manually selected.")
        else:
            st.error(f"Please select a valid option for {column_type.replace('_', ' ').capitalize()} before confirming.")

    # Update the status indicator after checking the session state
    display_status(f"{column_type.capitalize()} Column Confirmed", st.session_state.get(f'{column_type}_column_confirmed', False))

    # Return the selected column for further processing
    return selected_column

##---------------------------------------------------------------------------------------------------------
def handle_column_selection(suggestions, column_type):
    suggested_column = suggestions[0] if suggestions else 'No selection'
    st.write(f"Is this the individual {column_type} column? → **{suggested_column}**")
    col1, col2 = st.columns(2)
    confirm_key = f'confirm_yes_{column_type}'
    select_key = f'{column_type}_selectbox'
    edu_data = st.session_state['edu_data']
    message_placeholder = st.empty()  # Place to show messages dynamically

    with col1:
        if st.button("Yes", key=confirm_key):
            if suggested_column != 'No selection':
                st.session_state[f'selected_{column_type}_column'] = suggested_column
                st.session_state[f'{column_type}_column_confirmed'] = True
                #st.success(f"{column_type.capitalize()} column '{suggested_column}' has been confirmed.")                                
                message_placeholder.success(f"{column_type.capitalize()} column '{suggested_column}' has been confirmed.")

    with col2:
        if st.button("No", key=f'confirm_no_{column_type}'):
            selected_column = st.selectbox(
                f"Select the individual {column_type} column:",
                ['No selection'] + edu_data.columns.tolist(),
                key=select_key,
                #on_change=update_column_confirmation (column_type,message_placeholder)
                #args=(column_type,)
            )
            update_column_confirmation (column_type,message_placeholder)
##---------------------------------------------------------------------------------------------------------
def update_column_confirmation(column_type, placeholder):
    select_key = f'{column_type}_selectbox'
    selected_column = st.session_state.get(select_key)
    if selected_column and selected_column != 'No selection':
        st.session_state[f'selected_{column_type}_column'] = selected_column
        st.session_state[f'{column_type}_column_confirmed'] = True
        placeholder.success(f"{column_type.capitalize()} column '{selected_column}' has been manually selected.")
        #display_status(f"{column_type.capitalize()} Column Confirmed", st.session_state[f'{column_type}_column_confirmed'])

##---------------------------------------------------------------------------------------------------------
def update_combined_indicator():
    """Update the combined indicator status based on individual confirmations."""
    if all([
        st.session_state.get('education_access_column_confirmed', False),
        st.session_state.get('disruption_teacher_column_confirmed', False),
        st.session_state.get('disruption_idp_column_confirmed', False),
        st.session_state.get('disruption_armed_column_confirmed', False),
        st.session_state.get('barriers_column_confirmed', False)
    ]):
        st.session_state.indicators_confirmed = True
    else:
        st.session_state.indicators_confirmed = False

    display_status("Indicator Selection Confirmed", st.session_state.indicators_confirmed)
##---------------------------------------------------------------------------------------------------------
def find_barrier_details(barrier_variable, survey_data, choices_data, label_column):
    """
    Fetch all barriers for a given type from choices_data.
    """
    type_info = survey_data[survey_data['name'] == barrier_variable].iloc[0]['type']
    type_barrier = type_info.replace('select_one ', '')
    barrier_details = choices_data[choices_data['list_name'] == type_barrier]
    return barrier_details[label_column].tolist()
##---------------------------------------------------------------------------------------------------------
def show_barrier_selection(barrier_details, label_column):
    st.write("Select aggravating circumstances falling under **severity 4**:")
    selected_barriers = []
    for index, row in barrier_details.iterrows():
        if st.checkbox(f"{row[label_column]}", key=f"select_{row['name']}"):
            selected_barriers.append(row['name'])
    st.session_state.selected_barriers = selected_barriers  # Update session state
##---------------------------------------------------------------------------------------------------------
def select_severity_barriers(barrier_options, severity):
    """
    Allow the user to select barriers that correspond to a specified severity.
    """
    confirm_button_label = f"Confirm Severity {severity} Barriers"
    prompt_message = f"Select aggravating circumstances falling under **severity {severity}**, you can select multiple choices:"

    # Add special option for severity 5
    if severity == 5:
        barrier_options = barrier_options + ['---> None of the listed barriers <---']

    selected_barriers = st.multiselect(
            prompt_message,
            barrier_options,
            [])  # Start with no pre-selected barriers

    if st.button(confirm_button_label):
        if severity == 4:
            st.session_state.selected_severity_4_barriers = selected_barriers
            st.session_state['severity_4_confirmed'] = True

        elif severity == 5:
            st.session_state.selected_severity_5_barriers = selected_barriers
            st.session_state['severity_5_confirmed'] = True
            # Handle the special case when 'None of the listed barriers' is selected
            if '---> None of the listed barriers <---' in selected_barriers:
                selected_barriers = ['---> None of the listed barriers <---']
                st.session_state.selected_severity_5_barriers = selected_barriers

        st.success(f"Selected severity {severity} barriers have been confirmed.")
        st.write(f"Confirmed severity {severity} barriers:", selected_barriers)

    return selected_barriers
##---------------------------------------------------------------------------------------------------------
def display_combined_severity_status():
    severity_4_status = st.session_state.get('severity_4_confirmed', False)
    severity_5_status = st.session_state.get('severity_5_confirmed', False)
    all_severities_confirmed = severity_4_status and severity_5_status

    if all_severities_confirmed:
        status_description = "Selection of Aggravating Circumstances Confirmed"
        color = 'green'
    else:
        status_description = "Selection of Aggravating Circumstances Not Confirmed"
        color = 'gray'
    
    st.markdown(f"<span style='color: {color}; font-size: 20px; margin-right: 5px;'>●</span> {status_description}", unsafe_allow_html=True)
##---------------------------------------------------------------------------------------------------------
def check_for_duplicate_selections():
    selections = [
        st.session_state.get('selected_education_access_column'),
        st.session_state.get('selected_disruption_teacher_column'),
        st.session_state.get('selected_disruption_idp_column'),
        st.session_state.get('selected_disruption_armed_column'),
        st.session_state.get('selected_barriers_column')
    ]
    # Remove any 'None' or 'No selection' entries
    filtered_selections = [s for s in selections if s and s != 'No selection']
    
    # Check for duplicates
    if len(set(filtered_selections)) != len(filtered_selections):
        st.error("**Duplicate selections detected. Each variable should be used for only one category. Please adjust your selections.**")
##---------------------------------------------------------------------------------------------------------
def update_other_parameters_status():
    if (st.session_state.get('admin_level_confirmed', False) and
        st.session_state.get('school_start_month_confirmed', False) and
        (st.session_state.get('single_cycle', False) or st.session_state.get('upper_primary_end_confirmed', False))):
        st.session_state['other_parameters_confirmed'] = True
    else:
        st.session_state['other_parameters_confirmed'] = False

    display_status("Other Parameters Confirmed", st.session_state['other_parameters_confirmed'])

##---------------------------------------------------------------------------------------------------------
def show_step_content(step):
    if step == 0:
        st.write("### Step 1: Select Correct Data Frame")
        # Placeholder for data frame selection logic
        st.write("Here you would include your UI for selecting and validating the data frame.")
    elif step == 1:
        st.write("### Step 2: Select Education Indicators")
        # Placeholder for education indicators selection
        st.write("This is where users would pick and validate education indicators.")
    elif step == 2:
        st.write("### Step 3: Define Aggravating Circumstances")
        # Placeholder for defining aggravating circumstances
        st.write("Users would define the severity of barriers here.")
    elif step == 3:
        st.write("### Step 4: Define Disaggregation Variable")
        # Placeholder for disaggregation variables selection
        st.write("Setup for disaggregation by variables like administrative area, school cycle, etc.")

##---------------------------------------------------------------------------------------------------------
def find_matching_columns(dataframe, keywords):
    return [col for col in dataframe.columns if any(kw in col.lower() for kw in keywords)]
 
##---------------------------------------------------------------------------------------------------------
def handle_displacement_column_selection():
    if 'household_data' in st.session_state:
        household_data = st.session_state['household_data']
        displacement_keywords = [
            'hh_displaced', 'pop_group', 'i_type_pop', 'statut', 'hh_forcibly_displaced',
            'demo_situation_menage', 'pop_group_name', 'residency_status', 'pop_group',
            'statut_menage', 'population_group', 'd_statut_deplacement', 'B_1_hh_primary_residence',
            'statutMenage', 'B_1_hh_primary_residence', 'status', 'displacement', 'origin'
        ]
        displacement_suggestions = find_matching_columns(household_data, displacement_keywords)

        if 'show_manual_select' not in st.session_state:
            st.session_state['show_manual_select'] = False

        if displacement_suggestions and not st.session_state['show_manual_select']:
            selected_displacement = st.selectbox(
                'Select the variable that corresponds to the status (host community, IDP, returnee) of the household:',
                ['No selection'] + displacement_suggestions,
                key='displacement_selectbox'
            )
            col1, col2 = st.columns(2)
            with col1:
                if selected_displacement != 'No selection' and st.button("Confirm Displacement Column"):
                    st.session_state['status_var'] = selected_displacement
                    st.session_state['displacement_column_confirmed'] = True
                    st.success(f"Displacement column '{selected_displacement}' has been confirmed.")

            with col2:
                if st.button("It is not listed, select manually"):
                    st.session_state['show_manual_select'] = True

        if st.session_state['show_manual_select']:
            selected_displacement = st.selectbox(
                f"Select the variable that corresponds to the status (host community, IDP, returnee) of the household:",
                ['No selection'] + household_data.columns.tolist(),
                key='manual_displacement_selectbox'
            )
            if selected_displacement != 'No selection' and st.button("Confirm Selected Column"):
                st.session_state['status_var'] = selected_displacement
                st.success(f"Displacement column '{selected_displacement}' has been manually selected.")
                st.session_state['displacement_column_confirmed'] = True
                    

###########################################################################################################
##-----------------------------
# Function to handle uploading and selecting data
def upload_and_select_data():
    if 'uploaded_data' in st.session_state:
        st.subheader('Selection of the relevant sheets in the MSNA data file')

        data = st.session_state['uploaded_data']

        if isinstance(data, dict):
            col1, col2 = st.columns(2)

            survey_sheet_guess = [col for col in list(data.keys()) if any(kw in col.lower() for kw in ['survey', 'questionnaire', 'enquête'])]
            choice_sheet_guess = [col for col in list(data.keys()) if any(kw in col.lower() for kw in ['choice', 'choix'])]
            
            with col1:
                selected_sheet = st.selectbox('Select the Household data sheet:', ['No selection'] + list(data.keys()), key='household_key')
                selected_survey_sheet = survey_sheet_guess[0] if survey_sheet_guess else st.selectbox('Select the Survey/kobo sheet:', ['No selection'] + list(data.keys()), key='survey_key')

            with col2:
                selected_edu_sheet = st.selectbox('Select the Education loop (or individual loop) data sheet:', ['No selection'] + list(data.keys()), key='edu_key')
                selected_choice_sheet = choice_sheet_guess[0] if choice_sheet_guess else st.selectbox('Select the Kobo choice sheet:', ['No selection'] + list(data.keys()), key='choice_key')


            if st.button('Confirm Data Selections') and not any(x == 'No selection' for x in [selected_sheet, selected_survey_sheet, selected_edu_sheet, selected_choice_sheet]):
                st.session_state['household_data'] = data[selected_sheet]
                st.session_state['survey_data'] = data[selected_survey_sheet]
                st.session_state['edu_data'] = data[selected_edu_sheet]
                st.session_state['choice_data'] = data[selected_choice_sheet]
                st.session_state.data_selections_confirmed = True
                st.success("Data selections updated successfully!")

            if 'survey_data' in st.session_state:
                survey_data = st.session_state['survey_data']
                label_columns = [col for col in survey_data.columns if col.startswith('label')]
                if label_columns:
                    selected_label = st.selectbox('Select the desired label column:', ['No selection'] + label_columns, key='selected_label')
                    if selected_label != 'No selection':
                        st.session_state['label'] = selected_label
                        if st.button("Confirm Label Language"):
                            st.session_state.label_selected = True
                            st.success(f"Label column '{selected_label}' has been selected.")
                            st.markdown("""
                                        <div style='background-color: #f0f8ff; padding: 10px; border-radius: 5px;'>
                                            <span style='color: #014bb4;'><strong>Proceed to the next step:</strong></span>
                                            <span style='color: #014bb4; font-style: italic; font-size: 20px;'>Select education indicators</span>
                                        </div>
                                        """, unsafe_allow_html=True)
    else:
        st.warning("No data uploaded. Please go to the previous page and upload data.")     
##-----------------------------
# Function to select indicators
def select_indicators():
    if 'edu_data' in st.session_state and st.session_state.get('label_selected', False) :
        edu_data = st.session_state['edu_data']
        st.subheader("Select the correct variables and indicators.")
        st.markdown("""
            <div style='background-color: #FDFD96; border-radius: 5px; padding: 10px; margin: 10px 0;'>
            <h6 style='color: #162AFD; margin: 0; padding: 0;'>Please check carefully when selecting a variable. The choice of indicators directly impacts the PiN calculation.⚠️</h6>
            </div>
            """, unsafe_allow_html=True)

        age_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['age', 'âge', 'year'])]
        gender_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['sex', 'gender', 'sexe', 'genre'])]
        education_indicator_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['edu', 'education', 'school', 'ecole'])]

        # Checkbox to show/hide the data header
        if st.checkbox('Display Education Data Header (This can assist in inspecting the contents of the data frame)'):
            st.dataframe(edu_data.head())
        
        if age_suggestions:
            st.session_state['age_var'] = handle_column_selection(age_suggestions, 'age')
        if gender_suggestions:
            st.session_state['gender_var'] = handle_column_selection(gender_suggestions, 'gender')

        if education_indicator_suggestions:
            st.session_state['access_var'] = handle_full_selection(education_indicator_suggestions, 'education_access', "% of children accessing education:")    
            st.session_state['teacher_disruption_var'] =  handle_full_selection(education_indicator_suggestions, 'disruption_teacher', "% of children whose access to education was disrupted due to teacher strikes or absenteeism:") 
            st.session_state['idp_disruption_var'] =  handle_full_selection(education_indicator_suggestions, 'disruption_idp', "% of children whose access to education was disrupted due the school being used as a shelter by displaced persons:") 
            st.session_state['armed_disruption_var'] =  handle_full_selection(education_indicator_suggestions, 'disruption_armed', "% of children whose access to education was disrupted due the school being occupied by armed groups:") 
            st.session_state['barrier_var'] = handle_full_selection(education_indicator_suggestions, 'barriers', "main barriers to access education:") 
            check_for_duplicate_selections()
        if st.button("Confirm Indicators"):
            st.session_state.indicators_confirmed = True
            st.success("Indicators confirmed!")
            st.markdown("""
            <div style='background-color: #f0f8ff; padding: 10px; border-radius: 5px;'>
                <span style='color: #014bb4;'><strong>Proceed to the next step:</strong></span>
                <span style='color: #014bb4; font-style: italic; font-size: 20px;'>Define aggravating cicurmstances</span>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.warning("Please go to the previous step and first select the right sheets in the MSNA data file.") 
##-----------------------------
# Function to define severity of barriers
def define_severity():
    if 'survey_data' in st.session_state and 'choice_data' in st.session_state and 'edu_data' in st.session_state and st.session_state.get('indicators_confirmed', False):
        survey_data = st.session_state['survey_data']
        choices_data = st.session_state['choice_data']
        edu_data = st.session_state['edu_data']

        barrier_var = st.session_state.get('barrier_var', 'Default Value if not set')
        selected_label = st.session_state['label'] 
        barrier_options = find_barrier_details(barrier_var, survey_data, choices_data, selected_label)

        # Encapsulate descriptions within a single box with a light gray background
        st.markdown("""
            <div style='background-color: #f8f9fa; padding: 20px; border-radius: 10px; border: 1px solid #ccc;'>
                <h5 style='color: #c0474a;'>For the next part, please refer to the PiN methodology shared on the home page!</h5>
                <h6 style='color: #014bb4;'>Severity 4 aggravating circumstances:</h6>
                <ul>
                    <li>Child marriage and child labour (work at home or on the household's own farm -- in income generating activities).</li>
                    <li>Protection risks while traveling to/at the school (includes dangers and injuries, physical and emotional maltreatment, sexual and gender based violence, and verbal harassment, mental health and psychosocial distress).</li>
                    <li>Lack of documentation for school enrollment: household is recently displaced and this is the reason why they do not have documentation.</li>
                    <li>Discrimination or stigmatization affecting access to education.</li>
                </ul>
                <h6 style='color: #014bb4;'>Severity 5 aggravating circumstances:</h6>
                <ul>
                    <li>Children recruitment by armed groups.</li>
                    <li>In some specific contexts, the presence of bans preventing children from attending education.</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
        
        selected_severity_4_barriers = select_severity_barriers(barrier_options, 4)
        barrier_options_5 = [option for option in barrier_options if option not in selected_severity_4_barriers]
        selected_severity_5_barriers = select_severity_barriers(barrier_options_5, 5)
        if st.button("Confirm Severity Definitions"):
            st.session_state.severity_confirmed = True
            st.success("Severity definitions confirmed!")
            st.markdown("""
                    <div style='background-color: #f0f8ff; padding: 10px; border-radius: 5px;'>
                        <span style='color: #014bb4;'><strong>Proceed to the next step:</strong></span>
                        <span style='color: #014bb4; font-style: italic; font-size: 20px;'>Define disaggregation variables</span>
                    </div>
                    """, unsafe_allow_html=True)  
    else:
        st.warning("Please return to the previous step and first select the correct variable referring to barriers to access education.")          
##-----------------------------
# Function to handle administrative details and final confirmations
def finalize_details():
    if st.session_state.get('severity_confirmed', False):
        st.subheader("Choose the correct disaggregation variables and their values")



        #admin_level_options = ['No selection', 'Admin0', 'Admin1', 'Admin2', 'Admin3']

        # Check if the country has been selected on the first page
            # Check if the country has been selected on the first page
        if 'country' in st.session_state and st.session_state['country'] != 'no selection':
            selected_country = st.session_state['country']
            # Get the administrative levels for the selected country
            admin_level_options = ['No selection'] + admin_levels_per_country.get(selected_country, [])
        else:
            # Default to a generic or empty option if no country is selected
            selected_country = "No selection"
            admin_level_options = ['No selection']

        # Display the selectbox with an integrated markdown for instructions
        st.markdown(
            f"What is the smallest administrative level in **{selected_country}** at which we can calculate the PiN to ensure the results are representative? <span style='color: darkred; font-weight: bold;'>Please ensure that the selected administrative level corresponds to the same administrative level as that of the OCHA population data.</span>",
            unsafe_allow_html=True
        )


        selected_admin_level = st.selectbox(
            "Select",
            admin_level_options,
            index=0,  # Default to 'No selection'
            key='admin_target'
        )

        if st.button('Confirm Admin Level', key='confirm_admin_level'):
            if selected_admin_level != 'No selection':
                st.session_state.admin_level_confirmed = True
                st.success(f"Administrative level '{selected_admin_level}' confirmed!")
            else:
                st.error("Please select a valid administrative level.")




        months = ['No selection','January', 'February', 'March', 'April', 'May', 'June', 
                'July', 'August', 'September', 'October', 'November', 'December']
        start_school = st.selectbox(
            "When does the school year officially start? Needed for the estimate of the correct age",
            months,
            index=0,  # Default to 'No selection'
            key='start_school'
        )
        if st.button('Confirm School Start Month', key='confirm_start_school'):
            if start_school != 'No selection':
                start_school = st.session_state['start_school']
                st.session_state.school_start_month_confirmed = True
                st.success("School start month confirmed!")
            else:
                st.error("Please select a valid month.")
            #update_other_parameters_status()



        ## -------------------- school cycle -----------------------------------
        upper_primary_start = st.session_state['lower_primary_end'] +1
        lower_primary_end = st.slider(
            "Which is the age range for the lower primary school cycle?",
            min_value=6, 
            max_value=17, 
            value=st.session_state['lower_primary_end'],
            step=1,
            key='lower_primary_end'
        )

        col1, col2 = st.columns(2)
        with col1:
            if st.button('OK'):
                st.session_state.single_cycle = False
                lower_primary_end = st.session_state['lower_primary_end'] 
                upper_primary_start = lower_primary_end +1                   

        with col2:
            if st.button('There is only one cycle of primary, followed by secondary school'):
                st.session_state.single_cycle = True

        # Depending on user selection, show different sliders or information
        if 'single_cycle' in st.session_state and not st.session_state['single_cycle']:
            # Slider for the upper primary school cycle age range
            upper_primary_end = st.slider(
                "Which is the age range for the upper primary school?",
                min_value=upper_primary_start, 
                max_value=17, 
                value=st.session_state['upper_primary_end'],
                step=1,
                key='upper_primary_end'
            )
            upper_primary_start = st.session_state['lower_primary_end'] + 1
            secondary_start = st.session_state['upper_primary_end'] + 1
            # Button to confirm the final age ranges
            if st.button('Confirm Age Ranges'):
                st.session_state.upper_primary_end_confirmed = True
                if upper_primary_end != st.session_state['upper_primary_end']:
                    upper_primary_end = st.session_state['upper_primary_end'] 
                vect1 =  st.session_state['lower_primary_end']  
                vect2 =  st.session_state['upper_primary_end']
                st.session_state['vector_cycle'] = [vect1,vect2]
                st.markdown(f"""
                <div style="border: 1px solid #cccccc; border-radius: 5px; padding: 10px; margin-top: 5px; background-color: #f0f0f0;">
                    <h6 style="color: #555555; margin-bottom: 5px;">Confirmed Age Ranges:</h6>
                    <div><strong>Lower Primary School:</strong> 6 - {lower_primary_end}</div>
                    <div><strong>Upper Primary School:</strong> {upper_primary_start} - {upper_primary_end}</div>
                    <div><strong>Secondary School:</strong> {secondary_start} - 17</div>
                </div>
                """, unsafe_allow_html=True)


        elif 'single_cycle' in st.session_state and st.session_state['single_cycle']:
            # Directly display age ranges for primary and secondary
            primary_end = st.session_state['lower_primary_end']
            secondary_start = primary_end + 1
            vect1 =  st.session_state['lower_primary_end']  
            vect2 =  0
            st.session_state['vector_cycle'] = [vect1,vect2]
            st.markdown(f"""
                <div style="border: 1px solid #cccccc; border-radius: 5px; padding: 10px; margin-top: 5px; background-color: #f0f0f0;">
                    <h6 style="color: #555555; margin-bottom: 5px;">Confirmed Age Ranges:</h6>
                    <div><strong>Primary School:</strong> 6 - {primary_end}</div>
                    <div><strong>Secondary School:</strong> {secondary_start} - 17</div>
                </div>
                """, unsafe_allow_html=True)
            
        if st.button("Confirm school-age ranges"):
            st.session_state.final_confirmed = True
            #vector_cycle = st.session_state['vector_cycle']
            #st.write (vector_cycle)
            st.success("School-age ranges confirmed")

        handle_displacement_column_selection()

        if st.session_state.get('displacement_column_confirmed', False):
            st.success(f"Displacement column confirmed: {st.session_state['status_var']}")
            pippo = st.session_state['status_var']
            st.write(f"Selected Displacement Column: {pippo}")

        if st.button("Finalize and Confirm"):
            st.session_state.final_confirmed = True
            st.success("All details confirmed and finalized!")

        st.markdown("---")
  
###########################################################################################################
###########################################################################################################

def display_step_content():
    if st.session_state['current_step'] == 0:
        upload_and_select_data()
    elif st.session_state['current_step'] == 1:
        select_indicators()
    elif st.session_state['current_step'] == 2:
        define_severity()
    elif st.session_state['current_step'] == 3:
        finalize_details()          

display_step_content()


st.markdown("---")  # Markdown horizontal rule


# Always show status indicators
# Always show status indicators
display_status("Data Selections Confirmed", st.session_state.data_selections_confirmed)
display_status("Label Selected", st.session_state.label_selected)
display_status("Age Column Confirmed", st.session_state.age_column_confirmed)
display_status("Gender Column Confirmed", st.session_state.gender_column_confirmed)

update_combined_indicator()
display_combined_severity_status()
update_other_parameters_status()

if all([
    st.session_state.get('data_selections_confirmed', False),
    st.session_state.get('label_selected', False),
    st.session_state.get('age_column_confirmed', False),
    st.session_state.get('gender_column_confirmed', False),
    st.session_state.get('indicators_confirmed', False),
    st.session_state.get('severity_4_confirmed', False),
    st.session_state.get('severity_5_confirmed', False),
    st.session_state.get('other_parameters_confirmed', False)
]):
    st.markdown("""
                <div style='background-color: #90EE90; padding: 10px; border-radius: 5px;'>
                    <span style='color: black; font-size: 20px;'><strong>Proceed to PiN calculation!!!</strong></span>
                </div>
                """, unsafe_allow_html=True)  
