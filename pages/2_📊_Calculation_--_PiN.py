import streamlit as st
import pandas as pd
import extra_streamlit_components as stx
from shared_utils import language_selector

st.logo('pics/GEC Global English logo_Colour_JPEG.jpg')

st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico', layout='wide')

# Call the language selector function
language_selector()

# Access the translations
translations = st.session_state.translations

st.title(translations["title_page2"])


if 'password_correct' not in st.session_state:
    st.error(translations["no_user"])
    st.stop()

if 'current_step' not in st.session_state:
    st.session_state['current_step'] = 0

# Define the steps for the stepper bar
steps = [translations["step1"],translations["step2"],translations["step3"],translations["step4"]]
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
    'Afghanistan -- AFG': ['Admin_1: Region', 'Admin_2: Province', 'Admin_3: Districts'],
    'Burkina Faso -- BFA': ['Admin_1: Regions (Région)', 'Admin_2: Province', 'Admin_3: Department (Département)'],
    'Cameroon -- CMR': ['Admin_1', 'Admin_2', 'Admin_3'],
    'Central African Republic -- CAR': ['Admin_1: Prefectures (préfectures)', 'Admin_2: Sub-prefectures (sous-préfectures)', 'Admin_3: Communes'],
    'Democratic Republic of the Congo -- DRC': ['Admin_1: Provinces', 'Admin_2: Territories', 'Admin_3: Sectors/chiefdoms/communes'],
    'Haiti -- HTI': ['Admin_1: Departments (départements)', 'Admin_2: Arrondissements', 'Admin_3: Communes'],
    'Iraq -- IRQ': ['Admin_1: Governorates', 'Admin_2: Districts (aqḍyat)', 'Admin_3: Sub-districts (naḥiyat)'],
    'Lemuria -- LMR':['Admin_1: Province', 'Admin_2: District', 'Admin_3: Subdistrict'] ,
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
    'Somalia -- SOM': ['Admin_1: States', 'Admin_2: Districts', 'Admin_3: sub-Districts']
}



##---------------------------------------------------------------------------------------------------------
# Function to display status indicators
def display_status(description, status):
    color = 'green' if status else 'gray'
    st.markdown(f"<span style='color: {color}; font-size: 20px; margin-right: 5px;'>●</span> {description}", unsafe_allow_html=True)
##---------------------------------------------------------------------------------------------------------
def handle_full_selection(current_country, suggestions, column_type, custom_message):
    # Construct the full message with custom formatting
    message_template = translations["full_message"]
    full_message = message_template.format(custom_message=custom_message)
    #full_message = f"Please select the variable required for measuring the <span style='color: #014bb4;'><strong>{custom_message}</strong></span>"
    st.markdown(full_message, unsafe_allow_html=True)

    # Use a key that ensures the selectbox is unique and not accidentally re-used
    select_key = f'{column_type}_selectbox'
    selected_column = st.selectbox(
        "Choose one/Choisissez une option:",
        ['No selection'] + suggestions,
        key=select_key
    )

    # Load the translated messages from session state
    confirm_button_label = translations["confirm_button_label"].format(column_type=column_type.replace('_', ' ').capitalize())
    column_confirmed_message = translations["column_confirmed_message"]
    error_message = translations["error_message"]

    if column_type == 'disruption_idp' and current_country == 'Burkina Faso -- BFA':
        st.text_area("Message important", 
            """Conformément à la méthodologie convenue, la dimension de l'environnement protégé est déterminée par trois indicateurs de protection dans l'évaluation MSNA :
            1) École utilisée comme abri par des personnes déplacées
            2) Incidents de protection sur le trajet de l'école (violences, harcèlement verbal/physique, VBG, EEI, etc.)
            3) Incidents de protection au sein de l'école (violences, harcèlement verbal/physique, VBG, etc.) concourent à attribuer l'enfant dans la dimension de l'environnement protégé.""",
            height=150
        )
    # Add a confirmation button
    if st.button(confirm_button_label, key=f'confirm_{column_type}'):
        if selected_column != 'No selection':
            st.session_state[f'selected_{column_type}_column'] = selected_column
            st.session_state[f'{column_type}_column_confirmed'] = True
            st.success(column_confirmed_message.format(column_type=column_type.capitalize(), selected_column=selected_column))
        else:
            st.error(error_message.format(column_type=column_type.replace('_', ' ').capitalize()))

    # Update the status indicator after checking the session state
    display_status(f"{column_type.capitalize()} Column Confirmed", st.session_state.get(f'{column_type}_column_confirmed', False))

    # Return the selected column for further processing
    return selected_column

##-----------------------------
def handle_armed_disruption_selection(current_country, suggestions):
    # Display a checkbox to indicate if this indicator was not collected
    translated_text = translations["no_armed_disruption_indicator"]
    no_indicator_collected = st.checkbox(f"{translated_text}")
    column_type = 'disruption_armed'
    # If checkbox is checked, mark armed disruption as 'no_indicator' and skip the selectbox
    if no_indicator_collected:
        armed_disruption_var = "no_indicator"
        st.session_state[f'selected_{column_type}_column'] = "no_indicator"
        st.session_state[f'{column_type}_column_confirmed'] = True
    else:
        # If checkbox is not checked, proceed with the regular selection
        armed_disruption_var = handle_full_selection(current_country, suggestions, 'disruption_armed', translations["armed_disruption_var_prompt"])

    return armed_disruption_var
##-----------------------------
def handle_natural_hazard_disruption_selection(current_country, suggestions):
    # Display a checkbox for the natural hazard disruption indicator
    translated_text = translations["yes_natural_hazard_disruption_indicator"]
    indicator_collected = st.checkbox(f"{translated_text}")
    column_type = 'disruption_natural_hazard'
    
    # If checkbox is checked, proceed with the regular selection
    if indicator_collected:
        natural_hazard_disruption_var = handle_full_selection(
            current_country, 
            suggestions, 
            'disruption_natural_hazard', 
            translations["natural_hazard_disruption_var_prompt"]
        )
    else:
        # If checkbox is not checked, mark natural hazard disruption as 'no_indicator'
        natural_hazard_disruption_var = "no_indicator"
        st.session_state[f'selected_{column_type}_column'] = "no_indicator"
        st.session_state[f'{column_type}_column_confirmed'] = True

    return natural_hazard_disruption_var
##---------------------------------------------------------------------------------------------------------
def handle_column_selection(suggestions, column_type):
    suggested_column = suggestions[0] if suggestions else 'No selection'
    message_template = translations["is_this_individual_column_message"]

    st.write(message_template.format(column_type=column_type.replace('_', ' ').capitalize(), suggested_column=suggested_column))
    col1, col2 = st.columns(2)
    confirm_key = f'confirm_yes_{column_type}'
    select_key = f'{column_type}_selectbox'
    edu_data = st.session_state['edu_data']
    message_placeholder = st.empty()  # Place to show messages dynamically

    with col1:
        if st.button("Yes/Oui", key=confirm_key):
            if suggested_column != 'No selection':
                st.session_state[f'selected_{column_type}_column'] = suggested_column
                st.session_state[f'{column_type}_column_confirmed'] = True
                #st.success(f"{column_type.capitalize()} column '{suggested_column}' has been confirmed.")                                
                message_placeholder.success(f"{column_type.capitalize()} column '{suggested_column}' has been confirmed.")

    with col2:
        if st.button("No/Non", key=f'confirm_no_{column_type}'):
            suggested_column = st.selectbox(
                f"Select the individual {column_type} column:",
                ['No selection'] + edu_data.columns.tolist(),
                key=select_key,
                #on_change=update_column_confirmation (column_type,message_placeholder)
                #args=(column_type,)
            )
            update_column_confirmation (column_type,message_placeholder)
    return suggested_column
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

    display_status(translations["indicator_selection_confirmed"], st.session_state.indicators_confirmed)
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
    st.write(translations["select_aggravating_circumstances_message"])
    selected_barriers = []
    for index, row in barrier_details.iterrows():
        if st.checkbox(f"{row[label_column]}", key=f"select_{row['name']}"):
            selected_barriers.append(row['name'])
    st.session_state.selected_barriers = selected_barriers  # Update session state
##---------------------------------------------------------------------------------------------------------

def select_severity_barriers(barrier_options, severity):

    # Load translated messages from session state
    confirm_button_label = translations["confirm_button_label_severity"].format(severity=severity)
    prompt_message = translations["prompt_message"].format(severity=severity)
    none_of_listed_barriers = translations["none_of_listed_barriers"]
    success_message = translations["success_message"].format(severity=severity)
    confirmed_barriers_message = translations["confirmed_barriers_message"].format(severity=severity)

    # Add special option for severity 5
    if severity == 5:
        barrier_options = barrier_options + [none_of_listed_barriers]

    # Multiselect for barriers
    selected_barriers = st.multiselect(
            prompt_message,
            barrier_options,
            [])  # Start with no pre-selected barriers

    # Confirmation button
    if st.button(confirm_button_label):
        if severity == 4:
            st.session_state.selected_severity_4_barriers = selected_barriers
            st.session_state['severity_4_confirmed'] = True

        elif severity == 5:
            st.session_state.selected_severity_5_barriers = selected_barriers
            st.session_state['severity_5_confirmed'] = True
            # Handle the special case when 'None of the listed barriers' is selected
            if none_of_listed_barriers in selected_barriers:
                selected_barriers = [none_of_listed_barriers]
                st.session_state.selected_severity_5_barriers = selected_barriers

        # Display success message
        st.success(success_message)
        st.write(confirmed_barriers_message, selected_barriers)

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
        st.error(translations["duplicate_selections_error"])
##---------------------------------------------------------------------------------------------------------
def update_other_parameters_status():
    if (st.session_state.get('admin_level_confirmed', False) and
        st.session_state.get('school_start_month_confirmed', False) and
        (st.session_state.get('single_cycle', False) or st.session_state.get('upper_primary_end_confirmed', False)) and
        st.session_state.get('displacement_column_confirmed', False)) :
        st.session_state['other_parameters_confirmed'] = True
    else:
        st.session_state['other_parameters_confirmed'] = False

    display_status("Other Parameters Confirmed", st.session_state['other_parameters_confirmed'])


##---------------------------------------------------------------------------------------------------------
def find_matching_columns(dataframe, keywords):
    return [col for col in dataframe.columns if any(kw in col.lower() for kw in keywords)]
 
##---------------------------------------------------------------------------------------------------------
def handle_displacement_column_selection():
    if 'household_data' in st.session_state:
        household_data = st.session_state['household_data']
        displacement_keywords = [
            'hh_displaced', 'pop_group', 'i_type_pop', 'statut', 'hh_forcibly_displaced','statut',
            'demo_situation_menage', 'pop_group_name', 'residency_status', 'pop_group','population',
            'statut_menage', 'population_group', 'd_statut_deplacement', 'B_1_hh_primary_residence',
            'statutMenage', 'B_1_hh_primary_residence', 'status', 'displacement', 'origin', 'urbanity', 'urban', "depl_situation_menage_final"
        ]
        displacement_suggestions = find_matching_columns(household_data, displacement_keywords)

        if 'show_manual_select' not in st.session_state:
            st.session_state['show_manual_select'] = False

        if displacement_suggestions and not st.session_state['show_manual_select']:
            selected_displacement = st.selectbox(
                translations["select_status"],
                ['No selection'] + displacement_suggestions,
                key='displacement_selectbox'
            )
            col1, col2 = st.columns(2)
            with col1:
                if selected_displacement != 'No selection' and st.button(translations["confirm_status"]):
                    st.session_state['status_var'] = selected_displacement
                    st.session_state['displacement_column_confirmed'] = True
                    st.success(translations["success_status"])

            with col2:
                if st.button("It is not listed, select manually"):
                    st.session_state['show_manual_select'] = True

        if st.session_state['show_manual_select']:
            selected_displacement = st.selectbox(
                translations["select_status"],
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
        st.subheader(translations["sheet"])

        data = st.session_state['uploaded_data']

        if isinstance(data, dict):
            col1, col2 = st.columns(2)

            survey_sheet_guess = [col for col in list(data.keys()) if any(kw in col.lower() for kw in ['survey', 'questionnaire', 'enquête'])]
            choice_sheet_guess = [col for col in list(data.keys()) if any(kw in col.lower() for kw in ['choice', 'choix'])]
                        # Load the translated strings from session state
            select_household_data_sheet = translations["select_household_data_sheet"]
            select_survey_kobo_sheet = translations["select_survey_kobo_sheet"]
            select_education_loop_data_sheet = translations["select_education_loop_data_sheet"]
            select_kobo_choice_sheet = translations["select_kobo_choice_sheet"]

            # Example usage in your app
            with col1:
                selected_sheet = st.selectbox(select_household_data_sheet, ['No selection'] + list(data.keys()), key='household_key')
                selected_survey_sheet = survey_sheet_guess[0] if survey_sheet_guess else st.selectbox(select_survey_kobo_sheet, ['No selection'] + list(data.keys()), key='survey_key')

            with col2:
                selected_edu_sheet = st.selectbox(select_education_loop_data_sheet, ['No selection'] + list(data.keys()), key='edu_key')
                selected_choice_sheet = choice_sheet_guess[0] if choice_sheet_guess else st.selectbox(select_kobo_choice_sheet, ['No selection'] + list(data.keys()), key='choice_key')

            label_confirm_1 = translations["confirm_1"]
            label_success_1 = translations["success_1"]
            label_confirm_2 = translations["confirm_2"]

            if st.button(label_confirm_1) and not any(x == 'No selection' for x in [selected_sheet, selected_survey_sheet, selected_edu_sheet, selected_choice_sheet]):
                st.session_state['household_data'] = data[selected_sheet]
                st.session_state['survey_data'] = data[selected_survey_sheet]
                st.session_state['edu_data'] = data[selected_edu_sheet]
                st.session_state['choice_data'] = data[selected_choice_sheet]
                st.session_state.data_selections_confirmed = True
                st.success(label_success_1)

            if 'survey_data' in st.session_state:
                survey_data = st.session_state['survey_data']
                label_columns = [col for col in survey_data.columns if col.startswith('label')]
                if label_columns:
                    selected_label = st.selectbox(translations["label_json"], ['No selection'] + label_columns, key='selected_label')
                    if selected_label != 'No selection':
                        st.session_state['label'] = selected_label
                        if st.button(label_confirm_2):
                            st.session_state.label_selected = True
                            message_label_sucess =  translations["success_label_kobo"].format(selected_label=selected_label)
                            st.success(message_label_sucess)
                            st.markdown(translations["proceed_to_next_step"], unsafe_allow_html=True)

    else:
        st.warning(translations["no_data"])     

##-----------------------------
# Function to select indicators
def select_indicators():
    if 'edu_data' in st.session_state and st.session_state.get('label_selected', False) :
        edu_data = st.session_state['edu_data']
        # Display the translated subheader
        st.subheader(translations["select_variables_and_indicators_subheader"])

        # Display the translated HTML content
        st.markdown(translations["check_variable_warning_html"], unsafe_allow_html=True)

        age_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['age', 'âge'])]
        gender_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['sex', 'gender', 'sexe', 'genre'])]
        education_indicator_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['edu', 'education', 'school', 'ecole', 'scolarise', 'enseignant', 'formel', 'access'])]

        # Checkbox to show/hide the data header
        if st.checkbox(translations["display_education_data_header_checkbox"]):
            st.dataframe(edu_data.head())
        
        if age_suggestions:
            age_found = handle_column_selection(age_suggestions, 'age')
            st.session_state['age_var'] = age_found
        if gender_suggestions:
            st.session_state['gender_var'] = handle_column_selection(gender_suggestions, 'gender')

        if 'country' in st.session_state and st.session_state['country'] != 'no selection':
            current_country = st.session_state['country']
        if education_indicator_suggestions:
            st.markdown("---")  # Markdown horizontal rule
            st.session_state['access_var'] = handle_full_selection(current_country,education_indicator_suggestions, 'education_access', translations["access_var_prompt"])    
            st.markdown("---")  # Markdown horizontal rule
            st.session_state['teacher_disruption_var'] =  handle_full_selection(current_country,education_indicator_suggestions, 'disruption_teacher',translations["teacher_disruption_var_prompt"]) 
            st.markdown("---")  # Markdown horizontal rule
            st.session_state['natural_hazard_disruption_var'] =  handle_natural_hazard_disruption_selection(current_country, education_indicator_suggestions) 
            st.markdown("---")  # Markdown horizontal rule
            st.session_state['idp_disruption_var'] =  handle_full_selection(current_country,education_indicator_suggestions, 'disruption_idp', translations["idp_disruption_var_prompt"]) 
            st.markdown("---")  # Markdown horizontal rule
            st.session_state['armed_disruption_var'] =  handle_armed_disruption_selection(current_country, education_indicator_suggestions)  
            st.markdown("---")  # Markdown horizontal rule
            st.session_state['barrier_var'] = handle_full_selection(current_country, education_indicator_suggestions, 'barriers', translations["barrier_var_prompt"]) 
            check_for_duplicate_selections()
        if st.button(translations["confirm_indicators"]):
            st.session_state.indicators_confirmed = True
            st.success(translations["success_indicator"])

            # Display the HTML content
            st.markdown(translations["proceed_to_next_step3"], unsafe_allow_html=True)
    else:
        st.warning(translations["no_data"]) 
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
        st.markdown(translations["severity_circumstances_html"], unsafe_allow_html=True)
        
        selected_severity_4_barriers = select_severity_barriers(barrier_options, 4)
        barrier_options_5 = [option for option in barrier_options if option not in selected_severity_4_barriers]
        selected_severity_5_barriers = select_severity_barriers(barrier_options_5, 5)
        if st.button(translations["confirm_severity_all"]):
            st.session_state.severity_confirmed = True
            st.success(translations["success_severity_all"])
            # Display the HTML content
            st.markdown(translations["disaggregation_variables_html"], unsafe_allow_html=True)
    else:
        st.warning(translations["barriers_warning_message"])          
##-----------------------------
# Function to handle administrative details and final confirmations
def finalize_details():
    if st.session_state.get('severity_confirmed', False):
        st.subheader(translations["choose_disaggregation_variables_subheader"])
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
        admin_message = translations["smallest_admin_level"]
        markdown_message = admin_message.format(selected_country=selected_country)
        st.markdown(markdown_message, unsafe_allow_html=True)


        admin_target = st.selectbox(
            "Select",
            admin_level_options,
            index=0,  # Default to 'No selection'
            key='admin_target'
        )

        st.markdown(
            f"""
            <div style="font-size:18px; font-weight:bold;">
                {translations['mismatch_admin_message']}
            </div>
            """, unsafe_allow_html=True
        )

        # Display a simple checkbox below the message
        mismatch_admin_checkbox = st.checkbox(translations['check_box'])
        
        if st.button(translations["confirm_admin"], key='confirm_admin_level'):
            if admin_target != 'No selection':
                st.session_state['admin_var'] = admin_target
                st.session_state.admin_level_confirmed = True
                if mismatch_admin_checkbox:
                    st.session_state['mismatch_admin'] = True
                success_message_admin=  translations["success_admin"].format(admin_target=admin_target)
    
                st.success(success_message_admin)
            else:
                st.error("Please select a valid administrative level.")




        months = ['No selection','January', 'February', 'March', 'April', 'May', 'June', 
                'July', 'August', 'September', 'October', 'November', 'December']
        start_school_selection = st.selectbox(
            translations["start_message"],
            months,
            index=0,  # Default to 'No selection'
            key='start_school_selection'
        )
        if st.button(translations["confirm_school"], key='confirm_start_school'):
            if start_school_selection != 'No selection':
                st.session_state.school_start_month_confirmed = True
                st.session_state['start_school'] = start_school_selection 

                st.success(translations["success_start_school"])
                pluto =  st.session_state['start_school']
                st.write(pluto)
            else:
                st.error("Please select a valid month.")
            #update_other_parameters_status()



        ## -------------------- school cycle -----------------------------------
        upper_primary_start = st.session_state['lower_primary_end'] +1
        if st.session_state['country'] != 'Afghanistan -- AFG':
            lower_primary_end = st.slider(
                translations["school1"],
                min_value=6, 
                max_value=17, 
                value=st.session_state['lower_primary_end'],
                step=1,
                key='lower_primary_end'
            )
        else:
            lower_primary_end = st.slider(
                translations["school1"],
                min_value=7, 
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
            if st.button(translations["school2"]):
                st.session_state.single_cycle = True

        # Depending on user selection, show different sliders or information
        if 'single_cycle' in st.session_state and not st.session_state['single_cycle']:
            # Slider for the upper primary school cycle age range
            upper_primary_end = st.slider(
                translations["school3"],
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
                if st.session_state['country'] != 'Afghanistan -- AFG': school4_message = translations["school4"]
                else: school4_message = translations["school4_afg"]
                school4_content = school4_message.format(
                lower_primary_end=lower_primary_end,
                upper_primary_start=upper_primary_start,
                upper_primary_end=upper_primary_end,
                secondary_start=secondary_start
                )
                st.markdown(school4_content, unsafe_allow_html=True)



        elif 'single_cycle' in st.session_state and st.session_state['single_cycle']:
            # Directly display age ranges for primary and secondary
            primary_end = st.session_state['lower_primary_end']
            secondary_start = primary_end + 1
            vect1 =  st.session_state['lower_primary_end']  
            vect2 =  0
            st.session_state['vector_cycle'] = [vect1,vect2]
            if st.session_state['country'] != 'Afghanistan -- AFG': school5_message = translations["school5"]
            else: school5_message = translations["school5_afg"]
            # Insert the dynamic values into the HTML template
            school5_content = school5_message.format(
                primary_end=primary_end,
                secondary_start=secondary_start
            )

            # Display the HTML content
            st.markdown(school5_content, unsafe_allow_html=True)
        

            
        handle_displacement_column_selection()

        if st.button(translations["last_confirm"]):
            st.session_state.final_confirmed = True
            #st.success("All details confirmed and finalized!")



        #st.markdown("---")
  
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
        <div style='background-color: #90EE90; padding: 10px; border-radius: 5px; display: inline-block;'>
            <span style='color: black; font-size: 20px;'><strong>Completed / Terminé !!!!</strong></span>
        </div>
        """, unsafe_allow_html=True)  
    #st.markdown("---")  # Markdown horizontal rule
    col1, col2 = st.columns([0.60, 0.40])
    with col2: 
        st.page_link("pages/3_📋_Download_--_PiN_figures_and_other_outputs.py", label=translations['to_page3'], icon='📋')
    
    #if st.button('Calculate PiN'):

    st.markdown("---")  # Markdown horizontal rule
    st.markdown("---")  # Markdown horizontal rule
    st.markdown("---")  # Markdown horizontal rule
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
    st.write("Natural hazard Variable:", st.session_state.get('natural_hazard_disruption_var'))
    st.write("Barrier Variable:", st.session_state.get('barrier_var'))
    st.write("Selected Severity 4 Barriers:", st.session_state.get('selected_severity_4_barriers', []))
    st.write("Selected Severity 5 Barriers:", st.session_state.get('selected_severity_5_barriers', []))
    st.write("Admin Variable:", st.session_state.get('admin_target'))


    start_school =  st.session_state.get('start_school')
    vector_cycle =  st.session_state.get('vector_cycle')
    country =  st.session_state.get('country')
    edu_data =  st.session_state.get('edu_data').to_dict()  # Convert DataFrame to dict
    household_data =  st.session_state.get('household_data').to_dict()  # Convert DataFrame to dict
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
    admin_var =  st.session_state.get('admin_target')

# Always show status indicators
# Always show status indicators
display_status("Data Selections Confirmed", st.session_state.data_selections_confirmed)
display_status("Label Selected", st.session_state.label_selected)
display_status("Age Column Confirmed", st.session_state.age_column_confirmed)
display_status("Gender Column Confirmed", st.session_state.gender_column_confirmed)
update_combined_indicator()
display_combined_severity_status()
update_other_parameters_status()


