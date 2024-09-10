import streamlit as st
import pandas as pd
st.logo('pics/logos.png')


st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico', layout='wide')
st.title('Indicator Selection and severity categorization')


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
        'selected_barriers': []  # List to store user-selected barriers
        
    })
if 'lower_primary_end' not in st.session_state:
    st.session_state['lower_primary_end'] = 11  # Default end age for lower primary
if 'upper_primary_end' not in st.session_state:
    st.session_state['upper_primary_end'] = 16  # Default end age for upper primary


##---------------------------------------------------------------------------------------------------------
# Function to display status indicators
def display_status(description, status):
    color = 'green' if status else 'gray'
    st.markdown(f"<span style='color: {color};'>●</span> {description}", unsafe_allow_html=True)
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
    with col1:
        if st.button("Yes", key=confirm_key):
            if suggested_column != 'No selection':
                st.session_state[f'selected_{column_type}_column'] = suggested_column
                st.session_state[f'{column_type}_column_confirmed'] = True
                st.success(f"{column_type.capitalize()} column '{suggested_column}' has been confirmed.")
    with col2:
        if st.button("No", key=f'confirm_no_{column_type}'):
            selected_column = st.selectbox(
                f"Select the individual {column_type} column:",
                ['No selection'] + edu_data.columns.tolist(),
                key=select_key,
                on_change=update_column_confirmation,
                args=(column_type,)
            )
##---------------------------------------------------------------------------------------------------------
def update_column_confirmation(column_type):
    select_key = f'{column_type}_selectbox'
    selected_column = st.session_state.get(select_key)
    if selected_column and selected_column != 'No selection':
        st.session_state[f'selected_{column_type}_column'] = selected_column
        st.session_state[f'{column_type}_column_confirmed'] = True
        st.success(f"{column_type.capitalize()} column '{selected_column}' has been manually selected.")
        display_status(f"{column_type.capitalize()} Column Confirmed", st.session_state[f'{column_type}_column_confirmed'])

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
    #prompt_message = f"Select aggravating circumstances falling under **severity {severity}**:"
    confirm_button_label = f"Confirm Severity {severity} Barriers"
    prompt_message = f"Select aggravating circumstances falling under **severity {severity}**, you can select multiple choices:"
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
    
    st.markdown(f"<span style='color: {color};'>●</span> {status_description}", unsafe_allow_html=True)
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
        st.error("Duplicate selections detected. Each variable should be used for only one category. Please adjust your selections.")


###########################################################################################################
###########################################################################################################
###########################################################################################################
if 'uploaded_data' in st.session_state:
    data = st.session_state['uploaded_data']
    if isinstance(data, dict):
        col1, col2 = st.columns(2)
        with col1:
            selected_sheet = st.selectbox('Select the Household data sheet:', ['No selection'] + list(data.keys()), key='household_key')
            selected_survey_sheet = st.selectbox('Select the Survey/kobo sheet:', ['No selection'] + list(data.keys()), key='survey_key')
        with col2:
            selected_edu_sheet = st.selectbox('Select the Education loop (or individual loop) data sheet:', ['No selection'] + list(data.keys()), key='edu_key')
            selected_choice_sheet = st.selectbox('Select the Kobo choice sheet:', ['No selection'] + list(data.keys()), key='choice_key')

        if st.button('Confirm Data Selections') and not any(x == 'No selection' for x in [selected_sheet, selected_survey_sheet, selected_edu_sheet, selected_choice_sheet]):
            st.session_state['household_data'] = data[selected_sheet]
            st.session_state['survey_data'] = data[selected_survey_sheet]
            st.session_state['edu_data'] = data[selected_edu_sheet]
            st.session_state['choice_data'] = data[selected_choice_sheet]
            st.session_state.data_selections_confirmed = True
            st.success("Data selections updated successfully!")

        if 'survey_data' in st.session_state:
            survey_data = st.session_state['survey_data']
            label_columns = [col for col in survey_data.columns if col.startswith('label::')]
            if label_columns:
                selected_label = st.selectbox('Select the desired label column:', ['No selection'] + label_columns, key='selected_label')
                if selected_label != 'No selection':
                    st.session_state['label'] = selected_label
                    st.session_state.label_selected = True
                    st.success(f"Label column '{selected_label}' has been selected.")

        if 'edu_data' in st.session_state:
            edu_data = st.session_state['edu_data']
            st.subheader("Now we need to select the correct variables and indicators.")
            st.markdown("""
                <div style='background-color: #ffdddd; border-radius: 5px; padding: 10px; margin: 10px 0;'>
                <h6 style='color: #b30000; margin: 0; padding: 0;'>Please check carefully when selecting a variable. The choice of indicators directly impacts the PiN calculation.⚠️</h6>
                </div>
                """, unsafe_allow_html=True)

            age_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['age', 'âge', 'year'])]
            gender_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['sex', 'gender', 'sexe', 'genre'])]
            education_indicator_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['edu', 'education', 'school', 'ecole'])]

            # Checkbox to show/hide the data header
            if st.checkbox('Display Education Data Header (This can assist in inspecting the contents of the data frame)'):
                st.dataframe(edu_data.head())
            
            if age_suggestions:
                handle_column_selection(age_suggestions, 'age')
            if gender_suggestions:
                handle_column_selection(gender_suggestions, 'gender')

            if education_indicator_suggestions:
                access_var = handle_full_selection(education_indicator_suggestions, 'education_access', "% of children accessing education:")    
                teacher_disruption_var =  handle_full_selection(education_indicator_suggestions, 'disruption_teacher', "% of children whose access to education was disrupted due to teacher strikes or absenteeism:") 
                idp_disruption_var =  handle_full_selection(education_indicator_suggestions, 'disruption_idp', "% of children whose access to education was disrupted due the school being used as a shelter by displaced persons:") 
                armed_disruption_var =  handle_full_selection(education_indicator_suggestions, 'disruption_armed', "% of children whose access to education was disrupted due the school being occupied by armed groups:") 
                barrier_var = handle_full_selection(education_indicator_suggestions, 'barriers', "main barriers to access education:") 
                check_for_duplicate_selections()
            if st.session_state.get('barriers_column_confirmed', False):
                if 'survey_data' in st.session_state and 'choice_data' in st.session_state and 'edu_data' in st.session_state:
                    survey_data = st.session_state['survey_data']
                    choices_data = st.session_state['choice_data']
                    edu_data = st.session_state['edu_data']

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

            st.markdown("---")  # Markdown horizontal rule
            col1, col2 = st.columns(2)
            with col1:
                admin_level_options = ['Admin0', 'Admin1', 'Admin2', 'Admin3']
                admin_target = st.selectbox(
                    "What is the smallest administrative level at which we can calculate the PiN to ensure the results are representative?",
                    admin_level_options,
                    key='admin_target'
                )
                if st.button('Confirm Admin Level', key='confirm_admin_level'):
                    st.session_state.admin_target_confirmed = True
                    st.success("Administrative level confirmed!")
            with col2:
                months = ['January', 'February', 'March', 'April', 'May', 'June', 
                        'July', 'August', 'September', 'October', 'November', 'December']
                start_school = st.selectbox(
                    "When does the school year officially start? Needed for the estimate of the correct age",
                    months,
                    key='start_school'
                )
                if st.button('Confirm School Start Month', key='confirm_start_school'):
                    st.session_state.start_school_confirmed = True
                    st.success("School start month confirmed!")

            upper_primary_start = st.session_state['lower_primary_end'] +1
            
            lower_primary_end = st.slider(
                "Which is the age range for the lower primary school cycle?",
                min_value=6, 
                max_value=18, 
                value=st.session_state['lower_primary_end'],
                step=1,
                key='lower_primary_end'
            )

            # Intermediate OK button to confirm the lower primary end selection
            if st.button('OK'):
                lower_primary_end = st.session_state['lower_primary_end'] 
                upper_primary_start = lower_primary_end +1
                # Ensure the upper primary starts at least a year after lower primary
                #st.session_state['upper_primary_end'] = lower_primary_end + 1

            # Slider for the upper primary school cycle age range
            upper_primary_end = st.slider(
                "Which is the age range for the upper primary school?",
                min_value=upper_primary_start, 
                max_value=18, 
                value=st.session_state['upper_primary_end'],
                step=1,
                key='upper_primary_end'
            )

            # Display the chosen age ranges
            upper_primary_start = st.session_state['lower_primary_end'] + 1
            secondary_start = st.session_state['upper_primary_end'] + 1



            # Button to confirm the final age ranges
            if st.button('Confirm Age Ranges'):
                if upper_primary_end != st.session_state['upper_primary_end']:
                    st.session_state['upper_primary_end'] = upper_primary_end
                st.success(f"Age ranges confirmed: Lower Primary up to {lower_primary_end}, Upper Primary up to {upper_primary_end}")
                st.write(f"Lower Primary School: 6 - {lower_primary_end}")
                st.write(f"Upper Primary School: {upper_primary_start} - {upper_primary_end}")
                st.write(f"Secondary School: {secondary_start} - 18")   

    else:
        st.warning("Survey data or educational data is not available. Please upload and select the data.")
else:
    st.warning("No data uploaded. Please go to the previous page and upload data.")






st.markdown("---")  # Markdown horizontal rule


# Always show status indicators
# Always show status indicators
display_status("Data Selections Confirmed", st.session_state.data_selections_confirmed)
display_status("Label Selected", st.session_state.label_selected)
display_status("Age Column Confirmed", st.session_state.age_column_confirmed)
display_status("Gender Column Confirmed", st.session_state.gender_column_confirmed)

update_combined_indicator()
display_combined_severity_status()



# Add or modify the section where OCHA data is uploaded
if 'uploaded_data' in st.session_state:
    if 'uploaded_ocha_data' in st.session_state:
        ocha_data = st.session_state['uploaded_ocha_data']
        st.write(translations["ok_upload"])
        st.dataframe(ocha_data.head())  # Show a preview of the data
    else:
        # OCHA data uploader
        uploaded_ocha_file = st.file_uploader(translations["upload_ocha"], type=["csv", "xlsx"])
        if uploaded_ocha_file is not None:
            ocha_data = pd.read_excel(uploaded_ocha_file, engine='openpyxl')
            check_message = perform_ocha_data_checks(ocha_data)
            if check_message == "Data is valid":
                st.session_state['uploaded_ocha_data'] = ocha_data
                st.success(translations["ok_upload"])
            else:
                st.error(check_message)  # Display the error message if checks fail

    # Iterate over the DataFrame rows to create the bullet points for each population group
    for _, row_pop in final_overview_df_OCHA.iterrows():
        total_population_in_need = row_pop[label_tot]
        strata = row_pop['Strata']
        
        if strata not in not_pop_group_columns:
            # Remove the substring '(5-17 y.o.)' and convert to uppercase
            strata_cleaned = strata.replace('(5-17 y.o.)', '').strip().upper()
            
            # Create a bullet point for each population group with indentation
            bullet_point = doc.add_paragraph(style='List Bullet')
            bullet_point_format = bullet_point.paragraph_format
            bullet_point_format.left_indent = Inches(1)  # Adjust this value for the desired indentation
            bullet_text = bullet_point.add_run(f"{format_number(total_population_in_need)} are {strata_cleaned} population group;")
            bullet_text.font.size = Pt(12)
            bullet_text.font.name = 'Calibri'




status_var = 'pop_group'
access_var = 'edu_access'
teacher_disruption_var = 'edu_disrupted_teacher'
idp_disruption_var = 'edu_disrupted_displaced'
armed_disruption_var = 'edu_disrupted_occupation'#'edu_disrupted_occupation'no_indicator
barrier_var = 'edu_barrier'
selected_severity_4_barriers = [
    "Protection/safety risks while commuting to school",
    "Protection/safety risks while at school",
    "Child needs to work at home or on the household's own farm (i.e. is not earning an income for these activities, but may allow other family members to earn an income)",
    "Child participating in income generating activities outside of the home",
    "Child marriage, engagement or pregnancies",
    "Discrimination or stigmatization of the child for any reason",
    "Unable to enroll in school due to lack of documentation"]
selected_severity_5_barriers = ["Child is associated with armed forces or armed groups "]
#"---> None of the listed barriers <---"
#"Child is associated with armed forces or armed groups "
age_var = 'ind_age'
gender_var = 'ind_gender'
start_school = 'September'
country= 'Myanmar -- MMR'

#admin_var = 'Admin_3: Townships'#'Admin_2: Regions'
 
# 'Admin_3: Townships'
admin_var = 'Admin_1: States/Regions'#'Admin_2: Regions' 

vector_cycle = [10,14]
single_cycle = (vector_cycle[1] == 0)
primary_start = 6
secondary_end = 17
label = 'label::English'

# Path to your Excel file
excel_path = 'input/REACH_MMR_MMR2402_MSNA_Dataset_VALIDATED.xlsx'
excel_path_ocha = 'input/ocha_pop_MMR.xlsx'
#excel_path_ocha = 'input/test_ocha.xlsx'




        # Iterate through each column in the edu_data dataframe
        for col in edu_data.columns:
            # Convert the column to strings to ensure type consistency
            column_data = edu_data[col].astype(str)            
            # Check if any value from prefix_list is present in the current column
            matching_values = column_data.isin(prefix_list)

            # If there are any matches, add the column to the list
            if matching_values.any():
                admin_column_rapresentative.append(col)





                 dimension_admin_status_list = run_mismatch_admin_analysis(df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)
        dimension_admin_status_in_need_list = run_mismatch_admin_analysis(in_need_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)      
        severity_female_list = run_mismatch_admin_analysis(female_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='severity_category',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)  
        severity_male_list = run_mismatch_admin_analysis(male_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='severity_category',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)  
        dimension_female_list = run_mismatch_admin_analysis(female_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)        
        dimension_male_list = run_mismatch_admin_analysis(male_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict) 
        dimension_ece_list = run_mismatch_admin_analysis(ece_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict) 
        dimension_primary_list = run_mismatch_admin_analysis(primary_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict) 
        dimension_secondary_list = run_mismatch_admin_analysis(secondary_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)  
        if not single_cycle:      
            dimension_intermediate_list = run_mismatch_admin_analysis(intermediate_df, admin_var,admin_column_rapresentative,pop_group_var,
                                analysis_variable='dimension_pin',
                                admin_low_ok_list = admin_low_ok_list, prefix_list = admin_up_msna,grouped_dict = grouped_dict)   