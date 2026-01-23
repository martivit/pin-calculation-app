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

data_combination = st.session_state.get('data_combination', "")
#st.write (data_combination)


if 'password_correct' not in st.session_state:
    st.error(translations["no_user"])
    st.stop()

if 'current_step' not in st.session_state:
    st.session_state['current_step'] = 0

# Define the steps for the stepper bar
steps = [translations["step1"],translations["step2"],translations["step3"],translations["step4"]]
steps_nomsna = [translations["step4"]]

current_step = st.session_state['current_step']
if 'm' in data_combination:
    steps_to_show = steps
else:
    steps_to_show = steps_nomsna

# call stepper_bar exactly once
new_step = stx.stepper_bar(steps=steps_to_show)

# update your session state
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
        'other_parameters_confirmed': False, 
        'custom_indicator_mappings': {},   # central dictionary: {column_name: {dimension, severity, column_type}}
        'custom_indicator_mappings_2': {},   # central dictionary: {column_name: {dimension, severity, column_type}}
        'additional_indicators': []  ,     # list of {"column", "dimension", "severity", "column_type"}
        'additional_indicator_enable': False,  # NEW: explicit off by default
        'additional_2_indicators': []  ,     # list of {"column", "dimension", "severity", "column_type"}
        'additional_2_indicator_enable': False  # NEW: explicit off by default

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
    'Democratic Republic of the Congo -- DRC': ['Admin_1', 'Admin_2', 'Admin_3'],
    'Ethiopia -- ETH':['Admin_1', 'Admin_2', 'Admin_3'],
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
    'Somalia -- SOM': ['Admin_1: States', 'Admin_2: Districts', 'Admin_3: sub-Districts'],
    'South Sudan -- SSD':['Admin_1', 'Admin_2', 'Admin_3'], 
    'Sudan -- SDN':['Admin_1', 'Admin_2', 'Admin_3'],
    'Sparkea -- SPR' :['Admin_1: Region', 'Admin_2: District', 'Admin_3: Commune']
}


##-----------------------------------------  functions   --------------------------------------------------

##---------------------------------------------------------------------------------------------------------
# Function to display status indicators
def display_status(description, status):
    color = 'green' if status else 'gray'
    st.markdown(f"<span style='color: {color}; font-size: 20px; margin-right: 5px;'>●</span> {description}", unsafe_allow_html=True)
##---------------------------------------------------------------------------------------------------------
def handle_full_selection(current_country, suggestions, column_type, custom_message, custom_message_part2):
    # Construct the full message with custom formatting
    message_template = translations["full_message"]
    #full_message = message_template.format(custom_message=custom_message)
    full_message = message_template.format(custom_message=custom_message, custom_message_part2=custom_message_part2)

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
        armed_disruption_var = handle_full_selection(current_country, suggestions, 'disruption_armed', translations["armed_disruption_var_prompt"],translations["armed_disruption_var_prompt_2"] )

    return armed_disruption_var
##-----------------------------
def select_dimension_and_severity(key_prefix: str):
    """
    Renders a dimension/severity selector and stores results in session_state:
      - f"{key_prefix}_dimension" in {"learning_conditions", "protected_environment"}
      - f"{key_prefix}_severity" in {3, 4}
      - f"{key_prefix}_dimension_confirmed" in {True/False}
    """
    # Translations (with safe fallbacks)
    q_text = translations.get(
        "dimension_severity_question",
        "Which dimension & severity does this indicator fall under?"
    )
    opt1 = translations.get(
        "dimension_learning_conditions_opt",
        "Learning conditions (Severity 3)"
    )
    opt2 = translations.get(
        "dimension_protected_env_opt",
        "Protected environment (Severity 4)"
    )
    confirm_lbl = translations.get("confirm_dimension_severity", "Confirm selection")

    choice = st.radio(
        q_text,
        [opt1, opt2],
        key=f"{key_prefix}_dimsev_radio"
    )

    if st.button(confirm_lbl, key=f"{key_prefix}_dimsev_confirm"):
        if choice == opt1:
            st.session_state[f"{key_prefix}_dimension"] = "learning_conditions"
            st.session_state[f"{key_prefix}_severity"] = 3
        else:
            st.session_state[f"{key_prefix}_dimension"] = "protected_environment"
            st.session_state[f"{key_prefix}_severity"] = 4
        st.session_state[f"{key_prefix}_dimension_confirmed"] = True
        st.success(translations.get("dimension_severity_saved", "Saved."))

##-----------------------------
##-----------------------------
def handle_natural_hazard_disruption_selection(current_country, suggestions):
    """
    Same logic as before, but when the indicator is collected and the column
    is confirmed, we also ask the user to choose the severity (3 or 4) and save it to
    st.session_state['natural_hazard_disruption_severity'].
    """
    translated_text = translations["yes_natural_hazard_disruption_indicator"]
    indicator_collected = st.checkbox(f"{translated_text}", key="natural_hazard_collected")
    column_type = 'disruption_natural_hazard'

    if indicator_collected:
        # 1) Select the column (unchanged)
        natural_hazard_disruption_var = handle_full_selection(
            current_country,
            suggestions,
            column_type,
            translations["natural_hazard_disruption_var_prompt"],
            translations["natural_hazard_disruption_var_prompt_2"]
        )

        # 2) If the column is confirmed, ask for SEVERITY (only 3 or 4)
        selected_col = st.session_state.get(f"selected_{column_type}_column")
        confirmed = st.session_state.get(f"{column_type}_column_confirmed", False)

        if confirmed and selected_col and selected_col != "no_indicator":
            # Labels with safe fallbacks
            q_text = translations.get(
                "natural_hazard_severity_question",
                "Which severity does this indicator fall under?"
            )
            opt_s3 = translations.get(
                "natural_hazard_severity3_opt",
                "Severity 3 (Learning conditions)"
            )
            opt_s4 = translations.get(
                "natural_hazard_severity4_opt",
                "Severity 4 (Protected environment)"
            )
            confirm_lbl = translations.get("confirm_natural_hazard_severity", "Confirm severity")

            choice = st.radio(
                q_text,
                [opt_s3, opt_s4],
                key=f"{column_type}_severity_radio"
            )

            if st.button(confirm_lbl, key=f"{column_type}_severity_confirm"):
                st.session_state['natural_hazard_disruption_severity'] = 3 if choice == opt_s3 else 4
                st.success(translations.get("natural_hazard_severity_saved", "Severity saved."))

    else:
        # If checkbox is not checked, mark natural hazard disruption as 'no_indicator'
        natural_hazard_disruption_var = "no_indicator"
        st.session_state[f'selected_{column_type}_column'] = "no_indicator"
        st.session_state[f'{column_type}_column_confirmed'] = True
        # Clear severity if previously set
        st.session_state['natural_hazard_disruption_severity'] = None

    return natural_hazard_disruption_var

##---------------------------------------------------------------------------------------------------------
def handle_additional_selection(current_country, suggestions):
    """
    Optional additional indicator with an explicit on/off toggle.
    When OFF, we clear any previously saved additional indicator state.
    When ON, user can pick a column and set severity (3/4).
    """
    st.markdown(
        translations.get(
            "additional_indicator_title",
            "<b>Optional: Add an additional indicator</b>"
        ),
        unsafe_allow_html=True
    )

    enable = st.checkbox(
        translations.get("additional_indicator_enable", "Add an additional indicator"),
        key="additional_indicator_enable"
    )

    if not enable:
        # ✅ Set explicit "no_indicator" instead of None
        st.session_state["additional_indicators"] = [{"column": "no_indicator", "severity": None}]
        st.session_state["additional_indicator_last_var"] = "no_indicator"
        st.session_state["additional_indicator_last_severity"] = None
        st.session_state["additional_indicator_last_dimension"] = None

        # Also remove any "additional" entries from the central mapping
        mapping = st.session_state.get("custom_indicator_mappings", {})
        mapping = {k: v for k, v in mapping.items() if v.get("column_type") != "additional"}
        st.session_state["custom_indicator_mappings"] = mapping

        st.info(translations.get("additional_indicator_disabled_info",
                                  "No additional indicator will be used."))
        return None  # nothing selected

    # --- If enabled: allow picking + assigning severity ---
    pick = st.selectbox(
        translations.get("additional_indicator_pick", "Pick an indicator to add:"),
        ['No selection'] + suggestions,
        key="additional_indicator_selectbox"
    )

    if pick == 'No selection':
        st.warning(translations.get("additional_indicator_none_selected",
                                    "Select an indicator to continue, or untick the box to skip."))
        return "no_indicator"

    with st.container(border=True):
        st.markdown(
            translations.get(
                "additional_indicator_dimsev",
                "Assign this indicator to a dimension & severity:"
            ),
            unsafe_allow_html=True
        )

        prefix = f"additional_{pick}"
        select_dimension_and_severity(key_prefix=prefix)

        if st.session_state.get(f"{prefix}_dimension_confirmed", False):
            record = {
                "column": pick,
                "dimension": st.session_state[f"{prefix}_dimension"],
                "severity": st.session_state[f"{prefix}_severity"],  # 3 or 4
                "column_type": "additional"
            }

            # Append or update in list (avoid duplicates)
            add_list = st.session_state.get("additional_indicators", [])
            existing_idx = next((i for i, r in enumerate(add_list) if r.get("column") == pick), None)
            if existing_idx is not None:
                add_list[existing_idx] = record
            else:
                add_list.append(record)
            st.session_state["additional_indicators"] = add_list

            # Central mapping by column name
            mapping = st.session_state.get("custom_indicator_mappings", {})
            mapping[pick] = {
                "dimension": record["dimension"],
                "severity": record["severity"],
                "column_type": "additional"
            }
            st.session_state["custom_indicator_mappings"] = mapping

            # Convenience “last picked” keys
            st.session_state['additional_indicator_last_var'] = pick
            st.session_state['additional_indicator_last_severity'] = record["severity"]
            st.session_state['additional_indicator_last_dimension'] = record["dimension"]

            st.success(translations.get("additional_indicator_saved", "Additional indicator saved."))
            return pick


##---------------------------------------------------------------------------------------------------------
def handle_additional_2_selection(current_country, suggestions):
    """
    Optional additional indicator with an explicit on/off toggle.
    When OFF, we clear any previously saved additional indicator state.
    When ON, user can pick a column and set severity (3/4).
    """
    st.markdown(
        translations.get(
            "additional_indicator_title",
            "<b>Optional: Add an additional indicator</b>"
        ),
        unsafe_allow_html=True
    )

    enable = st.checkbox(
        translations.get("additional_indicator_enable", "Add an additional indicator"),
        key="additional_2_indicator_enable"
    )

    if not enable:
        # ✅ Set explicit "no_indicator" instead of None
        st.session_state["additional_2_indicators"] = [{"column": "no_indicator", "severity": None}]
        st.session_state["additional_2_indicator_last_var"] = "no_indicator"
        st.session_state["additional_2_indicator_last_severity"] = None
        st.session_state["additional_2_indicator_last_dimension"] = None

        # Also remove any "additional_2" entries from the central mapping
        mapping = st.session_state.get("custom_indicator_mappings_2", {})
        mapping = {k: v for k, v in mapping.items() if v.get("column_type") != "additional_2"}
        st.session_state["custom_indicator_mappings_2"] = mapping

        st.info(translations.get("additional_2_indicator_disabled_info",
                                  "No additional indicator will be used."))
        return None  # nothing selected

    # --- If enabled: allow picking + assigning severity ---
    pick = st.selectbox(
        translations.get("additional_indicator_pick", "Pick an indicator to add:"),
        ['No selection'] + suggestions,
        key="additional_2_indicator_selectbox"
    )

    if pick == 'No selection':
        st.warning(translations.get("additional_2_indicator_none_selected",
                                    "Select an indicator to continue, or untick the box to skip."))
        return "no_indicator"

    with st.container(border=True):
        st.markdown(
            translations.get(
                "additional_2_indicator_dimsev",
                "Assign this indicator to a dimension & severity:"
            ),
            unsafe_allow_html=True
        )

        prefix = f"additional_2_{pick}"
        select_dimension_and_severity(key_prefix=prefix)

        if st.session_state.get(f"{prefix}_dimension_confirmed", False):
            record = {
                "column": pick,
                "dimension": st.session_state[f"{prefix}_dimension"],
                "severity": st.session_state[f"{prefix}_severity"],  # 3 or 4
                "column_type": "additional_2"
            }

            # Append or update in list (avoid duplicates)
            add_list = st.session_state.get("additional_2_indicators", [])
            existing_idx = next((i for i, r in enumerate(add_list) if r.get("column") == pick), None)
            if existing_idx is not None:
                add_list[existing_idx] = record
            else:
                add_list.append(record)
            st.session_state["additional_2_indicators"] = add_list

            # Central mapping by column name
            mapping = st.session_state.get("custom_indicator_mappings_2", {})
            mapping[pick] = {
                "dimension": record["dimension"],
                "severity": record["severity"],
                "column_type": "additional_2"
            }
            st.session_state["custom_indicator_mappings_2"] = mapping

            # Convenience “last picked” keys
            st.session_state['additional_2_indicator_last_var'] = pick
            st.session_state['additional_2_indicator_last_severity'] = record["severity"]
            st.session_state['additional_2_indicator_last_dimension'] = record["dimension"]

            st.success(translations.get("additional_2_indicator_saved", "Additional indicator saved."))
            return pick


##---------------------------------------------------------------------------------------------------------
def handle_column_selection(suggestions, column_type):
    suggested_column = suggestions[0] if suggestions else "No selection"
    message_template = translations["is_this_individual_column_message"]

    st.write(message_template.format(
        column_type=column_type.replace("_", " ").capitalize(),
        suggested_column=suggested_column
    ))

    col1, col2 = st.columns(2)
    edu_data = st.session_state["edu_data"]
    select_key = f"{column_type}_selectbox"
    show_manual_key = f"show_manual_{column_type}"
    message_placeholder = st.empty()

    # init flag
    if show_manual_key not in st.session_state:
        st.session_state[show_manual_key] = False

    with col1:
        if st.button("Yes/Oui", key=f"confirm_yes_{column_type}"):
            if suggested_column != "No selection":
                st.session_state[f"selected_{column_type}_column"] = suggested_column
                st.session_state[f"{column_type}_column_confirmed"] = True
                st.session_state[show_manual_key] = False  # hide manual UI if previously shown
                message_placeholder.success(
                    f"{column_type.capitalize()} column '{suggested_column}' has been confirmed."
                )

    with col2:
        if st.button("No/Non", key=f"confirm_no_{column_type}"):
            st.session_state[show_manual_key] = True  # persist manual selection UI

    # If user chose "No", render selectbox (persistently) and confirm
    if st.session_state[show_manual_key]:
        st.selectbox(
            f"Select the individual {column_type} column:",
            ["No selection"] + edu_data.columns.tolist(),
            key=select_key
        )

        if st.button(translations.get("confirm_manual", "Confirm"), key=f"confirm_manual_{column_type}"):
            selected_column = st.session_state.get(select_key, "No selection")
            if selected_column != "No selection":
                st.session_state[f"selected_{column_type}_column"] = selected_column
                st.session_state[f"{column_type}_column_confirmed"] = True
                message_placeholder.success(
                    f"{column_type.capitalize()} column '{selected_column}' has been manually selected."
                )
            else:
                message_placeholder.error(translations.get("error_message", "Please select a valid option."))

    # Always return the confirmed selection if available, otherwise suggested (or No selection)
    return st.session_state.get(f"selected_{column_type}_column", suggested_column)

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
def find_barrier_details(barrier_variable, survey_data, choices_data, label_column, country=None):
    """
    Fetch all barrier labels (from choices_data[label_column]) for a given barrier question.

    Matching strategy (in survey_data):
      1) match barrier_variable to survey_data['name']
      2) if not found, match barrier_variable to survey_data[label_column]
         (then use the corresponding survey_data['name'] row)

    Assumes survey_data has columns: 'name', 'type', and label_column
    Assumes choices_data has columns: 'list_name' and label_column
    """

    if barrier_variable in [None, "", "no_indicator"]:
        return []

    if "name" not in survey_data.columns:
        raise KeyError("survey_data must contain a 'name' column.")
    if "type" not in survey_data.columns:
        raise KeyError("survey_data must contain a 'type' column.")
    if label_column not in survey_data.columns:
        raise KeyError(f"survey_data must contain the label column '{label_column}'.")
    if "list_name" not in choices_data.columns:
        raise KeyError("choices_data must contain a 'list_name' column.")
    if label_column not in choices_data.columns:
        raise KeyError(f"choices_data must contain the label column '{label_column}'.")

    v = str(barrier_variable).strip().lower()

    # 1) Try match on NAME
    s_name = survey_data["name"].astype(str).str.strip().str.lower()
    hit = survey_data.loc[s_name == v]

    # 2) If not found, try match on LABEL
    if hit.empty:
        s_lab = survey_data[label_column].astype(str).str.strip().str.lower()
        hit = survey_data.loc[s_lab == v]

    if hit.empty:
        raise KeyError(
            f"Barrier variable '{barrier_variable}' was not found in survey_data['name'] "
            f"nor in survey_data['{label_column}']."
        )

    # Use the first matched row
    type_info = str(hit.iloc[0]["type"]).strip()

    # Expect: "select_one <list>" (or "select_multiple <list>")
    if "select_one" in type_info:
        list_name = type_info.replace("select_one", "").strip()
    elif "select_multiple" in type_info:
        list_name = type_info.replace("select_multiple", "").strip()
    else:
        raise ValueError(
            f"survey_data['type'] for '{barrier_variable}' is '{type_info}', "
            "expected 'select_one <list>' or 'select_multiple <list>'."
        )

    barrier_details = choices_data.loc[choices_data["list_name"].astype(str).str.strip() == list_name]

    # Return barrier labels (drop blanks/NaN)
    out = (
        barrier_details[label_column]
        .dropna()
        .astype(str)
        .str.strip()
    )
    out = out[out != ""].tolist()

    return out
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
        st.session_state.get('displacement_column_confirmed', False) and st.session_state.get('pop_group_value_map_confirmed', False)) :
        st.session_state['other_parameters_confirmed'] = True
    else:
        st.session_state['other_parameters_confirmed'] = False

    display_status("Other Parameters Confirmed", st.session_state['other_parameters_confirmed'])


##---------------------------------------------------------------------------------------------------------
def find_matching_columns(dataframe, keywords):
    return [col for col in dataframe.columns if any(kw in col.lower() for kw in keywords)]
##--------------------------------------------------------------------------------------------------------
def display_filtered_kobo(filtered_edu_kobo, string_second_column = 'label'):
    """Displays the filtered data with enhanced styling."""
    if not filtered_edu_kobo.empty:       
        filtered_edu_kobo = filtered_edu_kobo.reset_index(drop=True)
        
        # Display as a table to remove row numbers completely
        st.dataframe(filtered_edu_kobo.style.set_properties(
            **{'background-color': '#DDE6D5', 'border': '1px solid #4CAF50'}
        ),
        use_container_width=True,  # Adjust the table to fit the container width
        hide_index=True,
        column_config={
            "name" : "Name indicator",
            string_second_column: "Question"
        }
        )
    else:
        st.warning(translations.get('no_data_found', "No matching data found for the selected criteria."))
 
##---------------------------------------------------------------------------------------------------------
def handle_displacement_column_selection():
    if 'household_data' in st.session_state:
        household_data = st.session_state['household_data']
        displacement_keywords = [
            'hh_displaced', 'pop_group', 'i_type_pop', 'statut', 'hh_forcibly_displaced','statut','hoh_dis','l_stratification',
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
##---------------------------------------------------------------------------------------------------------
def _get_status_values(df: pd.DataFrame, col: str):
    """Return sorted unique non-empty string values from a column."""
    if col not in df.columns:
        return []
    s = df[col].dropna().astype(str).str.strip()
    s = s[s != ""]
    # Keep stable ordering-ish but sorted is usually nicest:
    return sorted(s.unique().tolist())

##---------------------------------------------------------------------------------------------------------
def handle_displacement_value_mapping():
    """
    After displacement column is confirmed, show dropdowns to map:
      - host/non-displaced/general population (MANDATORY)
      - idp/displaced (optional)
      - returnee (optional)
      - refugees (optional)
      - other (optional, SINGLE value)
    Saves results into session_state["pop_group_value_map"].
    """
    if not st.session_state.get("displacement_column_confirmed", False):
        return

    df = st.session_state.get("household_data")
    status_col = st.session_state.get("status_var")

    if df is None or not status_col:
        return

    values = _get_status_values(df, status_col)
    if not values:
        st.warning(translations.get(
            "no_status_values_found",
            "No values found in the selected status column."
        ))
        return

    # --- Intro / warning about OCHA template logic ---
    st.info(translations[
        "pop_group_mapping_ocha_logic_note"])

    st.markdown(translations["status_value_mapping_title"], unsafe_allow_html=True)


    # --- init ---
    st.session_state.setdefault("pop_group_value_map", {})
    st.session_state.setdefault("pop_group_value_map_confirmed", False)

    with st.expander(translations["show_status_values"]):
        st.write(values)

    host_label = translations[
        "map_host_label"]
    idp_label = translations[
        "map_idp_label"]
    ret_label = translations[
        "map_returnee_label"]
    ref_label = translations[
        "map_refugee_label"]
    other_label = translations[
        "map_other_label"]
    confirm_lbl = translations["confirm_mapping"]
    error_lbl = translations[
        "mapping_error"]

    # --- dropdowns ---
    host_val = st.selectbox(host_label, ["No selection"] + values, key="map_host_value")
    idp_val  = st.selectbox(idp_label,  ["No selection"] + values, key="map_idp_value")
    ret_val  = st.selectbox(ret_label,  ["No selection"] + values, key="map_returnee_value")
    ref_val  = st.selectbox(ref_label,  ["No selection"] + values, key="map_refugee_value")

    used = {v for v in [host_val, idp_val, ret_val, ref_val] if v and v != "No selection"}
    remaining = [v for v in values if v not in used]

    other_val = st.selectbox(
        other_label,
        ["No selection"] + remaining,
        key="map_other_value"
    )

    # --- confirm ---
    if st.button(confirm_lbl, key="confirm_pop_group_mapping"):
        required_ok = (host_val != "No selection")

        # normalize optionals to None
        idp_norm = None if idp_val == "No selection" else idp_val
        ret_norm = None if ret_val == "No selection" else ret_val
        ref_norm = None if ref_val == "No selection" else ref_val
        oth_norm = None if other_val == "No selection" else other_val

        chosen = [host_val] + [v for v in [idp_norm, ret_norm, ref_norm, oth_norm] if v is not None]
        no_dupes = (len(set(chosen)) == len(chosen))

        if required_ok and no_dupes:
            st.session_state["pop_group_value_map"] = {
                "status_column": status_col,
                "host": host_val,       # required
                "idp": idp_norm,        # optional
                "returnee": ret_norm,   # optional
                "refugee": ref_norm,    # optional
                "other": oth_norm       # optional, SINGLE
            }
            st.session_state["pop_group_value_map_confirmed"] = True
            st.success(translations["mapping_saved"])
        else:
            st.error(error_lbl)

    display_status(
        translations.get("mapping_confirmed_status", "Population group mapping confirmed"),
        st.session_state.get("pop_group_value_map_confirmed", False)
    )



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
                # Extract label columns from the survey data
                label_columns = [col for col in survey_data.columns if col.lower().startswith('label')]
                
                if label_columns:
                    # Define prioritized labels (case-insensitive)
                    prioritized_labels = ['label::english', 'label::french', 'label']
                    selected_label = None
                    
                    # Convert columns to lowercase for comparison
                    label_columns_lower = [col.lower() for col in label_columns]
                    
                    # Find the first matching prioritized label (case-insensitive)
                    for priority_label in prioritized_labels:
                        for idx, label in enumerate(label_columns_lower):
                            if priority_label == label:
                                selected_label = label_columns[idx]  # Preserve original case
                                break
                        if selected_label:
                            break

                    # If no prioritized label is found, select the first available label or ask the user
                    if not selected_label:
                        if len(label_columns) == 1:
                            selected_label = label_columns[0]  # Automatically select the only available label
                        else:
                            # If no match, fallback to asking the user
                            selected_label = st.selectbox(
                                translations["label_json"], 
                                ['No selection'] + label_columns, 
                                key='selected_label'
                            )
                    
                    # Set the selected label in session state
                    if selected_label and selected_label != 'No selection':
                        st.session_state['label'] = selected_label
                        st.session_state.label_selected = True
                        #message_label_success = translations["success_label_kobo"].format(selected_label=selected_label)
                        #st.success(message_label_success)
                        st.markdown(translations["proceed_to_next_step"], unsafe_allow_html=True)
                else:
                    st.warning("No label columns found in the survey data.")

            

    else:
        st.warning(translations["no_data"])     

##-----------------------------
# Function to select indicators
edu_access_strings = [ 'edu_access', 'e_enfant_scolarise_formel', 'enrolled_school', 'Q5_6_1_edu_attendance', 'education_access', 'edu_scolarise']
def select_indicators():
    if 'edu_data' in st.session_state and 'survey_data' in st.session_state and st.session_state.get('label_selected', False) :
        edu_data = st.session_state['edu_data']

        # taking out the edu_kobo
        label =  st.session_state.get('label')
        survey_data = st.session_state['survey_data']
        
        extracted_columns_edu_kobo = survey_data[['name', label]]

        filtered_edu_access = extracted_columns_edu_kobo[extracted_columns_edu_kobo.iloc[:, 0].isin(edu_access_strings)]
        first_match_index = filtered_edu_access.index.min()
        filtered_edu_kobo = extracted_columns_edu_kobo.iloc[first_match_index:first_match_index + 15]


        # Display the translated subheader
        st.subheader(translations["select_variables_and_indicators_subheader"])

        # Display the translated HTML content
        st.markdown(translations["check_variable_warning_html"], unsafe_allow_html=True)

        age_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['age', 'âge'])]
        gender_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['sex', 'gender', 'sexe', 'genre'])]
        education_indicator_suggestions = [col for col in edu_data.columns if any(kw in col.lower() for kw in ['edu', 'education','teacher','Teacher','hazard','natural', 'school', 'ecole', 'scolarise', 'enseignant', 'formel', 'access'])]

        # Checkbox to show/hide the data header
        if st.checkbox(translations["display_education_data_header_checkbox"]):
            st.dataframe(edu_data.head())

        if age_suggestions:
            age_found = handle_column_selection(age_suggestions, 'age')
            st.session_state['age_var'] = age_found
        if gender_suggestions:
            st.session_state['gender_var'] = handle_column_selection(gender_suggestions, 'gender')

        st.markdown(
            """
            <div style="background-color: #f9f9f9; border-left: 2px solid #21B1FF; padding: 1px; margin-bottom: 0px;">
                <h4 style="color: #21B1FF;">{}</h4>
            </div>
            """.format(translations["show_kobo"]),
            unsafe_allow_html=True
        )
        with st.expander(translations["expand_kobo"]):
            display_filtered_kobo(filtered_edu_kobo, label)


        if 'country' in st.session_state and st.session_state['country'] != 'no selection':
            current_country = st.session_state['country']
        if education_indicator_suggestions:
            with st.container(border=True):
                st.session_state['access_var'] = handle_full_selection(current_country,education_indicator_suggestions, 'education_access', translations["access_var_prompt"], translations["access_var_prompt_2"])    
            with st.container(border=True):
                st.session_state['teacher_disruption_var'] =  handle_full_selection(current_country,education_indicator_suggestions, 'disruption_teacher',translations["teacher_disruption_var_prompt"], translations["teacher_disruption_var_prompt_2"]) 
            with st.container(border=True):
                st.session_state['natural_hazard_disruption_var'] =  handle_natural_hazard_disruption_selection(current_country, education_indicator_suggestions) 
            with st.container(border=True):
                st.session_state['idp_disruption_var'] =  handle_full_selection(current_country,education_indicator_suggestions, 'disruption_idp', translations["idp_disruption_var_prompt"], translations["idp_disruption_var_prompt_2"]) 
            with st.container(border=True):
                st.session_state['armed_disruption_var'] =  handle_armed_disruption_selection(current_country, education_indicator_suggestions)  
            with st.container(border=True):
                st.session_state['barrier_var'] = handle_full_selection(current_country, education_indicator_suggestions, 'barriers', translations["barrier_var_prompt"], translations["barrier_var_prompt_2"]) 
            with st.container(border=True):
                st.session_state['additional_var'] =handle_additional_selection(current_country, education_indicator_suggestions)
            with st.container(border=True):
                st.session_state['additional_2_var'] =handle_additional_2_selection(current_country, education_indicator_suggestions)
            
            check_for_duplicate_selections()
        if st.button(translations["confirm_indicators"]):
            st.session_state.indicators_confirmed = True
            st.success(translations["success_indicator"])

            # Display the HTML content
            st.markdown(translations["proceed_to_next_step3"], unsafe_allow_html=True)
            with st.expander("🔎 Debug – Natural hazard indicator"):
                st.write("natural_hazard_disruption_var:", st.session_state.get('natural_hazard_disruption_var'))
                st.write("selected_natural_hazard_disruption_column:", st.session_state.get('selected_disruption_natural_hazard_column'))
                st.write("natural_hazard_disruption_column_confirmed:", st.session_state.get('disruption_natural_hazard_column_confirmed'))
                st.write("natural_hazard_disruption_severity:", st.session_state.get('natural_hazard_disruption_severity'))


    else:
        st.warning(translations["no_data"]) 
##-----------------------------
# Function to define severity of barriers
def define_severity(country):
    if 'survey_data' in st.session_state and 'choice_data' in st.session_state and 'edu_data' in st.session_state and st.session_state.get('indicators_confirmed', False):
        survey_data = st.session_state['survey_data']
        choices_data = st.session_state['choice_data']
        edu_data = st.session_state['edu_data']

        barrier_var = st.session_state.get('barrier_var', 'Default Value if not set')
        selected_label = st.session_state['label'] 
        st.write(selected_label)

        barrier_options = find_barrier_details(barrier_var, survey_data, choices_data, selected_label, country=country)

        # Encapsulate descriptions within a single box with a light gray background
        st.markdown(translations["severity_circumstances_html"], unsafe_allow_html=True)
        with st.container(border=True):
            selected_severity_4_barriers = select_severity_barriers(barrier_options, 4)
        barrier_options_5 = [option for option in barrier_options if option not in selected_severity_4_barriers]
        with st.container(border=True):
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

        with st.container(border=True):
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
                <div style="font-size:16px; font-weight:bold;">
                    {translations['mismatch_admin_message']}<br>
                    <span style="color:darkred;">{translations['mismatch_admin_example']}</span>
                </div>
                """, unsafe_allow_html=True
            )
            scope_fix = st.session_state.get('scope_fix', False)

            # Display a simple checkbox below the message
            mismatch_admin_checkbox = st.checkbox(translations['check_box'])
            if mismatch_admin_checkbox:
                            if not scope_fix:
                                st.error(
                                    f"### {translations['scope_fix_empty_error_title']}"
                                    f"\n\n{translations['scope_fix_empty_error_message']}"
                                    f"\n\n⚠️ **{translations['scope_fix_warning']}**"
                                    )

            
            if st.button(translations["confirm_admin"], key='confirm_admin_level'):
                if admin_target != 'No selection':
                    st.session_state['admin_var'] = admin_target
                    st.session_state.admin_level_confirmed = True
                    if mismatch_admin_checkbox:
                        if scope_fix:
                            st.session_state['mismatch_admin'] = True
                        else:
                            st.session_state['mismatch_admin'] = False

                    success_message_admin=  translations["success_admin"].format(admin_target=admin_target)
        
                    st.success(success_message_admin)
                else:
                    st.error("Please select a valid administrative level.")


        #st.markdown("---")  # Markdown horizontal rule


        months = ['No selection','January', 'February', 'March', 'April', 'May', 'June', 
                'July', 'August', 'September', 'October', 'November', 'December']
        with st.container(border=True):
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
                else:
                    st.error("Please select a valid month.")
                #update_other_parameters_status()



        ## -------------------- school cycle -----------------------------------
        
        with st.container(border=True):
            school_cycle_question = translations["school_cycle_question"]["question"]
            option_two = translations["school_cycle_question"]["option_two"]
            option_three = translations["school_cycle_question"]["option_three"]
            # Render the question and options in two columns
            col1, col2 = st.columns([1, 1])  # Adjust width proportions as needed
            school_cycle_count = 0
            # First column: Question and explanation
            with col1:
                school_cycle_question = translations["school_cycle_question"]["question"]
                option_two = translations["school_cycle_question"]["option_two"]
                option_three = translations["school_cycle_question"]["option_three"]

                st.markdown(
                    f"""
                    <div style="margin-bottom: 0;">
                        <strong style="font-size: 18px;">{school_cycle_question}</strong>
                        <p style="font-size: 16px; color: #333; margin-top: 5px;">
                            {option_two}<br>
                            {option_three}
                        </p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            # Second column: Radio button
            with col2:
                school_cycle_count = st.radio(
                    label="",  # Leave empty as the question is rendered in col1
                    options=[2, 3],
                    index=0,
                    key="school_cycle_count"
                )





            if school_cycle_count == 3:
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


                    lower_primary_end = st.session_state['lower_primary_end'] 
                    upper_primary_start = lower_primary_end +1                   

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

                if st.button(translations["school_confirm_3"]):
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

            elif school_cycle_count == 2:
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

                if st.button(translations["school_confirm_2"]):
                    st.session_state.upper_primary_end_confirmed = True
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



        with st.container(border=True):    
            handle_displacement_column_selection()
        with st.container(border=True):
            handle_displacement_value_mapping()

        if st.button(translations["last_confirm"]):
            st.session_state.final_confirmed = True
            #st.success("All details confirmed and finalized!")

##-----------------------------
# Function to handle administrative details and final confirmations
def finalize_details_nomsna():

    st.subheader(translations["choose_disaggregation_variables_subheader"])
    #admin_level_options = ['No selection', 'Admin0', 'Admin1', 'Admin2', 'Admin3']

    with st.container(border=True):
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


        
        if st.button(translations["confirm_admin"], key='confirm_admin_level'):
            if admin_target != 'No selection':
                st.session_state['admin_var'] = admin_target
                st.session_state.admin_level_confirmed = True
                success_message_admin=  translations["success_admin"].format(admin_target=admin_target)
    
                st.success(success_message_admin)
            else:
                st.error("Please select a valid administrative level.")


    #st.markdown("---")  # Markdown horizontal rule


    months = ['No selection','January', 'February', 'March', 'April', 'May', 'June', 
            'July', 'August', 'September', 'October', 'November', 'December']
    with st.container(border=True):
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
            else:
                st.error("Please select a valid month.")
            #update_other_parameters_status()



    ## -------------------- school cycle -----------------------------------
    
    with st.container(border=True):
        school_cycle_question = translations["school_cycle_question"]["question"]
        option_two = translations["school_cycle_question"]["option_two"]
        option_three = translations["school_cycle_question"]["option_three"]
        # Render the question and options in two columns
        col1, col2 = st.columns([1, 1])  # Adjust width proportions as needed
        school_cycle_count = 0
        # First column: Question and explanation
        with col1:
            school_cycle_question = translations["school_cycle_question"]["question"]
            option_two = translations["school_cycle_question"]["option_two"]
            option_three = translations["school_cycle_question"]["option_three"]

            st.markdown(
                f"""
                <div style="margin-bottom: 0;">
                    <strong style="font-size: 18px;">{school_cycle_question}</strong>
                    <p style="font-size: 16px; color: #333; margin-top: 5px;">
                        {option_two}<br>
                        {option_three}
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )

        # Second column: Radio button
        with col2:
            school_cycle_count = st.radio(
                label="",  # Leave empty as the question is rendered in col1
                options=[2, 3],
                index=0,
                key="school_cycle_count"
            )





        if school_cycle_count == 3:
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


                lower_primary_end = st.session_state['lower_primary_end'] 
                upper_primary_start = lower_primary_end +1                   

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

            if st.button(translations["school_confirm_3"]):
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

        elif school_cycle_count == 2:
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

            if st.button(translations["school_confirm_2"]):
                st.session_state.upper_primary_end_confirmed = True
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


    if st.button(translations["last_confirm"]):
        st.session_state.final_confirmed = True
        #st.success("All details confirmed and finalized!")



        #st.markdown("---")
  
###########################################################################################################
###########################################################################################################

def display_step_content():
    data_combination = st.session_state.get('data_combination', "")
    selected_country = st.session_state['country']
    # Check if 'm' is in the data_combination string
    if 'm' in data_combination:
        if st.session_state['current_step'] == 0:
            upload_and_select_data()
        elif st.session_state['current_step'] == 1:
            select_indicators()
        elif st.session_state['current_step'] == 2:
            define_severity(country=selected_country)
        elif st.session_state['current_step'] == 3:
            finalize_details()
    else:
        # Skip directly to final step if 'm' is absent
        st.session_state['current_step'] = 0  # Ensure stepper starts at finalize_details()
        st.session_state['data_selections_confirmed'] = True
        st.session_state['age_column_confirmed'] = True
        st.session_state['label_selected'] = True
        st.session_state['gender_column_confirmed'] = True
        st.session_state['severity_4_confirmed'] = True
        st.session_state['severity_5_confirmed'] = True
        st.session_state['education_access_column_confirmed'] = True
        st.session_state['disruption_teacher_column_confirmed'] = True
        st.session_state['disruption_idp_column_confirmed'] = True
        st.session_state['disruption_armed_column_confirmed'] = True
        st.session_state['barriers_column_confirmed'] = True
        st.session_state['indicators_confirmed'] = True
        st.session_state['displacement_column_confirmed'] = True
        st.session_state['pop_group_value_map_confirmed'] = True

        finalize_details_nomsna()  # Call modified finalize_details directly

display_step_content()


st.markdown("---")  # Markdown horizontal rule

all_steps_confirmed = all([
    st.session_state.get('data_selections_confirmed', False),
    st.session_state.get('label_selected', False),
    st.session_state.get('age_column_confirmed', False),
    st.session_state.get('gender_column_confirmed', False),
    st.session_state.get('indicators_confirmed', False),
    st.session_state.get('severity_4_confirmed', False),
    st.session_state.get('severity_5_confirmed', False),
    st.session_state.get('other_parameters_confirmed', False)
])

#st.write(all_steps_confirmed)

if all_steps_confirmed:
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


# Always show status indicators
# Always show status indicators
display_status("Data Selections Confirmed", st.session_state.data_selections_confirmed)
display_status("Label Selected", st.session_state.label_selected)
display_status("Age Column Confirmed", st.session_state.age_column_confirmed)
display_status("Gender Column Confirmed", st.session_state.gender_column_confirmed)
update_combined_indicator()
display_combined_severity_status()
update_other_parameters_status()


