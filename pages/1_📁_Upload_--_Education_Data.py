import streamlit as st
import pandas as pd
import time


st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico',  layout='wide')
st.title('Upload MSNA and OCHA data')


def check_conditions_and_proceed():
    if selected_country != 'no selection' and 'uploaded_data' in st.session_state:
        if 'uploaded_ocha_data' in st.session_state:
            st.session_state.ready_to_proceed = True
        else:
            st.session_state.ready_to_proceed = False
            st.warning("Please upload the OCHA data to proceed.")
    else:
        st.session_state.ready_to_proceed = False
        if selected_country == 'no selection':
            st.warning("Please select a valid country to proceed.")
        else:
            st.warning("Please upload the MSNA data to proceed.")

    # Display success message if ready to proceed
    if st.session_state.get('ready_to_proceed', False):
        st.success("You have completed all necessary steps!")


# Country selection setup
countries = ['no selection',
    'Afghanistan -- AFG', 'Burkina Faso -- BFA', 'Central African Republic -- CAR', 
    'Democratic Republic of the Congo -- DRC', 'Haiti -- HTI', 'Iraq -- IRQ', 'Kenya -- KEN', 
    'Bangladesh -- BGD', 'Lebanon -- LBN', 'Moldova -- MDA', 'Mali -- MLI', 'Mozambique -- MOZ', 
    'Myanmar -- MMR', 'Niger -- NER', 'Syria -- SYR', 'Ukraine -- UKR', 'Somalia -- SOM'
]
selected_country = st.selectbox(
    'Which country do you want to calculate the PiN for?',
    countries,
    index=countries.index(st.session_state.get('country', 'no selection'))
)
st.session_state['country'] = selected_country



# Check if data already uploaded and preserved in session state
if 'uploaded_data' in st.session_state:
    data = st.session_state['uploaded_data']
    st.write("MSNA Data already uploaded. If you want to change the data, just refresh 🔄 the page")
else:
    # MSNA data uploader
    uploaded_file = st.file_uploader("Upload your input MSNA file.", type=["csv", "xlsx"])

    if uploaded_file is not None:
        st.write('This process may take some time; please wait a bit longer. 🐌🐌🐌🐌')
        bar = st.progress(0)

        try:
            # Increment progress for file load initialization
            bar.progress(10)
            time.sleep(0.1)  # simulate delay for starting the load

            # Actual data loading
            data = pd.read_excel(uploaded_file, engine='openpyxl', sheet_name=None)
            st.session_state['uploaded_data'] = data
            # Update progress post data load
            bar.progress(50)
            time.sleep(0.1)  # simulate delay for post-load processing

            # Finalize the loading process
            bar.progress(100)
            st.success("MSNA Data uploaded successfully!")
        except Exception as e:
            st.error(f"Failed to process the uploaded file: {e}")
            bar.progress(0)


st.markdown("---")  # Markdown horizontal rule

# Provide a message and download button for the OCHA data template with a highlighted section
st.markdown(
    """
    <div style="background-color: #e6f7ff; padding: 10px; border-radius: 5px;">
        <p style="color: #00529B;">
            You can download the template and fill it with the population figures provided by OCHA.
            Please ensure that you follow the template format.
        </p>
         <p style="color: red; margin-top: 0;">
            YELLOW COLUMNS ARE MANDATORY
        </p>
    </div>
    """, unsafe_allow_html=True
)

# Function to load the existing template from the file system
def load_template():
    with open('input/Template_Population_figures.xlsx', 'rb') as f:
        template = f.read()
    return template

# Add a download button for the existing template
st.download_button(label="Download OCHA Population Data Template",
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

    # Prepare the data by replacing NaN with 0 and rounding to the nearest whole number
    filled_ocha_data = ocha_data.fillna(0)

    # Prepare to collect error messages
    errors = []

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
    if 'uploaded_ocha_data' in st.session_state:
        ocha_data = st.session_state['uploaded_ocha_data']
        st.write("OCHA Data already uploaded.")
        st.dataframe(ocha_data.head())  # Show a preview of the data
    else:
        # OCHA data uploader
        uploaded_ocha_file = st.file_uploader("Upload OCHA population data file", type=["csv", "xlsx"])
        if uploaded_ocha_file is not None:
            ocha_data = pd.read_excel(uploaded_ocha_file, engine='openpyxl')
            check_message = perform_ocha_data_checks(ocha_data)
            if check_message == "Data is valid":
                st.session_state['uploaded_ocha_data'] = ocha_data
                st.success("OCHA Data uploaded successfully!")
            else:
                st.error(check_message)  # Display the error message if checks fail

# Check conditions to allow proceeding
check_conditions_and_proceed()

col1, col2 = st.columns([0.60, 0.40])

with col2: 
    st.page_link("pages/2_📊_Calculation_--_PiN.py", label="Proceed to the PiN Calculation page 	:arrow_right:", icon='📊')
