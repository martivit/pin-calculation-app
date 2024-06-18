import streamlit as st
import pandas as pd
import time


st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico',  layout='wide')
st.title('Upload MSNA and OCHA data')


# Define the function at the beginning of your script
def check_conditions_and_proceed():
    if selected_country != 'no selection' and 'uploaded_data' in st.session_state:
        if has_ocha_data:
            if 'uploaded_ocha_data' in st.session_state:
                st.session_state.ready_to_proceed = True
            else:
                st.session_state.ready_to_proceed = False
                st.warning("Please upload the OCHA data to proceed.")
        else:
            st.session_state.ready_to_proceed = True
    else:
        st.session_state.ready_to_proceed = False
        if selected_country == 'no selection':
            st.warning("Please select a valid country to proceed.")
        else:
            st.warning("Please upload the MSNA data to proceed.")

    # Display success message if ready to proceed
    if st.session_state.get('ready_to_proceed', False):
        st.success("You have completed all necessary steps. Please proceed to the next page from the sidebar.")


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

# Checkbox to ask if the user has OCHA data
has_ocha_data = st.checkbox('I want to upload OCHA population data', value=st.session_state.get('has_ocha_data', False))
st.session_state['has_ocha_data'] = has_ocha_data

if has_ocha_data:
    # Display an example DataFrame structure for OCHA data
    example_data = {
        "admin": ["Region A", "Region B"],
        "adminPcode": ["A01", "B01"],
        "tot Population": [10000, 20000],
        "boys (6-17)": [5000, 10000],
        "girls (6-17)": [5000, 10000]
    }
    example_df = pd.DataFrame(example_data)
    st.write("Example of the required OCHA population data structure:")
    st.dataframe(example_df)

    if 'uploaded_ocha_data' in st.session_state:
        ocha_data = st.session_state['uploaded_ocha_data']
        st.write("OCHA Data already uploaded.")
        st.dataframe(ocha_data.head())  # Show a preview of the data
    else:
        # OCHA data uploader appears only if checkbox is checked
        uploaded_ocha_file = st.file_uploader("Upload OCHA population data file", type=["csv", "xlsx"])
        if uploaded_ocha_file is not None:
            ocha_data = pd.read_excel(uploaded_ocha_file, engine='openpyxl')
            st.session_state['uploaded_ocha_data'] = ocha_data
            st.success("OCHA Data uploaded successfully!")



# Check conditions to allow proceeding
check_conditions_and_proceed()

st.page_link("pages/2_📊_Calculation_--_PiN.py", label="If you have successfully uploaded the necessary data, proceed to the PiN Calculation page 	:arrow_forward:", icon='📊')
