import streamlit as st

# Page configuration
st.set_page_config(page_title="GEC PiN", layout="wide")

# Path to your logos
logo_path_1 = 'pics/GEC_logo.png'
logo_path_2 = 'pics/impact_initiatives_logo.jpeg'

# Display logos at the top of the sidebar
st.sidebar.image(logo_path_1, width=80)  # Adjust width as needed
st.sidebar.image(logo_path_2, width=80)  # Adjust width as needed

# Sidebar title
st.sidebar.title("GEC PiN")

# Navigation or other sidebar elements
st.sidebar.header("Navigation")
st.sidebar.button("Upload -- Education Indicators")
st.sidebar.button("Calculation -- PiN")
st.sidebar.button("Download -- PiN figures")

# Page contents would be handled by the page scripts




# Check and display whether the user has OCHA data
if st.session_state.get('has_ocha_data', False):
    st.write("User has uploaded OCHA data.")
    # Optionally display the OCHA data
    if 'uploaded_ocha_data' in st.session_state:
        st.dataframe(st.session_state['uploaded_ocha_data'])
else:
    st.write("User has not uploaded OCHA data.")



# Additional inquiries with selectboxes
if 'household_data' in st.session_state:
    # Example of another inquiry
    column_names = household_data.columns.tolist()
    selected_variable = st.selectbox('Select a variable to analyze:', column_names, key='selected_variable')
    st.session_state['selected_variable'] = selected_variable
    # Display some data based on the selected variable
    if selected_variable:
        st.write(f"Data for selected variable - {selected_variable}:")
        st.write(household_data[selected_variable].head())




