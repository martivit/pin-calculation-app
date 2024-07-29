import streamlit as st
#import numpy as np
#import pandas as pd
import streamlit_authenticator as stauth
from pathlib import Path
import hmac
from hashlib import sha256



st.logo('pics/logos.png')


st.set_page_config(
    page_title="People in Need (PiN) and severity methodology",
    page_icon='icon/global_education_cluster_gec_logo.ico',
    layout='wide')


## ----- user authenthicator ------
# Load the configuration from 'test.yaml' safely
def check_password():
    """Returns `True` if the user had a correct password."""
 
    def login_form():
        """Form with widgets to collect user information"""
        with st.form("Credentials"):
            st.text_input("Username", key="username")
            st.text_input("Password", type="password", key="password")
            st.form_submit_button("Log in", on_click=password_entered)
 
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        h = sha256()
        pasw = st.session_state["password"]
        h.update(pasw.encode('utf-8'))
        hash = h.hexdigest()
        if st.session_state["username"] in st.secrets[
            "passwords"
          ] and hash in st.secrets.passwords[st.session_state["username"]]:
        # and hmac.compare_digest(   
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store the username or password.
            del st.session_state["username"]
        else:
            st.session_state["password_correct"] = False
 
    # Return True if the username + password is validated.
    if st.session_state.get("password_correct", False):
        return True
 
    # Show inputs for username + password.
    login_form()
    if "password_correct" in st.session_state:
        st.error("😕 User not known or password incorrect")
    return False
 
 
if not check_password():
    st.stop()




spacer1, spacer2 = st.sidebar.empty(), st.sidebar.empty()

for _ in range(150):  # You might need to adjust this number
    spacer1.write("")
    spacer2.write("")
logo_path_1 = 'pics/GEC_logo.png'
logo_path_2 = 'pics/IMPACT_Logo_TransparentBackground_WH.png'
st.sidebar.image(logo_path_1, width=200)  # Adjust width as necessary
st.sidebar.image(logo_path_2, width=200)  # Adjust width as necessary


st.image('pics/pinheader.jpg')


st.write("# Calculating Education Cluster People in Need (PiN)")

st.markdown(
    """

    Calculating the People in Need (PiN) for the Education sector within a Humanitarian Needs Overview (HNO)
    involves aggregating data on educational needs from various geographic regions and groups affected by crises.

    This process, guided by UN's OCHA, helps in planning and resource allocation by detailing the size and severity of educational
    needs across different areas and populations, including displaced persons and locals.
    The PiN calculation provides crucial insights into where and how resources should be distributed,
    with data broken down by geographic location, affected group, severity of educational conditions, and gender.

    Full guidance can be found here: 

    ### How to use this PiN calculator tool?

    **👈 Following the order shown in the sidebar** 
    
    - Select your country, **upload the MSNA data** and the OCHA population data (if present);
    - Check and pick the education indicators and sort out the education barriers according to the severity category;
    - Calculate the PiN!
    - Download the final tables. If OCHA population figures were provided, the results will show the final number of People in Need.
    - The results are disaggregated
    per: 
        - Gender
        - Geographical area
        - Different affected groups as defined in the humanitarian profile of the country (IDPs, residents, returnees, refugees)
        - Severity of education conditions, determined by the degree of unmet needs across the four dimensions
        - PiN dimensions

    #### Questions? Issues?
    You can contact:
"""
)