import streamlit as st
#import numpy as np
#import pandas as pd
import streamlit_authenticator as stauth
import pickle
from pathlib import Path

st.logo('pics/logos.png')


st.set_page_config(
    page_title="People in Need (PiN) and severity methodology",
    page_icon='icon/global_education_cluster_gec_logo.ico',
    layout='wide')


## ---- USER AUTHENTICATION --- 
names = ['Martina Vit', "Marie-Amandine Grand"]
usernames = ['mvit','mgrand']
credentials = {"usernames":{}}

file_path = Path(__file__).parent / 'hashed_pw.pkl'
with file_path.open("rb")as file:
    hashed_passwords = pickle.load(file)


authenticator = stauth.Authenticate(names, usernames, hashed_passwords,
    "Pin_tool_platform", "abcdef", cookie_expiry_days=30)

name, authentication_status, username = authenticator.login("Login", "main")


#if "shared" not in st.session_state:
   #st.session_state["shared"] = True

if authentication_status == False:
    st.session_state.authentication_status = False
    st.error("Username/password is incorrect")

if authentication_status == None:
    st.session_state.authentication_status = None
    st.warning("Please enter your username and password, if you don't have them contact martina.vit@impact.initiatives.org")

if authentication_status:
    st.session_state.authentication_status = True




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
