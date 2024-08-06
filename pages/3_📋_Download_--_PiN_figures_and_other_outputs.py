import streamlit as st
import numpy as np
import pandas as pd
#from backup import provatest

st.logo('pics/logos.png')

st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico',  layout='wide')

if 'password_correct' not in st.session_state:
    st.error('Please Login from the Home page and try again.')
    st.stop()



edu_data = st.session_state['edu_data']
country =  st.session_state['country']
age_var = st.session_state['age_var']
start_school = 'May'

#provatest (country, edu_data, age_var, start_school)
