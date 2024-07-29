import streamlit as st
import numpy as np
import pandas as pd
st.logo('pics/logos.png')

st.set_page_config(page_icon='icon/global_education_cluster_gec_logo.ico',  layout='wide')

if 'password_correct' not in st.session_state:
    st.error('Please Login from the Home page and try again.')
    st.stop()


hide_github_icon = """ 
    #GithubIcon { visibility: hidden; } 
    """
st.markdown(hide_github_icon, unsafe_allow_html=True)

hide_streamlit_style = """
    <style>
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    </style>
    """ 
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

