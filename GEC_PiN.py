import streamlit as st
from hashlib import sha256
import hmac
import json
from shared_utils import language_selector


st.logo('pics/GEC Global English logo_Colour_JPEG.jpg')


st.set_page_config(
    page_title="People in Need (PiN) and severity methodology",
    page_icon='icon/global_education_cluster_gec_logo.ico',
    layout='wide')


# Call the language selector function
language_selector()

# Access the translations
translations = st.session_state.translations


## ----- user authenthicator ------
# Define user credentials
try:
    users = {
        "vit": st.secrets["users"]["vit"]["pwd"],
        "grand": st.secrets["users"]["grand"]["pwd"],
        "impact": st.secrets["users"]["impact"]["pwd"],
        "cluster": st.secrets["users"]["cluster"]["pwd"],
        "retreat2024": st.secrets["users"]["retreat2024"]["pwd"]
    }
except KeyError as e:
    st.error(f"Error loading user credentials: {e}")
    st.stop()

SECRET_KEY = b'supersecretkey'
def hmac_hash(password):
    return hmac.new(SECRET_KEY, password.encode('utf-8'), 'sha256').hexdigest()

# Utility function to validate passwords
def check_password():
    """Returns `True` if the user has entered a correct password."""

    def login_form():
        """Form with widgets to collect user information"""
        with st.form("Credentials"):
            st.text_input("Username", key="username")
            st.text_input("Password", type="password", key="password")
            st.form_submit_button("Log in", on_click=password_entered)

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        user = st.session_state["username"]
        if user in users:
            entered_password_hash = hmac_hash(st.session_state["password"])
            # Compare the entered password with the stored password
            if hmac.compare_digest(entered_password_hash, users[user]):
                st.session_state["password_correct"] = True
                del st.session_state["password"]  # Don't store the username or password.
                del st.session_state["username"]
            else:
                st.session_state["password_correct"] = False
        else:
            st.session_state["password_correct"] = False

    # Return True if the username + password is validated.
    if st.session_state.get("password_correct", False):
        return True

    # Show inputs for username + password.
    login_form()
    if "password_correct" in st.session_state and not st.session_state["password_correct"]:
        st.error("😕 User not known or password incorrect")
        st.info("In case of issues, write to unit@email")

    return False


#if not check_password():
#    st.stop()

spacer1, spacer2 = st.sidebar.empty(), st.sidebar.empty()




st.image('pics/pinheader.jpg')


st.write(translations["Title"])

st.markdown(translations["methodology"])