# shared_utils.py

import streamlit as st
import json

# Load translation files
def load_translation(language):
    try:
        with open(f"{language}.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        st.error("Translation file not found.")
        st.stop()

# Function to handle the language selection in the sidebar
def language_selector():
    # Available languages
    languages = {"English": "en", "French": "fr"}

    # Initialize session state for the selected language
    if "selected_language" not in st.session_state:
        st.session_state.selected_language = "English"  # Default to English

    # Sidebar language selector
    language = st.sidebar.selectbox("Select language", list(languages.keys()), index=list(languages.keys()).index(st.session_state.selected_language))

    # Update the selected language in session state
    st.session_state.selected_language = language

    # Get the selected language code
    lang_code = languages[st.session_state.selected_language]

    # Load the translations for the selected language
    translations = load_translation(lang_code)

    # Store the translations in session state
    st.session_state.translations = translations
