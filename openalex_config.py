# openalex_config.py
import streamlit as st

def get_openalex_api_key():
    """Returns the OpenAlex key entered by the user at login."""
    return st.session_state.get("openalex_api_key", "") or None