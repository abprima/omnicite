# app.py
# OmniCite Auditor - Multi-Style Citation Checker
# Run: streamlit run app.py

import streamlit as st

st.set_page_config(
    page_title="OmniCite Auditor",
    page_icon="📚",
    layout="wide",
)

# =========================================================
# SHARED CONFIG
# =========================================================

if "openalex_api_key" not in st.session_state:
    st.session_state["openalex_api_key"] = ""
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# =========================================================
# GLOBAL BUTTON STYLING (blue = extract, green = AI)
# =========================================================

st.markdown(
    """
    <style>
    /* ---- BLUE buttons: Extract & Review ---- */
    div[class*="st-key-blue_btn"] button {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border: 1px solid #1d4ed8 !important;
        font-weight: 600 !important;
        transition: background-color 0.15s ease;
    }
    div[class*="st-key-blue_btn"] button:hover {
        background-color: #1d4ed8 !important;
        color: #ffffff !important;
        border-color: #1e40af !important;
    }
    div[class*="st-key-blue_btn"] button:focus {
        box-shadow: 0 0 0 0.2rem rgba(37, 99, 235, 0.4) !important;
    }

    /* ---- GREEN buttons: Automated Processing Check ---- */
    div[class*="st-key-green_btn"] button {
        background-color: #16a34a !important;
        color: #ffffff !important;
        border: 1px solid #15803d !important;
        font-weight: 600 !important;
        transition: background-color 0.15s ease;
    }
    div[class*="st-key-green_btn"] button:hover {
        background-color: #15803d !important;
        color: #ffffff !important;
        border-color: #166534 !important;
    }
    div[class*="st-key-green_btn"] button:focus {
        box-shadow: 0 0 0 0.2rem rgba(22, 163, 74, 0.4) !important;
    }

    /* ---- GREEN download buttons ---- */
    div[data-testid="stDownloadButton"] > button {
        background-color: #16a34a !important;
        color: #ffffff !important;
        border: 1px solid #15803d !important;
        font-weight: 600 !important;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        background-color: #15803d !important;
        color: #ffffff !important;
        border-color: #166534 !important;
    }

    /* ---- Centered dataframe text ---- */
    div[data-testid="stDataFrame"] th,
    div[data-testid="stDataFrame"] td {
        text-align: center !important;
    }
    div[data-testid="stDataFrame"] th > div,
    div[data-testid="stDataFrame"] td > div {
        justify-content: center !important;
        text-align: center !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# LOGIN GATE
# =========================================================

def login_screen():
    st.title("OmniCite Auditor")
    st.caption("Multi-style citation integrity checker — APA 7 · IEEE · Chicago")

    with st.form("login_form"):
        password = st.text_input("Access password", type="password")
        openalex_key = st.text_input(
            "OpenAlex API key",
            type="password",
            help="Get a free key at https://openalex.org/ (Settings → API key)",
        )
        submitted = st.form_submit_button("Enter", use_container_width=True)

    if submitted:
        if password != "12345":
            st.error("Invalid password.")
            return
        if not openalex_key.strip():
            st.error("OpenAlex API key is required.")
            return
        st.session_state["authenticated"] = True
        st.session_state["openalex_api_key"] = openalex_key.strip()
        st.rerun()


if not st.session_state["authenticated"]:
    login_screen()
    st.stop()

# =========================================================
# STYLE SELECTOR
# =========================================================

st.sidebar.title("OmniCite Auditor")
st.sidebar.caption("Choose a citation style")

style = st.sidebar.radio(
    "Citation style",
    ["APA 7th Edition", "IEEE", "Chicago Notes & Bibliography"],
    key="style_selector",
)

st.sidebar.divider()
st.sidebar.caption(f"OpenAlex key: `...{st.session_state['openalex_api_key'][-6:]}`")

if st.sidebar.button("Log out", use_container_width=True):
    st.session_state["authenticated"] = False
    st.session_state["openalex_api_key"] = ""
    st.rerun()

# =========================================================
# DISPATCH
# =========================================================

if style == "APA 7th Edition":
    import apa
    apa.render()
elif style == "IEEE":
    import ieee
    ieee.render()
else:
    import chicago
    chicago.render()