from openai import OpenAI
import streamlit as st

st.title("OpenAI API Connection Test")

try:
    api_key = st.secrets["OPENAI_API_KEY"]

    # Do not display the actual key
    st.success("OPENAI_API_KEY found in Streamlit secrets.")

except Exception as e:
    st.error("OPENAI_API_KEY was not found in Streamlit secrets.")
    st.error(str(e))
    st.stop()


try:
    client = OpenAI(api_key=api_key)

    st.info("Connecting to OpenAI...")

    response = client.responses.create(
        model="gpt-4o-mini",
        input="Reply with exactly: API WORKS"
    )

    st.success("API CONNECTION SUCCESSFUL")
    st.write(response.output_text)

except Exception as e:
    st.error("API CONNECTION FAILED")
    st.write("Error type:", type(e).__name__)
    st.write("Error:", str(e))
    st.write("Cause:", repr(e.__cause__))