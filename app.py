import streamlit as st
import subprocess
import os
import sys
from datetime import datetime

st.set_page_config(page_title="Friska Kitchen Panel", layout="centered")

# -------- UI STYLE --------
st.markdown("""
<style>
.block-container {max-width:700px;}

div.stButton > button {
    width:100%;
    padding:16px;
    font-size:18px;
    border-radius:12px;
    border:none;
    background:#1976d2;
    color:white;
    font-weight:600;
}

div.stButton > button:hover {
    background:#1565c0;
}
</style>
""", unsafe_allow_html=True)

# -------- LOGO --------
if os.path.exists("logo.png"):
    st.image("logo.png", width=160)

st.title("Friska Daily Kitchen Panel")

date = st.date_input("Select Date")

date_str = date.strftime("%Y-%m-%d")


def generate():

    with st.spinner("Preparing kitchen list..."):

        subprocess.run(
            [sys.executable, "daily_list_pc_version.py", date_str],
            capture_output=True,
            text=True
        )

    file_date = datetime.strptime(date_str,"%Y-%m-%d").strftime("%d-%b-%y")
    filename = f"{file_date} list.pdf"

    if os.path.isfile(filename):

        with open(filename,"rb") as f:

            st.success("Kitchen list ready")

            st.download_button(
                "Download Full Kitchen List",
                data=f,
                file_name=filename,
                mime="application/pdf"
            )

    else:

        st.error("PDF not generated.")
        st.write("Files present:", os.listdir("."))


if st.button("Generate Kitchen List"):
    generate()
