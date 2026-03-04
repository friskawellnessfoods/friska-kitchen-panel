import streamlit as st
import subprocess
import os
import sys
from datetime import datetime

st.set_page_config(page_title="Friska Kitchen Panel", layout="centered")

# ---------- UI STYLE ----------
st.markdown("""
<style>

.block-container {
    max-width:700px;
}

.logo-container {
    text-align:center;
    margin-bottom:10px;
}

div.stButton > button {
    width:100%;
    padding:14px;
    font-size:16px;
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

# ---------- LOGO ----------
if os.path.exists("logo.png"):
    st.image("logo.png", width=160)

# ---------- HEADER ----------
st.title("Friska Daily Kitchen Panel")

date = st.date_input("Select Date")
date_str = date.strftime("%Y-%m-%d")

# ---------- GENERATOR ----------
def run(mode):

    with st.spinner("Preparing file..."):

        subprocess.run(
            [sys.executable, "daily_list_pc_version.py", date_str, mode],
            capture_output=True,
            text=True
        )

    date_obj = datetime.strptime(date_str,"%Y-%m-%d")
    file_date = date_obj.strftime("%d-%b-%y")

    filename = f"{file_date} list.pdf"

    if os.path.isfile(filename):

        st.success("File Ready")

        with open(filename,"rb") as f:

            st.download_button(
                "⬇ Download PDF",
                data=f,
                file_name=filename,
                mime="application/pdf"
            )

    else:

        st.error("PDF not generated.")

# ---------- BUTTON GRID ----------
col1, col2 = st.columns(2)

with col1:

    if st.button("Full List"):
        run("full")

    if st.button("Meal List"):
        run("meals")

    if st.button("Tags"):
        run("tags")

with col2:

    if st.button("Client List"):
        run("client")

    if st.button("Delivery List"):
        run("delivery")
