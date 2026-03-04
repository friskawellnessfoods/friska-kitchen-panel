import streamlit as st
import subprocess
import os
from datetime import datetime

st.set_page_config(page_title="Friska Kitchen Panel", layout="centered")

st.title("Friska Daily Kitchen Panel")

date = st.date_input("Select Date")

date_str = date.strftime("%Y-%m-%d")

def run(mode):

    with st.spinner("Generating PDF..."):

        result = subprocess.run(
            ["python", "daily_list_pc_version.py", date_str, mode],
            capture_output=True,
            text=True
        )

        st.text(result.stdout)

    date_obj = datetime.strptime(date_str,"%Y-%m-%d")
    file_date = date_obj.strftime("%d-%b-%y")

    filename = f"{file_date} list.pdf"

    if os.path.exists(filename):

        st.success("PDF Ready")

        with open(filename,"rb") as f:

            st.download_button(
                label="Download PDF",
                data=f,
                file_name=filename
            )

    else:
        st.error("PDF not generated.")

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
