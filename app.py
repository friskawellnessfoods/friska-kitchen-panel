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

    progress = st.progress(0)
    status = st.empty()

    # Convert date to month/day for script input
    month = date.strftime("%b")   # Mar
    day = date.day                # 5

    process = subprocess.Popen(
        [sys.executable, "daily_list_pc_version.py"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1
    )

    # Send inputs exactly like keyboard
    process.stdin.write(f"{month}\n")
    process.stdin.write(f"{day}\n")
    process.stdin.flush()

    percent = 0

    for line in process.stdout:

        if "%" in line:
            try:
                percent = int(line.split("%")[0].split()[-1])
                progress.progress(percent)
                status.text(f"Preparing kitchen list... {percent}%")
            except:
                pass

    process.wait()

    progress.progress(100)
    status.text("Finalizing PDF...")

    file_date = date.strftime("%d-%b-%y")
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
