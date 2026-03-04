import streamlit as st
import subprocess
import os
from datetime import datetime
import streamlit.components.v1 as components

# check query parameters
params = st.query_params

if "mode" in params and "date" in params:

    mode = params["mode"]
    date = params["date"]

    subprocess.run(["python", "daily_list_pc_version.py", date, mode])

    date_obj = datetime.strptime(date, "%Y-%m-%d")
    file_date = date_obj.strftime("%d-%b-%y")

    filename = f"{file_date} list.pdf"

    if os.path.exists(filename):
        with open(filename, "rb") as f:
            st.download_button("Download PDF", f, file_name=filename)

# load your existing UI
with open("kitchen_operations.html", "r", encoding="utf-8") as f:
    components.html(f.read(), height=900, scrolling=True)
