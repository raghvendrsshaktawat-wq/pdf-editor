Fenesta WCS Survey Editor
A Streamlit web app for Fenesta surveyors to upload WCS PDF reports, edit survey dimensions, and generate overlaid PDFs.

Features
Upload multiple Fenesta WCS PDFs at once

Auto-extracts all sales line items (Sales Line, Qty, Description, System, Order W/H, Reference, Location, Glazing)

Excel-like editable grid for Survey W, Survey H, Room, Remarks

Tolerance color coding: 🟢 ≤75mm · 🟡 76–200mm · 🔴 >200mm

Overlays survey data back onto the original WCS PDF

Combined Excel export (all WCS in one file, one sheet per WCS)

Fenesta brand UI (orange + navy)

How to run locally
bash
pip install -r requirements.txt
streamlit run app.py
Deploy to Streamlit Cloud
Push this repo to GitHub

Go to share.streamlit.io

Click New app → select your repo → set app.py as the main file

Click Deploy

Files
app.py — main Streamlit app

requirements.txt — Python dependencies

README.md — this file
