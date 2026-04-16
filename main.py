import streamlit as st

from home_visit import house_visit_lookup
from monthly_final_report import mothly_final_report
from monthly_repoert import monthly_repo
from mothly_diary import monthly_diary

st.set_page_config(page_title="MPW monthly report and Monthly Diary", layout="wide")



# Tabs for navigation
tabs = ["Monthly Diary", "Monthly Report", "Home Visit", "Other Reports"]

tab = st.sidebar.radio("Navigate", tabs)


if tab == "Monthly Diary":
    monthly_diary()
elif tab == "Home Visit":
    house_visit_lookup()
elif tab == "Monthly Report":
    mothly_final_report()
elif tab == "Other Reports":
    monthly_repo()

