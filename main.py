import streamlit as st

st.set_page_config(page_title="MPW monthly report and Monthly Diary", layout="wide")

# ── Safe imports with error catching ──────────────────────────
try:
    from home_visit import house_visit_lookup
except Exception as e:
    st.error(f"❌ home_visit import failed: {e}")
    st.stop()

try:
    from monthly_final_report import mothly_final_report
except Exception as e:
    st.error(f"❌ monthly_final_report import failed: {e}")
    st.stop()

try:
    from mothly_diary import monthly_diary
except Exception as e:
    st.error(f"❌ mothly_diary import failed: {e}")
    st.stop()

# ── Navigation ────────────────────────────────────────────────
tabs = ["Monthly Diary", "Monthly Report", "Home Visit"]
tab  = st.sidebar.radio("Navigate", tabs)

if tab == "Monthly Diary":
    monthly_diary()
elif tab == "Home Visit":
    house_visit_lookup()
elif tab == "Monthly Report":
    mothly_final_report()