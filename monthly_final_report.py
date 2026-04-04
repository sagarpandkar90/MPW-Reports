import streamlit as st
import json
import base64
import streamlit.components.v1 as components
from pathlib import Path

# ── Page config – call ONCE at top level ─────────────────────────────────────
st.set_page_config(page_title="मासिक आहवाल", layout="wide")


# ── Safe int conversion (no crash on bad input) ───────────────────────────────
def safe_int(val):
    try:
        return int(str(val).strip() or 0)
    except (ValueError, TypeError):
        return 0


def mothly_final_report():
    st.title("📊 मासिक आहवाल")

    # ── Initialize session state ONCE ────────────────────────────────────────
    if 'sheet_data' not in st.session_state:
        st.session_state.sheet_data = {}
    if 'meta_month_year' not in st.session_state:
        st.session_state.meta_month_year = 'मार्च २०२६'
    if 'meta_phc_name' not in st.session_state:
        st.session_state.meta_phc_name = 'शेळगाव'
    if 'meta_taluka' not in st.session_state:
        st.session_state.meta_taluka = 'इंदापूर'
    if 'meta_district' not in st.session_state:
        st.session_state.meta_district = 'पुणे'
    if 'meta_sub_center' not in st.session_state:
        st.session_state.meta_sub_center = ''
    if 'meta_population' not in st.session_state:
        st.session_state.meta_population = ''

    # ── Helper: get saved value ───────────────────────────────────────────────
    def sv(sheet, *keys, default=""):
        try:
            d = st.session_state.sheet_data.get(sheet, {})
            for k in keys:
                d = d[k]
            return d if d is not None else default
        except (KeyError, TypeError, IndexError):
            return default

    def sv_row(sheet, table_key, row_idx, col_key, default=""):
        try:
            rows = st.session_state.sheet_data.get(sheet, {}).get(table_key, [])
            return rows[row_idx].get(col_key, default) if row_idx < len(rows) else default
        except (KeyError, TypeError, IndexError):
            return default

    # ── Sidebar metadata ──────────────────────────────────────────────────────
    st.sidebar.header("मुख्य माहिती")
    month_year = st.sidebar.text_input("महिना/वर्ष",    key="meta_month_year")
    phc_name   = st.sidebar.text_input("प्रा. आ. केंद्र", key="meta_phc_name")
    taluka     = st.sidebar.text_input("तालुका",         key="meta_taluka")
    district   = st.sidebar.text_input("जिल्हा",         key="meta_district")
    sub_center = st.sidebar.text_input("उपकेंद्र",       key="meta_sub_center")
    population = st.sidebar.text_input("लोकसंख्या",      key="meta_population")

    # ── Sidebar: JSON export/import ───────────────────────────────────────────
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 💾 डेटा सुरक्षा")

    if st.sidebar.button("📥 JSON डाउनलोड करा", key="sidebar_export"):
        export_data = {
            'metadata': {
                'month_year': month_year, 'phc_name': phc_name,
                'taluka': taluka, 'district': district,
                'sub_center': sub_center, 'population': population,
            },
            'sheet_data': st.session_state.sheet_data
        }
        json_str = json.dumps(export_data, ensure_ascii=False, indent=2)
        b64 = base64.b64encode(json_str.encode('utf-8')).decode()
        st.sidebar.markdown(
            f'<a href="data:application/json;base64,{b64}" '
            f'download="masik_ahwal_backup.json" '
            f'style="background:#4CAF50;color:white;padding:8px 16px;border-radius:4px;'
            f'text-decoration:none;display:block;text-align:center;margin-top:8px;">'
            f'⬇️ डाउनलोड करा</a>',
            unsafe_allow_html=True
        )

    uploaded_backup = st.sidebar.file_uploader("📤 JSON बॅकअप लोड करा", type=['json'], key="backup_uploader")
    if uploaded_backup is not None:
        try:
            backup_data = json.loads(uploaded_backup.read().decode('utf-8'))
            if 'sheet_data' in backup_data:
                st.session_state.sheet_data = backup_data['sheet_data']
                if 'metadata' in backup_data:
                    meta = backup_data['metadata']
                    for k, v in [('meta_month_year', 'month_year'), ('meta_phc_name', 'phc_name'),
                                  ('meta_taluka', 'taluka'), ('meta_district', 'district'),
                                  ('meta_sub_center', 'sub_center'), ('meta_population', 'population')]:
                        st.session_state[k] = meta.get(v, '')
                st.success("✅ डेटा यशस्वीरित्या लोड झाला!")
                st.rerun()
        except Exception as e:
            st.sidebar.error(f"❌ लोड करताना त्रुटी: {e}")

    # ── Tabs ──────────────────────────────────────────────────────────────────
    tabs = st.tabs([
        "1️⃣ रक्त नमुना", "2️⃣ थुंकी संकलन", "3️⃣ कुष्ठरुग्ण",
        "4️⃣ क्षय रुग्ण",  "5️⃣ कंटेनर",      "6️⃣ डासउत्पत्ती",
        "7️⃣ मोतीबिंदू",   "8️⃣ प्रयोगशाळा",  "9️⃣ PDF तयार करा"
    ])

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 1 — रक्त नमुना
    # ════════════════════════════════════════════════════════════════════════
    with tabs[0]:
        st.subheader("रक्त नमुना मासिक अहवाल")

        st.markdown("### उपकेंद्र माहिती")
        c1, c2, c3, c4 = st.columns(4)
        subcenter_sr   = c1.text_input("अ. क्र.",                    value=sv('sheet1','subcenter_sr', default="1"), key="s1_basic_sr")
        subcenter_name = c2.text_input("उपकेंद्राचे नाव",             value=sv('sheet1','subcenter_name'),           key="s1_basic_name")
        subcenter_pop  = c3.text_input("लोकसंख्या",                   value=sv('sheet1','subcenter_pop'),            key="s1_basic_pop")
        annual_target  = c4.text_input("रक्त नमुना वार्षिक उद्दिष्ट", value=sv('sheet1','annual_target'),           key="s1_annual_target")

        st.markdown("---")
        st.markdown("### कर्मचारी माहिती")

        saved_staff_count = len(st.session_state.sheet_data.get('sheet1', {}).get('staff_data', [])) or 3
        num_staff = st.number_input("कर्मचारी संख्या", min_value=1, max_value=20,
                                    value=saved_staff_count, key="rows1")

        designation_options = ["आरोग्य सेवक", "आरोग्य सेविका", "आरोग्य सेविका NHM"]
        staff_data = []

        for i in range(num_staff):
            st.markdown(f"**कर्मचारी {i+1}**")
            cols = st.columns([2, 2, 1, 1, 1, 1])
            staff_name = cols[0].text_input("नाव", value=sv_row('sheet1','staff_data',i,'नाव'), key=f"s1_name_{i}")
            saved_desg = sv_row('sheet1','staff_data',i,'पदनाम')
            desg_idx   = designation_options.index(saved_desg) if saved_desg in designation_options else 0
            staff_designation = cols[1].selectbox("पदनाम", options=designation_options,
                                                   index=desg_idx, key=f"s1_desg_{i}")
            row = {
                'नाव':            staff_name,
                'पदनाम':          staff_designation,
                'पहिला_पंधरावडा': cols[2].text_input("मासिक पहिला",  value=sv_row('sheet1','staff_data',i,'पहिला_पंधरावडा','0'), key=f"s1_f1_{i}"),
                'दुसरा_पंधरावडा': cols[3].text_input("मासिक दुसरा",  value=sv_row('sheet1','staff_data',i,'दुसरा_पंधरावडा','0'), key=f"s1_f2_{i}"),
                'प्रगती_पहिला':   cols[4].text_input("प्रगती पहिला", value=sv_row('sheet1','staff_data',i,'प्रगती_पहिला','0'),   key=f"s1_p1_{i}"),
                'प्रगती_दुसरा':   cols[5].text_input("प्रगती दुसरा", value=sv_row('sheet1','staff_data',i,'प्रगती_दुसरा','0'),   key=f"s1_p2_{i}"),
            }
            staff_data.append(row)

        st.markdown("---")
        st.markdown("### एकूण आशा कार्यकर्ती")
        total_asha_count = st.text_input("संख्या", value=sv('sheet1','total_asha_count', default="0"), key="s1_total_asha")
        cols_asha = st.columns([2, 2, 1, 1, 1, 1])
        asha_f1 = cols_asha[2].text_input("मासिक पहिला",  value=sv('sheet1','asha_data','f1', default="0"), key="s1_asha_f1")
        asha_f2 = cols_asha[3].text_input("मासिक दुसरा",  value=sv('sheet1','asha_data','f2', default="0"), key="s1_asha_f2")
        asha_p1 = cols_asha[4].text_input("प्रगती पहिला", value=sv('sheet1','asha_data','p1', default="0"), key="s1_asha_p1")
        asha_p2 = cols_asha[5].text_input("प्रगती दुसरा", value=sv('sheet1','asha_data','p2', default="0"), key="s1_asha_p2")

        total_f1       = sum(safe_int(s['पहिला_पंधरावडा']) for s in staff_data) + safe_int(asha_f1)
        total_f2       = sum(safe_int(s['दुसरा_पंधरावडा']) for s in staff_data) + safe_int(asha_f2)
        total_monthly  = total_f1 + total_f2
        total_p1       = sum(safe_int(s['प्रगती_पहिला'])   for s in staff_data) + safe_int(asha_p1)
        total_p2       = sum(safe_int(s['प्रगती_दुसरा'])   for s in staff_data) + safe_int(asha_p2)
        total_progress = total_p1 + total_p2

        st.info(f"📊 एकूण — मासिक: पहिला={total_f1}, दुसरा={total_f2}, एकूण={total_monthly} | "
                f"प्रगती: पहिला={total_p1}, दुसरा={total_p2}, एकूण={total_progress}")

        # Save only when tab is active (not on every keystroke globally)
        st.session_state.sheet_data['sheet1'] = {
            'title':          'राष्ट्रीय कीटकजन्य रोग नियंत्रण कार्यक्रम, जिल्हा पुणे',
            'subtitle':       'रक्त नमुना मासिक अहवाल',
            'month_year':     month_year,
            'subcenter_sr':   subcenter_sr,
            'subcenter_name': subcenter_name,
            'subcenter_pop':  subcenter_pop,
            'annual_target':  annual_target,
            'staff_data':     staff_data,
            'total_asha_count': total_asha_count,
            'asha_data':      {'f1': asha_f1, 'f2': asha_f2, 'p1': asha_p1, 'p2': asha_p2},
            'totals': {
                'f1': str(total_f1), 'f2': str(total_f2), 'monthly': str(total_monthly),
                'p1': str(total_p1), 'p2': str(total_p2), 'progress': str(total_progress)
            }
        }

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 2 — थुंकी संकलन
    # ════════════════════════════════════════════════════════════════════════
    with tabs[1]:
        st.subheader("थुंकी संकलन अहवाल")

        st.markdown("### तक्ता १: गावनिहाय थुंकी संकलन")
        saved_2a     = len(st.session_state.sheet_data.get('sheet2', {}).get('table1', [])) or 4
        num_rows_2a  = st.number_input("नोंदी (तक्ता १)", min_value=1, max_value=20, value=saved_2a, key="rows2a")

        keys2a   = ['अ. क्र.','गावाचे नाव','लोकसंख्या','कर्मचारी',
                    'मासिक_पुरुष','मासिक_स्त्री','मासिक_एकूण',
                    'वार्षिक_पुरुष','वार्षिक_स्त्री','वार्षिक_एकूण']
        labels2a = ["अ.क्र.","गाव","लोकसंख्या","कर्मचारी",
                    "मासिक पुरुष","मासिक स्त्री","मासिक एकूण",
                    "वार्षिक पुरुष","वार्षिक स्त्री","वार्षिक एकूण"]
        wkeys2a  = ["s2a_sr","s2a_village","s2a_pop","s2a_staff",
                    "s2a_m_m","s2a_m_f","s2a_m_t","s2a_y_m","s2a_y_f","s2a_y_t"]

        sheet2_table1 = []
        for i in range(num_rows_2a):
            cols = st.columns([1,2,2,2,1,1,1,1,1,1])
            row  = {}
            defaults = [str(i+1)] + [""]*9
            for j,(k,wk,lbl,dft) in enumerate(zip(keys2a, wkeys2a, labels2a, defaults)):
                row[k] = cols[j].text_input(lbl, value=sv_row('sheet2','table1',i,k,dft), key=f"{wk}_{i}")
            sheet2_table1.append(row)

        st.markdown("---")
        st.markdown("### तक्ता २: संशयीत रुग्ण माहिती")
        saved_2b    = len(st.session_state.sheet_data.get('sheet2', {}).get('table2', [])) or 4
        num_rows_2b = st.number_input("नोंदी (तक्ता २)", min_value=1, max_value=20, value=saved_2b, key="rows2b")

        keys2b   = ['अ. क्र.','गावाचे नाव','संशयीत_रुग्ण','वय','लिंग',
                    'नमुना_दिनांक','तपासणी_दिनांक','लॅब_क्रमांक','कर्मचारी']
        labels2b = ["अ.क्र.","गाव","रुग्ण नाव","वय","लिंग",
                    "नमुना दिनांक","तपासणी दिनांक","लॅब क्रमांक","कर्मचारी"]
        wkeys2b  = ["s2b_sr","s2b_village","s2b_name","s2b_age","s2b_gender",
                    "s2b_sample","s2b_test","s2b_lab","s2b_staff"]

        sheet2_table2 = []
        for i in range(num_rows_2b):
            cols = st.columns([1,2,2,1,1,2,2,2,2])
            row  = {}
            defaults = [str(i+1)] + [""]*8
            for j,(k,wk,lbl,dft) in enumerate(zip(keys2b, wkeys2b, labels2b, defaults)):
                row[k] = cols[j].text_input(lbl, value=sv_row('sheet2','table2',i,k,dft), key=f"{wk}_{i}")
            sheet2_table2.append(row)

        st.session_state.sheet_data['sheet2'] = {
            'title1': 'थुंकी संकलन गावनिहाय अहवाल',
            'title2': 'संशयीत क्षयरुग्ण तपासणी अहवाल',
            'table1': sheet2_table1, 'table2': sheet2_table2
        }

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 3 — कुष्ठरुग्ण
    # ════════════════════════════════════════════════════════════════════════
    with tabs[2]:
        st.subheader("कुष्ठरुग्ण मासिक अहवाल")

        st.markdown("### तक्ता १: गावनिहाय कुष्ठरुग्ण")
        saved_3a    = len(st.session_state.sheet_data.get('sheet3', {}).get('table1', [])) or 4
        num_rows_3a = st.number_input("नोंदी (तक्ता १)", min_value=1, max_value=20, value=saved_3a, key="rows3a")

        keys3a   = ['अ. क्र.','गावाचे नाव','लोकसंख्या',
                    'संबंधित_मुले','संबंधित_प्रौढ','संबंधित_एकूण',
                    'MB_मुले','MB_प्रौढ','MB_एकूण',
                    'PB_मुले','PB_प्रौढ','PB_एकूण',
                    'औषधोपचार_मुले','औषधोपचार_प्रौढ','औषधोपचार_एकूण']
        labels3a = ["अ.क्र.","गाव","लोकसंख्या",
                    "सं.मुले","सं.प्रौढ","सं.एकूण",
                    "MB मुले","MB प्रौढ","MB एकूण",
                    "PB मुले","PB प्रौढ","PB एकूण",
                    "औ.मुले","औ.प्रौढ","औ.एकूण"]
        wkeys3a  = [f"s3a_{x}" for x in ["sr","village","pop",
                    "rel_c","rel_a","rel_t","mb_c","mb_a","mb_t",
                    "pb_c","pb_a","pb_t","tr_c","tr_a","tr_t"]]

        sheet3_table1 = []
        for i in range(num_rows_3a):
            cols = st.columns([1,2,2,1,1,1,1,1,1,1,1,1,1,1,1])
            row  = {}
            defaults = [str(i+1)] + [""]*14
            for j,(k,wk,lbl,dft) in enumerate(zip(keys3a, wkeys3a, labels3a, defaults)):
                row[k] = cols[j].text_input(lbl, value=sv_row('sheet3','table1',i,k,dft), key=f"{wk}_{i}")
            sheet3_table1.append(row)

        st.markdown("---")
        st.markdown("### तक्ता २: कुष्ठरुग्ण तपशील")
        saved_3b    = len(st.session_state.sheet_data.get('sheet3', {}).get('table2', [])) or 3
        num_rows_3b = st.number_input("नोंदी (तक्ता २)", min_value=1, max_value=20, value=saved_3b, key="rows3b")

        sheet3_table2 = []
        for i in range(num_rows_3b):
            cols = st.columns([1,2,2,1,1,3,2])
            row  = {
                'अ. क्र.':       cols[0].text_input("अ.क्र.",   value=sv_row('sheet3','table2',i,'अ. क्र.',str(i+1)), key=f"s3b_sr_{i}"),
                'गावाचे नाव':    cols[1].text_input("गाव",      value=sv_row('sheet3','table2',i,'गावाचे नाव'),       key=f"s3b_village_{i}"),
                'रुग्णाचे_नाव':  cols[2].text_input("रुग्ण नाव",value=sv_row('sheet3','table2',i,'रुग्णाचे_नाव'),    key=f"s3b_name_{i}"),
                'वय':            cols[3].text_input("वय",       value=sv_row('sheet3','table2',i,'वय'),               key=f"s3b_age_{i}"),
                'लिंग':          cols[4].text_input("लिंग",     value=sv_row('sheet3','table2',i,'लिंग'),             key=f"s3b_gender_{i}"),
                'लक्षणे':        cols[5].text_input("लक्षणे",   value=sv_row('sheet3','table2',i,'लक्षणे'),           key=f"s3b_symptoms_{i}"),
                'कर्मचारी':      cols[6].text_input("कर्मचारी", value=sv_row('sheet3','table2',i,'कर्मचारी'),         key=f"s3b_staff_{i}"),
            }
            sheet3_table2.append(row)

        st.session_state.sheet_data['sheet3'] = {
            'title1': 'कुष्ठरुग्ण गावनिहाय मासिक अहवाल',
            'title2': 'कुष्ठरुग्ण तपशीलवार माहिती',
            'table1': sheet3_table1, 'table2': sheet3_table2
        }

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 4 — क्षय रुग्ण
    # ════════════════════════════════════════════════════════════════════════
    with tabs[3]:
        st.subheader("क्षय रुग्ण अहवाल")

        st.markdown("### तक्ता १: गावनिहाय क्षय रुग्ण")
        saved_4a    = len(st.session_state.sheet_data.get('sheet4', {}).get('table1', [])) or 3
        num_rows_4a = st.number_input("नोंदी (तक्ता १)", min_value=1, max_value=20, value=saved_4a, key="rows4a")

        keys4a   = ['अ. क्र.','गावाचे नाव','लोकसंख्या','कर्मचारी',
                    'मासिक_पुरुष','मासिक_स्त्री','मासिक_एकूण',
                    'वार्षिक_पुरुष','वार्षिक_स्त्री','वार्षिक_एकूण']
        labels4a = ["अ.क्र.","गाव","लोकसंख्या","कर्मचारी",
                    "मासिक पुरुष","मासिक स्त्री","मासिक एकूण",
                    "वार्षिक पुरुष","वार्षिक स्त्री","वार्षिक एकूण"]
        wkeys4a  = ["s4a_sr","s4a_village","s4a_pop","s4a_staff",
                    "s4a_m_m","s4a_m_f","s4a_m_t","s4a_y_m","s4a_y_f","s4a_y_t"]

        sheet4_table1 = []
        for i in range(num_rows_4a):
            cols = st.columns([1,2,2,2,1,1,1,1,1,1])
            row  = {}
            defaults = [str(i+1)] + [""]*9
            for j,(k,wk,lbl,dft) in enumerate(zip(keys4a, wkeys4a, labels4a, defaults)):
                row[k] = cols[j].text_input(lbl, value=sv_row('sheet4','table1',i,k,dft), key=f"{wk}_{i}")
            sheet4_table1.append(row)

        st.markdown("---")
        st.markdown("### तक्ता २: क्षयरुग्ण तपशील")
        saved_4b    = len(st.session_state.sheet_data.get('sheet4', {}).get('table2', [])) or 3
        num_rows_4b = st.number_input("नोंदी (तक्ता २)", min_value=1, max_value=20, value=saved_4b, key="rows4b")

        sheet4_table2 = []
        for i in range(num_rows_4b):
            cols = st.columns([1,2,2,1,1,2,2,2,2])
            row  = {
                'अ. क्र.':          cols[0].text_input("अ.क्र.",   value=sv_row('sheet4','table2',i,'अ. क्र.',str(i+1)), key=f"s4b_sr_{i}"),
                'गावाचे नाव':       cols[1].text_input("गाव",      value=sv_row('sheet4','table2',i,'गावाचे नाव'),       key=f"s4b_village_{i}"),
                'क्षयरुग्णाचे_नाव': cols[2].text_input("रुग्ण नाव",value=sv_row('sheet4','table2',i,'क्षयरुग्णाचे_नाव'),key=f"s4b_name_{i}"),
                'वय':               cols[3].text_input("वय",       value=sv_row('sheet4','table2',i,'वय'),               key=f"s4b_age_{i}"),
                'लिंग':             cols[4].text_input("लिंग",     value=sv_row('sheet4','table2',i,'लिंग'),             key=f"s4b_gender_{i}"),
                'कॅटेगरी':          cols[5].text_input("कॅटेगरी",  value=sv_row('sheet4','table2',i,'कॅटेगरी'),          key=f"s4b_cat_{i}"),
                'औषधोपचार_दिनांक':  cols[6].text_input("दिनांक",  value=sv_row('sheet4','table2',i,'औषधोपचार_दिनांक'), key=f"s4b_date_{i}"),
                'TB_नंबर':          cols[7].text_input("TB नं.",   value=sv_row('sheet4','table2',i,'TB_नंबर'),          key=f"s4b_tb_{i}"),
                'कर्मचारी':         cols[8].text_input("कर्मचारी", value=sv_row('sheet4','table2',i,'कर्मचारी'),         key=f"s4b_staff_{i}"),
            }
            sheet4_table2.append(row)

        st.session_state.sheet_data['sheet4'] = {
            'title1': 'क्षयरुग्ण गावनिहाय मासिक अहवाल',
            'title2': 'उपचार घेणारे क्षयरुग्ण तपशील',
            'table1': sheet4_table1, 'table2': sheet4_table2
        }

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 5 — कंटेनर सर्वेक्षण
    # ════════════════════════════════════════════════════════════════════════
    with tabs[4]:
        st.subheader("कंटेनर सर्वेक्षण अहवाल")

        saved_5    = len(st.session_state.sheet_data.get('sheet5', {}).get('data', [])) or 3
        num_rows_5 = st.number_input("नोंदी संख्या", min_value=1, max_value=20, value=saved_5, key="rows5")

        keys5   = ['अ. क्र.','गावाचे नाव','लोकसंख्या','एकूण_घरे','तपासलेले_घरे','दूषित_घरे',
                   'तपासलेली_भांडी','दूषित_भांडी','House_Index','Container_Index','Breteau_Index',
                   'रिकामी_भांडी','अँबेट_भांडी']
        labels5 = ["अ.क्र.","गाव","लोकसंख्या","एकूण घरे","तपासले","दूषित",
                   "भांडी","दूषित भांडी","HI","CI","BI","रिकामी","अँबेट"]
        wkeys5  = ["s5_sr","s5_village","s5_pop","s5_house","s5_check","s5_cont",
                   "s5_cont_c","s5_cont_cc","s5_hi","s5_ci","s5_bi","s5_empty","s5_abate"]

        sheet5_data = []
        for i in range(num_rows_5):
            cols = st.columns([1,2,2,2,2,2,2,2,2,2,2,2,2])
            row  = {}
            defaults = [str(i+1)] + [""]*12
            for j,(k,wk,lbl,dft) in enumerate(zip(keys5, wkeys5, labels5, defaults)):
                row[k] = cols[j].text_input(lbl, value=sv_row('sheet5','data',i,k,dft), key=f"{wk}_{i}")
            sheet5_data.append(row)

        st.session_state.sheet_data['sheet5'] = {
            'title': 'कंटेनर सर्वेक्षण मासिक अहवाल',
            'data': sheet5_data
        }

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 6 — डासउत्पत्ती
    # ════════════════════════════════════════════════════════════════════════
    with tabs[5]:
        st.subheader("डासउत्पत्ती स्थाने")

        def _das_table(sheet_key, table_key, title, num_key, saved_count, row_prefix):
            st.markdown(f"### {title}")
            num = st.number_input(f"नोंदी ({title})", min_value=1, max_value=20,
                                   value=saved_count, key=num_key)
            rows = []
            for i in range(num):
                cols = st.columns([1,2,2,1,1,3])
                row  = {
                    'अ. क्र.':          cols[0].text_input("अ.क्र.",    value=sv_row(sheet_key,table_key,i,'अ. क्र.',str(i+1)), key=f"{row_prefix}_sr_{i}"),
                    'उपकेंद्राचे नाव': cols[1].text_input("उपकेंद्र", value=sv_row(sheet_key,table_key,i,'उपकेंद्राचे नाव'),   key=f"{row_prefix}_sub_{i}"),
                    'गावाचे नाव':       cols[2].text_input("गाव",       value=sv_row(sheet_key,table_key,i,'गावाचे नाव'),       key=f"{row_prefix}_village_{i}"),
                    'कायम':             cols[3].text_input("कायम",      value=sv_row(sheet_key,table_key,i,'कायम'),             key=f"{row_prefix}_perm_{i}"),
                    'हंगामी':           cols[4].text_input("हंगामी",    value=sv_row(sheet_key,table_key,i,'हंगामी'),           key=f"{row_prefix}_seas_{i}"),
                    'ठिकाण':            cols[5].text_input("ठिकाण",    value=sv_row(sheet_key,table_key,i,'ठिकाण'),            key=f"{row_prefix}_loc_{i}"),
                }
                rows.append(row)
            return num, rows

        saved_6a = len(st.session_state.sheet_data.get('sheet6', {}).get('table1', [])) or 4
        saved_6b = len(st.session_state.sheet_data.get('sheet6', {}).get('table2', [])) or 4

        _, sheet6_table1 = _das_table('sheet6','table1','तक्ता १: डासउत्पत्ती स्थाने','rows6a',saved_6a,'s6a')
        st.markdown("---")
        _, sheet6_table2 = _das_table('sheet6','table2','तक्ता २: गप्पी मासे पैदास केंद्र','rows6b',saved_6b,'s6b')

        st.session_state.sheet_data['sheet6'] = {
            'title1': 'डासउत्पत्ती स्थानांची गावनिहाय यादी',
            'title2': 'गप्पी मासे पैदास केंद्राची यादी',
            'table1': sheet6_table1, 'table2': sheet6_table2
        }

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 7 — मोतीबिंदू
    # ════════════════════════════════════════════════════════════════════════
    with tabs[6]:
        st.subheader("मोतीबिंदू मासिक अहवाल")

        st.markdown("### तक्ता १: गावनिहाय मोतीबिंदू")
        saved_7a    = len(st.session_state.sheet_data.get('sheet7', {}).get('table1', [])) or 4
        num_rows_7a = st.number_input("नोंदी (तक्ता १)", min_value=1, max_value=20, value=saved_7a, key="rows7a")

        sheet7_table1 = []
        for i in range(num_rows_7a):
            cols = st.columns([1,2,2,1,1,1,1,1,1])
            row  = {
                'अ. क्र.':        cols[0].text_input("अ.क्र.",   value=sv_row('sheet7','table1',i,'अ. क्र.',str(i+1)), key=f"s7a_sr_{i}"),
                'गावाचे नाव':     cols[1].text_input("गाव",      value=sv_row('sheet7','table1',i,'गावाचे नाव'),       key=f"s7a_village_{i}"),
                'लोकसंख्या':      cols[2].text_input("लोकसंख्या",value=sv_row('sheet7','table1',i,'लोकसंख्या'),        key=f"s7a_pop_{i}"),
                'संशयीत_मुले':   cols[3].text_input("सं.मुले",  value=sv_row('sheet7','table1',i,'संशयीत_मुले'),     key=f"s7a_susp_c_{i}"),
                'संशयीत_प्रौढ':  cols[4].text_input("सं.प्रौढ", value=sv_row('sheet7','table1',i,'संशयीत_प्रौढ'),   key=f"s7a_susp_a_{i}"),
                'संशयीत_एकूण':   cols[5].text_input("सं.एकूण",  value=sv_row('sheet7','table1',i,'संशयीत_एकूण'),    key=f"s7a_susp_t_{i}"),
                'नवीन_मुले':      cols[6].text_input("नवीन मुले",value=sv_row('sheet7','table1',i,'नवीन_मुले'),       key=f"s7a_new_c_{i}"),
                'नवीन_प्रौढ':     cols[7].text_input("नवीन प्रौढ",value=sv_row('sheet7','table1',i,'नवीन_प्रौढ'),    key=f"s7a_new_a_{i}"),
                'नवीन_एकूण':      cols[8].text_input("नवीन एकूण",value=sv_row('sheet7','table1',i,'नवीन_एकूण'),      key=f"s7a_new_t_{i}"),
            }
            sheet7_table1.append(row)

        st.markdown("---")
        st.markdown("### तक्ता २: मोतीबिंदू रुग्ण तपशील")
        saved_7b    = len(st.session_state.sheet_data.get('sheet7', {}).get('table2', [])) or 4
        num_rows_7b = st.number_input("नोंदी (तक्ता २)", min_value=1, max_value=20, value=saved_7b, key="rows7b")

        sheet7_table2 = []
        for i in range(num_rows_7b):
            cols = st.columns([1,2,2,1,1,3,2])
            row  = {
                'अ. क्र.':       cols[0].text_input("अ.क्र.",    value=sv_row('sheet7','table2',i,'अ. क्र.',str(i+1)), key=f"s7b_sr_{i}"),
                'गावाचे नाव':    cols[1].text_input("गाव",       value=sv_row('sheet7','table2',i,'गावाचे नाव'),       key=f"s7b_village_{i}"),
                'रुग्णाचे_नाव':  cols[2].text_input("रुग्ण नाव", value=sv_row('sheet7','table2',i,'रुग्णाचे_नाव'),    key=f"s7b_name_{i}"),
                'वय':            cols[3].text_input("वय",        value=sv_row('sheet7','table2',i,'वय'),               key=f"s7b_age_{i}"),
                'लिंग':          cols[4].text_input("लिंग",      value=sv_row('sheet7','table2',i,'लिंग'),             key=f"s7b_gender_{i}"),
                'लक्षणे':        cols[5].text_input("लक्षणे",    value=sv_row('sheet7','table2',i,'लक्षणे'),           key=f"s7b_symptoms_{i}"),
                'कर्मचारी':      cols[6].text_input("कर्मचारी",  value=sv_row('sheet7','table2',i,'कर्मचारी'),         key=f"s7b_staff_{i}"),
            }
            sheet7_table2.append(row)

        st.session_state.sheet_data['sheet7'] = {
            'title1': 'मोतीबिंदू गावनिहाय मासिक अहवाल',
            'title2': 'मोतीबिंदू रुग्ण तपशीलवार माहिती',
            'table1': sheet7_table1, 'table2': sheet7_table2
        }

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 8 — प्रयोगशाळा
    # ════════════════════════════════════════════════════════════════════════
    with tabs[7]:
        st.subheader("प्रयोगशाळा अहवाल")

        st.markdown("### तक्ता १: प्रयोगशाळा तपासणी")
        saved_8a    = len(st.session_state.sheet_data.get('sheet8', {}).get('table1', [])) or 3
        num_rows_8a = st.number_input("नोंदी (तक्ता १)", min_value=1, max_value=20, value=saved_8a, key="rows8a")

        sheet8_table1 = []
        for i in range(num_rows_8a):
            with st.expander(f"नोंद {i+1}", expanded=(i==0)):
                c1, c2, c3 = st.columns(3)
                sr             = c1.text_input("अ.क्र.",           value=sv_row('sheet8','table1',i,'अ. क्र.',str(i+1)), key=f"s8a_sr_{i}")
                subcenter_val  = c2.text_input("उपकेंद्र",          value=sv_row('sheet8','table1',i,'उपकेंद्र'),         key=f"s8a_sub_{i}")
                total_surveys  = c3.text_input("एकूण सार्व. उद्भव", value=sv_row('sheet8','table1',i,'एकूण_सार्व'),       key=f"s8a_total_{i}")

                st.markdown("**जैविक पाणी**")
                c1,c2,c3,c4 = st.columns(4)
                bio_m_taken = c1.text_input("महिना घेतले",   value=sv_row('sheet8','table1',i,'जै_मही_घे'),    key=f"s8a_bio_mt_{i}")
                bio_m_cont  = c2.text_input("महिना दूषित",   value=sv_row('sheet8','table1',i,'जै_मही_दू'),    key=f"s8a_bio_mc_{i}")
                bio_y_taken = c3.text_input("प्रगती घेतले",  value=sv_row('sheet8','table1',i,'जै_प्रगती_घे'), key=f"s8a_bio_yt_{i}")
                bio_y_cont  = c4.text_input("प्रगती दूषित",  value=sv_row('sheet8','table1',i,'जै_प्रगती_दू'), key=f"s8a_bio_yc_{i}")

                st.markdown("**रासायनिक पाणी**")
                c1,c2,c3,c4 = st.columns(4)
                chem_m_taken = c1.text_input("महिना घेतले",  value=sv_row('sheet8','table1',i,'रा_मही_घे'),    key=f"s8a_chem_mt_{i}")
                chem_m_cont  = c2.text_input("महिना दूषित",  value=sv_row('sheet8','table1',i,'रा_मही_दू'),    key=f"s8a_chem_mc_{i}")
                chem_y_taken = c3.text_input("प्रगती घेतले", value=sv_row('sheet8','table1',i,'रा_प्रगती_घे'), key=f"s8a_chem_yt_{i}")
                chem_y_cont  = c4.text_input("प्रगती दूषित", value=sv_row('sheet8','table1',i,'रा_प्रगती_दू'), key=f"s8a_chem_yc_{i}")

                st.markdown("**TCL साठा**")
                c1,c2,c3,c4 = st.columns(4)
                tcl_start = c1.text_input("प्रारंभ", value=sv_row('sheet8','table1',i,'TCL_प्रारंभ'), key=f"s8a_tcl_s_{i}")
                tcl_buy   = c2.text_input("खरेदी",   value=sv_row('sheet8','table1',i,'TCL_खरेदी'),   key=f"s8a_tcl_b_{i}")
                tcl_use   = c3.text_input("खर्च",    value=sv_row('sheet8','table1',i,'TCL_खर्च'),    key=f"s8a_tcl_u_{i}")
                tcl_end   = c4.text_input("शेवट",    value=sv_row('sheet8','table1',i,'TCL_शेवट'),    key=f"s8a_tcl_e_{i}")

                st.markdown("**अतिरिक्त नमुने**")
                c1,c2,c3,c4,c5,c6 = st.columns(6)
                tcl_sample_m_t = c1.text_input("TCL मही घे",  value=sv_row('sheet8','table1',i,'TCL_नमुना_मही_घे'), key=f"s8a_tcls_mt_{i}")
                tcl_sample_m_c = c2.text_input("TCL मही दू",  value=sv_row('sheet8','table1',i,'TCL_नमुना_मही_दू'), key=f"s8a_tcls_mc_{i}")
                stool_m_t      = c3.text_input("शौच मही घे",  value=sv_row('sheet8','table1',i,'शौच_मही_घे'),       key=f"s8a_st_mt_{i}")
                stool_m_c      = c4.text_input("शौच मही दू",  value=sv_row('sheet8','table1',i,'शौच_मही_दू'),       key=f"s8a_st_mc_{i}")
                salt_m_t       = c5.text_input("मीठ मही घे",  value=sv_row('sheet8','table1',i,'मीठ_मही_घे'),       key=f"s8a_salt_mt_{i}")
                salt_m_c       = c6.text_input("मीठ मही दू",  value=sv_row('sheet8','table1',i,'मीठ_मही_दू'),       key=f"s8a_salt_mc_{i}")

            sheet8_table1.append({
                'अ. क्र.': sr, 'उपकेंद्र': subcenter_val, 'एकूण_सार्व': total_surveys,
                'जै_मही_घे': bio_m_taken, 'जै_मही_दू': bio_m_cont,
                'जै_प्रगती_घे': bio_y_taken, 'जै_प्रगती_दू': bio_y_cont,
                'रा_मही_घे': chem_m_taken, 'रा_मही_दू': chem_m_cont,
                'रा_प्रगती_घे': chem_y_taken, 'रा_प्रगती_दू': chem_y_cont,
                'TCL_प्रारंभ': tcl_start, 'TCL_खरेदी': tcl_buy,
                'TCL_खर्च': tcl_use, 'TCL_शेवट': tcl_end,
                'TCL_नमुना_मही_घे': tcl_sample_m_t, 'TCL_नमुना_मही_दू': tcl_sample_m_c,
                'शौच_मही_घे': stool_m_t, 'शौच_मही_दू': stool_m_c,
                'मीठ_मही_घे': salt_m_t, 'मीठ_मही_दू': salt_m_c,
            })

        st.markdown("---")
        st.markdown("### तक्ता २: गावनिहाय TCL साठा")
        saved_8b    = len(st.session_state.sheet_data.get('sheet8', {}).get('table2', [])) or 4
        num_rows_8b = st.number_input("नोंदी (तक्ता २)", min_value=1, max_value=20, value=saved_8b, key="rows8b")

        keys8b   = ['अ. क्र.','उपकेंद्र','ग्रामपंचायत','गावे','TCL_साठा','TCL_साठवण','पाणी_शुद्धी','TCL_नसलेले']
        labels8b = ["अ.क्र.","उपकेंद्र","ग्रामपंचायत","गावे","TCL साठा","TCL साठवण","पाणी शुद्धी","TCL नसलेले"]
        wkeys8b  = ["s8b_sr","s8b_sub","s8b_gp","s8b_villages","s8b_stock","s8b_storage","s8b_purif","s8b_no_tcl"]

        sheet8_table2 = []
        for i in range(num_rows_8b):
            cols = st.columns([1,2,2,2,2,2,2,2])
            row  = {}
            defaults = [str(i+1)] + [""]*7
            for j,(k,wk,lbl,dft) in enumerate(zip(keys8b, wkeys8b, labels8b, defaults)):
                row[k] = cols[j].text_input(lbl, value=sv_row('sheet8','table2',i,k,dft), key=f"{wk}_{i}")
            sheet8_table2.append(row)

        st.session_state.sheet_data['sheet8'] = {
            'title1': 'राज्य आयोग्य प्रयोगशाळा विविध नमुने तपासणी अहवाल',
            'title2': 'गावनिहाय TCL साठा अहवाल',
            'table1': sheet8_table1, 'table2': sheet8_table2
        }

    # ════════════════════════════════════════════════════════════════════════
    # TAB 9 — PDF Generate  (unchanged JS logic, cleaned embedding)
    # ════════════════════════════════════════════════════════════════════════
    with tabs[8]:
        st.subheader("📄 PDF तयार करा")
        st.info("सर्व डेटा भरल्यानंतर येथे PDF तयार करा")

        BASE_DIR  = Path(__file__).resolve().parent
        font_path = BASE_DIR / "fonts" / "NotoSerifDevanagari-VariableFont_wdth,wght.ttf"
        imgpath   = BASE_DIR / "fonts" / "img.png"

        if not font_path.exists():
            st.error(f"❌ Font file missing at: {font_path}")
            st.markdown("""
            **Download font from:**
            - [Google Fonts - Noto Serif Devanagari](https://fonts.google.com/noto/specimen/Noto+Serif+Devanagari)
            """)
            return

        font_b64       = base64.b64encode(font_path.read_bytes()).decode()
        img            = base64.b64encode(imgpath.read_bytes()).decode()
        all_sheets_json = json.dumps(st.session_state.sheet_data, ensure_ascii=False)

        components.html(
            f"""
            <html>
            <head>
              <meta charset="utf-8"/>
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.72/pdfmake.min.js"></script>
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.72/vfs_fonts.js"></script>
              <style>
                body {{ padding:20px; font-family:Arial,sans-serif; }}
                button {{
                  padding:14px 28px; margin-right:10px; border:none;
                  border-radius:6px; cursor:pointer; font-size:16px; font-weight:bold;
                  box-shadow:0 3px 6px rgba(0,0,0,.3);
                }}
                .preview-btn  {{ background:#2196F3; color:white; }}
                .download-btn {{ background:#4CAF50; color:white; }}
                #status {{ margin-top:15px; padding:12px; border-radius:4px; display:none; }}
              </style>
            </head>
            <body>
              <button class="preview-btn"  onclick="previewPDF()">👁️ Preview PDF</button>
              <button class="download-btn" onclick="downloadPDF()">⬇️ Download PDF</button>
              <div id="status"></div>
              <script>
                const logoImage = "data:image/png;base64,{img}";
                const allData   = {all_sheets_json};
                const metadata  = {{
                  monthYear: "{month_year}", phcName: "{phc_name}",
                  taluka: "{taluka}", district: "{district}",
                  subcenter: "{sub_center}", population: "{population}"
                }};
                pdfMake.vfs["MarathiFont.ttf"] = "{font_b64}";
                pdfMake.fonts = {{
                  MarathiFont: {{
                    normal:"MarathiFont.ttf", bold:"MarathiFont.ttf",
                    italics:"MarathiFont.ttf", bolditalics:"MarathiFont.ttf"
                  }}
                }};

                function showStatus(msg, color) {{
                  const s = document.getElementById('status');
                  s.innerHTML = msg; s.style.background = color; s.style.display = 'block';
                  setTimeout(() => s.style.display = 'none', 3000);
                }}
                function displayValue(val) {{
                  return (val === '0' || val === 0 || val === '' || val == null) ? '' : val;
                }}
                function addPageHeading(title, monthYear) {{
                  return [
                    {{ text:'राष्ट्रीय कीटकजन्य रोग नियंत्रण कार्यक्रम, जिल्हा पुणे', alignment:'center', bold:true, fontSize:13, margin:[0,20,0,4] }},
                    {{ text:'प्रा. आ. केंद्र '+metadata.phcName+' तालुका '+metadata.taluka+' जिल्हा '+metadata.district, fontSize:12, bold:true, alignment:'center', margin:[0,0,0,8] }},
                    {{ text:title+' '+(monthYear||metadata.monthYear), alignment:'center', bold:true, fontSize:12, margin:[0,0,0,15] }}
                  ];
                }}
                function addPageHeading2(title) {{
                  return [{{ text:title, alignment:'center', bold:true, fontSize:12, margin:[0,10,0,15] }}];
                }}
                function tableLayout() {{
                  return {{ hLineWidth:()=>0.5, vLineWidth:()=>0.5, paddingLeft:()=>4, paddingRight:()=>4, paddingTop:()=>4, paddingBottom:()=>4 }};
                }}
                function cell(val, opts) {{
                  return Object.assign({{ text:displayValue(val), fontSize:9, alignment:'center' }}, opts||{{}});
                }}
                function headerCell(text, opts) {{
                  return Object.assign({{ text:text, bold:true, fontSize:9, alignment:'center' }}, opts||{{}});
                }}
                function rowHasData(row) {{
                  return Object.values(row).some(v => v !== '' && v !== '0');
                }}

                function createPDFDefinition() {{
                  const content = [];

                  // Cover page
                  content.push({{ stack:[
                    {{ image:logoImage, width:750, alignment:'center', margin:[0,20,5,10] }},
                    {{ text:'मासिक आहवाल माहे : '+metadata.monthYear, alignment:'center', bold:true, fontSize:20 }},
                    {{ text:'प्राथमिक आरोग्य केंद्र '+metadata.phcName, alignment:'center', bold:true, fontSize:17 }},
                    {{ text:'उपकेंद्र '+metadata.subcenter, alignment:'center', bold:true, fontSize:17 }}
                  ], margin:[40,10,40,10] }});

                  // Sheet 1
                  if (allData.sheet1) {{
                    content.push({{ text:'', pageBreak:'before' }});
                    const s1 = allData.sheet1;
                    content.push({{ text:s1.title||'', alignment:'center', bold:true, fontSize:13, margin:[0,30,0,8] }});
                    content.push({{ text:'प्रा. आ. केंद्र '+metadata.phcName+' तालुका '+metadata.taluka+' जिल्हा '+metadata.district, fontSize:12, bold:true, alignment:'center', margin:[0,0,0,8] }});
                    content.push({{ text:(s1.subtitle||'')+' '+(s1.month_year||''), alignment:'center', bold:true, fontSize:12, margin:[0,0,0,5] }});
                    content.push({{ text:'उपकेंद्र: '+(s1.subcenter_name||''), bold:true, fontSize:11, margin:[10,5,0,15] }});
                    const tb1 = [];
                    tb1.push([
                      headerCell('अ. क्र.',{{rowSpan:2}}), headerCell('उपकेंद्राचे नाव',{{rowSpan:2}}),
                      headerCell('लोकसंख्या',{{rowSpan:2}}), headerCell('कर्मचाऱ्यांचे नाव वर्गवारी',{{rowSpan:2}}),
                      headerCell('रक्त नमुना वार्षिक उद्दिष्ट',{{rowSpan:2}}),
                      headerCell('मासिक घेतलेले रक्त नमुने',{{colSpan:3}}),{{}},{{}},
                      headerCell('प्रगतीपथावर घेतलेले रक्त नमुने',{{colSpan:3}}),{{}},{{}}
                    ]);
                    tb1.push([{{}},{{}},{{}},{{}},{{}},
                      headerCell('पहिला पंधरावडा'), headerCell('दुसरा पंधरावडा'), headerCell('एकूण'),
                      headerCell('पहिला पंधरावडा'), headerCell('दुसरा पंधरावडा'), headerCell('एकूण')
                    ]);
                    const totalRows = (s1.staff_data?.length||0) + 2;
                    (s1.staff_data||[]).forEach((staff, idx) => {{
                      const m1=parseInt(staff.पहिला_पंधरावडा||0), m2=parseInt(staff.दुसरा_पंधरावडा||0);
                      const p1=parseInt(staff.प्रगती_पहिला||0),   p2=parseInt(staff.प्रगती_दुसरा||0);
                      tb1.push([
                        {{ text:idx===0?(s1.subcenter_sr||'1'):'', fontSize:9, alignment:'center', rowSpan:idx===0?totalRows:1 }},
                        {{ text:idx===0?(s1.subcenter_name||''):'', fontSize:9, alignment:'center', rowSpan:idx===0?totalRows:1 }},
                        {{ text:idx===0?(s1.subcenter_pop||''):'', fontSize:9, alignment:'center', rowSpan:idx===0?totalRows:1 }},
                        cell((staff.नाव||'')+(staff.पदनाम?' ('+staff.पदनाम+')':'')),
                        {{ text:idx===0?(s1.annual_target||''):'', fontSize:9, alignment:'center', rowSpan:idx===0?totalRows:1 }},
                        cell(displayValue(m1)), cell(displayValue(m2)), cell(displayValue(m1+m2)),
                        cell(displayValue(p1)), cell(displayValue(p2)), cell(displayValue(p1+p2))
                      ]);
                    }});
                    const af1=parseInt(s1.asha_data?.f1||0), af2=parseInt(s1.asha_data?.f2||0);
                    const ap1=parseInt(s1.asha_data?.p1||0), ap2=parseInt(s1.asha_data?.p2||0);
                    tb1.push([cell(''),cell(''),cell(''),
                      {{ text:'एकूण आशा कार्यकर्ती\\nसंख्या '+(s1.total_asha_count||''), bold:true, fontSize:9, alignment:'center' }},
                      cell(''), cell(displayValue(af1)), cell(displayValue(af2)), cell(displayValue(af1+af2)),
                      cell(displayValue(ap1)), cell(displayValue(ap2)), cell(displayValue(ap1+ap2))
                    ]);
                    tb1.push([cell(''),cell(''),cell(''),
                      headerCell('एकूण रक्तनमुने'), cell(''),
                      cell(displayValue(s1.totals?.f1),{{bold:true}}), cell(displayValue(s1.totals?.f2),{{bold:true}}), cell(displayValue(s1.totals?.monthly),{{bold:true}}),
                      cell(displayValue(s1.totals?.p1),{{bold:true}}), cell(displayValue(s1.totals?.p2),{{bold:true}}), cell(displayValue(s1.totals?.progress),{{bold:true}})
                    ]);
                    content.push({{ table:{{ widths:['5%','12%','10%','15%','8%','8%','8%','8%','8%','8%','10%'], body:tb1 }}, layout:tableLayout() }});
                  }}

                  // Sheet 2
                  if (allData.sheet2) {{
                    content.push({{ text:'', pageBreak:'before' }});
                    const s2 = allData.sheet2;
                    content.push(...addPageHeading(s2.title1||'थुंकी संकलन', metadata.monthYear));
                    const tb2a = [];
                    tb2a.push([
                      headerCell('अ. क्र.',{{rowSpan:2}}), headerCell('गावाचे नाव',{{rowSpan:2}}),
                      headerCell('लोकसंख्या',{{rowSpan:2}}), headerCell('कर्मचारी नाव',{{rowSpan:2}}),
                      headerCell('मासिक',{{colSpan:3}}),{{}},{{}},
                      headerCell('वार्षिक',{{colSpan:3}}),{{}},{{}}
                    ]);
                    tb2a.push([{{}},{{}},{{}},{{}},
                      headerCell('पुरुष'),headerCell('स्त्री'),headerCell('एकूण'),
                      headerCell('पुरुष'),headerCell('स्त्री'),headerCell('एकूण')
                    ]);
                    (s2.table1||[]).filter(rowHasData).forEach(r => tb2a.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:['5%','15%','12%','15%','9%','9%','9%','9%','9%','8%'], body:tb2a }}, layout:tableLayout(), margin:[0,0,0,20] }});
                    content.push(...addPageHeading2(s2.title2||'संशयीत क्षयरुग्ण तपासणी'));
                    const tb2b = [[
                      headerCell('अ. क्र.'),headerCell('गावाचे नाव'),headerCell('संशयीत रुग्णाचे नाव'),
                      headerCell('वय'),headerCell('लिंग'),headerCell('नमुना घेतलेला दिनांक'),
                      headerCell('तपासणी दिनांक'),headerCell('लॅब क्रमांक'),headerCell('कर्मचारी नाव')
                    ]];
                    (s2.table2||[]).filter(rowHasData).forEach(r => tb2b.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:['5%','12%','15%','7%','7%','12%','12%','12%','12%'], body:tb2b }}, layout:tableLayout() }});
                  }}

                  // Sheet 3
                  if (allData.sheet3) {{
                    content.push({{ text:'', pageBreak:'before' }});
                    const s3 = allData.sheet3;
                    content.push(...addPageHeading(s3.title1||'कुष्ठरुग्ण', metadata.monthYear));
                    const tb3a = [];
                    tb3a.push([
                      headerCell('अ. क्र.',{{rowSpan:2}}),headerCell('गावाचे नाव',{{rowSpan:2}}),headerCell('लोकसंख्या',{{rowSpan:2}}),
                      headerCell('संबंधित कुष्ठ रुग्ण',{{colSpan:3}}),{{}},{{}},
                      headerCell('नवीन शोधलेले (एम.बी.)',{{colSpan:3}}),{{}},{{}},
                      headerCell('नवीन शोधलेले (पी.बी.)',{{colSpan:3}}),{{}},{{}},
                      headerCell('नियमित औषधोपचार',{{colSpan:3}}),{{}},{{}}
                    ]);
                    tb3a.push([{{}},{{}},{{}},
                      headerCell('मुले'),headerCell('प्रौढ'),headerCell('एकूण'),
                      headerCell('मुले'),headerCell('प्रौढ'),headerCell('एकूण'),
                      headerCell('मुले'),headerCell('प्रौढ'),headerCell('एकूण'),
                      headerCell('मुले'),headerCell('प्रौढ'),headerCell('एकूण')
                    ]);
                    (s3.table1||[]).filter(rowHasData).forEach(r => tb3a.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:Array(15).fill('*'), body:tb3a }}, layout:tableLayout(), margin:[0,0,0,20] }});
                    content.push(...addPageHeading2(s3.title2||'कुष्ठरुग्ण तपशील'));
                    const tb3b = [[
                      headerCell('अ. क्र.'),headerCell('गावाचे नाव'),headerCell('संबंधित रुग्णाचे नाव'),
                      headerCell('वय'),headerCell('लिंग'),headerCell('लक्षणे'),headerCell('कर्मचारी नाव')
                    ]];
                    (s3.table2||[]).filter(rowHasData).forEach(r => tb3b.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:['8%','15%','18%','8%','8%','25%','18%'], body:tb3b }}, layout:tableLayout() }});
                  }}

                  // Sheet 4
                  if (allData.sheet4) {{
                    content.push({{ text:'', pageBreak:'before' }});
                    const s4 = allData.sheet4;
                    content.push(...addPageHeading(s4.title1||'क्षयरुग्ण', metadata.monthYear));
                    const tb4a = [];
                    tb4a.push([
                      headerCell('अ. क्र.',{{rowSpan:2}}),headerCell('गावाचे नाव',{{rowSpan:2}}),
                      headerCell('लोकसंख्या',{{rowSpan:2}}),headerCell('कर्मचारी नाव',{{rowSpan:2}}),
                      headerCell('मासिक',{{colSpan:3}}),{{}},{{}},
                      headerCell('वार्षिक',{{colSpan:3}}),{{}},{{}}
                    ]);
                    tb4a.push([{{}},{{}},{{}},{{}},
                      headerCell('पुरुष'),headerCell('स्त्री'),headerCell('एकूण'),
                      headerCell('पुरुष'),headerCell('स्त्री'),headerCell('एकूण')
                    ]);
                    (s4.table1||[]).filter(rowHasData).forEach(r => tb4a.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:['5%','15%','12%','15%','9%','9%','9%','9%','9%','8%'], body:tb4a }}, layout:tableLayout(), margin:[0,0,0,20] }});
                    content.push(...addPageHeading2(s4.title2||'क्षयरुग्ण तपशील'));
                    const tb4b = [[
                      headerCell('अ. क्र.'),headerCell('गावाचे नाव'),headerCell('क्षयरुग्णाचे नाव'),
                      headerCell('वय'),headerCell('लिंग'),headerCell('कॅटेगरी'),
                      headerCell('औषधोपचार सुरू दिनांक'),headerCell('टी. बी. नंबर'),headerCell('कर्मचारी नाव')
                    ]];
                    (s4.table2||[]).filter(rowHasData).forEach(r => tb4b.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:['5%','12%','15%','7%','7%','10%','12%','12%','15%'], body:tb4b }}, layout:tableLayout() }});
                  }}

                  // Sheet 5
                  if (allData.sheet5) {{
                    content.push({{ text:'', pageBreak:'before' }});
                    const s5 = allData.sheet5;
                    content.push(...addPageHeading(s5.title||'कंटेनर सर्वेक्षण', metadata.monthYear));
                    const tb5 = [[
                      headerCell('अ. क्र.'),headerCell('गावाचे नाव'),headerCell('लोकसंख्या'),
                      headerCell('एकूण घरांची संख्या'),headerCell('तपासलेले घरे'),headerCell('दूषित घरे'),
                      headerCell('तपासलेली भांडी'),headerCell('दूषित भांडी'),
                      headerCell('हाऊस इंडेक्स'),headerCell('कंटेनर इंडेक्स'),headerCell('ब्रँट्यू इंडेक्स'),
                      headerCell('रिकामी केलेली भांडी'),headerCell('अँबेट टाकलेली भांडी')
                    ]];
                    (s5.data||[]).filter(rowHasData).forEach(r => tb5.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:Array(13).fill('*'), body:tb5 }}, layout:tableLayout() }});
                  }}

                  // Sheet 6
                  if (allData.sheet6) {{
                    content.push({{ text:'', pageBreak:'before' }});
                    const s6 = allData.sheet6;
                    content.push(...addPageHeading(s6.title1||'डासउत्पत्ती', metadata.monthYear));
                    const tb6a = [];
                    tb6a.push([
                      headerCell('अ. क्र.',{{rowSpan:2}}),headerCell('उपकेंद्राचे नाव',{{rowSpan:2}}),
                      headerCell('गावाचे नाव',{{rowSpan:2}}),headerCell('डास उत्पत्ती स्थाने',{{colSpan:2}}),{{}},
                      headerCell('डासउत्पत्ती स्थानाचे ठिकाण',{{rowSpan:2}})
                    ]);
                    tb6a.push([{{}},{{}},{{}},headerCell('कायम'),headerCell('हंगामी'),{{}}]);
                    (s6.table1||[]).filter(rowHasData).forEach(r => tb6a.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:['8%','18%','18%','12%','12%','32%'], body:tb6a }}, layout:tableLayout(), margin:[0,0,0,20] }});
                    content.push(...addPageHeading2(s6.title2||'गप्पी मासे पैदास केंद्र'));
                    const tb6b = [];
                    tb6b.push([
                      headerCell('अ. क्र.',{{rowSpan:2}}),headerCell('उपकेंद्राचे नाव',{{rowSpan:2}}),
                      headerCell('गावाचे नाव',{{rowSpan:2}}),headerCell('गप्पी मासे पैदास केंद्र',{{colSpan:2}}),{{}},
                      headerCell('गप्पी मासे पैदास केंद्राचे ठिकाण',{{rowSpan:2}})
                    ]);
                    tb6b.push([{{}},{{}},{{}},headerCell('कायम'),headerCell('हंगामी'),{{}}]);
                    (s6.table2||[]).filter(rowHasData).forEach(r => tb6b.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:['8%','18%','18%','12%','12%','32%'], body:tb6b }}, layout:tableLayout() }});
                  }}

                  // Sheet 7
                  if (allData.sheet7) {{
                    content.push({{ text:'', pageBreak:'before' }});
                    const s7 = allData.sheet7;
                    content.push(...addPageHeading(s7.title1||'मोतीबिंदू', metadata.monthYear));
                    const tb7a = [];
                    tb7a.push([
                      headerCell('अ. क्र.',{{rowSpan:2}}),headerCell('गावाचे नाव',{{rowSpan:2}}),headerCell('लोकसंख्या',{{rowSpan:2}}),
                      headerCell('संशयीत मोतीबिंदू',{{colSpan:3}}),{{}},{{}},
                      headerCell('नवीन शोधलेले मोतीबिंदू',{{colSpan:3}}),{{}},{{}}
                    ]);
                    tb7a.push([{{}},{{}},{{}},
                      headerCell('मुले'),headerCell('प्रौढ'),headerCell('एकूण'),
                      headerCell('मुले'),headerCell('प्रौढ'),headerCell('एकूण')
                    ]);
                    (s7.table1||[]).filter(rowHasData).forEach(r => tb7a.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:['8%','15%','12%','10%','10%','10%','10%','10%','10%'], body:tb7a }}, layout:tableLayout(), margin:[0,0,0,20] }});
                    content.push(...addPageHeading2(s7.title2||'मोतीबिंदू रुग्ण तपशील'));
                    const tb7b = [[
                      headerCell('अ. क्र.'),headerCell('गावाचे नाव'),headerCell('संबंधित रुग्णाचे नाव'),
                      headerCell('वय'),headerCell('लिंग'),headerCell('लक्षणे'),headerCell('कर्मचारी नाव')
                    ]];
                    (s7.table2||[]).filter(rowHasData).forEach(r => tb7b.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:['8%','15%','18%','8%','8%','25%','18%'], body:tb7b }}, layout:tableLayout() }});
                  }}

                  // Sheet 8
                  if (allData.sheet8) {{
                    content.push({{ text:'', pageBreak:'before' }});
                    const s8 = allData.sheet8;
                    content.push(...addPageHeading(s8.title1||'प्रयोगशाळा', metadata.monthYear));
                    const tb8a = [];
                    tb8a.push([
                      headerCell('अ. क्र.',{{rowSpan:3}}),headerCell('उपकेंद्राचे नाव',{{rowSpan:3}}),headerCell('एकूण सार्व. उद्भव',{{rowSpan:3}}),
                      headerCell('जैविक पाणी नमुने',{{colSpan:4}}),{{}},{{}},{{}},
                      headerCell('रासायनिक पाणी नमुने',{{colSpan:4}}),{{}},{{}},{{}},
                      headerCell('TCL प्रारंभ',{{rowSpan:3}}),headerCell('TCL खरेदी',{{rowSpan:3}}),
                      headerCell('TCL खर्च',{{rowSpan:3}}),headerCell('TCL शेवट',{{rowSpan:3}}),
                      headerCell('TCL नमुने',{{colSpan:2}}),{{}},
                      headerCell('शौच नमुने',{{colSpan:2}}),{{}},
                      headerCell('मीठ नमुने',{{colSpan:2}}),{{}}
                    ]);
                    tb8a.push([{{}},{{}},{{}},
                      headerCell('महिन्यात',{{colSpan:2}}),{{}},headerCell('प्रगती पर',{{colSpan:2}}),{{}},
                      headerCell('महिन्यात',{{colSpan:2}}),{{}},headerCell('प्रगती पर',{{colSpan:2}}),{{}},
                      {{}},{{}},{{}},{{}},
                      headerCell('महिन्यात',{{colSpan:2}}),{{}},headerCell('महिन्यात',{{colSpan:2}}),{{}},headerCell('महिन्यात',{{colSpan:2}}),{{}}
                    ]);
                    tb8a.push([{{}},{{}},{{}},
                      headerCell('घेतलेले'),headerCell('दुषित'),headerCell('घेतलेले'),headerCell('दुषित'),
                      headerCell('घेतलेले'),headerCell('दुषित'),headerCell('घेतलेले'),headerCell('दुषित'),
                      {{}},{{}},{{}},{{}},
                      headerCell('घेतलेले'),headerCell('दुषित'),headerCell('घेतलेले'),headerCell('दुषित'),headerCell('घेतलेले'),headerCell('दुषित')
                    ]);
                    (s8.table1||[]).filter(rowHasData).forEach(r => {{
                      tb8a.push([
                        cell(r['अ. क्र.']),cell(r['उपकेंद्र']),cell(r['एकूण_सार्व']),
                        cell(r['जै_मही_घे']),cell(r['जै_मही_दू']),cell(r['जै_प्रगती_घे']),cell(r['जै_प्रगती_दू']),
                        cell(r['रा_मही_घे']),cell(r['रा_मही_दू']),cell(r['रा_प्रगती_घे']),cell(r['रा_प्रगती_दू']),
                        cell(r['TCL_प्रारंभ']),cell(r['TCL_खरेदी']),cell(r['TCL_खर्च']),cell(r['TCL_शेवट']),
                        cell(r['TCL_नमुना_मही_घे']),cell(r['TCL_नमुना_मही_दू']),
                        cell(r['शौच_मही_घे']),cell(r['शौच_मही_दू']),
                        cell(r['मीठ_मही_घे']),cell(r['मीठ_मही_दू'])
                      ]);
                    }});
                    content.push({{ table:{{ widths:Array(21).fill('*'), body:tb8a }},
                      layout:{{ hLineWidth:()=>0.5, vLineWidth:()=>0.5, paddingLeft:()=>1, paddingRight:()=>1, paddingTop:()=>3, paddingBottom:()=>3 }},
                      margin:[0,0,0,20], fontSize:7 }});
                    content.push(...addPageHeading2(s8.title2||'गावनिहाय TCL साठा'));
                    const tb8b = [[
                      headerCell('अ. क्र.'),headerCell('उपकेंद्राचे नाव'),headerCell('एकूण ग्रामपंचायत'),
                      headerCell('एकूण गावे'),headerCell('टीसीएल साठा'),headerCell('टीसीएल साठवण'),
                      headerCell('पाणी शुद्धीकरण'),headerCell('टीसीएल नसलेले गावे')
                    ]];
                    (s8.table2||[]).filter(rowHasData).forEach(r => tb8b.push(Object.values(r).map(v=>cell(v))));
                    content.push({{ table:{{ widths:['8%','18%','12%','12%','12%','12%','14%','12%'], body:tb8b }}, layout:tableLayout() }});
                  }}

                  return {{
                    pageSize:'A4', pageOrientation:'landscape',
                    pageMargins:[15,15,15,15],
                    content:content,
                    defaultStyle:{{ font:'MarathiFont', fontSize:9 }}
                  }};
                }}

                function previewPDF() {{
                  try {{
                    showStatus('📄 PDF तयार करत आहे...','#FFF9C4');
                    pdfMake.createPdf(createPDFDefinition()).open();
                    showStatus('✅ PDF तयार झाली!','#C8E6C9');
                  }} catch(e) {{
                    showStatus('❌ त्रुटी: '+e.message,'#FFCDD2');
                    console.error('PDF Error:',e);
                  }}
                }}
                function downloadPDF() {{
                  try {{
                    showStatus('⬇️ PDF डाउनलोड करत आहे...','#FFF9C4');
                    pdfMake.createPdf(createPDFDefinition()).download('Masik_Ahwal_'+metadata.monthYear.replace(/\\s+/g,'_')+'.pdf');
                    showStatus('✅ PDF डाउनलोड झाली!','#C8E6C9');
                  }} catch(e) {{
                    showStatus('❌ त्रुटी: '+e.message,'#FFCDD2');
                    console.error('PDF Error:',e);
                  }}
                }}
              </script>
            </body>
            </html>
            """,
            height=120
        )


if __name__ == "__main__":
    mothly_final_report()