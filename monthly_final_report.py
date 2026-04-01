import streamlit as st
import json
import base64
import streamlit.components.v1 as components
from pathlib import Path


def mothly_final_report():
    """
    Manual data entry for 8 different health reports with landscape PDF generation
    """

    st.title("📊 मासिक आहवाल - Manual Data Entry")

    # Common metadata
    st.sidebar.header("मुख्य माहिती")
    month_year = st.sidebar.text_input("महिना/वर्ष", value="मार्च २०२६")
    phc_name = st.sidebar.text_input("प्राथमिक आरोग्य केंद्र", value="शेळगाव")
    taluka = st.sidebar.text_input("तालुका", value="इंदापूर")
    district = st.sidebar.text_input("जिल्हा", value="पुणे")
    sub_center = st.sidebar.text_input("उपकेंद्र", value="")
    population = st.sidebar.text_input("लोकसंख्या", value="")

    # Tab selection for different sheets
    tabs = st.tabs([
        "1️⃣ रक्त नमुना",
        "2️⃣ थुंकी संकलन (2 Tables)",
        "3️⃣ कुष्ठरुग्ण (2 Tables)",
        "4️⃣ क्षय रुग्ण (2 Tables)",
        "5️⃣ कंटेनर सर्वेक्षण",
        "6️⃣ डासउत्पत्ती (2 Tables)",
        "7️⃣ मोतीबिंदू (2 Tables)",
        "8️⃣ प्रयोगशाळा (2 Tables)",
        "9️⃣ PDF तयार करा"
    ])

    # Initialize session state for all sheets
    if 'sheet_data' not in st.session_state:
        st.session_state.sheet_data = {}

    # Sheet 1: रक्त नमुना मासिक अहवाल (NO CHANGES - Keep as is)
    with tabs[0]:
        st.subheader("रक्त नमुना मासिक अहवाल")

        # Basic subcenter info - entered once
        st.markdown("### उपकेंद्र माहिती")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            subcenter_sr = st.text_input("अ. क्र.", value="1", key="s1_basic_sr")
        with col2:
            subcenter_name = st.text_input("उपकेंद्राचे नाव", value="", key="s1_basic_name")
        with col3:
            subcenter_pop = st.text_input("लोकसंख्या", value="", key="s1_basic_pop")
        with col4:
            annual_target = st.text_input("रक्त नमुना वार्षिक उद्दिष्ट", value="", key="s1_annual_target")

        st.markdown("---")
        st.markdown("### कर्मचारी माहिती (आरोग्य सेविका व सेवक)")

        num_staff = st.number_input("कर्मचारी संख्या", min_value=1, max_value=20, value=3, key="rows1")

        staff_data = []
        designation_options = ["आरोग्य सेवक", "आरोग्य सेविका", "आरोग्य सेविका NHM"]

        for i in range(num_staff):
            st.markdown(f"**कर्मचारी {i + 1}**")
            cols = st.columns([2, 2, 1, 1, 1, 1])

            staff_name = cols[0].text_input(f"नाव", value="", key=f"s1_name_{i}")
            staff_designation = cols[1].selectbox(f"पदनाम", options=designation_options, key=f"s1_desg_{i}")

            row_data = {
                'नाव': staff_name,
                'पदनाम': staff_designation,
                'पहिला_पंधरावडा': cols[2].text_input(f"मासिक पहिला", value="0", key=f"s1_f1_{i}"),
                'दुसरा_पंधरावडा': cols[3].text_input(f"मासिक दुसरा", value="0", key=f"s1_f2_{i}"),
                'प्रगती_पहिला': cols[4].text_input(f"प्रगती पहिला", value="0", key=f"s1_p1_{i}"),
                'प्रगती_दुसरा': cols[5].text_input(f"प्रगती दुसरा", value="0", key=f"s1_p2_{i}"),
            }
            staff_data.append(row_data)

        st.markdown("---")
        st.markdown("### एकूण आशा कार्यकर्ती")
        total_asha_count = st.text_input("एकूण आशा कार्यकर्ती संख्या", value="0", key="s1_total_asha")

        # Get monthly and progress values for ASHA row
        cols_asha = st.columns([2, 2, 1, 1, 1, 1])
        with cols_asha[0]:
            st.text("(मासिक व प्रगती)")
        asha_f1 = cols_asha[2].text_input("मासिक पहिला", value="0", key="s1_asha_f1")
        asha_f2 = cols_asha[3].text_input("मासिक दुसरा", value="0", key="s1_asha_f2")
        asha_p1 = cols_asha[4].text_input("प्रगती पहिला", value="0", key="s1_asha_p1")
        asha_p2 = cols_asha[5].text_input("प्रगती दुसरा", value="0", key="s1_asha_p2")

        # Calculate totals automatically
        total_f1 = sum(int(s['पहिला_पंधरावडा'] or 0) for s in staff_data) + int(asha_f1 or 0)
        total_f2 = sum(int(s['दुसरा_पंधरावडा'] or 0) for s in staff_data) + int(asha_f2 or 0)
        total_monthly = total_f1 + total_f2
        total_p1 = sum(int(s['प्रगती_पहिला'] or 0) for s in staff_data) + int(asha_p1 or 0)
        total_p2 = sum(int(s['प्रगती_दुसरा'] or 0) for s in staff_data) + int(asha_p2 or 0)
        total_progress = total_p1 + total_p2

        st.info(
            f"📊 एकूण रक्तनमुने - मासिक: पहिला={total_f1}, दुसरा={total_f2}, एकूण={total_monthly} | प्रगती: पहिला={total_p1}, दुसरा={total_p2}, एकूण={total_progress}")

        st.session_state.sheet_data['sheet1'] = {
            'title': 'राष्ट्रीय कीटकजन्य रोग नियंत्रण कार्यक्रम, जिल्हा पुणे',
            'subtitle': 'रक्त नमुना मासिक अहवाल',
            'month_year': month_year,
            'subcenter_sr': subcenter_sr,
            'subcenter_name': subcenter_name,
            'subcenter_pop': subcenter_pop,
            'annual_target': annual_target,
            'staff_data': staff_data,
            'total_asha_count': total_asha_count,
            'asha_data': {
                'f1': asha_f1,
                'f2': asha_f2,
                'p1': asha_p1,
                'p2': asha_p2
            },
            'totals': {
                'f1': str(total_f1),
                'f2': str(total_f2),
                'monthly': str(total_monthly),
                'p1': str(total_p1),
                'p2': str(total_p2),
                'progress': str(total_progress)
            }
        }

    # Sheet 2: थुंकी संकलन - TWO TABLES
    with tabs[1]:
        st.subheader("थुंकी संकलन अहवाल - 2 Tables")

        # Table 1
        st.markdown("### तक्ता १: गावनिहाय थुंकी संकलन")
        num_rows_2a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=4, key="rows2a")

        sheet2_table1 = []
        for i in range(num_rows_2a):
            cols = st.columns([1, 2, 2, 2, 1, 1, 1, 1, 1, 1])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s2a_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value="", key=f"s2a_village_{i}"),
                'लोकसंख्या': cols[2].text_input(f"लोकसंख्या", value="", key=f"s2a_pop_{i}"),
                'कर्मचारी': cols[3].text_input(f"कर्मचारी", value="", key=f"s2a_staff_{i}"),
                'मासिक_पुरुष': cols[4].text_input(f"मासिक पुरुष", value="", key=f"s2a_m_m_{i}"),
                'मासिक_स्त्री': cols[5].text_input(f"मासिक स्त्री", value="", key=f"s2a_m_f_{i}"),
                'मासिक_एकूण': cols[6].text_input(f"मासिक एकूण", value="", key=f"s2a_m_t_{i}"),
                'वार्षिक_पुरुष': cols[7].text_input(f"वार्षिक पुरुष", value="", key=f"s2a_y_m_{i}"),
                'वार्षिक_स्त्री': cols[8].text_input(f"वार्षिक स्त्री", value="", key=f"s2a_y_f_{i}"),
                'वार्षिक_एकूण': cols[9].text_input(f"वार्षिक एकूण", value="", key=f"s2a_y_t_{i}"),
            }
            sheet2_table1.append(row_data)

        st.markdown("---")

        # Table 2
        st.markdown("### तक्ता २: संशयीत रुग्ण माहिती")
        num_rows_2b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=4, key="rows2b")

        sheet2_table2 = []
        for i in range(num_rows_2b):
            cols = st.columns([1, 2, 2, 1, 1, 2, 2, 2, 2])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s2b_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value="", key=f"s2b_village_{i}"),
                'संशयीत_रुग्ण': cols[2].text_input(f"रुग्णाचे नाव", value="", key=f"s2b_name_{i}"),
                'वय': cols[3].text_input(f"वय", value="", key=f"s2b_age_{i}"),
                'लिंग': cols[4].text_input(f"लिंग", value="", key=f"s2b_gender_{i}"),
                'नमुना_दिनांक': cols[5].text_input(f"नमुना दिनांक", value="", key=f"s2b_sample_{i}"),
                'तपासणी_दिनांक': cols[6].text_input(f"तपासणी दिनांक", value="", key=f"s2b_test_{i}"),
                'लॅब_क्रमांक': cols[7].text_input(f"लॅब क्रमांक", value="", key=f"s2b_lab_{i}"),
                'कर्मचारी': cols[8].text_input(f"कर्मचारी", value="", key=f"s2b_staff_{i}"),
            }
            sheet2_table2.append(row_data)

        st.session_state.sheet_data['sheet2'] = {
            'title1': 'थुंकी संकलन गावनिहाय अहवाल',
            'title2': 'संशयीत क्षयरुग्ण तपासणी अहवाल',
            'table1': sheet2_table1,
            'table2': sheet2_table2
        }

    # Sheet 3: कुष्ठरुग्ण - TWO TABLES
    with tabs[2]:
        st.subheader("कुष्ठरुग्ण मासिक अहवाल - 2 Tables")

        # Table 1
        st.markdown("### तक्ता १: गावनिहाय कुष्ठरुग्ण")
        num_rows_3a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=4, key="rows3a")

        sheet3_table1 = []
        for i in range(num_rows_3a):
            cols = st.columns([1, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s3a_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value="", key=f"s3a_village_{i}"),
                'लोकसंख्या': cols[2].text_input(f"लोकसंख्या", value="", key=f"s3a_pop_{i}"),
                'संबंधित_मुले': cols[3].text_input(f"सं.मुले", value="", key=f"s3a_rel_c_{i}"),
                'संबंधित_प्रौढ': cols[4].text_input(f"सं.प्रौढ", value="", key=f"s3a_rel_a_{i}"),
                'संबंधित_एकूण': cols[5].text_input(f"सं.एकूण", value="", key=f"s3a_rel_t_{i}"),
                'MB_मुले': cols[6].text_input(f"MB मुले", value="", key=f"s3a_mb_c_{i}"),
                'MB_प्रौढ': cols[7].text_input(f"MB प्रौढ", value="", key=f"s3a_mb_a_{i}"),
                'MB_एकूण': cols[8].text_input(f"MB एकूण", value="", key=f"s3a_mb_t_{i}"),
                'PB_मुले': cols[9].text_input(f"PB मुले", value="", key=f"s3a_pb_c_{i}"),
                'PB_प्रौढ': cols[10].text_input(f"PB प्रौढ", value="", key=f"s3a_pb_a_{i}"),
                'PB_एकूण': cols[11].text_input(f"PB एकूण", value="", key=f"s3a_pb_t_{i}"),
                'औषधोपचार_मुले': cols[12].text_input(f"औ.मुले", value="", key=f"s3a_tr_c_{i}"),
                'औषधोपचार_प्रौढ': cols[13].text_input(f"औ.प्रौढ", value="", key=f"s3a_tr_a_{i}"),
                'औषधोपचार_एकूण': cols[14].text_input(f"औ.एकूण", value="", key=f"s3a_tr_t_{i}"),
            }
            sheet3_table1.append(row_data)

        st.markdown("---")

        # Table 2
        st.markdown("### तक्ता २: कुष्ठरुग्ण तपशील")
        num_rows_3b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=3, key="rows3b")

        sheet3_table2 = []
        for i in range(num_rows_3b):
            cols = st.columns([1, 2, 2, 1, 1, 3, 2])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s3b_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value="", key=f"s3b_village_{i}"),
                'रुग्णाचे_नाव': cols[2].text_input(f"रुग्ण नाव", value="", key=f"s3b_name_{i}"),
                'वय': cols[3].text_input(f"वय", value="", key=f"s3b_age_{i}"),
                'लिंग': cols[4].text_input(f"लिंग", value="", key=f"s3b_gender_{i}"),
                'लक्षणे': cols[5].text_input(f"लक्षणे", value="", key=f"s3b_symptoms_{i}"),
                'कर्मचारी': cols[6].text_input(f"कर्मचारी", value="", key=f"s3b_staff_{i}"),
            }
            sheet3_table2.append(row_data)

        st.session_state.sheet_data['sheet3'] = {
            'title1': 'कुष्ठरुग्ण गावनिहाय मासिक अहवाल',
            'title2': 'कुष्ठरुग्ण तपशीलवार माहिती',
            'table1': sheet3_table1,
            'table2': sheet3_table2
        }

    # Sheet 4: क्षय रुग्ण - TWO TABLES
    with tabs[3]:
        st.subheader("क्षय रुग्ण अहवाल - 2 Tables")

        # Table 1
        st.markdown("### तक्ता १: गावनिहाय क्षय रुग्ण")
        num_rows_4a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=3, key="rows4a")

        sheet4_table1 = []
        for i in range(num_rows_4a):
            cols = st.columns([1, 2, 2, 2, 1, 1, 1, 1, 1, 1])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s4a_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value="", key=f"s4a_village_{i}"),
                'लोकसंख्या': cols[2].text_input(f"लोकसंख्या", value="", key=f"s4a_pop_{i}"),
                'कर्मचारी': cols[3].text_input(f"कर्मचारी", value="", key=f"s4a_staff_{i}"),
                'मासिक_पुरुष': cols[4].text_input(f"मासिक पुरुष", value="", key=f"s4a_m_m_{i}"),
                'मासिक_स्त्री': cols[5].text_input(f"मासिक स्त्री", value="", key=f"s4a_m_f_{i}"),
                'मासिक_एकूण': cols[6].text_input(f"मासिक एकूण", value="", key=f"s4a_m_t_{i}"),
                'वार्षिक_पुरुष': cols[7].text_input(f"वार्षिक पुरुष", value="", key=f"s4a_y_m_{i}"),
                'वार्षिक_स्त्री': cols[8].text_input(f"वार्षिक स्त्री", value="", key=f"s4a_y_f_{i}"),
                'वार्षिक_एकूण': cols[9].text_input(f"वार्षिक एकूण", value="", key=f"s4a_y_t_{i}"),
            }
            sheet4_table1.append(row_data)

        st.markdown("---")

        # Table 2
        st.markdown("### तक्ता २: क्षयरुग्ण तपशील")
        num_rows_4b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=3, key="rows4b")

        sheet4_table2 = []
        for i in range(num_rows_4b):
            cols = st.columns([1, 2, 2, 1, 1, 2, 2, 2, 2])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s4b_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value="", key=f"s4b_village_{i}"),
                'क्षयरुग्णाचे_नाव': cols[2].text_input(f"रुग्ण नाव", value="", key=f"s4b_name_{i}"),
                'वय': cols[3].text_input(f"वय", value="", key=f"s4b_age_{i}"),
                'लिंग': cols[4].text_input(f"लिंग", value="", key=f"s4b_gender_{i}"),
                'कॅटेगरी': cols[5].text_input(f"कॅटेगरी", value="", key=f"s4b_cat_{i}"),
                'औषधोपचार_दिनांक': cols[6].text_input(f"दिनांक", value="", key=f"s4b_date_{i}"),
                'TB_नंबर': cols[7].text_input(f"TB नं.", value="", key=f"s4b_tb_{i}"),
                'कर्मचारी': cols[8].text_input(f"कर्मचारी", value="", key=f"s4b_staff_{i}"),
            }
            sheet4_table2.append(row_data)

        st.session_state.sheet_data['sheet4'] = {
            'title1': 'क्षयरुग्ण गावनिहाय मासिक अहवाल',
            'title2': 'उपचार घेणारे क्षयरुग्ण तपशील',
            'table1': sheet4_table1,
            'table2': sheet4_table2
        }

    # Sheet 5: कंटेनर सर्वेक्षण - ONE TABLE
    with tabs[4]:
        st.subheader("कंटेनर सर्वेक्षण अहवाल")

        num_rows_5 = st.number_input("नोंदी संख्या", min_value=1, max_value=20, value=3, key="rows5")

        sheet5_data = []
        for i in range(num_rows_5):
            cols = st.columns([1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s5_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value="", key=f"s5_village_{i}"),
                'लोकसंख्या': cols[2].text_input(f"लोकसंख्या", value="", key=f"s5_pop_{i}"),
                'एकूण_घरे': cols[3].text_input(f"एकूण घरे", value="", key=f"s5_house_{i}"),
                'तपासलेले_घरे': cols[4].text_input(f"तपासले", value="", key=f"s5_check_{i}"),
                'दूषित_घरे': cols[5].text_input(f"दूषित", value="", key=f"s5_cont_{i}"),
                'तपासलेली_भांडी': cols[6].text_input(f"भांडी", value="", key=f"s5_cont_c_{i}"),
                'दूषित_भांडी': cols[7].text_input(f"दूषित भांडी", value="", key=f"s5_cont_cc_{i}"),
                'House_Index': cols[8].text_input(f"HI", value="", key=f"s5_hi_{i}"),
                'Container_Index': cols[9].text_input(f"CI", value="", key=f"s5_ci_{i}"),
                'Breteau_Index': cols[10].text_input(f"BI", value="", key=f"s5_bi_{i}"),
                'रिकामी_भांडी': cols[11].text_input(f"रिकामी", value="", key=f"s5_empty_{i}"),
                'अँबेट_भांडी': cols[12].text_input(f"अँबेट", value="", key=f"s5_abate_{i}"),
            }
            sheet5_data.append(row_data)

        st.session_state.sheet_data['sheet5'] = {
            'title': 'कंटेनर सर्वेक्षण मासिक अहवाल',
            'data': sheet5_data
        }

    # Sheet 6: डासउत्पत्ती - TWO TABLES
    with tabs[5]:
        st.subheader("डासउत्पत्ती स्थाने - 2 Tables")

        # Table 1
        st.markdown("### तक्ता १: डासउत्पत्ती स्थाने")
        num_rows_6a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=4, key="rows6a")

        sheet6_table1 = []
        for i in range(num_rows_6a):
            cols = st.columns([1, 2, 2, 1, 1, 3])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s6a_sr_{i}"),
                'उपकेंद्राचे नाव': cols[1].text_input(f"उपकेंद्र", value="", key=f"s6a_sub_{i}"),
                'गावाचे नाव': cols[2].text_input(f"गाव", value="", key=f"s6a_village_{i}"),
                'कायम': cols[3].text_input(f"कायम", value="", key=f"s6a_perm_{i}"),
                'हंगामी': cols[4].text_input(f"हंगामी", value="", key=f"s6a_seas_{i}"),
                'ठिकाण': cols[5].text_input(f"ठिकाण", value="", key=f"s6a_loc_{i}"),
            }
            sheet6_table1.append(row_data)

        st.markdown("---")

        # Table 2
        st.markdown("### तक्ता २: गप्पी मासे पैदास केंद्र")
        num_rows_6b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=4, key="rows6b")

        sheet6_table2 = []
        for i in range(num_rows_6b):
            cols = st.columns([1, 2, 2, 1, 1, 3])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s6b_sr_{i}"),
                'उपकेंद्राचे नाव': cols[1].text_input(f"उपकेंद्र", value="", key=f"s6b_sub_{i}"),
                'गावाचे नाव': cols[2].text_input(f"गाव", value="", key=f"s6b_village_{i}"),
                'कायम': cols[3].text_input(f"कायम", value="", key=f"s6b_perm_{i}"),
                'हंगामी': cols[4].text_input(f"हंगामी", value="", key=f"s6b_seas_{i}"),
                'ठिकाण': cols[5].text_input(f"ठिकाण", value="", key=f"s6b_loc_{i}"),
            }
            sheet6_table2.append(row_data)

        st.session_state.sheet_data['sheet6'] = {
            'title1': 'डासउत्पत्ती स्थानांची गावनिहाय यादी',
            'title2': 'गप्पी मासे पैदास केंद्राची यादी',
            'table1': sheet6_table1,
            'table2': sheet6_table2
        }

    # Sheet 7: मोतीबिंदू - TWO TABLES
    with tabs[6]:
        st.subheader("मोतीबिंदू मासिक अहवाल - 2 Tables")

        # Table 1
        st.markdown("### तक्ता १: गावनिहाय मोतीबिंदू")
        num_rows_7a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=4, key="rows7a")

        sheet7_table1 = []
        for i in range(num_rows_7a):
            cols = st.columns([1, 2, 2, 1, 1, 1, 1, 1, 1])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s7a_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value="", key=f"s7a_village_{i}"),
                'लोकसंख्या': cols[2].text_input(f"लोकसंख्या", value="", key=f"s7a_pop_{i}"),
                'संशयीत_मुले': cols[3].text_input(f"सं.मुले", value="", key=f"s7a_susp_c_{i}"),
                'संशयीत_प्रौढ': cols[4].text_input(f"सं.प्रौढ", value="", key=f"s7a_susp_a_{i}"),
                'संशयीत_एकूण': cols[5].text_input(f"सं.एकूण", value="", key=f"s7a_susp_t_{i}"),
                'नवीन_मुले': cols[6].text_input(f"नवीन मुले", value="", key=f"s7a_new_c_{i}"),
                'नवीन_प्रौढ': cols[7].text_input(f"नवीन प्रौढ", value="", key=f"s7a_new_a_{i}"),
                'नवीन_एकूण': cols[8].text_input(f"नवीन एकूण", value="", key=f"s7a_new_t_{i}"),
            }
            sheet7_table1.append(row_data)

        st.markdown("---")

        # Table 2
        st.markdown("### तक्ता २: मोतीबिंदू रुग्ण तपशील")
        num_rows_7b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=4, key="rows7b")

        sheet7_table2 = []
        for i in range(num_rows_7b):
            cols = st.columns([1, 2, 2, 1, 1, 3, 2])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s7b_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value="", key=f"s7b_village_{i}"),
                'रुग्णाचे_नाव': cols[2].text_input(f"रुग्ण नाव", value="", key=f"s7b_name_{i}"),
                'वय': cols[3].text_input(f"वय", value="", key=f"s7b_age_{i}"),
                'लिंग': cols[4].text_input(f"लिंग", value="", key=f"s7b_gender_{i}"),
                'लक्षणे': cols[5].text_input(f"लक्षणे", value="", key=f"s7b_symptoms_{i}"),
                'कर्मचारी': cols[6].text_input(f"कर्मचारी", value="", key=f"s7b_staff_{i}"),
            }
            sheet7_table2.append(row_data)

        st.session_state.sheet_data['sheet7'] = {
            'title1': 'मोतीबिंदू गावनिहाय मासिक अहवाल',
            'title2': 'मोतीबिंदू रुग्ण तपशीलवार माहिती',
            'table1': sheet7_table1,
            'table2': sheet7_table2
        }

    # Sheet 8: प्रयोगशाळा - TWO TABLES
    with tabs[7]:
        st.subheader("प्रयोगशाळा अहवाल - 2 Tables")

        st.warning(
            "⚠️ Sheet 8 has complex nested columns. Simplified data entry is shown. Full structure will be in PDF.")

        # Table 1 - Simplified
        st.markdown("### तक्ता १: प्रयोगशाळा तपासणी")
        num_rows_8a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=3, key="rows8a")

        sheet8_table1 = []
        for i in range(num_rows_8a):
            st.markdown(f"**नोंद {i + 1}**")
            cols = st.columns(3)
            sr = cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s8a_sr_{i}")
            subcenter = cols[1].text_input(f"उपकेंद्र", value="", key=f"s8a_sub_{i}")
            total_surveys = cols[2].text_input(f"एकूण सार्व. उद्भव", value="", key=f"s8a_total_{i}")

            st.markdown("##### जैविक पाणी")
            cols2 = st.columns(4)
            bio_m_taken = cols2[0].text_input(f"महिना घेतले", value="", key=f"s8a_bio_mt_{i}")
            bio_m_cont = cols2[1].text_input(f"महिना दूषित", value="", key=f"s8a_bio_mc_{i}")
            bio_y_taken = cols2[2].text_input(f"प्रगती घेतले", value="", key=f"s8a_bio_yt_{i}")
            bio_y_cont = cols2[3].text_input(f"प्रगती दूषित", value="", key=f"s8a_bio_yc_{i}")

            st.markdown("##### रासायनिक पाणी")
            cols3 = st.columns(4)
            chem_m_taken = cols3[0].text_input(f"महिना घेतले", value="", key=f"s8a_chem_mt_{i}")
            chem_m_cont = cols3[1].text_input(f"महिना दूषित", value="", key=f"s8a_chem_mc_{i}")
            chem_y_taken = cols3[2].text_input(f"प्रगती घेतले", value="", key=f"s8a_chem_yt_{i}")
            chem_y_cont = cols3[3].text_input(f"प्रगती दूषित", value="", key=f"s8a_chem_yc_{i}")

            st.markdown("##### TCL साठा")
            cols4 = st.columns(4)
            tcl_start = cols4[0].text_input(f"प्रारंभ", value="", key=f"s8a_tcl_s_{i}")
            tcl_buy = cols4[1].text_input(f"खरेदी", value="", key=f"s8a_tcl_b_{i}")
            tcl_use = cols4[2].text_input(f"खर्च", value="", key=f"s8a_tcl_u_{i}")
            tcl_end = cols4[3].text_input(f"शेवट", value="", key=f"s8a_tcl_e_{i}")

            # Additional samples (simplified)
            st.markdown("##### अतिरिक्त नमुने")
            cols5 = st.columns(6)
            tcl_sample_m_t = cols5[0].text_input(f"TCL मही घे", value="", key=f"s8a_tcls_mt_{i}")
            tcl_sample_m_c = cols5[1].text_input(f"TCL मही दू", value="", key=f"s8a_tcls_mc_{i}")
            stool_m_t = cols5[2].text_input(f"शौच मही घे", value="", key=f"s8a_st_mt_{i}")
            stool_m_c = cols5[3].text_input(f"शौच मही दू", value="", key=f"s8a_st_mc_{i}")
            salt_m_t = cols5[4].text_input(f"मीठ मही घे", value="", key=f"s8a_salt_mt_{i}")
            salt_m_c = cols5[5].text_input(f"मीठ मही दू", value="", key=f"s8a_salt_mc_{i}")

            row_data = {
                'अ. क्र.': sr,
                'उपकेंद्र': subcenter,
                'एकूण_सार्व': total_surveys,
                'जै_मही_घे': bio_m_taken,
                'जै_मही_दू': bio_m_cont,
                'जै_प्रगती_घे': bio_y_taken,
                'जै_प्रगती_दू': bio_y_cont,
                'रा_मही_घे': chem_m_taken,
                'रा_मही_दू': chem_m_cont,
                'रा_प्रगती_घे': chem_y_taken,
                'रा_प्रगती_दू': chem_y_cont,
                'TCL_प्रारंभ': tcl_start,
                'TCL_खरेदी': tcl_buy,
                'TCL_खर्च': tcl_use,
                'TCL_शेवट': tcl_end,
                'TCL_नमुना_मही_घे': tcl_sample_m_t,
                'TCL_नमुना_मही_दू': tcl_sample_m_c,
                'शौच_मही_घे': stool_m_t,
                'शौच_मही_दू': stool_m_c,
                'मीठ_मही_घे': salt_m_t,
                'मीठ_मही_दू': salt_m_c,
            }
            sheet8_table1.append(row_data)
            st.markdown("---")

        # Table 2 - Simpler
        st.markdown("### तक्ता २: गावनिहाय TCL साठा")
        num_rows_8b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=4, key="rows8b")

        sheet8_table2 = []
        for i in range(num_rows_8b):
            cols = st.columns([1, 2, 2, 2, 2, 2, 2, 2])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=str(i + 1), key=f"s8b_sr_{i}"),
                'उपकेंद्र': cols[1].text_input(f"उपकेंद्र", value="", key=f"s8b_sub_{i}"),
                'ग्रामपंचायत': cols[2].text_input(f"ग्रामपंचायत", value="", key=f"s8b_gp_{i}"),
                'गावे': cols[3].text_input(f"गावे", value="", key=f"s8b_villages_{i}"),
                'TCL_साठा': cols[4].text_input(f"TCL साठा", value="", key=f"s8b_stock_{i}"),
                'TCL_साठवण': cols[5].text_input(f"TCL साठवण", value="", key=f"s8b_storage_{i}"),
                'पाणी_शुद्धी': cols[6].text_input(f"पाणी शुद्धी", value="", key=f"s8b_purif_{i}"),
                'TCL_नसलेले': cols[7].text_input(f"TCL नसलेले", value="", key=f"s8b_no_tcl_{i}"),
            }
            sheet8_table2.append(row_data)

        st.session_state.sheet_data['sheet8'] = {
            'title1': 'राज्य आयोग्य प्रयोगशाळा विविध नमुने तपासणी अहवाल',
            'title2': 'गावनिहाय TCL साठा अहवाल',
            'table1': sheet8_table1,
            'table2': sheet8_table2
        }

    # Tab 9: Generate PDF
    with tabs[8]:
        st.subheader("📄 PDF तयार करा")

        st.info("सर्व डेटा भरल्यानंतर येथे PDF तयार करा")

        # Check for font file
        BASE_DIR = Path(__file__).resolve().parent
        font_path = BASE_DIR / "fonts" / "NotoSerifDevanagari-VariableFont_wdth,wght.ttf"
        imgpath = BASE_DIR / "fonts" / "img.png"
        if not font_path.exists():
            st.error(f"❌ Font file missing at: {font_path}")
            st.info(
                "Please create a 'fonts' folder in the same directory as this script and add the Marathi font file.")
            st.markdown("""
            **Download font from:**
            - [Google Fonts - Noto Serif Devanagari](https://fonts.google.com/noto/specimen/Noto+Serif+Devanagari)
            - Or use any other Devanagari Unicode font (.ttf format)
            """)
            return

        # Read and encode font
        font_b64 = base64.b64encode(font_path.read_bytes()).decode()
        img = base64.b64encode(imgpath.read_bytes()).decode()
        # Prepare all data for PDF
        all_sheets_json = json.dumps(st.session_state.sheet_data, ensure_ascii=False)

        # PDF generation component with local font
        components.html(
            f"""
            <html>
            <head>
              <meta charset="utf-8" />
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.72/pdfmake.min.js"></script>
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.72/vfs_fonts.js"></script>
              <style>
                body {{ padding: 20px; font-family: Arial, sans-serif; }}
                .button-container {{ margin: 20px 0; }}
                button {{
                  padding: 14px 28px;
                  margin-right: 10px;
                  border: none;
                  border-radius: 6px;
                  cursor: pointer;
                  font-size: 16px;
                  font-weight: bold;
                  box-shadow: 0 3px 6px rgba(0,0,0,0.3);
                }}
                .preview-btn {{ background: #2196F3; color: white; }}
                .download-btn {{ background: #4CAF50; color: white; }}
                #status {{
                  margin-top: 15px;
                  padding: 12px;
                  border-radius: 4px;
                  display: none;
                }}
              </style>
            </head>

            <body>
              <div class="button-container">
                <button class="preview-btn" onclick="previewPDF()">
                  👁️ Preview PDF
                </button>
                <button class="download-btn" onclick="downloadPDF()">
                  ⬇️ Download PDF
                </button>
              </div>

              <div id="status"></div>

              <script>
                const logoImage = "data:image/png;base64,{img}";
                const allData = {all_sheets_json};
                const metadata = {{
                  monthYear: "{month_year}",
                  phcName: "{phc_name}",
                  taluka: "{taluka}",
                  district: "{district}",
                  subcenter: "{sub_center}",
                  population: "{population}"
                }};

                pdfMake.vfs["MarathiFont.ttf"] = "{font_b64}";
                pdfMake.fonts = {{
                  MarathiFont: {{
                    normal: "MarathiFont.ttf",
                    bold: "MarathiFont.ttf",
                    italics: "MarathiFont.ttf",
                    bolditalics: "MarathiFont.ttf"
                  }}
                }};

                function showStatus(msg, color) {{
                  const status = document.getElementById('status');
                  status.innerHTML = msg;
                  status.style.background = color;
                  status.style.display = 'block';
                  setTimeout(() => {{ status.style.display = 'none'; }}, 3000);
                }}

                function displayValue(val) {{
                  if (val === '0' || val === 0 || val === '') return '';
                  return val;
                }}

                function addPageHeading2(title, monthYear) {{
                  return [

                    {{
                      text: title + ' ' + (monthYear || metadata.monthYear),
                      alignment: 'center',
                      bold: true,
                      fontSize: 12,
                      margin: [0, 10, 0, 15]
                    }}
                  ];
                }}
                function addPageHeading(title, monthYear) {{
                  return [
                    {{
                      text: 'राष्ट्रीय कीटकजन्य रोग नियंत्रण कार्यक्रम, जिल्हा पुणे',
                      alignment: 'center',
                      bold: true,
                      fontSize: 13,
                      margin: [0, 20, 0, 4]
                    }},
                    {{
                      text: 'प्रा. आ. केंद्र ' + metadata.phcName + ' तालुका ' + metadata.taluka + ' जिल्हा ' + metadata.district,
                      fontSize: 12,
                      bold: true,
                      alignment: 'center',
                      margin: [0, 0, 0, 8]
                    }},
                    {{
                      text: title + ' ' + (monthYear || metadata.monthYear),
                      alignment: 'center',
                      bold: true,
                      fontSize: 12,
                      margin: [0, 0, 0, 15]
                    }}
                  ];
                }}
                function createPDFDefinition() {{
                  const content = [];
                  content.push({{
                      stack: [
                        {{
                          image: logoImage,
                          width: 750,
                          alignment: 'center',
                          margin: [0, 20, 5 , 10]
                        }},
                        {{
                          text: 'मासिक आहवाल माहे : ' + metadata.monthYear,
                          alignment: 'center',
                          bold: true,
                          fontSize: 20  
                        }},
                        {{
                          text: 'प्राथमिक आरोग्य केंद्र ' + metadata.phcName,
                          alignment: 'center',
                          bold: true,
                          fontSize: 17 
                        }},
                        {{
                          text: 'उपकेंद्र '+ metadata.subcenter,
                          alignment: 'center',
                          bold: true,
                          fontSize: 17 
                        }}
                      ],
                      margin: [40, 10, 40, 10]
                    }});                  
                  content.push({{ text: '', pageBreak: 'before' }});
                  // SHEET 1 - रक्त नमुना (Keep as original)
                  if (allData.sheet1) {{
                    const s1 = allData.sheet1;
                    const tableBody1 = [];

                    content.push({{
                      text: s1.title || 'राष्ट्रीय कीटकजन्य रोग नियंत्रण कार्यक्रम, जिल्हा पुणे',
                      alignment: 'center',
                      bold: true,
                      fontSize: 13,
                      margin: [0, 30, 0, 8]
                    }});

                    content.push({{
                      text: 'प्रा. आ. केंद्र ' + metadata.phcName + ' तालुका ' + metadata.taluka + ' जिल्हा ' + metadata.district,
                      fontSize: 12,
                      bold: true,
                      alignment: 'center',
                      margin: [0, 0, 0, 8]
                    }});

                    content.push({{
                      text: (s1.subtitle || 'रक्त नमुना मासिक अहवाल') + ' ' + (s1.month_year || ''),
                      alignment: 'center',
                      bold: true,
                      fontSize: 12,
                      margin: [0, 0, 0, 5]
                    }});

                    content.push({{
                      text: 'उपकेंद्र: ' + (s1.subcenter_name || ''),
                      bold: true,
                      fontSize: 11,
                      margin: [10, 5, 0, 15]
                    }});

                    tableBody1.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'उपकेंद्राचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'लोकसंख्या', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'कर्मचाऱ्यांचे नाव वर्गवारी', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'रक्त नमुना वार्षिक उद्दिष्ट', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'मासिक घेतलेले रक्त नमुने', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}},
                      {{ text: 'प्रगतीपथावर घेतलेले रक्त नमुने जानेवारी २०२६ पासून', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}}
                    ]);

                    tableBody1.push([
                      {{}},{{}},{{}},{{}},{{}},
                      {{ text: 'पहिला पंधरावडा', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'दुसरा पंधरावडा', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'पहिला पंधरावडा', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'दुसरा पंधरावडा', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    const totalRows = (s1.staff_data ? s1.staff_data.length : 0) + 2;

                    if (s1.staff_data && s1.staff_data.length > 0) {{
                      s1.staff_data.forEach((staff, idx) => {{
                        const m1 = parseInt(staff.पहिला_पंधरावडा || 0);
                        const m2 = parseInt(staff.दुसरा_पंधरावडा || 0);
                        const p1 = parseInt(staff.प्रगती_पहिला || 0);
                        const p2 = parseInt(staff.प्रगती_दुसरा || 0);
                        const nameWithDesignation = (staff.नाव || '') + (staff.पदनाम ? ' (' + staff.पदनाम + ')' : '');

                        tableBody1.push([
                          {{ text: idx === 0 ? (s1.subcenter_sr || '1') : '', fontSize: 9, alignment: 'center', rowSpan: idx === 0 ? totalRows : 1 }},
                          {{ text: idx === 0 ? (s1.subcenter_name || '') : '', fontSize: 9, alignment: 'center', rowSpan: idx === 0 ? totalRows : 1 }},
                          {{ text: idx === 0 ? (s1.subcenter_pop || '') : '', fontSize: 9, alignment: 'center', rowSpan: idx === 0 ? totalRows : 1 }},
                          {{ text: nameWithDesignation, fontSize: 9, alignment: 'center' }},
                          {{ text: idx === 0 ? (s1.annual_target || '') : '', fontSize: 9, alignment: 'center', rowSpan: idx === 0 ? totalRows : 1 }},
                          {{ text: displayValue(m1), fontSize: 9, alignment: 'center' }},
                          {{ text: displayValue(m2), fontSize: 9, alignment: 'center' }},
                          {{ text: displayValue(m1 + m2), fontSize: 9, alignment: 'center' }},
                          {{ text: displayValue(p1), fontSize: 9, alignment: 'center' }},
                          {{ text: displayValue(p2), fontSize: 9, alignment: 'center' }},
                          {{ text: displayValue(p1 + p2), fontSize: 9, alignment: 'center' }}
                        ]);
                      }});
                    }}

                    const af1 = parseInt(s1.asha_data?.f1 || 0);
                    const af2 = parseInt(s1.asha_data?.f2 || 0);
                    const ap1 = parseInt(s1.asha_data?.p1 || 0);
                    const ap2 = parseInt(s1.asha_data?.p2 || 0);

                    tableBody1.push([
                      {{ text: '', alignment: 'center' }},
                      {{ text: '', alignment: 'center' }},
                      {{ text: '', alignment: 'center' }},
                      {{ text: 'एकूण आशा कार्यकर्ती\\nसंख्या ' + (s1.total_asha_count || ''), bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: '', alignment: 'center' }},
                      {{ text: displayValue(af1), fontSize: 9, alignment: 'center' }},
                      {{ text: displayValue(af2), fontSize: 9, alignment: 'center' }},
                      {{ text: displayValue(af1 + af2), fontSize: 9, alignment: 'center' }},
                      {{ text: displayValue(ap1), fontSize: 9, alignment: 'center' }},
                      {{ text: displayValue(ap2), fontSize: 9, alignment: 'center' }},
                      {{ text: displayValue(ap1 + ap2), fontSize: 9, alignment: 'center' }}
                    ]);

                    tableBody1.push([
                      {{ text: '', alignment: 'center' }},
                      {{ text: '', alignment: 'center' }},
                      {{ text: '', alignment: 'center' }},
                      {{ text: 'एकूण रक्तनमुने', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: '', alignment: 'center' }},
                      {{ text: displayValue(s1.totals?.f1), bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: displayValue(s1.totals?.f2), bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: displayValue(s1.totals?.monthly), bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: displayValue(s1.totals?.p1), bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: displayValue(s1.totals?.p2), bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: displayValue(s1.totals?.progress), bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    content.push({{
                      table: {{ 
                        widths: ['5%', '12%', '10%', '15%', '8%', '8%', '8%', '8%', '8%', '8%', '10%'], 
                        body: tableBody1 
                      }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}
                    }});
                  }}

                  // SHEET 2 - थुंकी संकलन - TWO TABLES
                  if (allData.sheet2) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s2 = allData.sheet2;

                    content.push(...addPageHeading(s2.title1 || 'थुंकी संकलन गावनिहाय अहवाल', metadata.monthYear));

                    const tableBody2a = [];
                    tableBody2a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'लोकसंख्या', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'मासिक', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}},
                      {{ text: 'वार्षिक', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}}
                    ]);
                    tableBody2a.push([
                      {{}},{{}},{{}},{{}},
                      {{ text: 'पुरुष', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'स्त्री', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'पुरुष', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'स्त्री', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    if (s2.table1) {{
                      s2.table1.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody2a.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: ['5%', '15%', '12%', '15%', '9%', '9%', '9%', '9%', '9%', '8%'], body: tableBody2a }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }},
                      margin: [0, 0, 0, 20]
                    }});

                    content.push(...addPageHeading2(s2.title2 || 'संशयीत क्षयरुग्ण तपासणी अहवाल', ''));

                    const tableBody2b = [];
                    tableBody2b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'संशयीत रुग्णाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'वय', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लिंग', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'नमुना घेतलेला दिनांक', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'तपासणी दिनांक', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लॅब क्रमांक', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    if (s2.table2) {{
                      s2.table2.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody2b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: ['5%', '12%', '15%', '7%', '7%', '12%', '12%', '12%', '12%'], body: tableBody2b }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}
                    }});
                  }}

                  // SHEET 3 - कुष्ठरुग्ण - TWO TABLES
                  if (allData.sheet3) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s3 = allData.sheet3;

                    content.push(...addPageHeading(s3.title1 || 'कुष्ठरुग्ण गावनिहाय मासिक अहवाल', metadata.monthYear));

                    const tableBody3a = [];
                    tableBody3a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'लोकसंख्या', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'संबंधित कुष्ठ रुग्ण', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}},
                      {{ text: 'अहवाल महिन्यात नवीन शोधलेले कुष्ठरुग्ण (एम.बी.)', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}},
                      {{ text: 'अहवाल महिन्यात नवीन शोधलेले कुष्ठरुग्ण (पी.बी.)', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}},
                      {{ text: 'नियमित औषधोपचार घेणारे कुष्ठरुग्ण', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}}
                    ]);
                    tableBody3a.push([
                      {{}},{{}},{{}},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    if (s3.table1) {{
                      s3.table1.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody3a.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: Array(15).fill('*'), body: tableBody3a }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }},
                      margin: [0, 0, 0, 20]
                    }});

                    content.push(...addPageHeading2(s3.title2 || 'कुष्ठरुग्ण तपशीलवार माहिती', ''));

                    const tableBody3b = [];
                    tableBody3b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'संबंधित रुग्णाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'वय', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लिंग', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लक्षणे', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    if (s3.table2) {{
                      s3.table2.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody3b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: ['8%', '15%', '18%', '8%', '8%', '25%', '18%'], body: tableBody3b }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}
                    }});
                  }}

                  // SHEET 4 - क्षय रुग्ण - TWO TABLES
                  if (allData.sheet4) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s4 = allData.sheet4;

                    content.push(...addPageHeading(s4.title1 || 'क्षयरुग्ण गावनिहाय मासिक अहवाल', metadata.monthYear));

                    const tableBody4a = [];
                    tableBody4a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'लोकसंख्या', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'मासिक', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}},
                      {{ text: 'वार्षिक', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}}
                    ]);
                    tableBody4a.push([
                      {{}},{{}},{{}},{{}},
                      {{ text: 'पुरुष', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'स्त्री', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'पुरुष', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'स्त्री', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    if (s4.table1) {{
                      s4.table1.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody4a.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: ['5%', '15%', '12%', '15%', '9%', '9%', '9%', '9%', '9%', '8%'], body: tableBody4a }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }},
                      margin: [0, 0, 0, 20]
                    }});

                    content.push(...addPageHeading2(s4.title2 || 'उपचार घेणारे क्षयरुग्ण तपशील', ''));

                    const tableBody4b = [];
                    tableBody4b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'क्षयरुग्णाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'वय', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लिंग', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'कॅटेगरी', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'औषधोपचार सुरू दिनांक', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'टी. बी. नंबर', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    if (s4.table2) {{
                      s4.table2.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody4b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: ['5%', '12%', '15%', '7%', '7%', '10%', '12%', '12%', '15%'], body: tableBody4b }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}
                    }});
                  }}

                  // SHEET 5 - कंटेनर सर्वेक्षण - ONE TABLE
                  if (allData.sheet5) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s5 = allData.sheet5;

                    content.push(...addPageHeading(s5.title || 'कंटेनर सर्वेक्षण मासिक अहवाल', metadata.monthYear));

                    const tableBody5 = [];
                    tableBody5.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लोकसंख्या', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण घरांची संख्या', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एडीएस डास अळी करता तपासलेले घरे', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एडीएस डास आळी करता दूषित आढळलेली घरे', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एडीएस डास अळी करिता तपासलेली भांडी', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एडीएस डास अळी करता दूषित आढळलेली भांडी', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'हाऊस इंडेक्स', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'कंटेनर इंडेक्स', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'ब्रँट्यू इंडेक्स', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'रिकामी केलेली भांडी', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'अँबेट टाकलेली भांडी', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    if (s5.data) {{
                      s5.data.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody5.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: Array(13).fill('*'), body: tableBody5 }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}
                    }});
                  }}

                  // SHEET 6 - डासउत्पत्ती - TWO TABLES
                  if (allData.sheet6) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s6 = allData.sheet6;

                    content.push(...addPageHeading(s6.title1 || 'डासउत्पत्ती स्थानांची गावनिहाय यादी', metadata.monthYear));

                    const tableBody6a = [];
                    tableBody6a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'उपकेंद्राचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'डास उत्पत्ती स्थाने', bold: true, fontSize: 9, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{ text: 'डासउत्पत्ती स्थानाचे ठिकाण', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }}
                    ]);
                    tableBody6a.push([
                      {{}},{{}},{{}},
                      {{ text: 'कायम', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'हंगामी', bold: true, fontSize: 9, alignment: 'center' }},
                      {{}}
                    ]);

                    if (s6.table1) {{
                      s6.table1.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody6a.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: ['8%', '18%', '18%', '12%', '12%', '32%'], body: tableBody6a }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }},
                      margin: [0, 0, 0, 20]
                    }});

                    content.push(...addPageHeading2(s6.title2 || 'गप्पी मासे पैदास केंद्राची यादी', ''));

                    const tableBody6b = [];
                    tableBody6b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'उपकेंद्राचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'गप्पी मासे पैदास केंद्र', bold: true, fontSize: 9, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{ text: 'गप्पी मासे पैदास केंद्राचे ठिकाण', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }}
                    ]);
                    tableBody6b.push([
                      {{}},{{}},{{}},
                      {{ text: 'कायम', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'हंगामी', bold: true, fontSize: 9, alignment: 'center' }},
                      {{}}
                    ]);

                    if (s6.table2) {{
                      s6.table2.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody6b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: ['8%', '18%', '18%', '12%', '12%', '32%'], body: tableBody6b }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}
                    }});
                  }}

                  // SHEET 7 - मोतीबिंदू - TWO TABLES
                  if (allData.sheet7) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s7 = allData.sheet7;

                    content.push(...addPageHeading(s7.title1 || 'मोतीबिंदू गावनिहाय मासिक अहवाल', metadata.monthYear));

                    const tableBody7a = [];
                    tableBody7a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'लोकसंख्या', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'संशयीत मोतीबिंदू', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}},
                      {{ text: 'अहवाल महिन्यात नवीन शोधलेले मोतीबिंदू', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},
                      {{}},{{}}
                    ]);
                    tableBody7a.push([
                      {{}},{{}},{{}},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    if (s7.table1) {{
                      s7.table1.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody7a.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: ['8%', '15%', '12%', '10%', '10%', '10%', '10%', '10%', '10%'], body: tableBody7a }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }},
                      margin: [0, 0, 0, 20]
                    }});

                    content.push(...addPageHeading2(s7.title2 || 'मोतीबिंदू रुग्ण तपशीलवार माहिती', ''));

                    const tableBody7b = [];
                    tableBody7b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'संबंधित रुग्णाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'वय', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लिंग', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लक्षणे', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    if (s7.table2) {{
                      s7.table2.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody7b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: ['8%', '15%', '18%', '8%', '8%', '25%', '18%'], body: tableBody7b }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}
                    }});
                  }}

                  // SHEET 8 - प्रयोगशाळा - TWO TABLES (COMPLEX)
                  if (allData.sheet8) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s8 = allData.sheet8;

                    content.push(...addPageHeading(s8.title1 || 'राज्य आयोग्य प्रयोगशाळा विविध नमुने तपासणी अहवाल', metadata.monthYear));

                    const tableBody8a = [];
                    // Complex nested header structure
                    tableBody8a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},
                      {{ text: 'उपकेंद्राचे नाव', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},
                      {{ text: 'एकूण सार्व. उद्भव', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},
                      {{ text: 'जैविक पाणी नमुने तपासणी', bold: true, fontSize: 8, colSpan: 4, alignment: 'center' }},
                      {{}},{{}},{{}},
                      {{ text: 'रासायनिक पाणी नमुने तपासणी', bold: true, fontSize: 8, colSpan: 4, alignment: 'center' }},
                      {{}},{{}},{{}},
                      {{ text: 'प्रारंभीची शिल्लक टीसीएल साठा (किग्रॅ)', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},
                      {{ text: 'चालू महिन्यात खरेदी केलेला साठा', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},
                      {{ text: 'चालू महिन्यात खर्च केलेला साठा', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},
                      {{ text: 'महिन्याच्या शेवटी शिल्लक साठा (किग्रॅ)', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},
                      {{ text: 'टी.सी.एल. नमुने', bold: true, fontSize: 8, colSpan: 4, alignment: 'center' }},
                      {{}},{{}},{{}},
                      {{ text: 'शौच नमुने', bold: true, fontSize: 8, colSpan: 4, alignment: 'center' }},
                      {{}},{{}},{{}},
                      {{ text: 'मीठ नमुने', bold: true, fontSize: 8, colSpan: 4, alignment: 'center' }},
                      {{}},{{}},{{}}
                    ]);

                    // Simplified for readability - showing structure
                    tableBody8a.push([
                      {{}},{{}},{{}},
                      {{ text: 'महिन्यात', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{ text: 'प्रगती पर', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{ text: 'महिन्यात', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{ text: 'प्रगती पर', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{}},{{}},{{}},{{}},
                      {{ text: 'महिन्यात', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{ text: 'प्रगती पर', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{ text: 'महिन्यात', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{ text: 'प्रगती पर', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{ text: 'महिन्यात', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},
                      {{}},
                      {{ text: 'प्रगती पर', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},
                      {{}}
                    ]);

                    tableBody8a.push([
                      {{}},{{}},{{}},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{}},{{}},{{}},{{}},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }}
                    ]);

                    if (s8.table1) {{
                      s8.table1.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          const rowData = [
                            {{ text: displayValue(row['अ. क्र.']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['उपकेंद्र']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['एकूण_सार्व']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['जै_मही_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['जै_मही_दू']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['जै_प्रगती_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['जै_प्रगती_दू']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['रा_मही_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['रा_मही_दू']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['रा_प्रगती_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['रा_प्रगती_दू']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['TCL_प्रारंभ']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['TCL_खरेदी']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['TCL_खर्च']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['TCL_शेवट']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['TCL_नमुना_मही_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['TCL_नमुना_मही_दू']), fontSize: 8, alignment: 'center' }},
                            {{ text: '', fontSize: 8, alignment: 'center' }},
                            {{ text: '', fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['शौच_मही_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['शौच_मही_दू']), fontSize: 8, alignment: 'center' }},
                            {{ text: '', fontSize: 8, alignment: 'center' }},
                            {{ text: '', fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['मीठ_मही_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['मीठ_मही_दू']), fontSize: 8, alignment: 'center' }},
                            {{ text: '', fontSize: 8, alignment: 'center' }},
                            {{ text: '', fontSize: 8, alignment: 'center' }}
                          ];
                          tableBody8a.push(rowData);
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: Array(27).fill('*'), body: tableBody8a }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 0.7, paddingRight: () => 0.7, paddingTop: () => 3, paddingBottom: () => 3 }},
                      margin: [0, 0, 0, 20],
                      fontSize: 7
                    }});

                    content.push(...addPageHeading2(s8.title2 || 'गावनिहाय TCL साठा अहवाल', ''));

                    const tableBody8b = [];
                    tableBody8b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'उपकेंद्राचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण ग्रामपंचायत', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण गावे', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'टीसीएल साठा', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'टीसीएल साठवण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'पाणी शुद्धीकरण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'टीसीएल नसलेले गावे', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);

                    if (s8.table2) {{
                      s8.table2.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          tableBody8b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }})));
                        }}
                      }});
                    }}

                    content.push({{
                      table: {{ widths: ['8%', '18%', '12%', '12%', '12%', '12%', '14%', '12%'], body: tableBody8b }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}
                    }});
                  }}

                  return {{
                    pageSize: 'A4',
                    pageOrientation: 'landscape',
                    pageMargins: [15, 15, 15, 15],
                    content: content,
                    defaultStyle: {{ font: 'MarathiFont', fontSize: 9 }}
                  }};
                }}

                function previewPDF() {{
                  try {{
                    showStatus('📄 PDF तयार करत आहे...', '#FFF9C4');
                    const docDef = createPDFDefinition();
                    pdfMake.createPdf(docDef).open();
                    showStatus('✅ PDF तयार झाली!', '#C8E6C9');
                  }} catch(e) {{
                    showStatus('❌ त्रुटी: ' + e.message, '#FFCDD2');
                    console.error('PDF Error:', e);
                  }}
                }}

                function downloadPDF() {{
                  try {{
                    showStatus('⬇️ PDF डाउनलोड करत आहे...', '#FFF9C4');
                    const docDef = createPDFDefinition();
                    const filename = 'Masik_Ahwal_' + metadata.monthYear.replace(/\\s+/g, '_') + '.pdf';
                    pdfMake.createPdf(docDef).download(filename);
                    showStatus('✅ PDF डाउनलोड झाली!', '#C8E6C9');
                  }} catch(e) {{
                    showStatus('❌ त्रुटी: ' + e.message, '#FFCDD2');
                    console.error('PDF Error:', e);
                  }}
                }}
              </script>
            </body>
            </html>
            """,
            height=250
        )


if __name__ == "__main__":
    mothly_final_report()