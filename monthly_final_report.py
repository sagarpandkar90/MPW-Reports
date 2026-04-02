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

    # ── LocalStorage bridge ───────────────────────────────────────────────────
    # This component reads ALL saved keys from localStorage and puts them into
    # a hidden Streamlit query-param-like mechanism via st.session_state.
    # We use a one-time JS injection on first load to restore saved values.

    STORAGE_KEY = "masik_ahwal_v1"

    # Inject auto-save + restore JS (runs once per page load)
    components.html(
        f"""
        <script>
        (function() {{
            const KEY = "{STORAGE_KEY}";

            // ── RESTORE: send saved data back to Streamlit via a hidden textarea ──
            function restore() {{
                try {{
                    const saved = localStorage.getItem(KEY);
                    if (saved) {{
                        // Write into a hidden element Streamlit can read via URL hash
                        const el = parent.document.getElementById('ls_data_bridge');
                        if (el) {{ el.value = saved; el.dispatchEvent(new Event('input', {{bubbles:true}})); }}
                    }}
                }} catch(e) {{ console.warn('LS restore error:', e); }}
            }}

            // ── AUTO-SAVE: watch for any input/change inside Streamlit iframe ──
            function saveAll() {{
                try {{
                    // Collect all inputs, textareas, selects inside the Streamlit app
                    const inputs = parent.document.querySelectorAll(
                        'input[data-testid="stTextInput-input"], ' +
                        'textarea[data-testid="stTextArea-Input"], ' +
                        'input[type="number"], ' +
                        '[data-testid="stSelectbox"] [aria-selected="true"]'
                    );
                    const data = {{}};
                    inputs.forEach(inp => {{
                        const label = inp.closest('[data-testid]')?.querySelector('label')?.innerText ||
                                      inp.placeholder || inp.id || inp.name;
                        const key = inp['data-widget-id'] || inp.id || label;
                        if (key) data[key] = inp.value;
                    }});

                    // Better approach: save Streamlit's own widget state from URL + sessionStorage
                    // Streamlit stores widget state in sessionStorage under specific keys
                    const stState = {{}};
                    for (let i = 0; i < sessionStorage.length; i++) {{
                        const k = sessionStorage.key(i);
                        if (k && (k.startsWith('stWidgetValue') || k.startsWith('_stcore'))) {{
                            try {{ stState[k] = sessionStorage.getItem(k); }} catch(e) {{}}
                        }}
                    }}
                    if (Object.keys(stState).length > 0) {{
                        localStorage.setItem(KEY + '_ss', JSON.stringify(stState));
                    }}
                }} catch(e) {{ console.warn('LS save error:', e); }}
            }}

            // Save every 10 seconds
            setInterval(saveAll, 10000);
            // Also save on any input change
            parent.document.addEventListener('change', saveAll, true);
            parent.document.addEventListener('blur', saveAll, true);
        }})();
        </script>
        """,
        height=0,
    )

    # ── Session-state based save/restore (works within a session) ─────────────
    # We store a JSON snapshot in session_state['manual_backup']
    # and also offer export/import as JSON file for true persistence.

    if 'manual_backup' not in st.session_state:
        st.session_state.manual_backup = {}

    # ── Top bar with Save/Restore controls ───────────────────────────────────
    with st.expander("💾 डेटा सुरक्षित करा / पुनर्संचयित करा", expanded=False):
        st.info(
            "⚠️ नेटवर्क बंद पडल्यास किंवा पेज रिफ्रेश झाल्यास डेटा जाऊ शकतो. "
            "खालील पर्यायांनी डेटा सुरक्षित करा."
        )
        col_a, col_b, col_c = st.columns(3)

        with col_a:
            if st.button("📥 JSON म्हणून डाउनलोड करा", key="btn_export"):
                if st.session_state.get('sheet_data'):
                    export_data = {
                        'metadata': {
                            'month_year': st.session_state.get('meta_month_year', ''),
                            'phc_name': st.session_state.get('meta_phc_name', ''),
                            'taluka': st.session_state.get('meta_taluka', ''),
                            'district': st.session_state.get('meta_district', ''),
                            'sub_center': st.session_state.get('meta_sub_center', ''),
                            'population': st.session_state.get('meta_population', ''),
                        },
                        'sheet_data': st.session_state.sheet_data
                    }
                    json_str = json.dumps(export_data, ensure_ascii=False, indent=2)
                    b64 = base64.b64encode(json_str.encode('utf-8')).decode()
                    st.markdown(
                        f'<a href="data:application/json;base64,{b64}" '
                        f'download="masik_ahwal_backup.json" '
                        f'style="background:#4CAF50;color:white;padding:8px 16px;border-radius:4px;text-decoration:none;">'
                        f'⬇️ डाउनलोड करा</a>',
                        unsafe_allow_html=True
                    )
                else:
                    st.warning("आधी डेटा भरा")

        with col_b:
            uploaded_backup = st.file_uploader("📤 JSON बॅकअप लोड करा", type=['json'], key="backup_uploader")
            if uploaded_backup is not None:
                try:
                    backup_data = json.loads(uploaded_backup.read().decode('utf-8'))
                    if 'sheet_data' in backup_data:
                        st.session_state.sheet_data = backup_data['sheet_data']
                        if 'metadata' in backup_data:
                            meta = backup_data['metadata']
                            st.session_state['meta_month_year'] = meta.get('month_year', '')
                            st.session_state['meta_phc_name'] = meta.get('phc_name', '')
                            st.session_state['meta_taluka'] = meta.get('taluka', '')
                            st.session_state['meta_district'] = meta.get('district', '')
                            st.session_state['meta_sub_center'] = meta.get('sub_center', '')
                            st.session_state['meta_population'] = meta.get('population', '')
                        st.success("✅ डेटा यशस्वीरित्या लोड झाला!")
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ लोड करताना त्रुटी: {e}")

        with col_c:
            st.markdown("**📋 ब्राउझर LocalStorage (Auto)**")
            components.html(
                f"""
                <div style="font-family:sans-serif;font-size:13px;">
                  <button onclick="saveNow()" style="padding:6px 14px;background:#2196F3;color:white;border:none;border-radius:4px;cursor:pointer;margin-bottom:6px;">
                    💾 आत्ता सेव्ह करा
                  </button>
                  &nbsp;
                  <button onclick="clearSaved()" style="padding:6px 14px;background:#f44336;color:white;border:none;border-radius:4px;cursor:pointer;margin-bottom:6px;">
                    🗑️ क्लियर करा
                  </button>
                  <div id="ls_status" style="margin-top:4px;font-size:12px;color:#555;"></div>
                </div>
                <script>
                const KEY = "{STORAGE_KEY}_inputs";
                function saveNow() {{
                    try {{
                        // Save all visible text inputs from the Streamlit app
                        const allInputs = {{}};
                        parent.document.querySelectorAll('input[aria-label], textarea[aria-label]').forEach(inp => {{
                            const label = inp.getAttribute('aria-label') || inp.id;
                            if (label) allInputs[label] = inp.value;
                        }});
                        // Also save number inputs
                        parent.document.querySelectorAll('input[type="number"]').forEach(inp => {{
                            const label = inp.closest('div[data-testid]')?.querySelector('label')?.innerText;
                            if (label) allInputs['num_' + label] = inp.value;
                        }});
                        localStorage.setItem(KEY, JSON.stringify(allInputs));
                        const ts = new Date().toLocaleTimeString('mr-IN');
                        document.getElementById('ls_status').innerText = '✅ ' + ts + ' ला सेव्ह झाले (' + Object.keys(allInputs).length + ' फील्ड)';
                    }} catch(e) {{ document.getElementById('ls_status').innerText = '❌ त्रुटी: ' + e.message; }}
                }}
                function clearSaved() {{
                    if (confirm('सर्व सेव्ह केलेला डेटा काढायचा?')) {{
                        localStorage.removeItem(KEY);
                        document.getElementById('ls_status').innerText = '🗑️ क्लियर झाले';
                    }}
                }}
                // Auto-save every 30 seconds
                setInterval(saveNow, 30000);
                // Show last saved info on load
                try {{
                    const saved = localStorage.getItem(KEY);
                    if (saved) {{
                        const parsed = JSON.parse(saved);
                        document.getElementById('ls_status').innerText = '📦 ' + Object.keys(parsed).length + ' फील्ड सेव्ह आहेत';
                    }}
                }} catch(e) {{}}
                </script>
                """,
                height=100,
            )

    # ── LocalStorage auto-fill on page load ──────────────────────────────────
    components.html(
        f"""
        <script>
        // On load, fill Streamlit inputs from localStorage
        (function() {{
            const KEY = "{STORAGE_KEY}_inputs";
            function fillInputs() {{
                try {{
                    const saved = localStorage.getItem(KEY);
                    if (!saved) return;
                    const data = JSON.parse(saved);
                    parent.document.querySelectorAll('input[aria-label], textarea[aria-label]').forEach(inp => {{
                        const label = inp.getAttribute('aria-label') || inp.id;
                        if (label && data[label] !== undefined && inp.value === '') {{
                            const nativeInputValueSetter = Object.getOwnPropertyDescriptor(
                                window.HTMLInputElement.prototype, 'value').set;
                            nativeInputValueSetter.call(inp, data[label]);
                            inp.dispatchEvent(new Event('input', {{bubbles: true}}));
                        }}
                    }});
                }} catch(e) {{ console.warn('Fill error:', e); }}
            }}
            // Try after a short delay to let Streamlit render
            setTimeout(fillInputs, 1500);
            setTimeout(fillInputs, 3000);
        }})();
        </script>
        """,
        height=0,
    )

    # Common metadata
    st.sidebar.header("मुख्य माहिती")
    month_year = st.sidebar.text_input("महिना/वर्ष", value=st.session_state.get('meta_month_year', 'मार्च २०२६'), key="meta_month_year")
    phc_name = st.sidebar.text_input("प्राथमिक आरोग्य केंद्र", value=st.session_state.get('meta_phc_name', 'शेळगाव'), key="meta_phc_name")
    taluka = st.sidebar.text_input("तालुका", value=st.session_state.get('meta_taluka', 'इंदापूर'), key="meta_taluka")
    district = st.sidebar.text_input("जिल्हा", value=st.session_state.get('meta_district', 'पुणे'), key="meta_district")
    sub_center = st.sidebar.text_input("उपकेंद्र", value=st.session_state.get('meta_sub_center', ''), key="meta_sub_center")
    population = st.sidebar.text_input("लोकसंख्या", value=st.session_state.get('meta_population', ''), key="meta_population")

    # ── Sidebar quick-save reminder ───────────────────────────────────────────
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 💾 डेटा सुरक्षा")
    st.sidebar.info("डेटा टाइप केल्यावर वरील **'💾 डेटा सुरक्षित करा'** section मधून JSON डाउनलोड करा.")
    if st.sidebar.button("📥 आत्ता JSON डाउनलोड करा", key="sidebar_export"):
        if st.session_state.get('sheet_data'):
            export_data = {
                'metadata': {
                    'month_year': month_year,
                    'phc_name': phc_name,
                    'taluka': taluka,
                    'district': district,
                    'sub_center': sub_center,
                    'population': population,
                },
                'sheet_data': st.session_state.sheet_data
            }
            json_str = json.dumps(export_data, ensure_ascii=False, indent=2)
            b64 = base64.b64encode(json_str.encode('utf-8')).decode()
            st.sidebar.markdown(
                f'<a href="data:application/json;base64,{b64}" '
                f'download="masik_ahwal_backup.json" '
                f'style="background:#4CAF50;color:white;padding:8px 16px;border-radius:4px;text-decoration:none;display:block;text-align:center;margin-top:8px;">'
                f'⬇️ डाउनलोड करा</a>',
                unsafe_allow_html=True
            )

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

    # Helper to get saved value from sheet_data backup
    def sv(sheet, *keys, default=""):
        """Get saved value from sheet_data if available"""
        try:
            d = st.session_state.sheet_data.get(sheet, {})
            for k in keys:
                d = d[k]
            return d if d is not None else default
        except (KeyError, TypeError, IndexError):
            return default

    def sv_row(sheet, table_key, row_idx, col_key, default=""):
        """Get saved value for a table row"""
        try:
            rows = st.session_state.sheet_data.get(sheet, {}).get(table_key, [])
            return rows[row_idx].get(col_key, default) if row_idx < len(rows) else default
        except (KeyError, TypeError, IndexError):
            return default

    # Sheet 1: रक्त नमुना मासिक अहवाल (NO CHANGES - Keep as is)
    with tabs[0]:
        st.subheader("रक्त नमुना मासिक अहवाल")

        # Basic subcenter info - entered once
        st.markdown("### उपकेंद्र माहिती")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            subcenter_sr = st.text_input("अ. क्र.", value=sv('sheet1', 'subcenter_sr', default="1"), key="s1_basic_sr")
        with col2:
            subcenter_name = st.text_input("उपकेंद्राचे नाव", value=sv('sheet1', 'subcenter_name'), key="s1_basic_name")
        with col3:
            subcenter_pop = st.text_input("लोकसंख्या", value=sv('sheet1', 'subcenter_pop'), key="s1_basic_pop")
        with col4:
            annual_target = st.text_input("रक्त नमुना वार्षिक उद्दिष्ट", value=sv('sheet1', 'annual_target'), key="s1_annual_target")

        st.markdown("---")
        st.markdown("### कर्मचारी माहिती (आरोग्य सेविका व सेवक)")

        saved_staff_count = len(st.session_state.sheet_data.get('sheet1', {}).get('staff_data', [])) or 3
        num_staff = st.number_input("कर्मचारी संख्या", min_value=1, max_value=20, value=saved_staff_count, key="rows1")

        staff_data = []
        designation_options = ["आरोग्य सेवक", "आरोग्य सेविका", "आरोग्य सेविका NHM"]

        for i in range(num_staff):
            st.markdown(f"**कर्मचारी {i + 1}**")
            cols = st.columns([2, 2, 1, 1, 1, 1])

            staff_name = cols[0].text_input(f"नाव", value=sv_row('sheet1', 'staff_data', i, 'नाव'), key=f"s1_name_{i}")
            saved_desg = sv_row('sheet1', 'staff_data', i, 'पदनाम')
            desg_idx = designation_options.index(saved_desg) if saved_desg in designation_options else 0
            staff_designation = cols[1].selectbox(f"पदनाम", options=designation_options, index=desg_idx, key=f"s1_desg_{i}")

            row_data = {
                'नाव': staff_name,
                'पदनाम': staff_designation,
                'पहिला_पंधरावडा': cols[2].text_input(f"मासिक पहिला", value=sv_row('sheet1', 'staff_data', i, 'पहिला_पंधरावडा', '0'), key=f"s1_f1_{i}"),
                'दुसरा_पंधरावडा': cols[3].text_input(f"मासिक दुसरा", value=sv_row('sheet1', 'staff_data', i, 'दुसरा_पंधरावडा', '0'), key=f"s1_f2_{i}"),
                'प्रगती_पहिला': cols[4].text_input(f"प्रगती पहिला", value=sv_row('sheet1', 'staff_data', i, 'प्रगती_पहिला', '0'), key=f"s1_p1_{i}"),
                'प्रगती_दुसरा': cols[5].text_input(f"प्रगती दुसरा", value=sv_row('sheet1', 'staff_data', i, 'प्रगती_दुसरा', '0'), key=f"s1_p2_{i}"),
            }
            staff_data.append(row_data)

        st.markdown("---")
        st.markdown("### एकूण आशा कार्यकर्ती")
        total_asha_count = st.text_input("एकूण आशा कार्यकर्ती संख्या", value=sv('sheet1', 'total_asha_count', default="0"), key="s1_total_asha")

        cols_asha = st.columns([2, 2, 1, 1, 1, 1])
        with cols_asha[0]:
            st.text("(मासिक व प्रगती)")
        asha_f1 = cols_asha[2].text_input("मासिक पहिला", value=sv('sheet1', 'asha_data', 'f1', default="0"), key="s1_asha_f1")
        asha_f2 = cols_asha[3].text_input("मासिक दुसरा", value=sv('sheet1', 'asha_data', 'f2', default="0"), key="s1_asha_f2")
        asha_p1 = cols_asha[4].text_input("प्रगती पहिला", value=sv('sheet1', 'asha_data', 'p1', default="0"), key="s1_asha_p1")
        asha_p2 = cols_asha[5].text_input("प्रगती दुसरा", value=sv('sheet1', 'asha_data', 'p2', default="0"), key="s1_asha_p2")

        total_f1 = sum(int(s['पहिला_पंधरावडा'] or 0) for s in staff_data) + int(asha_f1 or 0)
        total_f2 = sum(int(s['दुसरा_पंधरावडा'] or 0) for s in staff_data) + int(asha_f2 or 0)
        total_monthly = total_f1 + total_f2
        total_p1 = sum(int(s['प्रगती_पहिला'] or 0) for s in staff_data) + int(asha_p1 or 0)
        total_p2 = sum(int(s['प्रगती_दुसरा'] or 0) for s in staff_data) + int(asha_p2 or 0)
        total_progress = total_p1 + total_p2

        st.info(f"📊 एकूण रक्तनमुने - मासिक: पहिला={total_f1}, दुसरा={total_f2}, एकूण={total_monthly} | प्रगती: पहिला={total_p1}, दुसरा={total_p2}, एकूण={total_progress}")

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
            'asha_data': {'f1': asha_f1, 'f2': asha_f2, 'p1': asha_p1, 'p2': asha_p2},
            'totals': {
                'f1': str(total_f1), 'f2': str(total_f2), 'monthly': str(total_monthly),
                'p1': str(total_p1), 'p2': str(total_p2), 'progress': str(total_progress)
            }
        }

    # Sheet 2: थुंकी संकलन - TWO TABLES
    with tabs[1]:
        st.subheader("थुंकी संकलन अहवाल - 2 Tables")

        st.markdown("### तक्ता १: गावनिहाय थुंकी संकलन")
        saved_2a = len(st.session_state.sheet_data.get('sheet2', {}).get('table1', [])) or 4
        num_rows_2a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=saved_2a, key="rows2a")

        sheet2_table1 = []
        for i in range(num_rows_2a):
            cols = st.columns([1, 2, 2, 2, 1, 1, 1, 1, 1, 1])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=sv_row('sheet2','table1',i,'अ. क्र.',str(i+1)), key=f"s2a_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value=sv_row('sheet2','table1',i,'गावाचे नाव'), key=f"s2a_village_{i}"),
                'लोकसंख्या': cols[2].text_input(f"लोकसंख्या", value=sv_row('sheet2','table1',i,'लोकसंख्या'), key=f"s2a_pop_{i}"),
                'कर्मचारी': cols[3].text_input(f"कर्मचारी", value=sv_row('sheet2','table1',i,'कर्मचारी'), key=f"s2a_staff_{i}"),
                'मासिक_पुरुष': cols[4].text_input(f"मासिक पुरुष", value=sv_row('sheet2','table1',i,'मासिक_पुरुष'), key=f"s2a_m_m_{i}"),
                'मासिक_स्त्री': cols[5].text_input(f"मासिक स्त्री", value=sv_row('sheet2','table1',i,'मासिक_स्त्री'), key=f"s2a_m_f_{i}"),
                'मासिक_एकूण': cols[6].text_input(f"मासिक एकूण", value=sv_row('sheet2','table1',i,'मासिक_एकूण'), key=f"s2a_m_t_{i}"),
                'वार्षिक_पुरुष': cols[7].text_input(f"वार्षिक पुरुष", value=sv_row('sheet2','table1',i,'वार्षिक_पुरुष'), key=f"s2a_y_m_{i}"),
                'वार्षिक_स्त्री': cols[8].text_input(f"वार्षिक स्त्री", value=sv_row('sheet2','table1',i,'वार्षिक_स्त्री'), key=f"s2a_y_f_{i}"),
                'वार्षिक_एकूण': cols[9].text_input(f"वार्षिक एकूण", value=sv_row('sheet2','table1',i,'वार्षिक_एकूण'), key=f"s2a_y_t_{i}"),
            }
            sheet2_table1.append(row_data)

        st.markdown("---")
        st.markdown("### तक्ता २: संशयीत रुग्ण माहिती")
        saved_2b = len(st.session_state.sheet_data.get('sheet2', {}).get('table2', [])) or 4
        num_rows_2b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=saved_2b, key="rows2b")

        sheet2_table2 = []
        for i in range(num_rows_2b):
            cols = st.columns([1, 2, 2, 1, 1, 2, 2, 2, 2])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=sv_row('sheet2','table2',i,'अ. क्र.',str(i+1)), key=f"s2b_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value=sv_row('sheet2','table2',i,'गावाचे नाव'), key=f"s2b_village_{i}"),
                'संशयीत_रुग्ण': cols[2].text_input(f"रुग्णाचे नाव", value=sv_row('sheet2','table2',i,'संशयीत_रुग्ण'), key=f"s2b_name_{i}"),
                'वय': cols[3].text_input(f"वय", value=sv_row('sheet2','table2',i,'वय'), key=f"s2b_age_{i}"),
                'लिंग': cols[4].text_input(f"लिंग", value=sv_row('sheet2','table2',i,'लिंग'), key=f"s2b_gender_{i}"),
                'नमुना_दिनांक': cols[5].text_input(f"नमुना दिनांक", value=sv_row('sheet2','table2',i,'नमुना_दिनांक'), key=f"s2b_sample_{i}"),
                'तपासणी_दिनांक': cols[6].text_input(f"तपासणी दिनांक", value=sv_row('sheet2','table2',i,'तपासणी_दिनांक'), key=f"s2b_test_{i}"),
                'लॅब_क्रमांक': cols[7].text_input(f"लॅब क्रमांक", value=sv_row('sheet2','table2',i,'लॅब_क्रमांक'), key=f"s2b_lab_{i}"),
                'कर्मचारी': cols[8].text_input(f"कर्मचारी", value=sv_row('sheet2','table2',i,'कर्मचारी'), key=f"s2b_staff_{i}"),
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

        st.markdown("### तक्ता १: गावनिहाय कुष्ठरुग्ण")
        saved_3a = len(st.session_state.sheet_data.get('sheet3', {}).get('table1', [])) or 4
        num_rows_3a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=saved_3a, key="rows3a")

        sheet3_table1 = []
        for i in range(num_rows_3a):
            cols = st.columns([1, 2, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1])
            keys3a = ['अ. क्र.','गावाचे नाव','लोकसंख्या','संबंधित_मुले','संबंधित_प्रौढ','संबंधित_एकूण',
                      'MB_मुले','MB_प्रौढ','MB_एकूण','PB_मुले','PB_प्रौढ','PB_एकूण',
                      'औषधोपचार_मुले','औषधोपचार_प्रौढ','औषधोपचार_एकूण']
            widget_keys = [f"s3a_sr_{i}",f"s3a_village_{i}",f"s3a_pop_{i}",f"s3a_rel_c_{i}",f"s3a_rel_a_{i}",
                           f"s3a_rel_t_{i}",f"s3a_mb_c_{i}",f"s3a_mb_a_{i}",f"s3a_mb_t_{i}",f"s3a_pb_c_{i}",
                           f"s3a_pb_a_{i}",f"s3a_pb_t_{i}",f"s3a_tr_c_{i}",f"s3a_tr_a_{i}",f"s3a_tr_t_{i}"]
            labels = ["अ.क्र.","गाव","लोकसंख्या","सं.मुले","सं.प्रौढ","सं.एकूण","MB मुले","MB प्रौढ","MB एकूण",
                      "PB मुले","PB प्रौढ","PB एकूण","औ.मुले","औ.प्रौढ","औ.एकूण"]
            defaults = [str(i+1)] + [""]*14
            row_data = {}
            for j, (k, wk, lbl, dft) in enumerate(zip(keys3a, widget_keys, labels, defaults)):
                row_data[k] = cols[j].text_input(lbl, value=sv_row('sheet3','table1',i,k,dft), key=wk)
            sheet3_table1.append(row_data)

        st.markdown("---")
        st.markdown("### तक्ता २: कुष्ठरुग्ण तपशील")
        saved_3b = len(st.session_state.sheet_data.get('sheet3', {}).get('table2', [])) or 3
        num_rows_3b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=saved_3b, key="rows3b")

        sheet3_table2 = []
        for i in range(num_rows_3b):
            cols = st.columns([1, 2, 2, 1, 1, 3, 2])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=sv_row('sheet3','table2',i,'अ. क्र.',str(i+1)), key=f"s3b_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value=sv_row('sheet3','table2',i,'गावाचे नाव'), key=f"s3b_village_{i}"),
                'रुग्णाचे_नाव': cols[2].text_input(f"रुग्ण नाव", value=sv_row('sheet3','table2',i,'रुग्णाचे_नाव'), key=f"s3b_name_{i}"),
                'वय': cols[3].text_input(f"वय", value=sv_row('sheet3','table2',i,'वय'), key=f"s3b_age_{i}"),
                'लिंग': cols[4].text_input(f"लिंग", value=sv_row('sheet3','table2',i,'लिंग'), key=f"s3b_gender_{i}"),
                'लक्षणे': cols[5].text_input(f"लक्षणे", value=sv_row('sheet3','table2',i,'लक्षणे'), key=f"s3b_symptoms_{i}"),
                'कर्मचारी': cols[6].text_input(f"कर्मचारी", value=sv_row('sheet3','table2',i,'कर्मचारी'), key=f"s3b_staff_{i}"),
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

        st.markdown("### तक्ता १: गावनिहाय क्षय रुग्ण")
        saved_4a = len(st.session_state.sheet_data.get('sheet4', {}).get('table1', [])) or 3
        num_rows_4a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=saved_4a, key="rows4a")

        sheet4_table1 = []
        for i in range(num_rows_4a):
            cols = st.columns([1, 2, 2, 2, 1, 1, 1, 1, 1, 1])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=sv_row('sheet4','table1',i,'अ. क्र.',str(i+1)), key=f"s4a_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value=sv_row('sheet4','table1',i,'गावाचे नाव'), key=f"s4a_village_{i}"),
                'लोकसंख्या': cols[2].text_input(f"लोकसंख्या", value=sv_row('sheet4','table1',i,'लोकसंख्या'), key=f"s4a_pop_{i}"),
                'कर्मचारी': cols[3].text_input(f"कर्मचारी", value=sv_row('sheet4','table1',i,'कर्मचारी'), key=f"s4a_staff_{i}"),
                'मासिक_पुरुष': cols[4].text_input(f"मासिक पुरुष", value=sv_row('sheet4','table1',i,'मासिक_पुरुष'), key=f"s4a_m_m_{i}"),
                'मासिक_स्त्री': cols[5].text_input(f"मासिक स्त्री", value=sv_row('sheet4','table1',i,'मासिक_स्त्री'), key=f"s4a_m_f_{i}"),
                'मासिक_एकूण': cols[6].text_input(f"मासिक एकूण", value=sv_row('sheet4','table1',i,'मासिक_एकूण'), key=f"s4a_m_t_{i}"),
                'वार्षिक_पुरुष': cols[7].text_input(f"वार्षिक पुरुष", value=sv_row('sheet4','table1',i,'वार्षिक_पुरुष'), key=f"s4a_y_m_{i}"),
                'वार्षिक_स्त्री': cols[8].text_input(f"वार्षिक स्त्री", value=sv_row('sheet4','table1',i,'वार्षिक_स्त्री'), key=f"s4a_y_f_{i}"),
                'वार्षिक_एकूण': cols[9].text_input(f"वार्षिक एकूण", value=sv_row('sheet4','table1',i,'वार्षिक_एकूण'), key=f"s4a_y_t_{i}"),
            }
            sheet4_table1.append(row_data)

        st.markdown("---")
        st.markdown("### तक्ता २: क्षयरुग्ण तपशील")
        saved_4b = len(st.session_state.sheet_data.get('sheet4', {}).get('table2', [])) or 3
        num_rows_4b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=saved_4b, key="rows4b")

        sheet4_table2 = []
        for i in range(num_rows_4b):
            cols = st.columns([1, 2, 2, 1, 1, 2, 2, 2, 2])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=sv_row('sheet4','table2',i,'अ. क्र.',str(i+1)), key=f"s4b_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value=sv_row('sheet4','table2',i,'गावाचे नाव'), key=f"s4b_village_{i}"),
                'क्षयरुग्णाचे_नाव': cols[2].text_input(f"रुग्ण नाव", value=sv_row('sheet4','table2',i,'क्षयरुग्णाचे_नाव'), key=f"s4b_name_{i}"),
                'वय': cols[3].text_input(f"वय", value=sv_row('sheet4','table2',i,'वय'), key=f"s4b_age_{i}"),
                'लिंग': cols[4].text_input(f"लिंग", value=sv_row('sheet4','table2',i,'लिंग'), key=f"s4b_gender_{i}"),
                'कॅटेगरी': cols[5].text_input(f"कॅटेगरी", value=sv_row('sheet4','table2',i,'कॅटेगरी'), key=f"s4b_cat_{i}"),
                'औषधोपचार_दिनांक': cols[6].text_input(f"दिनांक", value=sv_row('sheet4','table2',i,'औषधोपचार_दिनांक'), key=f"s4b_date_{i}"),
                'TB_नंबर': cols[7].text_input(f"TB नं.", value=sv_row('sheet4','table2',i,'TB_नंबर'), key=f"s4b_tb_{i}"),
                'कर्मचारी': cols[8].text_input(f"कर्मचारी", value=sv_row('sheet4','table2',i,'कर्मचारी'), key=f"s4b_staff_{i}"),
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

        saved_5 = len(st.session_state.sheet_data.get('sheet5', {}).get('data', [])) or 3
        num_rows_5 = st.number_input("नोंदी संख्या", min_value=1, max_value=20, value=saved_5, key="rows5")

        sheet5_data = []
        keys5 = ['अ. क्र.','गावाचे नाव','लोकसंख्या','एकूण_घरे','तपासलेले_घरे','दूषित_घरे',
                 'तपासलेली_भांडी','दूषित_भांडी','House_Index','Container_Index','Breteau_Index','रिकामी_भांडी','अँबेट_भांडी']
        wkeys5 = [f"s5_sr",f"s5_village",f"s5_pop",f"s5_house",f"s5_check",f"s5_cont",
                  f"s5_cont_c",f"s5_cont_cc",f"s5_hi",f"s5_ci",f"s5_bi",f"s5_empty",f"s5_abate"]
        labels5 = ["अ.क्र.","गाव","लोकसंख्या","एकूण घरे","तपासले","दूषित","भांडी","दूषित भांडी","HI","CI","BI","रिकामी","अँबेट"]
        for i in range(num_rows_5):
            cols = st.columns([1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2])
            row_data = {}
            defaults5 = [str(i+1)] + [""]*12
            for j, (k, wk, lbl, dft) in enumerate(zip(keys5, wkeys5, labels5, defaults5)):
                row_data[k] = cols[j].text_input(lbl, value=sv_row('sheet5','data',i,k,dft), key=f"{wk}_{i}")
            sheet5_data.append(row_data)

        st.session_state.sheet_data['sheet5'] = {
            'title': 'कंटेनर सर्वेक्षण मासिक अहवाल',
            'data': sheet5_data
        }

    # Sheet 6: डासउत्पत्ती - TWO TABLES
    with tabs[5]:
        st.subheader("डासउत्पत्ती स्थाने - 2 Tables")

        st.markdown("### तक्ता १: डासउत्पत्ती स्थाने")
        saved_6a = len(st.session_state.sheet_data.get('sheet6', {}).get('table1', [])) or 4
        num_rows_6a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=saved_6a, key="rows6a")

        sheet6_table1 = []
        for i in range(num_rows_6a):
            cols = st.columns([1, 2, 2, 1, 1, 3])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=sv_row('sheet6','table1',i,'अ. क्र.',str(i+1)), key=f"s6a_sr_{i}"),
                'उपकेंद्राचे नाव': cols[1].text_input(f"उपकेंद्र", value=sv_row('sheet6','table1',i,'उपकेंद्राचे नाव'), key=f"s6a_sub_{i}"),
                'गावाचे नाव': cols[2].text_input(f"गाव", value=sv_row('sheet6','table1',i,'गावाचे नाव'), key=f"s6a_village_{i}"),
                'कायम': cols[3].text_input(f"कायम", value=sv_row('sheet6','table1',i,'कायम'), key=f"s6a_perm_{i}"),
                'हंगामी': cols[4].text_input(f"हंगामी", value=sv_row('sheet6','table1',i,'हंगामी'), key=f"s6a_seas_{i}"),
                'ठिकाण': cols[5].text_input(f"ठिकाण", value=sv_row('sheet6','table1',i,'ठिकाण'), key=f"s6a_loc_{i}"),
            }
            sheet6_table1.append(row_data)

        st.markdown("---")
        st.markdown("### तक्ता २: गप्पी मासे पैदास केंद्र")
        saved_6b = len(st.session_state.sheet_data.get('sheet6', {}).get('table2', [])) or 4
        num_rows_6b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=saved_6b, key="rows6b")

        sheet6_table2 = []
        for i in range(num_rows_6b):
            cols = st.columns([1, 2, 2, 1, 1, 3])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=sv_row('sheet6','table2',i,'अ. क्र.',str(i+1)), key=f"s6b_sr_{i}"),
                'उपकेंद्राचे नाव': cols[1].text_input(f"उपकेंद्र", value=sv_row('sheet6','table2',i,'उपकेंद्राचे नाव'), key=f"s6b_sub_{i}"),
                'गावाचे नाव': cols[2].text_input(f"गाव", value=sv_row('sheet6','table2',i,'गावाचे नाव'), key=f"s6b_village_{i}"),
                'कायम': cols[3].text_input(f"कायम", value=sv_row('sheet6','table2',i,'कायम'), key=f"s6b_perm_{i}"),
                'हंगामी': cols[4].text_input(f"हंगामी", value=sv_row('sheet6','table2',i,'हंगामी'), key=f"s6b_seas_{i}"),
                'ठिकाण': cols[5].text_input(f"ठिकाण", value=sv_row('sheet6','table2',i,'ठिकाण'), key=f"s6b_loc_{i}"),
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

        st.markdown("### तक्ता १: गावनिहाय मोतीबिंदू")
        saved_7a = len(st.session_state.sheet_data.get('sheet7', {}).get('table1', [])) or 4
        num_rows_7a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=saved_7a, key="rows7a")

        sheet7_table1 = []
        for i in range(num_rows_7a):
            cols = st.columns([1, 2, 2, 1, 1, 1, 1, 1, 1])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=sv_row('sheet7','table1',i,'अ. क्र.',str(i+1)), key=f"s7a_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value=sv_row('sheet7','table1',i,'गावाचे नाव'), key=f"s7a_village_{i}"),
                'लोकसंख्या': cols[2].text_input(f"लोकसंख्या", value=sv_row('sheet7','table1',i,'लोकसंख्या'), key=f"s7a_pop_{i}"),
                'संशयीत_मुले': cols[3].text_input(f"सं.मुले", value=sv_row('sheet7','table1',i,'संशयीत_मुले'), key=f"s7a_susp_c_{i}"),
                'संशयीत_प्रौढ': cols[4].text_input(f"सं.प्रौढ", value=sv_row('sheet7','table1',i,'संशयीत_प्रौढ'), key=f"s7a_susp_a_{i}"),
                'संशयीत_एकूण': cols[5].text_input(f"सं.एकूण", value=sv_row('sheet7','table1',i,'संशयीत_एकूण'), key=f"s7a_susp_t_{i}"),
                'नवीन_मुले': cols[6].text_input(f"नवीन मुले", value=sv_row('sheet7','table1',i,'नवीन_मुले'), key=f"s7a_new_c_{i}"),
                'नवीन_प्रौढ': cols[7].text_input(f"नवीन प्रौढ", value=sv_row('sheet7','table1',i,'नवीन_प्रौढ'), key=f"s7a_new_a_{i}"),
                'नवीन_एकूण': cols[8].text_input(f"नवीन एकूण", value=sv_row('sheet7','table1',i,'नवीन_एकूण'), key=f"s7a_new_t_{i}"),
            }
            sheet7_table1.append(row_data)

        st.markdown("---")
        st.markdown("### तक्ता २: मोतीबिंदू रुग्ण तपशील")
        saved_7b = len(st.session_state.sheet_data.get('sheet7', {}).get('table2', [])) or 4
        num_rows_7b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=saved_7b, key="rows7b")

        sheet7_table2 = []
        for i in range(num_rows_7b):
            cols = st.columns([1, 2, 2, 1, 1, 3, 2])
            row_data = {
                'अ. क्र.': cols[0].text_input(f"अ.क्र.", value=sv_row('sheet7','table2',i,'अ. क्र.',str(i+1)), key=f"s7b_sr_{i}"),
                'गावाचे नाव': cols[1].text_input(f"गाव", value=sv_row('sheet7','table2',i,'गावाचे नाव'), key=f"s7b_village_{i}"),
                'रुग्णाचे_नाव': cols[2].text_input(f"रुग्ण नाव", value=sv_row('sheet7','table2',i,'रुग्णाचे_नाव'), key=f"s7b_name_{i}"),
                'वय': cols[3].text_input(f"वय", value=sv_row('sheet7','table2',i,'वय'), key=f"s7b_age_{i}"),
                'लिंग': cols[4].text_input(f"लिंग", value=sv_row('sheet7','table2',i,'लिंग'), key=f"s7b_gender_{i}"),
                'लक्षणे': cols[5].text_input(f"लक्षणे", value=sv_row('sheet7','table2',i,'लक्षणे'), key=f"s7b_symptoms_{i}"),
                'कर्मचारी': cols[6].text_input(f"कर्मचारी", value=sv_row('sheet7','table2',i,'कर्मचारी'), key=f"s7b_staff_{i}"),
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
        st.warning("⚠️ Sheet 8 has complex nested columns. Simplified data entry is shown. Full structure will be in PDF.")

        st.markdown("### तक्ता १: प्रयोगशाळा तपासणी")
        saved_8a = len(st.session_state.sheet_data.get('sheet8', {}).get('table1', [])) or 3
        num_rows_8a = st.number_input("नोंदी संख्या (तक्ता १)", min_value=1, max_value=20, value=saved_8a, key="rows8a")

        sheet8_table1 = []
        for i in range(num_rows_8a):
            st.markdown(f"**नोंद {i + 1}**")
            cols = st.columns(3)
            sr = cols[0].text_input(f"अ.क्र.", value=sv_row('sheet8','table1',i,'अ. क्र.',str(i+1)), key=f"s8a_sr_{i}")
            subcenter = cols[1].text_input(f"उपकेंद्र", value=sv_row('sheet8','table1',i,'उपकेंद्र'), key=f"s8a_sub_{i}")
            total_surveys = cols[2].text_input(f"एकूण सार्व. उद्भव", value=sv_row('sheet8','table1',i,'एकूण_सार्व'), key=f"s8a_total_{i}")

            st.markdown("##### जैविक पाणी")
            cols2 = st.columns(4)
            bio_m_taken = cols2[0].text_input(f"महिना घेतले", value=sv_row('sheet8','table1',i,'जै_मही_घे'), key=f"s8a_bio_mt_{i}")
            bio_m_cont  = cols2[1].text_input(f"महिना दूषित", value=sv_row('sheet8','table1',i,'जै_मही_दू'), key=f"s8a_bio_mc_{i}")
            bio_y_taken = cols2[2].text_input(f"प्रगती घेतले", value=sv_row('sheet8','table1',i,'जै_प्रगती_घे'), key=f"s8a_bio_yt_{i}")
            bio_y_cont  = cols2[3].text_input(f"प्रगती दूषित", value=sv_row('sheet8','table1',i,'जै_प्रगती_दू'), key=f"s8a_bio_yc_{i}")

            st.markdown("##### रासायनिक पाणी")
            cols3 = st.columns(4)
            chem_m_taken = cols3[0].text_input(f"महिना घेतले", value=sv_row('sheet8','table1',i,'रा_मही_घे'), key=f"s8a_chem_mt_{i}")
            chem_m_cont  = cols3[1].text_input(f"महिना दूषित", value=sv_row('sheet8','table1',i,'रा_मही_दू'), key=f"s8a_chem_mc_{i}")
            chem_y_taken = cols3[2].text_input(f"प्रगती घेतले", value=sv_row('sheet8','table1',i,'रा_प्रगती_घे'), key=f"s8a_chem_yt_{i}")
            chem_y_cont  = cols3[3].text_input(f"प्रगती दूषित", value=sv_row('sheet8','table1',i,'रा_प्रगती_दू'), key=f"s8a_chem_yc_{i}")

            st.markdown("##### TCL साठा")
            cols4 = st.columns(4)
            tcl_start = cols4[0].text_input(f"प्रारंभ", value=sv_row('sheet8','table1',i,'TCL_प्रारंभ'), key=f"s8a_tcl_s_{i}")
            tcl_buy   = cols4[1].text_input(f"खरेदी",  value=sv_row('sheet8','table1',i,'TCL_खरेदी'), key=f"s8a_tcl_b_{i}")
            tcl_use   = cols4[2].text_input(f"खर्च",   value=sv_row('sheet8','table1',i,'TCL_खर्च'), key=f"s8a_tcl_u_{i}")
            tcl_end   = cols4[3].text_input(f"शेवट",   value=sv_row('sheet8','table1',i,'TCL_शेवट'), key=f"s8a_tcl_e_{i}")

            st.markdown("##### अतिरिक्त नमुने")
            cols5 = st.columns(6)
            tcl_sample_m_t = cols5[0].text_input(f"TCL मही घे", value=sv_row('sheet8','table1',i,'TCL_नमुना_मही_घे'), key=f"s8a_tcls_mt_{i}")
            tcl_sample_m_c = cols5[1].text_input(f"TCL मही दू", value=sv_row('sheet8','table1',i,'TCL_नमुना_मही_दू'), key=f"s8a_tcls_mc_{i}")
            stool_m_t      = cols5[2].text_input(f"शौच मही घे", value=sv_row('sheet8','table1',i,'शौच_मही_घे'), key=f"s8a_st_mt_{i}")
            stool_m_c      = cols5[3].text_input(f"शौच मही दू", value=sv_row('sheet8','table1',i,'शौच_मही_दू'), key=f"s8a_st_mc_{i}")
            salt_m_t       = cols5[4].text_input(f"मीठ मही घे", value=sv_row('sheet8','table1',i,'मीठ_मही_घे'), key=f"s8a_salt_mt_{i}")
            salt_m_c       = cols5[5].text_input(f"मीठ मही दू", value=sv_row('sheet8','table1',i,'मीठ_मही_दू'), key=f"s8a_salt_mc_{i}")

            row_data = {
                'अ. क्र.': sr, 'उपकेंद्र': subcenter, 'एकूण_सार्व': total_surveys,
                'जै_मही_घे': bio_m_taken, 'जै_मही_दू': bio_m_cont,
                'जै_प्रगती_घे': bio_y_taken, 'जै_प्रगती_दू': bio_y_cont,
                'रा_मही_घे': chem_m_taken, 'रा_मही_दू': chem_m_cont,
                'रा_प्रगती_घे': chem_y_taken, 'रा_प्रगती_दू': chem_y_cont,
                'TCL_प्रारंभ': tcl_start, 'TCL_खरेदी': tcl_buy,
                'TCL_खर्च': tcl_use, 'TCL_शेवट': tcl_end,
                'TCL_नमुना_मही_घे': tcl_sample_m_t, 'TCL_नमुना_मही_दू': tcl_sample_m_c,
                'शौच_मही_घे': stool_m_t, 'शौच_मही_दू': stool_m_c,
                'मीठ_मही_घे': salt_m_t, 'मीठ_मही_दू': salt_m_c,
            }
            sheet8_table1.append(row_data)
            st.markdown("---")

        st.markdown("### तक्ता २: गावनिहाय TCL साठा")
        saved_8b = len(st.session_state.sheet_data.get('sheet8', {}).get('table2', [])) or 4
        num_rows_8b = st.number_input("नोंदी संख्या (तक्ता २)", min_value=1, max_value=20, value=saved_8b, key="rows8b")

        sheet8_table2 = []
        for i in range(num_rows_8b):
            cols = st.columns([1, 2, 2, 2, 2, 2, 2, 2])
            row_data = {
                'अ. क्र.':    cols[0].text_input(f"अ.क्र.", value=sv_row('sheet8','table2',i,'अ. क्र.',str(i+1)), key=f"s8b_sr_{i}"),
                'उपकेंद्र':   cols[1].text_input(f"उपकेंद्र", value=sv_row('sheet8','table2',i,'उपकेंद्र'), key=f"s8b_sub_{i}"),
                'ग्रामपंचायत': cols[2].text_input(f"ग्रामपंचायत", value=sv_row('sheet8','table2',i,'ग्रामपंचायत'), key=f"s8b_gp_{i}"),
                'गावे':        cols[3].text_input(f"गावे", value=sv_row('sheet8','table2',i,'गावे'), key=f"s8b_villages_{i}"),
                'TCL_साठा':   cols[4].text_input(f"TCL साठा", value=sv_row('sheet8','table2',i,'TCL_साठा'), key=f"s8b_stock_{i}"),
                'TCL_साठवण':  cols[5].text_input(f"TCL साठवण", value=sv_row('sheet8','table2',i,'TCL_साठवण'), key=f"s8b_storage_{i}"),
                'पाणी_शुद्धी': cols[6].text_input(f"पाणी शुद्धी", value=sv_row('sheet8','table2',i,'पाणी_शुद्धी'), key=f"s8b_purif_{i}"),
                'TCL_नसलेले': cols[7].text_input(f"TCL नसलेले", value=sv_row('sheet8','table2',i,'TCL_नसलेले'), key=f"s8b_no_tcl_{i}"),
            }
            sheet8_table2.append(row_data)

        st.session_state.sheet_data['sheet8'] = {
            'title1': 'राज्य आयोग्य प्रयोगशाळा विविध नमुने तपासणी अहवाल',
            'title2': 'गावनिहाय TCL साठा अहवाल',
            'table1': sheet8_table1,
            'table2': sheet8_table2
        }

    # Tab 9: Generate PDF  — identical to original, zero changes
    with tabs[8]:
        st.subheader("📄 PDF तयार करा")
        st.info("सर्व डेटा भरल्यानंतर येथे PDF तयार करा")

        BASE_DIR = Path(__file__).resolve().parent
        font_path = BASE_DIR / "fonts" / "NotoSerifDevanagari-VariableFont_wdth,wght.ttf"
        imgpath = BASE_DIR / "fonts" / "img.png"
        if not font_path.exists():
            st.error(f"❌ Font file missing at: {font_path}")
            st.info("Please create a 'fonts' folder in the same directory as this script and add the Marathi font file.")
            st.markdown("""
            **Download font from:**
            - [Google Fonts - Noto Serif Devanagari](https://fonts.google.com/noto/specimen/Noto+Serif+Devanagari)
            - Or use any other Devanagari Unicode font (.ttf format)
            """)
            return

        font_b64 = base64.b64encode(font_path.read_bytes()).decode()
        img = base64.b64encode(imgpath.read_bytes()).decode()
        all_sheets_json = json.dumps(st.session_state.sheet_data, ensure_ascii=False)

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
                <button class="preview-btn" onclick="previewPDF()">👁️ Preview PDF</button>
                <button class="download-btn" onclick="downloadPDF()">⬇️ Download PDF</button>
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
                    normal: "MarathiFont.ttf", bold: "MarathiFont.ttf",
                    italics: "MarathiFont.ttf", bolditalics: "MarathiFont.ttf"
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
                  return [{{
                    text: title + ' ' + (monthYear || metadata.monthYear),
                    alignment: 'center', bold: true, fontSize: 12, margin: [0, 10, 0, 15]
                  }}];
                }}
                function addPageHeading(title, monthYear) {{
                  return [
                    {{ text: 'राष्ट्रीय कीटकजन्य रोग नियंत्रण कार्यक्रम, जिल्हा पुणे', alignment: 'center', bold: true, fontSize: 13, margin: [0, 20, 0, 4] }},
                    {{ text: 'प्रा. आ. केंद्र ' + metadata.phcName + ' तालुका ' + metadata.taluka + ' जिल्हा ' + metadata.district, fontSize: 12, bold: true, alignment: 'center', margin: [0, 0, 0, 8] }},
                    {{ text: title + ' ' + (monthYear || metadata.monthYear), alignment: 'center', bold: true, fontSize: 12, margin: [0, 0, 0, 15] }}
                  ];
                }}
                function createPDFDefinition() {{
                  const content = [];
                  content.push({{
                    stack: [
                      {{ image: logoImage, width: 750, alignment: 'center', margin: [0, 20, 5, 10] }},
                      {{ text: 'मासिक आहवाल माहे : ' + metadata.monthYear, alignment: 'center', bold: true, fontSize: 20 }},
                      {{ text: 'प्राथमिक आरोग्य केंद्र ' + metadata.phcName, alignment: 'center', bold: true, fontSize: 17 }},
                      {{ text: 'उपकेंद्र ' + metadata.subcenter, alignment: 'center', bold: true, fontSize: 17 }}
                    ],
                    margin: [40, 10, 40, 10]
                  }});
                  content.push({{ text: '', pageBreak: 'before' }});
                  if (allData.sheet1) {{
                    const s1 = allData.sheet1;
                    const tableBody1 = [];
                    content.push({{ text: s1.title || 'राष्ट्रीय कीटकजन्य रोग नियंत्रण कार्यक्रम, जिल्हा पुणे', alignment: 'center', bold: true, fontSize: 13, margin: [0, 30, 0, 8] }});
                    content.push({{ text: 'प्रा. आ. केंद्र ' + metadata.phcName + ' तालुका ' + metadata.taluka + ' जिल्हा ' + metadata.district, fontSize: 12, bold: true, alignment: 'center', margin: [0, 0, 0, 8] }});
                    content.push({{ text: (s1.subtitle || 'रक्त नमुना मासिक अहवाल') + ' ' + (s1.month_year || ''), alignment: 'center', bold: true, fontSize: 12, margin: [0, 0, 0, 5] }});
                    content.push({{ text: 'उपकेंद्र: ' + (s1.subcenter_name || ''), bold: true, fontSize: 11, margin: [10, 5, 0, 15] }});
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
                      {{ text: '', alignment: 'center' }},{{ text: '', alignment: 'center' }},{{ text: '', alignment: 'center' }},
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
                      {{ text: '', alignment: 'center' }},{{ text: '', alignment: 'center' }},{{ text: '', alignment: 'center' }},
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
                      table: {{ widths: ['5%', '12%', '10%', '15%', '8%', '8%', '8%', '8%', '8%', '8%', '10%'], body: tableBody1 }},
                      layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}
                    }});
                  }}
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
                    if (s2.table1) {{ s2.table1.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody2a.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: ['5%', '15%', '12%', '15%', '9%', '9%', '9%', '9%', '9%', '8%'], body: tableBody2a }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}, margin: [0, 0, 0, 20] }});
                    content.push(...addPageHeading2(s2.title2 || 'संशयीत क्षयरुग्ण तपासणी अहवाल', ''));
                    const tableBody2b = [];
                    tableBody2b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'गावाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'संशयीत रुग्णाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'वय', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लिंग', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'नमुना घेतलेला दिनांक', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'तपासणी दिनांक', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'लॅब क्रमांक', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);
                    if (s2.table2) {{ s2.table2.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody2b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: ['5%', '12%', '15%', '7%', '7%', '12%', '12%', '12%', '12%'], body: tableBody2b }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }} }});
                  }}
                  if (allData.sheet3) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s3 = allData.sheet3;
                    content.push(...addPageHeading(s3.title1 || 'कुष्ठरुग्ण गावनिहाय मासिक अहवाल', metadata.monthYear));
                    const tableBody3a = [];
                    tableBody3a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},{{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'लोकसंख्या', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'संबंधित कुष्ठ रुग्ण', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},{{}},{{}},
                      {{ text: 'अहवाल महिन्यात नवीन शोधलेले कुष्ठरुग्ण (एम.बी.)', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},{{}},{{}},
                      {{ text: 'अहवाल महिन्यात नवीन शोधलेले कुष्ठरुग्ण (पी.बी.)', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},{{}},{{}},
                      {{ text: 'नियमित औषधोपचार घेणारे कुष्ठरुग्ण', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},{{}},{{}}
                    ]);
                    tableBody3a.push([
                      {{}},{{}},{{}},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);
                    if (s3.table1) {{ s3.table1.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody3a.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: Array(15).fill('*'), body: tableBody3a }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}, margin: [0, 0, 0, 20] }});
                    content.push(...addPageHeading2(s3.title2 || 'कुष्ठरुग्ण तपशीलवार माहिती', ''));
                    const tableBody3b = [];
                    tableBody3b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'गावाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'संबंधित रुग्णाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'वय', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लिंग', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'लक्षणे', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);
                    if (s3.table2) {{ s3.table2.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody3b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: ['8%', '15%', '18%', '8%', '8%', '25%', '18%'], body: tableBody3b }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }} }});
                  }}
                  if (allData.sheet4) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s4 = allData.sheet4;
                    content.push(...addPageHeading(s4.title1 || 'क्षयरुग्ण गावनिहाय मासिक अहवाल', metadata.monthYear));
                    const tableBody4a = [];
                    tableBody4a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},{{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'लोकसंख्या', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},{{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'मासिक', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},{{}},{{}},
                      {{ text: 'वार्षिक', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},{{}},{{}}
                    ]);
                    tableBody4a.push([
                      {{}},{{}},{{}},{{}},
                      {{ text: 'पुरुष', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'स्त्री', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'पुरुष', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'स्त्री', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);
                    if (s4.table1) {{ s4.table1.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody4a.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: ['5%', '15%', '12%', '15%', '9%', '9%', '9%', '9%', '9%', '8%'], body: tableBody4a }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}, margin: [0, 0, 0, 20] }});
                    content.push(...addPageHeading2(s4.title2 || 'उपचार घेणारे क्षयरुग्ण तपशील', ''));
                    const tableBody4b = [];
                    tableBody4b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'गावाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'क्षयरुग्णाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'वय', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लिंग', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'कॅटेगरी', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'औषधोपचार सुरू दिनांक', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'टी. बी. नंबर', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);
                    if (s4.table2) {{ s4.table2.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody4b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: ['5%', '12%', '15%', '7%', '7%', '10%', '12%', '12%', '15%'], body: tableBody4b }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }} }});
                  }}
                  if (allData.sheet5) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s5 = allData.sheet5;
                    content.push(...addPageHeading(s5.title || 'कंटेनर सर्वेक्षण मासिक अहवाल', metadata.monthYear));
                    const tableBody5 = [];
                    tableBody5.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'गावाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लोकसंख्या', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एकूण घरांची संख्या', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एडीएस डास अळी करता तपासलेले घरे', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एडीएस डास आळी करता दूषित आढळलेली घरे', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एडीएस डास अळी करिता तपासलेली भांडी', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एडीएस डास अळी करता दूषित आढळलेली भांडी', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'हाऊस इंडेक्स', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'कंटेनर इंडेक्स', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'ब्रँट्यू इंडेक्स', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'रिकामी केलेली भांडी', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'अँबेट टाकलेली भांडी', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);
                    if (s5.data) {{ s5.data.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody5.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: Array(13).fill('*'), body: tableBody5 }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }} }});
                  }}
                  if (allData.sheet6) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s6 = allData.sheet6;
                    content.push(...addPageHeading(s6.title1 || 'डासउत्पत्ती स्थानांची गावनिहाय यादी', metadata.monthYear));
                    const tableBody6a = [];
                    tableBody6a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},{{ text: 'उपकेंद्राचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},{{ text: 'डास उत्पत्ती स्थाने', bold: true, fontSize: 9, colSpan: 2, alignment: 'center' }},{{}},
                      {{ text: 'डासउत्पत्ती स्थानाचे ठिकाण', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }}
                    ]);
                    tableBody6a.push([{{}},{{}},{{}},{{ text: 'कायम', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'हंगामी', bold: true, fontSize: 9, alignment: 'center' }},{{}}]);
                    if (s6.table1) {{ s6.table1.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody6a.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: ['8%', '18%', '18%', '12%', '12%', '32%'], body: tableBody6a }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}, margin: [0, 0, 0, 20] }});
                    content.push(...addPageHeading2(s6.title2 || 'गप्पी मासे पैदास केंद्राची यादी', ''));
                    const tableBody6b = [];
                    tableBody6b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},{{ text: 'उपकेंद्राचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},{{ text: 'गप्पी मासे पैदास केंद्र', bold: true, fontSize: 9, colSpan: 2, alignment: 'center' }},{{}},
                      {{ text: 'गप्पी मासे पैदास केंद्राचे ठिकाण', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }}
                    ]);
                    tableBody6b.push([{{}},{{}},{{}},{{ text: 'कायम', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'हंगामी', bold: true, fontSize: 9, alignment: 'center' }},{{}}]);
                    if (s6.table2) {{ s6.table2.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody6b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: ['8%', '18%', '18%', '12%', '12%', '32%'], body: tableBody6b }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }} }});
                  }}
                  if (allData.sheet7) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s7 = allData.sheet7;
                    content.push(...addPageHeading(s7.title1 || 'मोतीबिंदू गावनिहाय मासिक अहवाल', metadata.monthYear));
                    const tableBody7a = [];
                    tableBody7a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},{{ text: 'गावाचे नाव', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},
                      {{ text: 'लोकसंख्या', bold: true, fontSize: 9, rowSpan: 2, alignment: 'center' }},{{ text: 'संशयीत मोतीबिंदू', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},{{}},{{}},
                      {{ text: 'अहवाल महिन्यात नवीन शोधलेले मोतीबिंदू', bold: true, fontSize: 9, colSpan: 3, alignment: 'center' }},{{}},{{}}
                    ]);
                    tableBody7a.push([
                      {{}},{{}},{{}},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'मुले', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'प्रौढ', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एकूण', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);
                    if (s7.table1) {{ s7.table1.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody7a.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: ['8%', '15%', '12%', '10%', '10%', '10%', '10%', '10%', '10%'], body: tableBody7a }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }}, margin: [0, 0, 0, 20] }});
                    content.push(...addPageHeading2(s7.title2 || 'मोतीबिंदू रुग्ण तपशीलवार माहिती', ''));
                    const tableBody7b = [];
                    tableBody7b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'गावाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'संबंधित रुग्णाचे नाव', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'वय', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'लिंग', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'लक्षणे', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'कर्मचारी नाव', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);
                    if (s7.table2) {{ s7.table2.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody7b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: ['8%', '15%', '18%', '8%', '8%', '25%', '18%'], body: tableBody7b }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }} }});
                  }}
                  if (allData.sheet8) {{
                    content.push({{ text: '', pageBreak: 'before' }});
                    const s8 = allData.sheet8;
                    content.push(...addPageHeading(s8.title1 || 'राज्य आयोग्य प्रयोगशाळा विविध नमुने तपासणी अहवाल', metadata.monthYear));
                    const tableBody8a = [];
                    tableBody8a.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},{{ text: 'उपकेंद्राचे नाव', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},
                      {{ text: 'एकूण सार्व. उद्भव', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},{{ text: 'जैविक पाणी नमुने तपासणी', bold: true, fontSize: 8, colSpan: 4, alignment: 'center' }},{{}},{{}},{{}},
                      {{ text: 'रासायनिक पाणी नमुने तपासणी', bold: true, fontSize: 8, colSpan: 4, alignment: 'center' }},{{}},{{}},{{}},
                      {{ text: 'प्रारंभीची शिल्लक टीसीएल साठा (किग्रॅ)', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},{{ text: 'चालू महिन्यात खरेदी केलेला साठा', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},
                      {{ text: 'चालू महिन्यात खर्च केलेला साठा', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},{{ text: 'महिन्याच्या शेवटी शिल्लक साठा (किग्रॅ)', bold: true, fontSize: 8, rowSpan: 3, alignment: 'center' }},
                      {{ text: 'टी.सी.एल. नमुने', bold: true, fontSize: 8, colSpan: 4, alignment: 'center' }},{{}},{{}},{{}},
                      {{ text: 'शौच नमुने', bold: true, fontSize: 8, colSpan: 4, alignment: 'center' }},{{}},{{}},{{}},
                      {{ text: 'मीठ नमुने', bold: true, fontSize: 8, colSpan: 4, alignment: 'center' }},{{}},{{}},{{}}
                    ]);
                    tableBody8a.push([
                      {{}},{{}},{{}},
                      {{ text: 'महिन्यात', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},{{}},{{ text: 'प्रगती पर', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},{{}},
                      {{ text: 'महिन्यात', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},{{}},{{ text: 'प्रगती पर', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},{{}},
                      {{}},{{}},{{}},{{}},
                      {{ text: 'महिन्यात', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},{{}},{{ text: 'प्रगती पर', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},{{}},
                      {{ text: 'महिन्यात', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},{{}},{{ text: 'प्रगती पर', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},{{}},
                      {{ text: 'महिन्यात', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},{{}},{{ text: 'प्रगती पर', bold: true, fontSize: 8, colSpan: 2, alignment: 'center' }},{{}}
                    ]);
                    tableBody8a.push([
                      {{}},{{}},{{}},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{}},{{}},{{}},{{}},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},
                      {{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'घेतलेले', bold: true, fontSize: 8, alignment: 'center' }},{{ text: 'दुषित', bold: true, fontSize: 8, alignment: 'center' }}
                    ]);
                    if (s8.table1) {{
                      s8.table1.forEach(row => {{
                        if (Object.values(row).some(val => val !== '' && val !== '0')) {{
                          const rowData = [
                            {{ text: displayValue(row['अ. क्र.']), fontSize: 8, alignment: 'center' }},{{ text: displayValue(row['उपकेंद्र']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['एकूण_सार्व']), fontSize: 8, alignment: 'center' }},{{ text: displayValue(row['जै_मही_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['जै_मही_दू']), fontSize: 8, alignment: 'center' }},{{ text: displayValue(row['जै_प्रगती_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['जै_प्रगती_दू']), fontSize: 8, alignment: 'center' }},{{ text: displayValue(row['रा_मही_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['रा_मही_दू']), fontSize: 8, alignment: 'center' }},{{ text: displayValue(row['रा_प्रगती_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['रा_प्रगती_दू']), fontSize: 8, alignment: 'center' }},{{ text: displayValue(row['TCL_प्रारंभ']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['TCL_खरेदी']), fontSize: 8, alignment: 'center' }},{{ text: displayValue(row['TCL_खर्च']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['TCL_शेवट']), fontSize: 8, alignment: 'center' }},{{ text: displayValue(row['TCL_नमुना_मही_घे']), fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['TCL_नमुना_मही_दू']), fontSize: 8, alignment: 'center' }},{{ text: '', fontSize: 8, alignment: 'center' }},{{ text: '', fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['शौच_मही_घे']), fontSize: 8, alignment: 'center' }},{{ text: displayValue(row['शौच_मही_दू']), fontSize: 8, alignment: 'center' }},
                            {{ text: '', fontSize: 8, alignment: 'center' }},{{ text: '', fontSize: 8, alignment: 'center' }},
                            {{ text: displayValue(row['मीठ_मही_घे']), fontSize: 8, alignment: 'center' }},{{ text: displayValue(row['मीठ_मही_दू']), fontSize: 8, alignment: 'center' }},
                            {{ text: '', fontSize: 8, alignment: 'center' }},{{ text: '', fontSize: 8, alignment: 'center' }}
                          ];
                          tableBody8a.push(rowData);
                        }}
                      }});
                    }}
                    content.push({{ table: {{ widths: Array(27).fill('*'), body: tableBody8a }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 0.7, paddingRight: () => 0.7, paddingTop: () => 3, paddingBottom: () => 3 }}, margin: [0, 0, 0, 20], fontSize: 7 }});
                    content.push(...addPageHeading2(s8.title2 || 'गावनिहाय TCL साठा अहवाल', ''));
                    const tableBody8b = [];
                    tableBody8b.push([
                      {{ text: 'अ. क्र.', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'उपकेंद्राचे नाव', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'एकूण ग्रामपंचायत', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'एकूण गावे', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'टीसीएल साठा', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'टीसीएल साठवण', bold: true, fontSize: 9, alignment: 'center' }},
                      {{ text: 'पाणी शुद्धीकरण', bold: true, fontSize: 9, alignment: 'center' }},{{ text: 'टीसीएल नसलेले गावे', bold: true, fontSize: 9, alignment: 'center' }}
                    ]);
                    if (s8.table2) {{ s8.table2.forEach(row => {{ if (Object.values(row).some(val => val !== '' && val !== '0')) {{ tableBody8b.push(Object.values(row).map(val => ({{ text: displayValue(val), fontSize: 9, alignment: 'center' }}))); }} }}); }}
                    content.push({{ table: {{ widths: ['8%', '18%', '12%', '12%', '12%', '12%', '14%', '12%'], body: tableBody8b }}, layout: {{ hLineWidth: () => 0.5, vLineWidth: () => 0.5, paddingLeft: () => 4, paddingRight: () => 4, paddingTop: () => 4, paddingBottom: () => 4 }} }});
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
                  try {{ showStatus('📄 PDF तयार करत आहे...', '#FFF9C4'); pdfMake.createPdf(createPDFDefinition()).open(); showStatus('✅ PDF तयार झाली!', '#C8E6C9'); }}
                  catch(e) {{ showStatus('❌ त्रुटी: ' + e.message, '#FFCDD2'); console.error('PDF Error:', e); }}
                }}
                function downloadPDF() {{
                  try {{ showStatus('⬇️ PDF डाउनलोड करत आहे...', '#FFF9C4'); pdfMake.createPdf(createPDFDefinition()).download('Masik_Ahwal_' + metadata.monthYear.replace(/\\s+/g, '_') + '.pdf'); showStatus('✅ PDF डाउनलोड झाली!', '#C8E6C9'); }}
                  catch(e) {{ showStatus('❌ त्रुटी: ' + e.message, '#FFCDD2'); console.error('PDF Error:', e); }}
                }}
              </script>
            </body>
            </html>
            """,
            height=250
        )


if __name__ == "__main__":
    mothly_final_report()