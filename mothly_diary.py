def monthly_diary():
    import streamlit as st
    import pandas as pd
    import calendar
    import datetime
    from pathlib import Path
    import base64
    import json
    import streamlit.components.v1 as components
    from openpyxl import load_workbook

    st.title("मासिक डायरी – Arogya Sevak")

    BASE_DIR   = Path(__file__).resolve().parent
    excel_path = BASE_DIR / "movement.xlsx"

    if not excel_path.exists():
        st.error("❌ 'movement.xlsx' project folder मध्ये नाही!")
        return

    @st.cache_data
    def load_excel(path_str):
        wb   = load_workbook(path_str, read_only=True)
        ws   = wb.active
        rows = list(ws.iter_rows(values_only=True))
        header     = rows[0]
        month_cols = list(header[10:])
        mr_month_map = {
            'जाने': 1,  'फेब': 2,    'मार्च': 3,   'एप्रिल': 4,
            'मे':   5,  'जून': 6,    'जुलै': 7,    'ऑग': 8,
            'सेप्टें': 9, 'ऑक्टो': 10, 'नोव्हे': 11, 'डिसे': 12,
        }
        def parse_holiday_cell(val):
            if val is None:
                return set()
            result = set()
            for part in str(val).replace(" ", "").split(","):
                if part.isdigit():
                    result.add(int(part))
            return result

        data_rows   = []
        holiday_row = None
        for row in rows[1:]:
            gav   = str(row[0]).strip() if row[0] else ""
            vasti = str(row[1]).strip() if row[1] else ""
            if "सुट्टी" in gav or "सुट्टी" in vasti:
                holiday_row = list(row[10:])
            else:
                date_vals = list(row[10:])
                data_rows.append((gav, vasti, date_vals))

        if holiday_row:
            h_sets = [parse_holiday_cell(v) for v in holiday_row]
        else:
            h_sets = [set()] * 24

        half_meta = []
        for i, col_label in enumerate(month_cols):
            label     = str(col_label).strip()
            dash      = label.rfind('-')
            mr_name   = label[:dash].strip() if dash != -1 else label
            code      = label[dash+1:].strip() if dash != -1 else ""
            month_num = mr_month_map.get(mr_name, 0)
            part      = 1 if i % 2 == 0 else 2
            half_meta.append({
                "month_num": month_num,
                "part":      part,
                "code":      code,
                "label":     label,
                "col_idx":   i,
                "holidays":  h_sets[i] if i < len(h_sets) else set(),
            })
        return data_rows, half_meta

    data_rows, half_meta = load_excel(str(excel_path))

    EN_MONTHS = ["January","February","March","April","May","June",
                 "July","August","September","October","November","December"]
    MR_MONTHS = ["जानेवारी","फेब्रुवारी","मार्च","एप्रिल","मे","जून",
                 "जुलै","ऑगस्ट","सप्टेंबर","ऑक्टोबर","नोव्हेंबर","डिसेंबर"]

    # ── Common inputs ──────────────────────────────────────────
    cols = st.columns([1, 1, 2])
    month_en   = cols[0].selectbox("महिना", EN_MONTHS,
                    index=datetime.date.today().month - 1,
                    format_func=lambda x: MR_MONTHS[EN_MONTHS.index(x)])
    year       = cols[1].number_input("वर्ष", min_value=2000, max_value=2100,
                                       value=datetime.date.today().year)
    sevak_name = cols[2].text_input("आरोग्य सेवक नाव")

    c1, c2 = st.columns(2)
    upkendra   = c1.text_input("उपकेंद्राचे नाव")
    pra_kendra = c2.text_input("प्राथमिक आरोग्य केंद्र")

    month_num_sel = EN_MONTHS.index(month_en) + 1
    month_name_mr = MR_MONTHS[month_num_sel - 1]

    fixed_gav = data_rows[0][0] if data_rows else "शेळगाव"
    gav_input = st.text_input("गाव (बदलायचे असल्यास बदला)", value=fixed_gav)

    halves = [m for m in half_meta if m["month_num"] == month_num_sel]
    if len(halves) < 2:
        st.error("Excel मध्ये या महिन्याचा डेटा नाही!")
        return
    half1, half2 = halves[0], halves[1]

    def build_day_vasti_map(col_idx):
        mapping = {}
        for gav, vasti, date_vals in data_rows:
            if col_idx < len(date_vals):
                val = date_vals[col_idx]
                if val is not None:
                    try:
                        day = int(val)
                        if day > 0:
                            mapping[day] = vasti
                    except (ValueError, TypeError):
                        pass
        return mapping

    map1 = build_day_vasti_map(half1["col_idx"])
    map2 = build_day_vasti_map(half2["col_idx"])
    day_vasti_map    = {**map1, **map2}
    all_holiday_days = half1["holidays"] | half2["holidays"]

    FIXED_HOLIDAYS = {
        (1,  1):  "नवीन वर्ष",
        (1,  14): "मकर संक्रांति",
        (1,  26): "प्रजासत्ताक दिन",
        (2,  19): "छत्रपती शिवाजी महाराज जयंती",
        (4,  14): "डॉ. बाबासाहेब आंबेडकर जयंती",
        (5,  1):  "महाराष्ट्र दिन / कामगार दिन",
        (8,  15): "स्वातंत्र्य दिन",
        (10, 2):  "गांधी जयंती",
        (12, 25): "ख्रिसमस",
    }
    FLOATING_BY_MONTH = {
        3:  ["होळी", "धुळवड", "रमजान ईद (ईद-उल-फित्र)"],
        4:  ["गुड फ्रायडे"],
        5:  ["बुद्ध पौर्णिमा"],
        6:  ["ईद-उल-अधा (बकरी ईद)"],
        7:  ["रथयात्रा"],
        8:  ["पारसी नववर्ष (नवरोज)", "गणेश चतुर्थी"],
        9:  ["मोहरम"],
        10: ["दसरा (विजयादशमी)"],
        11: ["दिवाळी (लक्ष्मीपूजन)", "दिवाळी पाडवा", "गुरुनानक जयंती"],
    }

    def build_holiday_name_map(year_val):
        name_map = dict(FIXED_HOLIDAYS)
        for m, names in FLOATING_BY_MONTH.items():
            h1 = next((hm for hm in half_meta if hm["month_num"] == m and hm["part"] == 1), None)
            h2 = next((hm for hm in half_meta if hm["month_num"] == m and hm["part"] == 2), None)
            days_in_month = set()
            if h1: days_in_month |= h1["holidays"]
            if h2: days_in_month |= h2["holidays"]
            floating_days = sorted(d for d in days_in_month if (m, d) not in FIXED_HOLIDAYS)
            for i, day in enumerate(floating_days):
                name_map[(m, day)] = names[i] if i < len(names) else "सार्वजनिक सुट्टी"
        return name_map

    HOLIDAY_NAME_MAP = build_holiday_name_map(year)

    def get_holiday_name(m, d):
        return HOLIDAY_NAME_MAP.get((m, d), "सार्वजनिक सुट्टी")

    total_days = calendar.monthrange(year, month_num_sel)[1]
    dates      = [datetime.date(year, month_num_sel, d) for d in range(1, total_days + 1)]

    mondays  = [d for d in dates if d.weekday() == 0]
    tuesdays = [d for d in dates if d.weekday() == 1]

    first_monday  = mondays[0]  if len(mondays)  >= 1 else None
    first_tuesday = tuesdays[0] if len(tuesdays) >= 1 else None
    third_tuesday = tuesdays[2] if len(tuesdays) >= 3 else None

    # ── Font load ──────────────────────────────────────────────
    font_path = BASE_DIR / "fonts" / "NotoSerifDevanagari-VariableFont_wdth,wght.ttf"
    if not font_path.exists():
        st.error("❌ Font 'NotoSerifDevanagari-...' missing in fonts/ folder!")
        return
    font_b64 = base64.b64encode(font_path.read_bytes()).decode()

    # ══════════════════════════════════════════════════════════
    # TAB 1 — आगाऊ फिरती कार्यक्रम
    # ══════════════════════════════════════════════════════════
    tab1, tab2 = st.tabs(["📅 आगाऊ फिरती कार्यक्रम", "📒 मासिक दैनंदिनी"])

    with tab1:
        st.markdown("### 📅 मासिक आगाऊ फिरती कार्यक्रम")

        # Build DataFrame for Tab1
        rows1 = []
        for dt in dates:
            day        = dt.day
            is_sunday  = (dt.weekday() == 6)
            is_holiday = (day in all_holiday_days)

            if is_holiday:
                hname = get_holiday_name(month_num_sel, day)
                rows1.append({
                    "दिनांक":          dt.strftime("%d-%m-%Y"),
                    "कोठून":           "",
                    "कुठे":            "सा.सुट्टी",
                    "कामाचे स्वरूप":  hname,
                    "शेरा":            "",
                })
            elif is_sunday:
                rows1.append({
                    "दिनांक":          dt.strftime("%d-%m-%Y"),
                    "कोठून":           "",
                    "कुठे":            "---",
                    "कामाचे स्वरूप":  "रविवार",
                    "शेरा":            "",
                })
            elif dt == first_monday:
                vasti = day_vasti_map.get(day, "")
                rows1.append({
                    "दिनांक":          dt.strftime("%d-%m-%Y"),
                    "कोठून":           gav_input,
                    "कुठे":            vasti,
                    "कामाचे स्वरूप":  "नियमित लसीकरण भुजबळ वस्ती",
                    "शेरा":            "",
                })
            elif dt in (first_tuesday, third_tuesday):
                vasti = day_vasti_map.get(day, "")
                rows1.append({
                    "दिनांक":          dt.strftime("%d-%m-%Y"),
                    "कोठून":           gav_input,
                    "कुठे":            vasti,
                    "कामाचे स्वरूप":  "नियमित लसीकरण शेळगाव",
                    "शेरा":            "",
                })
            else:
                vasti = day_vasti_map.get(day, "")
                rows1.append({
                    "दिनांक":          dt.strftime("%d-%m-%Y"),
                    "कोठून":           gav_input,
                    "कुठे":            vasti,
                    "कामाचे स्वरूप":  f"घरभेट व कंटेनर सर्वेक्षण {vasti}" if vasti and vasti not in ("---","") else "बहुउद्देशीय कार्य प्रा. आ. केंद्र शेळगाव",
                    "शेरा":            "",
                })

        df1 = pd.DataFrame(rows1)
        edited_df1 = st.data_editor(
            df1,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "दिनांक":          st.column_config.TextColumn(width="small",  disabled=True),
                "कोठून":           st.column_config.TextColumn(width="small"),
                "कुठे":            st.column_config.TextColumn(width="medium"),
                "कामाचे स्वरूप":  st.column_config.TextColumn(width="large"),
                "शेरा":            st.column_config.TextColumn(width="medium"),
            },
            key="df1_editor",
        )

        st.info(
            f"📋 **Excel कोड:** `{half1['label']}` | `{half2['label']}`  \n"
            f"🎉 **सुट्ट्या (पहिला भाग):** {sorted(half1['holidays']) or 'नाही'}  \n"
            f"🎉 **सुट्ट्या (दुसरा भाग):** {sorted(half2['holidays']) or 'नाही'}"
        )

        # ── PDF 1 ──────────────────────────────────────────────
        data1_json = edited_df1.to_dict(orient="records")
        json1_js   = json.dumps(data1_json, ensure_ascii=False)

        components.html(
            f"""
            <html><head><meta charset="utf-8"/>
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.72/pdfmake.min.js"></script>
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.72/vfs_fonts.js"></script>
            </head>
            <body style="font-family:sans-serif;padding:10px;">
              <button onclick="previewPDF1()"
                style="padding:9px 18px;background:#2196F3;color:#fff;border:none;border-radius:6px;margin-right:10px;font-size:14px;cursor:pointer;">
                👁 Preview
              </button>
              <button onclick="downloadPDF1()"
                style="padding:9px 18px;background:#4CAF50;color:#fff;border:none;border-radius:6px;font-size:14px;cursor:pointer;">
                ⬇ Download PDF
              </button>
              <script>
                const data1 = {json1_js};
                pdfMake.vfs["Marathi.ttf"] = "{font_b64}";
                pdfMake.fonts = {{
                  MarathiFont: {{ normal:"Marathi.ttf", bold:"Marathi.ttf",
                                  italics:"Marathi.ttf", bolditalics:"Marathi.ttf" }}
                }};

                const HDR = "#BBDEFB", SUTTI = "#FFF9C4", RAVI = "#FCE4EC";

                function buildDoc1() {{
                  const body = [[
                    {{ text:"दिनांक",       bold:true, alignment:"center", fillColor:HDR }},
                    {{ text:"कोठून",        bold:true, alignment:"center", fillColor:HDR }},
                    {{ text:"कोठे",  bold:true, alignment:"center", fillColor:HDR }},
                    {{ text:"करावयाच्या कामाचा तपशील",         bold:true, alignment:"center", fillColor:HDR }},
                    {{ text:"शेरा",         bold:true, alignment:"center", fillColor:HDR }},
                  ]];

                  data1.forEach(r => {{
                    const vasti   = r["कुठे"]           || "";
                    const kothe   = r["कामाचे स्वरूप"] || "";
                    const isSutti = vasti === "सा.सुट्टी";
                    const isRavi  = kothe === "रविवार";
                    const fc      = isSutti ? SUTTI : isRavi ? RAVI : null;
                    body.push([
                      {{ text: r["दिनांक"]       || "", alignment:"center", fillColor:fc }},
                      {{ text: r["कोठून"]        || "", alignment:"center", fillColor:fc }},
                      {{ text: vasti,                   alignment:"center", fillColor:fc }},
                      {{ text: kothe,                   alignment:"center",   fillColor:fc }},
                      {{ text: r["शेरा"]         || "", alignment:"center", fillColor:fc }},
                    ]);
                  }});

                  // Signature row below table
                  const sigRow = {{
                    columns: [
                      {{ text: "\\n\\n   आरोग्य सेवक: {sevak_name}\\n    उपकेंद्र: {upkendra}", fontSize:10, width:"40%", alignment:"center" }},
                      {{ text:"\\n\\n", width:"20%"}},
                      {{ text: "\\n\\nमा. वैद्यकीय अधिकारी\\nप्रा. आ. केंद्र: {pra_kendra}", fontSize:10, width:"40%", alignment:"center" }},
                    ],
                    margin: [0, 14, 0, 0]
                  }};

                  return {{
                    pageSize: "A4",
                    pageMargins: [30, 20, 30, 30],
                    defaultStyle: {{ font:"MarathiFont", fontSize:10 }},
                    content: [
                      {{ text:"प्रा. आ. केंद्र {pra_kendra}                       उपकेंद्र {upkendra}", fontSize:13, alignment:"center", margin:[0,0,0,2] }},
                      {{ text:"मासिक आगाऊ फिरती कार्यक्रम", fontSize:13, bold:true, alignment:"center", margin:[0,0,0,2] }},
                      {{ text:"कर्मचारी नाव: {sevak_name}                                                                                    {month_name_mr} {year}", alignment:"left", fontSize:11, bold:true, margin:[0,0,0,6] }},

                      {{
                        table: {{ widths:["13%","13%","17%","40%","17%"], body }},
                        layout: {{ hLineWidth:()=>0.6, vLineWidth:()=>0.6,
                                   paddingTop:()=>2, paddingBottom:()=>2,
                                   paddingLeft:()=>3, paddingRight:()=>3 }}
                      }},
                      sigRow,
                    ],
                    styles: {{ title: {{ fontSize:15, bold:true }} }}
                  }};
                }}

                function previewPDF1()  {{ pdfMake.createPdf(buildDoc1()).open(); }}
                function downloadPDF1() {{ pdfMake.createPdf(buildDoc1()).download("Agau_Firti_{month_name_mr}_{year}.pdf"); }}
              </script>
            </body></html>
            """,
            height=80,
        )

    # ══════════════════════════════════════════════════════════
    # TAB 2 — मासिक दैनंदिनी
    # ══════════════════════════════════════════════════════════
    with tab2:
        st.markdown("### 📒 मासिक दैनंदिनी")

        rows2 = []
        work_days    = 0
        firti_days   = 0
        sutti_days   = 0
        raje_days    = 0   # रजा दिवस — user fills शेरा

        for dt in dates:
            day        = dt.day
            is_sunday  = (dt.weekday() == 6)
            is_holiday = (day in all_holiday_days)

            if is_holiday:
                hname = get_holiday_name(month_num_sel, day)
                sutti_days += 1
                rows2.append({
                    "दिनांक":          dt.strftime("%d-%m-%Y"),
                    "कोठून":           "",
                    "कुठे":            "सा.सुट्टी",
                    "कामाचे स्वरूप":  hname,
                    "निघण्याची वेळ":   "",
                    "परतीची वेळ":      "",
                    "शेरा":            "",
                })
            elif is_sunday:
                sutti_days += 1
                rows2.append({
                    "दिनांक":          dt.strftime("%d-%m-%Y"),
                    "कोठून":           "",
                    "कुठे":            "---",
                    "कामाचे स्वरूप":  "रविवार",
                    "निघण्याची वेळ":   "",
                    "परतीची वेळ":      "",
                    "शेरा":            "",
                })
            elif dt == first_monday:
                vasti = day_vasti_map.get(day, "")
                work_days  += 1
                firti_days += 1
                rows2.append({
                    "दिनांक":          dt.strftime("%d-%m-%Y"),
                    "कोठून":           gav_input,
                    "कुठे":            vasti,
                    "कामाचे स्वरूप":  "नियमित लसीकरण भुजबळ वस्ती",
                    "निघण्याची वेळ":   "9:00",
                    "परतीची वेळ":      "5:00",
                    "शेरा":            "",
                })
            elif dt in (first_tuesday, third_tuesday):
                vasti = day_vasti_map.get(day, "")
                work_days  += 1
                firti_days += 1
                rows2.append({
                    "दिनांक":          dt.strftime("%d-%m-%Y"),
                    "कोठून":           gav_input,
                    "कुठे":            vasti,
                    "कामाचे स्वरूप":  "नियमित लसीकरण शेळगाव",
                    "निघण्याची वेळ":   "9:00",
                    "परतीची वेळ":      "5:00",
                    "शेरा":            "",
                })
            else:
                vasti = day_vasti_map.get(day, "")
                work_days += 1
                if vasti and vasti not in ("---", ""):
                    firti_days += 1
                    kothe = f"घरभेट व कंटेनर सर्वेक्षण {vasti}"
                else:
                    kothe = "बहुउद्देशीय कार्य प्रा. आ. केंद्र शेळगाव"
                rows2.append({
                    "दिनांक":          dt.strftime("%d-%m-%Y"),
                    "कोठून":           gav_input,
                    "कुठे":            vasti,
                    "कामाचे स्वरूप":  kothe,
                    "निघण्याची वेळ":   "9:00",
                    "परतीची वेळ":      "5:00",
                    "शेरा":            "",
                })

        df2 = pd.DataFrame(rows2)
        edited_df2 = st.data_editor(
            df2,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "दिनांक":          st.column_config.TextColumn(width="small",  disabled=True),
                "कोठून":           st.column_config.TextColumn(width="small"),
                "कुठे":            st.column_config.TextColumn(width="medium"),
                "कामाचे स्वरूप":  st.column_config.TextColumn(width="large"),
                "निघण्याची वेळ":   st.column_config.TextColumn(width="small"),
                "परतीची वेळ":      st.column_config.TextColumn(width="small"),
                "शेरा":            st.column_config.TextColumn(width="medium"),
            },
            key="df2_editor",
        )

        # ── Summary stats below dataframe ──────────────────────
        st.markdown("---")
        st.markdown("#### 📊 महिन्याचा सारांश")
        sc1, sc2, sc3, sc4 = st.columns(4)
        ek_kam   = sc1.number_input("एकूण कामाचे दिवस",   value=work_days,  min_value=0, key="ek_kam")
        ek_firti = sc2.number_input("एकूण फिरतीचे दिवस",  value=firti_days, min_value=0, key="ek_firti")
        ek_sutti = sc3.number_input("एकूण सुट्टीचे दिवस", value=sutti_days, min_value=0, key="ek_sutti")
        ek_raje  = sc4.number_input("एकूण रजेचे दिवस",    value=raje_days,  min_value=0, key="ek_raje")

        # ── PDF 2 ──────────────────────────────────────────────
        data2_json = edited_df2.to_dict(orient="records")
        json2_js   = json.dumps(data2_json, ensure_ascii=False)

        components.html(
            f"""
            <html><head><meta charset="utf-8"/>
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.72/pdfmake.min.js"></script>
              <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.72/vfs_fonts.js"></script>
            </head>
            <body style="font-family:sans-serif;padding:10px;">
              <button onclick="previewPDF2()"
                style="padding:9px 18px;background:#2196F3;color:#fff;border:none;border-radius:6px;margin-right:10px;font-size:14px;cursor:pointer;">
                👁 Preview
              </button>
              <button onclick="downloadPDF2()"
                style="padding:9px 18px;background:#4CAF50;color:#fff;border:none;border-radius:6px;font-size:14px;cursor:pointer;">
                ⬇ Download PDF
              </button>
              <script>
                const data2 = {json2_js};
                pdfMake.vfs["Marathi.ttf"] = "{font_b64}";
                pdfMake.fonts = {{
                  MarathiFont: {{ normal:"Marathi.ttf", bold:"Marathi.ttf",
                                  italics:"Marathi.ttf", bolditalics:"Marathi.ttf" }}
                }};

                const HDR2 = "#BBDEFB", SUTTI2 = "#FFF9C4", RAVI2 = "#FCE4EC";

                function buildDoc2() {{
                  // ── Main table ────────────────────────────────
                  const body2 = [[
                    {{ text:"दिनांक",          bold:true, alignment:"center", fillColor:HDR2 }},
                    {{ text:"कोठून",           bold:true, alignment:"center", fillColor:HDR2 }},
                    {{ text:"कोठे",     bold:true, alignment:"center", fillColor:HDR2 }},
                    {{ text:"निघण्याची वेळ",   bold:true, alignment:"center", fillColor:HDR2 }},
                    {{ text:"परतीची वेळ",      bold:true, alignment:"center", fillColor:HDR2 }},
                    {{ text:"केलेल्या कामाचा तपशील",            bold:true, alignment:"center", fillColor:HDR2 }},
                    {{ text:"शेरा",            bold:true, alignment:"center", fillColor:HDR2 }},
                  ]];

                  data2.forEach(r => {{
                    const vasti   = r["कुठे"]           || "";
                    const kothe   = r["कामाचे स्वरूप"] || "";
                    const isSutti = vasti === "सा.सुट्टी";
                    const isRavi  = kothe === "रविवार";
                    const fc      = isSutti ? SUTTI2 : isRavi ? RAVI2 : null;
                    body2.push([
                      {{ text: r["दिनांक"]          || "", alignment:"center", fillColor:fc }},
                      {{ text: r["कोठून"]           || "", alignment:"center", fillColor:fc }},
                      {{ text: vasti,                      alignment:"center", fillColor:fc }},
                      {{ text: r["निघण्याची वेळ"]   || "", alignment:"center", fillColor:fc }},
                      {{ text: r["परतीची वेळ"]      || "", alignment:"center", fillColor:fc }},
                      {{ text: kothe,                      alignment:"center",   fillColor:fc }},
                      {{ text: r["शेरा"]            || "", alignment:"center", fillColor:fc }},
                    ]);
                  }});

                  // ── Summary stats small table ─────────────────
                  const statsBody = [[
                    {{ text:"एकूण कामाचे दिवस",   bold:true, alignment:"center", fillColor:"#E3F2FD" }},
                    {{ text:"एकूण फिरतीचे दिवस",  bold:true, alignment:"center", fillColor:"#E3F2FD" }},
                    {{ text:"एकूण सुट्टीचे दिवस", bold:true, alignment:"center", fillColor:"#E3F2FD" }},
                    {{ text:"एकूण रजेचे दिवस",    bold:true, alignment:"center", fillColor:"#E3F2FD" }},
                  ],[
                    {{ text:"{ek_kam}",   alignment:"center", fontSize:9 , bold:true }},
                    {{ text:"{ek_firti}", alignment:"center", fontSize:9 , bold:true }},
                    {{ text:"{ek_sutti}", alignment:"center", fontSize:9 , bold:true }},
                    {{ text:"{ek_raje}",  alignment:"center", fontSize:9 , bold:true }},
                  ]];

                  // ── Signature table (3 persons) ───────────────
                  const sigBody = [[
                    {{ text:"आरोग्य सेवक\\n{sevak_name}\\nउपकेंद्र: {upkendra}",
                       alignment:"center", fontSize:9 }},
                    {{ text:"मा. आरोग्य सहाय्यक\\nप्रा. आ. केंद्र: {pra_kendra}",
                       alignment:"center", fontSize:9 }},
                    {{ text:"मा. वैद्यकीय अधिकारी\\nप्रा. आ. केंद्र: {pra_kendra}",
                       alignment:"center", fontSize:9 }},
                  ]];

                  return {{
                    pageSize: "A4",
                    pageMargins: [25, 20, 25, 5 ],
                    defaultStyle: {{ font:"MarathiFont", fontSize:9 }},
                    content: [
                      {{ text:"प्रा. आ. केंद्र {pra_kendra}                       उपकेंद्र {upkendra}", style:"title", alignment:"center", margin:[0,0,0,2] }},
                      {{ text:"मासिक दैनंदिनी", fontSize:14, bold:true, alignment:"center", margin:[0,0,0,2] }},
                      {{ text:"कर्मचारी नाव: {sevak_name}                                                                                          {month_name_mr} {year}", alignment:"left", fontSize:11, bold:true, margin:[0,0,0,6] }},

                      // Main diary table
                      {{
                        table: {{ widths:["11%","11%","17%","10%","10%","31%","11%"], body:body2 }},
                        layout: {{ hLineWidth:()=>0.6, vLineWidth:()=>0.6,
                                   paddingTop:()=>2, paddingBottom:()=>2,
                                   paddingLeft:()=>3, paddingRight:()=>3 }}
                      }},

                      // Summary stats table
                      {{ text:"", margin:[0,4,0,4] }},
                      {{
                        table: {{ widths:["25%","25%","25%","25%"], body:statsBody }},
                        layout: {{ hLineWidth:()=>0.6, vLineWidth:()=>0.6,
                                   paddingTop:()=>2 , paddingBottom:()=>2  }}
                      }},

                      // Signature table
                      {{ text:"", margin:[0,16,0,25] }},
                      {{
                        table: {{ widths:["33%","33%","34%"], body:sigBody }},
                        layout: 'noBorders',
                      }},
                    ],
                    styles: {{ title: {{ fontSize:13, bold:true }} }}
                  }};
                }}

                function previewPDF2()  {{ pdfMake.createPdf(buildDoc2()).open(); }}
                function downloadPDF2() {{ pdfMake.createPdf(buildDoc2()).download("Masik_Daindini_{month_name_mr}_{year}.pdf"); }}
              </script>
            </body></html>
            """,
            height=80,
        )