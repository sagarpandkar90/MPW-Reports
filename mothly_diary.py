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

    # ─────────────────────────────────────────────────────────────
    # Load movement.xlsx (must be in project folder)
    # Structure:
    #   Col 0  → गाव  (fixed = शेळगाव for all data rows)
    #   Col 1  → वस्तीचे नाव
    #   Col 2-9 → village info (ignored for diary)
    #   Col 10+ → 24 half-month columns  (जाने-3, जाने-7, फेब-4, …)
    #   Last row → सा.सुट्टी  (holiday dates per half-month)
    # ─────────────────────────────────────────────────────────────
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

        header     = rows[0]                 # row 0 = column headers
        month_cols = list(header[10:])       # 24 half-month labels

        # ── Marathi short name → month number ────────────────────
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

        data_rows   = []   # list of (gav, vasti, [24 date values])
        holiday_row = None

        for row in rows[1:]:
            gav   = str(row[0]).strip() if row[0] else ""
            vasti = str(row[1]).strip() if row[1] else ""
            if "सुट्टी" in gav or "सुट्टी" in vasti:
                holiday_row = list(row[10:])
            else:
                date_vals = list(row[10:])
                data_rows.append((gav, vasti, date_vals))

        # Parse holiday dates per half (24 values)
        if holiday_row:
            h_sets = [parse_holiday_cell(v) for v in holiday_row]
        else:
            h_sets = [set()] * 24

        # Build half-month metadata list (24 entries)
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

    # ─────────────────────────────────────────────────────────────
    # Inputs
    # ─────────────────────────────────────────────────────────────
    EN_MONTHS = [
        "January","February","March","April","May","June",
        "July","August","September","October","November","December",
    ]
    MR_MONTHS = [
        "जानेवारी","फेब्रुवारी","मार्च","एप्रिल","मे","जून",
        "जुलै","ऑगस्ट","सप्टेंबर","ऑक्टोबर","नोव्हेंबर","डिसेंबर",
    ]

    cols       = st.columns([1, 1, 2])
    month_en   = cols[0].selectbox(
        "महिना", EN_MONTHS,
        index=datetime.date.today().month - 1,
        format_func=lambda x: MR_MONTHS[EN_MONTHS.index(x)]
    )
    year       = cols[1].number_input("वर्ष", min_value=2000, max_value=2100,
                                      value=datetime.date.today().year)
    sevak_name = cols[2].text_input("आरोग्य सेवक नाव")
    upkendra   = st.text_input("उपकेंद्राचे नाव")

    month_num_sel = EN_MONTHS.index(month_en) + 1
    month_name_mr = MR_MONTHS[month_num_sel - 1]

    # गाव is always read from Excel col 0 (शेळगाव) — user can override below
    fixed_gav = data_rows[0][0] if data_rows else "शेळगाव"
    gav_input = st.text_input("गाव (बदलायचे असल्यास बदला)", value=fixed_gav)

    # ─────────────────────────────────────────────────────────────
    # Find the two half-month columns for the selected month
    # ─────────────────────────────────────────────────────────────
    halves = [m for m in half_meta if m["month_num"] == month_num_sel]
    if len(halves) < 2:
        st.error("Excel मध्ये या महिन्याचा डेटा नाही!")
        return
    half1, half2 = halves[0], halves[1]

    # ─────────────────────────────────────────────────────────────
    # Build  day → वस्तीचे नाव  from Excel dates
    # ─────────────────────────────────────────────────────────────
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
    day_vasti_map   = {**map1, **map2}
    all_holiday_days = half1["holidays"] | half2["holidays"]

    # ─────────────────────────────────────────────────────────────
    # Holiday names — fixed dates that never change year to year
    # Floating festivals (Holi, Diwali, Ganesh, Eid, etc.) are
    # stored in the Excel with their correct date for that year,
    # so we identify them by their sequential position in the
    # month's holiday list rather than a hardcoded (month, day).
    # ─────────────────────────────────────────────────────────────

    # Fixed holidays: (month, day) → name  — these dates never shift
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

    # Floating festivals by month — names in chronological date order
    # for the non-fixed holidays in that month.
    # Fixed holidays (Jan 26, Feb 19, Apr 14, May 1, Aug 15, Oct 2, Dec 25)
    # are already handled by FIXED_HOLIDAYS above, so floating list only
    # covers the remaining dates in that month.
    FLOATING_BY_MONTH = {
        3:  ["होळी", "धुळवड", "रमजान ईद (ईद-उल-फित्र)"],
        4:  ["गुड फ्रायडे"],               # Apr 14 (Ambedkar) handled by FIXED
        5:  ["बुद्ध पौर्णिमा"],             # May 1 (Maharashtra Day) handled by FIXED
        6:  ["ईद-उल-अधा (बकरी ईद)"],
        7:  ["रथयात्रा"],
        8:  ["पारसी नववर्ष (नवरोज)", "गणेश चतुर्थी"],
        9:  ["मोहरम"],
        10: ["दसरा (विजयादशमी)"],          # Oct 2 (Gandhi Jayanti) handled by FIXED
        11: ["दिवाळी (लक्ष्मीपूजन)", "दिवाळी पाडवा", "गुरुनानक जयंती"],
    }

    # Pre-build a (month, day) → name map using Excel dates + ordered names
    def build_holiday_name_map(year_val):
        name_map = dict(FIXED_HOLIDAYS)  # start with fixed ones

        for m, names in FLOATING_BY_MONTH.items():
            # get all holiday days for this month from Excel halves
            h1 = next((hm for hm in half_meta if hm["month_num"] == m and hm["part"] == 1), None)
            h2 = next((hm for hm in half_meta if hm["month_num"] == m and hm["part"] == 2), None)
            days_in_month = set()
            if h1: days_in_month |= h1["holidays"]
            if h2: days_in_month |= h2["holidays"]

            # Remove days already covered by fixed holidays
            floating_days = sorted(
                d for d in days_in_month if (m, d) not in FIXED_HOLIDAYS
            )

            for i, day in enumerate(floating_days):
                if i < len(names):
                    name_map[(m, day)] = names[i]
                else:
                    name_map[(m, day)] = "सार्वजनिक सुट्टी"

        return name_map

    HOLIDAY_NAME_MAP = build_holiday_name_map(year)

    def get_holiday_name(m, d):
        return HOLIDAY_NAME_MAP.get((m, d), "सार्वजनिक सुट्टी")

    # ─────────────────────────────────────────────────────────────
    # Special days in the selected month
    # ─────────────────────────────────────────────────────────────
    total_days = calendar.monthrange(year, month_num_sel)[1]
    dates      = [datetime.date(year, month_num_sel, d) for d in range(1, total_days + 1)]

    mondays  = [d for d in dates if d.weekday() == 0]
    tuesdays = [d for d in dates if d.weekday() == 1]

    first_monday  = mondays[0]  if len(mondays)  >= 1 else None
    first_tuesday = tuesdays[0] if len(tuesdays) >= 1 else None
    third_tuesday = tuesdays[2] if len(tuesdays) >= 3 else None

    # ─────────────────────────────────────────────────────────────
    # Build DataFrame
    # ─────────────────────────────────────────────────────────────
    gav_col     = []
    vasti_col   = []   # भेटीचे गाव → now = वस्तीचे नाव
    tapshil_col = []
    shera_col   = []

    for dt in dates:
        day        = dt.day
        is_sunday  = (dt.weekday() == 6)
        is_holiday = (day in all_holiday_days)

        # ── Holiday checked FIRST (holiday > sunday > normal) ──────────
        if is_holiday:
            hname = get_holiday_name(month_num_sel, day)
            gav_col.append("")                # empty – as required for festivals
            vasti_col.append("सा.सुट्टी")
            tapshil_col.append(hname)
            shera_col.append("")              # blank – officer writes

        elif is_sunday:
            gav_col.append("")                # empty for Sundays
            vasti_col.append("---")
            tapshil_col.append("रविवार")
            shera_col.append("")              # blank – officer writes

        elif dt == first_monday:
            vasti = day_vasti_map.get(day, "")
            gav_col.append(gav_input)
            vasti_col.append(vasti)
            tapshil_col.append("नियमित लसीकरण भुजबळ वस्ती")
            shera_col.append("")

        elif dt in (first_tuesday, third_tuesday):
            vasti = day_vasti_map.get(day, "")
            gav_col.append(gav_input)
            vasti_col.append(vasti)
            tapshil_col.append("नियमित लसीकरण शेळगाव")
            shera_col.append("")

        else:
            vasti = day_vasti_map.get(day, "")
            gav_col.append(gav_input)
            vasti_col.append(vasti)
            if vasti and vasti not in ("---", ""):
                tapshil_col.append(f"घरभेट व कंटेनर सर्वेक्षण {vasti}")
            else:
                tapshil_col.append("बहुउद्देशीय कार्य प्रा. आ. केंद्र शेळगाव")
            shera_col.append("")

    # ─────────────────────────────────────────────────────────────
    # Info panel
    # ─────────────────────────────────────────────────────────────
    st.info(
        f"📋 **Excel कोड:** `{half1['label']}` | `{half2['label']}`  \n"
        f"🎉 **पहिल्या भागातील सुट्ट्या (तारखा):** {sorted(half1['holidays']) or 'नाही'}  \n"
        f"🎉 **दुसऱ्या भागातील सुट्ट्या (तारखा):** {sorted(half2['holidays']) or 'नाही'}"
    )

    # ─────────────────────────────────────────────────────────────
    # DataFrame  (column names updated per user request)
    # ─────────────────────────────────────────────────────────────
    df = pd.DataFrame({
        "भेटीचा दिनांक": [d.strftime("%d-%m-%Y") for d in dates],
        "गाव":            gav_col,
        "वस्तीचे नाव":   vasti_col,
        "कामाचा तपशील": tapshil_col,
        "शेरा":           shera_col,
    })

    st.write("### मासिक डायरी (Editable)")
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        column_config={
            "गाव":            st.column_config.TextColumn(width="small"),
            "वस्तीचे नाव":   st.column_config.TextColumn(width="medium"),
            "कामाचा तपशील": st.column_config.TextColumn(width="large"),
        },
    )

    # ─────────────────────────────────────────────────────────────
    # JSON → pdfMake
    # ─────────────────────────────────────────────────────────────
    data_json = edited_df.to_dict(orient="records")
    json_js   = json.dumps(data_json, ensure_ascii=False)

    # ─────────────────────────────────────────────────────────────
    # Font
    # ─────────────────────────────────────────────────────────────
    font_path = BASE_DIR / "fonts" / "NotoSerifDevanagari-VariableFont_wdth,wght.ttf"
    if not font_path.exists():
        st.error("❌ Font 'NotoSerifDevanagari-...' missing in fonts/ folder!")
        return
    font_b64 = base64.b64encode(font_path.read_bytes()).decode()

    # ─────────────────────────────────────────────────────────────
    # PDF via pdfMake
    # ─────────────────────────────────────────────────────────────
    components.html(
        f"""
        <html>
        <head><meta charset="utf-8"/>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.72/pdfmake.min.js"></script>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.72/vfs_fonts.js"></script>
        </head>
        <body style="font-family:sans-serif;padding:10px;">
          <button onclick="previewPDF()"
            style="padding:9px 16px;background:#2196F3;color:#fff;border:none;border-radius:6px;margin-right:8px;font-size:14px;cursor:pointer;">
            👁 Preview
          </button>
          <button onclick="downloadPDF()"
            style="padding:9px 16px;background:#4CAF50;color:#fff;border:none;border-radius:6px;font-size:14px;cursor:pointer;">
            ⬇ Download PDF
          </button>

          <script>
            const data = {json_js};

            pdfMake.vfs["Marathi.ttf"] = "{font_b64}";
            pdfMake.fonts = {{
              MarathiFont: {{
                normal:"Marathi.ttf", bold:"Marathi.ttf",
                italics:"Marathi.ttf", bolditalics:"Marathi.ttf"
              }}
            }};

            const monthLabel = "{month_name_mr} {year}";

            // ── table body ─────────────────────────────────────────
            const HDR_COLOR  = "#BBDEFB";   // blue-ish header
            const SUTTI_CLR  = "#FFF9C4";   // yellow  – holiday
            const RAVI_CLR   = "#FCE4EC";   // pink    – sunday

            const body = [[
              {{ text:"भेटीचा दिनांक", bold:true, alignment:"center", fillColor:HDR_COLOR }},
              {{ text:"गाव",            bold:true, alignment:"center", fillColor:HDR_COLOR }},
              {{ text:"कामाचा तपशील", bold:true, alignment:"center", fillColor:HDR_COLOR }},
              {{ text:"शेरा",           bold:true, alignment:"center", fillColor:HDR_COLOR }},
            ]];

            data.forEach(r => {{
              const vasti   = r["वस्तीचे नाव"]  || "";
              const tapshil = r["कामाचा तपशील"] || "";
              const shera   = r["शेरा"]           || "";
              const isSutti = vasti   === "सा.सुट्टी";
              const isRavi  = tapshil === "रविवार";
              const fc      = isSutti ? SUTTI_CLR : isRavi ? RAVI_CLR : null;

              body.push([
                {{ text: r["भेटीचा दिनांक"] || "", alignment:"center", fillColor:fc }},
                {{ text: r["गाव"]            || "", alignment:"center", fillColor:fc }},
                {{ text: tapshil,                   alignment:"center",   fillColor:fc}},
                {{ text: shera,                     alignment:"center", fillColor:fc}},
              ]);
            }});

            const docDef = {{
              pageSize:    "A4",
              pageMargins: [30, 14, 30, 45],
              background:  {{ canvas:[{{ type:"rect", x:7, y:7, w:581, h:828, lineWidth:1 }}] }},
              defaultStyle:{{ font:"MarathiFont", fontSize:11 }},
              content:[
                {{ text:"मासिक डायरी", style:"title", alignment:"center", margin:[0,4,0,3] }},
                {{ text:monthLabel, absolutePosition:{{x:420,y:38}}, fontSize:13, bold:true }},
                {{ text:"आरोग्य सेवक नाव: {sevak_name}", margin:[0,0,0,2], fontSize:12 }},
                {{ text:"उपकेंद्र: {upkendra}",           margin:[0,0,0,5], fontSize:12 }},
                {{
                  table:{{
                    widths:["17%","15%","38%","30%"],
                    body: body
                  }},
                  layout:{{
                    hLineWidth:()=>0.7, vLineWidth:()=>0.7,
                    paddingTop:()=>1.5, paddingBottom:()=>1.5,
                    paddingLeft:()=>3,  paddingRight:()=>3
                  }}
                }}
              ],
              styles:{{ title:{{ fontSize:17, bold:true }} }}
            }};

            function previewPDF()  {{ pdfMake.createPdf(docDef).open(); }}
            function downloadPDF() {{ pdfMake.createPdf(docDef).download("Masik_Diary_{month_name_mr}_{year}.pdf"); }}
          </script>
        </body>
        </html>
        """,
        height=800,
        scrolling=True,
    )