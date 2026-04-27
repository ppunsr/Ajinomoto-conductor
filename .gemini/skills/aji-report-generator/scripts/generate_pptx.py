import zipfile, re, os, shutil, openpyxl, sys
from datetime import datetime, date
import calendar

def to_excel_date(dt):
    if not isinstance(dt, datetime): return 0
    return (dt - datetime(1899, 12, 30)).days

def build_num_ref(sheet, col_letter, start_row, end_row, data, format_code='General'):
    formula = f"'{sheet}'!${col_letter}${start_row}:${col_letter}${end_row}"
    xml = f'<c:numRef><c:f>{formula}</c:f><c:numCache><c:formatCode>{format_code}</c:formatCode><c:ptCount val="{len(data)}"/>'
    for i, v in enumerate(data):
        xml += f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
    xml += '</c:numCache></c:numRef>'
    return xml

def build_str_ref(sheet, col_letter, start_row, end_row, data):
    formula = f"'{sheet}'!${col_letter}${start_row}:${col_letter}${end_row}"
    xml = f'<c:strRef><c:f>{formula}</c:f><c:strCache><c:ptCount val="{len(data)}"/>'
    for i, v in enumerate(data):
        xml += f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
    xml += '</c:strCache></c:strRef>'
    return xml

def build_cx_lvl(data, is_num=True, format_code='General'):
    tag = 'num' if is_num else 'str'
    fmt = f' formatCode="{format_code}"' if is_num else ''
    xml = f'<cx:lvl ptCount="{len(data)}"{fmt}>'
    for i, v in enumerate(data):
        xml += f'<cx:pt idx="{i}">{v}</cx:pt>'
    xml += '</cx:lvl>'
    return xml

def get_excel_data(wb, sheet_name, start_row, cols):
    if sheet_name not in wb.sheetnames: return []
    ws = wb[sheet_name]; data = []
    for r in range(start_row, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in cols]
        if not any(v is not None for v in vals): break
        data.append(vals)
    return data

def analyze_data(excel_path, month_str):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    m_target = month_str.lower()[:3]
    month_num = list(calendar.month_abbr).index(month_str[:3].capitalize())
    prev_month_num = month_num - 1 if month_num > 1 else 12
    m_prev = calendar.month_abbr[prev_month_num].lower()
    year = 2026 
    last_day = calendar.monthrange(year, month_num)[1]
    report_end_date = date(year, month_num, last_day)
    res = {'target_month': m_target, 'prev_month': m_prev, 'end_date': report_end_date}
    
    ws = wb['User Engagement']
    ue_rows = []
    for r in range(2, ws.max_row + 1):
        d = ws.cell(row=r, column=1).value
        if isinstance(d, datetime) and d.strftime('%b').lower() in [m_prev, m_target]:
            ue_rows.append({'row': r, 'date': to_excel_date(d), 'new': ws.cell(row=r, column=2).value, 'ret': ws.cell(row=r, column=3).value})
    res['ue'] = ue_rows
    
    ws_sc = wb['gameplay_report(score) ']
    score_rows = []; total_score_sum = 0
    for r in range(2, ws_sc.max_row + 1):
        d = ws_sc.cell(row=r, column=1).value
        if isinstance(d, datetime):
            val = ws_sc.cell(row=r, column=2).value or 0
            if d.date() <= report_end_date: total_score_sum += val
            if d.strftime('%b').lower() in [m_prev, m_target]: score_rows.append({'row': r, 'date': to_excel_date(d), 'val': val})
    res['score'] = score_rows
    
    ws_ti = wb['gameplay_report(time) ']
    time_rows = []; total_time_sum = 0
    for r in range(2, ws_ti.max_row + 1):
        d = ws_ti.cell(row=r, column=1).value
        if isinstance(d, datetime):
            val = ws_ti.cell(row=r, column=2).value or 0
            val_col3 = ws_ti.cell(row=r, column=3).value or 0
            if d.date() <= report_end_date: total_time_sum += val_col3
            if d.strftime('%b').lower() in [m_prev, m_target]: time_rows.append({'row': r, 'date': to_excel_date(d), 'val': val})
    res['time'] = time_rows
    
    ws_f = wb['User_funnel']
    f_col = 2
    for c in range(2, 10):
        val = ws_f.cell(row=5, column=c).value
        if val and str(val).lower()[:3] == m_target:
            f_col = c
            break
    res['funnel'] = {'cats': ['Totalclick', 'Register', 'Player'], 'vals': [ws_f.cell(row=2, column=f_col).value, ws_f.cell(row=3, column=f_col).value, ws_f.cell(row=4, column=f_col).value]}
    
    res['stats'] = {'dau': 0, 'prev_dau': 0, 'mau': 0, 'prev_mau': 0, 'stickiness': 0, 'prev_stickiness': 0, 'total_score': total_score_sum, 'avg_score': 0, 'prev_avg_score': 0, 'score_change': 0, 'total_time': total_time_sum, 'avg_time': 0, 'prev_avg_time': 0, 'time_change': 0}
    
    ws_ue = wb['User Engagement']
    target_col = None
    prev_col = None
    for r in range(1, 15):
        for c in range(1, 15):
            val = ws_ue.cell(row=r, column=c).value
            if isinstance(val, str) and val.lower() == 'month':
                for c2 in range(c+1, c+10):
                    m_val = ws_ue.cell(row=r, column=c2).value
                    if isinstance(m_val, str):
                        if m_val.lower().startswith(m_target): target_col = c2
                        elif m_val.lower().startswith(m_prev): prev_col = c2
                break
        if target_col: break
        
    for r in range(1, 15):
        for c in range(1, 15):
            v = ws_ue.cell(row=r, column=c).value
            if isinstance(v, str):
                if v == 'Monthly active user': 
                    if target_col: res['stats']['mau'] = ws_ue.cell(row=r, column=target_col).value
                    if prev_col: res['stats']['prev_mau'] = ws_ue.cell(row=r, column=prev_col).value
                if v == 'user stickiness': 
                    if target_col: res['stats']['stickiness'] = ws_ue.cell(row=r, column=target_col).value
                    if prev_col: res['stats']['prev_stickiness'] = ws_ue.cell(row=r, column=prev_col).value
                if 'Daily active' in v or 'Daily actuve' in v:
                    if target_col: res['stats']['dau'] = ws_ue.cell(row=r, column=target_col).value
                    if prev_col: res['stats']['prev_dau'] = ws_ue.cell(row=r, column=prev_col).value
            
    click, reg, play = res['funnel']['vals']
    res['stats']['conv_reg'] = (reg/click*100) if click else 0
    res['stats']['drop_off'] = ((reg-play)/reg*100) if reg else 0
    
    import re
    def match_month_header(val, target_m):
        if not val or not isinstance(val, str): return False
        v_clean = val.replace(' ', '').lower()
        return f"avg({target_m.lower()})" in v_clean or f"avg({target_m[:3].lower()})" in v_clean

    full_target_m = calendar.month_name[month_num]
    full_prev_m = calendar.month_name[prev_month_num]

    for c in range(1, 15):
        v = ws_sc.cell(row=2, column=c).value
        if match_month_header(v, full_target_m): 
            res['stats']['avg_score'] = ws_sc.cell(row=3, column=c).value
    
    for c in range(1, 15):
        v = ws_sc.cell(row=2, column=c).value
        if match_month_header(v, full_prev_m):
            ps = ws_sc.cell(row=3, column=c).value
            if ps: 
                res['stats']['prev_avg_score'] = ps
                res['stats']['score_change'] = (res['stats']['avg_score'] - ps)/ps*100
            
    for c in range(1, 15):
        v = ws_ti.cell(row=2, column=c).value
        if match_month_header(v, full_target_m): 
            res['stats']['avg_time'] = ws_ti.cell(row=3, column=c).value
            
    for c in range(1, 15):
        v = ws_ti.cell(row=2, column=c).value
        if match_month_header(v, full_prev_m):
            pt = ws_ti.cell(row=3, column=c).value
            if pt:
                res['stats']['prev_avg_time'] = pt
                res['stats']['time_change'] = (res['stats']['avg_time'] - pt)/pt*100
            
    return res

def update_pptx(excel_path, template_path, output_path, month):
    res = analyze_data(excel_path, month)
    
    import json
    json_path = f"analysis_output_{month}.json"
    
    p3 = p4 = p5 = p6 = p7 = None
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as jf:
            jdata = json.load(jf)
        
        p3 = next((p for p in jdata.get("pages", []) if p["page_number"] == 3), None)
        p4 = next((p for p in jdata.get("pages", []) if p["page_number"] == 4), None)
        p5 = next((p for p in jdata.get("pages", []) if p["page_number"] == 5), None)
        p6 = next((p for p in jdata.get("pages", []) if p["page_number"] == 6), None)
        p7 = next((p for p in jdata.get("pages", []) if p["page_number"] == 7), None)

        prev_m = None
        curr_m = None
        if p3:
            ue_c = p3["sections"]["User Engagement"]["comparison"]
            prev_m = ue_c["previous_month"]
            curr_m = ue_c["current_month"]
            
            funnel_m = p3["sections"]["User Funnel"]["metrics"]
            ue_m = p3["sections"]["User Engagement"]["metrics"]
            
            res['stats']['dau'] = float(ue_m["Daily Active Users (Avg.)"][curr_m])
            res['stats']['prev_dau'] = float(ue_m["Daily Active Users (Avg.)"][prev_m])
            res['stats']['mau'] = float(ue_m["Monthly Active Users"][curr_m])
            res['stats']['prev_mau'] = float(ue_m["Monthly Active Users"][prev_m])
            
            res['stats']['stickiness'] = float(str(ue_m["User Stickiness"][curr_m]).replace('%', ''))
            res['stats']['prev_stickiness'] = float(str(ue_m["User Stickiness"][prev_m]).replace('%', ''))
            
            res['stats']['conv_reg'] = float(str(funnel_m["Conversion rate"]).replace('%', ''))
            res['stats']['drop_off'] = float(str(funnel_m["Drop off"]).replace('%', ''))
            
        if p4 and curr_m and prev_m:
            sc_m = p4["metrics"]["AVG Score per Day"]
            res['stats']['avg_score'] = float(sc_m[curr_m])
            res['stats']['prev_avg_score'] = float(sc_m[prev_m])
            res['stats']['score_change'] = float(str(sc_m["difference"]).replace('%', ''))
            
        if p5 and curr_m and prev_m:
            ti_m = p5["metrics"]["AVG Time per Day"]
            res['stats']['avg_time'] = float(str(ti_m[curr_m]).replace(' minute', ''))
            res['stats']['prev_avg_time'] = float(str(ti_m[prev_m]).replace(' minute', ''))
            res['stats']['time_change'] = float(str(ti_m["difference"]).replace('%', ''))
            
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    temp_dir = 'temp_pptx_gen'
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    with zipfile.ZipFile(template_path, 'r') as zip_ref: zip_ref.extractall(temp_dir)
    
    charts_dir = os.path.join(temp_dir, 'ppt', 'charts')
    for filename in os.listdir(charts_dir):
        if not filename.endswith('.xml'): continue
        path = os.path.join(charts_dir, filename)
        with open(path, 'r', encoding='utf-8') as f: content = f.read()
        orig = content
        
        if filename != 'chartEx1.xml':
            content = re.sub(r'<c:(num|str)Cache>.*?</c:\1Cache>', '', content, flags=re.DOTALL)
            content = re.sub(r'<(c|cx):externalData[^>]*>.*?</\1:externalData>|<(c|cx):externalData[^>]*/>', '', content, flags=re.DOTALL)
            
        if 'User Engagement' in content and res['ue']:
            s, e = res['ue'][0]['row'], res['ue'][-1]['row']
            content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_num_ref("User Engagement", "A", s, e, [r["date"] for r in res["ue"]], "m/d/yy")}</c:cat>', content, flags=re.DOTALL)
            content = re.sub(r'(<c:ser>.*?<c:idx val="0".*?<c:val>).*?(</c:val>)', r'\g<1>' + build_num_ref('User Engagement', 'B', s, e, [r['new'] for r in res['ue']]) + r'\g<2>', content, flags=re.DOTALL)
            content = re.sub(r'(<c:ser>.*?<c:idx val="1".*?<c:val>).*?(</c:val>)', r'\g<1>' + build_num_ref('User Engagement', 'C', s, e, [r['ret'] for r in res['ue']]) + r'\g<2>', content, flags=re.DOTALL)
            content = re.sub(r'<c:strRef>\s*<c:f>.*?\$B\$1</c:f>.*?</c:strRef>', '<c:v>New user</c:v>', content)
            content = re.sub(r'<c:strRef>\s*<c:f>.*?\$C\$1</c:f>.*?</c:strRef>', '<c:v>returning User</c:v>', content)
            
        elif 'gameplay_report(score) ' in content and res['score']:
            s, e = res['score'][0]['row'], res['score'][-1]['row']
            content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_num_ref("gameplay_report(score) ", "A", s, e, [r["date"] for r in res["score"]], "m/d/yy")}</c:cat>', content, flags=re.DOTALL)
            content = re.sub(r'<c:val>.*?</c:val>', f'<c:val>{build_num_ref("gameplay_report(score) ", "B", s, e, [r["val"] for r in res["score"]], '0,"k"')}</c:val>', content, flags=re.DOTALL)
            content = content.replace('<a:schemeClr val="lt1"/>', '<a:srgbClr val="C00000"/>').replace('<a:schemeClr val="dk1"/>', '<a:schemeClr val="bg1"/>')
            
        elif 'gameplay_report(time) ' in content and res['time']:
            s, e = res['time'][0]['row'], res['time'][-1]['row']
            content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_num_ref("gameplay_report(time) ", "A", s, e, [r["date"] for r in res["time"]], "m/d/yy")}</c:cat>', content, flags=re.DOTALL)
            content = re.sub(r'<c:val>.*?</c:val>', f'<c:val>{build_num_ref("gameplay_report(time) ", "B", s, e, [r["val"] for r in res["time"]])}</c:val>', content, flags=re.DOTALL)
            
        elif 'state!' in content:
            rows = get_excel_data(wb, 'state', 12, (1, 2))
            if rows:
                content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_str_ref("state", "A", 12, 11+len(rows), [r[0] for r in rows])}</c:cat>', content, flags=re.DOTALL)
                content = re.sub(r'<c:val>.*?</c:val>', f'<c:val>{build_num_ref("state", "B", 12, 11+len(rows), [r[1] for r in rows])}</c:val>', content, flags=re.DOTALL)
                
        elif 'menu!' in content:
            rows = sorted(get_excel_data(wb, 'menu', 2, (1, 2)), key=lambda x: x[1] or 0, reverse=True)[::-1]
            if rows:
                content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_str_ref("menu", "A", 2, 1+len(rows), [r[0] for r in rows])}</c:cat>', content, flags=re.DOTALL)
                content = re.sub(r'<c:val>.*?</c:val>', f'<c:val>{build_num_ref("menu", "B", 2, 1+len(rows), [r[1] for r in rows])}</c:val>', content, flags=re.DOTALL)
                
        elif 'score!' in content:
            rows = sorted(get_excel_data(wb, 'score', 2, (2, 3)), key=lambda x: x[1] or 0, reverse=True)[:10][::-1]
            if rows:
                content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_str_ref("score", "B", 2, 1+len(rows), [r[0] for r in rows])}</c:cat>', content, flags=re.DOTALL)
                content = re.sub(r'<c:val>.*?</c:val>', f'<c:val>{build_num_ref("score", "C", 2, 1+len(rows), [r[1] for r in rows])}</c:val>', content, flags=re.DOTALL)
                
        elif filename == 'chartEx1.xml':
            f_col_letter = 'C' if month.lower().startswith('mar') else 'B'
            content = re.sub(r'<cx:strDim[^>]*>.*?</cx:strDim>', f'<cx:strDim type="cat"><cx:f>User_funnel!$A$2:$A$4</cx:f>{build_cx_lvl(res["funnel"]["cats"], False)}</cx:strDim>', content, flags=re.DOTALL)
            content = re.sub(r'<cx:numDim[^>]*>.*?</cx:numDim>', f'<cx:numDim type="val"><cx:f>User_funnel!${f_col_letter}$2:${f_col_letter}$4</cx:f>{build_cx_lvl(res["funnel"]["vals"], True, "General")}</cx:numDim>', content, flags=re.DOTALL)
            
        if content != orig:
            with open(path, 'w', encoding='utf-8') as f: f.write(content)
            
    drawings_dir = os.path.join(temp_dir, 'ppt', 'drawings')
    if os.path.exists(drawings_dir):
        for filename in os.listdir(drawings_dir):
            if filename == 'drawing1.xml':
                dpath = os.path.join(drawings_dir, filename)
                with open(dpath, 'r', encoding='utf-8') as df:
                    d_content = df.read()
                
                # Surgically remove only TextBox 33 (which contains the ghost data) to preserve the red rectangle
                d_content = re.sub(r'<cdr:relSizeAnchor(?:(?!</cdr:relSizeAnchor>).)*?name="TextBox 33".*?</cdr:relSizeAnchor>', '', d_content, flags=re.DOTALL)
                
                with open(dpath, 'w', encoding='utf-8') as df:
                    df.write(d_content)

    slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
    for filename in os.listdir(slides_dir):
        if not filename.endswith('.xml') or '_' in filename: continue
        path = os.path.join(slides_dir, filename)
        with open(path, 'r', encoding='utf-8') as f: content = f.read()
        orig = content

        content = content.replace("Data Period: Data Period:", "Data Period:")
        content = re.sub(r'2025/11/28\s*.\s*2026/0(?:</a:t>.*?<a:t>)*?3(?:</a:t>.*?<a:t>)*?/31', f"2025/11/28 - {res['end_date'].strftime('%Y/%m/%d')}", content)

        target_full = month.capitalize()
        prev_full = calendar.month_name[list(calendar.month_abbr).index(res['prev_month'].capitalize())]
        
        content = re.sub(r'February(</a:t>.*?<a:t>)-2026', fr'__PREV__\1-2026', content)
        content = content.replace("February-2026", "__PREV__-2026")
        
        content = re.sub(r'March(</a:t>.*?<a:t>)-2026', fr'__TARGET__\1-2026', content)
        content = content.replace("March-2026", "__TARGET__-2026")
        
        content = content.replace("__PREV__", prev_full)
        content = content.replace("__TARGET__", target_full)
            
        s = res['stats']
        
        # Dynamic Analysis Paragraph replacement
        def replace_paragraph(xml_content, marker, new_text):
            idx = xml_content.find(marker)
            if idx != -1:
                start_p = xml_content.rfind('<a:p>', 0, idx)
                end_p = xml_content.find('</a:p>', idx) + 6
                if start_p != -1 and end_p != -1:
                    new_p = f'<a:p><a:r><a:rPr lang="en-US" sz="1400"/><a:t>{new_text}</a:t></a:r></a:p>'
                    return xml_content[:start_p] + new_p + xml_content[end_p:]
            return xml_content
            
        if filename == 'slide3.xml' and p3 and 'key_finding' in p3:
            content = replace_paragraph(content, "Although user stickiness improved", p3['key_finding'])
        elif filename == 'slide4.xml' and p4 and 'key_finding' in p4:
            content = replace_paragraph(content, "Game performance experienced a significant downturn", p4['key_finding'])
        elif filename == 'slide5.xml' and p5 and 'key_finding' in p5:
            content = replace_paragraph(content, "Player engagement time has declined significantly", p5['key_finding'])
        elif filename == 'slide6.xml' and p6 and 'key_finding' in p6:
            content = replace_paragraph(content, "Users show a strong preference for", p6['key_finding'])
            
        if 'Daily Active Users' in content:
            idx = content.find('Daily Active Users')
            start_idx = content.rfind('<p:sp>', 0, idx)
            end_idx = content.find('</p:sp>', idx) + len('</p:sp>')
            orig_box = content[start_idx:end_idx]
            
            # The template only has the TARGET month's DAU box (which contains 3.0, 21, 14%).
            target_box = orig_box.replace('3.0</a:t>', f'{s["dau"]:.1f}</a:t>')
            target_box = target_box.replace('>21</a:t>', f'>{s["mau"]:.0f}</a:t>')
            target_box = target_box.replace('>14%</a:t>', f'>{s["stickiness"]:.2f}%</a:t>')
            
            # Create a box for the PREVIOUS month by cloning the target box and shifting it left
            prev_box = orig_box.replace('3.0</a:t>', f'{s["prev_dau"]:.1f}</a:t>')
            prev_box = prev_box.replace('>21</a:t>', f'>{s["prev_mau"]:.0f}</a:t>')
            prev_box = prev_box.replace('>14%</a:t>', f'>{s["prev_stickiness"]:.2f}%</a:t>')

            # Shift X position to center it within the left red rectangle (Rectangle 25)
            # Rectangle 25 center is 6672273. Prev box width is 1931628. 6672273 - (1931628 / 2) = 5706459
            prev_box = re.sub(r'<a:off x="\d+"', '<a:off x="5706459"', prev_box)

            # Ensure unique IDs and names for the new shape to avoid corruption or overlap
            import uuid
            new_id = str(uuid.uuid4().int % 10000)
            new_creation_id = str(uuid.uuid4()).upper()
            prev_box = re.sub(r'id="\d+"', f'id="{new_id}"', prev_box, count=1)
            prev_box = re.sub(r'name="TextBox \d+"', f'name="TextBox {new_id}"', prev_box, count=1)
            prev_box = re.sub(r'creationId id="\{[A-F0-9-]+\}"', f'creationId id="{{{new_creation_id}}}"', prev_box, count=1)

            content = content[:start_idx] + prev_box + target_box + content[end_idx:]
            
            # Remove the hardcoded previous month image (rId4) which contained the old Feb data
            if filename == 'slide3.xml':
                content = re.sub(r'<p:pic>.*?<a:blip r:embed="rId4".*?</p:pic>', '', content, flags=re.DOTALL)

            
        content = re.sub(r'88% conversion rate', f'{s["conv_reg"]:.0f}% conversion rate', content)
        content = re.sub(r'55% drop off', f'{s["drop_off"]:.0f}% drop off', content)
        content = re.sub(r'55% drop-off', f'{s["drop_off"]:.0f}% drop-off', content)
        
        # Use placeholders for Score to prevent double replacement
        content = content.replace('28,986', '__PREV_SCORE__')
        content = content.replace('12,265', '__TARGET_SCORE__')
        content = content.replace('__PREV_SCORE__', f'{s["prev_avg_score"]:,.0f}')
        content = content.replace('__TARGET_SCORE__', f'{s["avg_score"]:,.0f}')
        
        content = content.replace('7,796,142', f'{s["total_score"]:,}')
        content = content.replace('(-57.7%)', f'({s["score_change"]:.1f}%)')
        
        # Use placeholders for Time to prevent double replacement
        content = content.replace('32 minute', '__PREV_TIME__')
        content = content.replace('15 minute', '__TARGET_TIME__')
        content = content.replace('__PREV_TIME__', f'{s["prev_avg_time"]:.0f} minute')
        content = content.replace('__TARGET_TIME__', f'{s["avg_time"]:.0f} minute')
        
        content = content.replace('7,876', f'{int(s["total_time"]):,}')
        content = content.replace('(- 53.1%)', f'({s["time_change"]:.1f}%)')
        
        if 'Hours' in content: 
            content = re.sub(r'\([0-9.]+ Hours\)', f'({s["total_time"]/60:.2f} Hours)', content)
            
        if content != orig:
            with open(path, 'w', encoding='utf-8') as f: f.write(content)
            
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, ds, fs in os.walk(temp_dir):
            for file in fs:
                fpath = os.path.join(root, file); zipf.write(fpath, os.path.relpath(fpath, temp_dir))
    shutil.rmtree(temp_dir)
    print(f'Saved to {output_path}')

if __name__ == "__main__":
    if len(sys.argv) < 5: sys.exit(1)
    update_pptx(sys.argv[1], sys.argv[2], sys.argv[4], sys.argv[3])
