from flask import Flask, request, send_file, render_template, jsonify
from replacer import replace_in_docx
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter
import os, json, tempfile, zipfile, copy, io, re
from concurrent.futures import ThreadPoolExecutor

app = Flask(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.xlsx")

# ── WBS 상수 ──────────────────────────────────────────────────────────────────
ORIG_BASE_COL  = 6
ORIG_R7_START  = 8
ORIG_R14_START = 17
DEL_START      = 6
DEL_COUNT      = 12

ORIG_GANTT = {
    7:  [(8,5,0.8)],
    8:  [(c,5,0.8) for c in range(9,17)],
    9:  [(c,5,0.8) for c in range(9,17)],
    10: [(c,5,0.8) for c in range(9,17)],
    11: [(17,5,0.8),(20,5,0.8)],
    12: [(17,9,0.8),(20,9,0.8)],
    13: [(17,9,0.8),(20,9,0.8)],
    14: [(17,5,0.8)],
    15: [(c,2,-0.1) for c in range(18,22)],
    16: [(22,5,0.8)],
    17: [(c,2,-0.1) for c in range(23,27)],
    18: [(27,5,0.8)],
    19: [(c,2,-0.1) for c in range(28,40)],
    20: [(40,5,0.8)],
    21: [(c,2,-0.1) for c in range(41,45)],
    22: [(45,5,0.8)],
    23: [(46,5,0.8)],
}

ORIG_R24 = {
    21:'rgb',22:'theme',23:'theme',24:'theme',25:'theme',
    26:'rgb',27:'theme',28:'theme',29:'theme',30:'theme',
    31:'rgb',32:'theme',33:'theme',34:'theme',35:'theme',
    36:'theme',37:'theme',38:'theme',39:'theme',40:'theme',
    41:'theme',42:'theme',43:'theme',
    44:'rgb',45:'theme',46:'theme',47:'theme',48:'theme',
    49:'rgb',50:'rgb',
}

def nc(old):
    if old < DEL_START: return old
    if old < DEL_START + DEL_COUNT: return None
    return old - DEL_COUNT

def is_colored(f):
    if f.fill_type != 'solid': return False
    ft = f.fgColor.type
    if ft == 'theme' and f.fgColor.theme != 0: return True
    if ft == 'rgb':
        try: return f.fgColor.rgb not in ('00000000','FFFFFFFF','00FFFFFF')
        except: pass
    return False

def generate_wbs(client_name, start_date_str, include_vuln_self):
    from datetime import date
    sd            = date.fromisoformat(start_date_str)
    start_year    = sd.year
    start_month   = sd.month
    week_of_month = (sd.day - 1) // 7

    orig_start_col = ORIG_BASE_COL + (start_month - 1) * 4 + week_of_month
    SHIFT          = orig_start_col - ORIG_R7_START

    med    = Side(border_style="medium")
    none_s = Side(border_style=None)

    wb     = load_workbook(TEMPLATE_PATH)
    ws     = wb.active
    wb_src = load_workbook(TEMPLATE_PATH)
    ws_src = wb_src.active

    # 고객사명 치환
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell): continue
            if cell.value and isinstance(cell.value, str) and "고객사명" in cell.value:
                cell.value = cell.value.replace("고객사명", client_name)

    # 간트 초기화
    for rn in range(7, 25):
        for col in range(6, 73):
            c = ws.cell(rn, col)
            if isinstance(c, MergedCell): continue
            try:
                if is_colored(c.fill): c.fill = PatternFill(fill_type=None)
            except: pass

    # 간트 shift
    for rn, cells in ORIG_GANTT.items():
        for (oc, _, __) in cells:
            new_col = oc + SHIFT
            if new_col < 1: continue
            target = ws.cell(rn, new_col)
            if not isinstance(target, MergedCell):
                target.fill = copy.copy(ws_src.cell(rn, oc).fill)

    # Row24 취약점
    new_r14  = ORIG_R14_START + SHIFT
    adjust   = min(ORIG_R24) - ORIG_R14_START
    src_yell = ws_src.cell(24, 21)
    src_blue = ws_src.cell(24, 22)
    if include_vuln_self:
        for oc, ftype in ORIG_R24.items():
            new_col = new_r14 + (oc - ORIG_R14_START - adjust)
            if new_col < 1: continue
            c = ws.cell(24, new_col)
            if isinstance(c, MergedCell): continue
            c.fill = copy.copy(src_yell.fill if ftype == 'rgb' else src_blue.fill)

    # 병합 완전 해제
    for mr in [str(m) for m in ws.merged_cells.ranges]:
        ws.unmerge_cells(mr)

    # BF(orig 58) 헤더 설정 (삭제 전)
    ref54 = ws_src.cell(5, 54)
    for rn, val in [(4,None),(5,"2월"),(6,"1주")]:
        c = ws.cell(rn, 58)
        if val: c.value = val
        c.font=copy.copy(ref54.font); c.fill=copy.copy(ref54.fill)
        c.alignment=copy.copy(ref54.alignment); c.border=copy.copy(ref54.border)
    for rn in range(7, 25):
        ws.cell(rn, 58).border = copy.copy(ws_src.cell(rn, 54).border)

    # 헤더값 설정 (삭제 전)
    ws.cell(4, 54).value = f"{start_year+1}년"
    for i, oc in enumerate([18,22,26,30,34,38,42,46,50,54,58]):
        real_m = ((start_month + i - 1) % 12) + 1
        ws.cell(5, oc).value = f"{real_m}월"

    # 컬럼 삭제
    ws.delete_cols(DEL_START, DEL_COUNT)

    # extra 컬럼 (47~53)
    m_feb = ((start_month + 9  - 1) % 12) + 1
    m_mar = ((start_month + 10 - 1) % 12) + 1
    extra = [
        (47,f"{m_feb}월","2주",False,False),
        (48,f"{m_feb}월","3주",False,False),
        (49,f"{m_feb}월","4주",False,True),
        (50,f"{m_mar}월","1주",True, False),
        (51,f"{m_mar}월","2주",False,False),
        (52,f"{m_mar}월","3주",False,False),
        (53,f"{m_mar}월","4주",False,True),
    ]
    ref5=ws_src.cell(5,54); ref6=ws_src.cell(6,54); ref4=ws_src.cell(4,54)
    for col, month, week, is_first, is_last in extra:
        c6=ws.cell(6,col); c6.value=week
        c6.font=copy.copy(ref6.font); c6.fill=copy.copy(ref6.fill)
        c6.alignment=copy.copy(ref6.alignment); c6.border=copy.copy(ref6.border)
        c5=ws.cell(5,col); c5.value=month if is_first else None
        c5.font=copy.copy(ref5.font); c5.fill=copy.copy(ref5.fill)
        c5.alignment=copy.copy(ref5.alignment)
        c5.border=Border(
            left=med if is_first else none_s,
            right=med if (is_last or is_first) else none_s,
            top=med, bottom=med)
        c4=ws.cell(4,col)
        c4.font=copy.copy(ref4.font); c4.fill=copy.copy(ref4.fill)
        c4.alignment=copy.copy(ref4.alignment)
        c4.border=Border(left=none_s,right=none_s,top=med,bottom=none_s)
        for rn in range(7,25):
            ws.cell(rn,col).border=copy.copy(ws_src.cell(rn,54).border)
        ws.column_dimensions[get_column_letter(col)].width=4.9

    # 연도/월 대표셀 값 설정 (병합 전)
    ws.cell(4, 6).value  = f"{start_year}년"
    ws.cell(4, 42).value = f"{start_year+1}년"
    ws.cell(5, 46).value = f"{m_feb}월"
    ws.cell(5, 50).value = f"{m_mar}월"

    # F4 서식 복원
    f4 = ws.cell(4, 6)
    pf = PatternFill(fill_type="solid")
    pf.fgColor.type="theme"; pf.fgColor.theme=2; pf.fgColor.tint=-0.1
    f4.fill=pf
    f4.font=Font(name="맑은 고딕",bold=True,size=15,color="FF000000")
    f4.alignment=Alignment(horizontal="center",vertical="center")
    f4.border=Border(left=med,right=med,top=med,bottom=none_s)

    # 병합 재설정
    def sm(nsc,sr,nec,er):
        if nsc and nec and nsc<=nec:
            try: ws.merge_cells(start_row=sr,start_column=nsc,end_row=er,end_column=nec)
            except: pass

    for s in [nc(c) for c in range(18,51,4) if nc(c)]:
        sm(s,5,s+3,5)
    sm(nc(54),5,nc(54)+3,5)
    sm(nc(58),5,49,5)
    sm(50,5,53,5)
    sm(6,4,41,4)
    sm(42,4,53,4)
    for (sc,sr,ec,er) in [(2,4,5,5),(2,8,2,10),(2,11,2,13),
                           (3,8,3,10),(3,11,3,13),(3,15,3,18),(3,20,3,22),
                           (4,9,4,10),(4,12,4,13)]:
        try: ws.merge_cells(start_row=sr,start_column=sc,end_row=er,end_column=ec)
        except: pass

    # 열너비 4.9
    for col in range(6, 54):
        ws.column_dimensions[get_column_letter(col)].width=4.9

    # 전체 글자 검은색
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell,MergedCell): continue
            if cell.font:
                f=cell.font
                cell.font=Font(name=f.name,size=f.size,bold=f.bold,
                               italic=f.italic,underline=f.underline,
                               strike=f.strike,color="FF000000")

    # 취약점 자체점검 미포함
    if not include_vuln_self:
        ws.delete_rows(24,1)
        for col in range(2,6):
            c=ws.cell(23,col); b=c.border
            c.border=Border(left=b.left,right=b.right,top=b.top,bottom=med)

    safe_name=re.sub(r'[\\/:*?"<>|]','_',client_name)
    filename=f"{safe_name}_CSAP_간편등급_컨설팅_일정_v2_1.xlsx"
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf, filename


# ── 워드 치환 관련 ────────────────────────────────────────────────────────────
def process_file(file_data):
    file_bytes, filename, rules, tmpdir = file_data
    input_path  = os.path.join(tmpdir, f"in_{filename}")
    output_path = os.path.join(tmpdir, f"out_{filename}")
    with open(input_path, "wb") as f:
        f.write(file_bytes)
    replace_in_docx(input_path, output_path, rules)
    return output_path, filename


# ── 라우트 ───────────────────────────────────────────────────────────────────
@app.route("/")
def home():
    return render_template("home.html")

@app.route("/replacer")
def replacer():
    return render_template("index.html")   # 기존 워드 치환 페이지

@app.route("/replace", methods=["POST"])
def replace():
    if "files" not in request.files:
        return jsonify({"error": "파일이 없습니다"}), 400
    files = request.files.getlist("files")
    rules_json = request.form.get("rules", "{}")
    rules = json.loads(rules_json)
    if not rules:
        return jsonify({"error": "치환 규칙을 입력해주세요"}), 400

    with tempfile.TemporaryDirectory() as tmpdir:
        if len(files) == 1:
            file = files[0]
            if not file.filename.endswith(".docx"):
                return jsonify({"error": ".docx 파일만 지원합니다"}), 400
            input_path  = os.path.join(tmpdir, file.filename)
            output_path = os.path.join(tmpdir, f"out_{file.filename}")
            file.save(input_path)
            replace_in_docx(input_path, output_path, rules)
            return send_file(output_path, as_attachment=True, download_name=file.filename)
        else:
            file_data_list = []
            for file in files:
                if not file.filename.endswith(".docx"):
                    continue
                file_data_list.append((file.read(), file.filename, rules, tmpdir))
            with ThreadPoolExecutor() as executor:
                results = list(executor.map(process_file, file_data_list))
            zip_path = os.path.join(tmpdir, "replaced_files.zip")
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for output_path, filename in results:
                    zipf.write(output_path, filename)
            return send_file(zip_path, as_attachment=True, download_name="replaced_files.zip")

@app.route("/wbs")
def wbs():
    return render_template("wbs.html")

@app.route("/wbs/generate", methods=["POST"])
def wbs_generate():
    client_name       = request.form.get("client_name", "").strip()
    start_date        = request.form.get("start_date", "")
    include_vuln_self = request.form.get("include_vuln_self") == "true"
    if not client_name or not start_date:
        return jsonify({"error": "고객사명과 시작일을 입력해주세요."}), 400
    try:
        buf, filename = generate_wbs(client_name, start_date, include_vuln_self)
        return send_file(buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True, download_name=filename)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)
