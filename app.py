from flask import Flask, request, send_file, render_template, jsonify
from replacer import replace_in_docx
import os, json, tempfile, zipfile
from concurrent.futures import ThreadPoolExecutor  # 추가!

app = Flask(__name__)

def process_file(file_data):
    """각 파일을 처리하는 함수"""
    file_bytes, filename, rules, tmpdir = file_data
    input_path  = os.path.join(tmpdir, f"in_{filename}")
    output_path = os.path.join(tmpdir, f"out_{filename}")
    with open(input_path, "wb") as f:
        f.write(file_bytes)
    replace_in_docx(input_path, output_path, rules)
    return output_path, filename

@app.route("/")
def index():
    return render_template("index.html")

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
            # 파일 데이터를 미리 읽어두기 (멀티스레드 안전)
            file_data_list = []
            for file in files:
                if not file.filename.endswith(".docx"):
                    continue
                file_data_list.append((file.read(), file.filename, rules, tmpdir))

            # 병렬 처리 (CPU 코어 수만큼 동시 실행)
            with ThreadPoolExecutor() as executor:
                results = list(executor.map(process_file, file_data_list))

            # ZIP으로 묶기
            zip_path = os.path.join(tmpdir, "replaced_files.zip")
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for output_path, filename in results:
                    zipf.write(output_path, filename)

            return send_file(zip_path, as_attachment=True, download_name="replaced_files.zip")

if __name__ == "__main__":
    app.run(debug=True)
