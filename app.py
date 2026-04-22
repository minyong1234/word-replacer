from flask import Flask, request, send_file, render_template, jsonify
from replacer import replace_in_docx
import os, json, tempfile, zipfile

app = Flask(__name__)

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
            # 파일 1개면 바로 다운로드
            file = files[0]
            if not file.filename.endswith(".docx"):
                return jsonify({"error": ".docx 파일만 지원합니다"}), 400
            input_path = os.path.join(tmpdir, file.filename)
            output_path = os.path.join(tmpdir, f"output_{file.filename}")
            file.save(input_path)
            replace_in_docx(input_path, output_path, rules)
            return send_file(output_path, as_attachment=True, download_name=file.filename)
        else:
            # 파일 여러 개면 ZIP으로 묶어서 다운로드
            zip_path = os.path.join(tmpdir, "replaced_files.zip")
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for file in files:
                    if not file.filename.endswith(".docx"):
                        continue
                    input_path = os.path.join(tmpdir, file.filename)
                    output_path = os.path.join(tmpdir, f"output_{file.filename}")
                    file.save(input_path)
                    replace_in_docx(input_path, output_path, rules)
                    zipf.write(output_path, file.filename)
            return send_file(zip_path, as_attachment=True, download_name="replaced_files.zip")

if __name__ == "__main__":
    app.run(debug=True)
