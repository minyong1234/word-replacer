from flask import Flask, request, send_file, render_template, jsonify
from replacer import replace_in_docx
import os, uuid, json, tempfile
 
app = Flask(__name__)
 
@app.route("/")
def index():
    return render_template("index.html")
 
@app.route("/replace", methods=["POST"])
def replace():
    if "file" not in request.files:
        return jsonify({"error": "파일이 없습니다"}), 400
 
    file = request.files["file"]
    rules_json = request.form.get("rules", "{}")
    rules = json.loads(rules_json)
 
    if not file.filename.endswith(".docx"):
        return jsonify({"error": ".docx 파일만 지원합니다"}), 400
 
    if not rules:
        return jsonify({"error": "치환 규칙을 입력해주세요"}), 400
 
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path  = os.path.join(tmpdir, "input.docx")
        output_path = os.path.join(tmpdir, "output.docx")
        file.save(input_path)
        replace_in_docx(input_path, output_path, rules)
        return send_file(output_path, as_attachment=True, download_name="replaced.docx")
 
if __name__ == "__main__":
    app.run(debug=True)
