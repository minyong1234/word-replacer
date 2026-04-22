from docx import Document
from docx.shared import RGBColor
import re

def replace_in_docx(input_path, output_path, replacements, case_sensitive=True):
    doc = Document(input_path)

    def replace_run(run):
        for old, new in replacements.items():
            if case_sensitive:
                if old in run.text:
                    run.text = run.text.replace(old, new)
                    run.font.color.rgb = RGBColor(0, 0, 0)  # 검은색
            else:
                if re.search(re.escape(old), run.text, flags=re.IGNORECASE):
                    run.text = re.sub(re.escape(old), new, run.text, flags=re.IGNORECASE)
                    run.font.color.rgb = RGBColor(0, 0, 0)  # 검은색

    # 일반 단락 처리
    for para in doc.paragraphs:
        for run in para.runs:
            replace_run(run)

    # 표 안의 텍스트 처리
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        replace_run(run)

    doc.save(output_path)
