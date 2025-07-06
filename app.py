from flask import Flask, render_template, request, send_file
import fitz  # PyMuPDF
import os
import re
import openpyxl
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def extract_event_time_from_text(text):
    # 按行分割
    lines = text.splitlines()
    for line in lines:
        if "GMT" in line:
            # 尝试匹配日期时间格式
            m = re.search(r"(\w{3} \w{3} \d{1,2} \d{4} \d{2}:\d{2}:\d{2})\s*GMT([+-]\d{4})", line)
            if m:
                dt_part = m.group(1)
                tz_part = m.group(2)
                dt_str = f"{dt_part} {tz_part}"
                try:
                    dt_obj = datetime.strptime(dt_str, "%a %b %d %Y %H:%M:%S %z")
                    return dt_obj.strftime("%m/%d/%y %H:%M")
                except Exception as e:
                    # print("Datetime parse error:", e)
                    return ""
    return ""

def extract_cpsc_lines(pdf_path):
    records = []
    with fitz.open(pdf_path) as doc:
        full_text = ""
        for page in doc:
            full_text += page.get_text()

        entry_match = re.search(r'Entry # (\d+)', full_text)
        if not entry_match:
            return []
        entry_raw = entry_match.group(1).strip()
        entry_number = f"NVB-{entry_raw[:7]}-{entry_raw[-1]}"

        event_time = extract_event_time_from_text(full_text)
        print("Event Time:", event_time)  # Uncomment for debugging
        blocks = full_text.split("Line#")
        for block in blocks:
            if "Gov Agency: CPS" in block:
                line_match = re.search(r'^(\d+)', block.strip())
                if line_match:
                    records.append({
                        "entry": entry_number,
                        "status": "CPSC_check",
                        "event_time": event_time,
                        "timezone": "America/New_York",
                        "line": line_match.group(1)
                    })
    return records

def generate_excel(data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CPSC Check"

    headers = ["Entry Number", "Status", "Event Time", "Time Zone", "Line"]
    ws.append(headers)

    for item in data:
        ws.append([
            item["entry"],
            item["status"],
            item["event_time"],
            item["timezone"],
            item["line"]
        ])

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max_len + 2

    output = BytesIO()
    today_str = datetime.now().strftime("%Y%m%d")
    filename = f"CPSC_Check_Results_{today_str}.xlsx"
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        uploaded_files = request.files.getlist("pdfs")
        all_records = []
        for file in uploaded_files:
            if file and file.filename.endswith('.pdf'):
                filepath = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(filepath)
                all_records += extract_cpsc_lines(filepath)
        # ✅ 自动删除 uploads 下所有文件
        for file in os.listdir(UPLOAD_FOLDER):
            file_path = os.path.join(UPLOAD_FOLDER, file)
            try:
                os.remove(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}: {e}")
        if all_records:
            return generate_excel(all_records)
    return render_template("index.html")

import os

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)