"""
STARLUX SOP Checker - Flask Backend
"""

import os
import json
import uuid
import shutil
from flask import Flask, request, jsonify, send_file, render_template_string
from werkzeug.utils import secure_filename
from checker import check_sop, apply_fixes

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

UPLOAD_FOLDER = "/home/claude/sop-checker/uploads"
OUTPUT_FOLDER = "/home/claude/sop-checker/outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# In-memory session store
sessions = {}


@app.route('/')
def index():
    with open('/home/claude/sop-checker/templates/index.html', 'r', encoding='utf-8') as f:
        return f.read()


@app.route('/api/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({"error": "未收到檔案"}), 400
    file = request.files['file']
    if not file.filename.endswith('.docx'):
        return jsonify({"error": "請上傳 .docx 格式的檔案"}), 400

    session_id = str(uuid.uuid4())
    filename = f"{session_id}_original.docx"
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        issues = check_sop(filepath)
    except Exception as e:
        return jsonify({"error": f"分析失敗：{str(e)}"}), 500

    sessions[session_id] = {
        "original_path": filepath,
        "issues": issues,
        "original_filename": secure_filename(file.filename),
    }

    return jsonify({
        "session_id": session_id,
        "total_issues": len(issues),
        "issues": issues,
    })


@app.route('/api/apply_fixes', methods=['POST'])
def apply_fixes_endpoint():
    data = request.json
    session_id = data.get("session_id")
    fix_ids = data.get("fix_ids", [])  # list of issue IDs to fix
    fix_all = data.get("fix_all", False)

    if session_id not in sessions:
        return jsonify({"error": "Session 不存在，請重新上傳檔案"}), 400

    session = sessions[session_id]
    issues = session["issues"]

    if fix_all:
        to_fix = [i for i in issues if i.get("auto_fixable")]
    else:
        to_fix = [i for i in issues if i["id"] in fix_ids and i.get("auto_fixable")]

    if not to_fix:
        return jsonify({"error": "選取的項目中沒有可自動修改的問題"}), 400

    output_filename = f"{session_id}_fixed.docx"
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    # Use previously fixed version if exists, else original
    source_path = session.get("current_fixed_path", session["original_path"])

    try:
        apply_fixes(source_path, to_fix, output_path)
    except Exception as e:
        return jsonify({"error": f"修改失敗：{str(e)}"}), 500

    session["current_fixed_path"] = output_path

    # Mark issues as fixed
    fixed_ids = {i["id"] for i in to_fix}
    for issue in issues:
        if issue["id"] in fixed_ids:
            issue["user_action"] = "fixed"

    remaining = [i for i in issues if i.get("user_action") != "fixed" and i.get("user_action") != "skip"]

    return jsonify({
        "success": True,
        "fixed_count": len(to_fix),
        "remaining_count": len(remaining),
        "issues": issues,
    })


@app.route('/api/skip_issue', methods=['POST'])
def skip_issue():
    data = request.json
    session_id = data.get("session_id")
    issue_id = data.get("issue_id")

    if session_id not in sessions:
        return jsonify({"error": "Session 不存在"}), 400

    for issue in sessions[session_id]["issues"]:
        if issue["id"] == issue_id:
            issue["user_action"] = "skip"
            break

    return jsonify({"success": True})


@app.route('/api/download/<session_id>/<format>', methods=['GET'])
def download(session_id, format):
    if session_id not in sessions:
        return jsonify({"error": "Session 不存在"}), 400

    session = sessions[session_id]
    source = session.get("current_fixed_path", session["original_path"])
    orig_name = session["original_filename"].replace(".docx", "")

    if format == "docx":
        return send_file(source,
                         as_attachment=True,
                         download_name=f"{orig_name}_checked.docx",
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    elif format == "pdf":
        try:
            import subprocess
            pdf_path = source.replace(".docx", ".pdf")
            subprocess.run([
                "python", "/mnt/skills/public/docx/scripts/office/soffice.py",
                "--headless", "--convert-to", "pdf",
                "--outdir", OUTPUT_FOLDER, source
            ], check=True, timeout=30)
            # soffice outputs to same dir with same name
            expected_pdf = os.path.join(OUTPUT_FOLDER,
                                        os.path.basename(source).replace(".docx", ".pdf"))
            if os.path.exists(expected_pdf):
                return send_file(expected_pdf,
                                 as_attachment=True,
                                 download_name=f"{orig_name}_checked.pdf",
                                 mimetype='application/pdf')
            else:
                return jsonify({"error": "PDF 轉換失敗，請下載 Word 版本後自行轉換"}), 500
        except Exception as e:
            return jsonify({"error": f"PDF 轉換失敗：{str(e)}。請下載 Word 版本。"}), 500

    return jsonify({"error": "不支援的格式"}), 400


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=7860, debug=False)
