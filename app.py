import os
import io
import re
from difflib import SequenceMatcher
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
import docx

app = Flask(__name__, template_folder='templates')
CORS(app)

QUESTION_BANK = []

COMPETENCY_INDICATORS = {
    "HH1.1": {"keywords": ["h2so4", "đặc nóng", "so2", "fe", "oxi hóa khử"], "chapter": "Chương 1 - Đại cương Kim Loại"},
    "HH1.2": {"keywords": ["cân bằng", "n2", "3h2", "2nh3", "thuận nghịch", "áp suất"], "chapter": "Chương 2 - Tốc độ phản ứng & Cân bằng"},
    "HH2.1": {"keywords": ["este", "ch3cooc2h5", "xà phòng hóa", "hiệu suất", "glucozơ"], "chapter": "Chương 3 - Este - Lipit"}
}

def calculate_similarity(text1, text2):
    clean1 = re.sub(r'<[^>]+>', '', text1).strip().lower()
    clean2 = re.sub(r'<[^>]+>', '', text2).strip().lower()
    return SequenceMatcher(None, clean1, clean2).ratio()

def convert_run_to_html(run):
    text = run.text
    if not text: return ""
    if run.font.subscript: text = f"<sub>{text}</sub>"
    elif run.font.superscript: text = f"<sup>{text}</sup>"
    if run.bold: text = f"<b>{text}</b>"
    if run.italic: text = f"<i>{text}</i>"
    if run.font.underline: text = f"<u>{text}</u>"
    return text

def parse_docx_advanced(file_stream):
    doc = docx.Document(file_stream)
    parsed_questions = []
    current_q = None
    for p in doc.paragraphs:
        line_html = "".join([convert_run_to_html(r) for r in p.runs]).strip()
        if not line_html: continue
        if re.match(r'^(<b>)?Câu\s+\d+[:.]', line_html, re.IGNORECASE):
            if current_q: parsed_questions.append(current_q)
            current_q = {"content": line_html, "options": [], "short_answer": "", "explanation": "", "question_type": "mcq_4"}
        elif re.match(r'^(<b>)?[A-Da-d][.)]', line_html) and current_q:
            is_underline = "<u>" in line_html
            key_match = re.search(r'([A-Da-d])', line_html)
            key = key_match.group(1).upper() if key_match else "A"
            if key_match.group(1).islower(): current_q["question_type"] = "true_false"
            current_q["options"].append({"key": key, "content": line_html, "is_correct": is_underline})
        elif line_html.startswith("ĐS:") or line_html.startswith("<b>ĐS:</b>"):
            if current_q:
                current_q["question_type"] = "short_answer"
                current_q["short_answer"] = re.sub(r'^(<b>)?ĐS:\s*<\/b>?', '', line_html).strip()
        elif "HD:" in line_html or "Hướng dẫn giải" in line_html:
            if current_q: current_q["explanation"] = line_html
    if current_q: parsed_questions.append(current_q)
    return parsed_questions

def smart_tag_question(question_text):
    text_lower = question_text.lower()
    best_match_code = "HH1.1"
    matched_chapter = "Chương 1 - Đại cương"
    for code, info in COMPETENCY_INDICATORS.items():
        for kw in info["keywords"]:
            if kw in text_lower:
                best_match_code = code
                matched_chapter = info["chapter"]
                break
        if best_match_code == code: break
    bloom_level = "know"
    if any(w in text_lower for w in ["hiệu suất", "tính khối lượng", "phần trăm", "hỗn hợp"]): bloom_level = "apply"
    elif any(w in text_lower for w in ["cực trị", "tổng hợp", "biện luận"]): bloom_level = "high_apply"
    elif any(w in text_lower for w in ["giải thích", "nguyên nhân", "hiện tượng"]): bloom_level = "understand"
    return {"chapter_id": matched_chapter, "competency_code": best_match_code, "bloom_level": bloom_level, "difficulty_score": 3 if bloom_level in ["know", "understand"] else 4}

def audit_question_quality(question):
    warnings = []
    if question.get("question_type") == "mcq_4":
        if len(question.get("options", [])) < 4: warnings.append("Thiếu phương án lựa chọn.")
        correct_opts = [opt for opt in question.get("options", []) if opt.get("is_correct")]
        if len(correct_opts) == 0: warnings.append("Chưa gạch chân đáp án đúng.")
    if not question.get("explanation"): warnings.append("Thiếu hướng dẫn giải.")
    return {"is_valid": len(warnings) == 0, "warnings": warnings}

@app.route('/')
def index():
    return render_template('admin.html')

@app.route('/api/questions/import-word', methods=['POST'])
def import_word():
    file = request.files.get('file')
    if not file: return jsonify({"success": False, "error": "Chưa chọn file"}), 400
    questions = parse_docx_advanced(file)
    added_count = 0
    duplicate_warnings = []
    for q in questions:
        is_duplicate = False
        matched_q = None
        for existing in QUESTION_BANK:
            if calculate_similarity(q["content"], existing["content"]) >= 0.95:
                is_duplicate = True
                matched_q = existing
                break
        if is_duplicate:
            duplicate_warnings.append({"new": q["content"], "existing": matched_q["content"]})
        else:
            q["metadata"] = smart_tag_question(q["content"])
            q["quality_audit"] = audit_question_quality(q)
            QUESTION_BANK.append(q)
            added_count += 1
    return jsonify({"success": True, "added_count": added_count, "duplicates_found": len(duplicate_warnings), "duplicates": duplicate_warnings})

@app.route('/api/questions/export-word', methods=['GET'])
def export_word():
    doc = docx.Document()
    doc.add_heading('NGÂN HÀNG CÂU HỎI HÓA HỌC', level=1)
    for idx, q in enumerate(QUESTION_BANK):
        p = doc.add_paragraph()
        p.add_run(f"Câu {idx + 1}. ").bold = True
        p.add_run(re.sub(r'<[^>]+>', '', q['content']))
        for opt in q.get('options', []):
            opt_p = doc.add_paragraph()
            run = opt_p.add_run(f"{opt['key']}. {re.sub(r'<[^>]+>', '', opt['content'])}")
            if opt.get('is_correct'): run.underline = True
        if q.get('short_answer'): doc.add_paragraph(f"ĐS: {q['short_answer']}")
        if q.get('explanation'): doc.add_paragraph(f"HD: {re.sub(r'<[^>]+>', '', q['explanation'])}")
        doc.add_paragraph()
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return send_file(stream, as_attachment=True, download_name="Ngan_Hang_Xuat.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

if __name__ == '__main__':
    app.run(port=5000, debug=True)