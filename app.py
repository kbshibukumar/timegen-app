from flask import Flask, request, render_template, jsonify, send_file
import os
import json
import pandas as pd
import io
from werkzeug.utils import secure_filename
from scheduler import generate_timetable, generate_class_excel, generate_teacher_excel, sanitize

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

app_state = {
    'TT': {},
    'TS': {},
    'metadata': {}
}

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/parse_metadata', methods=['POST'])
def parse_metadata():
    try:
        course_file = request.files.get('course_teacher')
        partial_file = request.files.get('partial_tt')
        working_days = int(request.form.get('working_days', 5))
        periods_per_day = int(request.form.get('periods_per_day', 8))
        
        teachers = set()
        if course_file and course_file.filename:
            df_ct = pd.read_excel(course_file) if course_file.filename.endswith('.xlsx') else pd.read_csv(course_file)
            df_ct.columns = df_ct.columns.str.strip()
            t_col = next((c for c in df_ct.columns if 'TEACHER' in c.upper()), None)
            if t_col:
                for val in df_ct[t_col].dropna():
                    for t in str(val).split(','):
                        clean_t = sanitize(t)
                        if clean_t: teachers.add(clean_t)
        
        period_labels = [f"P{i+1}" for i in range(periods_per_day)]
        if partial_file and partial_file.filename:
            df_p = pd.read_excel(partial_file, header=None) if partial_file.filename.endswith('.xlsx') else pd.read_csv(partial_file, header=None)
            if len(df_p) > 1:
                row_2 = [sanitize(str(x)) if pd.notna(x) else "" for x in df_p.iloc[1].values]
                extracted = [x for x in row_2[1:] if x]
                for i in range(min(len(extracted), periods_per_day)):
                    period_labels[i] = extracted[i]
        
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"][:working_days]
        
        return jsonify({
            'teachers': sorted(list(teachers)),
            'period_labels': period_labels,
            'days': days
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate', methods=['POST'])
def generate():
    if 'partial_tt' not in request.files or 'course_teacher' not in request.files:
        return jsonify({'error': 'Missing files'}), 400

    partial_file = request.files['partial_tt']
    course_file = request.files['course_teacher']
    
    periods_per_day = int(request.form.get('periods_per_day', 7))
    working_days = int(request.form.get('working_days', 5))
    
    unsuitable_slots_str = request.form.get('unsuitable_slots', '[]')
    unsuitable_slots = json.loads(unsuitable_slots_str)

    partial_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(partial_file.filename))
    course_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(course_file.filename))
    
    partial_file.save(partial_path)
    course_file.save(course_path)

    try:
        TT, TS, metadata = generate_timetable(partial_path, course_path, periods_per_day, working_days, unsuitable_slots)
        app_state['TT'] = TT
        app_state['TS'] = TS
        app_state['metadata'] = metadata

        warnings = metadata.get('warnings', [])
        return jsonify({'message': 'Timetable generated successfully!', 'warnings': warnings})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/class_timetable')
def download_class_tt():
    if not app_state['TT']: return "No timetable generated yet", 400
    excel_file = generate_class_excel(app_state['TT'], app_state['metadata']['classes'], app_state['metadata']['periods'], app_state['metadata']['working_days'], app_state['metadata'].get('period_labels'))
    return send_file(excel_file, download_name="Class_TimeTable.xlsx", as_attachment=True)

@app.route('/download/teacher_timetable')
def download_teacher_tt():
    if not app_state['TT']: return "No timetable generated yet", 400
    excel_file = generate_teacher_excel(app_state['TT'], app_state['TS'], app_state['metadata']['classes'], app_state['metadata']['teachers'], app_state['metadata']['periods'], app_state['metadata']['working_days'], app_state['metadata'].get('period_labels'))
    return send_file(excel_file, download_name="Teacher_TimeTable.xlsx", as_attachment=True)

# --- NEW: Serve User-Provided Template Files ---

@app.route('/download/sample_course_teacher')
def download_sample_ct():
    try:
        # Looks inside the 'static' folder for your exact Excel file
        file_path = os.path.join(app.root_path, 'static', 'Sample_Course_Teacher.xlsx')
        return send_file(file_path, as_attachment=True)
    except FileNotFoundError:
        return "Template file not found on the server. Please ensure 'Sample_Course_Teacher.xlsx' is in the static folder.", 404

@app.route('/download/sample_partial_tt')
def download_sample_partial():
    try:
        # Looks inside the 'static' folder for your exact Excel file
        file_path = os.path.join(app.root_path, 'static', 'Sample_Partially_Filled.xlsx')
        return send_file(file_path, as_attachment=True)
    except FileNotFoundError:
        return "Template file not found on the server. Please ensure 'Sample_Partially_Filled.xlsx' is in the static folder.", 404

if __name__ == '__main__':
    app.run(debug=True)