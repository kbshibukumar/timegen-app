from flask import Flask, request, render_template, jsonify, send_file
import os
from werkzeug.utils import secure_filename
from scheduler import generate_timetable, generate_class_excel, generate_teacher_excel

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Global variables to hold state for downloads (suitable for single-user admin)
app_state = {
    'TT': {},
    'TS': {},
    'metadata': {}
}

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    if 'partial_tt' not in request.files or 'course_teacher' not in request.files:
        return jsonify({'error': 'Missing files'}), 400

    partial_file = request.files['partial_tt']
    course_file = request.files['course_teacher']
    periods_per_day = int(request.form.get('periods_per_day', 7))

    partial_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(partial_file.filename))
    course_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(course_file.filename))
    
    partial_file.save(partial_path)
    course_file.save(course_path)

    try:
        TT, TS, metadata = generate_timetable(partial_path, course_path, periods_per_day)
        
        # Save to global state for the download routes
        app_state['TT'] = TT
        app_state['TS'] = TS
        app_state['metadata'] = metadata

        return jsonify({'message': 'Timetable generated successfully!'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/class_timetable')
def download_class_tt():
    if not app_state['TT']:
        return "No timetable generated yet", 400
    excel_file = generate_class_excel(app_state['TT'], app_state['metadata']['classes'], app_state['metadata']['periods'])
    return send_file(excel_file, download_name="Class_TimeTable.xlsx", as_attachment=True)

@app.route('/download/teacher_timetable')
def download_teacher_tt():
    if not app_state['TT']:
        return "No timetable generated yet", 400
    excel_file = generate_teacher_excel(app_state['TT'], app_state['TS'], app_state['metadata']['classes'], app_state['metadata']['teachers'], app_state['metadata']['periods'])
    return send_file(excel_file, download_name="Teacher_TimeTable.xlsx", as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)