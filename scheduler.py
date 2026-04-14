import pandas as pd
import io

def sanitize(text):
    if pd.isna(text):
        return ""
    return str(text).strip().upper()

def generate_timetable(partial_tt_path, course_teacher_path, periods_per_day, working_days):
    # --- 1. DATA INGESTION ---
    df_ct = pd.read_excel(course_teacher_path) if course_teacher_path.endswith('.xlsx') else pd.read_csv(course_teacher_path)
    df_ct.columns = df_ct.columns.str.strip()
    
    # DYNAMIC COLUMN MAPPING
    column_map = {}
    for col in df_ct.columns:
        col_upper = col.upper()
        if "TEACHER" in col_upper: column_map['Teacher'] = col
        elif "COURSE" in col_upper and "TYPE" not in col_upper: column_map['Course'] = col
        elif "CLASS" in col_upper: column_map['Class'] = col
        elif "TYPE" in col_upper: column_map['Type'] = col
        elif "PERIOD" in col_upper or "HOUR" in col_upper: column_map['Hours'] = col

    # Apply the mapping
    df_ct = df_ct.rename(columns={v: k for k, v in column_map.items()})

    # Sanitize inputs with fallbacks for missing columns
    for col in ['Course', 'Class', 'Teacher']:
        if col in df_ct.columns:
            df_ct[col] = df_ct[col].apply(sanitize)
    
    if 'Type' not in df_ct.columns:
        df_ct['Type'] = 'L'
    else:
        df_ct['Type'] = df_ct['Type'].apply(sanitize).replace('', 'L')

    classes = list(df_ct['Class'].unique()) if 'Class' in df_ct.columns else []
    # ... rest of the logic remains the same ...

def generate_class_excel(TT, classes, periods_per_day, working_days):
    output = io.BytesIO()
    all_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    days = all_days[:working_days]
    periods = [f"Period {i+1}" for i in range(periods_per_day)]
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for class_id in classes:
            grid = {p: [] for p in periods}
            for day_idx, day in enumerate(days):
                for p_idx, p in enumerate(periods):
                    global_slot = (day_idx * periods_per_day) + p_idx 
                    course = TT.get(class_id, {}).get(global_slot, "-")
                    grid[p].append(course)
            
            df = pd.DataFrame(grid, index=days)
            df.index.name = f"Class: {class_id}"
            df.to_excel(writer, sheet_name=str(class_id)[:31])
    output.seek(0)
    return output

def generate_teacher_excel(TT, TS, classes, teachers, periods_per_day, working_days):
    output = io.BytesIO()
    all_days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    days = all_days[:working_days]
    
    columns = []
    global_slots_to_check = []
    for day_idx, day in enumerate(days):
        for p_idx in range(periods_per_day):
            columns.append(f"{day}-P{p_idx+1}")
            global_slots_to_check.append((day_idx * periods_per_day) + p_idx)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        master_data = []
        for class_id in classes:
            row_data = [TT.get(class_id, {}).get(slot, "-") for slot in global_slots_to_check]
            master_data.append(row_data)
        df_master = pd.DataFrame(master_data, index=classes, columns=columns)
        df_master.index.name = "Classes"
        df_master.to_excel(writer, sheet_name="Master Class Matrix")

        faculty_data = []
        for teacher in teachers:
            row_data = []
            total_periods = 0
            for slot in global_slots_to_check:
                assigned_class = TS.get(teacher, {}).get(slot, "-")
                if assigned_class not in ["-", "BUSY"]:
                    total_periods += 1
                row_data.append(assigned_class)
            row_data.append(total_periods)
            faculty_data.append(row_data)
            
        faculty_columns = columns + ["Total Periods / Week"]
        df_faculty = pd.DataFrame(faculty_data, index=teachers, columns=faculty_columns)
        df_faculty.index.name = "Faculty"
        df_faculty.to_excel(writer, sheet_name="Faculty Workload")

    output.seek(0)
    return output