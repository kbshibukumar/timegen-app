import pandas as pd
import io
import math

def sanitize(text):
    """Sanitizes input by stripping spaces and converting to uppercase."""
    if pd.isna(text):
        return ""
    return str(text).strip().upper()

def generate_timetable(partial_tt_path, course_teacher_path, periods_per_day):
    # --- 1. DATA INGESTION ---
    # Load Course-Teacher mapping
    df_ct = pd.read_excel(course_teacher_path) if course_teacher_path.endswith('.xlsx') else pd.read_csv(course_teacher_path)
    
    # Sanitize inputs
    df_ct['Teacher'] = df_ct['Teacher'].apply(sanitize)
    df_ct['Course'] = df_ct['Course'].apply(sanitize)
    df_ct['Class'] = df_ct['Class'].apply(sanitize)
    df_ct['Type'] = df_ct['Type'].apply(sanitize)

    classes = list(df_ct['Class'].unique())
    teachers = list(df_ct['Teacher'].unique())
    
    # Initialize Matrices
    TT = {c: {} for c in classes} # TT[class_id][global_slot] = course
    TS = {t: {} for t in teachers} # TS[teacher][global_slot] = class_id or 'BUSY'
    
    # Process Partially Filled TT
    # (Assuming a basic format where rows are classes and columns are global slots for simplicity in this skeleton)
    try:
        df_partial = pd.read_excel(partial_tt_path) if partial_tt_path.endswith('.xlsx') else pd.read_csv(partial_tt_path)
        # Add logic here to populate TT and TS with pre-filled laboratory or preferred slots
    except Exception as e:
        print(f"Warning: Could not process partial timetable completely: {e}")

    # Build Requirement Maps
    CHMap = {} # (Class, Course) -> Total Hours
    CTMap = {} # (Class, Course) -> {'type': 'L'/'O', 'teachers': []}
    
    for _, row in df_ct.iterrows():
        key = (row['Class'], row['Course'])
        if key not in CHMap:
            # Assuming an 'Hours' column exists, defaulting to 4 if not
            CHMap[key] = int(row.get('Hours', 4)) if pd.notna(row.get('Hours')) else 4
            CTMap[key] = {'type': row['Type'], 'teachers': []}
        
        if row['Teacher'] not in CTMap[key]['teachers']:
            CTMap[key]['teachers'].append(row['Teacher'])

    # --- 2. SLOT-FILLING ALGORITHM (Skeleton with your constraints) ---
    max_slots = 5 * periods_per_day # 5 working days

    for (class_id, course), hours in CHMap.items():
        course_info = CTMap[(class_id, course)]
        course_type = course_info['type']
        assigned_teachers = course_info['teachers']
        
        remaining_hrs = hours
        
        for slot in range(max_slots):
            if remaining_hrs <= 0:
                break
                
            # Check if class is free
            if TT[class_id].get(slot) is not None:
                continue
                
            # Check consecutive periods (simplified check for same day)
            if slot > 0 and (slot % periods_per_day != 0) and TT[class_id].get(slot - 1) == course:
                continue
                
            # Check Teacher Availability
            if course_type == 'O':
                # Simultaneous: All teachers must be free
                teachers_free = all(TS[t].get(slot) is None for t in assigned_teachers)
                if teachers_free:
                    TT[class_id][slot] = course
                    for t in assigned_teachers:
                        TS[t][slot] = class_id
                    remaining_hrs -= 1
            else:
                # Lecture: Find one free teacher (round-robin or lowest hours)
                for t in assigned_teachers:
                    if TS[t].get(slot) is None:
                        TT[class_id][slot] = course
                        TS[t][slot] = class_id
                        remaining_hrs -= 1
                        break # Move to next hour requirement

    return TT, TS, {'classes': classes, 'teachers': teachers, 'periods': periods_per_day}

def generate_class_excel(TT, classes, periods_per_day):
    output = io.BytesIO()
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
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
            df.to_excel(writer, sheet_name=str(class_id)[:31]) # Excel sheet names max 31 chars
    output.seek(0)
    return output

def generate_teacher_excel(TT, TS, classes, teachers, periods_per_day):
    output = io.BytesIO()
    days = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    
    columns = []
    global_slots_to_check = []
    for day_idx, day in enumerate(days):
        for p_idx in range(periods_per_day):
            columns.append(f"{day}-P{p_idx+1}")
            global_slots_to_check.append((day_idx * periods_per_day) + p_idx)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Part 1: Master Class Matrix
        master_data = []
        for class_id in classes:
            row_data = [TT.get(class_id, {}).get(slot, "-") for slot in global_slots_to_check]
            master_data.append(row_data)
        df_master = pd.DataFrame(master_data, index=classes, columns=columns)
        df_master.index.name = "Classes"
        df_master.to_excel(writer, sheet_name="Master Class Matrix")

        # Part 2: Faculty Workload Matrix
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