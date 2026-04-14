import pandas as pd
import io

def sanitize(text):
    if pd.isna(text):
        return ""
    return str(text).strip().upper()

def generate_timetable(partial_tt_path, course_teacher_path, periods_per_day, working_days):
    # --- 1. DATA INGESTION ---
    # Load Course-Teacher mapping
    if course_teacher_path.endswith('.xlsx'):
        df_ct = pd.read_excel(course_teacher_path)
    else:
        df_ct = pd.read_csv(course_teacher_path)
    
    # Strip invisible spaces from raw headers
    df_ct.columns = df_ct.columns.str.strip()
    
    # DYNAMIC COLUMN MAPPING: Find your descriptive headers by keywords
    column_map = {}
    for col in df_ct.columns:
        col_upper = col.upper()
        if "TEACHER" in col_upper: 
            column_map['Teacher'] = col
        elif "COURSE" in col_upper and "TYPE" not in col_upper: 
            column_map['Course'] = col
        elif "CLASS" in col_upper: 
            column_map['Class'] = col
        elif "TYPE" in col_upper: 
            column_map['Type'] = col
        elif "PERIOD" in col_upper or "HOUR" in col_upper or "MAXIMUM" in col_upper: 
            column_map['Hours'] = col

    # Rename your long headers to standard names used by the algorithm
    df_ct = df_ct.rename(columns={v: k for k, v in column_map.items()})
    
    # Sanitize standard columns
    for col in ['Course', 'Class']:
        if col in df_ct.columns:
            df_ct[col] = df_ct[col].apply(sanitize)
    
    # Handle 'Type' column specifically
    if 'Type' not in df_ct.columns:
        df_ct['Type'] = 'L'
    else:
        df_ct['Type'] = df_ct['Type'].apply(sanitize).replace('', 'L')

    classes = list(df_ct['Class'].unique()) if 'Class' in df_ct.columns else []
    
    # Build Requirement Maps and parse teachers (Handling Scenario 2: Comma separation)
    CHMap = {} 
    CTMap = {} 
    teachers_set = set() 
    
    for _, row in df_ct.iterrows():
        # Ensure we have the minimum required data
        if 'Class' not in df_ct.columns or 'Course' not in df_ct.columns:
            continue
            
        key = (row['Class'], row['Course'])
        if key not in CHMap:
            # Default to 4 hours if the 'Hours' column is missing or empty
            CHMap[key] = int(row.get('Hours', 4)) if pd.notna(row.get('Hours')) else 4
            CTMap[key] = {'type': row['Type'], 'teachers': []}
        
        # Split teachers separated by commas (your Scenario 2)
        if 'Teacher' in df_ct.columns:
            raw_teachers = str(row['Teacher']).split(',')
            for t in raw_teachers:
                clean_teacher = sanitize(t)
                if clean_teacher:
                    teachers_set.add(clean_teacher)
                    if clean_teacher not in CTMap[key]['teachers']:
                        CTMap[key]['teachers'].append(clean_teacher)

    teachers = list(teachers_set)
    
    # Initialize Matrices
    TT = {c: {} for c in classes} 
    TS = {t: {} for t in teachers} 
    
    # Process Partially Filled TT
    try:
        if partial_tt_path.endswith('.xlsx'):
            df_partial = pd.read_excel(partial_tt_path)
        else:
            df_partial = pd.read_csv(partial_tt_path)
        # Note: In a production version, you would map df_partial to TT and TS here
    except Exception as e:
        print(f"Warning: Could not process partial timetable: {e}")

    # --- 2. SLOT-FILLING ALGORITHM ---
    max_slots = working_days * periods_per_day 

    for (class_id, course), hours in CHMap.items():
        course_info = CTMap[(class_id, course)]
        course_type = course_info['type']
        assigned_teachers = course_info['teachers']
        
        remaining_hrs = hours
        
        for slot in range(max_slots):
            if remaining_hrs <= 0:
                break
                
            # Constraint 1: Class must be free
            if TT[class_id].get(slot) is not None:
                continue
                
            # Constraint 2: No consecutive classes of same subject (on same day)
            if slot > 0 and (slot % periods_per_day != 0) and TT[class_id].get(slot - 1) == course:
                continue
                
            if course_type == 'O' or "OTHER" in course_type:
                # Type 'O': ALL teachers must be free simultaneously
                teachers_free = all(TS.get(t, {}).get(slot) is None for t in assigned_teachers)
                if teachers_free:
                    TT[class_id][slot] = course
                    for t in assigned_teachers:
                        TS[t][slot] = class_id
                    remaining_hrs -= 1
            else:
                # Type 'L': Only one teacher needs to be free
                for t in assigned_teachers:
                    if TS.get(t, {}).get(slot) is None:
                        TT[class_id][slot] = course
                        TS[t][slot] = class_id
                        remaining_hrs -= 1
                        break 

    # Return values for the Flask app to unpack
    return TT, TS, {'classes': classes, 'teachers': teachers, 'periods': periods_per_day, 'working_days': working_days}

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