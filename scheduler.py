import pandas as pd
import io
import math

def sanitize(text):
    if pd.isna(text): return ""
    return str(text).strip().upper()

def generate_timetable(partial_tt_path, course_teacher_path, periods_per_day, working_days):
    # --- 1. DATA INGESTION ---
    if course_teacher_path.endswith('.xlsx'):
        df_ct = pd.read_excel(course_teacher_path)
    else:
        df_ct = pd.read_csv(course_teacher_path)
    
    df_ct.columns = df_ct.columns.str.strip()
    
    column_map = {}
    for col in df_ct.columns:
        col_upper = col.upper()
        if "TEACHER" in col_upper: column_map['Teacher'] = col
        elif "COURSE" in col_upper and "TYPE" not in col_upper: column_map['Course'] = col
        elif "CLASS" in col_upper: column_map['Class'] = col
        elif "TYPE" in col_upper: column_map['Type'] = col
        elif any(k in col_upper for k in ["PERIOD", "HOUR", "MAXIMUM"]): column_map['Hours'] = col

    df_ct = df_ct.rename(columns={v: k for k, v in column_map.items()})
    
    for col in ['Course', 'Class']:
        if col in df_ct.columns: df_ct[col] = df_ct[col].apply(sanitize)
    
    df_ct['Type'] = df_ct['Type'].apply(sanitize).replace('', 'L') if 'Type' in df_ct.columns else 'L'

    classes = list(df_ct['Class'].unique())
    CHMap = {} 
    CTMap = {} 
    teachers_set = set() 
    
    for _, row in df_ct.iterrows():
        key = (row['Class'], row['Course'])
        if key not in CHMap:
            CHMap[key] = int(row.get('Hours', 4)) if pd.notna(row.get('Hours')) else 4
            CTMap[key] = {'type': row['Type'], 'teachers': []}
        
        raw_teachers = str(row['Teacher']).split(',')
        for t in raw_teachers:
            clean_teacher = sanitize(t)
            if clean_teacher:
                teachers_set.add(clean_teacher)
                if clean_teacher not in CTMap[key]['teachers']:
                    CTMap[key]['teachers'].append(clean_teacher)

    teachers = sorted(list(teachers_set))
    TT = {c: {} for c in classes} 
    TS = {t: {} for t in teachers} 
    
    # Process Partial (Locked) Slots
    try:
        df_partial = pd.read_excel(partial_tt_path) if partial_tt_path.endswith('.xlsx') else pd.read_csv(partial_tt_path)
        # Note: Pre-fill logic for lab preferences could be added here
    except Exception: pass

    # --- 2. MAXIMALLY DISTANT SLOT-FILLING ALGORITHM ---
    max_slots = working_days * periods_per_day 

    for (class_id, course), total_hrs in CHMap.items():
        course_info = CTMap[(class_id, course)]
        is_simultaneous = "O" in course_info['type'] or "OTHER" in course_info['type']
        assigned_teachers = course_info['teachers']
        
        # Calculate Interval for Maximally Distant distribution 
        if total_hrs > 1:
            interval = max(1, max_slots // total_hrs)
        else:
            interval = max_slots // 2

        for i in range(total_hrs):
            # Target initial distant positions [cite: 268]
            target_slot = (i * interval) % max_slots
            
            # Move Left/Right logic to find the closest free slot 
            found = False
            for offset in range(max_slots):
                # Check target, then +1, -1, +2, -2...
                for direction in [1, -1]:
                    check_slot = (target_slot + offset * direction) % max_slots
                    
                    # Constraint check: Class Free
                    if TT[class_id].get(check_slot) is not None: continue
                    
                    # Constraint check: No consecutive periods of same course [cite: 125, 153]
                    is_consecutive = False
                    if check_slot > 0 and (check_slot % periods_per_day != 0):
                        if TT[class_id].get(check_slot - 1) == course: is_consecutive = True
                    if check_slot < max_slots - 1 and ((check_slot + 1) % periods_per_day != 0):
                        if TT[class_id].get(check_slot + 1) == course: is_consecutive = True
                    if is_consecutive: continue

                    # Constraint check: Teacher Availability [cite: 153, 172]
                    selected_teacher = None
                    if is_simultaneous:
                        if all(TS[t].get(check_slot) is None for t in assigned_teachers):
                            selected_teacher = "ALL"
                    else:
                        # Find the teacher with fewest hours so far to balance workload
                        assigned_teachers.sort(key=lambda t: sum(1 for s in TS[t].values() if s != 'BUSY'))
                        for t in assigned_teachers:
                            if TS[t].get(check_slot) is None:
                                selected_teacher = t
                                break
                    
                    if selected_teacher:
                        TT[class_id][check_slot] = course
                        if selected_teacher == "ALL":
                            for t in assigned_teachers: TS[t][check_slot] = class_id
                        else:
                            TS[selected_teacher][check_slot] = class_id
                        found = True
                        break
                if found: break

    return TT, TS, {'classes': classes, 'teachers': teachers, 'periods': periods_per_day, 'working_days': working_days}

def generate_class_excel(TT, classes, periods_per_day, working_days):
    output = io.BytesIO()
    all_days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    days = all_days[:working_days]
    periods = [f"Period {i+1}" for i in range(periods_per_day)]
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for class_id in classes:
            grid = {p: [] for p in periods}
            for d_idx in range(working_days):
                for p_idx in range(periods_per_day):
                    slot = (d_idx * periods_per_day) + p_idx
                    grid[f"Period {p_idx+1}"].append(TT.get(class_id, {}).get(slot, "-"))
            
            df = pd.DataFrame(grid, index=days)
            df.to_excel(writer, sheet_name=str(class_id)[:31])
    output.seek(0)
    return output

def generate_teacher_excel(TT, TS, classes, teachers, periods_per_day, working_days):
    output = io.BytesIO()
    all_days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    days = all_days[:working_days]
    
    columns = [f"{d}-P{p+1}" for d in days for p in range(periods_per_day)]
    global_slots = [(d_idx * periods_per_day) + p_idx for d_idx in range(working_days) for p_idx in range(periods_per_day)]

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Part 1: Master Class Matrix
        master_data = []
        for c in classes:
            master_data.append([TT.get(c, {}).get(s, "-") for s in global_slots])
        pd.DataFrame(master_data, index=classes, columns=columns).to_excel(writer, sheet_name="Master Class Matrix")

        # Part 2: Faculty Workload Matrix (Ensures every teacher is listed)
        faculty_data = []
        for t in teachers:
            row = []
            count = 0
            for s in global_slots:
                val = TS.get(t, {}).get(s, "-")
                if val not in ["-", "BUSY"]: count += 1
                row.append(val)
            row.append(count)
            faculty_data.append(row)
            
        pd.DataFrame(faculty_data, index=teachers, columns=columns + ["Total Periods"]).to_excel(writer, sheet_name="Faculty Workload")

    output.seek(0)
    return output