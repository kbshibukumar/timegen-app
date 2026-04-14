import pandas as pd
import io
import re

def sanitize(text):
    if pd.isna(text): return ""
    s = str(text).strip().upper()
    if s == 'NAN': return ""
    # Remove floating .0 from pure numbers (e.g., "3.0" becomes "3")
    if s.endswith('.0'): s = s[:-2]
    # Remove spaces around specific separators like /, -, &
    s = re.sub(r'\s*([/&+-])\s*', r'\1', s)
    # Condense any remaining multiple spaces into a single space
    s = re.sub(r'\s+', ' ', s)
    return s

def generate_timetable(partial_tt_path, course_teacher_path, periods_per_day, working_days, unsuitable_slots=None):
    if course_teacher_path.endswith('.xlsx'): df_ct = pd.read_excel(course_teacher_path)
    else: df_ct = pd.read_csv(course_teacher_path)
    
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
    
    if 'Class' in df_ct.columns:
        df_ct['Class'] = df_ct['Class'].replace(r'^\s*$', pd.NA, regex=True).ffill()
        df_ct['Class'] = df_ct['Class'].apply(sanitize)
        
    for col in ['Course']:
        if col in df_ct.columns: df_ct[col] = df_ct[col].apply(sanitize)
    
    df_ct['Type'] = df_ct['Type'].apply(sanitize).replace('', 'L') if 'Type' in df_ct.columns else 'L'

    classes = list(df_ct['Class'].unique()) if 'Class' in df_ct.columns else []
    classes = [c for c in classes if c] 
    
    CHMap = {} 
    CTMap = {} 
    teachers_set = set() 
    
    for _, row in df_ct.iterrows():
        if 'Class' not in df_ct.columns or 'Course' not in df_ct.columns: continue
        class_id = row['Class']
        course = row['Course']
        if not class_id or not course: continue
        
        key = (class_id, course)
        if key not in CHMap:
            CHMap[key] = int(row.get('Hours', 4)) if pd.notna(row.get('Hours')) else 4
            CTMap[key] = {'type': row['Type'], 'teachers': []}
        
        if 'Teacher' in df_ct.columns:
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
    max_slots = working_days * periods_per_day
    warnings = []
    
    period_labels = [f"P{i+1}" for i in range(periods_per_day)]

    # Lock unsuitable slots for specified teachers right away
    if unsuitable_slots:
        for us in unsuitable_slots:
            t_name = sanitize(us.get('teacher'))
            d_idx = int(us.get('day'))
            p_idx = int(us.get('period'))
            global_slot = (d_idx * periods_per_day) + p_idx
            
            if t_name in TS and global_slot < max_slots:
                TS[t_name][global_slot] = 'BUSY'

    def teacher_has_adjacent_class(t, slot):
        is_adj = False
        if slot > 0 and (slot % periods_per_day != 0) and TS[t].get(slot - 1) is not None and TS[t].get(slot - 1) != 'BUSY':
            is_adj = True
        if slot < max_slots - 1 and ((slot + 1) % periods_per_day != 0) and TS[t].get(slot + 1) is not None and TS[t].get(slot + 1) != 'BUSY':
            is_adj = True
        return is_adj

    # --- PRE-PROCESS PARTIALLY FILLED TIMETABLE ---
    try:
        if partial_tt_path.endswith('.xlsx'): df_partial = pd.read_excel(partial_tt_path, header=None)
        else: df_partial = pd.read_csv(partial_tt_path, header=None)
        
        if len(df_partial) > 1:
            row_2 = [sanitize(str(x)) if pd.notna(x) else "" for x in df_partial.iloc[1].values]
            extracted_labels = [x for x in row_2[1:] if x]
            if len(extracted_labels) > 0:
                for i in range(min(len(extracted_labels), periods_per_day)):
                    period_labels[i] = extracted_labels[i]

        for _, row in df_partial.iterrows():
            row_vals = [sanitize(str(x)) if pd.notna(x) else "" for x in row.values]
            class_id, start_col = None, 0
            for i, val in enumerate(row_vals):
                if val in classes:
                    class_id = val; start_col = i + 1; break
                    
            if class_id:
                slot = 0
                for i in range(start_col, len(row_vals)):
                    if slot >= max_slots: break
                    val = row_vals[i]
                    if val:
                        TT[class_id][slot] = val
                        key = (class_id, val)
                        if key in CHMap:
                            CHMap[key] -= 1  
                            for t in CTMap[key]['teachers']: TS[t][slot] = class_id 
                    slot += 1
    except Exception as e: pass

    # --- MAXIMALLY DISTANT SLOT-FILLING WITH TWO-PASS SOFT CONSTRAINTS ---
    for (class_id, course), remaining_hrs in CHMap.items():
        if remaining_hrs <= 0: continue 
            
        course_info = CTMap[(class_id, course)]
        is_simultaneous = "O" in course_info['type'] or "OTHER" in course_info['type']
        assigned_teachers = course_info['teachers']
        
        interval = max(1, max_slots // remaining_hrs) if remaining_hrs > 1 else max_slots // 2

        for i in range(remaining_hrs):
            target_slot = (i * interval) % max_slots
            best_slot, best_teacher, is_perfect_match = None, None, False
            
            for offset in range(max_slots):
                directions = [1, -1] if offset > 0 else [1]
                for direction in directions:
                    check_slot = (target_slot + offset * direction) % max_slots
                    
                    if TT[class_id].get(check_slot) is not None: continue
                    
                    is_consecutive_course = False
                    if check_slot > 0 and (check_slot % periods_per_day != 0) and TT[class_id].get(check_slot - 1) == course: is_consecutive_course = True
                    if check_slot < max_slots - 1 and ((check_slot + 1) % periods_per_day != 0) and TT[class_id].get(check_slot + 1) == course: is_consecutive_course = True
                    if is_consecutive_course: continue

                    if is_simultaneous:
                        if all(TS[t].get(check_slot) is None for t in assigned_teachers):
                            if all(not teacher_has_adjacent_class(t, check_slot) for t in assigned_teachers):
                                best_slot, best_teacher, is_perfect_match = check_slot, "ALL", True
                                break
                            elif best_slot is None: best_slot, best_teacher = check_slot, "ALL"
                    else:
                        assigned_teachers.sort(key=lambda t: sum(1 for s in TS[t].values() if s not in ['BUSY', '-']))
                        for t in assigned_teachers:
                            if TS[t].get(check_slot) is None:
                                if not teacher_has_adjacent_class(t, check_slot):
                                    best_slot, best_teacher, is_perfect_match = check_slot, t, True
                                    break
                                elif best_slot is None: best_slot, best_teacher = check_slot, t
                                    
                    if is_perfect_match: break
                if is_perfect_match: break
            
            if best_slot is not None:
                TT[class_id][best_slot] = course
                if best_teacher ==