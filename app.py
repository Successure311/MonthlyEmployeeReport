from flask import Flask, jsonify, send_from_directory, request, send_file
from flask_cors import CORS
import xlrd
import os
import glob
import re
import io
import xlsxwriter
from datetime import datetime
from werkzeug.utils import secure_filename
from process_attendance import process_report
# No DB imports

app = Flask(__name__, static_folder='static', static_url_path='/static')
CORS(app)

# No DB initialization

DATA_DIR = os.path.dirname(os.path.abspath(__file__))

OFFICIAL_HOLIDAYS_2026 = [
    "2026-01-15", "2026-01-26", "2026-03-03", "2026-03-26", "2026-03-31",
    "2026-04-03", "2026-04-14", "2026-05-01", "2026-05-28", "2026-06-26",
    "2026-09-14", "2026-10-02", "2026-10-20", "2026-11-10", "2026-11-24",
    "2026-12-25"
]


def parse_time_to_minutes(time_str):
    """Convert HH:MM time string to total minutes."""
    if not time_str or time_str == '00:00':
        return 0
    try:
        parts = time_str.strip().split(':')
        return int(parts[0]) * 60 + int(parts[1])
    except (ValueError, IndexError):
        return 0


def minutes_to_hhmm(minutes):
    """Convert total minutes to refined string format."""
    if not minutes or minutes == 0:
        return "00:00"
    h = int(minutes) // 60
    m = int(minutes) % 60
    if h == 0:
        return f"{m} min"
    return f"{h:02d}:{m:02d} h"


def parse_work_duration_report(filepath):
    """
    Parse WorkDurationReport .xls file.
    The file has a structured format with blocks per employee:
    - Employee row: 'Employee:', 'ID : Name', 'Total Work Duration: ... summary ...'
    - Status row: 'Status', day1_status, day2_status, ...
    - InTime row: 'InTime', day1_in, day2_in, ...
    - OutTime row: 'OutTime', day1_out, day2_out, ...
    - Duration row: 'Duration', day1_dur, day2_dur, ...
    - Late By row: 'Late By', day1_late, ...
    - Early By row: 'Early By', day1_early, ...
    - OT row: 'OT', ...
    - Shift row: 'Shift', day1_shift, ...
    """
    wb = xlrd.open_workbook(filepath)
    employees = []

    date_range_str = ""
    month_name = ""
    year_str = ""
    num_days = 31  # default

    for sheet_idx in range(wb.nsheets):
        ws = wb.sheet_by_index(sheet_idx)

        # Get month range from row 2/3 if not already found
        if not date_range_str:
            for r in range(min(10, ws.nrows)):
                for c in range(ws.ncols):
                    val = str(ws.cell_value(r, c)).strip()
                    if 'To' in val and any(yr in val for yr in ['2026', '2025', '2024']):
                        date_range_str = val
                        break
                if date_range_str:
                    parts = date_range_str.split('To')
                    if len(parts) == 2:
                        end_part = parts[1].strip()
                        try:
                            # Fix for possible multiple spaces
                            end_date_str = re.sub(r'\s+', ' ', end_part.strip())
                            # Try parsing typically formatted "Jan 31 2024" or extracting year directly
                            end_date = datetime.strptime(end_date_str, "%b %d %Y")
                            num_days = end_date.day
                            month_name = end_date.strftime("%B")
                            year_str = end_date.strftime("%Y")
                        except ValueError:
                            # Fallback: Extract year using regex if standard strptime fails
                            year_match = re.search(r'\b(202\d)\b', end_date_str)
                            if year_match:
                                year_str = year_match.group(1)
                            month_match = re.search(r'\b([A-Z][a-z]{2,8})\b', end_date_str, re.IGNORECASE)
                            if month_match:
                                # just take the first 3 letters capitalized
                                month_name = month_match.group(1)[:3].capitalize()
                    break

        # ---- Detect weekend/Saturday/Sunday columns ONCE per sheet (row index 6) ----
        sat_cols = []
        sun_cols = []
        col_to_date = {}
        if ws.nrows > 6:
            for c in range(1, ws.ncols):
                val_str = str(ws.cell_value(6, c)).strip().lower()
                if not val_str:
                    continue
                match = re.search(r'(\d+)', val_str)
                if match:
                    col_to_date[c] = int(match.group(1))
                # Saturday detection: "st", "sa", "sat"
                if re.search(r'\bst\b|\bsa\b|\bsat\b', val_str):
                    sat_cols.append(c)
                # Sunday detection: ends with " s" or is just "s" or contains "su"/"sun"
                elif re.search(r'\bsu\b|\bsun\b', val_str) or val_str.endswith(' s') or re.match(r'^\d+\s+s$', val_str):
                    sun_cols.append(c)

        report_saturday_count = len(sat_cols)
        report_sunday_count = len(sun_cols)
        print(f"Sheet {sheet_idx}: sat_cols={sat_cols[:5]}, sun_cols={sun_cols[:5]}, Sat={report_saturday_count}, Sun={report_sunday_count}", flush=True)

        r = 0
        while r < ws.nrows:
            # Check all columns in this row for "Employee:" marker
            row_first = str(ws.cell_value(r, 0)).strip()
            is_emp_row = row_first.startswith('Employee') and ':' in row_first

            if is_emp_row:
                # Find employee info and summary by scanning all columns
                emp_info = ""
                summary_str = ""
                for c in range(1, ws.ncols):
                    cell_val = str(ws.cell_value(r, c)).strip()
                    if not cell_val:
                        continue
                    # Employee info contains "ID : Name" pattern
                    if re.match(r'\d+\s*:', cell_val) and not emp_info:
                        emp_info = cell_val
                    # Summary contains "Total Work Duration"
                    elif 'Total Work Duration' in cell_val:
                        summary_str = cell_val

                # Parse employee name
                emp_name = emp_info
                emp_id = ""
                if ':' in emp_info:
                    parts = emp_info.split(':', 1)
                    emp_id = parts[0].strip()
                    emp_name = parts[1].strip()

                # FILTER: Only take actual names, not numbers
                EXCLUDED_EMPLOYEES = ['Deendayal Sevag', 'Manish Jajoo', 'NAND KUMAR', 'deepak', 'ram', 'Dinesh Jii', 'Anuj']
                is_excluded = any(ex.lower() in emp_name.lower() for ex in EXCLUDED_EMPLOYEES)
                
                # Normalize employee names
                EMP_NAME_MAP = {
                    'Sagar': 'Sagar Limbasiya',
                    'Sagar T': 'Sagar Tarsariya'
                }
                for old_name, new_name in EMP_NAME_MAP.items():
                    if emp_name.lower() == old_name.lower():
                        emp_name = new_name
                        break
                
                if emp_name and not emp_name.isnumeric() and emp_name != emp_id and not is_excluded:
                    # Parse summary
                    summary = parse_summary(summary_str)

                    # Read next rows for daily data
                    daily_data = {}
                    labels_to_find = ['Status', 'InTime', 'OutTime',
                                      'Duration', 'Late By', 'Early By', 'OT', 'Shift']
                    

                    search_r = r + 1
                    while search_r < min(r + 10, ws.nrows):
                        label = str(ws.cell_value(search_r, 0)).strip()
                        if label in labels_to_find:
                            values = []
                            for c in range(1, ws.ncols): # parse all columns first
                                values.append(str(ws.cell_value(search_r, c)).strip())
                            daily_data[label] = values
                        search_r += 1

                    # Build per-day records
                    days = []
                    
                    status_row = daily_data.get('Status', [])
                    start_col_offset = 1 if len(status_row) > 0 and status_row[0] == '' else 0
                    
                    has_any_data = False
                    
                    ns_shift_count = 0

                    for c in range(start_col_offset, len(status_row)):
                        sheet_col_idx = c + 1
                        
                        # Only process columns that we mapped to a calendar day
                        if sheet_col_idx not in col_to_date:
                            continue
                            
                        calendar_day_num = col_to_date[sheet_col_idx]
                        
                        # Check if this date is a holiday
                        is_holiday = False
                        if month_name and year_str:
                            try:
                                mon_idx = datetime.strptime(month_name[:3], "%b").month
                                date_str_key = f"{year_str}-{str(mon_idx).zfill(2)}-{str(calendar_day_num).zfill(2)}"
                                if date_str_key in OFFICIAL_HOLIDAYS_2026:
                                    is_holiday = True
                            except:
                                pass

                        status = status_row[c] if c < len(status_row) else ''
                        in_time = daily_data.get('InTime', [''])[c] if c < len(daily_data.get('InTime', [])) else ''
                        out_time = daily_data.get('OutTime', [''])[c] if c < len(daily_data.get('OutTime', [])) else ''
                        duration = daily_data.get('Duration', [''])[c] if c < len(daily_data.get('Duration', [])) else ''
                        late_by = daily_data.get('Late By', [''])[c] if c < len(daily_data.get('Late By', [])) else ''
                        early_by = daily_data.get('Early By', [''])[c] if c < len(daily_data.get('Early By', [])) else ''
                        ot = daily_data.get('OT', [''])[c] if c < len(daily_data.get('OT', [])) else ''
                        shift = daily_data.get('Shift', [''])[c] if c < len(daily_data.get('Shift', [])) else ''

                        if is_holiday:
                            shift = "HO"

                        shift_upper = shift.strip().upper()
                        
                        # If SHIFT is GS and STATUS is A but either In or Out Time is present, change status to P
                        if shift_upper == 'GS' and status.strip().upper() == 'A' and (in_time.strip() or out_time.strip()):
                            status = 'P'
                        
                        # If this column index in the sheet is a weekend, skip it
                        if sheet_col_idx in sat_cols or sheet_col_idx in sun_cols:
                            continue

                        # Weekdays only from here for NS count
                        if shift_upper == 'NS':
                            ns_shift_count += 1
                        # Count GS with Status A and both In/Out Time missing as Absent
                        if shift_upper == 'GS' and status.upper() == 'A' and not in_time.strip() and not out_time.strip():
                            ns_shift_count += 1
                            
                        # Format date string: we have month/year from top of file
                        date_label = f"Day {calendar_day_num}"
                        if month_name and year_str and calendar_day_num <= num_days:
                            try:
                                date_obj = datetime.strptime(f"{year_str} {month_name[:3]} {calendar_day_num}", "%Y %b %d")
                                day_name = date_obj.strftime("%A")
                                short_month = date_obj.strftime("%b")
                                date_label = f"{short_month} {str(calendar_day_num).zfill(2)}, {day_name}"
                            except ValueError:
                                date_label = f"{month_name[:3]} {str(calendar_day_num).zfill(2)}"

                        # Handle forgotten punches
                        is_forgot_punch_in = False
                        is_forgot_punch_out = False
                        late_by = "00:00"
                        early_by = "00:00"
                        late_out = "00:00"
                        late_out_mins = 0
                        is_work_status = status.strip().upper() == 'P'
                        if is_work_status or in_time.strip() or out_time.strip():
                            # All employees have 08:45 in time and 17:15 out time
                            target_in_str, target_out_str = ('08:45', '17:15')

                            is_forgot_punch_in = False
                            is_forgot_punch_out = False

                            # If either punch is missing, treat the whole day as forgotten punch (fill both)
                            # Also handle case where Status is 'P' but both times are missing
                            if not in_time.strip() or not out_time.strip():
                                in_time = target_in_str
                                out_time = target_out_str
                                is_forgot_punch_in = True
                                is_forgot_punch_out = True

                            # Always recalculate duration from in/out time
                            if in_time.strip() and out_time.strip():
                                try:
                                    t_in = datetime.strptime(in_time.strip(), '%H:%M')
                                    t_out = datetime.strptime(out_time.strip(), '%H:%M')
                                    dur_mins = int((t_out - t_in).total_seconds() / 60)
                                    if dur_mins >= 0:
                                        duration = f"{dur_mins // 60:02d}:{dur_mins % 60:02d}"
                                except ValueError:
                                    pass

                            early_arrival_mins = 0
                            early_going_mins = 0
                            late_mins = 0
                            late_out_mins = 0

                            if in_time.strip():
                                try:
                                    t_actual_in = datetime.strptime(in_time.strip(), '%H:%M')
                                    t_target_in = datetime.strptime(target_in_str, '%H:%M')
                                    diff_in_mins = int((t_actual_in - t_target_in).total_seconds() / 60)

                                    if diff_in_mins > 0:
                                        late_mins = diff_in_mins
                                    elif diff_in_mins < 0:
                                        early_arrival_mins = abs(diff_in_mins)
                                except ValueError:
                                    pass

                            if out_time.strip():
                                try:
                                    t_actual_out = datetime.strptime(out_time.strip(), '%H:%M')
                                    t_target_out = datetime.strptime(target_out_str, '%H:%M')
                                    diff_out_mins = int((t_target_out - t_actual_out).total_seconds() / 60)

                                    if diff_out_mins > 0:
                                        early_going_mins = diff_out_mins
                                    elif diff_out_mins < 0:
                                        late_out_mins = abs(diff_out_mins)
                                except ValueError:
                                    pass

                            # Set Late By
                            late_by = f"{late_mins // 60:02d}:{late_mins % 60:02d}"

                            # Early By represents "Early Departure" (leaving before Out Time)
                            early_by = f"{early_going_mins // 60:02d}:{early_going_mins % 60:02d}"

                            # Late Out represents extra time seated beyond Out Time
                            late_out = f"{late_out_mins // 60:02d}:{late_out_mins % 60:02d}"




                        # Check if this day has actual attendance data
                        has_data = status.strip() or in_time.strip() or out_time.strip() or duration.strip()
                        if has_data:
                            has_any_data = True
                            
                        # If the entire column represents a clear day (not just empty due to weekends, but cleared due to exclusion)
                        # We still increment valid day count to keep the dates aligned, but maybe we shouldn't append it if it's an excluded employee.
                        # Wait, we decided to not append the employee at all if they have no data.
                        
                        # Only append to the days array if this day has data (this cleanly omits weekends from the UI)
                        if has_data:
                            days.append({
                                'day': date_label,
                                'status': status,
                                'inTime': in_time,
                                'outTime': out_time,
                                'duration': duration,
                                'lateBy': late_by,
                                'earlyBy': early_by,
                                'lateOut': late_out,
                                'lateOutMinutes': late_out_mins,
                                'ot': ot,
                                'shift': shift,
                                'isForgotPunchIn': is_forgot_punch_in,
                                'isForgotPunchOut': is_forgot_punch_out
                            })
                    
                    # Only append employee if they actually have data
                    if has_any_data:
                        employees.append({
                            'id': emp_id,
                            'name': emp_name,
                            'summary': summary,
                            'days': days,
                            'saturdayCount': report_saturday_count,
                            'sundayCount': report_sunday_count,
                            'nsShiftCount': ns_shift_count
                        })
            r += 1

                    
    # Removed overwriting days since we extract them correctly now.
    

    return {
        'month': month_name,
        'year': year_str,
        'dateRange': date_range_str,
        'numDays': num_days,
        'employees': employees
    }


def parse_summary(summary_str):
    """Parse the summary string from WorkDurationReport."""
    result = {}
    patterns = {
        'totalWorkDuration': r'Total Work Duration:\s*([\d:]+)\s*Hrs',
        'totalOT': r'Total OT:\s*([\d:]+)\s*Hrs',
        'present': r'Present:\s*(\d+)',
        'absent': r'Absent:\s*(\d+)',
        'weeklyOff': r'WeeklyOff:\s*(\d+)',
        'holidays': r'Holidays:\s*(\d+)',
        'leavesTaken': r'Leaves Taken:\s*(\d+)',
        'lateByHrs': r'Late By Hrs:\s*([\d:]+)',
        'lateByDays': r'Late By Days:\s*(\d+)',
        'earlyByHrs': r'Early By Hrs:\s*([\d:]+)',
        'earlyGoingByDays': r'Early going By Days:\s*(\d+)',
        'totalDurationPlusOT': r'Total Duration\(\+OT\):\s*([\d:]+)',
        'avgWorkingHrs': r'Average Working Hrs:\s*([\d:]+)',
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, summary_str)
        if match:
            result[key] = match.group(1)
    return result


def compute_analysis(report_data):
    """Compute attendance analysis from parsed report data."""
    employees_analysis = []

    for emp in report_data['employees']:
        present_count = 0
        absent_count = 0
        wo_count = 0
        total_duration_min = 0
        total_late_min = 0
        total_early_min = 0
        late_count = 0
        early_departure_count = 0
        on_time_count = 0
        total_late_out_min = 0
        max_late_out_min = 0
        late_out_day_count = 0
        half_day_count = 0

        all_days = []
        for day in emp['days']:
            is_half_day = False
            status = day['status'].strip().upper()
            shift = day.get('shift', '').strip().upper()
            
            # Basic status counting
            if shift == 'HO':
                wo_count += 1
            elif status == 'P':
                present_count += 1
            elif status == 'A':
                absent_count += 1
            elif status in ['WO', 'OFF']:
                # These are weekly offs but not explicit holidays
                # They are already counted in saturdayCount/sundayCount elsewhere if applicable
                pass

            # Half Day logic: In time > 09:00 OR Out time < 15:30
            in_time_str = day.get('inTime', '').strip()
            out_time_str = day.get('outTime', '').strip()
            
            if (in_time_str and in_time_str != '-') or (out_time_str and out_time_str != '-'):
                try:
                    limit_in = datetime.strptime('09:00', '%H:%M')
                    limit_out = datetime.strptime('15:30', '%H:%M')
                    
                    is_late_in = False
                    is_early_out = False
                    
                    if in_time_str and in_time_str != '-':
                        t_in = datetime.strptime(in_time_str, '%H:%M')
                        if t_in > limit_in:
                            is_late_in = True
                            
                    if out_time_str and out_time_str != '-':
                        t_out = datetime.strptime(out_time_str, '%H:%M')
                        if t_out < limit_out:
                            is_early_out = True
                    
                    if is_late_in or is_early_out:
                        half_day_count += 1
                        is_half_day = True
                except ValueError:
                    pass
            else:
                is_half_day = False

            # Exclude HO shift from all time-based metrics
            if shift == 'HO':
                all_days.append({
                    'day': day['day'],
                    'status': day['status'],
                    'inTime': day['inTime'],
                    'outTime': day['outTime'],
                    'duration': day['duration'],
                    'lateBy': day['lateBy'],
                    'earlyBy': day['earlyBy'],
                    'lateOut': day.get('lateOut', '00:00'),
                    'shift': day['shift'],
                    'isForgotPunchIn': day.get('isForgotPunchIn', False),
                    'isForgotPunchOut': day.get('isForgotPunchOut', False)
                })
                continue

            # Performance Metrics (Only for non-HO)
            dur_min = parse_time_to_minutes(day['duration'])
            total_duration_min += dur_min
            
            late_min = parse_time_to_minutes(day['lateBy'])
            total_late_min += late_min
            if late_min > 0:
                late_count += 1
            else:
                if status == 'P':
                    on_time_count += 1
                    
            early_min = parse_time_to_minutes(day['earlyBy'])
            total_early_min += early_min
            if early_min > 0:
                early_departure_count += 1

            lo_min = day.get('lateOutMinutes', 0)
            total_late_out_min += lo_min
            if lo_min > 0:
                late_out_day_count += 1
                if lo_min > max_late_out_min:
                    max_late_out_min = lo_min
            
            all_days.append({
                'day': day['day'],
                'status': day['status'],
                'inTime': day['inTime'],
                'outTime': day['outTime'],
                'duration': day['duration'],
                'lateBy': day['lateBy'],
                'earlyBy': day['earlyBy'],
                'shift': day['shift'],
                'lateOut': day.get('lateOut', '00:00'),
                'lateOutMinutes': day.get('lateOutMinutes', 0),
                'isForgotPunchIn': day.get('isForgotPunchIn', False),
                'isForgotPunchOut': day.get('isForgotPunchOut', False),
                'isHalfDay': is_half_day
            })

        avg_duration_min = total_duration_min / present_count if present_count > 0 else 0
        avg_late_min = total_late_min / late_count if late_count > 0 else 0
        avg_early_min = total_early_min / early_departure_count if early_departure_count > 0 else 0
        avg_late_out_min = total_late_out_min / late_out_day_count if late_out_day_count > 0 else 0

        # User refinement: Only count as Leave days if Shift == 'NS'
        ns_count = emp.get('nsShiftCount', 0)
        final_absent_count = ns_count

        print(f"compute_analysis: {emp['name']} sat={emp.get('saturdayCount')} sun={emp.get('sundayCount')}", flush=True)
        employees_analysis.append({
            'id': emp['id'],
            'name': emp['name'],
            'summary': emp['summary'],
            'presentCount': present_count,
            'absentCount': final_absent_count,
            'halfDayCount': half_day_count,
            'lateOutCount': late_out_day_count,
            'maxLateOut': minutes_to_hhmm(max_late_out_min),
            'avgLateOut': minutes_to_hhmm(int(avg_late_out_min)),
            'totalLateOutMinutes': total_late_out_min,
            'weeklyOffCount': wo_count,
            'lateCount': late_count,
            'earlyDepartureCount': early_departure_count,
            'onTimeCount': on_time_count,
            'saturdayCount': emp.get('saturdayCount', 0),
            'sundayCount': emp.get('sundayCount', 0),
            'nsShiftCount': emp.get('nsShiftCount', 0),
            'totalDuration': minutes_to_hhmm(total_duration_min),
            'totalDurationMinutes': total_duration_min,
            'avgDuration': minutes_to_hhmm(int(avg_duration_min)),
            'avgDurationMinutes': round(avg_duration_min, 2),
            'totalLateBy': minutes_to_hhmm(total_late_min),
            'totalLateMinutes': total_late_min,
            'avgLateBy': minutes_to_hhmm(int(avg_late_min)),
            'avgLateMinutes': round(avg_late_min, 2),
            'totalEarlyBy': minutes_to_hhmm(total_early_min),
            'totalEarlyMinutes': total_early_min,
            'avgEarlyBy': minutes_to_hhmm(int(avg_early_min)),
            'avgEarlyMinutes': round(avg_early_min, 2),
            'dailyData': all_days
        })

    return employees_analysis


def extract_month_year_from_file(filepath):
    """Quickly extract month and year from the top of the excel file."""
    try:
        wb = xlrd.open_workbook(filepath, logfile=open(os.devnull, 'w'))
        ws = wb.sheet_by_index(0)
        for r in range(min(10, ws.nrows)):
            for c in range(min(5, ws.ncols)):
                val = str(ws.cell_value(r, c)).strip()
                if 'To' in val and any(yr in val for yr in ['2027', '2026', '2025', '2024', '2023']):
                    parts = val.split('To')
                    if len(parts) == 2:
                        end_part = parts[1].strip()
                        end_date_str = re.sub(r'\s+', ' ', end_part.strip())
                        try:
                            # Try standard parsing
                            end_date = datetime.strptime(end_date_str, "%b %d %Y")
                            return f"{end_date.strftime('%b')}-{end_date.strftime('%Y')}"
                        except ValueError:
                            # Fallback regex
                            year_match = re.search(r'\b(202\d)\b', end_date_str)
                            year_str = year_match.group(1) if year_match else ''
                            month_match = re.search(r'\b([A-Z][a-z]{2,8})\b', end_date_str, re.IGNORECASE)
                            month_str = month_match.group(1)[:3].capitalize() if month_match else ''
                            if month_str and year_str:
                                return f"{month_str}-{year_str}"
                            elif month_str:
                                return month_str
    except Exception:
        pass
    return None

def find_report_files():
    """Find all WorkDurationReport files in the data directory."""
    pattern = os.path.join(DATA_DIR, 'WorkDurationReport*.xls*')
    files = glob.glob(pattern)
    result = []
    for f in files:
        basename = os.path.basename(f)
        
        # 1. Try to read exact month/year from file contents
        content_date = extract_month_year_from_file(f)
        
        if content_date:
            month_str = content_date
        else:
            # 2. Fallback to extracting from filename
            match = re.search(r'WorkDurationReport[\s_]+([A-Za-z]+)(?:[\s_-]+(\d{4}))?', basename, re.IGNORECASE)
            if match:
                month_part = match.group(1).capitalize()
                month_part = month_part[:3]
                year_part = match.group(2)
                if year_part:
                    month_str = f"{month_part}-{year_part}"
                else:
                    month_str = month_part
            else:
                month_str = basename
                
        result.append({'filename': basename, 'month': month_str, 'path': f})
    return result


# ==================== ROUTES ====================

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/static/<path:path>')
def send_static(path):
    return send_from_directory('static', path)


@app.route('/api/reports')
def get_reports():
    """Get list of available report files from disk."""
    files = find_report_files()
    return jsonify({'reports': files})


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle report file upload and save to DB."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and file.filename.endswith(('.xls', '.xlsx')):
        filename = secure_filename(file.filename)
        if not filename.startswith('WorkDurationReport'):
            filename = 'WorkDurationReport ' + filename
            
        filepath = os.path.join(DATA_DIR, filename)
        file.save(filepath)
        
        try:
            # Process the file in-place (cleanup Sat/Sun etc.)
            process_report(filepath, filepath)
            
            return jsonify({'message': 'File uploaded and processed', 'filename': filename})
        except Exception as e:
            return jsonify({'error': f'Operation failed: {str(e)}'}), 500
        finally:
            pass
            
    return jsonify({'error': 'Invalid file type. Only .xls and .xlsx allowed.'}), 400

@app.route('/api/analysis/<filename>')
def get_analysis(filename):
    """Fetch analysis from disk."""
    try:
        filepath = os.path.join(DATA_DIR, filename)
        if not os.path.exists(filepath):
            return jsonify({'error': f'Report not found on disk: {filename}'}), 404
            
        report_data = parse_work_duration_report(filepath)
            
        analysis = compute_analysis(report_data)
        return jsonify({
            'month': report_data['month'],
            'year': report_data.get('year', ''),
            'dateRange': report_data.get('dateRange', ''),
            'numDays': report_data.get('numDays', 31),
            'employees': analysis
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/employee/<filename>/<emp_name>')
def get_employee_detail(filename, emp_name):
    """Get detailed daily data for a specific employee."""
    filepath = os.path.join(DATA_DIR, filename)
    if not os.path.exists(filepath):
        return jsonify({'error': f'File not found: {filename}'}), 404

    try:
        report_data = parse_work_duration_report(filepath)
        analysis = compute_analysis(report_data)
        for emp in analysis:
            if emp['name'].lower() == emp_name.lower():
                return jsonify(emp)
        return jsonify({'error': f'Employee not found: {emp_name}'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/leaves')
def get_all_leaves():
    """Aggregate leaves across all report files."""
    try:
        files = find_report_files()
        all_leaves = {}
        # all_leaves = {
        #   "Employee Name": {
        #     "total": 5,
        #     "monthly": {"Jan-2026": 2, "Feb-2026": 3},
        #     "dates": ["Jan 05, 2026, Monday", ...]
        #   }
        # }
        
        all_months = []

        for f_info in files:
            filepath = f_info['path']
            month_key = f_info['month']
            if month_key not in all_months:
                all_months.append(month_key)
                
            report_data = parse_work_duration_report(filepath)
            
            for emp in report_data['employees']:
                name = emp['name']
                # Apply same name normalization
                EMP_NAME_MAP = {
                    'Sagar': 'Sagar Limbasiya',
                    'Sagar T': 'Sagar Tarsariya'
                }
                for old_name, new_name in EMP_NAME_MAP.items():
                    if name.lower() == old_name.lower():
                        name = new_name
                        break
                
                if name not in all_leaves:
                    all_leaves[name] = {
                        "total": 0,
                        "monthly": {},
                        "dates": []
                    }
                
                if month_key not in all_leaves[name]["monthly"]:
                    all_leaves[name]["monthly"][month_key] = 0
                
                for day in emp['days']:
                    shift = day.get('shift', '').strip().upper()
                    status = day.get('status', '').strip().upper()
                    in_time = day.get('inTime', '')
                    out_time = day.get('outTime', '')
                    # User refinement: Count as Leave day if Shift == 'NS' or Shift == 'GS' with Status 'A' and both In/Out Time missing
                    if shift == 'NS' or (shift == 'GS' and status == 'A' and not in_time.strip() and not out_time.strip()):
                        all_leaves[name]["total"] += 1
                        all_leaves[name]["monthly"][month_key] += 1
                        all_leaves[name]["dates"].append({
                            "date": day['day'],
                            "month": month_key
                        })

        # Convert to list for easier frontend handling
        result = []
        for name, data in all_leaves.items():
            result.append({
                "name": name,
                "total": data["total"],
                "monthly": data["monthly"],
                "dates": data["dates"]
            })
            
        # Sort by total leaves descending
        result.sort(key=lambda x: x['total'], reverse=True)

        return jsonify({
            "leaves": result,
            "months": all_months
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/api/export_excel', methods=['POST'])
def export_excel_api():
    try:
        report_data = request.json
        if not report_data:
            return jsonify({'error': 'No data provided'}), 400

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output)
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'bg_color': '#1f4e78',
            'font_color': 'white',
            'border': 1
        })
        
        subheader_format = workbook.add_format({
            'bold': True,
            'bg_color': '#d9e1f2',
            'border': 1
        })
        
        table_header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472c4',
            'font_color': 'white',
            'border': 1,
            'align': 'center'
        })
        
        cell_format = workbook.add_format({'border': 1, 'align': 'center'})
        forgot_punch_format = workbook.add_format({'border': 1, 'align': 'center', 'font_color': 'red', 'bold': True})
        
        status_p_format = workbook.add_format({'border': 1, 'align': 'center', 'bg_color': '#e2f0d9', 'font_color': '#375623'})
        status_a_format = workbook.add_format({'border': 1, 'align': 'center', 'bg_color': '#fce4d6', 'font_color': '#843c0c'})
        status_wo_format = workbook.add_format({'border': 1, 'align': 'center', 'bg_color': '#f2eefb', 'font_color': '#604195'})
        
        kpi_label_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#f2f2f2'})
        kpi_value_format = workbook.add_format({'border': 1, 'align': 'center'})
        
        ho_row_format = workbook.add_format({
            'bg_color': '#fff2cc', # Light yellow
            'border': 1,
            'align': 'center'
        })

        employees = report_data.get('employees', [])
        
        # Add Summary Sheet if exporting multiple employees
        if len(employees) > 1:
            summary_sheet = workbook.add_worksheet("Summary")
            summary_sheet.merge_range('A1:L1', f"Attendance Summary - {report_data.get('month')} {report_data.get('year')}", header_format)
            
            summary_headers = ['Employee', 'Present', 'Absent | Half', 'Holidays', 'Late Days', 'On Time', 'Total Hrs', 'Avg Hrs', 'Avg Late', 'Avg Early', 'Avg Late Out', 'Max Late Out']
            for col_idx, text in enumerate(summary_headers):
                summary_sheet.write(2, col_idx, text, table_header_format)
            
            for row_idx, emp in enumerate(employees):
                curr_row = row_idx + 3
                summary_sheet.write(curr_row, 0, emp['name'], cell_format)
                summary_sheet.write(curr_row, 1, emp.get('presentCount'), cell_format)
                summary_sheet.write(curr_row, 2, f"{emp.get('absentCount')} | {emp.get('halfDayCount')}", cell_format)
                summary_sheet.write(curr_row, 3, emp.get('weeklyOffCount'), cell_format)
                summary_sheet.write(curr_row, 4, emp.get('lateCount'), cell_format)
                summary_sheet.write(curr_row, 5, emp.get('onTimeCount'), cell_format)
                summary_sheet.write(curr_row, 6, emp.get('totalDuration'), cell_format)
                summary_sheet.write(curr_row, 7, emp.get('avgDuration'), cell_format)
                summary_sheet.write(curr_row, 8, emp.get('avgLateBy'), cell_format)
                summary_sheet.write(curr_row, 9, emp.get('avgEarlyBy'), cell_format)
                summary_sheet.write(curr_row, 10, emp.get('avgLateOut'), cell_format)
                summary_sheet.write(curr_row, 11, emp.get('maxLateOut'), cell_format)
            
            summary_sheet.set_column('A:A', 20)
            summary_sheet.set_column('B:F', 10)
            summary_sheet.set_column('G:L', 15)

            # Add Leave Summary Table to Summary Sheet
            leave_data = report_data.get('leaveData')
            if leave_data and leave_data.get('leaves'):
                start_row = len(employees) + 6
                summary_sheet.merge_range(f'A{start_row}:D{start_row}', "Employee Leave Summary (Consolidated)", header_format)
                
                l_months = leave_data.get('months', [])
                l_headers = ['Employee', 'Total Leaves'] + l_months
                
                # Write headers
                for col_idx, text in enumerate(l_headers):
                    summary_sheet.write(start_row + 1, col_idx, text, table_header_format)
                
                # Write data
                for row_idx, l_emp in enumerate(leave_data.get('leaves', [])):
                    curr_row = start_row + 2 + row_idx
                    summary_sheet.write(curr_row, 0, l_emp['name'], cell_format)
                    summary_sheet.write(curr_row, 1, l_emp['total'], cell_format)
                    for m_idx, m_name in enumerate(l_months):
                        summary_sheet.write(curr_row, 2 + m_idx, l_emp.get('monthly', {}).get(m_name, 0), cell_format)

        for emp in employees:
            sheet_name = re.sub(r'[\\/*?:\[\]]', '', emp['name'])[:31]
            worksheet = workbook.add_worksheet(sheet_name)
            
            # Write Header
            worksheet.merge_range('A1:I1', f"Attendance Report - {emp['name']}", header_format)
            worksheet.merge_range('A2:I2', f"Period: {report_data.get('month')} {report_data.get('year')}", subheader_format)
            
            # Write KPIs
            worksheet.write('A4', 'Metric', kpi_label_format)
            worksheet.write('B4', 'Value', kpi_label_format)
            worksheet.write('D4', 'Metric', kpi_label_format)
            worksheet.write('E4', 'Value', kpi_label_format)
            worksheet.write('G4', 'Metric', kpi_label_format)
            worksheet.write('H4', 'Value', kpi_label_format)
            
            kpis = [
                ('Present Days', emp.get('presentCount')),
                ('Absent | Half', f"{emp.get('absentCount')} | {emp.get('halfDayCount')}"),
                ('Holidays', emp.get('weeklyOffCount')),
                ('Late Arrival', f"{emp.get('lateCount')} ( Avg. {emp.get('avgLateBy')} )"),
                ('Early Departure', f"{emp.get('earlyDepartureCount')} ( Avg. {emp.get('avgEarlyBy')} )"),
                ('Work Hours', f"{emp.get('totalDuration')} ( Avg. {emp.get('avgDuration')} )"),
                ('Avg. Late Out', emp.get('avgLateOut')),
                ('Max Late Out', emp.get('maxLateOut')),
                ('On Time Days', emp.get('onTimeCount'))
            ]
            
            for i, (label, val) in enumerate(kpis):
                row = 4 + (i % 3)
                col = (i // 3) * 3
                worksheet.write(row, col, label, kpi_label_format)
                worksheet.write(row, col + 1, str(val), kpi_value_format)
            
            # Write Daily Table
            table_headers = ['Day', 'Status', 'In Time', 'Out Time', 'Duration', 'Late Arrival', 'Early Dep.', 'Late Out', 'Shift']
            start_row = 9
            for col_idx, text in enumerate(table_headers):
                worksheet.write(start_row, col_idx, text, table_header_format)
            
            daily_data_list = emp.get('dailyData', [])
            for row_idx, day in enumerate(daily_data_list):
                curr_row = start_row + 1 + row_idx
                in_time = day.get('inTime', '-')
                out_time = day.get('outTime', '-')
                has_punch = (in_time and in_time != '-') or (out_time and out_time != '-')
                
                is_ho_punched = day.get('shift', '').upper() == 'HO' and has_punch
                base_fmt = ho_row_format if is_ho_punched else cell_format
                
                worksheet.write(curr_row, 0, day.get('day'), base_fmt)
                
                status = day.get('status', '-')
                fmt = base_fmt
                if not is_ho_punched:
                    if status == 'P': fmt = status_p_format
                    elif status == 'A': fmt = status_a_format
                    elif status in ('WO', 'WOP', 'HO') or day.get('shift', '').upper() == 'HO': 
                        fmt = status_wo_format
                
                worksheet.write(curr_row, 1, status, fmt)
                worksheet.write(curr_row, 2, day.get('inTime'), base_fmt)
                worksheet.write(curr_row, 3, day.get('outTime'), base_fmt)
                worksheet.write(curr_row, 4, day.get('duration'), base_fmt)
                worksheet.write(curr_row, 5, day.get('lateBy'), base_fmt)
                worksheet.write(curr_row, 6, day.get('earlyBy'), base_fmt)
                worksheet.write(curr_row, 7, day.get('lateOut'), base_fmt)
                worksheet.write(curr_row, 8, day.get('shift'), base_fmt)
            
            # Add Employee Wise Leave Details
            leave_data = report_data.get('leaveData')
            if leave_data and leave_data.get('leaves'):
                emp_leaves = next((l for l in leave_data['leaves'] if l['name'] == emp['name']), None)
                if emp_leaves and emp_leaves.get('dates'):
                    l_start_row = start_row + len(daily_data_list) + 3
                    worksheet.merge_range(f'A{l_start_row}:C{l_start_row}', f"Leave Dates: {emp['name']} (Total {len(emp_leaves['dates'])})", header_format)
                    
                    l_headers = ['#', 'Date', 'Month Report'] # Keeping it simple for Excel
                    for col_idx, text in enumerate(l_headers):
                        worksheet.write(l_start_row, col_idx, text, table_header_format)
                    
                    for idx, d_info in enumerate(emp_leaves['dates']):
                        r = l_start_row + 1 + idx
                        worksheet.write(r, 0, idx + 1, cell_format)
                        worksheet.write(r, 1, str(d_info['date']), cell_format)
                        worksheet.write(r, 2, d_info['month'], cell_format)

            worksheet.set_column('A:A', 15)
            worksheet.set_column('B:D', 12)
            worksheet.set_column('E:H', 15)
            worksheet.set_column('I:I', 12)

        workbook.close()
        output.seek(0)
        
        filename = f"Attendance_Report_{report_data.get('month')}_{report_data.get('year')}.xlsx"
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        print(f"Export Error: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5009)
