import xlrd
from xlutils.copy import copy
import datetime

def parse_time(time_str):
    if not time_str or time_str.strip() == "":
        return None
    try:
        return datetime.datetime.strptime(time_str.strip(), "%H:%M")
    except ValueError:
        return None

def format_duration(td):
    total_seconds = int(td.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{hours}:{minutes:02d}"

def process_report(input_file, output_file):
    rb = xlrd.open_workbook(input_file)
    sheet_read = rb.sheet_by_index(0)
    
    wb = copy(rb)
    sheet_write = wb.get_sheet(0)
    
    # Identify Saturday and Sunday columns from Row 7 (global default)
    global_sat_sun_cols = []
    # Row index 6 is the 7th row
    if sheet_read.nrows > 6:
        row_7 = [str(v).strip() for v in sheet_read.row_values(6)]
        for c in range(len(row_7)):
            day_str = row_7[c]
            # Detect Saturday/Sunday (Sa, Su, Sat, Sun, etc.)
            if any(day_str.endswith(s) for s in ["St", " S", "Sa", "Su", "Sat", "Sun"]):
                global_sat_sun_cols.append(c)
    
    print(f"Detected Saturday/Sunday columns: {global_sat_sun_cols}")

    # Custom thresholds
    DEFAULT_LATE_THRESHOLD = datetime.datetime.strptime("08:45", "%H:%M")
    DEFAULT_OUT_THRESHOLD = datetime.datetime.strptime("17:15", "%H:%M")

    SPECIAL_EMPLOYEES = ['Sagar T', 'Kalpan Shroff', 'Moin Shaikh', 'Sagar', 'Krunal']
    SPECIAL_OUT_THRESHOLD = datetime.datetime.strptime("17:15", "%H:%M")

    VATSAL_LATE_THRESHOLD = datetime.datetime.strptime("09:00", "%H:%M")
    VATSAL_OUT_THRESHOLD = datetime.datetime.strptime("17:15", "%H:%M")

    EXCLUDED_EMPLOYEES = ['Deendayal Sevag', 'Manish Jajoo', 'NAND KUMAR', 'deepak', 'ram', 'Dinesh Jii', 'Anuj']
    
    current_employee_format = "A" # Default to Format A
    current_employee_name = ""
    
    # Iterate through all rows
    for r in range(sheet_read.nrows):
        row_values = [str(sheet_read.cell_value(r, c)).strip() for c in range(sheet_read.ncols)]
        first_cell = row_values[0]
        
        # Detect employee and format
        if first_cell == "Employee:":
            # Check if name is at index 3 (Format A) or index 1 (Format B)
            if row_values[1] != "":
                current_employee_format = "B"
                current_employee_name = row_values[1]
            else:
                current_employee_format = "A"
                current_employee_name = row_values[3]
            
            print(f"Row {r}: Detected Employee {current_employee_name} - Format {current_employee_format}")
            continue

        # Skip logic: if current employee is excluded, clear all data rows for them
        is_excluded = any(ex.lower() in current_employee_name.lower() for ex in EXCLUDED_EMPLOYEES)
        
        active_sat_sun = global_sat_sun_cols
        
        # Determine thresholds for current employee
        late_threshold = DEFAULT_LATE_THRESHOLD
        out_threshold = DEFAULT_OUT_THRESHOLD
        
        if any(sp.lower() in current_employee_name.lower() for sp in SPECIAL_EMPLOYEES):
            out_threshold = SPECIAL_OUT_THRESHOLD
        elif 'vatsal' in current_employee_name.lower():
            late_threshold = VATSAL_LATE_THRESHOLD

        # 1. Remove Sat/Sun data OR clear all data for excluded employees
        if first_cell in ["Status", "InTime", "OutTime", "Duration", "Late By", "Early By", "OT", "Shift"]:
            if is_excluded:
                # Clear entire row for excluded employees
                for c in range(1, sheet_read.ncols):
                    sheet_write.write(r, c, "")
            else:
                # Just clear Sat/Sun
                for c in active_sat_sun:
                    if c < sheet_read.ncols:
                        sheet_write.write(r, c, "")
        
        # 2. Fix Late By and Early By logic
        if not is_excluded:
            if first_cell == "Late By":
                intime_row_idx = r - 3
                if intime_row_idx >= 0 and str(sheet_read.cell_value(intime_row_idx, 0)).strip() == "InTime":
                    start_col = 2 if current_employee_format == "A" else 1
                    for c in range(start_col, sheet_read.ncols):
                        if c in active_sat_sun:
                            continue
                            
                        intime_val = str(sheet_read.cell_value(intime_row_idx, c)).strip()
                        dt_intime = parse_time(intime_val)
                        
                        if dt_intime:
                            if dt_intime > late_threshold:
                                diff = dt_intime - late_threshold
                                sheet_write.write(r, c, format_duration(diff))
                            else:
                                sheet_write.write(r, c, "00:00")
                        else:
                            # Clear Late By if no InTime
                            sheet_write.write(r, c, "")
            
            if first_cell == "Early By":
                intime_row_idx = r - 4  # Early by is usually right below Late By, InTime is 4 above
                # Let's dynamically find InTime row index just above Early By
                search_idx = r - 1
                while search_idx >= 0:
                    if str(sheet_read.cell_value(search_idx, 0)).strip() == "InTime":
                        intime_row_idx = search_idx
                        break
                    search_idx -= 1
                    
                if intime_row_idx >= 0 and str(sheet_read.cell_value(intime_row_idx, 0)).strip() == "InTime":
                    start_col = 2 if current_employee_format == "A" else 1
                    for c in range(start_col, sheet_read.ncols):
                        if c in active_sat_sun:
                            continue
                            
                        intime_val = str(sheet_read.cell_value(intime_row_idx, c)).strip()
                        dt_intime = parse_time(intime_val)
                        
                        if dt_intime:
                            # Early arrival means getting there BEFORE the required InTime (late_threshold)
                            if dt_intime < late_threshold:
                                diff = late_threshold - dt_intime
                                sheet_write.write(r, c, format_duration(diff))
                            else:
                                sheet_write.write(r, c, "00:00")
                        else:
                            # Clear Early By if no InTime
                            sheet_write.write(r, c, "")

    wb.save(output_file)
    print(f"Processed report saved to {output_file}")

if __name__ == "__main__":
    process_report(r'E:\Monthly_Employe_Report\WorkDurationReport Jan.xls', 
                   r'E:\Monthly_Employe_Report\WorkDurationReport Jan.xls')
