import xlrd
import os
from process_attendance import process_report

def verify():
    input_file = r'e:\Monthly_Employe_Report\WorkDurationReport Jan.xls'
    output_file = r'e:\Monthly_Employe_Report\WorkDurationReport_Test.xls'
    
    if not os.path.exists(input_file):
        print(f"Input file not found: {input_file}")
        return

    print("Processing report...")
    process_report(input_file, output_file)
    
    print("Verifying results...")
    rb = xlrd.open_workbook(output_file)
    sheet = rb.sheet_by_index(0)
    
    EXCLUDED = ['Deendayal Sevag', 'Manish Jajoo', 'NAND KUMAR', 'deepak', 'ram']
    
    found_employees = []
    
    for r in range(sheet.nrows):
        first_cell = str(sheet.cell_value(r, 0)).strip()
        if first_cell == "Employee:":
            emp_name = str(sheet.cell_value(r, 1)).strip() or str(sheet.cell_value(r, 3)).strip()
            found_employees.append(emp_name)
            
            # Check if excluded employees have empty data rows
            is_excluded = any(ex.lower() in emp_name.lower() for ex in EXCLUDED)
            if is_excluded:
                # Check status row (r+1)
                for c in range(1, min(10, sheet.ncols)):
                    val = str(sheet.cell_value(r + 1, c)).strip()
                    if val:
                        print(f"ERROR: Excluded employee {emp_name} has data at row {r+1}, col {c}: {val}")
            
            # Check Vatsal's late by at a specific day if possible
            if 'vatsal' in emp_name.lower():
                # We need to find an InTime > 08:45 but < 09:00 if exists
                pass

    print(f"Found employees: {len(found_employees)}")
    for ex in EXCLUDED:
        if any(ex.lower() in name.lower() for name in found_employees):
            # They are still there as rows, but should have no data
            print(f"Verified: {ex} row exists but should be cleared.")
            
    # Detect Sat/Sun columns dynamically
    sat_sun_cols = []
    if sheet.nrows > 6:
        row_7 = [str(v).strip() for v in sheet.row_values(6)]
        for c in range(len(row_7)):
            if any(row_7[c].endswith(s) for s in ["Sa", "Su", "Sat", "Sun", "St", " S"]):
                sat_sun_cols.append(c)
    print(f"Verified Sat/Sun columns: {sat_sun_cols}")

    # Check Sat/Sun columns for an employee
    for r in range(sheet.nrows):
        label = str(sheet.cell_value(r, 0)).strip()
        if label in ["Status", "InTime", "OutTime"]:
            for c in sat_sun_cols:
                if c < sheet.ncols:
                    val = str(sheet.cell_value(r, c)).strip()
                    if val:
                        print(f"ERROR: Found data in Sat/Sun at row {r}, col {c}: {val}")

    print("Verification script finished.")

if __name__ == "__main__":
    verify()
