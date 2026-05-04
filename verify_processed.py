import xlrd

def verify_report(file_path):
    wb = xlrd.open_workbook(file_path)
    sheet = wb.sheet_by_index(0)
    
    # Sat/Sun columns identified: [5, 6, 14, 15, 21, 22, 28, 29, 36]
    sat_sun_cols = [5, 6, 14, 15, 21, 22, 28, 29, 36]
    
    print(f"Verifying {file_path}")
    
    # Check Row 61-68 (Krisha) for Sat/Sun
    print("\nChecking Sat/Sun columns for Employee 17 (Krisha):")
    for r in range(61, 69):
        row_label = str(sheet.cell_value(r, 0)).strip()
        if row_label in ["Status", "InTime", "OutTime", "Duration", "Late By", "Early By", "OT", "Shift"]:
            vals = [str(sheet.cell_value(r, c)).strip() for c in sat_sun_cols]
            print(f"Row {r} ({row_label}): {vals}")
            if any(vals):
                print(f"FAILED: Found data in Sat/Sun columns for row {r}")
            else:
                print(f"PASSED: Sat/Sun columns are empty for row {r}")

    # Check Late By calculation for Krisha
    print("\nChecking Late By calculation for Krisha (Row 65):")
    # InTime is Row 62
    for c in range(2, sheet.ncols):
        if c in sat_sun_cols:
            continue
        
        intime = str(sheet.cell_value(62, c)).strip()
        late_by = str(sheet.cell_value(65, c)).strip()
        
        if intime:
            print(f"Col {c}: InTime={intime}, Late By={late_by}")
            # Check 08:35 (should be 00:00)
            if intime == "08:35" and late_by != "00:00":
                print(f"ERROR: InTime 08:35 should have Late By 00:00, got {late_by}")
            # Check 11:14 (should be 02:29)
            if intime == "11:14" and late_by != "2:29":
                print(f"ERROR: InTime 11:14 should have Late By 2:29, got {late_by}")

if __name__ == "__main__":
    verify_report(r'E:\Monthly_Employe_Report\WorkDurationReport Jan_Updated.xls')
