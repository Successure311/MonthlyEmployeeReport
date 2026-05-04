import xlrd

def inspect_report(file_path):
    wb = xlrd.open_workbook(file_path)
    sheet = wb.sheet_by_index(0)
    
    print(f"File: {file_path}")
    print(f"Total Rows: {sheet.nrows}")
    print(f"Total Cols: {sheet.ncols}")
    
    # Print row 7 (index 6)
    if sheet.nrows > 6:
        row_7 = [str(sheet.cell_value(6, c)).strip() for c in range(sheet.ncols)]
        print("\nRow 7 (Days):")
        print(row_7)
    
    # Print first few rows to see header
    print("\nFirst 10 rows:")
    for r in range(min(10, sheet.nrows)):
        row_data = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
        print(f"Row {r}: {row_data}")

    # Find where employee data starts and how it's repeating
    print("\nEmployee data pattern (Rows 60-70):")
    for r in range(60, min(80, sheet.nrows)):
        row_data = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
        if any(row_data):
            print(f"Row {r}: {row_data}")

if __name__ == "__main__":
    inspect_report(r'E:\Monthly_Employe_Report\WorkDurationReport Jan.xls')
