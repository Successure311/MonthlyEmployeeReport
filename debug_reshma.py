import xlrd

def debug_reshma(file_path):
    rb = xlrd.open_workbook(file_path)
    sheet = rb.sheet_by_index(0)
    
    print(f"Row 80 (Employee): {[str(sheet.cell_value(80, c)) for c in range(sheet.ncols)]}")
    print(f"Row 81 (Status): {[str(sheet.cell_value(81, c)) for c in range(sheet.ncols)]}")

if __name__ == "__main__":
    debug_reshma(r'E:\Monthly_Employe_Report\WorkDurationReport Jan.xls')
