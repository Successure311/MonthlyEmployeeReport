import xlrd
import datetime

def find_weekday_late(file_path):
    rb = xlrd.open_workbook(file_path)
    sheet = rb.sheet_by_index(0)
    
    # Identify Sat/Sun columns
    sat_sun_cols = []
    row_7 = sheet.row_values(6)
    for c in range(2, len(row_7)):
        day_str = str(row_7[c]).strip()
        if day_str.endswith("St") or day_str.endswith(" S"):
            sat_sun_cols.append(c)
    
    threshold = datetime.datetime.strptime("08:45", "%H:%M")
    
    print(f"Searching for weekday late entries in {file_path}...")
    
    for r in range(sheet.nrows):
        if str(sheet.cell_value(r, 0)).strip() == "InTime":
            for c in range(2, sheet.ncols):
                if c in sat_sun_cols:
                    continue
                
                val = str(sheet.cell_value(r, c)).strip()
                if not val: continue
                
                try:
                    dt = datetime.datetime.strptime(val, "%H:%M")
                    if dt > threshold:
                        print(f"Found Late Entry: Row {r}, Col {c}, Time {val}")
                except ValueError:
                    pass

if __name__ == "__main__":
    find_weekday_late(r'E:\Monthly_Employe_Report\WorkDurationReport Jan.xls')
