import xlrd
import re
from datetime import datetime

def parse_work_duration_report(filepath):
    wb = xlrd.open_workbook(filepath)
    employees = []
    month_name = ""
    year_str = ""
    num_days = 31
    date_range_str = ""

    for sheet_idx in range(wb.nsheets):
        ws = wb.sheet_by_index(sheet_idx)
        
        # Detect weekend columns
        sat_cols = []
        sun_cols = []
        if ws.nrows > 6:
            for c in range(1, ws.ncols):
                val_str = str(ws.cell_value(6, c)).strip().lower()
                if not val_str: continue
                if any(x in val_str for x in ["sa", "st"]):
                    sat_cols.append(c)
                elif any(x in val_str for x in ["su", "sn", " s"]):
                    sun_cols.append(c)
                elif val_str.endswith(" s") or val_str == "s":
                    sun_cols.append(c)

        report_saturday_count = len(sat_cols)
        report_sunday_count = len(sun_cols)
        print(f"File: {filepath} - Sat: {report_saturday_count}, Sun: {report_sunday_count}")

# Test on all files
parse_work_duration_report("WorkDurationReport Jan.xls")
parse_work_duration_report("WorkDurationReport_Feb_2026.xls")
parse_work_duration_report("WorkDurationReport_March_2026.xls")
