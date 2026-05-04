import os
import glob
from app import parse_work_duration_report, find_report_files
from database import init_db, save_report_to_db

def migrate():
    print("Initializing Database...", flush=True)
    init_db()
    
    reports = find_report_files()
    print(f"Found {len(reports)} reports to migrate.", flush=True)
    
    for r in reports:
        filename = r['filename']
        filepath = r['path']
        print(f"Processing {filename}...", flush=True)
        try:
            report_data = parse_work_duration_report(filepath)
            save_report_to_db(report_data, filename)
            print(f"Successfully migrated {filename}", flush=True)
        except Exception as e:
            print(f"Failed to migrate {filename}: {e}", flush=True)

if __name__ == "__main__":
    migrate()
