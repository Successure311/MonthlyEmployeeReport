import os
import json
from app import find_report_files, parse_work_duration_report

def test():
    files = find_report_files()
    print(f"Found {len(files)} files")
    for f in files:
        print(f"File: {f['filename']}, Month: {f['month']}")
        data = parse_work_duration_report(f['path'])
        ns_count = 0
        for emp in data['employees']:
            for day in emp['days']:
                if day.get('shift', '').strip().upper() == 'NS':
                    ns_count += 1
        print(f"  NS shifts (leaves) found: {ns_count}")

if __name__ == "__main__":
    test()
