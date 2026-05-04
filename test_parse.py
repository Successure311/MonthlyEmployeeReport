import json
from app import parse_work_duration_report, compute_analysis

# Run the parser on the currently 'cleaned' file
filepath = r'E:\Monthly_Employe_Report\WorkDurationReport Jan.xls'

try:
    data = parse_work_duration_report(filepath)
    analysis_data = compute_analysis(data)
    
    for emp in analysis_data:
        print(f"\nCounts for {emp['name']}:")
        print(f"  Saturday Count: {emp.get('saturdayCount')}")
        print(f"  Sunday Count: {emp.get('sundayCount')}")
        print(f"  Avg Late By: {emp.get('avgLateBy')}")
        print(f"  Avg Early By: {emp.get('avgEarlyBy')}")
        # Print first few days to see shifts and forgotten punches
        for d in emp.get('dailyData', [])[:15]:
             print(f"    {d['day']}: Shift={d.get('shift', '-')}, In={d.get('inTime')}, Out={d.get('outTime')}, Dur={d.get('duration')}, Forgot={d.get('isForgotPunchIn')}/{d.get('isForgotPunchOut')}")
        break

except Exception as e:
    print(f"Error parsing: {e}")
