import mysql.connector
import os
import json
from dotenv import load_dotenv

load_dotenv()

def get_db_connection(use_db=True):
    # Priority: Environment Variables (Render) -> Credentials JSON (Local)
    host = os.getenv("TIDB_HOST")
    user = os.getenv("TIDB_USER")
    password = os.getenv("TIDB_PASSWORD")
    
    db_name = os.getenv("TIDB_DATABASE", "MonthlyEmployeeReport")
    port = int(os.getenv("TIDB_PORT", 4000))
    
    # SSL CA handle
    ssl_ca = os.getenv("TIDB_SSL_CA")
    if not ssl_ca:
        # Check current directory (root) first
        root_ca = os.path.join(os.path.dirname(__file__), "isrgrootx1.pem")
        # Check TiDB subdirectory as fallback
        tidb_ca = os.path.join(os.path.dirname(__file__), "TiDB", "isrgrootx1.pem")
        
        if os.path.exists(root_ca):
            ssl_ca = root_ca
        elif os.path.exists(tidb_ca):
            ssl_ca = tidb_ca

    if not host:
        # Try loading from JSON if env vars not set (local dev)
        cred_path = os.path.join(os.path.dirname(__file__), "TiDB", "TiDB_Credentials.json")
        if os.path.exists(cred_path):
            with open(cred_path, 'r') as f:
                creds = json.load(f).get("TiDB", {})
                host = creds.get("Host")
                user = creds.get("User")
                password = creds.get("Password")
                database = creds.get("Database", "MonthlyEmployeeReport")
                port = creds.get("Port", 4000)
                # Don't override ssl_ca if already found

    config = {
        "host": host,
        "user": user,
        "password": password,
        "port": port,
    }
    
    if use_db:
        config["database"] = db_name
    
    if ssl_ca:
        config["ssl_ca"] = ssl_ca

    return mysql.connector.connect(**config)

def init_db():
    # Connect without specific database to create it
    conn = get_db_connection(use_db=False)
    cursor = conn.cursor()
    
    db_name = os.getenv("TIDB_DATABASE", "MonthlyEmployeeReport")
    cursor.execute(f"CREATE DATABASE IF NOT EXISTS {db_name}")
    cursor.execute(f"USE {db_name}")
    
    # Reports table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS reports (
        id INT AUTO_INCREMENT PRIMARY KEY,
        filename VARCHAR(255) UNIQUE,
        month VARCHAR(50),
        year VARCHAR(10),
        uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # Employees table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS employees (
        id VARCHAR(50) PRIMARY KEY,
        name VARCHAR(255)
    )
    """)
    
    # Attendance table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS attendance (
        id INT AUTO_INCREMENT PRIMARY KEY,
        report_id INT,
        employee_id VARCHAR(50),
        date_label VARCHAR(100),
        status VARCHAR(50),
        in_time VARCHAR(20),
        out_time VARCHAR(20),
        duration VARCHAR(20),
        late_by VARCHAR(20),
        early_by VARCHAR(20),
        late_out VARCHAR(20),
        late_out_minutes INT,
        ot VARCHAR(20),
        shift VARCHAR(20),
        is_forgot_punch_in BOOLEAN,
        is_forgot_punch_out BOOLEAN,
        FOREIGN KEY (report_id) REFERENCES reports(id) ON DELETE CASCADE,
        FOREIGN KEY (employee_id) REFERENCES employees(id)
    )
    """)
    
    conn.commit()
    cursor.close()
    conn.close()
    print("Database initialized successfully.")

def save_report_to_db(report_data, filename):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        # 1. Insert/Get Report
        cursor.execute("INSERT INTO reports (filename, month, year) VALUES (%s, %s, %s) ON DUPLICATE KEY UPDATE month=%s, year=%s",
                       (filename, report_data['month'], report_data['year'], report_data['month'], report_data['year']))
        cursor.execute("SELECT id FROM reports WHERE filename = %s", (filename,))
        report_id = cursor.fetchone()[0]
        
        # Clear existing attendance for this report to avoid duplicates on re-upload
        cursor.execute("DELETE FROM attendance WHERE report_id = %s", (report_id,))
        
        for emp in report_data['employees']:
            # 2. Insert/Update Employee
            cursor.execute("INSERT INTO employees (id, name) VALUES (%s, %s) ON DUPLICATE KEY UPDATE name=%s",
                           (emp['id'], emp['name'], emp['name']))
            
            # 3. Insert Attendance
            for day in emp['days']:
                cursor.execute("""
                INSERT INTO attendance (
                    report_id, employee_id, date_label, status, in_time, out_time, 
                    duration, late_by, early_by, late_out, late_out_minutes, ot, 
                    shift, is_forgot_punch_in, is_forgot_punch_out
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    report_id, emp['id'], day['day'], day['status'], day['inTime'], day['outTime'],
                    day['duration'], day['lateBy'], day['earlyBy'], day['lateOut'], day.get('lateOutMinutes', 0),
                    day['ot'], day['shift'], day.get('isForgotPunchIn', False), day.get('isForgotPunchOut', False)
                ))
        
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        print(f"Error saving report to DB: {e}")
        raise e
    finally:
        cursor.close()
        conn.close()

def get_all_reports_from_db():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT filename, month, year FROM reports ORDER BY uploaded_at DESC")
    reports = cursor.fetchall()
    cursor.close()
    conn.close()
    return reports

def get_report_data_from_db(filename):
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    
    # Get report info
    cursor.execute("SELECT id, month, year FROM reports WHERE filename = %s", (filename,))
    report_info = cursor.fetchone()
    if not report_info:
        return None
    
    report_id = report_info['id']
    
    # Get employees
    cursor.execute("""
        SELECT DISTINCT e.id, e.name 
        FROM employees e
        JOIN attendance a ON e.id = a.employee_id
        WHERE a.report_id = %s
    """, (report_id,))
    employees = cursor.fetchall()
    
    for emp in employees:
        cursor.execute("""
            SELECT date_label as day, status, in_time as inTime, out_time as outTime, 
                   duration, late_by as lateBy, early_by as earlyBy, late_out as lateOut,
                   late_out_minutes as lateOutMinutes, ot, shift, 
                   is_forgot_punch_in as isForgotPunchIn, is_forgot_punch_out as isForgotPunchOut
            FROM attendance
            WHERE report_id = %s AND employee_id = %s
            ORDER BY id ASC
        """, (report_id, emp['id']))
        emp['days'] = cursor.fetchall()
        # We don't store summary directly, app.py can compute it or we can store it.
        # But compute_analysis in app.py handles it if we provide the 'days'.
        # Note: app.py expects 'summary' key in employee object (from Excel)
        # We'll just provide an empty summary or app.py will handle it.
        emp['summary'] = {} 
        
    cursor.close()
    conn.close()
    
    return {
        'month': report_info['month'],
        'year': report_info['year'],
        'employees': employees
    }

if __name__ == "__main__":
    init_db()
