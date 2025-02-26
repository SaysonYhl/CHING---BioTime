import sqlite3
import datetime
import openpyxl
import os
import subprocess
from config import load_config

# time constants
am_start = datetime.time(9, 0, 0) # 9:00 am start time
am_late = datetime.time(9, 30, 0) # 9:30 am late start
am_absent = datetime.time(10, 15, 0) # 10:15 am absent
am_end = datetime.time(14, 0, 0) # 2:00 pm out time
am_latest_out = datetime.time(15, 29, 0) # 3:29 pm latest out

pm_start = datetime.time(15, 30, 0) # 3:30 pm start time
pm_late = datetime.time(16, 0, 0) # 4:00 pm late start
pm_absent = datetime.time(16, 31, 0) # 4:30 pm absent
pm_end = datetime.time(21, 0, 0) # 9:00 pm out time
pm_latest_out = datetime.time(23, 59, 0) # 11:59 pm latest out

def process_dates(start_date, end_date, excel_filename):
    try:
        # Load configuration before each process
        config = load_config()
        db_path = config["db_path"]
        report_directory = config["report_directory"]
        daily_salary = config["daily_salary"]
        deduction_per_minute = config["deduction_per_minute"]

        start_dt = datetime.datetime.strptime(start_date, "%Y-%m-%d")
        end_dt = datetime.datetime.strptime(end_date, "%Y-%m-%d")

        num_days = (end_dt - start_dt).days + 1
        total_shifts = num_days * 2

        gross_salary = daily_salary * num_days

        with sqlite3.connect(db_path) as conn:
            cursor = conn.cursor()
            query = """
                SELECT em.id, em.emp_firstname, em.emp_lastname, ap.punch_time
                FROM hr_employee em
                LEFT JOIN att_punches ap ON em.id = ap.employee_id
                WHERE (date(ap.punch_time) BETWEEN ? AND ?) AND (em.emp_privilege = 0)
                ORDER BY em.id, ap.punch_time;
            """
            cursor.execute(query, (start_date, end_date))
            results = cursor.fetchall()

            employee_attendance = {}
            for emp_id, first_name, last_name, punch_time_str in results:
                punch_time = datetime.datetime.strptime(punch_time_str, "%Y-%m-%d %H:%M:%S")

                if emp_id not in employee_attendance:
                    employee_attendance[emp_id] = {
                        "first_name": first_name,
                        "last_name": last_name,
                        "total_shifts": total_shifts,
                        "late": 0,
                        "absent": 0,
                        "deductions": 0,
                        "gross_salary": gross_salary,
                        "net_salary": 0,
                        "punches": []
                    }

                employee_attendance[emp_id]["punches"].append(punch_time)

            for emp_id, data in employee_attendance.items():
                status = check_attendance(data["punches"], deduction_per_minute, daily_salary, start_date, end_date)
                data["late"] = status["Late Minutes"]
                data["absent"] = status["Absent"]
                data["deductions"] = status["Deductions"]
                data["net_salary"] = gross_salary - data["deductions"]

            if not os.path.exists(report_directory):
                os.makedirs(report_directory)
            generate_excel(excel_filename, employee_attendance, start_date, end_date)
            print(f"Excel report generated successfully: {os.path.join(report_directory, excel_filename)}")

    except sqlite3.Error as e:
        print(f"Database error: {e}")
    except ValueError as e:
        print(f"Invalid input: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

def check_attendance(punches, deduction_per_minute, daily_salary, start_date, end_date):
    status = {"Late Minutes": 0, "Absent": 0, "Deductions": 0}
    punches.sort()

    # Generate a list of all dates in the range
    start_dt = datetime.datetime.strptime(start_date, "%Y-%m-%d").date()
    end_dt = datetime.datetime.strptime(end_date, "%Y-%m-%d").date()
    all_dates = [start_dt + datetime.timedelta(days=i) for i in range((end_dt - start_dt).days + 1)]

    for current_date in all_dates:
        am_shift = []
        pm_shift = []

        # Filter punches for the current date
        for punch in punches:
            if punch.date() == current_date:
                if am_start <= punch.time() <= am_latest_out:
                    am_shift.append(punch)
                elif pm_start <= punch.time() <= pm_latest_out:
                    pm_shift.append(punch)

        # AM Shift Check
        punch_in_am = next((p for p in am_shift if am_start <= p.time() < am_absent), None)
        punch_out_am = next((p for p in am_shift if am_end <= p.time() <= am_latest_out), None)

        if punch_in_am and punch_out_am:
            if am_late <= punch_in_am.time() < am_absent:
                late_minutes = (datetime.datetime.combine(current_date, punch_in_am.time()) - datetime.datetime.combine(current_date, am_late)).total_seconds() // 60
                status["Late Minutes"] += int(late_minutes)
                status["Deductions"] += late_minutes * deduction_per_minute
        else:
            status["Absent"] += 1
            status["Deductions"] += daily_salary // 2.0

        # PM Shift Check
        punch_in_pm = next((p for p in pm_shift if pm_start <= p.time() < pm_absent), None)
        punch_out_pm = next((p for p in pm_shift if pm_end <= p.time() <= pm_absent), None)

        if punch_in_pm and punch_out_pm:
            if pm_late <= punch_in_pm.time() < pm_absent:
                late_minutes = (datetime.datetime.combine(current_date, punch_in_pm.time()) - datetime.datetime.combine(current_date, pm_late)).total_seconds() // 60
                status["Late Minutes"] += int(late_minutes)
                status["Deductions"] += late_minutes * deduction_per_minute
        else:
            status["Absent"] += 1
            status["Deductions"] += daily_salary // 2.0

    return status


def generate_excel(filename, employee_attendance, start_date, end_date):
    # Load config before generating excel
    config = load_config()
    report_directory = config["report_directory"]

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    report_title = f"Employee Attendance Report ({start_date} - {end_date})"
    title_cell = sheet.cell(row=1, column=1, value=report_title)
    title_cell.font = openpyxl.styles.Font(size=16, bold=True)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)

    header = ["ID", "First Name", "Last Name", "Total Shifts", "Late Minutes", "Shifts Absent", "Deductions", "Gross Salary", "Net Salary"]
    sheet.append(header)

    for emp_id, data in employee_attendance.items():
        row = [
            emp_id,
            data["first_name"],
            data["last_name"],
            data["total_shifts"],
            data["late"],
            data["absent"],
            data["deductions"],
            data["gross_salary"],
            data["net_salary"]
        ]
        sheet.append(row)


    full_path = os.path.join(report_directory, filename)
    workbook.save(full_path)

    try:
        os.startfile(full_path) if os.name == 'nt' else subprocess.call(['open', full_path])
    except Exception as e:
        print(f"Error opening Excel file: {e}")