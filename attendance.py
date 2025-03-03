import sqlite3
import datetime
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os
import subprocess
import json

# ===============================================================
# TIME CONSTANTS
# ===============================================================
am_start = datetime.time(7, 30, 0)  # 7:30 am start time
am_late = datetime.time(9, 30, 0)  # 9:30 am late start
am_absent = datetime.time(10, 0, 0)  # 10:15 am absent
am_end = datetime.time(14, 0, 0)  # 2:00 pm out time
am_latest_out = datetime.time(15, 0, 0)  # 3:00 pm latest out

pm_start = datetime.time(15, 1, 0)  # 3:01 pm start time
pm_late = datetime.time(16, 0, 0)  # 4:00 pm late start
pm_absent = datetime.time(16, 31, 0)  # 4:30 pm absent
pm_end = datetime.time(21, 0, 0)  # 9:00 pm out time
pm_latest_out = datetime.time(23, 59, 0)  # 11:59 pm latest out

config_file = "config.json"


# ===============================================================
# CONFIGURATION MANAGEMENT
# ===============================================================

# Load configuration from file or return defaults
def load_config():
    try:
        with open(config_file, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {
            "db_path": "C:/Program Files (x86)/ZKBio Time.Net/TimeNet.db",
            "report_directory": "C:/Users/Public/Documents",
            "department_salaries": {
                "Dining 1": 12750.0,
                "Dining 2": 12300.0,
                "Chief Cook": 28000.0,
                "Senior Cook": 25000.0,
                "Cook": 20000.0,
                "Chief Cutter": 27000.0,
                "Senior Cutter": 18000.0,
                "Cutter": 13000.0,
                "Quality Control": 16000.0,
                "Helper 1": 13000.0,
                "Helper 2": 12300.0,
                "Washer": 12300.0
            }
        }


# Get salary configuration for a specific department
def get_salary_config(dept_name):
    """Return salary configuration based on department role"""

    config = load_config()

    # Define deduction rates and absence deductions based on department salaries
    role_config = {
        "Dining 1": {
            "daily_salary": config["department_salaries"]["Dining 1"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Dining 1"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Dining 1"] / 30) / 2
        },
        "Dining 2": {
            "daily_salary": config["department_salaries"]["Dining 2"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Dining 2"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Dining 2"] / 30) / 2
        },
        "Chief Cook": {
            "daily_salary": config["department_salaries"]["Chief Cook"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Chief Cook"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Chief Cook"] / 30) / 2
        },
        "Senior Cook": {
            "daily_salary": config["department_salaries"]["Senior Cook"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Senior Cook"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Senior Cook"] / 30) / 2
        },
        "Cook": {
            "daily_salary": config["department_salaries"]["Cook"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Cook"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Cook"] / 30) / 2
        },
        "Chief Cutter": {
            "daily_salary": config["department_salaries"]["Chief Cutter"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Chief Cutter"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Chief Cutter"] / 30) / 2
        },
        "Senior Cutter": {
            "daily_salary": config["department_salaries"]["Senior Cutter"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Senior Cutter"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Senior Cutter"] / 30) / 2
        },
        "Cutter": {
            "daily_salary": config["department_salaries"]["Cutter"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Cutter"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Cutter"] / 30) / 2
        },
        "Quality Control": {
            "daily_salary": config["department_salaries"]["Quality Control"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Quality Control"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Quality Control"] / 30) / 2
        },
        "Helper 1": {
            "daily_salary": config["department_salaries"]["Helper 1"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Helper 1"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Helper 1"] / 30) / 2
        },
        "Helper 2": {
            "daily_salary": config["department_salaries"]["Helper 2"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Helper 2"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Helper 2"] / 30) / 2
        },
        "Washer": {
            "daily_salary": config["department_salaries"]["Washer"] / 30,
            "deduction_per_minute": ((config["department_salaries"]["Washer"] / 30) / 8) / 60,
            "absence_deduction": (config["department_salaries"]["Washer"] / 30) / 2
        }
    }

    return role_config.get(dept_name)


# ===============================================================
# MAIN PROCESSING FUNCTION
# ===============================================================

# Process attendance data for a date range and generate Excel report
def process_dates(start_date, end_date, excel_filename):
    try:
        # Load configuration
        config = load_config()
        db_path = config["db_path"]
        report_directory = config["report_directory"]

        start_dt = datetime.datetime.strptime(start_date, "%Y-%m-%d")
        end_dt = datetime.datetime.strptime(end_date, "%Y-%m-%d")

        num_days = (end_dt - start_dt).days + 1
        total_shifts = num_days * 2

        # Connect to database and retrieve punch data
        with sqlite3.connect(db_path) as conn:
            cursor = conn.cursor()
            query = """
                SELECT em.id, em.emp_firstname, em.emp_lastname, em.department_id, dep.dept_name, ap.punch_time
                FROM hr_employee em
                INNER JOIN hr_department dep ON em.department_id = dep.id
                LEFT JOIN att_punches ap ON em.id = ap.employee_id
                WHERE (date(ap.punch_time) BETWEEN ? AND ?) AND (em.emp_privilege=0)
                ORDER BY em.department_id, ap.punch_time;
            """
            cursor.execute(query, (start_date, end_date))
            results = cursor.fetchall()

            # Initialize employee attendance dictionary
            employee_attendance = {}
            for emp_id, first_name, last_name, department_id, dept_name, punch_time_str in results:
                punch_time = datetime.datetime.strptime(punch_time_str, "%Y-%m-%d %H:%M:%S")

                # Create employee record if doesn't exist
                if emp_id not in employee_attendance:
                    # Get salary configuration based on department
                    salary_config = get_salary_config(dept_name)
                    daily_salary = salary_config["daily_salary"]
                    gross_salary = daily_salary * num_days

                    employee_attendance[emp_id] = {
                        "first_name": first_name,
                        "last_name": last_name,
                        "department_id": department_id,
                        "dept_name": dept_name,
                        "daily_salary": daily_salary,  # Store daily salary for later use
                        "total_shifts": total_shifts,
                        "late": 0,
                        "absent": 0,
                        "deductions": 0,
                        "gross_salary": gross_salary,
                        "paid_leave_days": 0,  # Default to 0
                        "salary_advance": 0,  # Default to 0
                        "net_salary": 0,
                        "punches": []
                    }

                # Add punch time to employee's record
                employee_attendance[emp_id]["punches"].append(punch_time)

            # Process attendance for each employee
            for emp_id, data in employee_attendance.items():
                # Get role-specific configuration
                salary_config = get_salary_config(data["dept_name"])
                deduction_per_minute = salary_config["deduction_per_minute"]
                daily_salary = salary_config["daily_salary"]
                absence_deduction = salary_config["absence_deduction"]

                # Check attendance status
                status = check_attendance(
                    data["punches"],
                    deduction_per_minute,
                    absence_deduction,
                    start_date,
                    end_date
                )

                # Update employee record with attendance status
                data["late"] = status["Late Minutes"]
                data["absent"] = status["Absent"]
                data["deductions"] = status["Deductions"]

                # Net salary calculation will happen in the Excel sheet via formulas

            # Create report directory if it doesn't exist
            if not os.path.exists(report_directory):
                os.makedirs(report_directory)

            # Generate Excel report
            generate_excel(excel_filename, employee_attendance, start_date, end_date)
            print(f"Excel report generated successfully: {os.path.join(report_directory, excel_filename)}")

            conn.close()

    except sqlite3.Error as e:
        print(f"Database error: {e}")
    except ValueError as e:
        print(f"Invalid input: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


# ===============================================================
# ATTENDANCE CHECKING
# ===============================================================

# Check attendance for an employee within the date range
def check_attendance(punches, deduction_per_minute, absence_deduction, start_date, end_date):
    status = {"Late Minutes": 0, "Absent": 0, "Deductions": 0}
    punches.sort()

    # Generate a list of all dates in the range
    start_dt = datetime.datetime.strptime(start_date, "%Y-%m-%d").date()
    end_dt = datetime.datetime.strptime(end_date, "%Y-%m-%d").date()
    all_dates = [start_dt + datetime.timedelta(days=i) for i in range((end_dt - start_dt).days + 1)]

    # Process each date in the range
    for current_date in all_dates:
        am_shift = []
        pm_shift = []

        # Filter punches for the current date and split by shift
        for punch in punches:
            if punch.date() == current_date:
                if am_start <= punch.time() <= am_latest_out:
                    am_shift.append(punch)
                elif pm_start <= punch.time() <= pm_latest_out:
                    pm_shift.append(punch)

        # Check morning shift attendance
        punch_in_am = next((p for p in am_shift if am_start <= p.time() < am_absent), None)
        punch_out_am = next((p for p in am_shift if am_end <= p.time() <= am_latest_out), None)

        if punch_in_am and punch_out_am:
            # Employee was present but check if late
            if am_late <= punch_in_am.time() < am_absent:
                late_minutes = (datetime.datetime.combine(current_date, punch_in_am.time()) - datetime.datetime.combine(
                    current_date, am_late)).total_seconds() // 60
                status["Late Minutes"] += int(late_minutes)
                status["Deductions"] += late_minutes * deduction_per_minute
        else:
            # Employee was absent for morning shift
            status["Absent"] += 1
            status["Deductions"] += absence_deduction

        # Check afternoon shift attendance
        punch_in_pm = next((p for p in pm_shift if pm_start <= p.time() < pm_absent), None)
        punch_out_pm = next((p for p in pm_shift if pm_end <= p.time() <= pm_latest_out), None)

        if punch_in_pm and punch_out_pm:
            # Employee was present but check if late
            if pm_late <= punch_in_pm.time() < pm_absent:
                late_minutes = (datetime.datetime.combine(current_date, punch_in_pm.time()) - datetime.datetime.combine(
                    current_date, pm_late)).total_seconds() // 60
                status["Late Minutes"] += int(late_minutes)
                status["Deductions"] += late_minutes * deduction_per_minute
        else:
            # Employee was absent for afternoon shift
            status["Absent"] += 1
            status["Deductions"] += absence_deduction

    return status


# ===============================================================
# EXCEL REPORT GENERATION
# ===============================================================

# Generate Excel report with attendance and salary data
def generate_excel(filename, employee_attendance, start_date, end_date):
    # Load config before generating excel
    config = load_config()
    report_directory = config["report_directory"]

    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Set column widths for better readability
    column_widths = {
        1: 8,  # ID
        2: 15,  # First Name
        3: 15,  # Last Name
        4: 15,  # Position
        5: 12,  # Total Shifts
        6: 12,  # Late Minutes
        7: 12,  # Shifts Absent
        8: 14,  # Daily Salary
        9: 14,  # Gross Salary
        10: 14,  # Deductions
        11: 15,  # Salary Advance
        12: 15,  # Others
        13: 14,  # Net Salary
    }

    # Apply column widths
    for col_num, width in column_widths.items():
        sheet.column_dimensions[get_column_letter(col_num)].width = width

    # Create report title
    report_title = f"Employee Attendance Report ({start_date} - {end_date})"
    title_cell = sheet.cell(row=1, column=1, value=report_title)
    title_cell.font = Font(size=20, bold=True)
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=13)
    title_cell.alignment = Alignment(horizontal='center')

    # Create headers with formatting
    headers = ["ID", "First Name", "Last Name", "Position", "Total Shifts", "Late Minutes",
               "Shifts Absent", "Daily Salary", "Gross Salary", "Deductions", "Salary Advance",
               "Others", "Net Salary"]

    header_row = 2
    for col_num, header_text in enumerate(headers, 1):
        cell = sheet.cell(row=header_row, column=col_num, value=header_text)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')

    # Add employee data starting from row 3
    row_num = 3
    for emp_id, data in employee_attendance.items():
        # Map column data
        daily_salary = data["daily_salary"]
        num_days = data["gross_salary"] / daily_salary

        # Add the basic data cells
        sheet.cell(row=row_num, column=1, value=emp_id)
        sheet.cell(row=row_num, column=2, value=data["first_name"])
        sheet.cell(row=row_num, column=3, value=data["last_name"])
        sheet.cell(row=row_num, column=4, value=data["dept_name"])
        sheet.cell(row=row_num, column=5, value=data["total_shifts"])
        sheet.cell(row=row_num, column=6, value=data["late"])
        sheet.cell(row=row_num, column=7, value=data["absent"])

        # Daily Salary - now editable with initial calculated value
        daily_salary_cell = sheet.cell(row=row_num, column=8, value=daily_salary)
        daily_salary_cell.number_format = '#,##0.00'

        # Gross Salary now uses a formula referencing the Daily Salary cell
        gross_salary_cell = sheet.cell(row=row_num, column=9, value=f"=H{row_num}*{num_days}")
        gross_salary_cell.number_format = '#,##0.00'

        # Deductions - now editable with initial calculated value
        deductions_cell = sheet.cell(row=row_num, column=10, value=data["deductions"])
        deductions_cell.number_format = '#,##0.00'

        # Salary Advance (inputtable)
        salary_advance_cell = sheet.cell(row=row_num, column=11, value=0)
        salary_advance_cell.number_format = '#,##0.00'

        # Others (inputtable) - now after Salary Advance
        others_cell = sheet.cell(row=row_num, column=12, value=0)
        others_cell.number_format = '#,##0.00'

        # Net Salary formula updated with the new column positions
        net_salary_cell = sheet.cell(row=row_num, column=13, value=f"=I{row_num}-J{row_num}-K{row_num}+L{row_num}")
        net_salary_cell.number_format = '#,##0.00'

        # Highlight cells that are meant to be edited by the user
        input_cols = [8, 10, 11, 12]  # Daily Salary, Deductions, Salary Advance, Others
        for col in input_cols:
            cell = sheet.cell(row=row_num, column=col)
            cell.fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

        row_num += 1

    # Add instructions
    instruction_row = row_num + 2
    instructions = [
        "NOTE:",
        "- Yellow cells can be edited to adjust Daily Salary, Deductions, Salary Advances, and Others",
        "- The Others column can contain positive or negative values which will be reflected in the Net Salary",
        "- Changes to yellow cells will automatically update the Net Salary calculation"
    ]

    for i, instruction in enumerate(instructions):
        cell = sheet.cell(row=instruction_row + i, column=1, value=instruction)
        # Make all instruction text bold
        cell.font = Font(bold=True)
        sheet.merge_cells(start_row=instruction_row + i, start_column=1, end_row=instruction_row + i, end_column=7)

    # Apply borders to the main table
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for row in range(2, row_num):
        for col in range(1, 14):
            sheet.cell(row=row, column=col).border = thin_border

    # Save the workbook
    full_path = os.path.join(report_directory, filename)
    workbook.save(full_path)

    # Try to open the Excel file automatically
    try:
        os.startfile(full_path) if os.name == 'nt' else subprocess.call(['open', full_path])
    except Exception as e:
        print(f"Error opening Excel file: {e}")