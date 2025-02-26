import tkinter as tk
import attendance
import re
import json
from tkinter import filedialog, messagebox
import os
import datetime
from tkcalendar import DateEntry

config_file = "config.json"

default_config = {
    "db_path": r"C:\Program Files (x86)\ZKBio Time.Net\TimeNet.db",
    "report_directory": r"C:\Users\yhljo\PycharmProjects\PythonProject\HelloWorld",
    "daily_salary": float(410.0),
    "deduction_per_minute": float(0.85)
}


# Load configuration
def load_config():
    try:
        with open(config_file, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return default_config


# Save configuration
def save_config(config):
    with open(config_file, "w") as f:
        json.dump(config, f, indent=4)
    messagebox.showinfo("Success", "Configuration saved successfully!")


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("600x370")
        self.title("CHING - BioTime")
        self.previous_screen = "main"

        script_dir = os.path.dirname(os.path.abspath(__file__))
        # create the full path to the icon file
        icon_path = os.path.join(script_dir, "app_icon.ico")
        self.iconbitmap(icon_path)  # Sets the icon.

        # Create a container for all screens
        self.container = tk.Frame(self, bg="#FFFDF0")
        self.container.pack(fill="both", expand=True)

        # dictionary para sa tanan nga screens
        self.screens = {}

        # initialize the main screen
        self.main_screen = MainScreen(self.container, self)
        self.screens["main"] = self.main_screen

        # initialize the settings screen
        self.settings_screen = SettingsScreen(self.container, self)
        self.screens["settings"] = self.settings_screen

        # initialize the date range screen
        self.date_range_screen = DateRangeScreen(self.container, self)
        self.screens["date_range"] = self.date_range_screen

        # initialize the report screen
        self.generate_report_screen = GenerateReportScreen(self.container, self)
        self.screens["generate_report"] = self.generate_report_screen

        # show the main screen
        self.show_screen("main")

        # store start and end dates here
        self.start_date = None
        self.end_date = None

    def show_screen(self, screen_name):
        # Hide all screens
        for screen in self.screens.values():
            screen.pack_forget()
        # Show the requested screen
        self.screens[screen_name].pack(fill="both", expand=True)

class MainScreen(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        self.configure(bg="#FFFDF0")


        # Create a frame to center the content
        center_frame = tk.Frame(self, bg="#FFFDF0")
        center_frame.pack(expand=True, fill="both")

        # Settings button (top right)
        settings_button = tk.Button(self, text="⚙️", font=('Segoe UI', 18),
                                    command=lambda: self.open_settings(),
                                    relief=tk.FLAT, cursor="hand2", bg="#FFFDF0", fg="#A31D1D")
        settings_button.place(relx=0.98, rely=0.02, anchor="ne")

        # Update the label text
        label = tk.Label(center_frame, text="Before generating reports, ensure you have the latest attendance data.\nRetrieve new transactions from ZKBio Time.Net software before continuing.\n \nSteps:\nOpen ZKBio Time.Net software > Go to Device (Main Section) >\nSelect Device (Subsection) > Click 'Get Transactions'", font=('Segoe UI', 12), wraplength=780, justify=tk.CENTER, bg="#FFFDF0")
        label.pack(expand=True, fill="both")

        # Create a button that says "Continue"
        continue_button = tk.Button(self, text="Continue", font=('Segoe UI', 14), command=lambda: controller.show_screen("date_range"), bg="#6D2323", fg="white", cursor="hand2", pady=5, relief=tk.RIDGE)
        continue_button.pack(side="bottom", fill="x")

    def open_settings(self):
        self.controller.previous_screen = "main"
        self.controller.show_screen("settings")

class SettingsScreen(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        self.configure(bg="#FFFDF0") #configure the frame

        # Create a frame to center the content
        center_frame = tk.Frame(self, bg="#FFFDF0") #configure the frame
        center_frame.pack(expand=True, fill="both")

        # Back button to return to the main screen
        back_button = tk.Button(self, text="⬅", font=('Segoe UI', 17), bg="#FFFDF0", fg="#A31D1D",
                                command=lambda: controller.show_screen(controller.previous_screen), cursor="hand2",
                                relief=tk.FLAT)
        back_button.pack(pady=2, padx=2)
        back_button.place(relx=0.10, rely=0.01, anchor="ne")

        # Create the box_frame
        box_frame = tk.Frame(self, bg="#FFFDF0") #configure the frame
        box_frame.place(relx=0.5, rely=0.42, anchor="center")

        # Create the center_frame inside box_frame
        center_frame = tk.Frame(box_frame, bg="#FFFDF0") #configure the frame
        center_frame.pack(anchor="w", padx=20, pady=5)

        config = load_config()

        # Browsing functions
        def browse_db_path():
            path = filedialog.askopenfilename(title="Select Database File", filetypes=[("Database Files", "*.db")])
            if path:
                db_path_entry.delete(0, tk.END)
                db_path_entry.insert(0, path)

        def browse_report_directory():
            path = filedialog.askdirectory(title="Select Report Directory")
            if path:
                report_directory_entry.delete(0, tk.END)
                report_directory_entry.insert(0, path)

        def create_label_entry(label, entry_var, width=55):
            tk.Label(center_frame, text=label, font=('Segoe UI', 11), anchor="w", bg="#FFFDF0").pack(pady=0.5, anchor="w")
            entry = tk.Entry(center_frame, width=width, font=('Segoe UI', 11))
            entry.insert(0, entry_var)
            entry.pack(anchor="w")
            return entry

        db_path_entry = create_label_entry("Database Path:", config["db_path"])
        tk.Button(center_frame, text="Browse DB", command=browse_db_path, cursor="hand2", relief=tk.GROOVE, bg="snow2", width=10).pack(pady=5, anchor="w")
        report_directory_entry = create_label_entry("Report Directory:", config["report_directory"])
        tk.Button(center_frame, text="Browse Report Directory", command=browse_report_directory, cursor="hand2", relief=tk.GROOVE, bg="snow2", width=20).pack(pady=(5, 15),
                                                                                                      anchor="w")
        daily_salary_entry = create_label_entry("Daily Salary (PHP):", config["daily_salary"], width=10)
        deduction_entry = create_label_entry("Deduction per Minute (PHP):", config["deduction_per_minute"], width=10)

        def save():
            config = {
                "db_path": db_path_entry.get(),
                "report_directory": report_directory_entry.get(),
                "daily_salary": float(daily_salary_entry.get()),
                "deduction_per_minute": float(deduction_entry.get())
            }
            save_config(config)

        # Submit button
        save_button = tk.Button(self, text="Save", command=save, font=('Segoe UI', 14), bg="#6D2323", fg="white", cursor="hand2", pady=5, relief=tk.RIDGE)
        save_button.pack(side="bottom", fill="x")

class DateRangeScreen(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        self.configure(bg="#FFFDF0")

        # Center frame
        center_frame = tk.Frame(self, bg="#FFFDF0")
        center_frame.pack(expand=True, fill="both")

        # Configure grid to center widgets
        center_frame.grid_rowconfigure(0, weight=1)
        center_frame.grid_columnconfigure(0, weight=1)

        # Back button
        back_button = tk.Button(self, text="⬅", font=('Segoe UI', 17),
                                command=lambda: controller.show_screen("main"), cursor="hand2", relief=tk.FLAT,
                                bg="#FFFDF0", fg="#A31D1D")
        back_button.place(relx=0.10, rely=0.01, anchor="ne")

        # Settings button
        settings_button = tk.Button(self, text="⚙️", font=('Segoe UI', 18),
                                    command=lambda: self.open_settings(), relief=tk.FLAT, cursor="hand2", bg="#FFFDF0",
                                    fg="#A31D1D")
        settings_button.place(relx=0.98, rely=0.02, anchor="ne")

        # Vertical frame (holds date entries)
        vertical_frame = tk.Frame(center_frame, bg="#FFFDF0")
        vertical_frame.grid(row=0, column=0)

        # Horizontal frame for Start and End Date
        date_frame = tk.Frame(vertical_frame, bg="#FFFDF0")
        date_frame.pack(pady=20)

        # Start Date
        tk.Label(date_frame, text="Start Date:", font=('Segoe UI', 12), bg="#FFFDF0").grid(row=0, column=0, padx=(0,5))
        self.start_date_entry = DateEntry(date_frame, font=('Segoe UI', 12), date_pattern='yyyy-mm-dd',
                                          background="#6D2323", foreground="white", headersbackground="#6D2323",
                                          headersforeground="white", selectbackground="#A31D1D")
        self.start_date_entry.grid(row=0, column=1, padx=(0,70))

        # End Date
        tk.Label(date_frame, text="End Date:", font=('Segoe UI', 12), bg="#FFFDF0").grid(row=0, column=2, padx=(0,5))
        self.end_date_entry = DateEntry(date_frame, font=('Segoe UI', 12), date_pattern='yyyy-mm-dd',
                                        background="#6D2323", foreground="white", headersbackground="#6D2323",
                                        headersforeground="white", selectbackground="#A31D1D")
        self.end_date_entry.grid(row=0, column=3, padx=(0,0))

        # Submit button
        submit_button = tk.Button(self, text="Submit", font=('Segoe UI', 14), command=self.submit_dates, bg="#6D2323", fg="white", cursor="hand2", pady=5, relief=tk.RIDGE)
        submit_button.pack(side="bottom", fill="x")

    def submit_dates(self):
        start_date_str = self.start_date_entry.get()
        end_date_str = self.end_date_entry.get()

        if self.validate_date(start_date_str) and self.validate_date(end_date_str):
            try:
                start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
                end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()

                if end_date < start_date:
                    messagebox.showerror("Invalid Date Range", "End date must be later than or equal to start date.")
                    return

                self.controller.start_date = start_date_str
                self.controller.end_date = end_date_str
                self.controller.show_screen("generate_report")

            except ValueError:
                messagebox.showerror("Invalid Date", "Please enter dates in the format YYYY-MM-DD.")

        else:
            messagebox.showerror("Invalid Date", "Please enter dates in the format YYYY-MM-DD.")

    def validate_date(self, date_str):
        # Check if the date matches the YYYY-MM-DD format using regex
        date_pattern = r"^\d{4}-\d{2}-\d{2}$"
        return re.match(date_pattern, date_str) is not None

    def set_equal_button_size(self, buttons):
        # Calculate maximum width and height
        fixed_width = 15
        fixed_height = 1

        # Set all buttons to the maximum width and height
        for button in buttons:
            button.config(width=fixed_width, height=fixed_height)

    def open_settings(self):
        self.controller.previous_screen = "date_range"
        self.controller.show_screen("settings")


class GenerateReportScreen(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        self.configure(bg="#FFFDF0")  # configure the frame

        # Back button to return to the main screen
        back_button = tk.Button(self, text="⬅", font=('Segoe UI', 17),
                                command=lambda: controller.show_screen("date_range"), cursor="hand2", relief=tk.FLAT, bg="#FFFDF0", fg="#A31D1D")
        back_button.pack(pady=2, padx=2)
        back_button.place(relx=0.10, rely=0.01, anchor="ne")

        center_frame = tk.Frame(self, bg="#FFFDF0")
        center_frame.pack(expand=True)

        vertical_frame = tk.Frame(center_frame, bg="#FFFDF0")
        vertical_frame.pack(pady=20)

        filename_label = tk.Label(vertical_frame, text="Enter Excel File Name (Ex: 'attendance.xlsx'):", font=('Segoe UI', 12), bg="#FFFDF0")
        filename_label.pack(pady=(0, 5))

        # Excel File Name Entry with Placeholder
        self.filename_entry = tk.Entry(vertical_frame, font=('Segoe UI', 12), width=40, justify=tk.CENTER, fg="gray")
        self.filename_entry.insert(0, "attendance.xlsx")

        # Placeholder behavior
        def on_focus_in(event):
            if self.filename_entry.get() == "attendance.xlsx":
                self.filename_entry.delete(0, tk.END)
                self.filename_entry.config(fg="black")

        def on_focus_out(event):
            if not self.filename_entry.get().strip():
                self.filename_entry.insert(0, "attendance.xlsx")
                self.filename_entry.config(fg="gray")

        # Bind focus events
        self.filename_entry.bind("<FocusIn>", on_focus_in)
        self.filename_entry.bind("<FocusOut>", on_focus_out)

        self.filename_entry.pack(pady=(0, 20))

        generate_button = tk.Button(self, text="Generate Report", font=('Segoe UI', 14), command=self.generate_report, bg="#6D2323", fg="white", cursor="hand2", pady=5, relief=tk.RIDGE)
        generate_button.pack(side="bottom", fill="x")

    def generate_report(self):
        filename = self.filename_entry.get().strip()

        # Check if the filename ends with ".xlsx"
        if filename:
            if not filename.endswith(".xlsx"):
                messagebox.showerror("Invalid Filename", "The filename must end with '.xlsx'.")
                return

            try:
                attendance.process_dates(self.controller.start_date, self.controller.end_date, filename)
                messagebox.showinfo("Report Generated", "Report generated successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to generate report: {e}")
        else:
            messagebox.showerror("Error", "Please enter a valid filename.")


if __name__ == "__main__":
    try:
        app = Application()
        app.mainloop()
    except Exception as e:
        print(f"An error occurred: {e}")