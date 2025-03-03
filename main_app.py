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
    "report_directory": r"C:\Users\Public\Documents",
    "daily_salary": float(410.0),
    "deduction_per_minute": float(0.85),
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
        self.geometry("670x480")
        self.title("CHING - BioTime")

        script_dir = os.path.dirname(os.path.abspath(__file__))
        # create the full path to the icon file
        icon_path = os.path.join(script_dir, "app_icon.ico")
        self.iconbitmap(icon_path)  # Sets the icon.

        # Create a container for all screens
        self.container = tk.Frame(self, bg="#FFFDF0")
        self.container.pack(fill="both", expand=True)

        # Create frames for different screens
        self.frames = {}
        for F in (MainScreen, SettingsScreen):
            frame = F(self.container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # Configure the container grid
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        # Show main screen initially
        self.show_frame(MainScreen)

        # Store start and end dates here
        self.start_date = None
        self.end_date = None

    def show_frame(self, frame_class):
        """Raise the specified frame to the top"""
        frame = self.frames[frame_class]
        frame.tkraise()


class MainScreen(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        self.configure(bg="#FFFDF0")

        # Main content frame - using pack with fill="both" and expand=True
        content_frame = tk.Frame(self, bg="#FFFDF0")
        content_frame.pack(expand=True, fill="both", padx=20, pady=40)
        content_frame.place(relx=0.5, rely=0.1, anchor="n")

        # Settings button (top right) - placed directly on the frame, not in content_frame
        settings_button = tk.Button(self, text="⚙️", font=('Segoe UI', 20),
                                    command=lambda: controller.show_frame(SettingsScreen),
                                    relief=tk.FLAT, cursor="hand2", bg="#FFFDF0", fg="#A31D1D")
        settings_button.place(relx=0.98, rely=0.02, anchor="ne")

        # Top section - Instructions
        instructions_text = "Before generating reports, ensure you have the latest attendance data.\nRetrieve new transactions from ZKBio Time.Net software before continuing.\n \nSteps:\nOpen ZKBio Time.Net software > Go to Device (Main Section) >\nSelect Device (Subsection) > Click 'Get Transactions'"
        instructions_label = tk.Label(content_frame, text=instructions_text, font=('Segoe UI', 12),
                                      wraplength=750, justify=tk.CENTER, bg="#FFFDF0")
        instructions_label.pack(pady=(20, 30))

        # Middle section - Date Range Selection
        date_frame = tk.Frame(content_frame, bg="#FFFDF0")
        date_frame.pack(pady=20)

        # Start Date
        tk.Label(date_frame, text="Start Date:", font=('Segoe UI', 12), bg="#FFFDF0").grid(row=0, column=0, padx=(0, 5))
        self.start_date_entry = DateEntry(date_frame, font=('Segoe UI', 12), date_pattern='yyyy-mm-dd',
                                          background="#6D2323", foreground="white", headersbackground="#FFFDF0",
                                          headersforeground="#6D2323", selectbackground="#A31D1D")
        self.start_date_entry.grid(row=0, column=1, padx=(0, 70))

        # End Date
        tk.Label(date_frame, text="End Date:", font=('Segoe UI', 12), bg="#FFFDF0").grid(row=0, column=2, padx=(0, 5))
        self.end_date_entry = DateEntry(date_frame, font=('Segoe UI', 12), date_pattern='yyyy-mm-dd',
                                        background="#6D2323", foreground="white", headersbackground="#FFFDF0",
                                        headersforeground="#6D2323", selectbackground="#A31D1D")
        self.end_date_entry.grid(row=0, column=3, padx=(0, 0))

        # Bottom section - Excel File Name Entry
        file_frame = tk.Frame(content_frame, bg="#FFFDF0")
        file_frame.pack(pady=20)

        filename_label = tk.Label(file_frame, text="Enter Excel File Name (Ex: 'attendance.xlsx'):",
                                  font=('Segoe UI', 12), bg="#FFFDF0")
        filename_label.pack(pady=(0, 5))

        self.filename_entry = tk.Entry(file_frame, font=('Segoe UI', 12), width=40, justify=tk.CENTER, fg="gray")
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

        self.filename_entry.pack(pady=(0, 10))

        # Generate Report button
        generate_button = tk.Button(self, text="Generate Report", font=('Segoe UI', 13),
                                    command=self.generate_report, bg="#6D2323", fg="white",
                                    cursor="hand2", pady=2, width=20, relief=tk.RAISED)
        generate_button.place(relx=0.5, rely=0.96, anchor="s")

    def generate_report(self):
        start_date_str = self.start_date_entry.get()
        end_date_str = self.end_date_entry.get()
        filename = self.filename_entry.get().strip()

        if self.validate_date(start_date_str) and self.validate_date(end_date_str):
            try:
                start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
                end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()

                if end_date < start_date:
                    messagebox.showerror("Invalid Date Range", "End date must be later than or equal to start date.")
                    return

                if filename:
                    if not filename.endswith(".xlsx"):
                        messagebox.showerror("Invalid Filename", "The filename must end with '.xlsx'.")
                        return

                    try:
                        attendance.process_dates(start_date_str, end_date_str, filename)
                        messagebox.showinfo("Report Generated", "Report generated successfully!")
                    except Exception as e:
                        messagebox.showerror("Error", f"Failed to generate report: {e}")
                else:
                    messagebox.showerror("Error", "Please enter a valid filename.")

            except ValueError:
                messagebox.showerror("Invalid Date", "Please enter dates in the format YYYY-MM-DD.")
        else:
            messagebox.showerror("Invalid Date", "Please enter dates in the format YYYY-MM-DD.")

    def validate_date(self, date_str):
        # Check if the date matches the YYYY-MM-DD format using regex
        date_pattern = r"^\d{4}-\d{2}-\d{2}$"
        return re.match(date_pattern, date_str) is not None


class SettingsScreen(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.configure(bg="#FFFDF0")

        # Create a canvas with scrollbar container that fills the window width
        canvas_container = tk.Frame(self, bg="#FFFDF0")
        canvas_container.place(relx=0.5, rely=0.1, anchor="n", relwidth=1.0, relheight=0.8)

        # Create scrollbar first (will be on the far right)
        scrollbar = tk.Scrollbar(canvas_container, orient="vertical")
        scrollbar.pack(side="right", fill="both", padx=(0, 5))  # No padding to right edge

        # Create canvas with padding on the left to center content
        canvas = tk.Canvas(canvas_container, bg="#FFFDF0", highlightthickness=0)
        canvas.pack(side="left", fill="y", expand=True, padx=(0, 0))  # Add left padding only

        # Back button (top left)
        back_button = tk.Button(self, text="⬅️", font=('Segoe UI', 20),
                                command=lambda: controller.show_frame(MainScreen),
                                relief=tk.FLAT, cursor="hand2", bg="#FFFDF0", fg="#A31D1D")
        back_button.place(relx=0.02, rely=0.02, anchor="nw")

        # Page label
        page_label = tk.Label(self, text="Settings", font=('Segoe UI', 18, "bold"), relief=tk.FLAT,
                              bg="#FFFDF0", fg="#A31D1D")
        page_label.place(relx=0.5, rely=0.02, anchor="n")

        # Settings content frame that will contain all the settings
        settings_frame = tk.Frame(canvas, bg="#FFFDF0")

        # Configure the canvas
        settings_frame.bind("<Configure>",
                            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # Create window in canvas with settings_frame centered within available width
        canvas.create_window((0, 0), window=settings_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.configure(command=canvas.yview)

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

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

        def create_label_entry(label, entry_var, width=55, center=False):
            frame = tk.Frame(settings_frame, bg="#FFFDF0")
            frame.pack(anchor="w", pady=5)

            tk.Label(frame, text=label, font=('Segoe UI', 11), anchor="w", bg="#FFFDF0").pack(side=tk.TOP, anchor="w")
            justify = tk.CENTER if center else tk.LEFT
            entry = tk.Entry(frame, width=width, font=('Segoe UI', 10), justify=justify)
            entry.insert(0, entry_var)
            entry.pack(side=tk.TOP, anchor="w")
            return entry, frame

        tk.Label(settings_frame, text="Paths:", font=('Segoe UI', 11, "bold"),
                 anchor="w", bg="#FFFDF0").pack(anchor="w")

        # Report directory
        report_directory_entry, report_dir_frame = create_label_entry("Excel Report Directory:",
                                                                      config["report_directory"])
        tk.Button(report_dir_frame, text="Browse Directory", command=browse_report_directory, cursor="hand2",
                  relief=tk.RIDGE, bg="snow2", width=14).pack(pady=5, anchor="w")

        # Database path
        db_path_entry, db_path_frame = create_label_entry("Database Path:", config["db_path"])
        tk.Button(db_path_frame, text="Browse Path", command=browse_db_path, cursor="hand2",
                  relief=tk.RIDGE, bg="snow2", width=14).pack(pady=5, anchor="w")

        # Department salaries
        dept_salary_frame = tk.Frame(settings_frame, bg="#FFFDF0")
        dept_salary_frame.pack(anchor="w", pady=5)

        tk.Label(dept_salary_frame, text="Monthly Salaries:", font=('Segoe UI', 11, "bold"),
                 anchor="w", bg="#FFFDF0").pack(anchor="w")

        self.dept_salary_entries = {}
        dept_entries_frame = tk.Frame(dept_salary_frame, bg="#FFFDF0")
        dept_entries_frame.pack(anchor="w", pady=5, fill="x")

        row, col = 0, 0
        for dept, salary in config["department_salaries"].items():
            dept_frame = tk.Frame(dept_entries_frame, bg="#FFFDF0")
            dept_frame.grid(row=row, column=col, padx=(3,20), pady=5, sticky="w")

            # Label for department name
            tk.Label(dept_frame, text=(dept+ ":"), font=('Segoe UI', 11), anchor="w",
                     bg="#FFFDF0").pack(side=tk.LEFT, padx=(0, 5))

            # Entry for salary value
            entry = tk.Entry(dept_frame, width=10, font=('Segoe UI', 10), justify=tk.CENTER)
            entry.insert(0, salary)
            entry.pack(side=tk.LEFT)

            self.dept_salary_entries[dept] = entry

            # Arrange in 2 columns
            col += 1
            if col > 1:
                col = 0
                row += 1

        # Add extra whitespace at the bottom (approximately 1/4 of the window height)
        extra_space = tk.Frame(settings_frame, bg="#FFFDF0", height=70)
        extra_space.pack(side=tk.TOP, fill="x", pady=0)

        def save():
            try:
                config = {
                    "db_path": db_path_entry.get(),
                    "report_directory": report_directory_entry.get(),
                    "department_salaries": {dept: float(entry.get()) for dept, entry in
                                            self.dept_salary_entries.items()}
                }
                save_config(config)
                controller.show_frame(MainScreen)  # Return to main screen after saving
            except ValueError as e:
                messagebox.showerror("Invalid Input", f"Please enter valid numeric values for salary fields: {e}")

        # Save button
        save_button = tk.Button(self, text="Save", command=save, font=('Segoe UI', 13),
                                bg="#6D2323", fg="white", cursor="hand2", width=15, height=1, pady=2, relief=tk.RAISED)
        save_button.place(relx=0.5, rely=0.96, anchor="s")

        # Bind cleanup for mouse wheel when leaving this screen
        def _on_frame_leave(event):
            canvas.unbind_all("<MouseWheel>")

        self.bind("<Destroy>", _on_frame_leave)


if __name__ == "__main__":
    try:
        app = Application()
        app.mainloop()
    except Exception as e:
        print(f"An error occurred: {e}")