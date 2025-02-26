import json
import tkinter as tk
from tkinter import filedialog, messagebox
import os

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


# GUI for settings
def open_settings():
    root = tk.Tk()
    root.title("Settings")
    root.geometry("450x400")
    root.configure(bg="white")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    # create the full path to the icon file
    icon_path = os.path.join(script_dir, "app_icon.ico")  # replace your_icon.ico
    root.iconbitmap(icon_path)  # Sets the icon.

    box_frame = tk.Frame(root, bg="white")
    box_frame.place(relx=0.6, rely=0.5, anchor="center", width=500, height=350)

    center_frame = tk.Frame(box_frame, bg="white")
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
        tk.Label(center_frame, text=label, bg="white", font=('Arial', 10), anchor="w").pack(pady=2, anchor="w")
        entry = tk.Entry(center_frame, width=width)
        entry.insert(0, entry_var)
        entry.pack(anchor="w")
        return entry

    db_path_entry = create_label_entry("Database Path:", config["db_path"])
    tk.Button(center_frame, text="Browse DB", command=browse_db_path).pack(pady=5, anchor="w")
    report_directory_entry = create_label_entry("Report Directory:", config["report_directory"])
    tk.Button(center_frame, text="Browse Report Directory", command=browse_report_directory).pack(pady=5, anchor="w")
    daily_salary_entry = create_label_entry("Daily Salary (PHP):", config["daily_salary"], width=20)
    daily_salary_entry.pack_configure(pady=5)
    deduction_entry = create_label_entry("Deduction per Minute (PHP):", config["deduction_per_minute"], width=20)

    def save():
        config = {
            "db_path": db_path_entry.get(),
            "report_directory": report_directory_entry.get(),
            "daily_salary": float(daily_salary_entry.get()),
            "deduction_per_minute": float(deduction_entry.get())
        }
        save_config(config)

    # Save button centered
    button_frame = tk.Frame(box_frame, bg="white")
    button_frame.pack(fill="x", pady=10)

    save_button = tk.Button(root, text="Save", command=save, bg="green", fg="white", width=12)
    save_button.place(relx=0.5, rely=0.85, anchor="center")

    root.mainloop()
