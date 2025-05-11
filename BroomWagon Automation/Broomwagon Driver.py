import os
import json
import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import sys

# === Auto-install missing dependencies ===
required_packages = {
    "pystray": "pystray",
    "PIL": "Pillow",
}

def install_missing_packages():
    for module_name, package_name in required_packages.items():
        try:
            __import__(module_name)
        except ImportError:
            print(f"Installing missing package: {package_name}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

install_missing_packages()

# === Config ===
DEFAULT_DRIVERS = ["Hyndavi", "Sweta", "Shubham", "Joyita", "Shashwat", "Aman"]
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))  # Script folder
WATCH_FOLDER = SCRIPT_DIR  # Input files here
CONFIG_FOLDER = os.path.join(SCRIPT_DIR, "config")  # Config folder
CONFIG_PATH = os.path.join(CONFIG_FOLDER, "driver_status.json")

# === Helpers ===
def get_current_week():
    today = datetime.date.today()
    return f"{today.isocalendar()[0]}-W{today.isocalendar()[1]:02}"

def load_config():
    if not os.path.exists(CONFIG_PATH):
        return {}
    with open(CONFIG_PATH, "r") as f:
        return json.load(f)

def save_config(config):
    # Ensure the config folder exists
    os.makedirs(CONFIG_FOLDER, exist_ok=True)
    
    with open(CONFIG_PATH, "w") as f:
        json.dump(config, f, indent=4)

# === GUI ===
def launch_gui():
    current_week = get_current_week()
    config = load_config()
    current_drivers = config.get(current_week)

    if not current_drivers:
        current_drivers = DEFAULT_DRIVERS[:]

    root = tk.Tk()
    root.title("BroomWagon Driver Selector")
    root.geometry("350x350")

    tk.Label(root, text=f"Select drivers for {current_week}", font=("Arial", 12, "bold")).pack(pady=10)

    checks = {}
    for driver in DEFAULT_DRIVERS:
        var = tk.BooleanVar(value=driver in current_drivers)
        cb = tk.Checkbutton(root, text=driver, variable=var, font=("Arial", 10))
        cb.pack(anchor="w", padx=20)
        checks[driver] = var

    def save_and_close():
        selected = [d for d, v in checks.items() if v.get()]
        if not selected:
            messagebox.showwarning("No Drivers Selected", "Please select at least one driver.")
            return
        config[current_week] = selected
        save_config(config)
        messagebox.showinfo("Saved", f"Drivers for {current_week} updated.")
        root.destroy()

    def reset_to_default():
        for driver in DEFAULT_DRIVERS:
            checks[driver].set(True)

    # Buttons
    ttk.Button(root, text="Save", command=save_and_close).pack(pady=10)
    ttk.Button(root, text="Reset to Default Drivers", command=reset_to_default).pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    launch_gui()
