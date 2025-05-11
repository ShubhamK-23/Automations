
# BroomWagon Automation Script

This script automates the processing of BroomWagon Excel files, assigning tickets to drivers, and logging the processed files. It helps automate ticket distribution based on weekly driver availability.

## Features

- **Ticket Distribution**: Automatically distributes tickets to drivers, with a weekly rotation system.
- **Driver Availability Configuration**: Drivers are dynamically assigned based on a weekly configuration stored in a `driver_status.json` file.
- **Excel File Formatting**: The script adjusts the formatting of the Excel file, such as column widths, font styles, and cell colors.
- **Logging**: All actions are logged in a dedicated log file for traceability.
- **Automatic Dependency Installation**: Missing dependencies will be automatically installed upon script execution.

## Prerequisites

- Python 3.x
- The following Python libraries:
  - `pandas`
  - `openpyxl`
  - `win10toast`

These libraries are automatically installed if they are not already present.

## Installation

### Step 1: Download the Script

1. Download the `Broomwagon_Automation.py` and  `Broomwagon Driver.py` scripts and save it in a directory on your computer.

### Step 2: Install Dependencies

The script will automatically check for and install missing dependencies when you run it.

### Step 3: Configuration Files

- The script will **automatically create the `config` folder** if it doesn't exist.
- The `driver_status.json` and `processed_files.log` files will be automatically created in the `config` folder when the script runs for the first time.
- The `driver_status.json` file contains weekly driver availability data, and the `processed_files.log` file keeps track of files that have already been processed.

### Step 4: Prepare the Files to Process

1. Place the `BroomWagon_WK-X_XX.xlsx` BroomWagon files you want to process into the same directory as the script.

## Usage

1. **Run the Script**: 

   Open a terminal or command prompt, navigate to the script's directory, and run the following command:

   `python Broomwagon_Automation.py`

2. **Update Weekly Driver Availability**:

   The script will ask if you would like to update the weekly driver availability. If you select "yes", it will attempt to run a `Broomwagon Driver.py` script (if available) to update the driver configuration. If the driver script isn't found, it will continue without updating.

3. **Processing Files**:

   The script will automatically process all `.xlsx` files in the `WATCH_FOLDER`. For each file:
   - It will validate that the necessary columns exist.
   - It will assign tickets to drivers based on the current week’s configuration and rotation.
   - It will adjust the Excel file's formatting, including adding new columns and setting colors for drivers.
   - After processing, it will log the file as processed and delete the backup file.

   Processed files will be saved in the `OUTPUT` folder.

4. **Manual Trigger**:

   The script runs interactively, so you can trigger the process whenever needed. It will prompt you for confirmation before proceeding with the driver availability update and file processing.

## File Structure

```
/script-directory/
│
├── Broomwagon_Automation.py   # Main script for processing files
├── config/
│   ├── driver_status.json     # Weekly driver availability configuration (created automatically)
│   └── processed_files.log    # Log of processed files (created automatically)
├── logs/                      # Logs folder for the script's logs
│
└── Output/                    # Processed files will be saved here
```

## Customizing the Script

1. **Customizing Driver Rotation**:  
   You can customize the driver rotation and assignments by modifying the `driver_status.json` file to match your team's schedule.

2. **Updating Column Widths**:  
   You can modify the `special_columns` dictionary in the script to adjust the column widths for specific columns like 'Ticket action', 'Internal action', etc.

## Troubleshooting

- **Missing Columns**: If the script encounters a file that doesn't have all the expected columns, it will log an error and skip processing that file.
  
- **File Already Processed**: If a file has already been processed, the script will skip it and log the event.

- **Dependency Issues**: The script attempts to install missing Python libraries automatically. Ensure that you have internet access if dependencies need to be installed.

## License

This script is provided as-is and can be freely modified and used for internal purposes. Please note that it is intended to be used for automating the ticket distribution process and might require adaptation for specific use cases.
