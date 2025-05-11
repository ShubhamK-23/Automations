
# 🛠️ Automations

This repository contains personal productivity scripts to automate repetitive tasks, including:

- 🔐 **KeePass Auto-Opener** – securely manages and auto-opens KeePass databases with saved credentials.
- 📋 **Broomwagon Automation** – processes ticket reports and distributes them among drivers with reminders.

---

## 🔐 KeePass Auto-Opener

A secure utility to store encrypted KeePass credentials and auto-open databases via the KeePass executable.

### Features
- Interactive setup for KeePass database passwords.
- Passwords are AES-encrypted using Python’s `cryptography` module.
- Opens KeePass and unlocks selected databases automatically.

### Usage
```bash
python keepass_autoopener/keepass_auto_opener.py
```

### Setup
- On first run, the script will prompt you for passwords for each database.
- Encrypted configuration is saved locally and not committed to the repository.

---

## 📋 Broomwagon Automation

Processes ticket reports from the OTRS platform and distributes them among a pool of drivers, with reminders and tracking.

### Features
- Adds tracking and internal action columns to the report.
- Distributes tickets across drivers based on availability and rotation.
- Avoids reprocessing files already handled.
- Includes a simple GUI tool to manage weekly driver availability.
- Renames processed files for clarity and traceability.
- Logs all processing steps and errors.

### Usage
To process the BroomWagon file:
```bash
python broomwagon_automation/Broomwagon_Automation.py
```

To configure weekly driver availability:
```bash
python broomwagon_automation/Broomwagon Driver.py
```

---

## 📁 Repository Structure

```
Automations/
│
├── keepass_autoopener/
│   └── keepass_auto_opener.py
│
├── broomwagon_automation/
│   ├── Broomwagon_Automation.py
│   └── Broomwagon Driver.py
│
└── README.md
```

---

## 📌 Notes

- Scripts are tailored for internal use cases and Windows environments.
- KeePass Auto-Opener uses a default path:  
  `C:\Program Files\KeePass Password Safe 2\KeePass.exe`
---

## 🔒 Security Notice

Passwords are never stored in plain text. The `SecureConfigManager` module uses a local encryption key stored with restricted file permissions. Ensure proper OS-level security is in place to protect this key file.

---
