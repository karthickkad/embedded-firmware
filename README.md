# OMS Valve & Communication Report Generator

![Python](https://img.shields.io/badge/Python-3.11+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Version](https://img.shields.io/badge/Version-v5.0-orange.svg)

A Python-based desktop utility for analyzing OMS and AMS communication logs, generating communication statistics, extracting valve runtime information, and exporting professionally formatted Excel reports.

Designed for **Firmware QA**, **Embedded Validation**, and **Communication Log Analysis**.

---

# Features

- Load Master Sheet
- Support up to **5 Communication Log** files
- Automatic Master Sheet column detection
- Duplicate packet removal
- Packet validation
- Communication statistics calculation
- Valve runtime extraction
- Communication Report generation
- Valve Report generation
- Separate OMS and AMS communication reports
- Search devices using Serial Number
- Communication Report Search
- Valve Report Search
- DEVICE ID displayed in Search Mode
- Automatic Excel formatting
- Auto column sizing
- Freeze panes
- Duplicate Device ID detection
- Duplicate Serial Number detection
- Processing summary

---

# Requirements

- Python 3.11 or later

Install the required packages:

```bash
pip install pandas openpyxl
```

Tkinter is included with the standard Python installation on Windows.

---

# Project Structure

```
OMS_Valve_Communication_Report_Generator_v5.py
README.md
requirements.txt
LICENSE
.gitignore
```

---

# Master Sheet Format

The Master Sheet must contain the following columns.

| Column |
|----------|
| SERIAL .NO |
| DEVICE ID |

Example:

| SERIAL .NO | DEVICE ID |
|------------|-----------|
| OMSXSAM021 | 0ca8bb13-cd30-4596-81af-53e14ab9e151 |
| OMSX002432 | 41915c4c-6519-4467-908e-084801abc9d9 |
| AMSX001044 | f0c22892-2909-4370-beaf-37b1e6d77538 |

The program automatically:

- Detects Serial Number column
- Detects Device ID column
- Ignores blank rows
- Detects duplicate Serial Numbers
- Detects duplicate Device IDs

---

# Communication Log Format

Supported file types

- `.xlsx`
- `.xls`

Required columns

| Column |
|----------|
| DATE-TIME |
| MESSAGE |

The Communication Log header should begin at **Row 4**.

Maximum supported communication log files:

```
5
```

---

# Running the Program

Execute:

```bash
python OMS_Valve_Communication_Report_Generator_v5.py
```

The application will guide you through:

1. Selecting the Master Sheet
2. Selecting Communication Log Files
3. Report Generation
4. Search Mode

---

# Generated Reports

## Communication Report

Contains

- Report Date
- Serial Number
- Total Packets
- Communication Count
- No Communication Count
- Communication Percentage
- No Communication Percentage

Generated file:

```
DD-MM-YYYY_communication.xlsx
```

---

## Valve Report

Generated only for OMS devices.

Contains

- Valve Status
- Valve Start Time
- Configured Run Timer
- Valve Run Timer

Generated file:

```
DD-MM-YYYY_Valve_Report.xlsx
```

---

# Search Mode

## Communication Report Search

Search using a complete or partial Serial Number.

Displays

- DEVICE ID
- Report Date
- Serial Number
- Total Packets
- Communication Count
- No Communication Count
- Communication Percentage
- No Communication Percentage

---

## Valve Report Search

Search using a complete or partial Serial Number.

Displays

- DEVICE ID
- Valve Status
- Valve Start Time
- Configured Run Timer
- Valve Run Timer

---

# Supported Devices

- OMS Devices
- AMS Devices

---

# Packet Processing

The application performs:

- Packet validation
- Device ID verification
- Communication status calculation
- Valve runtime extraction
- Duplicate packet removal
- Communication statistics generation
- Excel report generation

---

# Technologies Used

- Python
- Pandas
- OpenPyXL
- Tkinter

---

# Version History

## Version 5.0

### Added

- Support for multiple communication log files
- Communication Report Search
- Valve Report Search
- DEVICE ID display in Search Mode
- Automatic Master Sheet column detection
- Duplicate packet removal
- Improved valve runtime extraction
- Excel formatting improvements
- Processing summary

### Improved

- Packet validation
- Communication statistics
- Valve status calculation
- Report formatting
- Search performance
- Code documentation

---

# License

This project is licensed under the MIT License.

---

# Author

**Karthick**

Firmware QA Engineer

Specialized in Embedded Systems, Firmware Validation, Communication Log Analysis, and Test Automation.

---

## Future Enhancements

- Packet format configuration
- CSV export support
- PDF report generation
- Command-line mode
- Performance optimization for large datasets
- Automated report scheduling