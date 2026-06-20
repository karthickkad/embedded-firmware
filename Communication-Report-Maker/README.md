# OMS Communication & Valve Report Generator

## Overview

OMS Communication & Valve Report Generator is a Python-based automation tool used to analyze OMS communication logs and generate detailed reports for device communication status and valve operation status.

The tool reads:

* Master Sheet (Device Mapping)
* Communication Log Excel Files (up to 5 files)

and generates:

1. Communication Report
2. Valve Report

---

## Features

### Communication Report

* Device ID to Serial Number Mapping
* Total Packets Count
* Communication Count
* No Communication Count
* Communication Percentage
* No Communication Percentage
* Supports multiple communication log files

### Valve Report

* OMS Devices Only
* Channel 1 (Valve 1–4)
* Channel 2 (Valve 5–8)
* Valve Status

  * ON
  * COMPLETED
  * FAIL
  * SKIP
* Valve Start Time
* Configured Run Timer
* Valve Run Timer

### Search Mode

#### Communication Search

Search by Serial Number:

* Report Date
* Device ID
* Total Packets
* Communication Count
* No Communication Count
* Communication Percentage

#### Valve Search

Search by Serial Number:

* Valve 1–8 Status
* Valve Start Time
* Configured Run Timer
* Valve Run Timer

---

## Input Files

### Master Sheet

Required Columns:

| SERIAL .NO | DEVICE ID |
| ---------- | --------- |

Example:

| SERIAL .NO | DEVICE ID                            |
| ---------- | ------------------------------------ |
| OMSXSAM021 | 0ca8bb13-cd30-4596-81af-53e14ab9e151 |

---

### Communication Log

Required Columns:

| DATE-TIME | DEVICE ID | REFERENCE TYPE | MESSAGE |
| --------- | --------- | -------------- | ------- |

Header Row: 4

Packet Example:

```text
##,2f991702-a780-43af-a095-92eb5bf7dcc8,0,1780217335,...,[3-1-02:00-01:47-07:00],...,1,...
```

---

## Packet Parsing

### Device ID

The tool extracts Device ID from:

```python
fields[1]
```

instead of the Communication Log DEVICE ID column.

### Communication Status

The communication flag is extracted from:

```python
fields[-5]
```

Values:

* 1 = Communication
* 0 = No Communication

### Valve Format

Example:

```text
3-1-02:00-01:47-07:00
```

Meaning:

| Position | Description          |
| -------- | -------------------- |
| 1        | Valve Number         |
| 2        | Valve Status         |
| 3        | Configured Run Timer |
| 4        | Valve Run Timer      |
| 5        | Valve Start Time     |

### Valve Status Logic

* ON → Valve Status = 1
* COMPLETED → Configured Run Timer = Valve Run Timer
* FAIL → Configured Run Timer ≠ Valve Run Timer
* SKIP → Valve not available for the selected day

---

## Generated Reports

### Communication Report

File Name:

```text
L&T_Communication_Report_YYYY-MM-DD.xlsx
```

### Valve Report

File Name:

```text
YYYY-MM-DD_Valve_Report.xlsx
```

Valve Report is generated only for OMS devices.

---

## Requirements

Install dependencies:

```bash
pip install pandas openpyxl
```

---

## Run

```bash
python OMS_Valve_Communication_Report_Generator_v4.py
```

---

## Author

Karthick

Firmware QA Engineer
