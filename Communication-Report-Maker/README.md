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
  * SKIP
* Actual Start Time (AST)
* Actual Run Time (ART)

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
* AST
* ART

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

##,2f991702-a780-43af-a095-92eb5bf7dcc8,0,1780217335,...,[3-1-02:00-01:47-07:00],...,1,...

---

## Packet Parsing

### Device ID

The tool extracts Device ID from:

fields[1]

instead of the Communication Log DEVICE ID column.

### Communication Status

The communication flag is extracted from:

fields[-5]

Values:

* 1 = Communication
* 0 = No Communication

### Valve Format

Example:

3-1-02:00-01:47-07:00

Meaning:

* Valve Number = 3
* Valve Status = 1
* Actual Run Time = 02:00
* Set Time = 01:47
* Actual Start Time = 07:00

---

## Generated Reports

### Communication Report

File Name:

L&T_Communication_Report_YYYY-MM-DD.xlsx

### Valve Report

File Name:

YYYY-MM-DD_Valve_Report.xlsx

---

## Requirements

Install dependencies:

pip install pandas openpyxl

---

## Run

python OMS_Valve_Communication_Report_Generator_v3.py

---

## Author

Karthick

Firmware QA Engineer
