# Communication Report Maker

## Overview

Communication Report Maker is a Python application that analyzes OMSX and AMSX communication logs and generates a communication summary report for all devices found in the uploaded log files.

The tool matches Device IDs from communication packets with Serial Numbers from the Master Sheet and generates an Excel report containing communication statistics.

## Features

* Upload Master Sheet Excel file
* Upload up to 5 Communication Log Excel files
* Supports:

  * MOTOR_STATUS (OMSX)
  * ENVIRONMENT_NODE_STATUS (AMSX)
* Extracts Device ID directly from packet data
* Calculates:

  * Total Packets
  * Communication Count
  * No Communication Count
  * Communication Percentage
  * No Communication Percentage
* Generates dated Excel reports
* Search devices using Serial Number
* Displays communication statistics in console

## Installation

```bash
pip install pandas openpyxl
```

## Run

```bash
python packet_seperation.py
```

## Generated Report

```text
Communication_Report_YYYY-MM-DD.xlsx
```

Example:

```text
Communication_Report_2026-06-07.xlsx
```

## Report Columns

* Report Date
* Serial Number
* Device Type
* Device ID
* Total Packets
* Communication Count
* No Communication Count
* Communication %
* No Communication %

## Supported Device Types

### OMSX

```text
MOTOR_STATUS
```

### AMSX

```text
ENVIRONMENT_NODE_STATUS
```

## Communication Status

```text
1 = Communication Available
0 = No Communication
```

## Workflow

1. Upload Master Sheet
2. Upload up to 5 Communication Log files
3. Generate Communication Report
4. Search using Serial Number
5. View device communication statistics
6. Export results to Excel

## Output File

```text
Communication_Report_YYYY-MM-DD.xlsx
```

## Author

Karthick

Firmware QA Engineer
