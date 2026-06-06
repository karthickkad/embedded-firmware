import pandas as pd
import os
from datetime import datetime
from tkinter import Tk, filedialog

# --------------------------------------------------
# FILE SELECTION
# --------------------------------------------------
root = Tk()
root.withdraw()

print("Select Master Sheet")

master_file = filedialog.askopenfilename(
    title="Select Master Sheet",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not master_file:
    print("Master Sheet not selected!")
    exit()

print("Select up to 5 Communication Log Files")

log_files = filedialog.askopenfilenames(
    title="Select Communication Log Files",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not log_files:
    print("No Communication Log Files Selected!")
    exit()

if len(log_files) > 5:
    print("Maximum 5 files allowed!")
    exit()

# --------------------------------------------------
# READ MASTER SHEET
# --------------------------------------------------
master_df = pd.read_excel(master_file)

master_df.columns = (
    master_df.columns.astype(str)
    .str.strip()
)

serial_col = "SERIAL .NO"
device_col = "DEVICE ID"

master_df[serial_col] = (
    master_df[serial_col]
    .astype(str)
    .str.strip()
)

master_df[device_col] = (
    master_df[device_col]
    .astype(str)
    .str.strip()
)

device_to_serial = dict(
    zip(
        master_df[device_col],
        master_df[serial_col]
    )
)

# --------------------------------------------------
# READ COMMUNICATION LOGS
# --------------------------------------------------
all_logs = []

for file in log_files:

    try:

        temp_df = pd.read_excel(
            file,
            header=3
        )

        temp_df.columns = (
            temp_df.columns.astype(str)
            .str.strip()
        )

        all_logs.append(temp_df)

        print(
            f"Loaded: {os.path.basename(file)}"
        )

    except Exception as e:

        print(
            f"Error loading file: {file}"
        )

        print(e)

if len(all_logs) == 0:
    print("No valid communication logs!")
    exit()

log_df = pd.concat(
    all_logs,
    ignore_index=True
)

print(
    f"\nTotal Records Loaded: {len(log_df)}"
)

# --------------------------------------------------
# CREATE SUMMARY
# --------------------------------------------------
device_summary = {}

for _, row in log_df.iterrows():

    try:

        packet = str(
            row["MESSAGE"]
        ).strip()

        if (
            packet == ""
            or packet.lower() == "nan"
        ):
            continue

        fields = [
            x.strip()
            for x in packet.split(",")
        ]

        if len(fields) < 20:
            continue

        packet_device_id = fields[1]

        try:
            comm_status = int(
                fields[-5]
            )
        except:
            comm_status = 0

        device_type = str(
            row["REFERENCE TYPE"]
        ).strip()

        if packet_device_id not in device_summary:

            device_summary[
                packet_device_id
            ] = {

                "device_type":
                    device_type,

                "total_packets":
                    0,

                "communication_count":
                    0,

                "no_communication_count":
                    0
            }

        device_summary[
            packet_device_id
        ]["total_packets"] += 1

        if comm_status == 1:

            device_summary[
                packet_device_id
            ]["communication_count"] += 1

        else:

            device_summary[
                packet_device_id
            ]["no_communication_count"] += 1

    except:
        continue

# --------------------------------------------------
# GENERATE REPORT
# --------------------------------------------------
today = datetime.now().strftime(
    "%Y-%m-%d"
)

report_rows = []

for device_id, stats in device_summary.items():

    serial_no = device_to_serial.get(
        device_id,
        "UNKNOWN"
    )

    total_packets = stats[
        "total_packets"
    ]

    communication_count = stats[
        "communication_count"
    ]

    no_communication_count = stats[
        "no_communication_count"
    ]

    if total_packets > 0:

        communication_percentage = (
            communication_count
            / total_packets
        ) * 100

        no_communication_percentage = (
            no_communication_count
            / total_packets
        ) * 100

    else:

        communication_percentage = 0

        no_communication_percentage = 0

    report_rows.append({

        "Report Date":
            today,

        "Serial Number":
            serial_no,

        "Device ID":
            device_id,

        "Total Packets":
            total_packets,

        "Communication Count":
            communication_count,

        "No Communication Count":
            no_communication_count,

        "Communication %":
            round(
                communication_percentage,
                2
            ),

        "No Communication %":
            round(
                no_communication_percentage,
                2
            )
    })

report_df = pd.DataFrame(
    report_rows
)

report_df = report_df.sort_values(
    by="Serial Number"
)

report_file = (
    f"L&T_Communication_Report_{today}.xlsx"
)

report_df.to_excel(
    report_file,
    index=False
)

print("\n")
print("=" * 70)
print("L&T COMMUNICATION REPORT GENERATED")
print("=" * 70)
print(f"Report File : {report_file}")
print(f"Total Devices : {len(report_df)}")
print("=" * 70)

# --------------------------------------------------
# SEARCH MODE
# --------------------------------------------------
while True:

    serial_no = input(
        "\nEnter Serial Number (or EXIT): "
    ).strip()

    if serial_no.upper() == "EXIT":

        print("\nProgram Closed")
        break

    result = report_df[
        report_df["Serial Number"]
        == serial_no
    ]

    if result.empty:

        print(
            "\nSerial Number Not Found!"
        )

        continue

    row = result.iloc[0]

    print("\n")
    print("=" * 90)
    print("DEVICE COMMUNICATION REPORT")
    print("=" * 90)

    print(f"{'Report Date':30} : {row['Report Date']}")
    print(f"{'Serial Number':30} : {row['Serial Number']}")
    print(f"{'Device ID':30} : {row['Device ID']}")
    print(f"{'Total Packets':30} : {row['Total Packets']}")
    print(f"{'Communication Count':30} : {row['Communication Count']}")
    print(f"{'No Communication Count':30} : {row['No Communication Count']}")
    print(f"{'Communication Percentage':30} : {row['Communication %']} %")
    print(f"{'No Communication Percentage':30} : {row['No Communication %']} %")

    print("=" * 90)