
# OMS Valve & Communication Report Generator
# Version 3 - Device ID Fix + Communication Status Fix + Search Mode

import pandas as pd
import os
import re
from tkinter import Tk, filedialog

def extract_valves(packet):
    valves = []
    m = re.search(r"\[(.*?)\]", packet)
    if not m:
        return valves

    for item in m.group(1).split(","):
        parts = item.strip().split("-")
        if len(parts) >= 5:
            try:
                valves.append({
                    "valve_no": int(parts[0]),
                    "status": int(parts[1]),
                    "art": parts[2],
                    "set_time": parts[3],
                    "ast": parts[4]
                })
            except:
                pass
    return valves

try:

    root = Tk()
    root.withdraw()

    print("-" * 70)
    print("SELECT MASTER SHEET")
    print("-" * 70)

    master_file = filedialog.askopenfilename(
        title="Select Master Sheet",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not master_file:
        raise SystemExit("Master Sheet Not Selected")

    print(master_file)

    log_files = filedialog.askopenfilenames(
        title="Select Communication Logs (Max 5)",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not log_files:
        raise SystemExit("No Communication Logs Selected")

    master_df = pd.read_excel(master_file)
    master_df.columns = master_df.columns.astype(str).str.strip()

    serial_col = "SERIAL .NO"
    device_col = "DEVICE ID"

    device_to_serial = dict(
        zip(
            master_df[device_col].astype(str).str.strip(),
            master_df[serial_col].astype(str).str.strip()
        )
    )

    all_logs = []

    for f in log_files:

        try:

            print(f"Loaded: {os.path.basename(f)}")

            df = pd.read_excel(
                f,
                header=3
            )

            df.columns = (
                df.columns.astype(str)
                .str.strip()
            )

            all_logs.append(df)

        except Exception as e:

            print(f"Error loading {f}")
            print(e)

    if len(all_logs) == 0:
        raise SystemExit("No Valid Logs Found")

    log_df = pd.concat(
        all_logs,
        ignore_index=True
    )

    print(f"\nTotal Records Loaded: {len(log_df)}")

    report_date = pd.to_datetime(
        log_df["DATE-TIME"].iloc[0],
        dayfirst=True
    ).strftime("%Y-%m-%d")

    device_summary = {}
    valve_summary = {}

    for _, row in log_df.iterrows():

        packet = str(
            row.get("MESSAGE", "")
        ).strip()

        if not packet or packet.lower() == "nan":
            continue

        fields = [
            x.strip()
            for x in packet.split(",")
        ]

        if len(fields) < 5:
            continue

        try:
            packet_device_id = fields[1].strip()
        except:
            continue

        if packet_device_id not in device_to_serial:
            continue

        device_id = packet_device_id
        
        serial_no = device_to_serial.get(
            device_id,
            ""
        )
        
        #communication report (OMSX + AMSX)

        if device_id not in device_summary:

            device_summary[device_id] = {
                "total": 0,
                "comm": 0,
                "no_comm": 0
            }

        device_summary[device_id]["total"] += 1

        try:
            communication_status = int(fields[-5])
        except:
            communication_status = 0

        if communication_status == 1:
            device_summary[device_id]["comm"] += 1
        else:
            device_summary[device_id]["no_comm"] += 1
        
        #valve report (OMSX only)
        if serial_no.startswith("OMSX"):
            
            if device_id not in valve_summary:
                
                valve_summary[device_id] = {
                    "valves": {
                        i: None for i in range(1, 9)
                        }
                }
            
            for valve in extract_valves(packet):
                valve_summary[device_id]["valves"][valve["valve_no"]] = valve
            
    report_rows = []

    for device_id, stats in device_summary.items():

        total_packets = stats["total"]
        comm_count = stats["comm"]
        no_comm_count = stats["no_comm"]

        comm_pct = round(
            (comm_count / total_packets) * 100, 2
        ) if total_packets else 0

        no_comm_pct = round(
            (no_comm_count / total_packets) * 100, 2
        ) if total_packets else 0

        report_rows.append({

            "Report Date": report_date,
            "Serial Number": device_to_serial.get(device_id, "UNKNOWN"),
            "Device ID": device_id,
            "Total Packets": total_packets,
            "Communication Count": comm_count,
            "No Communication Count": no_comm_count,
            "Communication %": comm_pct,
            "No Communication %": no_comm_pct
        })

    comm_df = pd.DataFrame(report_rows)

    comm_file = f"{report_date}_communication.xlsx"
    comm_df.to_excel(comm_file, index=False)

    devices = sorted(
        valve_summary.keys(),
        key=lambda x: device_to_serial.get(x, x)
    )

    rows = []

    for valve_no in range(1, 9):

        channel = "CH1" if valve_no <= 4 else "CH2"

        status_row = {
            "CHANNEL": channel,
            "ITEM": f"Valve {valve_no}"
        }

        ast_row = {
            "CHANNEL": "",
            "ITEM": "AST"
        }

        art_row = {
            "CHANNEL": "",
            "ITEM": "ART"
        }

        for dev in devices:

            serial = device_to_serial.get(dev, dev)
            valve = valve_summary[dev]["valves"][valve_no]

            if valve is None:

                status_row[serial] = "SKIP"
                ast_row[serial] = "-"
                art_row[serial] = "-"

            else:

                status_row[serial] = (
                    "ON"
                    if valve["status"] == 1
                    else "COMPLETED"
                )

                ast_row[serial] = valve["ast"]
                art_row[serial] = valve["art"]

        rows.extend([status_row, ast_row, art_row])

    valve_df = pd.DataFrame(rows)

    valve_file = f"{report_date}_Valve_Report.xlsx"
    valve_df.to_excel(valve_file, index=False)

    print("\n" + "-" * 70)
    print("REPORT GENERATED SUCCESSFULLY")
    print("-" * 70)

    print(os.path.abspath(comm_file))
    print(os.path.abspath(valve_file))

    while True:

        print("\n" + "-" * 90)
        print("ENTERED TO SEARCH MODE")
        print("-" * 90)
        print("1. Communication Report")
        print("2. Valve Report")
        print("3. Exit")

        choice = input("\nEnter Choice : ").strip()

        if choice == "1":

            serial_no = input(
                "\nEnter Serial Number : "
            ).strip()

            result = comm_df[
                comm_df["Serial Number"] == serial_no
            ]

            if result.empty:
                print("Serial Number Not Found!")
                continue

            row = result.iloc[0]

            print("\n" + "-" * 90)
            print("DEVICE COMMUNICATION REPORT")
            print("-" * 90)

            for col in result.columns:
                print(f"{col:30} : {row[col]}")

            print("-" * 90)

        elif choice == "2":

            serial_no = input(
                "\nEnter Serial Number : "
            ).strip()

            device_id = None

            for dev, ser in device_to_serial.items():
                if ser == serial_no:
                    device_id = dev
                    break

            if device_id is None or device_id not in valve_summary:
                print("Serial Number Not Found!")
                continue

            print("\n" + "-" * 90)
            print("DEVICE VALVE REPORT")
            print("-" * 90)

            for valve_no in range(1, 9):

                valve = valve_summary[device_id]["valves"][valve_no]

                if valve is None:

                    print(f"\nValve {valve_no} : SKIP")

                else:

                    status = (
                        "ON"
                        if valve["status"] == 1
                        else "COMPLETED"
                    )

                    print(f"\nValve {valve_no}")
                    print(f"Status : {status}")
                    print(f"AST    : {valve['ast']}")
                    print(f"ART    : {valve['art']}")

            print("-" * 90)

        elif choice == "3":

            print("Program Closed")
            break

        else:

            print("Invalid Choice")

except Exception as e:

    print("\nPROGRAM FAILED")
    print(str(e))
    input("Press Enter To Exit...")
