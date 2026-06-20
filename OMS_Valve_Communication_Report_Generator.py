
# OMS Valve & Communication Report Generator

import pandas as pd
import os
import re
from datetime import datetime
from tkinter import Tk, filedialog

def battery_health(voltage):
    if voltage < 12.0:
        return "CRITICAL"
    elif voltage < 12.5:
        return "LOW"
    elif voltage <= 13.5:
        return "GOOD"
    return "HIGH"

def extract_valves(packet):
    valves = []
    m = re.search(r"\[(.*?)\]", packet)
    if not m:
        return valves

    content = m.group(1)
    for item in content.split(","):
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

root = Tk()
root.withdraw()

master_file = filedialog.askopenfilename(
    title="Select Master Sheet",
    filetypes=[("Excel Files","*.xlsx *.xls")]
)

if not master_file:
    raise SystemExit("Master sheet not selected")

log_files = filedialog.askopenfilenames(
    title="Select Communication Logs (max 5)",
    filetypes=[("Excel Files","*.xlsx *.xls")]
)

if not log_files:
    raise SystemExit("No logs selected")

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
    df = pd.read_excel(f, header=3)
    df.columns = df.columns.astype(str).str.strip()
    all_logs.append(df)

log_df = pd.concat(all_logs, ignore_index=True)

report_date = pd.to_datetime(
    log_df["DATE-TIME"].iloc[0],
    dayfirst=True
).strftime("%d-%m-%Y")

# Communication report
device_summary = {}
valve_summary = {}

for _, row in log_df.iterrows():
    packet = str(row.get("MESSAGE","")).strip()
    if not packet or packet.lower() == "nan":
        continue

    fields = [x.strip() for x in packet.split(",")]
    if len(fields) < 10:
        continue

    device_id = str(row["DEVICE ID"]).strip()

    if device_id not in device_summary:
        device_summary[device_id] = {
            "total":0,
            "comm":0,
            "no_comm":0
        }

    device_summary[device_id]["total"] += 1

    try:
        comm_status = int(fields[-1].replace("$",""))
    except:
        comm_status = 1

    if comm_status:
        device_summary[device_id]["comm"] += 1
    else:
        device_summary[device_id]["no_comm"] += 1

    if device_id not in valve_summary:
        valve_summary[device_id] = {
            "battery":[],
            "solar":[],
            "valves":{i:None for i in range(1,9)}
        }

    try:
        battery = int(fields[-9]) / 10
        solar = int(fields[-8]) / 10
        valve_summary[device_id]["battery"].append(battery)
        valve_summary[device_id]["solar"].append(solar)
    except:
        pass

    for v in extract_valves(packet):
        valve_summary[device_id]["valves"][v["valve_no"]] = v

comm_rows = []
for dev, s in device_summary.items():
    total = s["total"]
    comm_rows.append({
        "Report Date": report_date,
        "Serial Number": device_to_serial.get(dev,"UNKNOWN"),
        "Device ID": dev,
        "Total Packets": total,
        "Communication Count": s["comm"],
        "No Communication Count": s["no_comm"],
        "Communication %": round((s["comm"]/total)*100,2) if total else 0
    })

comm_df = pd.DataFrame(comm_rows)
comm_file = f"{report_date}_Communication_Report.xlsx"
comm_df.to_excel(comm_file,index=False)

# Valve report (Option B layout)
devices = sorted(valve_summary.keys(), key=lambda x: device_to_serial.get(x,x))

rows = []

for valve_no in range(1,9):
    channel = "CH1" if valve_no <= 4 else "CH2"

    status_row = {"CHANNEL":channel, "ITEM":f"Valve {valve_no}"}
    ast_row = {"CHANNEL":"", "ITEM":"AST"}
    art_row = {"CHANNEL":"", "ITEM":"ART"}

    for dev in devices:
        v = valve_summary[dev]["valves"][valve_no]

        col = device_to_serial.get(dev, dev)

        if v is None:
            status_row[col] = "SKIP"
            ast_row[col] = "-"
            art_row[col] = "-"
        else:
            if v["status"] == 1:
                status = "ON"
            elif v["art"] == v["set_time"]:
                status = "COMPLETED"
            else:
                status = "FAIL"

            status_row[col] = status
            ast_row[col] = v["ast"]
            art_row[col] = v["art"]

    rows.extend([status_row, ast_row, art_row])

battery_row = {"CHANNEL":"","ITEM":"Avg Battery"}
solar_row = {"CHANNEL":"","ITEM":"Avg Solar"}
health_row = {"CHANNEL":"","ITEM":"Battery Health"}

for dev in devices:
    col = device_to_serial.get(dev, dev)

    bats = valve_summary[dev]["battery"]
    sols = valve_summary[dev]["solar"]

    avg_bat = round(sum(bats)/len(bats),2) if bats else 0
    avg_sol = round(sum(sols)/len(sols),2) if sols else 0

    battery_row[col] = avg_bat
    solar_row[col] = avg_sol
    health_row[col] = battery_health(avg_bat)

rows.extend([battery_row, solar_row, health_row])

valve_df = pd.DataFrame(rows)
valve_file = f"{report_date}_Valve_Report.xlsx"
valve_df.to_excel(valve_file,index=False)

print("Generated:")
print(comm_file)
print(valve_file)
