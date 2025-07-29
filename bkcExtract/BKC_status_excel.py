#!/usr/intel/bin/python3.7.4

import os
from openpyxl import Workbook

def find_pythonsv_version_in_files(file_paths):
    version_info_dict = {}

    for file_path in file_paths:
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r') as file:
                    lines = file.readlines()
                    for line in lines:
                        if "pythonsv_version" in line:
                            server_name = file_path.split("\\")[2]  # Extract server name
                            version_info = line.strip()  # Extract version info
                            version_info_dict[server_name] = {"version_info": version_info}
                            break  # Stop after finding the first occurrence
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
        else:
            print(f"File does not exist: {file_path}")

    return version_info_dict

def find_latest_triplet_file(triplet_paths, version_info_dict):
    for path in triplet_paths:
        latest_file_name = None
        latest_time = 0
        server_name = path.split("\\")[2]  # Extract server name

        if os.path.exists(path):
            try:
                for file_name in os.listdir(path):
                    if file_name.startswith("Triplet"):
                        file_path = os.path.join(path, file_name)
                        file_time = os.path.getmtime(file_path)
                        if file_time > latest_time:
                            latest_time = file_time
                            latest_file_name = file_name
            except Exception as e:
                print(f"Error accessing {path}: {e}")
        else:
            print(f"Path does not exist: {path}")

        if latest_file_name:
            if server_name in version_info_dict:
                version_info_dict[server_name]["triplet_info"] = latest_file_name
            else:
                version_info_dict[server_name] = {"triplet_info": latest_file_name}

def print_combined_info(version_info_dict):
    print(f"{'Server':<20} {'Version Info':<40} {'Triplet Info':<60}")  # Table header
    print("=" * 120)  # Separator line

    for server, info in version_info_dict.items():
        version_info = info.get("version_info", "N/A")
        triplet_info = info.get("triplet_info", "N/A")
        print(f"{server:<20} {version_info:<40} {triplet_info:<60}")

def save_to_excel(version_info_dict, file_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Server Info"

    # Write headers
    sheet.append(["Server", "Version Info", "Triplet Info"])

    # Write data
    for server, info in version_info_dict.items():
        version_info = info.get("version_info", "N/A")
        triplet_info = info.get("triplet_info", "N/A")
        sheet.append([server, version_info, triplet_info])

    # Save the workbook
    workbook.save(file_path)
    print(f"Data saved to {file_path}")

# Example usage
version_file_paths = [
    "\\\\pg07tcmv0036\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0037\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0039\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0043\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0046\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0048\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0050\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0080\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0081\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0081\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0083\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0084\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0085\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0086\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0087\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0088\\c$\\pythonsv\\version.ini",
    "\\\\pg07tcmv0089\\c$\\pythonsv\\version.ini"
 
]

triplet_paths = [
    "\\\\pg07tcmv0036\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0037\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0039\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0043\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0046\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0048\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0050\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0080\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0081\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0082\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0083\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0084\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0085\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0086\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0087\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0088\\c$\\Intel\\Triplet_Logs",
    "\\\\pg07tcmv0089\\c$\\Intel\\Triplet_Logs"
]

version_info_dict = find_pythonsv_version_in_files(version_file_paths)
find_latest_triplet_file(triplet_paths, version_info_dict)
print_combined_info(version_info_dict)

# Define the path where the Excel file will be saved
excel_file_path = "I:\\mtl\\users\\ctio\\script\\Playground\\Server_Info.xlsx"
save_to_excel(version_info_dict, excel_file_path)