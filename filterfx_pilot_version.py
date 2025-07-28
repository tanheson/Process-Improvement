import os
import shutil
import time
from datetime import datetime

# Define multiple source folders (network drives) and a single destination
source_folders = [
    r"//PG07TCMV0080/c$/Results/ARL/S681/A0", 
    r"//PG07TCMV0081/c$/Results/ARL/S681/A0", 
    r"//PG07TCMV0082/c$/Results/ARL/S681/A0", 
    r"//PG07TCMV0083/c$/Results/ARL/S681/A0", 
    r"//PG07TCMV0084/c$/Results/ARL/S681/A0", 
    r"//PG07TCMV0085/c$/Results/ARL/S681/A0", 
    r"//PG07TCMV0086/c$/Results/ARL/S681/A0", 
    r"//PG07TCMV0088/c$/Results/ARL/S681/A0", 
    r"C:/Results/NVL/HX/A0"  # Local drive as an example
]
destination_folder = r"U:/NVL/HX/A0/results_production"

# Required file keyword to check
required_file_keyword = "HotVmin.xlsx"

def get_latest_hotvmin_file(folder_path):
    """Check if a folder contains a file with the required keyword and get the latest modified file."""
    hotvmin_files = []
    for root, _, files in os.walk(folder_path):
        hotvmin_files.extend([os.path.join(root, f) for f in files if required_file_keyword in f])
    if not hotvmin_files:
        return None, False  # No "HotVmin.xlsx" file found
    latest_file = max(hotvmin_files, key=os.path.getmtime, default=None)
    return latest_file, True  # Return latest file path and flag indicating presence

def parse_timestamp_folder_name(folder_name):
    """Parse timestamp from folder name (e.g., '2025.07.25_19.28.10') and return datetime object."""
    try:
        return datetime.strptime(folder_name, "%Y.%m.%d_%H.%M.%S")
    except ValueError:
        return None

def get_latest_timestamp_folder(hotvmin_path):
    """Get the latest timestamp folder under a HotVmin folder."""
    timestamp_folders = []
    for item in os.listdir(hotvmin_path):
        item_path = os.path.join(hotvmin_path, item)
        if os.path.isdir(item_path):
            timestamp = parse_timestamp_folder_name(item)
            if timestamp:
                timestamp_folders.append((item_path, timestamp))
    if not timestamp_folders:
        return None
    return max(timestamp_folders, key=lambda x: x[1])[0]  # Return path of latest folder

try:
    # Check if destination parent directory exists
    destination_parent = os.path.dirname(destination_folder)
    if not os.path.exists(destination_parent):
        raise OSError(f"Parent directory '{destination_parent}' does not exist. Please create it first.")

    # Create destination folder
    os.makedirs(destination_folder, exist_ok=True)

    # Process each source folder
    for source_folder in source_folders:
        # Check if source folder exists
        if not os.path.isdir(source_folder):
            print(f"Warning: Source folder '{source_folder}' does not exist. Skipping...")
            continue

        print(f"Processing source folder: {source_folder}")

        # Traverse the source folder
        for root, dirs, files in os.walk(source_folder):
            # Calculate the relative path from source_folder
            relative_path = os.path.relpath(root, source_folder)
            dest_path = os.path.join(destination_folder, relative_path)

            # Filter folders under "D3" and "D4"
            if "D3" in relative_path or "D4" in relative_path:
                # Check each directory in dirs
                dirs_to_keep = []
                for dir_name in dirs:
                    dir_full_path = os.path.join(root, dir_name)
                    if "HotVmin" in dir_name:
                        # Get the latest timestamp folder
                        latest_timestamp_folder = get_latest_timestamp_folder(dir_full_path)
                        if not latest_timestamp_folder:
                            print(f"Ignoring folder '{dir_full_path}' as it lacks valid timestamp folders.")
                            continue
                        # Check for HotVmin.xlsx in the latest timestamp folder
                        latest_file, has_hotvmin = get_latest_hotvmin_file(latest_timestamp_folder)
                        if not has_hotvmin:
                            print(f"Ignoring folder '{dir_full_path}' as the latest timestamp folder lacks a '{required_file_keyword}' file.")
                            continue
                        print(f"Keeping folder '{dir_full_path}' with latest timestamp folder '{os.path.basename(latest_timestamp_folder)}' and '{os.path.basename(latest_file)}' (modified: {time.ctime(os.path.getmtime(latest_file))})")
                        dirs_to_keep.append(dir_name)  # Keep the HotVmin folder
                    else:
                        dirs_to_keep.append(dir_name)  # Keep non-"HotVmin" folders
                # Update dirs to only include folders to keep (modifies os.walk behavior)
                dirs[:] = dirs_to_keep

                # Filter timestamp folders within HotVmin
                if "HotVmin" in os.path.basename(root):
                    latest_timestamp_folder = get_latest_timestamp_folder(root)
                    if latest_timestamp_folder:
                        # Keep only the latest timestamp folder
                        timestamp_dirs_to_keep = []
                        for dir_name in dirs:
                            dir_full_path = os.path.join(root, dir_name)
                            if dir_full_path == latest_timestamp_folder:
                                timestamp_dirs_to_keep.append(dir_name)
                            else:
                                print(f"Ignoring timestamp folder '{dir_full_path}' as it is not the latest.")
                        dirs[:] = timestamp_dirs_to_keep  # Restrict to only the latest timestamp folder

            # Copy files and create directories
            for file in files:
                source_path = os.path.join(root, file)
                destination_path = os.path.join(dest_path, file)
                if not "HotVmin" in os.path.basename(root):  # Avoid copying files directly under HotVmin
                    os.makedirs(os.path.dirname(destination_path), exist_ok=True)
                    shutil.copy2(source_path, destination_path)
                    print(f"Copied: {file} to {destination_path}")
                # Handle copying from the latest timestamp folder under HotVmin
                if "HotVmin" in os.path.basename(os.path.dirname(root)) and os.path.basename(root).replace(".", "_").replace(":", "").isdigit():
                    latest_timestamp_folder = get_latest_timestamp_folder(os.path.dirname(root))
                    if root.startswith(latest_timestamp_folder):
                        if required_file_keyword in file:
                            latest_file, _ = get_latest_hotvmin_file(root)
                            if os.path.abspath(source_path) != os.path.abspath(latest_file):
                                print(f"Ignoring older '{file}' in '{root}'.")
                                continue
                        os.makedirs(os.path.dirname(destination_path), exist_ok=True)
                        shutil.copy2(source_path, destination_path)
                        print(f"Copied: {file} to {destination_path}")

    print("All desired folders and files from all sources are copied")

    # Log the completion time
    current_time = time.strftime("%H:%M:%S %Z, %Y-%m-%d", time.localtime())
    print(f"Copy completed at {current_time} (e.g., 20:09 +08, 2025-07-28)")

except PermissionError:
    print(f"Error: Permission denied while accessing a source or '{destination_folder}'. Ensure you have appropriate rights.")
except FileNotFoundError as e:
    print(f"Error: {str(e)}")
except OSError as e:
    print(f"Error: Failed to copy folder. {str(e)}")
except Exception as e:
    print(f"Unexpected error: {str(e)}")