import os
import shutil
import time
import logging
from datetime import datetime

# Define multiple source folders and a single destination
source_folders = [
    r"C:/Results/NVL/HX/A0",               # Local path 
    # r"C:/Results/ARL/S681/A0",                # Local path 
    # r"//PG07TCMV0088/c$/Results/ARL/S681/A0"
]
destination_folder = r"U:/NVL/HX/A0/results_production"

# Required file keywords to check
required_file_keywords = ["HotVmin.xlsx", "HotGNG.xlsx"]

# Define custom log directory
custom_log_dir = r"U:/users/Hs/script/Process-Improvement/runResultFilter/debuglog"  # Change this to your desired log folder path
os.makedirs(custom_log_dir, exist_ok=True)
log_filename = os.path.join(custom_log_dir, f"copy_log_{time.strftime('%Y%m%d_%H%M%S')}.log")
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename),
        logging.StreamHandler()  # Output to console as well
    ]
)

def get_latest_hotvmin_file(folder_path, keyword):
    """Check if a folder contains a file with the required keyword and get the latest modified file."""
    hotvmin_files = []
    for root, _, files in os.walk(folder_path):
        hotvmin_files.extend([os.path.join(root, f) for f in files if keyword in f])
    if not hotvmin_files:
        return None, False  # No matching file found
    latest_file = max(hotvmin_files, key=os.path.getmtime, default=None)
    return latest_file, True  # Return latest file path and flag indicating presence

def parse_timestamp_folder_name(folder_name):
    """Parse timestamp from folder name (e.g., '2025.07.25_19.28.10') and return datetime object."""
    try:
        return datetime.strptime(folder_name, "%Y.%m.%d_%H.%M.%S")
    except ValueError:
        return None

def get_latest_timestamp_folder(folder_path):
    """Get the latest timestamp folder under a HotVmin or HotGNG folder."""
    timestamp_folders = []
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isdir(item_path):
            timestamp = parse_timestamp_folder_name(item)
            if timestamp:
                timestamp_folders.append((item_path, timestamp))
    if not timestamp_folders:
        return None, None
    latest_folder, latest_timestamp = max(timestamp_folders, key=lambda x: x[1])
    return latest_folder, latest_timestamp  # Return path and timestamp of latest folder

try:
    # Check if destination parent directory exists
    destination_parent = os.path.dirname(destination_folder)
    if not os.path.exists(destination_parent):
        raise OSError(f"Parent directory '{destination_parent}' does not exist. Please create it first.")

    # Create destination folder
    os.makedirs(destination_folder, exist_ok=True)

    # Dictionary to track the latest timestamp folder and its timestamp for each HotVmin/HotGNG folder
    latest_timestamp_info = {}  # {folder_path: (latest_folder, latest_timestamp)}

    # Process each source folder to find the latest timestamp
    for source_folder in source_folders:
        if not os.path.isdir(source_folder):
            logging.warning(f"Source folder '{source_folder}' does not exist. Skipping...")
            continue

        logging.info(f"Processing source folder: {source_folder}")

        for root, dirs, _ in os.walk(source_folder):
            relative_path = os.path.relpath(root, source_folder)
            if "D3" in relative_path or "D4" in relative_path:
                for dir_name in dirs:
                    dir_full_path = os.path.join(root, dir_name)
                    if dir_name == "99999999_999_+99_+99":
                        logging.info(f"Ignoring folder '{dir_full_path}' as it matches the excluded name '99999999_999_+99_+99'.")
                        continue
                    if "HotVmin" in dir_name or "HotGNG" in dir_name:
                        latest_timestamp_folder, latest_timestamp = get_latest_timestamp_folder(dir_full_path)
                        if latest_timestamp_folder:
                            # Check for the presence of required files
                            has_file = False
                            for keyword in required_file_keywords:
                                latest_file, file_exists = get_latest_hotvmin_file(latest_timestamp_folder, keyword)
                                if file_exists:
                                    has_file = True
                                    break
                            if not has_file:
                                logging.info(f"Ignoring folder '{dir_full_path}' as the latest timestamp folder lacks a '{keyword}' file.")
                                continue
                            # Update latest timestamp info if this is the latest
                            dest_relative_path = os.path.join(os.path.relpath(root, source_folder), dir_name)
                            dest_folder_path = os.path.join(destination_folder, dest_relative_path)
                            if dest_folder_path not in latest_timestamp_info or latest_timestamp > latest_timestamp_info[dest_folder_path][1]:
                                latest_timestamp_info[dest_folder_path] = (latest_timestamp_folder, latest_timestamp)
                                logging.info(f"Updated latest timestamp folder for '{dest_folder_path}' to '{os.path.basename(latest_timestamp_folder)}' (modified: {time.ctime(os.path.getmtime(latest_timestamp_folder))})")

    # Copy and replace in destination
    for source_folder in source_folders:
        if not os.path.isdir(source_folder):
            continue

        for root, dirs, files in os.walk(source_folder):
            relative_path = os.path.relpath(root, source_folder)
            if "D3" in relative_path or "D4" in relative_path:
                dirs_to_keep = []
                for dir_name in dirs:
                    dir_full_path = os.path.join(root, dir_name)
                    if dir_name == "99999999_999_+99_+99":
                        logging.info(f"Ignoring folder '{dir_full_path}' as it matches the excluded name '99999999_999_+99_+99'.")
                        continue
                    if "HotVmin" in dir_name or "HotGNG" in dir_name:
                        latest_timestamp_folder, latest_timestamp = get_latest_timestamp_folder(dir_full_path)
                        if latest_timestamp_folder:
                            dest_relative_path = os.path.join(os.path.relpath(root, source_folder), dir_name)
                            dest_folder_path = os.path.join(destination_folder, dest_relative_path)
                            if dest_folder_path in latest_timestamp_info:
                                latest_source_folder, latest_timestamp = latest_timestamp_info[dest_folder_path]
                                # Remove all existing timestamp subfolders in destination
                                if os.path.exists(dest_folder_path):
                                    for item in os.listdir(dest_folder_path):
                                        item_path = os.path.join(dest_folder_path, item)
                                        if os.path.isdir(item_path):
                                            item_timestamp = parse_timestamp_folder_name(item)
                                            if item_timestamp:  # Remove only timestamp folders
                                                shutil.rmtree(item_path)
                                                logging.info(f"Removed timestamp folder '{item_path}'.")
                                # Copy the entire latest timestamp folder
                                if latest_source_folder == latest_timestamp_folder:
                                    dest_timestamp_folder = os.path.join(dest_folder_path, os.path.basename(latest_timestamp_folder))
                                    if os.path.exists(dest_timestamp_folder):
                                        shutil.rmtree(dest_timestamp_folder)
                                        logging.info(f"Removed existing '{os.path.basename(dest_timestamp_folder)}' to replace with latest.")
                                    shutil.copytree(latest_timestamp_folder, dest_timestamp_folder, dirs_exist_ok=True)
                                    logging.info(f"Copied timestamp folder '{os.path.basename(latest_timestamp_folder)}' to '{dest_timestamp_folder}'")
                            dirs_to_keep.append(dir_name)
                    else:
                        dirs_to_keep.append(dir_name)
                dirs[:] = dirs_to_keep

    logging.info("All desired folders and files from all sources are copied")

    # Log the completion time
    current_time = time.strftime("%H:%M:%S %Z, %Y-%m-%d", time.localtime())
    logging.info(f"Copy completed at {current_time} (e.g., 13:25 +08, 2025-09-11)")

except PermissionError:
    logging.error(f"Error: Permission denied while accessing a source or '{destination_folder}'. Ensure you have appropriate rights.")
except FileNotFoundError as e:
    logging.error(f"Error: {str(e)}")
except OSError as e:
    logging.error(f"Error: Failed to copy folder. {str(e)}")
except Exception as e:
    logging.error(f"Unexpected error: {str(e)}")