import os
import shutil
import time
from pathlib import Path

# Define source and destination paths
source_folder = r"C:/Results/NVL/HX/A0"
destination_folder = r"U:/NVL/HX/A0/results_production"

# Required file keyword to check
required_file_keyword = "HotVmin.xlsx"

try:
    # Check if source folder exists
    if not os.path.isdir(source_folder):
        raise FileNotFoundError(f"Source folder '{source_folder}' does not exist.")

    # Check if destination parent directory exists
    destination_parent = os.path.dirname(destination_folder)
    if not os.path.exists(destination_parent):
        raise OSError(f"Parent directory '{destination_parent}' does not exist. Please create it first.")

    # Create destination folder
    os.makedirs(destination_folder, exist_ok=True)

    # Function to check if a folder contains a file with the required keyword and get the latest modified file
    def get_latest_hotvmin_file(folder_path):
        hotvmin_files = []
        for root, _, files in os.walk(folder_path):
            hotvmin_files.extend([os.path.join(root, f) for f in files if required_file_keyword in f])
        if not hotvmin_files:
            return None, False  # No "HotVmin.xlsx" file found
        latest_file = max(hotvmin_files, key=os.path.getmtime, default=None)
        return latest_file, True  # Return latest file path and flag indicating presence

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
                latest_file, has_hotvmin = get_latest_hotvmin_file(dir_full_path)
                if "HotVmin" in dir_name and not has_hotvmin:
                    print(f"Ignoring folder '{dir_full_path}' as it lacks a '{required_file_keyword}' file.")
                elif "HotVmin" in dir_name and has_hotvmin:
                    print(f"Keeping folder '{dir_full_path}' with latest '{os.path.basename(latest_file)}' (modified: {time.ctime(os.path.getmtime(latest_file))})")
                    dirs_to_keep.append(dir_name)
                else:
                    dirs_to_keep.append(dir_name)  # Keep non-"HotVmin" folders
            # Update dirs to only include folders to keep (modifies os.walk behavior)
            dirs[:] = dirs_to_keep

        # Copy files and create directories
        for file in files:
            source_path = os.path.join(root, file)
            destination_path = os.path.join(dest_path, file)
            # Only copy the latest "HotVmin.xlsx" if it matches, ignore others
            if "HotVmin" in os.path.basename(root) and required_file_keyword in file:
                latest_file, _ = get_latest_hotvmin_file(root)
                if os.path.abspath(source_path) != os.path.abspath(latest_file):
                    print(f"Ignoring older '{file}' in '{root}'.")
                    continue  # Skip older "HotVmin.xlsx" files
            os.makedirs(os.path.dirname(destination_path), exist_ok=True)
            shutil.copy2(source_path, destination_path)
            print(f"Copied: {file} to {destination_path}")

    print("All desired folders and files are copied")

    # Log the completion time
    current_time = time.strftime("%H:%M:%S %Z, %Y-%m-%d", time.localtime())
    print(f"Copy completed at {current_time} (e.g., 10:20 +08, 2025-07-28)")

except PermissionError:
    print(f"Error: Permission denied while accessing '{source_folder}' or '{destination_folder}'. Ensure you have appropriate rights.")
except FileNotFoundError as e:
    print(f"Error: {str(e)}")
except OSError as e:
    print(f"Error: Failed to copy folder. {str(e)}")
except Exception as e:
    print(f"Unexpected error: {str(e)}")