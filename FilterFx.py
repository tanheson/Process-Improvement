import os
import shutil
import time

# Define source and destination paths
# source_folder = r"C:/Results/ARL/S681/A0"
source_folder = r"C:/Results/NVL/HX/A0"
destination_folder = r"U:/NVL/HX/A0/results_production"

try:
    # Check if source folder exists
    if not os.path.isdir(source_folder):
        raise FileNotFoundError(f"Source folder '{source_folder}' does not exist.")

    # Check if destination parent directory exists
    destination_parent = os.path.dirname(destination_folder)
    if not os.path.exists(destination_parent):
        raise OSError(f"Parent directory '{destination_parent}' does not exist. Please create it first.")

    # # Delete existing destination folder if it exists
    # if os.path.exists(destination_folder):
    #     shutil.rmtree(destination_folder)
    #     print(f"Deleted existing folder: '{destination_folder}'")

    # Create destination folder
    os.makedirs(destination_folder, exist_ok=True)

    # Function to check if a folder contains a file with "HotVmin" in its name
    def has_hotvmin_file(folder_path):
        for root, _, files in os.walk(folder_path):
            if any("HotVmin" in f for f in files):
                return True
        return False

    # Copy folders that contain "HotVmin" files
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
                if has_hotvmin_file(dir_full_path):
                    dirs_to_keep.append(dir_name)
                else:
                    print(f"Ignoring folder '{dir_full_path}' as it does not contain a 'HotVmin' file.")
            # Update dirs to only include folders to keep (modifies os.walk behavior)
            dirs[:] = dirs_to_keep

        # Copy files and create directories
        for file in files:
            source_path = os.path.join(root, file)
            destination_path = os.path.join(dest_path, file)
            os.makedirs(os.path.dirname(destination_path), exist_ok=True)
            shutil.copy2(source_path, destination_path)
            print(f"Copied: {file} to {destination_path}")

    print("All desired folders and files are copied")

    # Log the completion time
    current_time = time.strftime("%H:%M:%S %Z, %Y-%m-%d", time.localtime())
    print(f"Copy completed at {current_time} (e.g., 09:54 +08, 2025-07-22)")

except PermissionError:
    print(f"Error: Permission denied while accessing '{source_folder}' or '{destination_folder}'. Ensure you have appropriate file.")
except FileNotFoundError as e:
    print(f"Error: {str(e)}")
except OSError as e:
    print(f"Error: Failed to copy folder. {str(e)}")
except Exception as e:
    print(f"Unexpected error: {str(e)}")



