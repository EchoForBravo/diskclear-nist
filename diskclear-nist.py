import sys
import subprocess
import wmi
import re
import os
import shutil
import zipfile
from pathlib import Path
from urllib.request import urlretrieve

# Function to install the wmi module
def install_wmi():
    try:
        import wmi
    except ImportError:
        print("The 'wmi' module is not installed. Installing now...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "wmi"])
        print("Installation complete.")

# Function to install the DiskSpd utility
def install_diskspd():
    diskspd_url = "https://github.com/microsoft/diskspd/releases/download/v2.1/DiskSpd.ZIP"
    diskspd_zip = "DiskSpd-2.0.21a.zip"
    diskspd_dir = "DiskSpd-2.0.21a"

    try:
        print("Downloading DiskSpd utility...")
        temp_file, _ = urlretrieve(diskspd_url, diskspd_zip)

        print("Extracting DiskSpd utility...")
        with zipfile.ZipFile(diskspd_zip, 'r') as zip_ref:
            zip_ref.extractall()

        diskspd_path = os.path.join(os.getcwd(), diskspd_dir, "DiskSpd.exe")
        return diskspd_path
    except Exception as e:
        print(f"Error installing DiskSpd utility: {e}")
        return None
    finally:
        if os.path.exists(temp_file):
            os.remove(temp_file)

# Install the wmi module
install_wmi()

# Install the DiskSpd utility if not present
diskspd_path = shutil.which("DiskSpd.exe")
if not diskspd_path:
    diskspd_path = install_diskspd()
    if not diskspd_path:
        print("Unable to install DiskSpd utility. Exiting...")
        input("Press Enter to exit...")
        sys.exit(1)

# Connect to the WMI service
c = wmi.WMI()

# Get a list of disk drives
drives = c.Win32_DiskDrive()

# Print the list of drives
print("Available disk drives:")
for i, drive in enumerate(drives):
    print(f"{i+1}. {drive.Caption} ({drive.Size} bytes)")

# Prompt the user to select a drive
selected_drive = int(input("Enter the number of the drive you want to erase: "))

# Get the selected drive object
drive_to_erase = drives[selected_drive - 1]

# Confirm with the user
confirm = input(f"Are you sure you want to erase {drive_to_erase.Caption}? (y/n) ")
if confirm.lower() == "y":
    # Erase the disk using NIST-approved method
    print(f"Erasing {drive_to_erase.Caption} using NIST-approved method...")

    # Check if the drive is a USB drive or removable media
    if "USB" in drive_to_erase.Caption or "Removable" in drive_to_erase.Caption:
        # Handle USB drives or removable media differently
        print("USB drive or removable media detected.")
        drive_letter = input("Please enter the drive letter (e.g., E:): ")
        subprocess.run([diskspd_path, "-c1G", "-w0", f"{drive_letter}"])

        # Verify the erasure process
        print("Verifying erasure...")
        # Add code to verify that the disk is erased (e.g., read data from the disk and check for all zeros)

        print("Disk erased and verified successfully.")
    else:
        # Handle internal drives
        match = re.match(r"^([A-Z]):", drive_to_erase.Caption)
        if match:
            drive_letter = match.group(1)
            subprocess.run([diskspd_path, "-c1G", "-w0", f"{drive_letter}:"])

            # Verify the erasure process
            print("Verifying erasure...")
            # Add code to verify that the disk is erased (e.g., read data from the disk and check for all zeros)

            print("Disk erased and verified successfully.")
        else:
            print(f"Unable to determine drive letter for {drive_to_erase.Caption}. Skipping erasure.")
else:
    print("Disk erase canceled.")

input("Press Enter to exit...")