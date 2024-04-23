import sys
import subprocess
import wmi
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

# Function to download and extract DBAN
def download_dban():
    dban_url = "https://sourceforge.net/projects/dban/files/latest/download"
    dban_zip = "dban-2.3.0.zip"
    dban_dir = "dban-2.3.0"

    try:
        print("Downloading DBAN utility...")
        temp_file, _ = urlretrieve(dban_url, dban_zip)

        print("Extracting DBAN utility...")
        with zipfile.ZipFile(dban_zip, 'r') as zip_ref:
            zip_ref.extractall()

        dban_path = os.path.join(os.getcwd(), dban_dir, "dban.iso")
        return dban_path
    except Exception as e:
        print(f"Error downloading DBAN utility: {e}")
        return None
    finally:
        if os.path.exists(temp_file):
            os.remove(temp_file)

# Install the wmi module
install_wmi()

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
    # Download and extract DBAN
    dban_path = download_dban()
    if not dban_path:
        print("Unable to download DBAN utility. Exiting...")
        input("Press Enter to exit...")
        sys.exit(1)

    print(f"DBAN utility downloaded and extracted: {dban_path}")

    # Check if the drive is a USB drive or removable media
    if "USB" in drive_to_erase.Caption or "Removable" in drive_to_erase.Caption:
        print("USB drive or removable media detected.")
        print("Please follow these steps to erase the drive using DBAN:")
        print("1. Create a bootable USB drive or CD/DVD using the DBAN ISO file.")
        print("2. Boot from the DBAN media and follow the on-screen instructions to erase the drive.")
    else:
        print("Internal drive detected.")
        print("Please follow these steps to erase the drive using DBAN:")
        print("1. Create a bootable USB drive or CD/DVD using the DBAN ISO file.")
        print("2. Boot from the DBAN media and follow the on-screen instructions to erase the drive.")

    input("Press Enter to exit...")
else:
    print("Disk erase canceled.")

input("Press Enter to exit...")