# UNIVET Dataset: Bulk-Renaming & EXIF Extraction Script
## Description
This folder contains a Python automation script developed to rename a large collection of photos according to a specific format defined in an Excel file.

The core logic relies on matching the chronological acquisition order of the images with the sequential rows defined in the Excel protocol, ensuring zero data-entry errors during post-production.

## Features
- **Chronological Sorting**: the script sorts photos based on their EXIF "Acquisition Time" metadata, ensuring the correct order of renaming.
- **Safe Renaming**: original photos remain unchanged, and renamed copies are created in a new location.
- **Automated Excel Update**: the script updates the Excel file with EXIF "Focal Length" data for each renamed photo.


## How to Use
1. Place the photos you want to rename in the appropriate folders (e.g., `photos/samsung` for Samsung photos and `photos/iphone` for iPhone photos).
2. Update the `excel_path` variable in the `rename.py` script to point to your Excel file that contains the renaming format and other relevant information.
3. Run the script. The safely renamed photos will appear in the "renamed" folder, and the Excel database will be automatically updated with the required metadata.

## Environment Setup & Dependencies
The script utilizes Python's standard libraries (`os`, `shutil`, `time`) for system operations, but requires external packages for Excel manipulation and EXIF extraction.

Create a virtual environment and install the required dependencies:

```bash
# Create the virtual environment
python -m venv .venv

# Activate the virtual environment
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate

# Install external dependencies
pip install -r requirements.txt
```