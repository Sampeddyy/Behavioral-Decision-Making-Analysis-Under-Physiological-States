import os
import shutil

# Main folder path containing all 80 Excel files
main_folder = r"S:\lmassignment\LM_A2_2025_data"  # Using raw string to handle backslashes

# Ensure the main folder exists
if not os.path.exists(main_folder):
    print(f"Error: Main folder '{main_folder}' does not exist. Please check the path.")
    exit()

# Get list of all .xlsx files in the main folder
excel_files = [f for f in os.listdir(main_folder) if f.endswith(".xlsx")]

# Check if we have the expected number of files
print(f"Found {len(excel_files)} .xlsx files in {main_folder}")
if len(excel_files) != 80:
    print("Warning: Expected 80 files. Please verify all files are present.")

# Function to extract participant ID from file name
def get_participant_id(filename):
    # Assuming the ID is the part before the first hyphen or category indicator
    # e.g., "AD-food-C.xlsx" -> "AD"
    return filename.split('-')[0].strip()

# Create folders for each participant and move files
participant_ids = set()  # To track unique participant IDs
for file in excel_files:
    participant_id = get_participant_id(file)
    participant_ids.add(participant_id)
    
    # Create participant folder if it doesn't exist
    participant_folder = os.path.join(main_folder, participant_id)
    if not os.path.exists(participant_folder):
        os.makedirs(participant_folder)
        print(f"Created folder: {participant_folder}")
    
    # Define source and destination paths
    source_path = os.path.join(main_folder, file)
    dest_path = os.path.join(participant_folder, file)
    
    # Move the file to the participant's folder
    try:
        shutil.move(source_path, dest_path)
        print(f"Moved {file} to {participant_folder}")
    except Exception as e:
        print(f"Error moving {file}: {e}")

# Summary
print(f"\nSorting complete!")
print(f"Processed {len(excel_files)} files into {len(participant_ids)} participant folders.")
print(f"Participant IDs found: {sorted(participant_ids)}")