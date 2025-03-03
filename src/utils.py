import os
import glob

def cleanup_data_directory():
    """Delete all JSON files in the data directory except for essential files."""
    data_dir = "data"
    essential_files = []  # Add any essential files that should not be deleted

    # Get a list of all JSON files in the data directory
    json_files = glob.glob(os.path.join(data_dir, "*.json"))

    # Delete each JSON file that is not essential
    for json_file in json_files:
        if os.path.basename(json_file) not in essential_files:
            try:
                os.remove(json_file)
                print(f"Deleted: {json_file}")
            except Exception as e:
                print(f"Error deleting {json_file}: {e}")