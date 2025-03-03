import asyncio
import time
import schedule
from src.scraper import scrape_tippmix
from src.odds_extractor import extract_odds
from src.historical_data import main as historical_data_main
from src.telegram_bot import compare_odds_with_stats
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

def ensure_data_directory_exists():
    """Create the data directory if it doesn't exist."""
    data_dir = "data"
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
        print(f"Created directory: {data_dir}")

async def run_workflow():
    try:
        print("Starting workflow...")
        await scrape_tippmix()
        extract_odds()
        historical_data_main()
        compare_odds_with_stats("data/games_with_odds.json")
        print("Workflow completed.")
    except Exception as e:
        print(f"Error in workflow: {e}")
    finally:
        # Clean up JSON files after the workflow
        cleanup_data_directory()
        print("Cleanup completed. Waiting for the next run...")

def schedule_workflow():
    # Ensure the data directory exists
    ensure_data_directory_exists()

    # Schedule the workflow to run every 2.5 minutes
    schedule.every(2.5).minutes.do(lambda: asyncio.run(run_workflow()))

    # Keep the script running
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    schedule_workflow()
