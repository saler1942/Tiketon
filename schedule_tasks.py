import time
import subprocess
import os
from datetime import datetime, timedelta

def run_command(command):
    """Run a management command"""
    print(f"Running command: {command}")
    result = subprocess.run(command, shell=True, capture_output=True, text=True)
    print(f"Output: {result.stdout}")
    if result.stderr:
        print(f"Error: {result.stderr}")
    return result

# Calculate time until next run (2:00 AM)
def wait_until_next_run():
    now = datetime.now()
    # If we're past 2 AM, wait until tomorrow at 2 AM
    if now.hour >= 2:
        next_run = now.replace(day=now.day+1, hour=2, minute=0, second=0, microsecond=0)
    else:
        next_run = now.replace(hour=2, minute=0, second=0, microsecond=0)
    
    # Special handling for July 13, 2024 start date
    start_date = datetime(2024, 7, 13, 2, 0, 0)
    if now < start_date:
        next_run = start_date
        print(f"Waiting until initial start date: {next_run}")
    
    # Calculate seconds to wait
    wait_seconds = (next_run - now).total_seconds()
    if wait_seconds < 0:
        wait_seconds = 0
    
    return wait_seconds

def main():
    print(f"Scheduler started at: {datetime.now()}")
    print(f"Automatic event purge will begin on July 13, 2024 at 2:00 AM")
    
    while True:
        # Calculate time to wait until next run
        wait_seconds = wait_until_next_run()
        
        # Sleep until next run time
        print(f"Next run scheduled at: {datetime.now() + timedelta(seconds=wait_seconds)}")
        print(f"Sleeping for {wait_seconds/3600:.2f} hours until next run")
        time.sleep(wait_seconds)
        
        # Run the purge_events command
        print(f"Running scheduled task at: {datetime.now()}")
        run_command("python manage.py purge_events")
        
        # Sleep for a minute to avoid immediate re-run
        time.sleep(60)

if __name__ == "__main__":
    main() 