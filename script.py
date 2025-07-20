import os
import shutil
import threading
import time
from datetime import datetime
from pathlib import Path

backup_tasks = []

def force_copy(source_path, dest_path):
    """Force copy a file using binary read/write."""
    with open(source_path, 'rb') as src, open(dest_path, 'wb') as dst:
        dst.write(src.read())

def backup_file(source_path, dest_dir, interval_minutes):
    source = Path(source_path)
    dest = Path(dest_dir)
    task_name = source.name

    while True:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"{source.stem}_{timestamp}{source.suffix}"
        backup_path = dest / backup_filename

        try:
            force_copy(source, backup_path)
            print(f"[{task_name}] Forced backup successful at {timestamp}")
        except Exception as e:
            print(f"[{task_name}] Forced backup failed: {e}")

        time.sleep(interval_minutes * 60)

def start_backup():
    source_path = input("Enter the full path of the Excel file to back up: ").strip()
    dest_dir = input("Enter the destination folder for backups: ").strip()
    interval = int(input("Enter backup interval in minutes: ").strip())

    if not os.path.exists(source_path):
        print("❌ Source file does not exist.")
        return
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)

    thread = threading.Thread(target=backup_file, args=(source_path, dest_dir, interval), daemon=True)
    thread.start()

    backup_tasks.append({
        "file": source_path,
        "dest": dest_dir,
        "interval": interval,
        "thread": thread
    })

    print(f"✅ Backup started for '{source_path}' every {interval} minutes.")

def show_status():
    print("\n📋 Current Backup Tasks:")
    for i, task in enumerate(backup_tasks, 1):
        print(f"{i}. File: {task['file']}")
        print(f"   Destination: {task['dest']}")
        print(f"   Interval: {task['interval']} minutes\n")

def main():
    print("=== Excel Auto Backup Manager ===")
    while True:
        print("\nOptions:")
        print("1. Start new backup task")
        print("2. Show current backup status")
        print("3. Exit")

        choice = input("Select an option (1/2/3): ").strip()
        if choice == "1":
            start_backup()
        elif choice == "2":
            show_status()
        elif choice == "3":
            print("👋 Exiting backup manager.")
            break
        else:
            print("❌ Invalid choice. Try again.")

if __name__ == "__main__":
    main()
