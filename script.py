import os
import shutil
import threading
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class BackupApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Auto Backup Manager")

        self.source_path = tk.StringVar()
        self.dest_path = tk.StringVar()
        self.interval = tk.IntVar(value=5)
        self.method = tk.StringVar(value="Timestamped")
        self.countdown = tk.StringVar(value="")

        self.setup_ui()
        self.backup_thread = None
        self.stop_event = threading.Event()

    def setup_ui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frame, text="Source Excel File:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.source_path, width=50).grid(row=0, column=1)
        ttk.Button(frame, text="Browse", command=self.browse_source).grid(row=0, column=2)

        ttk.Label(frame, text="Destination Folder:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.dest_path, width=50).grid(row=1, column=1)
        ttk.Button(frame, text="Browse", command=self.browse_dest).grid(row=1, column=2)

        ttk.Label(frame, text="Backup Interval (minutes):").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.interval, width=10).grid(row=2, column=1, sticky="w")

        ttk.Label(frame, text="Backup Method:").grid(row=3, column=0, sticky="w")
        ttk.Combobox(frame, textvariable=self.method, values=["Basic", "Timestamped"], state="readonly").grid(row=3, column=1, sticky="w")

        ttk.Button(frame, text="Start Backup", command=self.start_backup).grid(row=4, column=1, pady=10)

        ttk.Label(frame, text="Next Backup In:").grid(row=5, column=0, sticky="w")
        ttk.Label(frame, textvariable=self.countdown).grid(row=5, column=1, sticky="w")

        self.log = tk.Text(frame, height=10, width=70, state="disabled")
        self.log.grid(row=6, column=0, columnspan=3, pady=10)

    def browse_source(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.source_path.set(file_path)

    def browse_dest(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.dest_path.set(folder_path)

    def log_message(self, message):
        self.log.config(state="normal")
        self.log.insert("end", f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def force_copy(self, source, destination):
        with open(source, 'rb') as src, open(destination, 'wb') as dst:
            dst.write(src.read())

    def backup_loop(self, source, dest, interval, method):
        while not self.stop_event.is_set():
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            source_path = Path(source)
            if method == "Basic":
                backup_filename = source_path.name
            else:
                backup_filename = f"{source_path.stem}_{timestamp}{source_path.suffix}"
            backup_path = Path(dest) / backup_filename

            try:
                self.force_copy(source, backup_path)
                self.log_message(f"Backup successful: {backup_filename}")
            except Exception as e:
                self.log_message(f"Backup failed: {e}")

            for remaining in range(interval * 60, 0, -1):
                if self.stop_event.is_set():
                    break
                mins, secs = divmod(remaining, 60)
                self.countdown.set(f"{mins:02}:{secs:02}")
                time.sleep(1)

    def start_backup(self):
        source = self.source_path.get()
        dest = self.dest_path.get()
        interval = self.interval.get()
        method = self.method.get()

        if not os.path.exists(source):
            messagebox.showerror("Error", "Source file does not exist.")
            return
        if not os.path.exists(dest):
            os.makedirs(dest)

        self.stop_event.clear()
        self.backup_thread = threading.Thread(target=self.backup_loop, args=(source, dest, interval, method), daemon=True)
        self.backup_thread.start()
        self.log_message(f"Started backup for '{Path(source).name}' every {interval} minutes using '{method}' method.")

if __name__ == "__main__":
    root = tk.Tk()
    app = BackupApp(root)
    root.mainloop()
