import time
import os
import shutil
import win32print
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class FileHandler(FileSystemEventHandler):
    def process_file(self, filepath):
        if os.path.isfile(filepath):  # make sure the file still exists (it could have been deleted by another process)
            filename = os.path.basename(filepath)
            if filename.startswith('~$'):  # ignore temporary files created by Microsoft Office
                return
            ext = os.path.splitext(filename)[1].lower()
            if ext in ('.pdf', '.docx', '.xlsx', '.pptx', '.txt'):  # specify the file extensions to print
                print(f'Printing {filename}')
                printer_name_start = filename.find("PRT_")
                if printer_name_start == -1:
                    print("Error: Printer name not found in filename.")
                    new_filename = os.path.splitext(filename)[0] + '.failed'
                    os.rename(filepath, os.path.join(os.path.dirname(filepath), new_filename))
                    return
                printer_name = filename[printer_name_start + 4:-4].strip()
                try:
                    hPrinter = win32print.OpenPrinter(printer_name)
                except:
                    print(f"Error: Printer {printer_name} not found.")
                    new_filename = os.path.splitext(filename)[0] + '.failed'
                    os.rename(filepath, os.path.join(os.path.dirname(filepath), new_filename))
                    return
                time.sleep(1)
                with open(filepath, "rb") as file:
                    data = file.read()
                try:
                    hJob = win32print.StartDocPrinter(hPrinter, 1, (filename, None, "TEXT"))
                    try:
                        win32print.StartPagePrinter(hPrinter)
                        win32print.WritePrinter(hPrinter, data)
                        win32print.EndPagePrinter(hPrinter)
                    finally:
                        win32print.EndDocPrinter(hPrinter)
                finally:
                    win32print.ClosePrinter(hPrinter)
                dest_folder = "C:/AUTOPRINT/PROCESSADOS"  # replace with the destination folder for processed files
                if not os.path.exists(dest_folder):
                    os.makedirs(dest_folder)
                shutil.move(filepath, os.path.join(dest_folder, filename))

    def on_created(self, event):
        filepath = event.src_path
        self.process_file(filepath)

observer = Observer()
event_handler = FileHandler()
folder_path = 'C:/AUTOPRINT/'  # replace with the path to the folder you want to monitor

# Process existing files
for filename in os.listdir(folder_path):
    filepath = os.path.join(folder_path, filename)
    event_handler.process_file(filepath)

observer.schedule(event_handler, folder_path, recursive=False)
observer.start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()
observer.join()