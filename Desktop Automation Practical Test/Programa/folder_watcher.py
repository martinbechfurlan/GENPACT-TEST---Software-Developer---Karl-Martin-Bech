import os
import threading
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Define the FolderWatcher class, inheriting from FileSystemEventHandler
class FolderWatcher(FileSystemEventHandler):
    def __init__(self, folder_to_watch, excel_consolidator):
        # Initialize the class with the folder to watch and an instance of ExcelConsolidator
        self.folder_to_watch = folder_to_watch
        self.excel_consolidator = excel_consolidator
        # Create an Observer that will monitor the folder
        self.observer = Observer()

    def run(self):
        # Set up the Observer to use this instance as the event handler
        self.observer.schedule(self, self.folder_to_watch, recursive=True)
        # Start the Observer
        self.observer.start()
        try:
            # Keep the script running to continue monitoring the folder
            while True:
                time.sleep(5)
        except KeyboardInterrupt:
            # Stop the Observer if interrupted (e.g., with Ctrl+C)
            self.observer.stop()
        # Ensure the Observer shuts down properly
        self.observer.join()

    def on_created(self, event):
        # Method called when a new file or folder is created
        if event.is_directory:
            # If the event is a directory, do nothing
            return None
        elif event.src_path.endswith('.xls') or event.src_path.endswith('.xlsx'):
            # If the created file is an Excel file, process it if it's not in 'processed' or 'not_applicable'
            if 'processed' not in event.src_path and 'not_applicable' not in event.src_path:
                self.excel_consolidator.consolidate(event.src_path)
        else:
            # If the created file is not an Excel file and is not in 'processed' or 'not_applicable', move it to 'not_applicable'
            if 'processed' not in event.src_path and 'not_applicable' not in event.src_path:
                self.excel_consolidator.move_to_not_applicable(event.src_path)

    def start(self):
        # Start a thread to run the run() method in the background
        thread = threading.Thread(target=self.run)
        thread.daemon = True  # The thread will automatically stop when the program exits
        thread.start()
