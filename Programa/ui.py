import os
import tkinter as tk
from tkinter import filedialog, scrolledtext
from folder_watcher import FolderWatcher
from excel_consolidator import ExcelConsolidator
from PIL import Image, ImageTk

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Consolidation Tool")  # Set the window title
        self.root.geometry("700x500")  # Set the window size
        self.root.configure(bg='#e0f7fa')  # Set the background color

        # Load and resize the logo
        self.logo_image = Image.open('genpact.png')  # Ensure the file is in the same directory as the script
        self.logo_image = self.logo_image.resize((100, 100))  # Resize the logo
        self.logo_photo = ImageTk.PhotoImage(self.logo_image)

        # Place the logo in the top-right corner
        self.logo_label = tk.Label(root, image=self.logo_photo, bg='#e0f7fa')
        self.logo_label.place(x=580, y=100)  # Adjust the position as needed

        # Header section
        self.header_frame = tk.Frame(root, bg='#00acc1')
        self.header_frame.pack(fill=tk.X, padx=10, pady=10)

        self.header_label = tk.Label(
            self.header_frame, 
            text="Excel Consolidation Tool", 
            bg='#00acc1', 
            fg='white', 
            font=("Arial", 16, "bold")
        )
        self.header_label.pack(pady=10)

        # Folder selection section
        self.select_folder_frame = tk.Frame(root, bg='#e0f7fa')
        self.select_folder_frame.pack(pady=10)

        self.label = tk.Label(
            self.select_folder_frame, 
            text="Select folder to watch:", 
            bg='#e0f7fa', 
            font=("Arial", 14)
        )
        self.label.pack(pady=5)

        self.select_folder_button = tk.Button(
            self.select_folder_frame, 
            text="Select Folder", 
            command=self.select_folder, 
            bg='#004d40', 
            fg='white', 
            font=("Arial", 12)
        )
        self.select_folder_button.pack(pady=5)

        # Exit button
        self.exit_button = tk.Button(
            root, 
            text="Exit", 
            command=self.exit_application, 
            bg='#d32f2f', 
            fg='white', 
            font=("Arial", 12, "bold")
        )
        self.exit_button.pack(pady=10)

        # Log area
        self.log_frame = tk.Frame(root, bg='#e0f7fa')
        self.log_frame.pack(pady=10)

        self.log_text = scrolledtext.ScrolledText(
            self.log_frame, 
            width=85, 
            height=20, 
            font=("Arial", 10),
            bg='#ffffff', 
            fg='#000000',
            padx=10,
            pady=10
        )
        self.log_text.pack()

    def log(self, message):
        # Append messages to the log area and ensure the latest message is visible
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)

    def select_folder(self):
        # Open a folder selection dialog
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            master_file = os.path.join(folder_selected, "master.xlsx")
            # Create an instance of ExcelConsolidator
            consolidator = ExcelConsolidator(master_file, self.log)
            # Create an instance of FolderWatcher
            watcher = FolderWatcher(folder_selected, consolidator)
            self.log(f"Watching folder: {folder_selected}")
            # Start watching the folder in a new thread
            watcher.start()

            # Process existing files in the selected folder
            self.process_existing_files(folder_selected, consolidator, master_file)

    def process_existing_files(self, folder, consolidator, master_file):
        # Iterate through files in the selected folder
        for file_name in os.listdir(folder):
            file_path = os.path.join(folder, file_name)
            if os.path.isfile(file_path):
                if file_name.lower() == 'master.xlsx':
                    continue  # Skip the master file itself

                if file_name.lower().endswith('.xls') or file_name.lower().endswith('.xlsx'):
                    # Consolidate Excel files
                    consolidator.consolidate(file_path)
                else:
                    # Move non-Excel files to 'not_applicable' folder
                    consolidator.move_to_not_applicable(file_path)

    def exit_application(self):
        # Close the application
        self.root.destroy()
