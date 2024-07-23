import os
import tkinter as tk
from ui import App  # Import the App class from the ui module

def main():
    root = tk.Tk()  # Create the main window of the Tkinter application
    app = App(root)  # Initialize the App class with the main window
    root.mainloop()  # Enter the Tkinter event loop

if __name__ == "__main__":
    main()  # Call the main function if this script is executed directly
