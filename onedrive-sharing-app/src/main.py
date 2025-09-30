import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from dotenv import load_dotenv
from onedrive import OneDriveAPI

# Load environment variables from .env file
load_dotenv()

class OneDriveSharingApp:
    def __init__(self, master):
        self.master = master
        master.title("OneDrive Sharing Link Generator")

        self.label = tk.Label(master, text="Select a file or folder from OneDrive:")
        self.label.pack()

        self.select_button = tk.Button(master, text="Select", command=self.select_file_or_folder)
        self.select_button.pack()

        self.link_label = tk.Label(master, text="")
        self.link_label.pack()

        self.generate_button = tk.Button(master, text="Generate Sharing Link", command=self.generate_link)
        self.generate_button.pack()

    def select_file_or_folder(self):
        self.selected_path = filedialog.askopenfilename() or filedialog.askdirectory()
        if self.selected_path:
            self.link_label.config(text=f"Selected: {self.selected_path}")

    def generate_link(self):
        api = OneDriveAPI()
        if hasattr(self, 'selected_path'):
            try:
                # Assuming selected_path is the item ID in OneDrive
                link = api.generate_sharing_link(self.selected_path)
                messagebox.showinfo("Sharing Link", f"Generated Link: {link}")
            except Exception as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showwarning("Warning", "Please select a file or folder first.")

if __name__ == "__main__":
    root = tk.Tk()
    app = OneDriveSharingApp(root)
    root.mainloop()