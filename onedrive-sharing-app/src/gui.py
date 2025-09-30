from tkinter import Tk, Label, Button, filedialog, messagebox
import os
from dotenv import load_dotenv
from onedrive import OneDriveAPI  # Assuming this function is defined in onedrive.py

class OneDriveSharingApp:
    def __init__(self, master):
        self.master = master
        master.title("OneDrive Sharing Link Generator")

        self.label = Label(master, text="Select a file or folder from OneDrive:")
        self.label.pack(pady=10)

        self.select_button = Button(master, text="Select File/Folder", command=self.select_file_or_folder)
        self.select_button.pack(pady=5)

        self.generate_button = Button(master, text="Generate Sharing Link", command=self.generate_link)
        self.generate_button.pack(pady=5)

        self.selected_path = None

    def select_file_or_folder(self):
        self.selected_path = filedialog.askopenfilename() or filedialog.askdirectory()
        if self.selected_path:
            messagebox.showinfo("Selected", f"You selected: {self.selected_path}")

    def generate_link(self):
        api = OneDriveAPI()
        if not self.selected_path:
            messagebox.showwarning("Warning", "Please select a file or folder first.")
            return
        
        try:
            link = api.generate_sharing_link(self.selected_path)
            messagebox.showinfo("Sharing Link", f"Generated Sharing Link:\n{link}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    load_dotenv()  # Load environment variables from .env
    root = Tk()
    app = OneDriveSharingApp(root)
    root.mainloop()