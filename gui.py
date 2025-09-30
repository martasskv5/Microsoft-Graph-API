import os
import pyperclip
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from ms_graph import get_access_token
from onedrive import OneDrive, get_root_path
from dotenv import load_dotenv
from redirect import Redirect, AVAIBLE_SERVICES
from win11toast import toast as notification
from datetime import datetime, timedelta

# Load environment variables from .env file
load_dotenv()

class OneDriveSharingApp:
    def __init__(self, master):
        self.access_token = get_access_token()
        self.master = master
        master.title("OneDrive Sharing Link Generator")
        self.client = None
        self.onedrive = OneDrive(self.access_token)
        self.setup_gui(self.master)

    def setup_gui(self, master):
        self.label = tk.Label(master, text="Select a file from OneDrive:")
        self.label.pack()

        self.select_button = tk.Button(master, text="Select File", command=self.select_file)
        self.select_button.pack()

        self.file_label = tk.Label(master, text="")
        self.file_label.pack()

        self.sharing_type_label = tk.Label(master, text="Select sharing type:")
        self.sharing_type_label.pack()

        self.sharing_type = tk.StringVar(value="view")
        self.view_radio = tk.Radiobutton(master, text="View", variable=self.sharing_type, value="view")
        self.view_radio.pack()
        self.edit_radio = tk.Radiobutton(master, text="Edit", variable=self.sharing_type, value="edit")
        self.edit_radio.pack()

        self.scope_label = tk.Label(master, text="Select sharing scope:")
        self.scope_label.pack()

        self.scope = tk.StringVar(value="anonymous")
        self.anonymous_radio = tk.Radiobutton(master, text="Anonymous", variable=self.scope, value="anonymous")
        self.anonymous_radio.pack()
        self.organization_radio = tk.Radiobutton(master, text="Organization", variable=self.scope, value="organization")
        self.organization_radio.pack()

        self.date_label = tk.Label(master, text="Select expiration date:")
        self.date_label.pack()
        default_date = datetime.now() + timedelta(weeks=1)
        self.expiration = DateEntry(master, width=12, background='darkblue', foreground='white', borderwidth=2, year=default_date.year, month=default_date.month, day=default_date.day, mindate=datetime.now())
        self.expiration.pack()

        self.clear_date_button = tk.Button(master, text="Clear Date", command=self.clear_date)
        self.clear_date_button.pack()

        self.redirect = tk.BooleanVar(value=True)
        self.redirect.trace_add("write", self.toggle_redirect_url_entry)
        self.redirect_btn = tk.Checkbutton(master, text="Use custom redirect URL", variable=self.redirect)
        self.redirect_btn.pack()

        self.redirect_url_services = ttk.Combobox(master, values=AVAIBLE_SERVICES, state="readonly")
        self.redirect_url_services.set(AVAIBLE_SERVICES[1])  # Set default value
        self.redirect_url_services.pack()

        self.redirect_url_alias = tk.StringVar(value="Enter redirect alias")
        self.redirect_url_entry = tk.Entry(master, textvariable=self.redirect_url_alias)
        self.redirect_url_entry.pack()

        self.generate_button = tk.Button(master, text="Generate Sharing Link", command=self.generate_link)
        self.generate_button.pack()

        self.link_label = tk.Label(master, text="")
        self.link_label.pack()

        self.toggle_redirect_url_entry()   

    def select_file(self):
        self.file_path = filedialog.askopenfilename()
        if self.file_path:
            self.file_label.config(text=f"Selected: {self.file_path}")
            self.redirect_url_alias.set(get_root_path(self.file_path).replace(" ", "_"))

    def toggle_redirect_url_entry(self, *args):
        if self.redirect.get():
            self.redirect_url_services.pack()
            self.redirect_url_entry.pack()
        else:
            self.redirect_url_services.pack_forget()
            self.redirect_url_entry.pack_forget()

    def clear_date(self):
        self.expiration.set_date(None)

    def generate_link(self):
        if not hasattr(self, 'file_path'):
            messagebox.showwarning("Warning", "Please select a file first.")
            return

        try:
            link = self.onedrive.create_sharing_link(self.file_path, self.sharing_type.get(), self.scope.get())

            if self.redirect.get():
                redirect = Redirect(link, self.redirect_url_alias.get())
                if self.redirect_url_services.get() == "tinyurl":
                    link = redirect.tinyUrl()
                elif self.redirect_url_services.get() == "infinityfree":
                    link = redirect.infinityfree(os.getenv("INFINITYFREE_URL"))
            self.link_label.config(text=f"Generated Link: {link}")
            pyperclip.copy(link)
            notification("Link Generated", f"Link copied to clipboard: {link}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = OneDriveSharingApp(root)
    root.mainloop()