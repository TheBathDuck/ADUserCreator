import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from pyad import *
import win32net
import win32netcon
import win32security
import ntsecuritycon
import os

def create_share(share_name, path, userid):
    share_info = {
        'netname': share_name,
        'path': path,
        'type': win32netcon.STYPE_DISKTREE,
    }
    
    try:
        win32net.NetShareAdd(None, 2, share_info)
        print(f"Share '{share_name}' created successfully.")

        user_sid = win32security.LookupAccountName(None, userid)[0]
        sd = win32security.GetFileSecurity(path, win32security.DACL_SECURITY_INFORMATION)
        dacl = win32security.ACL()
        admin_sid = win32security.LookupAccountName(None, "Administrators")[0]
        dacl.AddAccessAllowedAce(win32security.ACL_REVISION_DS, ntsecuritycon.FILE_ALL_ACCESS, admin_sid)
        dacl.AddAccessAllowedAce(win32security.ACL_REVISION_DS, ntsecuritycon.FILE_ALL_ACCESS, user_sid)
        sd.SetSecurityDescriptorDacl(1, dacl, 0)
        win32security.SetFileSecurity(path, win32security.DACL_SECURITY_INFORMATION, sd)


        print(user_sid)

    except Exception as e:
        print(f"Failed to create share '{share_name}': {e}")

class BasicGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Cerberus - User Manager")
        self.master.iconbitmap('logo.ico')

        self.load_button = tk.Button(self.master, text="Load File", command=self.load_file)
        self.load_button.grid(row=0, column=0, padx=5, pady=5)

        self.create_button = tk.Button(self.master, text="Create", command=self.create_something)
        self.create_button.grid(row=0, column=1, padx=5, pady=5)

        self.clear_button = tk.Button(self.master, text="Clear", command=self.clear_users)
        self.clear_button.grid(row=0, column=2, padx=5, pady=5)

        self.close_button = tk.Button(self.master, text="Close", command=self.master.quit)
        self.close_button.grid(row=0, column=3, padx=5, pady=5)

        self.users_listbox = tk.Listbox(self.master, width=70, height=10)
        self.users_listbox.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                self.data = pd.read_csv(file_path)
                print("File loaded successfully!")
                self.display_users()
            except Exception as e:
                print("Error:", e)

    def clear_users(self):
        print("Clearing")
        self.data = None
        self.users_listbox.delete(0, tk.END)

    def display_users(self):
        self.users_listbox.delete(0, tk.END)
        for index, row in self.data.iterrows():
            user_info = f"{index} - {row['Personeelsnummer']} - {row['Roepnaam']} {row['Tussenvoegsel']} {row['Achternaam']}"
            self.users_listbox.insert(tk.END, user_info)

    def create_something(self):
        pyad.set_defaults(ldap_server="testbedrijf.local")
        ou_dn = "OU=TestUsers,DC=testbedrijf,DC=local"
        ou = pyad.adcontainer.ADContainer.from_dn(ou_dn)

        print(ou)

        for index, row in self.data.iterrows():
            personeelsnummer = str(row['Personeelsnummer'])
            roepnaam = row['Roepnaam']
            tussenvoegsel = row['Tussenvoegsel']
            achternaam = row['Achternaam']
            adres = row['Adres']
            pc = row['PC']
            plaats = row['Plaats']
            tel = row['Tel']
            geboortedatum = row['Geboortedatum']

            new_user = pyad.aduser.ADUser.create(personeelsnummer, container_object=ou)
            new_user.update_attribute("givenName", roepnaam)
            new_user.update_attribute("sn", f"{tussenvoegsel} {achternaam}") # Fix space bij gebruikers zonder tussenvoegsel.
            new_user.update_attribute("description", personeelsnummer)
            new_user.update_attribute("streetAddress", adres)
            new_user.update_attribute("postalCode", pc)
            new_user.update_attribute("l", plaats)
            new_user.update_attribute("telephoneNumber", tel)

            os.makedirs('C:\Shares\S-' + personeelsnummer, exist_ok=True)
            create_share("S-" + personeelsnummer, 'C:\Shares\S-' + personeelsnummer, personeelsnummer)
            print("User created", personeelsnummer, roepnaam, tussenvoegsel, achternaam)
        messagebox.showinfo("Success", "Users created successfully!")
        pass

def main():
    root = tk.Tk()
    app = BasicGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
