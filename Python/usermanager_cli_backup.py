from pyad import *
import os
import csv
import win32net
import win32netcon
import win32security
import ntsecuritycon
import requests
import winreg

registry_path = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Accent"
value_name = "AccentColorMenu"
accent_color = "ff00b9ff"

try:
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path, 0, winreg.KEY_SET_VALUE)
    winreg.SetValueEx(key, value_name, 0, winreg.REG_DWORD, int(accent_color, 16))
except FileNotFoundError:
    print("Registry key not found.")
except PermissionError:
    print("Permission denied to modify registry key.")
except Exception as e:
    print(f"Error occurred: {e}")

notepadpp_path = "C:\\Program Files\\Notepad++\\notepad++.exe"
if os.path.exists(notepadpp_path):
    print("Notepad++ is already installed, skipping..")
else:
    download_url = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.6.4/npp.8.6.4.Installer.x64.exe"
    installer_filename = "npp.8.6.4.Installer.x64.exe"
    installer_path = os.path.join(os.getcwd(), installer_filename)

    print("Downloading Notepad++ installer...")
    with open(installer_path, 'wb') as f:
        response = requests.get(download_url)
        f.write(response.content)

    print("Installing Notepad++...")
    os.system(installer_path + " /S")
    os.remove(installer_path)

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

pyad.set_defaults(ldap_server="testbedrijf.local")
ou_dn = "OU=TestUsers,DC=testbedrijf,DC=local"
ou = pyad.adcontainer.ADContainer.from_dn(ou_dn)

with open('klanten.csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    
    # Loop through each row in the CSV file
    for row in reader:
        # Access values by column names
        personeelsnummer = row['Personeelsnummer']
        roepnaam = row['Roepnaam']
        tussenvoegsel = row['Tussenvoegsel']
        achternaam = row['Achternaam']
        adres = row['Adres']
        pc = row['PC']
        plaats = row['Plaats']
        tel = row['Tel']
        geboortedatum = row['Geboortedatum']

        new_user = pyad.aduser.ADUser.create(personeelsnummer, container_object=ou)

        new_user.update_attribute("givenName", row['Roepnaam'])
        new_user.update_attribute("sn", f"{row['Tussenvoegsel']} {row['Achternaam']}") # Fix space bij gebruikers zonder tussenvoegsel.
        new_user.update_attribute("description", row['Personeelsnummer'])

        new_user.update_attribute("streetAddress", row['Adres'])
        new_user.update_attribute("postalCode", row['PC'])
        new_user.update_attribute("l", row['Plaats'])
        new_user.update_attribute("telephoneNumber", row['Tel'])

        os.makedirs('C:\Shares\S-' + personeelsnummer, exist_ok=True)
        create_share("S-" + personeelsnummer, 'C:\Shares\S-' + personeelsnummer, personeelsnummer)
        print("User created", personeelsnummer, roepnaam, tussenvoegsel, achternaam)