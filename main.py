import os
import win32com.client
from pprint import pprint


folder = r"C:\Users\asorensen\OneDrive - CVRx Inc\2024_COMP_OPS\COMP_STATEMENTS"

search_terms = [
    "abesamis",
    "campos",
    "avalos"
]

template = r"C:\Users\asorensen\OneDrive - CVRx Inc\Desktop\msg.oft"


def get_files(directory: str, terms: list):
    files = []
    for term in terms:
        for dirpath, dirnames, filenames in os.walk(directory):
            for file in filenames:
                if term.lower() in file.lower():
                    if "preliminary".lower() not in file.lower():
                        full_path = os.path.join(dirpath, file)
                        files.append(full_path)
    return files


def send_email():
    outlook = win32com.client.Dispatch('Outlook.Application')
    mail = outlook.CreateItemFromTemplate(template)
    mail.Subject = "Statements for Abesamis, Da Silva Campos, Avalos"
    mail.To = 'kryan@cvrx.com'
    # mail.CC = 'jmoore@cvrx.com'
    for file in get_files(directory=folder, terms=search_terms):
        mail.Attachments.Add(file)
    mail.Send()
