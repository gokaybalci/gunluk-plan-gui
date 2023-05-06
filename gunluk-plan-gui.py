import tkinter as tk
import requests
import docx

from bs4 import BeautifulSoup
from tkinter import messagebox
from docx import Document


def download_daily_plan():
    kademe = kademe_entry.get()
    hafta = hafta_entry.get()
    oisim = oisim_entry.get()
    misim = misim_entry.get()

    # Gives an error if there is a missing input
    if kademe == "" or hafta == "" or oisim == "" or misim == "":
        messagebox.showerror("Error", "Lütfen eksik alanları doldurunuz.")
        return

    kademe_url = str("https://www.ingilizceciyiz.com/"+kademe+"-sinif-ingilizce-gunluk-plan/")

    # Requests URL and get response object
    response = requests.get(kademe_url)

    # Parse text obtained
    soup = BeautifulSoup(response.text, 'html.parser')

    eslesen_haftalar = soup.find_all(lambda tag: len(tag.find_all('a')) == 0 and hafta +". Hafta" in tag.text)

    for link in eslesen_haftalar:
        if ('.docx' in link.get('href', [])):
            status_label.config(text="Dosya indiriliyor...")
            # Get response object for link
            response = requests.get(link.get('href'))

            # Write content in docx file
            docx = open(kademe + ". Sınıf " + hafta + ". Hafta" + ".docx", 'wb')
            docx.write(response.content)
            docx.close()

    status_label.config(text="İşlem tamamlandı.")

    # Load the document and change teacher and principal name
    document = Document(kademe + ". Sınıf " + hafta + ". Hafta" + ".docx")
    dot_count = 0  # Counter for number of occurrences of 'dots' in given documents
    for paragraph in document.paragraphs:
        if '…' in paragraph.text:
            dot_count += 1
        if dot_count == 1:
            paragraph.text = paragraph.text.replace('…', oisim, 1)
        elif dot_count == 2:
            paragraph.text = paragraph.text.replace('…', misim, 1)

    document.save(kademe + ". Sınıf " + hafta + ". Hafta" + ".docx")

    status_label.config(text="Dosya güncellendi.")

root = tk.Tk()
root.title("Gunluk Plan GUI")

# Set the window size
root.geometry("400x300")

# create input widgets for user inputs
kademe_label = tk.Label(root, text="Kademe giriniz:")
kademe_label.pack()
kademe_entry = tk.Entry(root)
kademe_entry.pack()

hafta_label = tk.Label(root, text="Günlük plan haftası:")
hafta_label.pack()
hafta_entry = tk.Entry(root)
hafta_entry.pack()

oisim_label = tk.Label(root, text="Öğretmen ismi:")
oisim_label.pack()
oisim_entry = tk.Entry(root)
oisim_entry.pack()

misim_label = tk.Label(root, text="İdareci ismi:")
misim_label.pack()
misim_entry = tk.Entry(root)
misim_entry.pack()

# create a button to execute the code
execute_button = tk.Button(root, text="İndir", command=download_daily_plan)
execute_button.pack()

# Create a label to show the status of the process
status_label = tk.Label(root, text="")
status_label.pack()

root.mainloop()
