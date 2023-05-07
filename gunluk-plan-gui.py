import tkinter as tk
import requests
import docx
import tkinter.ttk as ttk

from bs4 import BeautifulSoup
from tkinter import messagebox, ttk
from docx import Document


def download_daily_plan():
    kademe = kademe_entry.get()
    hafta = hafta_entry.get()
    oisim = oisim_entry.get()
    misim = misim_entry.get()

    # Gives an error if there is a missing input
    if kademe == "" or hafta == "" or oisim == "" or misim == "":
        messagebox.showerror("Hata!", "Lütfen eksik alanları doldurunuz.")
        return

    kademe_url = str("https://www.ingilizceciyiz.com/"+kademe+"-sinif-ingilizce-gunluk-plan/")

    # Requests URL and get response object
    response = requests.get(kademe_url)

    # Parse text obtained
    soup = BeautifulSoup(response.text, 'html.parser')



    eslesen_haftalar = soup.find_all(lambda tag: len(tag.find_all('a')) == 0 and str(hafta) +". Hafta" in tag.text)

    for link in eslesen_haftalar:
        if '.docx' in link['href']:
            response = requests.get(link['href'])
            docx = open(kademe + ". Sınıf " + hafta + ". Hafta" + ".docx", 'wb')
            docx.write(response.content)
            docx.close()
            break  # exit the loop after the first valid link is downloaded



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

    status_label.config(text="Dosyanız hazır.")

root = tk.Tk()
root.title("Gunluk Plan GUI")
root.resizable(False, False)

# Set the window size
root.geometry("350x375")

# font setup
input_font = ('Century Gothic', 13)
button_font = ('Century Gothic', 17)

# create input widgets for user inputs
kademe_label = tk.Label(root, text="Kademe giriniz:", font=input_font)
kademe_label.pack(side="top", padx=3, pady=3)
kademe_entry = tk.Entry(root)
kademe_entry.pack(side="top", padx=1, pady=1)

hafta_label = tk.Label(root, text="Günlük plan haftası:", font=input_font)
hafta_label.pack(side="top", padx=3, pady=3)
hafta_entry = tk.Entry(root)
hafta_entry.pack(side="top", padx=1, pady=1)

oisim_label = tk.Label(root, text="Öğretmen ismi:", font=input_font)
oisim_label.pack(side="top", padx=3, pady=3)
oisim_entry = tk.Entry(root)
oisim_entry.pack(side="top", padx=1, pady=1)

misim_label = tk.Label(root, text="İdareci ismi:", font=input_font)
misim_label.pack(side="top", padx=3, pady=3)
misim_entry = tk.Entry(root)
misim_entry.pack(side="top", padx=1, pady=1)

# create a button to execute the code
execute_button = tk.Button(root, text="İndir", command=download_daily_plan, font=button_font)
execute_button.pack(side="top", padx=5, pady=25)

# Create a label to show the status of the process
status_label = tk.Label(root, text="", font=button_font)
status_label.pack()

root.mainloop()
