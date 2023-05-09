import tkinter as tk
import requests
import docx
import tkinter.ttk as ttk

from bs4 import BeautifulSoup
from tkinter import messagebox, ttk
from docx import Document


def download_daily_plan():
    kademe_list = []
    for k in kademe_entry.get().replace(" ", "").split(","):
        k = k.strip()
        if k.isdigit() and int(k) > 0 and int(k) <= 8:
            kademe_list.append(int(k))
    hafta_list = []
    for h in hafta_entry.get().replace(" ", "").split(","):
        h = h.strip()
        if h.isdigit() and int(h) > 0 and int(h) <= 52:
            hafta_list.append(int(h))
    oisim = oisim_entry.get()
    misim = misim_entry.get()

    # Gives an error if there is a missing input
    if kademe_list == "" or hafta_list == "" or oisim == "" or misim == "":
        messagebox.showerror("Hata!", "Lütfen eksik alanları doldurunuz.")
        return


    for kademe_name in kademe_list:
        kademe_url = f"https://www.ingilizceciyiz.com/{kademe_name}-sinif-ingilizce-gunluk-plan/"
        # Requests URL and get response object
        response = requests.get(kademe_url)

        # Parse text obtained
        soup = BeautifulSoup(response.text, 'html.parser')

        for hafta in hafta_list:
            # Search for the link that contains the week number and download the file
            eslesen_haftalar = soup.find_all(lambda tag: len(tag.find_all('a')) == 0 and str(hafta) + ". Hafta" in tag.text)
            for link in eslesen_haftalar:
                if '.docx' in link['href']:
                    response = requests.get(link['href'])
                    docx = open(f"{kademe_name}. Sınıf {hafta}. Hafta.docx", 'wb')
                    docx.write(response.content)
                    docx.close()
                    break  # exit the loop after the first valid link is downloaded

            # Load the downloaded document and change teacher and principal name
            document = Document(f"{kademe_name}. Sınıf {hafta}. Hafta.docx")
            dot_count = 0  # Counter for number of occurrences of 'dots' in given documents
            for paragraph in document.paragraphs:
                if '…' in paragraph.text:
                    dot_count += 1
                if dot_count == 1:
                    paragraph.text = paragraph.text.replace('…', oisim, 1)
                elif dot_count == 2:
                    paragraph.text = paragraph.text.replace('…', misim, 1)

            document.save(f"{kademe_name}. Sınıf {hafta}. Hafta.docx")


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

hafta_label = tk.Label(root, text="Haftalar:", font=input_font)
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
