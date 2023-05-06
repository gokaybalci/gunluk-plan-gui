import requests
import docx
from bs4 import BeautifulSoup

from docx import Document

kademe = input("Kademe seçiniz: ") # Grade level
hafta = input("Hafta seçiniz: ") # Week number for daily plans
oisim = input("İsminiz: ") # Name of the teacher
misim = input("İdareci ismi: ") # Name of the principle
dot_count = 0  # Counter for number of occurrences of 'dots' in given documents

kademe_url = str("https://www.ingilizceciyiz.com/"+kademe+"-sinif-ingilizce-gunluk-plan/")

# Requests URL and get response object
response = requests.get(kademe_url)

# Parse text obtained
soup = BeautifulSoup(response.text, 'html.parser')
   
eslesen_haftalar = soup.find_all(lambda tag: len(tag.find_all('a')) == 0 and hafta +". Hafta" in tag.text)

for link in eslesen_haftalar:
    if ('.docx' in link.get('href', [])):
        print("Dosya indiriliyor... ")
    
    # Get response object for link
        response = requests.get(link.get('href'))

    # Write content in docx file
        docx = open(kademe + ". Sınıf " + hafta + ". Hafta" + ".docx", 'wb')
        docx.write(response.content)
        docx.close()
print("İşlem tamamlandı.")


# Load the document and change teacher and principal name

document = Document(kademe + ". Sınıf " + hafta + ". Hafta" + ".docx")
for paragraph in document.paragraphs:
    if '…' in paragraph.text:
        dot_count += 1
    if dot_count == 1:
        paragraph.text = paragraph.text.replace('…', oisim, 1)
    elif dot_count == 2:
        paragraph.text = paragraph.text.replace('…', misim, 1)
    
document.save(kademe + ". Sınıf " + hafta + ". Hafta" + ".docx")

