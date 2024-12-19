from docx import Document
from docx2pdf import convert
import yaml
import os

def main():
    name = ""

    doc = Document("E:\\Studijos\\Praktika\\Script_2\\Konf Sutartis.docx")
    shared_data = r"E:\\Studijos\\Praktika\\Script_2\\shared_data.yml"

    if(os.path.exists(shared_data)):
        with open("shared_data.yml", "r", encoding="utf-8") as file:
            data = yaml.safe_load(file)
    
        name = data["name"]
        DOB = data["dob"]
        address = data["address"]
        tel_nr = input("Phone Number: ")
        contract_number = input("Contract number: ")
        contract_date = input("Contract signing date: ")
        os.remove(shared_data)
    else:
        name = input("Student Name: ")
        tel_nr = input("Phone Number: ")
        contract_number = input("Contract number: ")
        contract_date = input("Contract signing date: ")
        DOB = input("Date Of Birth: ")
        address = input("Address: ")
        
    change(doc, "[vardas]", name)
    change(doc, "[tel_nr]", tel_nr)
    change(doc, "[sutarties_nr]", contract_number)
    change(doc, "[pasirasymo_data]", contract_date)
    change(doc, "[gim_data]", DOB)
    change(doc, "[adresas]", address)
    write(doc, name)


def change(doc, placeholder, replacement_word):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, replacement_word)

            
def write(doc, name):
    doc.save(r"E:\\Studijos\Praktika\\Script_2\\Updated Konf Sutartis.docx")
    convert(r"E:\\Studijos\Praktika\\Script_2\\Updated Konf Sutartis.docx", fr"E:\\Studijos\Praktika\\Script_2\\{name}.pdf")
    os.remove(r"E:\\Studijos\Praktika\\Script_2\\Updated Konf Sutartis.docx")
    


main()




    