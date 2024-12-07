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
            data = yaml.load(file)
    
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
        
    change(doc, name, tel_nr, contract_date, contract_number, DOB, address)
    write(doc, name)


# Changes all the [templates] to chosen words
def change(doc, name, tel_nr, contract_number, contract_date, DOB, address):
    for paragraph in doc.paragraphs:
        if "[vardas]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[vardas]", name)
        if "[tel_nr]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[tel_nr]", tel_nr)
        if "[sutarties_nr]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[sutarties_nr]", contract_number)
        if "[pasirasymo_data]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[pasirasymo_data]", contract_date)
        if "[gim_data]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[gim_data]", DOB)
        if "[adresas]" in paragraph.text:
            paragraph.text = paragraph.text.replace("[adresas]", address)
            

def write(doc, name):
    doc.save(r"E:\\Studijos\Praktika\\Script_2\\Updated Konf Sutartis.docx")
    convert(r"E:\\Studijos\Praktika\\Script_2\\Updated Konf Sutartis.docx", fr"E:\\Studijos\Praktika\\Script_2\\{name}.pdf")
    os.remove(r"E:\\Studijos\Praktika\\Script_2\\Updated Konf Sutartis.docx")
    


main()




    