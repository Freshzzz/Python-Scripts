import re
from docx import Document
import yaml
doc = Document("E:\\Studijos\\Praktika\\Script_1\\kolegijos praktikos Sutartis.docx")


# Clears the previous data from the output.txt file
open('output.txt', 'w', encoding="utf-8").close()

def main():
    
    # Variables
    allText: list[str] = []
    allNames: list[str] = []
    allDOB: list[str] = []
    allAdresses: list[str] = []
    
    # Reads the text from a .docx file
    for paragraph in doc.paragraphs:
        allText.append(paragraph.text)
    
    # Reads data from a .yml file
    with open("config.yml", "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
        
    name_start = config['name_start']
    name_end = config['name_end']
    
    dob_start = config['dob_start']
    
    add_start = config['address_start']
    add_end = config['address_end']

    content = " ".join(allText)
    words = content.split()

    # Finds all the Names, DOBs & Adresses
    for i in range(len(words)-1):
        # Gets rids of the [()] on the side of the string
        words[i] = re.sub(r"[()]", "", words[i])
        if(words[i] in name_start):
            if(words[i+1].istitle()):
                name_list(i, words, allNames, name_end)   
        if(words[i] == dob_start):
            dob_list(i, words, allDOB)
        if(words[i] == add_start):
            adress_list(i, words, allAdresses, add_end)
            
    write(allNames, allDOB, allAdresses)
    

# Saves First & Last names
def name_list(i, words, allNames, name_end):
    stop = 0
    y = 1
    temporary_string = ""
        
    while stop == 0:
        if(words[i+y] not in name_end):
            temporary_string = temporary_string + " " + words[i+y]
        if(words[i+y] in name_end):
            allNames.append(temporary_string.strip(","))
            break
        y += 1
        

# Saves the Date Of Birth
def dob_list(i, words, allDOB):
    allDOB.append(words[i+1])
    

# Saves the address
def adress_list(i, words, allAdresses, add_end):
    stop = 0
    y = 1
    tempAdress = ""
    while stop == 0:
        if(words[i+y] not in add_end):
            if(words[i+y].endswith(',') and words[i+y].istitle()):
                words[i+y] = words[i+y].rstrip(',')
            tempAdress = tempAdress + " " + words[i+y]
        if(words[i+y] in add_end):
            allAdresses.append(tempAdress)
            break
        y += 1
    

# Saves the data to the output.txt file
def write(allNames, allDOB, allAdresses):
    with open('E:\\Studijos\\Praktika\\Script_1\\output.txt', 'a', encoding='utf-8') as f:
        print("Printing")
        f.write('	Names \n')
        f.write('\n'.join(allNames) + '\n')
        f.write('\n	Date Of Birth \n')
        f.write('\n'.join(allDOB) + '\n')
        f.write('\n 	Adresses \n')
        f.write('\n'.join(allAdresses))


main()


    

