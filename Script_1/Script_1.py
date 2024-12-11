import re
from docx import Document
import yaml


# Clears the previous data from the output.txt file
open('output.txt', 'w', encoding="utf-8").close()

def load_config(file_path):
    required_keys = ["name_start", "name_end", "dob_start", "address_start", "address_end"]
    try:
        # Reads data from a .yml file
        with open("config.yml", "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
            
        for key in required_keys:
            if key not in config:
                raise KeyError(f"Missing: {key}")
            
        return config
    
    except FileNotFoundError:
        print(f"Error: Configuration file '{file_path}' not found. Please provide a valid path.")
        exit(1)  # Exit the program if the file is missing
    except KeyError as e:
        print(f"Error in config.yml: {e}")
        exit(1)
    except yaml.YAMLError as e:
        print(f"Error reading config.yml: {e}")
        exit(1)


def load_doc(file_path):
    try:
        doc = Document(file_path)
        return doc
    except FileNotFoundError:
        print("Error: File Not Found. Please check the path")
        exit(1)
    except PermissionError:
        print("Error: Insufficient permissions to access the file")
        exit(1)
    except Exception as e:
        print("Error: There has been an error opening the file")
        exit(1)
    
        

def main():
    
    # Variables
    allText: list[str] = []
    allNames: list[str] = []
    allDOB: list[str] = []
    allAdresses: list[str] = []
    
    doc = load_doc("E:\\Studijos\\Praktika\\Script_1\\kolegijos praktikos Sutartis.docx")
    
    # Reads the text from a .docx file
    for paragraph in doc.paragraphs:
        allText.append(paragraph.text)
    
    config = load_config("config.yml")
    
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
    y = 1
    temporary_words = []
        
    while True:
        if(not_in_list(words[i+y], name_end)):
            temporary_words.append(words[i+y])
        if(in_list(words[i+y], name_end)):
            temporary_string = " ".join(temporary_words)
            allNames.append(temporary_string.strip(","))
            break
        y += 1
        

# Saves the Date Of Birth
def dob_list(i, words, allDOB):
    allDOB.append(words[i+1])
    

# Saves the address
def adress_list(i, words, allAdresses, add_end):
    y = 1
    temporary_words = []
    while True:
        if(not_in_list(words[i+y], add_end)):
            if(words[i+y].endswith(',') and words[i+y].istitle()):
                words[i+y] = words[i+y].rstrip(',')
            temporary_words.append(words[i+y])
        if(in_list(words[i+y], add_end)):
            temporary_address = ' '.join(temporary_words)
            allAdresses.append(temporary_address)
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
        

def not_in_list(word, word_list):
    return word not in word_list

def in_list(word, word_list):
    return word in word_list


main()


    

