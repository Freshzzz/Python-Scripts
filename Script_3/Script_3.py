import subprocess
import os
import yaml

shared_data_PATH = r"E:\\Studijos\\Praktika\\Script_2\\shared_data.yml"
output_file_PATH = r'E:\\Studijos\\Praktika\\Script_1\\output.txt'

# Saves the chosen data from output.txt file to a .json file
def save_to_file(name, dob, address):
    data = {
        "name": name,
        "dob": dob,
        "address": address
        }
    
    with open(shared_data_PATH, "w", encoding="utf-8") as file:
        yaml.dump(data, file, default_flow_style=False)

# User choice function
def choose(options):
    choice = 0
    while not (choice > 0 and choice < options+1):
        try:
            choice = int(input("Your choice: "))
        except ValueError:
            print("Please enter a valid number")
    return choice
            

# Reads from the output.txt file
def output_file_read(names, dobs, adresses):
    try:
        f = open(output_file_PATH, encoding="utf-8")
        for x in f:
            x = x.lstrip()
            x = x.rstrip()
        
            match x:
                case "Names":
                    category = 1
                    continue
                case "Date Of Birth":
                    category = 2
                    continue
                case "Adresses":
                    category = 3
                    continue
        
            if(x != ''):
                match category:
                    case 1:
                        names.append(x)
                    case 2:
                        dobs.append(x)
                    case 3:
                        adresses.append(x)
    
    except FileNotFoundError:
        print("output.txt file was not found!")
    except Exception as e:
        print("Error has occured while reading the output.txt file")
        

# Gives user the choice to choose the data they want to use
def choose_info(all_options):
    os.system('cls')
    print("Choose the option you want to use:")
    for i, option in enumerate(all_options, 1):
        print(f"{i} - {option}")
    
    index = choose(len(all_options))
    return all_options[index - 1]


# Launches the 1st Script
def first_Script():
    script_1_path = os.path.abspath("../Script_1/Script_1.py")
    script_1_dir = os.path.dirname(script_1_path)
    subprocess.run(
        ["python", script_1_path],
        cwd=script_1_dir
        )


# Launches the 2nd Script
def second_script():
    script_2_path = os.path.abspath("../Script_2/Script_2.py")
    script_2_dir = os.path.dirname(script_2_path)
    subprocess.run(
        ["python", script_2_path],
        cwd=script_2_dir
        )

def main():
    names: list[str] = []
    dobs: list[str] = []
    adresses: list[str] = []
    options = 2
    
    chosen_name = ""
    chosen_dob = ""
    chosen_ad = ""
    
    print("1 - Choose data from a .docx file")
    print("2 - Enter your information")
    
    choice = choose(options)
    
    if(choice == 1):
        first_Script()
        output_file_read(names, dobs, adresses)
        chosen_name = choose_info(names)
        chosen_dob = choose_info(dobs)
        chosen_ad = choose_info(adresses)
        
        save_to_file(chosen_name, chosen_dob, chosen_ad)
        second_script()
    if(choice == 2):
        subprocess.run(["python", "../Script_2/Script_2.py"])
        
    

main()