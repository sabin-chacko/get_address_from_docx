import os
import docx
import requests
from urllib.parse import quote
#location of excel files
data_folder = "data_folder"

def get_docx_filename() -> list:
    """to get docx file from the folder

    Returns:
        list: list of docx file name
    """    
    file_list = []

    if os.path.exists(data_folder):
        for f in os.listdir(data_folder):
            #check f is of type file and f is of xlsx format
            if os.path.isfile(os.path.join(data_folder, f)) and\
                f.split(".")[1] == "docx":
                file_list.append(f)
    return file_list

def get_address_from_each_line_with_wit(line: str) -> str:
    """process each line to get address
        using wit.ai get address string

    Args:
        line (str): line from document

    Returns:
        str: address
    """
    try:
        if not line:
            return ""
        addr = ""
        # calling an app created in wit.ai to check whether the given string contains address
        resp = requests.get("https://api.wit.ai/message?v=20220507&q=" + quote(line[0:260]), headers = {"Authorization": "Bearer HMISHBBUQ3SGZNGCSVW7FRRCP5WZD35R"})
        data = resp.json()
        if 'wit$location:location' in data["entities"]:
            addr = " ".join([x['body'] for x in data["entities"]['wit$location:location']])
        return addr
    except Exception as e:
        return ""
    


def get_address_from_docx(filename: str) -> str:
    """process each file and return address

    Args:
        filename (str): filename

    Returns:
        str: address
    """
    # initialise address string
    address = ""
    
    doc = docx.Document(os.path.join(data_folder, filename))
    for i in doc.paragraphs:
        address_in_each_line = get_address_from_each_line_with_wit(i.text)
        if address_in_each_line:
            address += "\n" + address_in_each_line
    return address

def process_docx_file(file_list: list) -> None:
    """proccess docx files

    Args:
        file_list (list): list of docx filename
    """
    for f in file_list:
        address_text = get_address_from_docx(f)
        #printing the address obtained from each file
        print(f'address_text from {f} :{address_text.strip()}')


def main():
    """main method
    """
    file_list = get_docx_filename()
    process_docx_file(file_list)


if __name__ == "__main__":
    main()