import re

# regex help documents: https://docs.python.org/3/howto/regex.html and https://docs.python.org/3/library/re.html#match-objects
# Look for module level functions for ones where creating an object is not required.

def return_lists_for_filepath_and_filestem(folder_path):



    # Loop through all the title names in the folder and only retrieve the cal sheets, not the cover sheets or templates. 
    # (Warning: There are a few templates such as TT- template and RTD 1-element templates that have a dash in their names.)
    list_filestem = []
    list_filepath = []
    for file in folder_path.glob('*-*.docx'): #glob for recursive search.
        if re.findall(r"[\w]{1,6}-[\d]+",file.stem): #If it has a 'tag-like' appearance (further screening)
            # print(file.stem)
            list_filestem.append(file.stem.upper()) #makes sure all words are capitalized anyways.
            list_filepath.append(str(file))
    
    
    # print(list_filestem)
    # print(list_filepath)
    return list_filepath,list_filestem #return as tuple of lists.



    # Deprecated: Not useful anymore. 
    # for title_of_doc_without_extension in list_filestem:
    #     if ("TE" or "TI") in title_of_doc_without_extension:
    #         instrument_type = "RTD"
    #     elif "T-" in title_of_doc_without_extension:
    #         instrument_type = "TRANSMITTER"
    #     elif "V-" in title_of_doc_without_extension:
    #         instrument_type = "VALVE"
    #     elif title_of_doc_without_extension.startswith("LS"):
    #         instrument_type = "LS_SWITCH"
    #     elif title_of_doc_without_extension.startswith("P") and ("S" in title_of_doc_without_extension):
    #         instrument_type = "PS_SWITCH"
    #     print(instrument_type)



# path of folder that contains the word documents.
# folder_path = Path(r'C:\Users\brennan.ukrainetz\Desktop\Mail Merge Project Temporary Folders\Mail MergeTEST Files')

# print(return_lists_for_filepath_and_filestem(folder_path))
