 

from pathlib import Path






# regex help documents: https://docs.python.org/3/howto/regex.html and https://docs.python.org/3/library/re.html#match-objects
# Look for module level functions for ones where creating an object is not required.

folder_path = Path(r'C:\Users\brennan.ukrainetz\Desktop\Mail Merge Project Temporary Folders\Mail MergeTEST Files')



# Loop through all the title names in the folder and only retrieve the cal sheets, not the cover sheets or templates. (Warning: There are a few templates such as TT- template and RTD 1-element templates that have a dash in their names.)
file_list = []
for file in folder_path.glob('*-*.docx'):
    print(file.stem)
    file_list.append(file.stem.upper()) #makes sure all words are capitalized anyways.
print(file_list)

for title_of_doc_without_extension in file_list:
    if ("TE" or "TI") in title_of_doc_without_extension:
        instrument_type = "RTD"
    elif "T-" in title_of_doc_without_extension:
        instrument_type = "TRANSMITTER"
    elif "V-" in title_of_doc_without_extension:
        instrument_type = "VALVE"
    elif title_of_doc_without_extension.startswith("LS"):
        instrument_type = "LS_SWITCH"
    elif title_of_doc_without_extension.startswith("P") and ("S" in title_of_doc_without_extension):
        instrument_type = "PS_SWITCH"
    print(instrument_type)




