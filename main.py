from pathlib import Path
from uploading_to_excel import upload_to_excel
from create_dictionary_from_sheet_string import create_dictionary_from_sheet_string
from process_document_text import process_text
from choosing_files import return_lists_for_filepath_and_filestem
# Take file, prcoess, and create and print dictionary.


# path of folder that contains the word documents.
folder_path = Path(r'C:\Users\brennan.ukrainetz\Desktop\Mail Merge Project Temporary Folders\TAKE1')
file_paths, file_stems = return_lists_for_filepath_and_filestem(folder_path)
# print(file_stems)

list_of_dictionaries = []
for file_path in file_paths:
    haystack =  process_text(file_path)
    # print(haystack)
    dictionary = create_dictionary_from_sheet_string(haystack)
    list_of_dictionaries.append(dictionary)

upload_to_excel(list_of_dictionaries,file_stems)

# Testing excel and dictionary together
# file_path = r"C:\Users\brennan.ukrainetz\Desktop\Mail Merge Project Temporary Folders\Mail MergeTEST Filesfile_stemsdocx"
# haystack =  process_text(file_path)
# dictionary = create_dictionary_from_sheet_string(haystack)
# list_of_dicts = [dictionary]
# upload_to_excel(list_of_dicts,"PT-2685")



