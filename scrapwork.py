import docx2txt
import re
from pprint import pprint

def create_dictionary_from_sheet_string():


    haystack = """

    Air Liquide Canada
    Switch Calibrations
    Scotford ASU
    N39/40 High L.O. Outlet Filter DP


    Tag	PDSH-0445	Calibrated range    103 KPA  

    Make: MAKEY MAKER BOI              DEV ID: 809123                    
    
    Model: QUADRA POWER II	Serial: Chunky BOI 1908231 

    Calibration Data

    FOUND   	 LEFT

        RESET


    Physical Checks
    	Inspected conduit and wiring
    	Checked wire terminals
    	Performed visual checks

    Calibration Equipment used



    Calibrated By


    Comments:




    Instrument Tech: 	       Reviewed by:                                 

    Date:	       Date:                                                    

    """



    delimiter_list = ["SCOTFORD ASU", "TAG", "CALIBRATED RANGE","MAKE","MODEL","SERIAL", "CALIBRATION DATA"]
    # special_field = []
    # Add these words to list, if they exist in the document. Won't have both at the same time.
    for word in ["DEV ID", "Size"]:
        if word in haystack:
            delimiter_list.insert(delimiter_list.index("MAKE")+1,word)
            # special_field.append(word)


    # print(delimiter_list)

    # Initializing list, converting all text to uppercase.
    leftover_string = haystack.upper()
    wanted_strings_list = []

    # Split string by delimiters and take the strings we want.
    for delimiter in delimiter_list:
        list_of_splitted_strs = leftover_string.split(delimiter,1)
        wanted_strings_list.append(list_of_splitted_strs[0])
        leftover_string = list_of_splitted_strs[1]
    wanted_strings_list.pop(0) #Getting rid of the first string which isn't useful.

    cleaned_list = [string.strip(' \n\t:') for string in wanted_strings_list]
    # print(cleaned_list)

    # Preparing the list for converting to dictionary.
    delimiter_list[0] = "BLURB" #Replace SFD ASU text
    dictionary_keys = delimiter_list

    # Zip the keys and values together as a dictionary and return.
    return dict(zip(dictionary_keys,cleaned_list))

    # ENCODING FOR WHEN ERRORS OCCUR IN READING WEIRD ASCII CHARACTERS.
    # for entry in leftover_string:
    #     # print(entry)
    #     print(entry.encode('ascii','ignore'))
    #     # print("\n")
