

def create_dictionary_from_sheet_string(sheet_string):

    sheet_string = sheet_string.upper()
    delimiter_list = ["SCOTFORD ASU", "TAG", "CALIBRATED RANGE","MAKE","MODEL","SERIAL", "CALIBRATION DATA"]
    # special_field = []
    # Add these words to list, if they exist in the document. Won't have both at the same time.
    for word in ["DEV ID", "SIZE"]:
        if word in sheet_string:
            delimiter_list.insert(delimiter_list.index("MAKE")+1,word)
            # special_field.append(word)


    # print(delimiter_list)

    # Initializing list, converting all text to uppercase.
    leftover_string = sheet_string.upper()
    wanted_strings_list = []

    # Split string by delimiters and take the strings we want.
    for delimiter in delimiter_list:
        list_of_splitted_strs = leftover_string.split(delimiter,1)
        wanted_strings_list.append(list_of_splitted_strs[0])
        if not len(list_of_splitted_strs) == 2: #If the split wasn't successful; return an error.
                # raise ValueError(f"Didn't find {delimiter} in {wanted_strings_list[1]}")
                wanted_strings_list = ["Can't Find: " + delimiter for value in delimiter_list] #Creates a list of proper length that tells me in the excel sheet what it couldn't find.
                break
        leftover_string = list_of_splitted_strs[1]
    wanted_strings_list.pop(0) #Getting rid of the first string which isn't useful.

    cleaned_list = [string.strip(' \n\t:') for string in wanted_strings_list]
    # print(cleaned_list)

    # Preparing the list for converting to dictionary.
    delimiter_list[0] = "BLURB" #Replace SFD ASU text
    dictionary_keys = delimiter_list

    # Zip the keys and values together as a dictionary and return. 
    # NOTE: Calibration data is cut off because the cleaned list is 1 shorter than the dictionary keys.
    return dict(zip(dictionary_keys,cleaned_list))

    # ENCODING FOR WHEN ERRORS OCCUR IN READING WEIRD ASCII CHARACTERS.
    # for entry in leftover_string:
    #     # print(entry)
    #     print(entry.encode('ascii','ignore'))
    #     # print("\n")


# haystack = """

#     Air Liquide Canada
#     Switch Calibrations
#     Scotford ASU
#     N39/40 High L.O. Outlet Filter DP


#     Tag	PDSH-0445	Calibrated range    103 KPA  

#     Make: MAKEY MAKER BOI              DEV ID: 809123                    
    
#     Model: QUADRA POWER II	Serial: Chunky BOI 1908231 

#     Calibration Data

#     FOUND   	 LEFT

#         RESET


#     Physical Checks
#     	Inspected conduit and wiring
#     	Checked wire terminals
#     	Performed visual checks

#     Calibration Equipment used



#     Calibrated By


#     Comments:




#     Instrument Tech: 	       Reviewed by:                                 

#     Date:	       Date:                                                    

#     """
