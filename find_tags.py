import docx2txt
import re

haystack = """

Tag	PDSH-0445	Calibrated range    103 KPA  

Make:                                         
  
Model:	Serial 

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

class Sheet():
    def __init__(self) -> None:
        normal_flags = re.IGNORECASE|re.ASCII |re.MULTILINE
        special_flags = re.IGNORECASE|re.ASCII |re.MULTILINE | re.DOTALL

        self.fields = ({"Tag":["Tag:?(.+)(?=Calib)",normal_flags], 
        "Calibrated Range":["Calibrated Range:?(.+)",normal_flags],
        "Make":["Make:?(.+)", normal_flags],
        "Model":["Model:?(.+)(?=Serial)", normal_flags], 
        "Serial":["Serial:?(.+)",normal_flags],
        "Size":["Size:?(.+)", normal_flags],
        "DEV ID":["DEV ID:?(.+)", normal_flags],
        "Blurb": ["Scotford ASU(.*)Tag",special_flags]})
        #Parentheses for readability.https://stackoverflow.com/questions/4172448/is-it-possible-to-break-a-long-line-to-multiple-lines-in-python

# ! FIRST TO-DO: Create a Sheet values dictionary that will scrape the keys off of the self.fields dictionary, and use these to put default values of "N/A" for each field. Then, if a match is found in the for loop, overwrite the "N/A" with the actual value. Then after all that is done I can work on inputting into a spreadsheet.

class TESheet(Sheet):
    def __init__(self):
        super().__init__()
        # Define the fields we actually care about for a TE sheet.
        self.field_list = ["Tag", "Calibrated Range", ]

sheet = TESheet()
print(sheet.field_list)






for field_name in sheet.field_list:
    
    match = re.findall( sheet.fields[field_name][0], #Access regex pattern for that field
    haystack,
    sheet.fields[field_name][1]) # Access flags for that field 

    if match:
        print(match)

        continue

    # matches = re.findall(fields[field_name][0],haystack,fields[field_name][1])
    # # NOTE: Eventually add in a unicode thing for UnicodeEncodeError: 'charmap' codec can't encode character '\uf0ad' in position 13: character maps to <undefined>
    # for match in matches:
    #     print(match)


