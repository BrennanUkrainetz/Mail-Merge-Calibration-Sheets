import docx2txt
import re


# document_text = docx2txt.process(r"C:\Users\brennan.ukrainetz\Desktop\Mail Merge Project Temporary Folders\Mail MergeTEST Files\LSL-0480.docx")
# document_text = docx2txt.process(r"C:\Users\brennan.ukrainetz\Desktop\Mail Merge Project Temporary Folders\Mail MergeTEST Files\PT-0783.docx")
document_text = docx2txt.process(r"C:\Users\brennan.ukrainetz\Desktop\Mail Merge Project Temporary Folders\Mail MergeTEST Files\PT-2685.docx")
# document_text = docx2txt.process(r"C:\Users\brennan.ukrainetz\Desktop\Mail Merge Project Temporary Folders\Mail MergeTEST Files\TE-2811C.docx")
# document_text = docx2txt.process(r"C:\Users\brennan.ukrainetz\Desktop\Mail Merge Project Temporary Folders\Mail MergeTEST Files\TI-0403.docx")
# document_text = docx2txt.process(r"C:\Users\brennan.ukrainetz\Desktop\Mail Merge Project Temporary Folders\Mail MergeTEST Files\TV-2801.docx")


# print(document_text)
# Pattern = re.compile(r"Calibrated range:?(?P<range>.+)", re.IGNORECASE | re.MULTILINE) #Works because the '.' by default does not match newline characters, so it stops once it gets to the end of the line.
Pattern = re.compile(r"Tag:?(.+)(?=Calib)", re.IGNORECASE | re.MULTILINE) #Works because the '.' by default does not match newline characters, 
# so it stops once it gets to the end of the line.

match = Pattern.findall(document_text) #When capturing groups are present, returns the group as a list.
if match: #has 'truthy' value if match found.
    range = match[0].strip() #Strips whitespace.
    print(range)
