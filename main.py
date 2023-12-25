import PyPDF2
import re #RegEx
import openpyxl

# a function to write the U numbers to excel
def write_U_numbers_to_excel(roll_numbers):
    output_file='CDA3103_Attendance.xlsx'

    # Create a new Excel workbook
    workbook = openpyxl.Workbook()

    # Select the active sheet
    sheet = workbook.active

    # Write header
    sheet['A1'] = 'Roll Number'

    # Write roll number to the Excel sheet
    row_number = 2
    for num in roll_numbers:
        sheet.cell(row = row_number, column = 1, value = num)
        row_number += 1        

    # Save the workbook to the output file
    workbook.save(output_file)

# a function to find the list of U numbers
def find_U_numbers(text):
    #Define the regular expression pattern
    format = r'U\d+' # Match the character 'U' followed by one or more digits

    #Use re.findall to find all matches in the text
    matches = re.findall(format,text)

    return matches

conversion = 0
try:
    user_input = input("Enter PDF file name: ")
    file_name = user_input + ".pdf"

    pdfFileObject = open(file_name, "rb")


    pdf_reader = PyPDF2.PdfReader(pdfFileObject) # PdfReader object

    # Get the total number of pages
    num_pages  = len(pdf_reader.pages)

    output_text = ""

    for i in range(num_pages):
        page = pdf_reader.pages[i]
        output_text += page.extract_text()
                
    conversion = 1

except FileNotFoundError:
    print(f"Error: The file '{file_name}' was not found.")
    

if(conversion):
    print("File found")
    result = find_U_numbers(output_text)
    write_U_numbers_to_excel(result)
    print("Success! Check the excel file")
else:
    print("Error")




