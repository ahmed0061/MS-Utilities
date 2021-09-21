import os, sys, openpyxl


# vars
directory = r"C:\Users\ahmed.mosaad\Desktop\aaa"    # directory where excel files
TargetPartNumber="74LVT14D,112"                     # string need to search for 
col_no = 2


# change working dir to where files exists
os.chdir(directory)


# generate excel files in dir
xl_files = ( file for file in os.listdir(r"C:\Users\ahmed.mosaad\Desktop\aaa") if file.endswith(".xlsx")       )



# function to get file related objs
def file_objects(file_name):
    file_obj = openpyxl.load_workbook(file_name)
    file_sheet = file_obj.get_sheet_by_name("Sheet1")
    file_rows = file_sheet.get_highest_row()

    return file_obj, file_sheet, file_rows


# loop through each file,   try get file objs,  loop through specific column where expecting the string 
for file in list(xl_files):

    try:
        file_obj, file_sheet, file_rows = file_objects(file)
        
        for i in range(1, file_rows+1):
            if file_sheet.cell(row=i, column=col_no).value == TargetPartNumber : 
                print("Found in : ", file)
    
    except:
        pass