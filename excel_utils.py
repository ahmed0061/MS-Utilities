import pandas as pd  
import os


def partion_file(file, limit, expected_files_count):
	""" A FUNCTION TO PARTION EXCEL FILE INTO FILES WITH SPECIFIED LIMITS AND EXPECTED NUMBER OF FILES 

		FUNCTION PARAMETERS :
		limit : is required no of rows in file
		expected_files_count : is no of files that will be generation by this partioning


		EXAMPLE :
		partion_file("all.xlsx", 999, 35)
	"""

	# function defaults variables
	sheetname= "Sheet1"
	output_dir = r"D:\Desktop\py\EXCEL_UTILS\output\Stocked\\"

	# read input file
	dfs = pd.read_excel(file, sheet_name=sheetname)


	# set python to output_dir
	os.chdir(output_dir)

	# init loop
	start = 0
	end = limit
	expected_files_count += 1	
	for i in range(1, expected_files_count):
		df = dfs[start:end]
		df.to_excel(str(i) + ".xlsx", index=False)	
		start += limit + 1	# increment row that included in end . 
		end += limit

partion_file("stocked-imp.xlsx", 99999, 8)


# def combine_files():
# 	# variables
# 	directory = r"C:\Users\ahmed.mosaad\Desktop\py\EXCEL_UTILS\Combine\\"
# 	extension = ".xls"
	
# 	# set work directory
# 	os.chdir(directory)
# 	files = [file for file in os.listdir(directory) if file.endswith(extension)]
# 	df = pd.DataFrame()
# 	for file in files:
# 	     df = df.append(pd.read_excel(file), ignore_index=True)
# 	df.to_excel("Combined" + ".xlsx", index=False) 


# # combine_files()




