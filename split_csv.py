# THIS MODULE
# take csv file ,
# split frame according to desired,
# save each partition in excel,
# take 10 min for 650k excel file




import pandas as pd  
import os



def partion_file(file, limit, expected_files_count):
	""" A FUNCTION TO PARTION EXCEL FILE INTO FILES WITH SPECIFIED LIMITS AND EXPECTED NUMBER OF FILES 

		FUNCTION PARAMETERS :
		limit : is required no of rows in file
		expected_files_count : is no of files that will be generation by this partioning


		EXAMPLE :
		partion_file("all.csv", 999, 35)
	"""

	# function defaults variables
	sheetname= "Sheet1"
	output_dir = r"D:\Desktop\py\EXCEL_UTILS\output\c2\\"

	# read input file
	# dfs = pd.read_excel(file, sheet_name=sheetname)
	dfs = pd.read_csv(file)


	# set python to output_dir
	os.chdir(output_dir)

	# init loop
	start = 0
	end = limit
	expected_files_count += 1	
	for i in range(1, expected_files_count):
		df = dfs[start:end]
		df.to_excel(str(i+11) + ".xlsx", index=False)
		# df.to_csv(str(i) + ".csv", sep='\t')	
		start += limit + 1	# increment row that included in end . 
		end += limit

partion_file("lead1.csv", 99999, 11)