import pandas as pd
# import numpy as np
import re

import pdfplumber
from tabula import read_pdf
import openpyxl


class ExtractPDFTables:

	# Initializing contructor variables
	def __init__(self, file_path, page_range=False, area=False):
		self.file_path = file_path #string with file path
		self.page_range = page_range # defined page range with tables
		self.area = area # Table area

	# Getting tables from PDF SDG Linked to GRI
	def getTablesSDG_GRI_0_Deprecated(self):

	# 	'''This method requires to add a nested list with 2 page ranges,
	# 	   with this approach, we avoid scanning a non table on page 73
	# 	'''
	# 	if any(isinstance(i, list) for i in self.page_range) == False:
	# 		return print("WARNING: page_range should be a nested list\n eg: page_range = [list(range(3, 73)), list(range(74, 99))]\n Why? the SDG-GRI pdf file has a page in the middle that is\n not a table, getTablesSDG_GRI method iterates the first\n range of pages an then the other one.")

	# 	else:
	# 		# Gets all the tables from first page range
	# 		pdf = read_pdf(self.file_path, stream=True, pages = self.page_range[0],
	# 					   area = self.area, multiple_tables=False)

	# 		#renaming "sources" column to "source" (same column name in the reamining tables)
	# 		pdf[0].rename(columns = {'Sources':'Source'}, inplace = True)
			
	# 		# Getting the remaining tables
	# 		pdf_ = read_pdf(self.file_path, stream=True, pages = self.page_range[1],
	# 						area = self.area, multiple_tables=False )

	# 		#Merging all the tables in one pandas dataframe
	# 		pdf[0] = pdf[0].append(pdf_[0])

	# 		# Renaming variable name
	# 		df = pdf[0]
	# 		df = df[df.Target != 'Target']

	# 		structuringApproach = input('Are you going to map GRI on SDG sheet? (yes or no): ')

	# 		if structuringApproach == 'yes':
	# 			print('Structured GRI for SDG sheet, this data can be mapped on SDG excel sheet with the GRI Disclosures')
	# 			df = df.dropna()
	# 			# df = df.drop(['Available Business Disclosures'], axis=1)
	# 			df = df.drop(labels = 'Target',axis = 1).groupby(df['Target'].mask(df['Target']==' ').ffill()).agg(', '.join).reset_index()

	# 		elif structuringApproach == 'no':
	# 			print('Structured SDG to GRI, this data can be mapped on GRI excel sheet with the SDG Targets')
	# 			df = df.drop(['Available Business Disclosures'], axis=1)
	# 			df = df.dropna()

	# 		else:
	# 			print("WARNING: please run the script again and choose 'yes' or 'no' ") 

			return df

	def getTablesSDG_GRI(self, sdg=None):

		if sdg != None:
					sdg_df = pd.read_csv(sdg)
		
		'''This method requires to add a nested list with 2 page ranges,
		   with this approach, we avoid scanning a non table on page 73
		'''
		if any(isinstance(i, list) for i in self.page_range) == False:
			return print("WARNING: page_range should be a nested list\n eg: page_range = [list(range(3, 73)), list(range(74, 99))]\n Why? the SDG-GRI pdf file has a page in the middle that is\n not a table, getTablesSDG_GRI method iterates the first\n range of pages an then the other one.")

		else:
			# Gets all the tables from first page range
			pdf = read_pdf(self.file_path, stream=True, pages = self.page_range[0],
						   area = self.area, multiple_tables=False)

			#renaming "sources" column to "source" (same column name in the reamining tables)
			pdf[0].rename(columns = {'Sources':'Source'}, inplace = True)
			
			# Getting the remaining tables
			pdf_ = read_pdf(self.file_path, stream=True, pages = self.page_range[1],
							area = self.area, multiple_tables=False )

			#Merging all the tables in one pandas dataframe
			pdf[0] = pdf[0].append(pdf_[0])

			# Renaming variable name
			df = pdf[0]
			df = df[df.Target != 'Target']

			
			array_agg = lambda x: '\n'.join(x.astype(str))

			df = df.fillna(method='ffill')
			df = df.groupby(['Target', 'Disclosure'], as_index=False).agg({'Available Business Disclosures': array_agg})

			structuringApproach = input('Are you going to map GRI on SDG sheet? (yes or no): ')

			if structuringApproach == 'yes':
				print('Structured GRI for SDG sheet, this data can be mapped on SDG excel sheet with the GRI Disclosures')

				df = df.groupby(['Target'], as_index=False).agg({'Available Business Disclosures': array_agg, 'Disclosure':array_agg})
				# df = df.drop(labels = 'Target',axis = 1).groupby(df['Target'].mask(df['Target']==' ').ffill()).agg(', '.join).reset_index()

				df.rename(columns = {'Target':'SDG_Target', 'Available Business Disclosures': 'GRI Available Business Disclosures', 
									 'Disclosure': 'GRI_Disclosure'}, inplace = True)
			
				if sdg != None:
					df = pd.merge(df, sdg_df, on=['SDG_Target'], how='right')
					df = df[['SDG_Target', 'SDG Description','GRI_Disclosure', 'GRI Available Business Disclosures']]
				
				# df = df.dropna()

			elif structuringApproach == 'no':

				df.rename(columns = {'Target':'SDG_Target', 'Available Business Disclosures': 'GRI Available Business Disclosures', 
									 'Disclosure': 'GRI_Disclosure'}, inplace = True)
				
				if sdg != None:
					df = pd.merge(df, sdg_df, on=['SDG_Target'], how='right')
					df = df[['SDG_Target', 'SDG Description','GRI_Disclosure', 'GRI Available Business Disclosures']]
				

			else:
				print("WARNING: please run the script again and choose 'yes' or 'no' ") 

			return df


	# Getting tables from PDF GRI Linked to COH4B
	def getTablesGRI_COH4B(self):
		
		pdf = pdfplumber.open(self.file_path, pages=self.page_range)
		frames = []

		for i in range(0, len(self.page_range)):
			try:
				page = pdf.pages[i]
				table = page.extract_table()
				frames.append(pd.DataFrame(table))
			
			except:
				pass

		df =  pd.concat(frames)
		df = df.drop_duplicates()

		return df
	
	# Getting tables from PDF TCFD Linked to GRI
	def getTablesTCFD_GRI(self):
		# pdf = read_pdf(self.file_path, pages=self.page_range, area = self.area, multiple_tables=True)
		pdf = pdfplumber.open(self.file_path, pages=self.page_range)

		frames = []

		for i in range(0, len(self.page_range)):

			page = pdf.pages[i]
			table = page.debug_tablefinder()
			req_table = table.tables[0] 

			cells = req_table.cells

			for cell in cells[0:len(cells)]:
			    page.crop(cell).extract_words()
			
			data = page.extract_table()
			df = pd.DataFrame(data)
			df = pd.DataFrame(data[2:],columns=data[1])
			df.iloc[:, 0] = df.iloc[:, 0].str.replace('\n' , ' ')
			# df.iloc[:, 0] = df.iloc[:, 0].str.replace('Governance' , '')
			df.iloc[:, 1] = df.iloc[:, 1].str.replace('\n' , ' ')
			frames.append(df)


		df =  pd.concat(frames)
		df = df.drop_duplicates()

		return df

	# Funtions needed in some of the extracted dataframes
	def extractDisclosures(self, df, column, newColumn, regex, method):
		
		values = []

		for i in range(0, len(df[column])):

			try:
				match = method(regex, df[column][i])
				values.append(df[column][i][match.start():match.end()])

			except:
				values.append('No value :(')
				pass
		df[newColumn] = values

		return df

	def extractDisclosures_DEPRECATERED(self, df, column, newColumn, regex, method_):
		
		values = []

		for i in range(0, len(df[column])):

			try:
				match = method_(regex, df[column].values[i])

				if len(match) > 2:
					values.append(', '.join(match))
				else:
					values.append(''.join(match[0]))

			except:
				values.append('No value :(')
				pass
		df[newColumn] = values

		return df

	def df_gri_tcdf(self, df):

		df_GRI_TCDF = df[['GRI Standards', 'id', 'Recommended \nDisclosures \n(TCFD Framework)']]
		df_GRI_TCDF.rename(columns = {'GRI Standards':'GRI_Standards'}, inplace = True)
		df_GRI_TCDF['GRI_Standards'].str.split(', ')
		df_GRI_TCDF = df_GRI_TCDF.assign(GRI_Standards=df_GRI_TCDF['GRI_Standards'].str.split(', ')).explode('GRI_Standards')

		return df_GRI_TCDF

	'''This methods are usefull in case the gathered table
	   does now shows the columns as headers, but columns
	   are shown as values, we run the method with the
	   gathered dataframe and the row index we want as header. 
	'''
	def setHeaders(self, df, rowIndex):
		headers = df.iloc[rowIndex]
		df = pd.DataFrame(df.values[rowIndex+1:], columns=headers)

		return df

	'''There is a header with None value and the correct value is
	   in the column to the left when choosing that one header 'row 1 header', 
	   the None header as None and the value from column 1 that is suppose to 
	   be an id,
	'''
	def headerSwap(self, df, h1, h2, newh1):

		df.rename(columns = {h1:newh1, h2:h1}, inplace = True)
        
		return df

	''' It is convinient to add a . to the id numbers, to easily match
		with the same values on the excel files and do the mapping.'''
	def addDot(self, df, column):
		df[column] = df[column].str[:-1] + df[column].str[-1] + '.'

		return df

		

class MapLinks2Excel:

	def __init__(self, df, sheet, path_file):
		self.df = df
		self.sheet = sheet
		self.path_file = path_file
	
	def MapSDG_GRI(self):
		regex = '[+-]?[0-9]+\.-?[0-9a-zA-Z_]+'

		wb = openpyxl.load_workbook(self.path_file)
		ws = wb[self.sheet]

		rows = ws.max_row

		for i in range(1, rows):
			if ws.cell(row=i, column=1).value != None:
				target_cell = ws.cell(row=i, column=1).value
				

				if (re.search(regex, target_cell)):
					target = re.search(regex, target_cell).group()

					try:
						disclosure_to_add = self.df.loc[self.df['SDG_Target'] == target]['GRI_Disclosure'].item()
						description_to_add = self.df.loc[self.df['SDG_Target'] == target]['GRI Available Business Disclosures'].item()
						
						ws.cell(row=i, column=3, value=str(disclosure_to_add))	
						ws.cell(row=i, column=4, value=str(description_to_add))
					except:
						pass

		wb.save(self.path_file)

		return f'{self.sheet} sheet from Excel file have bee mapped'
	
	def MapGRI_SDG(self):

		wb = openpyxl.load_workbook(self.path_file)
		ws = wb[self.sheet]

		rows = ws.max_row

		for i in range(1, rows):

			if ws.cell(row=i, column=2).value != None:
				target_cell = ws.cell(row=i+1, column=2).value

				if target_cell != None:
					try:
						if len(self.df[self.df['GRI_Disclosure'] == target_cell]) != 0:
							target_to_add = self.df[self.df.GRI_Disclosure==target_cell].squeeze()['SDG_Target'].values
							ws.cell(row=i+1, column=6, value='\n '.join(target_to_add))

							disclosure_to_add = self.df[self.df.GRI_Disclosure==target_cell].squeeze()['SDG Description'].values
							ws.cell(row=i+1, column=7, value='\n '.join(disclosure_to_add))
					except:
						pass

		wb.save(self.path_file)

		return f'{self.sheet} sheet from Excel file have bee mapped'

	def MapCOH4B_GRI(self):
		regex = "[+-]?[0-9]+\."

		wb = openpyxl.load_workbook(self.path_file)
		ws = wb[self.sheet]
		rows = ws.max_row

		''' In case of loading df from csv file, the commented code below
			will fixe and issue with the id column, since the df it's being taking
			from instance tableGRI_COH4B, it is no needed to excute that code.
		'''

		# pdf_tables = pd.read_csv(pdf_tables)
		# pdf_tables['id'] = pdf_tables['id'].astype(str).apply(lambda x: x.replace('.0','.'))

		for i in range(1, rows):
			if ws.cell(row=i, column=1).value == None:
				pass

			else:
				target_cell = ws.cell(row=i, column=1).value

				if (re.search(regex, target_cell)):
					target = re.search(regex, target_cell).group()

					try:
						value_to_add = [self.df.loc[self.df['id'] == target]['C. GRI \nStandards'].item(),
										self.df.loc[self.df['id'] == target]['D. GRI disclosures'].item()]
						
						ws.cell(row=i, column=3, value=str(value_to_add[0]))
						ws.cell(row=i, column=4, value=str(value_to_add[1]))

					except:
						pass

		wb.save(self.path_file)

		return f'{self.sheet} sheet from Excel file have bee mapped'

	def MapGRI_COH4B(self):

		regex = '^[0-9\.\-\/]+$'

		wb = openpyxl.load_workbook(self.path_file)
		ws = wb[self.sheet]
		rows = ws.max_row

		for i in range(1, rows):
			if ws.cell(row=i, column=2).value == None:
				pass

			else:
				target_cell = ws.cell(row=i, column=2).value
				if(re.search(regex, target_cell)):
					target = re.search(regex, target_cell).group()

					try:
						value_to_add = self.df.loc[self.df['GRI Standards'] == target]['id'].item()
						# print(value_to_add)
						ws.cell(row=i, column=10, value=value_to_add)

						value_to_add_2 = self.df.loc[self.df['GRI Standards'] == target]['A. COHBP & \ndefinition'].item()
						ws.cell(row=i, column=11, value=value_to_add_2)
					except:
						pass

		wb.save(self.path_file)

		return f'{self.sheet} sheet from Excel file have bee mapped' 

	def mapTCFD_GRI(self):

		regex = '[a-zA-Z]+\)'
		wb = openpyxl.load_workbook(self.path_file)
		ws = wb[self.sheet]
		rows = ws.max_row

		count = 0

		for i in range(1, rows):
			if ws.cell(row=i, column=2).value != None and count <= len(self.df):
				target_cell = ws.cell(row=i, column=2).value

				if re.search(regex, target_cell):

					target = re.search(regex, target_cell).group()
					target = re.sub('[abc]\)', self.df['id'].values[count], target)

					value_to_add = self.df.loc[self.df['id'] == target]['Related \ncode/\nparagraph'].item()
					value_to_add = re.sub(', ', ' ', value_to_add)
					value_to_add = re.sub('\s\s+', ' ', value_to_add)
					value_to_add = re.sub(' and| with', '', value_to_add)

					value_to_add_2 = self.df.loc[self.df['id'] == target]['Description'].item()

					ws.cell(row=i, column=4, value=value_to_add)

					ws.cell(row=i, column=5, value=value_to_add_2)
					
					count += 1

		wb.save(self.path_file)

	def mapGRI_TCFD(self):

		wb = openpyxl.load_workbook(self.path_file)
		ws = wb[self.sheet]
		rows = ws.max_row

		for i in range(1, rows+1):
			if ws.cell(row=i, column=2).value != None:
				target_cell = ws.cell(row=i, column=2).value

				if target_cell != None:
					value = self.df.loc[self.df['GRI_Standards'] == target_cell]['id'].values
					value_2 =self.df.loc[self.df['GRI_Standards'] == target_cell]['Recommended \nDisclosures \n(TCFD Framework)'].values
					if len(value):
						value_to_add = ' '.join(value)
						ws.cell(row=i, column=8, value=value_to_add)
						# print(value_to_add)

				
					if len(value_2):
						value_to_add_2 = ' \n'.join(value_2)

						ws.cell(row=i, column=9, value=value_to_add_2)

						# print(value_to_add_2)
						# print('--------------------------------')

		

		wb.save(self.path_file)		

	def mapGRI2016_2021(self):
		wb = openpyxl.load_workbook(self.path_file)
		ws = wb[self.sheet]
		rows = ws.max_row

		for i in range(3, rows):
			if ws.cell(row=i, column=2).value != None:
				target_cell = ws.cell(row=i, column=2).value
        
        
				if target_cell != None:
					try:
						disclosure_to_add = self.df.loc[self.df['Disclosure Number 2016'] == target_cell]['Disclosure Number 2021'].item()
						ws.cell(row=i, column=4, value=disclosure_to_add)
						section_to_add = self.df.loc[self.df['Disclosure Number 2016'] == target_cell]['Section 2021'].item()
						ws.cell(row=i, column=5, value=section_to_add)	

					except:
						pass

		wb.save(self.path_file)

		print(f"{self.sheet} sheet from Excel file have bee mapped with it's GRI 2021 equivalent")

	def mapGRI2021_2016(self):
		pass