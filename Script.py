import pandas as pd
import numpy as np
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
	def getTablesSDG_GRI(self):

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

			structuringApproach = input('Are you going to map SDG to GRI? (yes or no): ')

			if structuringApproach == 'yes':
				print('Structured SDG to GRI, this data can be mapped on SDG excel sheet with the GRI Disclosures')
				df = df.dropna()
				df = df.drop(['Available Business Disclosures'], axis=1)
				df = df.drop(labels = 'Target',axis = 1).groupby(df['Target'].mask(df['Target']==' ').ffill()).agg(', '.join).reset_index()

			elif structuringApproach == 'no':
				print('Structured GRI to SDG, this data can be mapped on GRI excel sheet with the SDG Tagets')
				df = df.drop(['Available Business Disclosures'], axis=1)
				df = df.dropna()

			else:
				print("WARNINH: please run, the script again and choose 'yes' or 'no' ") 

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

	def extractDisclosures(self, df, column, newColumn):
		
		values = []

		for i in range(0, len(df[column])):

			try:
				match = re.search(r'[0-9]{3}-[0-9]{2}|[0-9]{3}-[0-9]{1}', df[column][i])
				values.append(df[column][i][match.start():match.end()])

			except:
				values.append('No value :(')
				pass
		df[newColumn] = values

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
			if ws.cell(row=i, column=1).value == None:
				pass

			else:
				target_cell = ws.cell(row=i, column=1).value

				if (re.search(regex, target_cell)):
					target = re.search(regex, target_cell).group()

					try:
						value_to_add = self.df.loc[self.df['Target'] == target]['Disclosure'].item()
						ws.cell(row=i, column=3, value=str(value_to_add))	
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
						if len(self.df[self.df['Disclosure'] == target_cell]) != 0:
							value_to_add = self.df[self.df.Disclosure==target_cell].squeeze()['Target'].values
							ws.cell(row=i+1, column=4, value=', '.join(value_to_add))
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
						ws.cell(row=i, column=6, value=value_to_add[0])

					except:
						pass

		wb.save(self.path_file)

		return f'{self.sheet} sheet from Excel file have bee mapped' 

	def mapTCFD_GRI(self):
		pass

	def mapGRI_TCFD(self):
		pass