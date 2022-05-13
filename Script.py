import pandas as pd
import numpy as np
import re

import pdfplumber
from tabula import read_pdf
import openpyxl


class ExtractPDFTables:

	def __init__(self, file_path, page_range=False, firstTablePage=False, area=False):
		self.file_path = file_path
		self.page_range = page_range
		self.firstTablePage = firstTablePage
		self.area = area

	def getTables1(self):
		
		pdf = pdfplumber.open(self.file_path)
		frames = []

		for i in range(self.firstTablePage, len(pdf.pages)):
			try:
				page = pdf.pages[i]
				table = page.extract_table()
				frames.append(pd.DataFrame(table))
			
			except:
				pass

		df =  pd.concat(frames)
		df = df.drop_duplicates()

		return df

	def getTables2(self):
		# pdf = read_pdf(self.file_path, pages=self.page_range, area = self.area, multiple_tables=True)
		pdf = pdfplumber.open(self.file_path, pages=self.page_range)

		'''Ojo con esto: Code below is based on just one page
		   A soon as I manage to extract the desired data
		   I gotta test with last pages.
		'''
		page = pdf.pages[0]
		table = page.debug_tablefinder()
		req_table = table.tables[0] 

		cells = req_table.cells

		for cell in cells[0:len(cells)]:
		    page.crop(cell).extract_words() 

		data = page.extract_table()

		df = pd.DataFrame(data[2:],columns=data[1])

		df.iloc[:, 0] = df.iloc[:, 0].str.replace('\n' , ' ')
		df.iloc[:, 0] = df.iloc[:, 0].str.replace('Governance' , '')
		df.iloc[:, 1] = df.iloc[:, 1].str.replace('\n' , ' ')
		return df
		

	def setHeaders(self, df, rowIndex):
		headers = df.iloc[rowIndex]
		df = pd.DataFrame(df.values[rowIndex+1:], columns=headers)

		return df

	def headerSwap(self, df, h1, h2, newh1):

		df.rename(columns = {h1:newh1, h2:h1}, inplace = True)
        
		return df

	def addDot(self, df, column):
		df[column] = df[column].str[:-1] + df[column].str[-1] + '.'

		return df






