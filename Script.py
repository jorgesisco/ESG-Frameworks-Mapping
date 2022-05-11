import pandas as pd
import numpy as np
import re

import pdfplumber
from tabula import read_pdf
import openpyxl


class ExtractPDFTables:

	def __init__(self, file_path, firstTablePage):
		self.file_path = file_path
		self.firstTablePage = firstTablePage

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

	# def getTables2(self):
	# 	pdf = read_pdf(self.file_path)
		

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






