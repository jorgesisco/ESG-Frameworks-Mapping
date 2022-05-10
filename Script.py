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

	def getTables(self, file_path, firstTablePage):
		
		pdf = read_pdf(file_path)
		frames = []

		for i in range(firstTablePage, len(pdf.pages)):
		    try:
		        page = pdf.pages[i]
		        table = page.extract_table()
		        frames.append(pd.DataFrame(table))
		    except:
		        pass


pdf_path = 'ESG-Frameworks/Mapping-Standards/sdg-gri.pdf'
firstTablePage = 3
table = ExtractPDFTables(pdf_path, firstTablePage)


# print(result.sorter(df))