import getTableFromPdf
from getTableFromPdf import getTable, Mapping



path = 'ESG-Frameworks/Mapping-Standards/sdg-gri.pdf'
page = 3

output_path = 'ESG-Frameworks/Mapping-Standards/SDG-GRI-DF.csv'

# getTableSDG_GRI(path, page, output_path)

excelFile = 'ESG-Frameworks/Outputs/ESG Alignment Product Outputs.xlsx'

Mapping(excelFile, output_path)