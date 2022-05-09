import getTableFromPdf
from getTableFromPdf import getTableSDG_GRI



path = 'ESG-Frameworks/Mapping-Standards/sdg-gri.pdf'
page = 3

extracted_tables = 'ESG-Frameworks/Mapping-Standards/SDG-GRI.csv'

table_structured = 'ESG-Frameworks/Mapping-Standards/SDG_GRI.csv'

getTableSDG_GRI(path, page, extracted_tables, table_structured )

excelFile = 'ESG-Frameworks/Outputs/ESG Alignment Product Outputs.xlsx'

# Mapping(excelFile, output_path)