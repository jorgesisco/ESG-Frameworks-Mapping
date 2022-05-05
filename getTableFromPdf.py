from tabula import read_pdf
# from tabulate import tabulate
import pandas as pd
import numpy as np


def getTable(path, page):
    df = read_pdf(path, stream=True, pages = page, area = [80.51, 90.42, 561.96, 814.18], multiple_tables=False)

    for i in range(page+1, 73):
        df[0] = df[0].append(read_pdf('ESG-Frameworks/Mapping-Standars/sdg-gri.pdf', stream=True, pages = i, area = [80.51, 90.42, 561.96, 814.18], multiple_tables=False )[0], ignore_index=True)
    
    df[0].rename(columns = {'Sources':'Source'}, inplace = True)

    for i in range(74, 99):
        df[0] = df[0].append(read_pdf('ESG-Frameworks/Mapping-Standars/sdg-gri.pdf', stream=True, pages = i, area = [80.51, 90.42, 561.96, 814.18], multiple_tables=False )[0], ignore_index=True)
    
    df[0].to_csv('SDG-GRI-DF.csv')

    clean_table()
    return None


def clean_table():
    df = pd.read_csv('SDG-GRI-DF.csv')
    df = df.drop(['Unnamed: 0'], axis=1)
    df = df.replace(np.nan, '', regex=True)
    df = df.replace('', np.nan).ffill().bfill()
    df = df.groupby(['Target', 'Disclosure'])['Available Business Disclosures'].apply(', \n'.join).reset_index()
    df = df[["Target", "Available Business Disclosures", "Disclosure"]]
    df.to_csv('SDG_GRI.csv')
    




if __name__ == "__main__":
    getTable()
