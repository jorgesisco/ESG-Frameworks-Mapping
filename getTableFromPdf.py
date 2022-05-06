from tabula import read_pdf
# from tabulate import tabulate
import pandas as pd
import numpy as np


def getTableSDG_GRI(path, page, output_path):
    df = read_pdf(path, stream=True, pages = page, area = [80.51, 90.42, 561.96, 814.18], multiple_tables=False)

    for i in range(page+1, 73):
        df[0] = df[0].append(read_pdf(path, stream=True, pages = i, area = [80.51, 90.42, 561.96, 814.18], multiple_tables=False )[0], ignore_index=True)
    
    df[0].rename(columns = {'Sources':'Source'}, inplace = True)

    for i in range(74, 99):
        df[0] = df[0].append(read_pdf(path, stream=True, pages = i, area = [80.51, 90.42, 561.96, 814.18], multiple_tables=False )[0], ignore_index=True)
    
    df[0].to_csv(output_path)

    clean_table(output_path)
    return None


def clean_table(output_path):
    df = pd.read_csv(output_path)
    df = df.drop(['Unnamed: 0'], axis=1)
    df = df.replace(np.nan, '', regex=True)
    df['Target_'] = df['Target']
    df = df[["Target", "Target_", "Available Business Disclosures", "Disclosure"]]
    df = df.drop(labels = 'Target_',axis = 1).groupby(df['Target_'].mask(df['Target_']=='').ffill()).agg('\n'.join).reset_index()
    df.rename(columns = {'Target_':'Target'}, inplace = True)
    df.to_csv(output_path)
    

def Mapping(file, df):
    return None


if __name__ == "__main__":
    getTableSDG_GRI()
