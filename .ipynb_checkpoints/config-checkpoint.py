import camelot
import pandas as pd
#Frameworks Linking Standars

SDG_GRI = camelot.read_pdf('ESG-Frameworks/Mapping-Standars/sdg-gri.pdf', flavor="stream", edge_tol=500, pages='3')
GRID_COH4B = camelot.read_pdf('ESG-Frameworks/Mapping-Standars/gri-coh4b.pdf')

def load_files(file):

    df = pd.Datafram
    print(file[0].df)


if __name__ == "__main__":
    print(load_files(SDG_GRI))