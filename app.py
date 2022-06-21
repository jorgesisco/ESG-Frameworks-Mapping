from logging import PlaceHolder
import pandas as pd
import numpy as np
import base64
from turtle import onclick
import streamlit as st
import subprocess
from subprocess import STDOUT, check_call
import os
import base64
import camelot.io as camelot

from Script_2 import ExtractPDFTables

# to run this only once and it's cached
@st.cache
def gh():
    # Install ghostscript in machine
    proc = subprocess.Popen('apt-get install -y ghostscript', shell=True, stdin=None, stdout=open(os.devnull, 'wb'), stderr=STDOUT, executable="/bin/bash")
    proc.wait()

gh()

st.title("ESG PDF Table Extractor")
st.subheader("Script currently working only with ADX-GRI16 PDF")
st.text("Note: Only table extration from PDF words on this app, mapping on excel file will be\nimplemented later on")


input_pdf = st.file_uploader(label = "upload your pdf here", type = 'pdf')


get_codes_check = st.checkbox('Would you like to get columns with ESGs codes? (optional)')

if get_codes_check:
    st.markdown("### ESG #1")
    esg_1 = st.text_input("Enter ESG Name:", placeholder='e.g: GRI_2016, ADX, SDG', value='GRI_16')

    st.markdown("### ESG #2")
    esg_2 = st.text_input("Enter ESG Name:", placeholder=' e.g: GRI_2016, ADX', value='ADX')

st.markdown("### Page Number")
page_number = st.text_input("Enter the page or pages:", placeholder=' e.g: 1, 1-10', value='10-11')

st.markdown('### Flavor Approach')
flavor_ = st.text_input("Type 'lattice' for tables with visible gridlines or 'stream' for tables with not gridlines at all", placeholder='stream or lattice', value='lattice')


Extract_Tables_button = st.button(label='Extract Tables', key=None, help=None, on_click=None, args=None, kwargs=None, disabled=False)




if input_pdf != None and  Extract_Tables_button:

    try:
        # byte object into a PDF file 
        with open("input.pdf", "wb") as f:
            base64_pdf = base64.b64encode(input_pdf.read()).decode('utf-8')
            f.write(base64.b64decode(base64_pdf))
        f.close()

        # read the pdf and parse it using stream
        # table = camelot.read_pdf("input.pdf", pages = page_number, flavor = flavor_)
        if get_codes_check:
            df = ExtractPDFTables("input.pdf",
                                    page_range = page_number,
                                    flavor = flavor_,
                                    strip_text = '\n',
                                    get_codes=[esg_1, esg_2]).getTablesCamelot()
            
        else:
            df = ExtractPDFTables("input.pdf",
                                page_range = page_number,
                                flavor = flavor_,
                                strip_text = '\n').getTablesCamelot()
        
         
        if len(df) > 0:
            st.dataframe(df)

            def convert_df(df):
                return df.to_csv().encode('utf-8')

            csv = convert_df(df)

            st.download_button(
                 label="Download data as CSV",
                 data=csv,
                file_name='ESG_Table.csv',
                mime='text/csv',
            )

            st.markdown("### Map Data on Excel")
            map_to_excel = st.button(label='Map to excel', key=None, help=None, on_click=None, args=None, kwargs=None, disabled=False)
    except:
        pass