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
st.subheader("Developed by Jorge Sisco")

input_pdf = st.file_uploader(label = "upload your pdf here", type = 'pdf')

st.markdown("### Page Number")
page_number = st.text_input("Enter the page or pages:", placeholder='e.g: 1, 1-10')

st.markdown('### Flavor Approach')
flavor_ = st.text_input("Type 'lattice' for tables with visible gridlines or 'stream' for tables with not gridlines at all", value='lattice')


Extract_Tables_button = st.button(label='Extract Tables', key=None, help=None, on_click=None, args=None, kwargs=None, disabled=False)

if input_pdf != None and Extract_Tables_button:

    

    try:
        # byte object into a PDF file 
        with open("input.pdf", "wb") as f:
            base64_pdf = base64.b64encode(input_pdf.read()).decode('utf-8')
            f.write(base64.b64decode(base64_pdf))
        f.close()

        # read the pdf and parse it using stream
        # table = camelot.read_pdf("input.pdf", pages = page_number, flavor = flavor_)
        table = ExtractPDFTables("input.pdf",
                                page_range = page_number,
                                flavor = flavor_,
                                strip_text = '\n').getTablesCamelot()
            

        if len(table) > 0:

            st.dataframe(table)

            def convert_df(table):
                return table.to_csv().encode('utf-8')

            csv = convert_df(table)

            st.download_button(
                 label="Download data as CSV",
                 data=csv,
                file_name='ESG_Table.csv',
                mime='text/csv',
            )

            map_to_excel = st.button(label='Map to excel', key=None, help=None, on_click=None, args=None, kwargs=None, disabled=False)
    except:
        pass