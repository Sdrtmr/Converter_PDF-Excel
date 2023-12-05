# installing tabula

 # pip install tabula-py

 # import pandas and convert

import tabula 
import pandas as pd

pdf_in = "C:/Users/user/Documents/Trofimova-2013.pdf"

PDF = tabula.read_pdf(pdf_in, pages = 'all', multiple_tables=True) 

# view result

print ('\nTables from PDF file\n'+str(PDF))

pdf_out_xlsx = "D:\Temp\From_PDF.xlsx"

# converting to Excel 

PDF = pd.DataFrame(PDF)
PDF.to_excel(pdf_out_xlsx,index=False) 

print("Done") 
