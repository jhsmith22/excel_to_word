#!/usr/bin/env python
# coding: utf-8

# In[116]:


# coding: utf-8
# citations: https://pythonmana.com/2021/03/20210329161147051K.html
# citation: https://stackoverflow.com/questions/43637211/retrieve-document-content-with-document-structure-with-python-docx

# Import Libraries
get_ipython().system('pip install docx')
# !pip install logging
# !pip install re
# !pip install os
# !pip install xlrd
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import pandas as pd
import numpy as np
import re
from os.path import exists
import logging 


# In[117]:


########################################################################
# Load Project and Table Settings 
########################################################################
# Master File
master = 'Control.xlsx' 

# Project Settings
project_settings = pd.read_excel(master, sheet_name='project settings', header=None)

# Table Settings
table_settings = pd.read_excel(master, sheet_name='table settings')

# Create a caption list that stores all the parsed captions for all tables
caption_list_parsed = table_settings['word_table_caption_text'].tolist()
print(f"Captions: {caption_list_parsed}")

# Word Table Shells INPUT Filename
document = Document(project_settings.iloc[0,1].strip())

# Word Table Shells OUTPUT Filename
output_doc = project_settings.iloc[1,1].strip()
if not output_doc:
    output_doc = "results.docx"
print(f"Word Table Shells OUTPUT Filename: {output_doc}")

# Table Format Style
table_style = project_settings.iloc[2,1].strip()
if not table_style:
    table_style = "__Table Style-AIR 2021"
print(f"All tables in Word will be in this style: {table_style}")

# Table Caption Style
caption_style = project_settings.iloc[3,1].strip()
if not caption_style:
    caption_style = "Exhibit Title"
print(f"All table captions in Word have this style: {caption_style}")

print("DONE LOADING PROJECT AND TABLE SETTINGS")


# In[118]:


########################################################################
# Populate Table Data in Word
########################################################################

# Loop through the paragraphs & table pairs in the Word document
# citation: source: https://theprogrammingexpert.com/write-table-fast-python-docx/
''' Define script to identify table 'child' within paragraph 'parent' based on document order
    Each returned value is an instance of either Table or Paragraph.'''
def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)      

'''Identifies table meta-data from crosswalk file'''
def findtables(document, table_settings):

    # Iterate through paragraphs
    for block in iter_block_items(document):

        # Identify paragraphs
        if isinstance(block, Paragraph):
            
            # Find the element in the caption_list_parsed that matches
            if block.style.name==caption_style:
                
                print(f"Table Caption extracted Word Table Shells: {block.text}")

                for caption in caption_list_parsed:
                    result = re.search(caption, block.text)
                    if result is not None:
                        row = caption_list_parsed.index(caption)
                                         
                # Load corresponding table data
                worksheet_name = table_settings.iloc[row, 1].strip()
                print(f"Worksheet Name extracted from Control file: {worksheet_name}")

                workbook_name = table_settings.iloc[row, 2].strip()
                print(f"Workbook Name extracted from Control file: {workbook_name}")
                
                tables_data = pd.read_excel(
                    workbook_name, sheet_name=worksheet_name, header=None)
                print(f"Table Data extracted from worksheet: {tables_data}")
                
                # Load "nofill_first_x_rows" setting
                skip_table_rows = table_settings.iloc[row, 3]
                if not skip_table_rows:
                    skip_table_rows = 0

                # Load "nofill_first_y_cols" setting
                skip_table_cols = table_settings.iloc[row, 4]
                if not skip_table_cols:
                    skip_table_cols = 0
                
                # Load "skip_merged_rows" setting
                skip_merged_rows = table_settings.iloc[row, 5].strip()
                if not skip_merged_rows:
                    skip_merged_rows = "y"
                
                print(f"Done loading table settings for: {result}")
                
                # Convert decimal to percent
                format_as_percent = table_settings.iloc[row, 6]
                print(format_as_percent)
                
                # Number of decimals
                percent_decimal_places = table_settings.iloc[row, 7]
                print(percent_decimal_places)
                
                # Move on to the corresponding table object
                continue

        # Load corresponding table object
        elif isinstance(block,Table):
            
            tablepopulate(block, tables_data, skip_table_rows, skip_table_cols, 
                          skip_merged_rows, worksheet_name, format_as_percent, percent_decimal_places)

'''Extracts that table's formatting specifications and fills in the data'''
def tablepopulate(block, df, skip_table_rows, skip_table_cols, 
                  skip_merged_rows, worksheet_name, format_as_percent, percent_decimal_places):
     
        # Apply decimal to percent formatting conversion
#         df.round(mrate * 100, 1)
#         mratepct = str(mrate_rnd) + '%'
        
#         if format_as_percent == "y":
#             for (columnName, columnData) in df.iteritems():
#                 df.round(columnName * 100, percent_decimal_places)
        
        # loop through rows and cols of the dataframe to populate the table object
        for i in range(skip_table_rows, df.shape[0]):
            for j in range(skip_table_cols, df.shape[1]):
                
                # Skip merged cells in the first column
                if j == 0 and skip_merged_rows == 'y':
                    c = block.cell(i,0)

                    if c._tc.right > 1:
                        continue
                                      
                # Skip over blank (nan) cells in the dataframe
                if (str(df.values[i,j])) != 'nan':
                    block.cell(i,j).text = str(df.values[i,j])
        
        # Add table styles and formats                        
        block.style = table_style
        
        print(f"Done populating data from: {worksheet_name}")

findtables(document, table_settings)

# #3. Save the outputted Word document
document.save('result.docx')
print(f"Program complete. Output file is ready to open: {output_doc}")


# In[ ]:





# In[ ]:




