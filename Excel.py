import pandas as pd
import numpy as np

def auto_adjust_column_widths(writer, dataframe, sheet_name):
    # Code from: https://stackoverflow.com/questions/45985358/how-to-wrap-text-for-an-entire-column-using-pandas
    workbook  = writer.book    
    wrap_format = workbook.add_format({
        'text_wrap': True,
        'align': 'center'
    })

    # Code from: https://towardsdatascience.com/how-to-auto-adjust-the-width-of-excel-columns-with-pandas-excelwriter-60cee36e175e
    # Auto-adjust columns' width
    # NOTE: Need XlsxWriter package!
    MAX_COLUMN_WIDTH = 20
    OFFSET = 2
    for column in dataframe:
        data_len = dataframe[column].astype(str).map(len).max() + OFFSET
        name_len = len(column) + OFFSET
        
        # Column width should NOT exceed max
        column_width = min(data_len, MAX_COLUMN_WIDTH)
        # ...unless the name is longer
        column_width = max(column_width, name_len)
        
        # Get index of column
        col_idx = dataframe.columns.get_loc(column)    

        # Set width (while also making sure wrapping is enabled)     
        writer.sheets[sheet_name].set_column(col_idx, col_idx, column_width, wrap_format)
        
def convert_pandas_col_to_excel_col(dataframe, col_name):
    # Get column index
    index = dataframe.columns.get_loc(col_name)    
    # Use base-26 (where the "digits" are A, B, etc.)
    num_text = ""
        
    while index > -1:
        # Get remainder
        rem = index % 26        
        digit = chr(rem + ord('A'))
        num_text = digit + num_text
        index = index // 26
        index -= 1
            
    return num_text

        