class ExcelFile
=========

This class allow to write in excel files using the openpyxl library.

Attributes
__________
file_path : string  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The path where the excel file is stored

excel_file :  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;excel file modelling with a Workbook object from openpyxl library

Parameters
__________
file_path : string  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The path where you want to write your excel file

Methods
_______
create_empty_excel_file()  
&nbsp;&nbsp;&nbsp;&nbsp;Create an empty excel file in file_path and store it in excel_file attribute

load_excel_file()  
&nbsp;&nbsp;&nbsp;&nbsp;Load an existing excel file and store it in excel_file attribute

write_in_cell_of_sheet(sheet, row, column, entry)  
&nbsp;&nbsp;&nbsp;&nbsp;Write the <entry> value in the row, column cell of the given sheet of the excel_file object

write_list_in_line_of_sheet(self, sheet, line_number, list_entry, start_column=1)  
&nbsp;&nbsp;&nbsp;&nbsp;Write the <list_entry> in the row, column cell of the given sheet of the excel_file object