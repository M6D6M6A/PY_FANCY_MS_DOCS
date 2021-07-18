# py_fancy_ms_docs

## py_fancy_ms_docs	- Let you focus on data AND file formats

- Author:		      Philipp Reuter
- Version:      	1.0.0
- Generated:    	May 16, 2020
- Last Update:    	Jul 18, 2021
- Idea based on:	http://docs.pyexcel.org/en/latest/


### Introduction
py_fancy_ms_docs provides one application programming interface to read, manipulate and write data in .xlsx excel formats (Maybe other too, idc). 
This library makes information processing involving excel files an enjoyable task.
The original library focuses on data processing using excel files as storage media hence fonts, colors and charts were not and will not be considered.

So I went a different way and extract the file as a zip file to memory.
You can insert data with row and column into a fully formatted Excel and the file keeps all formatting from the original.
The idea originated from the common usability problem:
You want to automatically insert data into an Excel, but pandas and pyexcel destroy all the formatting?


### How to use
> `from py_fancy_ms_docs.py_fancy_excel import Excel`

> `excel_file = Excel("Path/to/File")` To create / override with empty Excel, add empty=True

> `excel_file.add_data("Test", 1, 1, 1)` (data, row_index, column_index, sheet_index)
>   > row_index, column_index and sheet_index range from 1, 2, ...

> `excel_file.save_excel(path="Path/to/new/File.xlsx")` Save the edited Excel as file (Feel free to test different file extensions, .zip works!)

### Extra Features
> `excel_file.save_as_folder()` # Extracts the Excel file to Folder to see Excel File Contents

> `excel_file.save_as_json()` # Extracts the Excel file to Json to see Excel File Contents, like they are stored in the Excel() class


### Planned Features (not ordered)
- [ ] Port for all Microsoft Word Document Types
- [ ] Editing the Format in Python
- [ ] Clone / Add Sheets
- [ ] Clone / Add / Create Tables
- [ ] Insert formulas
- [ ] Edit Font, Colors & Borders
- [ ] Improving Code Qualtity and Speed
- [ ] Improve the ease of use
- [ ] Image support
- [ ] Plots and Charts
