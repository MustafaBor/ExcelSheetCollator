# ExcelSheetCollator
App to collate, view, search, open specific excel sheet tabs for a large collection of worksheets (i.e. thousands) contained in multiple workbooks

# About
A simple app to idenfity and access large number of excel files and worksheets.

# Background
I've created the app to be able to quickly find a worksheet (via list or search function) in a large collection of workbooks (around 50 of them), each containing anywhere from 5 to 250 worksheets. 

# Requirements
  - C# 7.3
  - Microsoft.Office.Interop.Excel Assembly

# Usage
  1. Choose directory where the excel workbooks are located (default set to C:\\) by clicking on the Select Folder button
  2. Click on Load button to populate the listbox (might take some time depending on the total number of worksheets)
  3. Single-click on the item from the list to open the relevant workbook and worksheet
  (or enter the keyword in the textbox to search for specific worksheet, click on Search button and single-click on the item from the refined list in the right listbox)

# Known/Potential Issues
  - Occassionally excel files launch after system reboot (only after first reboot).
  - Tested on a few thousand worksheets without issues but not tested with larger size which might cause freezing, etc.
