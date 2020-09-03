# ExcelSheetCollator
App to collate, view, search, open specific Excel sheet tabs for a large collection of worksheets (i.e. thousands) contained in multiple Excel files (ie. workbooks).

# About
A simple app to idenfity and access large number of Excel worksheets contained in multiple files. The app lists worksheet names followed by file names for all files in the chosen directory. The user can search for keywords to refine listing and the specific tab will open in excel upon clicking on the list item.

# Background
I've created the app to be able to quickly find a worksheet (via list or search function) in a large collection of workbooks (around 50 of them), each containing anywhere from 5 to 250 worksheets. 

# Requirements
  - C# 7.3
  - Microsoft.Office.Interop.Excel Assembly

# Usage
  1. Choose directory where the Excel workbooks are located (default set to C:\\) by clicking on the `Select Folder` button
  
  <img src="/images/Excel_1.jpg" alt="Excel Image 1" width="650">


  2. Click on Load button to populate the listbox (might take some time depending on the total number of worksheets). The list will contain worksheet name and workbook name separated by a dash surrounded by double spaces.
  
  <img src="/images/Excel_2.jpg" alt="Excel Image 2" width="650">

  3. Enter keyword in the textbox to search for specific worksheet and click on Search button. This will populate right listbox with all worksheets that contain the keyword.
  
  <img src="/images/Excel_3.jpg" alt="Excel Image 3" width="650">

  4. Single-click on the item from the refined list in the right listbox to open the relevant workbook and worksheet. This will open the Excel file and automatically select the worksheet tab chosen.
  
  <img src="/images/Excel_4.jpg" alt="Excel Image 4" width="650">

# Known/Potential Issues
  - Occassionally excel files launch after system reboot (only after first reboot).
  - Tested on a few thousand worksheets without issues but not tested with larger size which might cause freezing, etc.
