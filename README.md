# Autofit-CSV
## A small Excel add-in to make CSV files human-readable upon opening
---
##
The add-in currently performs: 
- Adds wrap-text to header row (first row), so auto-sizing is only fit to the subsequent data rows.
- Auto-Resizes columns from the first row to the first completely empty row in the file (to avoid auto-sizing based on unrelated file info tables present at the bottom of certain files)
- Adjusts size of large auto-sized columns to a maximum width based on the monitor size and number of columns present
- Unless manual adjustments are made, there is no save-as prompt for the format modifications performed by the add-in (which cannot be saved in a csv file anyways)
  
*Confirmed Excel Version compatibility: **Excel 365 (Windows 10)**

---
## Install Instructions (Windows only)
#### Script-based install
1. Download Autosize-CSV.xaml or clone the repo
2. Click on Install-AutosizeCsv.ink

#### Manual install
1. Download Autosize-CSV.xaml or clone the repo
2. Right click on Autosize-CSV.xaml, open properties and check the box beside "Unblock File"
3. Copy Autosize-CSV.xaml to your excel add-in folder
4. Enable developer tab in the customize ribbon in Excel's options panel
5. Click on the developer tab and the click on Excel Add-ins
6. Check the box beside Autofit-CSV and click OK
7. Reopen any currently open csv files or open new csv files to have your csv file appear human readable!

### Upcoming Features
- Disable prompt to saving on any format modifications that will not be saved anyways
- Freeze header row (first row) so that remains present while scrolling
- Add a tab or box section to toggle autofitting and header row freezing
- Test for machine's OS and implement appropriate OS compatible functions (to work on mac)
- Test older versions of excel
- Add a textbox section allowing min and max widths after the autosize to be set and saved for use in all subsequent csv files