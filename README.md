# ExcelMacros
This repository includes the VBA code for videos posted on my Linkedin page. 

https://www.linkedin.com/in/urinussbaum/

# Multiple Column Vlookup

VBA Custom Function to lookup a range of values in mutliple columns. 

Link to post: http://bit.ly/2mphQ4w

# Sum Text

VBA Custom Function to improve on Excel Concat Function. 

This adds together Excel range of text while adding a seperator between each cell. 

Link to post: http://bit.ly/2mAMw2M

# Color Fun

VBA Custom Function that sums up columns by cell color and function that sums up columns by font color.

First Argument is Range to sum, while second agrument is color. (Must be range that contains the color criteria)

Link to post: http://bit.ly/2HdLNMs

# Mul Vlookup
VBA Custom Function (known as UDF) that returns all matches for a Vlookup in an array.

This returns an array which in excel requires selecting mutliple cells then clicking CTRL + SHFT + ENTER.

parameters of Function are:
1. SearchValue - The value you are attempting to search in Table. Can be hard coded or Cell.
2. SearchInCol - Column where you are matching the search value to.
3. ReturnVal - Column which you are pulling the values from. Should be same size as parameter 2. 

Link to post: http://bit.ly/2OJ2l4U

# Capture
VBA code to capture full screen and attach to email. 
Second Module captures screen based off size of rows and columns and attaches to email. 

Requirements:
You will need to download Boxcutter application http://keepnote.org/boxcutter/. 
Change the path to application in macro according to where installed on your computer. 

Add References to Microsoft Outlook and Microsoft Scriptime Runtime in VBA Tools>References. 

Link to post: https://bit.ly/332lD8I

# Comments

The first macro finds comments on your data and prints them out in first open column but on same row.

The second macro moves entire row to new sheet for more examination. 

Link to post: https://bit.ly/2XZC3eQ

# Split Data

This Macro splits table rows (based off column F) into seperate worksheets. 
It then creates new workbooks for each worksheet and saved them in a location you are prompted to choose. 

It starts looping only from row 2. It also only copies column A:J. Change according to your needs. 

Just copy and paste code into your VBA Module and change the Column Number to your desired column.

# Ordinal Suffix Function

This switches numbers from cardinal numbers into ordinal numbers. Used for addresses in many locations. 

Link to post: https://bit.ly/2wSGwnZ

# Fit By Column Header

This code fits columns by header size in row 1. It will work regardless of Font Size. 

Link to post: https://bit.ly/2MSpwJf

# Full Contact API

###### Requirements:

1. A Full Contact API Key generated here: https://bit.ly/2Rj8RwW
2. Json Converter Library for VBA. Link: https://bit.ly/2ZCrO0J
3. Add reference to Microsoft Scripting Runtime under tools>references.

###### Description:

This Macro pulls the full contact info for the email in the active cell.
You can run for company or phone numbers by editing end points.
Full Documentation is: https://bit.ly/2WQQgP4

Link to post: https://bit.ly/2wZtgxX

# Filter By Last Week

###### Description:
This code will filter data to only the last week. In the code the date was in column C. Change as needed.

Link to post: https://bit.ly/31RHeA5

# Protect Sheet

###### Description:
The first sub will lock sheet excluding your selected cell or range. Second sub will unlock sheet and reset your selected cell/range lock property.

Link to post: https://bit.ly/2NvM8Q6

