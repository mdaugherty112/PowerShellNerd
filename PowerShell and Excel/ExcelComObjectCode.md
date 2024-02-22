# PowerShell Excel Com Objects
This page will cover the basics of using PowerShell and Excel Com objects. 

This relates to my YouTube video [Automate Like a Pro: PowerShell's Excel Com Objects Magic](https://www.youtube.com/watch?v=sYlXXaIWAzA)

🔔 [Make sure to Subscribe at my Channel](https://www.youtube.com/@matthewdaugherty462?sub_confrimation=1) 🔔 

## Create New Excel Object
```PowerShell
$excel = New-Object -Com Excel.Application
```
## Open Existing Excel File
```PowerShell
$importFile = 'C:\Users\User\Downloads\exceldoc.xlsx'
$wb = $excel.Workbooks.Open($importFile)
```

## Make Excel visible
```PowerShell
$excel.Visible = $true
```

## Add Workbook
```PowerShell
$wb = $excel.workbooks.add()
```

## Worksheets
### Select Worksheet
```PowerShell
$ws = $wb.Worksheets.Item(1) 
```
### Add New Worksheet
```PowerShell
# Use [void] so no output prints to screen
[void]$wb.Worksheets.Add()
```
### Visually change worksheet
```PowerShell
$ws.Activate() 
```
### Rename Worksheet
```PowerShell
$ws.name = "metrics"
```
### Get Last Row and Column used

```PowerShell
$ws.UsedRange.columns.count
$ws.UsedRange.rows.count
```
### Read value from cell
```PowerShell
$ws.Cells.Item(1, 1).text
```

### Add or Modify Value in a Cell 
```PowerShell
#update the cell with 'New value'
$ws.Cells.Item(1,1).Value="New value"
```

### Delete Entire Row
```PowerShell
# We use [void] so no output goes to screen
[void]$ws.Cells.Item(1,1).EntireRow.Delete()
```

### Delete Entire Column
```PowerShell
# We use [void] so no output goes to screen
[void]$ws.Cells.Item(1,1).EntireColumn.Delete()
```

### Search Worksheet for string
```PowerShell
# Searching worksheet for 'string' text
$search = $ws.Cells.Find("String")

# Show row 
$search.Row

# Show Column
$search.Column

# Update
$ws.cells.Item($Search.Row, $Search.Column).value = "Updated"
```

### Autosize Columns
```PowerShell
$ws.Columns.Autofit()

# OR

$ws.UsedRange.Columns.autofit()
```

### Autosize Rows
```PowerShell
$ws.Rows.AutoFit()

# OR

$ws.UsedRange.Rows.AutoFit() 
```

### Using vlookup
```PowerShell
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open("C:\path\to\your\file.xlsx") # Update with your file path

# Get the worksheets
$sheet1 = $workbook.Worksheets.Item('Sheet1') 
$sheet2 = $workbook.Worksheets.Item('Sheet2')

# Find the last row in Column A of Sheet2
$lastRow = $sheet2.UsedRange.Rows.Count

# Loop through each row in Column A of sheet2 and put the VLOOKUP formula into column B
for ($i = 1; $i -le $lastRow; $i++){
    # Note: the $i+1 in the formula is to adjust for zero-based index
    $sheet2.Cells($i, 2).formula = "=VLOOKUP(A$i,Sheet1!A:B,2,FALSE)"
}
```

### Autofill cells
```PowerShell
# make sure you have your worksheet already selected
# Get the cell you want to increment under in our case
# B2 has a vlookup 
$range1 = $ws.Range("B2")

# Now Select the range you want to pull it down to
# so you can easily tell this is $ws.Range("B2:B20")
$range2 = $ws.Range("B2:B20") 

# second param speicfies the fill type and 0 is the default
$range1.AutoFill($range2, 0) 
 ```

### Insert Table
```PowerShell
# Select range of cells in worksheet
$range = $ws.Cells.Range("A1:B3") 

# Provide that range to the code below
$table = $ws.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,$range,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)

# Set table style
# NOTE: Get name from excel if you hover tables
$table.TableStyle = "TableStyleDark10"
$table.TableStyle = "TableStyleMedium3"
$table.TableStyle = "TableStyleMedium20"
```

## Workbook
### SaveAs
```PowerShell
$excel.ActiveWorkbook.SaveAs("C:\temp\myexcel.xlsx") 
```
### Save
```PowerShell
$excel.Save() 
```

### Quit (close excel)
```PowerShell
$excel.Quit()
```

















