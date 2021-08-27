# How to convert an Excel file using a PowerShell script

A document repository can also be found in my profile article at [Medium](https://guimatheus92.medium.com/convert-excel-file-to-csv-from-a-powershell-script-3b998b9e8c2f "Medium").

------------
For the script EXCEL_TO_CSV_ALLTABS.ps1, where we convert all tabs of an Excel file, just enter the name of the directory, the name of the file and the file format.
The script use the function to convert Excel files with XLSX to CSV format, however this format can be changed in the script.

If you want, you can change the folder and your filename
```shell
# Defines the directory where the file is located
$dir = "D:\"

# Defines the name of the Excel file
$excelFileName = "YOUR_FILE_NAME"
```

For the second script EXCEL_TO_CSV.ps1, where we convert just one tab of an Excel file, just enter the name of the directory and the name of the file in the script below.

