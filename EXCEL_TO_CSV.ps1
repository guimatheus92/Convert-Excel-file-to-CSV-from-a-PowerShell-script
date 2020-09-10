# Defines the directory and name of the file to be exported to the CSV file
$dir = "D:\"

# Defines the name of the Excel file
$excelFileName = "YOUR_FILE_NAME"

# Define a function to convert the file
Function ExportWSToCSV ($excelFileName, $csvLoc)
{
    $excelFile = $dir + $excelFileName + ".xlsx"
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false
    $wb = $E.Workbooks.Open($excelFile)
    foreach ($ws in $wb.Worksheets)
    {
        $n = $excelFileName
        $ws.SaveAs($csvLoc + $n + ".csv", 6)
    }
    $E.Quit()
}

# For each file in the directory with the xlsx format, convert to CSV using the function above
$ens = Get-ChildItem $dir -filter *.xlsx
foreach($e in $ens)
{
    ExportWSToCSV -excelFileName $e.BaseName -csvLoc $dir
}