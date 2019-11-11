$SourceFolder = "Y:\My Documents\test" # Does not require final \
$DestinationFolder = "Y:\My Documents\test\out\" # Must include final \

Write-Host "This script is running. Please be patient."

if (!(Test-Path $DestinationFolder)) {
	New-Item $DestinationFolder -type directory
}
ForEach ($Workbook in (Get-ChildItem -Path $SourceFolder -Include "*.xls" -Recurse)) {
	$MSXL = New-Object -ComObject excel.application
	$MSXL.Application.DisplayAlerts = $false
    $MSXL.Visible = $false
		$WorkbookTypes = Add-Type -AssemblyName 'Microsoft.Office.Interop.Excel' -Passthru
		$WorkbookSaveFormat = $WorkbookTypes | Where {$_.Name -eq "xlSaveFormat"}	
		$OpenWorkbook = $MSXL.Workbooks.Open("$workbook", 2, $true)
		$Filename = ($DestinationFolder + ($Workbook.name -replace '\.xls', ''))
		$OpenWorkbook.SaveAs($Filename, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)
	Stop-Process -name "EXCEL" -force
	$MSEXCEL = $Null	
	if (!(Test-Path -path "$Filename.csv")) {
		"FAILURE: $Workbook" | Out-File -Append "$DestinationFolder\Failed.txt"
	}
}

Write-Host "."
Write-Host "."
Write-Host "."
Write-Host "."
Write-Host "Completed. A failure log will appear if there are errors. This window will close in 10 seconds"
Start-Sleep -s 10
