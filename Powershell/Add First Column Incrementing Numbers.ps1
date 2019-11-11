------Add first column called 'Row' with incrementing number

$csv = Import-Csv "C:\Users\username\Desktop\splitfile_3.csv"
 $global:i = 1000000; $csv | Select-Object @{ Name = 'Row'; Expression = { $global:i.ToString("0000"); $global:i += 1 } }, * | Export-Csv "C:\Users\username\Desktop\file.csv"
