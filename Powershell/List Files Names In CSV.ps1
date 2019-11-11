Get-ChildItem *.* | Select-Object FullName | Export-Csv FileList.csv -NoTypeInformation
