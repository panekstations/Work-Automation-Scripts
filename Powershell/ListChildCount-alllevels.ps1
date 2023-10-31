# Displays all folder (and sub-folders) with the count of all files below each level. 

Get-ChildItem -Recurse | Where-Object { $_.PSIsContainer } | ForEach-Object {
    $FolderName = $_.FullName
    $FileCount = (Get-ChildItem -File -Recurse $FolderName | Measure-Object).Count
    Write-Host "$FolderName - Child File Count: $FileCount"
}


