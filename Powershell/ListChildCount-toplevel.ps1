# Run in a folder and it will list all the top-level folders and how many child files are contained in all subdirectories

Get-ChildItem -Directory | ForEach-Object {
    $FolderName = $_.FullName
    $FileCount = (Get-ChildItem -File -Recurse $FolderName | Measure-Object).Count
    Write-Host "$FolderName - Child File Count: $FileCount"
}
