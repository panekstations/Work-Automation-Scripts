# good way to check from new processes, malware, etc.

# save all processes to a file on desktop to XML file
get-process | export-clixml -path C:\users\username\Desktop\processlist.xml

# compare all currently running processes to original list (importing xml list)
Compare-Object -ReferenceObject (Import-Clixml C:\Users\username\Desktop\processlist.xml) -DifferenceObject (Get-Process) -Property Name
