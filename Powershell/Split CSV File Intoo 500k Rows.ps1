
<#Split CSV into different files of 500000 row increments with the name splitfile #>

<#Directory to Save #>
c:\Users\username\Desktop


<#Replace Location of file #>
$i=0; Get-Content c:\Users\username\Desktop\file.csv -ReadCount 500000 | %{$i++; $_ | Out-File c:\Users\spanek\Desktop\splitfile2_$i.csv}
