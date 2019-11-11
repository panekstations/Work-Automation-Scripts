$getFirstLine = $true

get-childItem "*.csv" | foreach {
    $filePath = $_

    $lines = Get-Content $filePath  
    $linesToWrite = switch($getFirstLine) {
           $true  {$lines}
           $false {$lines | Select -Skip 1}

    }

    $getFirstLine = $false
    Add-Content "final.csv" $linesToWrite
    }
