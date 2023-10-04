$sourceDirectory = "H:\BackUp\ABR1\ARCHIVELOG"
$destinationDirectory = $sourceDirectory + "1"

$sourceDirectoryF=$destinationDirectory + "\*"
$archiveThreshold = (Get-Date).AddMonths(-2)
$archiveThreshold = Get-Date -Year $archiveThreshold.Year -Month $archiveThreshold.Month -Day 1
$winrarPath = "C:\Program Files\WinRAR\WinRAR.exe"
$archiveThreshold1 = $archiveThreshold.AddMonths(-1)
$archiveName = $sourceDirectory + "\" + ($archiveThreshold1.Month.ToString() + '-' + $archiveThreshold1.Year.ToString() + ".rar")
$archiveThresholdDelete = (Get-Date).AddYears(-1)


# Check if the source directory exists
    if (Test-Path $sourceDirectory -PathType Container) {
   
   
        # Check if the destination directory exists, and if not, create it
       
       
        if (-not (Test-Path $destinationDirectory -PathType Container)) {
            New-Item -Path $destinationDirectory -ItemType Directory | Out-Null
        }
    
  
        $files = Get-ChildItem $sourceDirectory | Where-Object {$_.PSIsContainer -eq $true -and $_.LastWriteTime -lt $archiveThreshold }
                
        
        foreach ($file in $files) {
            $sourcePath = $file.FullName
            $destinationPath = Join-Path -Path $destinationDirectory -ChildPath $file.Name

            # Use double quotes around the paths to handle spaces in file names
            Move-Item -Path "$sourcePath" -Destination "$destinationPath"            

            Write-Host $file.FullName

            Sleep -Milliseconds 500
        }
      
      
        Start-Process -FilePath $winrarPath -ArgumentList "a -r -df  -x*.rar $archiveName $sourceDirectoryF " -Wait

       #Check the directory exists before attempting to delete it
                
        
        if (Test-Path -Path $destinationDirectory -PathType Container) {
            #Delete the directory and its contents

           Remove-Item -Path $destinationDirectory -Recurse -Force

          Write-Host "Directory '$destinationDirectory' has been deleted."

        } else {
           Write-Host "Directory '$archiveName' does not exist."
        }

            $files = Get-ChildItem $sourceDirectory | Where-Object { $_.PSIsContainer -eq $false -and [System.IO.Path]::GetExtension($_.FullName) -eq '.rar' }

            foreach ($file in $files) {
            # Determine the format based on the length of the month component
           
               $month, $year = $file.BaseName -split '-'
           
                if ($month.Length -eq 1) {
                    $format = "M-yyyy"  # Single-digit month without leading zero
                } else {
                    $format = "MM-yyyy" # Two-digit month
                }

                $dateTime = [datetime]::ParseExact("$month-$year", $format, [System.Globalization.CultureInfo]::InvariantCulture)


                # Compare the difference to 1 year
                if ($dateTime -lt $archiveThresholdDelete) {

                    Remove-Item -Path $file.FullName  -Force

                } else {
                    Write-Host "The difference between the two dates is 1 year or less."
                }

                Sleep -Milliseconds 500
            }
 
    } else {
        Write-Host "Source directory does not exist: $sourceDirectory"
    }