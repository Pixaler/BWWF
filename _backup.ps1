# Check that you in a D:\BWWF folder
Write-Host "---You in a D drive?---"
Write-Host "."
Start-Sleep -Seconds 2
Write-Host "."
Write-Host "---Checking location----"
Write-Host "."
$checklocation = Get-Location
Write-Host "."
if ($checklocation -match 'D:' -eq $false) {
  Clear-Host
  Write-Host "___Change script folder location!___"
  Write-Host "."
  Start-Sleep -Seconds 5
  Write-Host "."
  Exit
}

# Close all apps
Clear-Host
Write-Host "---Close all apps---"
Write-Host "." 
(get-process | ? { $_.mainwindowtitle -ne "" -and $_.processname -ne "powershell" } )| stop-process # Close all windows, except explorer.exe and powershell.exe
Start-Sleep -Seconds 5

# Dismount K drive
Clear-Host
Write-Host "---Dismount K drive---"
Write-Host "."
C:\!MOE\VeraCrypt\VeraCrypt-x64.exe /l K /d /f /q # /l K - letter of drive that you dismount. Change it, if its diffrent or remove command
Write-Host "."
Start-Sleep -Seconds 5

# Add !MOE, PortableApps and BATCH to archive and save archive in D:\ drive
Clear-Host
Write-Host "---Add all to archive---"
Write-Host "."
$date = Get-Date -Format "dd-MM-yyyy"
Write-Host "."
./7Zip/x64/7za.exe a !BACKUP.zip C:\!MOE C:\PortableApps C:\BATCH -oD:\ -mx=9 -tzip # Add all folders to archive with 7za.exe
Write-Host "---Copy to root of D and add $date to name---"  
Write-Host "."  
Move-Item -Path D:\BWWF\!BACKUP.zip -Destination "D:\!BACKUP($date).zip"
Write-Host "."  
Write-Host "All done!"
Start-Sleep -Seconds 5
Exit