# Check that you in a D:\BWWF folder
Write-Host "---Backup or restore?:---"
Write-Host "."
$setup_my_apps = Read-Host -Prompt "You options is (b/r)"
if ($setup_my_apps -like 'b') {
    # Close all apps
    Clear-Host
    Write-Host "---Close all apps---"
    Write-Host "." 
    # Close all windows, except explorer.exe and powershell.exe
    (get-process | ? { $_.mainwindowtitle -ne "" -and $_.processname -ne "powershell" } )| stop-process -force 
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
    Write-Host "."
    $date = Get-Date -UFormat "%d.%m.%Y"
    $path1 = "..\!BACKUP(" + $date + ").zip"
    $path2 = "..\!BACKUPDATA(" + $date+ ").zip"
    ./7Zip/x64/7za.exe a $path1 C:\!MOE C:\PortableApps C:\BATCH -mx=5 -tzip # Add all folders to archive with 7za.exe
	./7Zip/x64/7za.exe a $path2 D:\07_Projects D:\08_Notes -mx=5 -tzip # Add all folders to archive with 7za.exe
	# Move-Item -Path "!BACKUP*.zip" -Destination "D:\!BACKUP"
	# Move-Item -Path "!BACKUPDATA*.zip" -Destination "D:\!BACKUP"
    Start-Sleep -Seconds 5

}
if ($setup_my_apps -like 'r') {
    Write-Host "---Starting extract files---"
    Write-Host "."
    ./7Zip/x64/7za.exe x D:\!BACKUP\!BACKUP*.zip -oC:\
    Exit
}
