powershell.exe Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
start "" caffeine64.exe
start "" powershell.exe "./_backup.ps1"