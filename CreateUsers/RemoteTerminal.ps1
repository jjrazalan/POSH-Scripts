$RemoteComputer = Read-Host "Enter computer name to remote into"
Enter-PSSession -ComputerName $RemoteComputer