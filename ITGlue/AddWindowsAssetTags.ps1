<# 
.SYNOPSIS 
Create ITGlue configuration with asset tag info

.DESCRIPTION
Uses API key from ITGlue to create a new asset and assign asset tag number. Done by computerinfo and created through a post request to the REST API.
This script must be run before ninja rmm gets installed on the machine.

.AUTHOR
    Josh Razalan
#>
Function SetBIOSAssetTag {
    #By manufacture install bios utilties to set asset tag info
    if ($Manufacturer -like "*Dell*") {
        $BiosUtil = "\\Path\to\Dell-Command-Configure.EXE"
        Start-Process -FilePath $BiosUtil -ArgumentList "/s -NoNewWindow -Wait -PassThru"
        Start-Sleep -s 10

        #Call program to set Asset Tag Number
        $SetAssetNum = { Start-Process -FilePath "C:\Program Files (x86)\Dell\Command Configure\X86_64\cctk.exe" -ArgumentList "--asset=$AssetTagNum" }
        Invoke-Command -ScriptBlock $SetAssetNum
        Write-Host "Successfully wrote to $Manufacturer bios."
    }

    if ($Manufacturer -like "*Lenovo*") {
        $BiosUtil = "\\Path\to\giaw03ww.exe"
        Start-Process -FIlePath $BiosUtil -ArgumentList "/VERYSILENT -NoNewWindow -Wait -PassThru"
        Start-Sleep -s 10

        #Call program to set Asset Tag Number
        $SetAssetNum = { C:\DRIVERS\WINAIA\WinAIA.exe -silent -set "USERASSETDATA.ASSET_NUMBER=$AssetTagNum" }
        Invoke-Command -ScriptBlock $SetAssetNum
        Write-Host "Successfully wrote to $Manufacturer bios."
    }

    if ($Manufacturer -like "*Microsoft*") {
        #Determine type of architecture to copy correct program over
        if ((Get-CimInStance Win32_OperatingSystem).OSArchitecture -eq "64-Bit") {
            Copy-Item -Path "\\Path\to\AssetTag.exe" -Destination "C:\AssetTag.exe"
            $BiosUtil = "C:\AssetTag.exe"
        }
        else {
            Copy-Item -Path "\\Path\to\AssetTag_x86.exe" -Destination "C:\AssetTag.exe"
            $BiosUtil = "C:\AssetTag_x86.exe"
        }
        $SetAssetNum = { Start-Process -FilePath $BiosUtil -ArgumentList "-s $AssetTagNum" }
        Invoke-Command -ScriptBlock $SetAssetNum
        Write-Host "Successfully wrote to $Manufacturer bios."
        Remove-Item -Path $BiosUtil
    }

    if (!($SetAssetNum)) {
        Write-Host "Please add asset tag information to BIOS manually."
    }
}

function InstallNinja {
    #Set path for ninja install by organization
    $NinjaOrg = @{
        1  = '\\Path\to\Colorado.msi';
    }
    $FilePath = $NinjaOrg.$locnum
    $Arguments = "/i `"$filepath`" /quiet"
    Start-Process msiexec.exe -ArgumentList $Arguments -wait    
}

#Prompt for Asset Tag Number to write to Bios if supported
$AssetTagNum = Read-Host "Please enter asset tag number"

#Get required computer information
$Computer = Get-ComputerInfo 
$Manufacturer = $Computer.CsManufacturer
$ComputerName = $Computer.CsName
$SerialNum = $Computer.BiosSeralNumber
$Model = $Computer.CsModel

#Get number location for office to set ITGlue org and Ninja org
Do {
    $locnum = Read-Host "Enter number for location:`n1.Colorado `n"
    [int]$locnum = [convert]::ToInt32($locnum, 10)
    if ($locnum -ge 1 -and $locnum -le 1) {
        $valid = $true
    }
}
until ($valid) { }  

#Call function to write asset tag number to bios
SetBIOSAssetTag $Manufacturer $AssetTagNum

#Call function to install Ninja by organization
InstallNinja $locnum

#Call function to create POST Request
. \\Path\to\POSTNewWorkstation.ps1
POSTRequest $ComputerName $AssetTagNum $SerialNum $Manufacturer $Model $locnum 