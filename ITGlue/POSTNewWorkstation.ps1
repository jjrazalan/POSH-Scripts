<# 
.SYNOPSIS 
Create POST Request for ITGlue new workstation.

.DESCRIPTION
Call this function in other scripts to prompt for the correct office and then generate the POST request to create a 
new workstation with the correct asset tag information. IDs were discovered from the GET requests follwing the ITGlue
API Documentation.

.AUTHOR
    Josh Razalan
#>

function POSTRequest {
    #Set location ID from ITGlue
    $Location = @{
        1  = '';
    }
    $ManufacturerHash = @{
        "DELL"                  = ''
        "LENOVO"                = ''
        "Microsoft Corporation" = ''
    }

    #Required variables for the web request
    $LocationID = $Location.$locnum
    $ConfigURI = "https://api.itglue.com/organizations/$LocationID/relationships/configurations/"
    $APIKey = ""
    $ManufacturerID = $ManufacturerHash.$Manufacturer
    $Notes = "Model: $Model"

    #Create hashtable and convert to json for body of request
    $body = @{
        data = @{
            type       = "configurations"
            attributes = @{
                name                    = $ComputerName
                "configuration-type-id" = ""
                "asset-tag"             = $AssetTagNum
                "serial-number"         = $SerialNum
                "Manufacturer-id"       = $ManufacturerID
                notes                   = $Notes
            }
        }
    } | convertto-json

    invoke-restmethod -method POST -uri $ConfigURI -body $body -header @{"x-api-key" = $APIKey } -ContentType application/vnd.api+json
    Write-Host "Successfully created $ComputerName In ITGLue"

}