function Convert-Anchor {
    # v1.1
    param (			
        [Parameter(Mandatory = $true)]    
        [string] $Value
    )

    # trim input
    $Value = $Value.ToString().Trim()
    
    # validate GUID
    function Validate-GUID {
        param([string]$GUID)
        $SampleGUID = [System.Guid]::Empty
        [GUID]::TryParse($GUID, [ref]$SampleGUID)  
    }
    
    # validate Base64
    function Validate-Base64 {
        param([string]$Base64)
        try {
            [Convert]::FromBase64String($Base64) | Out-Null
            return $true
        } catch {
            return $false
        }
    }

    # validate HEX
    function Validate-Hex {
        param([string]$HEX)
        $HEX -match '^([0-9A-Fa-f]{2}\s){15}[0-9A-Fa-f]{2}$'
    }

    # export object
    $ExportObject = [PSCustomObject]@{
        GUID = ""
        Base64 = ""
        HEX = ""
        ByteArray = "" # AD takes ByteArray
    }

    # switch input
    switch -Regex ($Value) {
        '^[\da-fA-F]{8}-([\da-fA-F]{4}-){3}[\da-fA-F]{12}$' {
            if (Validate-GUID($Value)) {
                $ExportObject.GUID = $Value
                $ExportObject.Base64 = [system.convert]::ToBase64String(([GUID]$Value).ToByteArray())
                $ExportObject.HEX = (([GUID]$Value).ToByteArray() | ForEach-Object ToString X2) -join ' '
                $ExportObject.ByteArray = ([GUID]$Value).ToByteArray()
            } else {
                Write-Error "Invalid GUID: $Value"
            }
        }
        '^[A-Za-z0-9\+\/\=]+$' {
            if (Validate-Base64($Value)) {
                $ExportObject.GUID = ([GUID]([system.convert]::FromBase64String($Value))).Guid
                $ExportObject.Base64 = $Value
                $ExportObject.HEX = ([system.convert]::FromBase64String($Value) | Foreach-Object ToString X2) -join ' '
                $ExportObject.ByteArray = ([system.convert]::FromBase64String($Value))
            } else {
                Write-Error "Invalid Base64: $Value"
            }
        }
        '^(?:[0-9A-Fa-f]{2}\s){15}[0-9A-Fa-f]{2}$' {
            if (Validate-Hex($Value)) {
                $ExportObject.GUID = [GUID]([byte[]] (-split (($Value -replace ' ', '') -replace '..', '0x$& ')))
                $ExportObject.Base64 = [system.convert]::ToBase64String([byte[]] (-split (($Value -replace ' ', '') -replace '..', '0x$& ')))
                $ExportObject.HEX = $Value
                $ExportObject.ByteArray = [System.Convert]::FromHexString($Value)
            } else {
                Write-Error "Invalid HEX: $Value"
            }
        }
        default {
            Write-Error "Invalid input format. Please provide a valid GUID, HEX, or Base64 string."
        }
    }

    # return
    return $ExportObject
}
