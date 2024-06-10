# invoke the Microsoft Graph API
Function Invoke-MgGraphAPI {
    # version 2.3
    # https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0#using-invoke-mggraphrequest
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]
        [String]$Method,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]
        [String]$Endpoint,
        [Parameter(Mandatory = $false)]
        $Body
    )
    $URI = "https://graph.microsoft.com/" + $Endpoint
    $Data = [System.Collections.Generic.List[Object]]@()

    $maxAttempts = 5
    $attempts = 0
    $success = $false

    if (-not $Method) {
        $Method = "GET"
    }

    while ($attempts -lt $maxAttempts -and !$success) {
        try {
            switch ($Method) {
                "GET" {
                    do {
                        $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -ContentType "application/json"
                        $success = $true
                        # if $Response contains multiple objects
                        if ($null -ne $Response.Value) {
                            $Data.AddRange($Response.value) 
                            # if $Response contains '@odata.nextLink' - paging
                            if ($Response.'@odata.nextLink') {
                                $URI = $Response.'@odata.nextlink'
                            } 
                            else {
                                $URI = $null
                            }              
                        }
                        # $Response contains a single object
                        else {
                            $Data.Add($Response)
                            $URI = $null
                        }
                    } until ($null -eq $URI)     
                    return $Data
                }
                "POST" {
                    $Body = $Body | ConvertTo-Json
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -Body $Body -ContentType "application/json"
                    $success = $true
                    return $Response
                }
                "PUT" {
                    $Body = $Body | ConvertTo-Json
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -Body $Body -ContentType "application/json"
                    $success = $true
                    return $Response
                }
                "PATCH" {
                    $Body = $Body | ConvertTo-Json
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -Body $Body -ContentType "application/json"
                    $success = $true
                    return $Response
                }
                "DELETE" {
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI
                    $success = $true
                    return $Response
                }
            }    
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq '429') {
                Write-Host "Encountered throttling. Retrying...`nAttempt: $($attempts + 1) of $($maxAttempts + 1)" -ForegroundColor Cyan
                #If the API call fails due to rate limiting, get the Retry-After header and wait for the specified time
                $retryAfter = $_.Exception.Response.Headers['Retry-After']
                Start-Sleep -Seconds $retryAfter
                #Increment the number of attempts
                $attempts++
            }
            else{
                Write-Host -ForegroundColor Red "Exception type: $($_.Exception.GetType().FullName)"
                Write-Host -ForegroundColor Red "Exception message: $($_.Exception.Message)"
                if ($($_.ErrorDetails.Message | ConvertFrom-Json).error.code) { Write-Host -ForegroundColor Red "Error detail code:" $($_.ErrorDetails.Message | ConvertFrom-Json).error.code }
                if ($($_.ErrorDetails.Message | ConvertFrom-Json).error.message) { Write-Host -ForegroundColor Red "Error detail message:" $($_.ErrorDetails.Message | ConvertFrom-Json).error.message }
            }
        }
    }
}
