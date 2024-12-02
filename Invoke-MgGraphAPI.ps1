Function Invoke-MgGraphAPI {
    # version 2.8
    # https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0#using-invoke-mggraphrequest
    param (
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]
        [String]$Method,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]
        [String]$Endpoint,
        [Parameter(Mandatory = $false)]
        $Body,
        [Parameter(Mandatory = $false)]
        [switch]$Beta
    )

    # If the -Beta switch is used, set the API version to beta
    $APIVersion = if ($PSBoundParameters.ContainsKey('Beta')) { "beta" } else { "v1.0" }

    # Check if the endpoint already contains the API version
    $URI = if ($Endpoint -match "/v1.0/" -or $Endpoint -match "/beta/") {
        "https://graph.microsoft.com$Endpoint"
    } else {
        # Generalize to adapt any endpoint
        if ($Endpoint -match "^v1.0/") {
            "https://graph.microsoft.com/$Endpoint"
        } elseif ($Endpoint -match "^/") {
            "https://graph.microsoft.com/v1.0$Endpoint"
        } else {
            "https://graph.microsoft.com/v1.0/$Endpoint"
        }
    }

    # Create an empty list to store the data
    $Data = [System.Collections.Generic.List[Object]]@()

    # Set the maximum number of attempts
    $MaxAttempts = 5

    # Set the initial number of attempts 
    $Attempts = 1

    # Set the initial value of the success flag to false
    $Success = $false

    # If the method parameter is not specified, set the method to GET
    if (-not $Method) {
        $Method = "GET"
    }

    while ($Attempts -le $MaxAttempts -and -not $Success) {
        try {
            switch ($Method) {
                # If the method is GET, loop until there are no more pages of data
                "GET" {
                    do {
                        $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -ContentType "application/json"
                        $Success = $true
                        # If the response contains a value property, add the value property to the data list
                        if ($null -ne $Response.Value) {
                            $Data.AddRange($Response.value) 
                            # If the response contains an @odata.nextLink property, set the URI to the value of the @odata.nextLink property (paging)
                            if ($Response.'@odata.nextLink') {
                                # Set the URI to the value of the @odata.nextLink property
                                $URI = $Response.'@odata.nextlink'
                            } 
                            else {
                                # If there is no subsequent @odata.nextLink property, set the URI to null to exit the loop
                                $URI = $null
                            }              
                        }
                        # If the response does not contain a value property, add the response to the data list and set the URI to null to exit the loop
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
                    $Success = $true
                    return $Response
                }
                "PUT" {
                    $Body = $Body | ConvertTo-Json
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -Body $Body -ContentType "application/json"
                    $Success = $true
                    return $Response
                }
                "PATCH" {
                    $Body = $Body | ConvertTo-Json
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI -Body $Body -ContentType "application/json"
                    $Success = $true
                    return $Response
                }
                "DELETE" {
                    $Response = Invoke-MgGraphRequest -Method $Method -Uri $URI
                    $Success = $true
                    return $Response
                }
            }    
        }
        catch {
            # If the call fails due to rate limiting (throttling), retry the call
            if ($_.Exception.Response.StatusCode -eq '429') {
                Write-Host "Encountered throttling. Retrying...`nAttempt: $($Attempts) of $($MaxAttempts)" -ForegroundColor Yellow
                # If the response contains a Retry-After header, wait for the specified number of seconds
                $RetryAfter = $_.Exception.Response.Headers['Retry-After']
                Start-Sleep -Seconds $RetryAfter
                # Increment the number of attempts
                $Attempts++
            }
            else {
                Write-Host -ForegroundColor Red "StatusCode: $($_.Exception.Response.StatusCode)"
                Write-Host -ForegroundColor Red "Exception message: $($_.Exception.Message)"
                # Exit the loop if the exception is not due to throttling
                $Success = $true
            }
        }
    }

    return $Data

}
