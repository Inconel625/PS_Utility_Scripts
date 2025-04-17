# Useful Powershell Functions

### Compute SHA1 hash of an input string
```powershell
function Get-SHA1OfString {
    param (
        [string]$InputString
    )
    $MemoryStream = [IO.MemoryStream]::new([byte[]][char[]]$InputString)
    Get-FileHash -InputStream $MemoryStream -Algorithm SHA1
}
```

### Convert an array to a markdown formatted table
```powershell
function ConvertTo-MarkdownTable { # Accepts an input array and returns a markdown format table
    param (
        [Parameter(Mandatory=$true)]
        [array]$InputArray
    )
    if ($InputArray.Count -eq 0) { return "Array is empty!" }
    $headers = $InputArray[0].PSObject.Properties.Name
    $table = @()
    $table += "| " + ($headers -join " | ") + " |" # Add headers
    foreach ($item in $InputArray) { # Add rows
        $row = $headers | ForEach-Object { $item.$_ }
        $table += "| " + ($row -join " | ") + " |"
    }
    return $table -join "`n"
}
```

### IP Abuse Check
```powershell
function Get-IPAbuseData{
    # Function to do a API abuse check of an input ip address using abuseipdb
    # Must supply your own API key
    Param (
        [Parameter(Mandatory,ValueFromPipeline=$True)]
        [string]
        $IPAddress
    )
    # Define the URL and query parameters
    $uri = 'https://api.abuseipdb.com/api/v2/check'
    $querystring = @{
        ipAddress = $IPAddress
    }
    # Define the headers
    $headers = @{
        Accept = 'application/json'
        Key = '________________'
    }
    # Make the GET request
    $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -Body $querystring
    # Output the response
    return $response.Data
}
```

### Geo-IP Lookup function
```powershell
$script:GeoIP_Reference_List = @()
function Get-IPGeolocation {
    # Function to do a API lookup to find the GEOIP location of an input ip address using freeipapi.com
    # Saves reference list of IP's to avoid duplicate lookups
    Param (
        [Parameter(Mandatory,ValueFromPipeline=$True)]
        [string]
        $IPAddress
    )
    # Removes brackets and port if present in input string, should work for both IPv4 and v6 addresses
    $ipAddress = $ipAddress -replace '\[|\]|:\d+$'
    # Check if the IP is already in the $script:GeoIP_Reference_List array, this saves limited API queries and boosts script speed
    $existingGeoIP = $script:GeoIP_Reference_List | Where-Object { $_.IP -eq $IPAddress }
    if ($existingGeoIP) {
        # If the IP is already in the list, return the existing data
        #Write-Host "IP already checked"
        return $existingGeoIP
    }
    else {
        # If the IP is not in the list, invoke the API to get geolocation data
        $script:GeoIPAPIQueries ++ # Increments counter every time the geoip function is run, more than 45 queries per minute gets you banned from ip-api.com
        #Write-Host "QueryCounter "$script:GeoIPAPIQueries
        If ($script:GeoIPAPIQueries -gt 59) { # Pauses queries if needed to avoid going over limit
            Write-Host "Query rate exceeded, pausing for 70s to avoid IP blocking" -ForegroundColor Yellow
            Start-Sleep 60
            $script:GeoIPAPIQueries = 0
            }
        #Error handling function to avoid queries being missed if the API request returns an error
        while ($true) {
            try {
                $request = Invoke-RestMethod -Method Get -Uri "https://freeipapi.com/api/json/$IPAddress"
                # If successful, break out of the loop
                break
            } catch {
                # Display the error message, wait for 10 seconds before re-trying
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Waiting 10s before re-trying" -ForegroundColor Red
                Start-Sleep -Seconds 10
            }
        }
        # Create a PSCustomObject with geolocation data
        $geoIPData = [PSCustomObject]@{
            IP      = $request.ipAddress
            City    = $request.cityName
            Country = $request.countryName
            IsProxy = $request.isProxy
        }
        # Add the data to the $script:GeoIP_Reference_List
        $script:GeoIP_Reference_List += $geoIPData
        # Return the geolocation data
        return $geoIPData
    }
}
```

### Function to create folders if they do not already exist
```powershell
function New-FolderIfNotPresent {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )
    
    if (-not (Test-Path -Path $FolderPath -PathType Container)) {
        Write-Host "Folder $FolderPath does not exist. Creating folder..." -ForegroundColor Yellow
        New-Item -Path $FolderPath -ItemType Directory | Out-Null
        Write-Host "Folder created successfully." -ForegroundColor blue
    } else {
        Write-Host "Folder already exists." -ForegroundColor blue
    }
}
```

### Function to prompt for user input with a default value
```powershell
function Get-InputWithDefault { # Prompt for input with a default value
    param (
        [string]$DefaultVar, #can change var type as needed
        [string]$Prompt
    )
    Write-Host "$Prompt [Default:$($DefaultVar)]" -ForegroundColor Cyan
    $OutputVar = Read-Host -Prompt ">"
    $OutputVar = ($DefaultVar,$OutputVar)[[bool]$OutputVar]
    return $OutputVar
}
```

### Function to check if a string contains any one of a set of terms to search for 
```powershell
function Find-WordInArray {
    param (
        [string]$InputString,
        [string[]]$WordArray
    )
    # Split the input string into an array of words
    $InputWords = $InputString -split '\s+'
    # Check if any word in the input matches any word in the array
    $Match = $InputWords | Where-Object { $WordArray -contains $_ }
    # Return true if there is a match, otherwise return false
    if ($Match) {
        return $true
    } else {
        return $false
    }
}
```

### Function to get the number of objects in an 
```powershell
function Get-OUObjectCount {
    param (
        [string]$SearchBase
    )
    # Get all OUs recursively starting from the specified search base
    $ous = Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase
    # Iterate through each OU and get the count of each object type
    foreach ($ou in $ous) {
        $ouPath = $ou.DistinguishedName

        # Get the count of each object type in the OU
        $objectCount = Get-ADObject -Filter * -SearchBase $ouPath -Properties ObjectClass |
            Group-Object ObjectClass |
            Select-Object Name, Count
        # Display OU information and object counts
        Write-Output "OU: $ouPath"
        $objectCount | Format-Table -AutoSize
        Write-Output "`n"  # Add a newline for better readability
    }
}
# Example:
Get-OUObjectCount -SearchBase "DC=company,DC=location,DC=local"
```

### Functions for dealing with shortcut/.lnk files:
```powershell
function Get-Shortcut {
    # Gets all start menu shortcuts by default
    param(
    $path = $null
    )
    $obj = New-Object -ComObject WScript.Shell
    if ($path -eq $null) {
    $pathUser = [System.Environment]::GetFolderPath('StartMenu')
    $pathCommon = $obj.SpecialFolders.Item('AllUsersStartMenu')
    $path = dir $pathUser, $pathCommon -Filter *.lnk -Recurse 
    }
    if ($path -is [string]) {
    $path = dir $path -Filter *.lnk
    }
    $path | ForEach-Object { 
    if ($_ -is [string]) {
        $_ = dir $_ -Filter *.lnk
    }
    if ($_) {
        $link = $obj.CreateShortcut($_.FullName)

        $info = @{}
        $info.Hotkey = $link.Hotkey
        $info.TargetPath = $link.TargetPath
        $info.LinkPath = $link.FullName
        $info.Arguments = $link.Arguments
        $info.Target = try {Split-Path $info.TargetPath -Leaf } catch { 'n/a'}
        $info.Link = try { Split-Path $info.LinkPath -Leaf } catch { 'n/a'}
        $info.WindowStyle = $link.WindowStyle
        $info.IconLocation = $link.IconLocation

        New-Object PSObject -Property $info
    }
    }
}
```

```powershell
function Set-Shortcut {
    param(
        [Parameter(ValueFromPipelineByPropertyName=$true)]
        $LinkPath,
        $Hotkey,
        $IconLocation,
        $Arguments,
        $TargetPath
    )
    begin {
        $shell = New-Object -ComObject WScript.Shell
    }
    process {
        $link = $shell.CreateShortcut($LinkPath)
    $PSCmdlet.MyInvocation.BoundParameters.GetEnumerator() |
        Where-Object { $_.key -ne 'LinkPath' } |
        ForEach-Object { $link.$($_.key) = $_.value }
    $link.Save()
    }
}
```
