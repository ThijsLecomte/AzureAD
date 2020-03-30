<#
.SYNOPSIS
    This script syncs trusted named locations from Conditional Access to Corporate IP Ranges in MCAS

.DESCRIPTION
    This script is meant to run in Azure Automatinon and will query 3 three variables/credentials:
        - Credential 'GraphAPI' which has an applicationID as username and applicationSecret as password
            The application should have 'Policy.Read.All' application permissions
        - tenantID
        - Credential 'MCAS' which uses the MCAS URL as username and MCAS API KEY als password.
            URL should the in the following format: 365bythijs.eu2.portal.cloudappsecurity.com

    This script will run once and copy all trusted named locations from Conditional Access to corporate IP ranges in MCAS.
    BE AWARE: if you currently have corporate IP ranges in MCAS, these will all be overwritten.

    This script uses some undocumented MCAS API endpoints.

.EXAMPLE
    Use script to authenticate with O365
    ..\Get-AADLicenseErrors..ps1 -ReportSender 'example@contoso.com' -ReportRecipient 'Example2@contoso.com' -SMTPServer "smtp.office365.com" -SMTPPort 587 -SMTPSSL True

.EXAMPLE
    Change the default logpath with the use of the parameter logPath
    ..\Get-AADLicenseErrors..ps1 -logPath "C:\Windows\Temp\CustomScripts\Get-AADLicenseErrors.txt"

.NOTES
    File Name  : Invoke-NamedLocationsToMCASSyncLocal.ps1  
    Author     : Thijs Lecomte 
    Company    : The Collective Consulting
#>

function Invoke-MCASRestMethod {
    [CmdletBinding()]
    param (
        # Specifies the credential object containing tenant as username (e.g. 'contoso.us.portal.cloudappsecurity.com') and the 64-character hexadecimal Oauth token as the password.
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
                ($_.GetNetworkCredential().username).EndsWith('.portal.cloudappsecurity.com')
            })]
        [ValidateScript( {
                $_.GetNetworkCredential().Password -match ($MCAS_TOKEN_VALIDATION_PATTERN)
            })]
        [System.Management.Automation.PSCredential]$Credential,

        # Specifies the relative path of the full uri being invoked (e.g. - '/api/v1/alerts/')
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
                $_.StartsWith('/')
            })]
        [string]$Path,

        # Specifies the HTTP method to be used for the request
        [Parameter(Mandatory = $true)]
        [ValidateSet('Get', 'Post', 'Put', 'Delete')]
        [string]$Method,

        # Specifies the body of the request, not including MCAS query filters, which should be specified separately in the -FilterSet parameter
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        $Body,

        # Specifies the content type to be used for the request
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$ContentType = 'application/json',

        # Specifies the MCAS query filters to be used, which will be added to the body of the message
        [Parameter(Mandatory = $false)]
        [ValidateNotNull()]
        $FilterSet,

        # Specifies the retry interval, in seconds, if a call to the MCAS web API is throttled. Default = 5 (seconds)
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [int]$RetryInterval = 5,

        # Specifies that a single item is to be fetched, skipping any processing for lists, such as checking result count totals
        #[switch]$Fetch,

        # Specifies use Invoke-WebRequest instead of Invoke-RestMethod, enabling the caller to get the raw response from the MCAS API without any JSON conversion
        [switch]$Raw
    )
    #Ensure TLS 1.2 is used.
    if([Net.ServicePointManager]::SecurityProtocol -notmatch 'Tls12'){
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }

    if ($Raw) {
        $cmd = 'Invoke-WebRequest'
        Write-Verbose "-Raw parameter was specified"
    }
    else {
        $cmd = 'Invoke-RestMethod'
        Write-Verbose "-Raw parameter was not specified"
    }
    Write-Verbose "$cmd will be used"

    $tenant = ($Credential.GetNetworkCredential().username)
    Write-Verbose "Tenant name is $tenant"

    Write-Verbose "Relative path is $Path"

    Write-Verbose "Method is $Method"

    $token = $Credential.GetNetworkCredential().Password
    #MK - Commenting out this line for security reasons. Not sure I like having the raw token in the verbose output.
    #Write-Verbose "OAuth token is $token"

    $headers = 'Authorization = "Token {0}"' -f $token | ForEach-Object {
        "@{$_}"
    }
    Write-Verbose "Request headers are $headers"

    # Construct base MCAS call before processing -Body and -FilterSet
    $mcasCall = '{0} -Uri ''https://{1}{2}'' -Method {3} -Headers {4} -ContentType {5} -UseBasicParsing' -f $cmd, $tenant, $Path, $Method, $headers, $ContentType

    if ($Method -eq 'Get') {
        Write-Verbose "A request using the Get HTTP method cannot have a message body."
    }
    else {
        $jsonBody = $Body | ConvertTo-Json -Compress -Depth 4
        Write-Verbose "Base request body is $jsonBody"

        if ($FilterSet) {
            Write-Verbose "Request body before query filters is $jsonBody"
            $jsonBody = $jsonBody.TrimEnd('}') + ',' + '"filters":{' + ((ConvertTo-MCASJsonFilterString $FilterSet).TrimStart('{')) + '}'
            Write-Verbose "Request body after query filters is $jsonBody"
        }
        else {
            Write-Verbose "No filters were added to the request body"
        }
        Write-Verbose "Final request body is $jsonBody"

        # Add -Body to the constructed MCAS call, when the http method is not 'Get'
        $mcasCall = '{0} -Body ''{1}''' -f $mcasCall, $jsonBody
    }

    Write-Verbose "Constructed call to MCAS is to follow:"
    $mcasCall2 = '{0} -Uri ''https://{1}{2}'' -Method {3} -ContentType {5} -UseBasicParsing' -f $cmd, $tenant, $Path, $Method, $headers, $ContentType

    Write-Verbose $mcasCall2

    Write-Verbose "Retry interval if MCAS call is throttled is $RetryInterval seconds"

    # This loop is the actual call to MCAS. It includes automatic retry if the API call is throttled
    do {
        $retryCall = $false

        try {
            Write-Verbose "Attempting call to MCAS..."
            $response = Invoke-Expression -Command $mcasCall
        }
        catch {
            if ($_ -like 'The remote server returned an error: (429) TOO MANY REQUESTS.') {
                Write-Warning "429 - Too many requests. The MCAS API throttling limit has been hit, the call will be retried in $RetryInterval second(s)..."
                $retryCall = $true
                Write-Verbose "Sleeping for $RetryInterval seconds"
                Start-Sleep -Seconds $RetryInterval
            }
            ElseIf ($_ -match 'throttled') {
                Write-Warning "Too many requests. Usually the throttle time for this call is 1 minute. Next request will resume in 1 minute..."
                $retryCall = $true
                Write-Verbose "Sleeping for 60 seconds"
                Start-Sleep -Seconds 60
            }
            ElseIf ($_ -like 'The remote server returned an error: (504)') {
                Write-Warning "504 - Gateway Timeout. The call will be retried in $RetryInterval second(s)..."
                $retryCall = $true
                Write-Verbose "Sleeping for $RetryInterval seconds"
                Start-Sleep -Seconds $RetryInterval
            }
            else {
                throw $_
            }
        }

        # Uncomment following two lines if you want to see raw responses in -Verbose output
        #Write-Verbose 'MCAS response to follow:'
        #Write-Verbose $response
    }
    while ($retryCall)

    # Provide the total record count in -Verbose output and as InformationVariable, if appropriate
    if (@('Get', 'Post') -contains $Method) {
        if ($response.total) {
            Write-Verbose 'Checking total matching record count via the response properties...'
            $recordTotal = $response.total
        }
        elseif ($response.Content) {
            try {
                Write-Verbose 'Checking total matching record count via raw JSON response...'
                $recordTotal = (($response.content).Replace('"Level":','"Level_2":') | ConvertFrom-Json).total
            }
            catch {
                Write-Verbose 'JSON conversion failed. Checking total matching record count via raw response string extraction...'
                #below linew as commented out as it breaks with the new activities_kusto endpoint.
                #$recordTotal = ($response.Content.Split(',', 3) | Where-Object {$_.StartsWith('"total"')} | Select-Object -First 1).Split(':')[1]
            }
        }
        else {
            Write-Verbose 'Could not check total matching record count, perhaps because zero or one records were returned. Zero will be returned as the matching record count.'
            $recordTotal = 0
        }

        Write-Verbose ('The total number of matching records was {0}' -f $recordTotal)
        #removing the below line because it is now breaking certain cmdlets such as Get-MCASFile when retriving a file by identity
        #Write-Information $recordTotal
    }
    $response
}

#Function to update MCAS IP Range
Function Update-MCASIPRange(){
    [CmdletBinding()]
    param (
        [switch]$Add,
        [switch]$Remove,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $MCASIPRange,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $subnet,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $credential
    )

    $IPRanges = @()

    if($Add){
        Write-Verbose "Adding $($subnet) to IPRange $($MCASIPRange.Name)" -Verbose

        $IPRanges = $MCASIPRange.subnets.OriginalString + $subnet
    }
    elseif($Remove){
        Write-Verbose "Removing $($subnet.originalString) from IPRange $($MCASIPRange.Name)" -Verbose

        #as subnets is an array, we cannot remove one element easily. So we need to recreate the array
        $newSubnets = @()
        foreach ($MCASsubnet in $MCASIPRange.Subnets)
        {
            if ($MCASsubnet.OriginalString -ne $subnet.OriginalString)
            {
                $newSubnets += $MCASsubnet.originalString
            }
        }

        $IPRanges = $newSubnets
    }

    Write-Verbose "Updating to IPranges $IPRanges"

    $body=@{
        "name"=$MCASIPRange.Name
        "category"=1
        "subnets"=$IPRanges
    }
    
    try{
        Write-Verbose "Updating MCASIPRange $($MCASIPRange.Name)" -Verbose
        $updatedRange = Invoke-MCASRestMethod -Credential $credential -Path "/api/v1/subnet/$($MCASIPRange._id)/update_rule/" -Method Post -Body $body
    }
    catch{
        Write-Error "Error updating MCAS IP Range"
        throw "Error calling MCAS API. The exception was: $_"
    }

    $id = $MCASIPRange._id
    
    $updatedRange = $updatedRange | Add-Member -MemberType NoteProperty -Name '_id' -Value $id -PassThru

    return $updatedRange
}

#Function to create new MCAS IP Ranges
Function Create-MCASIPRange{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DisplayName,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $IPRanges,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $credential
    )
    Write-Verbose "Creating MCAS IP Range. Name: $DisplayName, subnets $IPRanges" -Verbose

    $body=@{
        "name"=$DisplayName
        "category"=1
        "subnets"=$IPRanges
    }

    try{
        $data = Invoke-MCASRestMethod -Credential $credential -Path '/api/v1/subnet/create_rule/' -Method Post -Body $body
        Write-Verbose "Created MCAS IP Range with ID $data" -Verbose
    }
    catch{
        Write-Error "Error creating MCAS IP Range"
        throw "Error calling MCAS API. The exception was: $_"
    }
}

#Function to remove an MCAS IP Range through id
Function Remove-MCASIPRange{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$id,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $credential
    )

    Write-Verbose "Removing MCAS IP Range $id" -Verbose
    
    try{
        Invoke-MCASRestMethod -Credential $credential -Path "/api/v1/subnet/$id/" -Method Delete
        Write-Verbose "Succesfully deleted MCAS IP Range $id" -Verbose
    }
    catch{
        Write-Error "Error removing MCAS IP Range"
        throw "Error calling MCAS API. The exception was: $_"
    }
    
}


#This function will retrieve all trusted named locations from AAD
Function Get-NamedLocations(){
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $clientsecret,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$clientid,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$tenantId
    )

    $body=@{
        client_id=$clientid
        client_secret=$clientsecret
        scope="https://graph.microsoft.com/.default"
        grant_type="client_credentials"
    }

    $accesstoken = (Invoke-WebRequest -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body -Method Post -UseBasicParsing).content | ConvertFrom-Json

    $authHeader = @{
        'Content-Type'='application/json'
        'Authorization'="Bearer " + $accessToken.access_token
        'ExpiresOn'=$accessToken.expires_in
    }

    $uri = "https://graph.microsoft.com/beta/conditionalAccess/namedLocations?`$select=displayName,microsoft.graph.ipNamedLocation/ipRanges/&`$filter=microsoft.graph.ipNamedLocation/isTrusted"

    $namedlocations = @()
    do{
        try {
            Write-Verbose "Querying AAD for named locations" -Verbose
            $data = (Invoke-RestMethod -Uri $uri -Headers $authHeader -Method Get –UseBasicParsing)
            $namedlocations += $data.Value

            $uri = $data.'@odata.nextLink'
        }

        catch {
            Write-Error "Error getting named locations from AAD"
            Write-Error "$($_.Exception.Message)"
        }
    }
    while($uri)

    Write-Verbose "Found $($namedlocations.count) named locations" -Verbose
    return $namedlocations
}

#This function will retrieve all MCAS corporate IP-addresses
Function Get-MCASCorpIPRanges{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $credential
    )
    $ipranges = @()
    do{
        try {
            Write-Verbose "Querying MCAS for IP Ranges" -Verbose
            $filterSet = @(@{'category'=@{'eq'=1}})
            $data = Invoke-MCASRestMethod -Credential $credential -Path '/api/v1/subnet/' -Method Get -FilterSet $filterSet
            $ipranges += $data.data
        }

        catch {
            Write-Error "Error getting MCAS IP Ranges"
            Write-Error "$($_.Exception.Message)"
        }
    }
    while($data.hasNext)

    Write-Verbose "Found $($ipranges.count) corporate IP ranges in MCAS" -Verbose
    return $ipranges
}

$GraphCreds = Get-AutomationPSCredential -Name 'GraphAPI'

$namedlocations = Get-NamedLocations -clientsecret $GraphCreds.GetNetworkCredential().password -clientid $GraphCreds.username -tenantId (Get-AutomationVariable -Name 'tenantID')

[Array]$ipranges = Get-MCASCorpIPRanges -credential (Get-AutomationPSCredential -Name 'MCAS')

#Loop over named locations
foreach($named in $namedlocations){
    Write-Verbose "Checking named location $($named.displayName)" -Verbose

    #Check if named location currently exists in MCAS
    if($ipranges.count -ne 0 -and $ipranges.Name.contains($named.displayName)){
        Write-Verbose "Named location currently exists in MCAS" -Verbose

        $MCASIPRange = $ipranges.Get($ipranges.Name.IndexOf($named.displayName))
        
        Write-Verbose "Start looping over IP ranges" -Verbose
        foreach($iprange in $named.ipranges){
            Write-Verbose "Checking IPrange $($ipRange.cidrAddress)" -Verbose

            if(!$MCASIPRange.Subnets.OriginalString.Contains($iprange.cidrAddress)){
                Write-Verbose "IPrange doesnt exist in MCAS, creating" -Verbose
                
                $MCASIPRange = Update-MCASIPRange -MCASIPRange $MCASIPRange -Subnet $iprange.cidrAddress -Add -credential (Get-AutomationPSCredential -Name 'MCAS')
            }
        }

        Write-Verbose "Check MCAS IP Ranges to see if they exist in AAD." -Verbose
        #Loop MCAS IP Ranges to check it they exist in AAD
        foreach($MCASsubnet in $MCASIPRange.subnets){
            if(!$named.ipranges.cidrAddress.contains($MCASsubnet.originalString)){
                Write-Verbose "MCAS range $($MCASsubnet.originalString) doesn't exist in AAD, removing" -Verbose

                $MCASIPRange = Update-MCASIPRange -MCASIPRange $MCASIPRange -Subnet $MCASsubnet -Remove -credential (Get-AutomationPSCredential -Name 'MCAS')
            }
        }
    }
    else {
        #named location doesn't exist in MCAS - Create ip range in MCAS
        Write-Verbose "Named location $($named.displayName) doesn't exist in MCAS, creating" -Verbose
        Create-MCASIPRange -DisplayName $named.displayName -IPRanges $named.ipRanges.cidrAddress  -credential (Get-AutomationPSCredential -Name 'MCAS')
    }
}

Write-Verbose "Checking if there are any MCAS Corp IP Ranges that don't exist in AAD"
foreach($iprange in $ipranges){
    Write-Verbose "Checking $($iprange.name)"

    if(!$namedlocations.DisplayName.Contains($iprange.Name)){
        Write-Verbose "$($iprange.name) not found in AAD, removing"

        Remove-MCASIPRange -id $iprange._id -credential (Get-AutomationPSCredential -Name 'MCAS')
    }
}