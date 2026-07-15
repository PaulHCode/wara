<#
.SYNOPSIS
    Invokes an Azure Resource Graph query.

.DESCRIPTION
    The `Invoke-WAFQuery` function executes an Azure Resource Graph query and returns the results. It handles pagination and consolidates results from multiple subscriptions if provided.

.PARAMETER Query
    The Kusto query string to execute against Azure Resource Graph.

.PARAMETER SubscriptionId
    An array of subscription IDs to scope the query to.

.INPUTS
    System.String. The query string.
    System.String[]. The array of subscription IDs.

.OUTPUTS
    System.Object[]. Returns an array of query results.

.EXAMPLE
    PS> $query = "Resources | where type =~ 'Microsoft.Compute/virtualMachines'"
    PS> $results = Invoke-WAFQuery -Query $query -SubscriptionId @("59f6f1ab-6d68-4c90-b4e5-ad2d71cefc57")

    This example retrieves all virtual machines within the specified subscription.

.EXAMPLE
    PS> $results = Invoke-WAFQuery -Query $query -SubscriptionId $subscriptionIds

    This example executes the query across multiple subscriptions.

.NOTES
    Author: Kyle Poineal
    Date: [Today's Date]
#>
function Invoke-WAFQuery {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]] $SubscriptionIds,

        [Parameter(Mandatory = $false)]
        [string] $Query = 'resources | project name, type, location, resourceGroup, subscriptionId, id'
    )

    $result = $SubscriptionIds ? (Search-AzGraph -Query $Query -First 1000 -Subscription $SubscriptionIds) : (Search-AzGraph -Query $Query -First 1000 -UseTenantScope) # -first 1000 returns the first 1000 results and subsequently reduces the amount of queries required to get data.

    # Collection to store all resources
    $allResources = @($result)

    # Loop to paginate through the results using the skip token
    $result = while ($result.SkipToken) {
        # Retrieve the next set of results using the skip token
        $result = $SubscriptionIds ? (Search-AzGraph -Query $Query -SkipToken $result.SkipToken -Subscription $SubscriptionIds -First 1000) : (Search-AzGraph -Query $Query -SkipToken $result.SkipToken -First 1000 -UseTenantScope)
        # Add the results to the collection
        Write-Output $result
    }

    $allResources += $result

    # Output all resources
    return , $allResources
}

<#
.SYNOPSIS
    Invokes an Azure REST API then returns the response.

.DESCRIPTION
    The Invoke-AzureRestApi function invokes an Azure REST API with the specified parameters then return the response.

.PARAMETER Method
    The HTTP method to invoke the Azure REST API. The accepted values are GET, POST, PUT, PATCH, and DELETE.

.PARAMETER SubscriptionId
    The subscription ID that constitutes the URI for invoke the Azure REST API.

.PARAMETER ResourceGroupName
    The resource group name that constitutes the URI for invoke the Azure REST API.

.PARAMETER ResourceProviderName
    The resource provider name that constitutes the URI for invoke the Azure REST API. It's usually as the XXXX.XXXX format.

.PARAMETER ResourceType
    The resource type that constitutes the URI for invoke the Azure REST API.

.PARAMETER Name
    The resource name that constitutes the URI for invoke the Azure REST API.

.PARAMETER ApiVersion
    The Azure REST API version that constitutes the URI for invoke the Azure REST API. It's usually as the yyyy-mm-dd format.

.PARAMETER QueryString
    The query string that constitutes the URI for invoke the Azure REST API.

.PARAMETER RequestBody
    The request body for invoke the Azure REST API.

.OUTPUTS
    Returns a REST API response as the PSHttpResponse.

.EXAMPLE
    PS> $response = Invoke-AzureRestApi -Method 'GET' -SubscriptionId '11111111-1111-1111-1111-111111111111' -ResourceProviderName 'Microsoft.ResourceHealth' -ResourceType 'events' -ApiVersion '2024-02-01' -QueryString 'queryStartTime=2024-10-02T00:00:00'

.NOTES
    Author: Takeshi Katano
    Date: 2024-10-23

    This function requires the Az.Accounts module to be installed and imported.
#>
function Invoke-AzureRestApi {
    [CmdletBinding()]
    [OutputType([Microsoft.Azure.Commands.Profile.Models.PSHttpResponse])]
    param (
        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')]
        [string] $Method,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $true)]
        [ValidateScript({ Test-WAFIsGuid -StringGuid $_ })]
        [string] $SubscriptionId,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [ValidateLength(1, 90)]
        [string] $ResourceGroupName,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $true)]
        [string] $ResourceProviderName,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $true)]
        [string] $ResourceType,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [string] $Name,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $true)]
        [ValidatePattern('^[0-9]{4}(-[0-9]{2}){2}$')]
        [string] $ApiVersion,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $false)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $false)]
        [string] $QueryString,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $false)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $false)]
        [string] $RequestBody
    )

    # Built the Azure REST API URI path.
    $cmdletParams = @{
        SubscriptionId       = $SubscriptionId
        ResourceProviderName = $ResourceProviderName
        ResourceType         = $ResourceType
        ApiVersion           = $ApiVersion
    }
    if ($PSBoundParameters.ContainsKey('ResourceGroupName')) { $cmdletParams.ResourceGroupName = $ResourceGroupName }
    if ($PSBoundParameters.ContainsKey('Name')) { $cmdletParams.Name = $Name }
    if ($PSBoundParameters.ContainsKey('QueryString')) { $cmdletParams.QueryString = $QueryString }
    $path = Get-AzureRestMethodUriPath @cmdletParams

    # Invoke the Azure REST API using the URI path.
    $cmdletParams = @{
        Method = $Method
        Path   = $path
    }
    if ($PSBoundParameters.ContainsKey('RequestBody')) { $cmdletParams.Payload = $RequestBody }
    return Invoke-AzRestMethod @cmdletParams
}

<#
.SYNOPSIS
    Retrieves the path of the Azure REST API URI.

.DESCRIPTION
    The Get-AzureRestMethodUriPath function retrieves the formatted path of the Azure REST API URI based on the specified URI parts as parameters.
    The path represents the Azure REST API URI without the protocol (e.g. https), host (e.g. management.azure.com). For example,
    /subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/rg1/providers/Microsoft.Storage/storageAccounts/stsample1234?api-version=2024-01-01

.PARAMETER SubscriptionId
    The subscription ID that constitutes the path of Azure REST API URI.

.PARAMETER ResourceGroupName
    The resource group name that constitutes the path of Azure REST API URI.

.PARAMETER ResourceProviderName
    The resource provider name that constitutes the path of Azure REST API URI. It's usually as the XXXX.XXXX format.

.PARAMETER ResourceType
    The resource type that constitutes the path of Azure REST API URI.

.PARAMETER Name
    The resource name that constitutes the path of Azure REST API URI.

.PARAMETER ApiVersion
    The Azure REST API version that constitutes the path of Azure REST API URI. It's usually as the yyyy-mm-dd format.

.PARAMETER QueryString
    The query string that constitutes the path of Azure REST API URI.

.OUTPUTS
    Returns a URI path to call Azure REST API.

.EXAMPLE
    PS> $path = Get-AzureRestMethodUriPath -SubscriptionId '11111111-1111-1111-1111-111111111111' -ResourceGroupName 'rg1' -ResourceProviderName 'Microsoft.Storage' -ResourceType 'storageAccounts' -Name 'stsample1234' -ApiVersion '2024-01-01' -QueryString 'param1=value1'

.NOTES
    Author: Takeshi Katano
    Date: 2024-10-23
#>
function Get-AzureRestMethodUriPath {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $true)]
        [ValidateScript({ Test-WAFIsGuid -StringGuid $_ })]
        [string] $SubscriptionId,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [ValidateLength(1, 90)]
        [string] $ResourceGroupName,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $true)]
        [string] $ResourceProviderName,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $true)]
        [string] $ResourceType,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [string] $Name,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $true)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $true)]
        [ValidatePattern('^[0-9]{4}(-[0-9]{2}){2}$')]
        [string] $ApiVersion,

        [Parameter(ParameterSetName = 'WithResourceGroup', Mandatory = $false)]
        [Parameter(ParameterSetName = 'WithoutResourceGroup', Mandatory = $false)]
        [string] $QueryString
    )

    $additionalQueryString = if ($PSBoundParameters.ContainsKey('QueryString')) { '&' + $QueryString } else { '' }
    $path = if ($PSCmdlet.ParameterSetName -eq 'WithResourceGroup') {
        '/subscriptions/{0}/resourcegroups/{1}/providers/{2}/{3}/{4}?api-version={5}{6}' -f $SubscriptionId, $ResourceGroupName, $ResourceProviderName, $ResourceType, $Name, $ApiVersion, $additionalQueryString
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'WithoutResourceGroup') {
        '/subscriptions/{0}/providers/{1}/{2}?api-version={3}{4}' -f $SubscriptionId, $ResourceProviderName, $ResourceType, $ApiVersion, $additionalQueryString
    }
    else {
        throw "The parameter set name [$($PSCmdlet.ParameterSetName)] is invalid."
    }
    return $path
}

<#
.SYNOPSIS
    Imports configuration data from a file.

.DESCRIPTION
    The `Import-WAFConfigFileData` function reads the content of a configuration file, extracts sections, and returns the data as a `PSCustomObject`. The configuration file should have sections defined by square brackets `[SectionName]` and key-value pairs within each section.

.PARAMETER ConfigFile
    The path to the configuration file.

.INPUTS
    System.String. The function accepts a string representing the path to the configuration file.

.OUTPUTS
    System.Management.Automation.PSCustomObject. Returns a custom object containing the configuration data.

.EXAMPLE
    PS> $configData = Import-WAFConfigFileData -ConfigFile "C:\config\settings.txt"

    This example imports configuration data from the specified file.

.EXAMPLE
    PS> Import-WAFConfigFileData -ConfigFile "config.txt"

    This example imports configuration data from 'config.txt' in the current directory.

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
function Import-WAFConfigFileData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [string] $ConfigFile
    )

    # Read the file content and store it in a variable
    $filecontent, $linetable, $objarray, $count, $start, $stop, $configsection = $null
    $filepath = (Resolve-Path -Path $configfile).Path
    $filecontent = (Get-content $filepath).trim().tolower()

    # Create an array to store the line number of each section
    $linetable = @()
    $objarray = [ordered]@{}

    $filecontent = $filecontent | Where-Object { $_ -ne '' -and $_ -notlike '*#*' }

    #Remove empty space.
    foreach ($line in $filecontent) {
        $index = $filecontent.IndexOf($line)
        if ($line -match '^\[([^\]]+)\]$' -and ($filecontent[$index + 1] -match '^\[([^\]]+)\]$' -or [string]::IsNullOrEmpty($filecontent[$index + 1]))) {
            # Set this line to empty because the next line is a section as well.
            # This is to avoid the section name being added to the object since it has no parameters.
            # This is because if we were to keep the note-property it would mess up logic for determining if a section is empty.
            # Powershell will return $true on an emtpy note property - Because the property exists.
            $filecontent[$index] = ''
        }
    }

    #Remove empty space again.
    $filecontent = $filecontent | Where-Object { $_ -ne '' -and $_ -notlike '*#*' }

    # Iterate through the file content and store the line number of each section
    foreach ($line in $filecontent) {
        if (-not [string]::IsNullOrWhiteSpace($line) -and -not $line.startswith('#')) {
            #Get the Index of the current line
            $index = $filecontent.IndexOf($line)
            # If the line is a section, store the line number
            if ($line -match '^\[([^\]]+)\]$') {
                # Store the section name and line number. Remove the brackets from the section name
                $linetable += $filecontent.indexof($line)
            }
        }
    }

    # Iterate through the line numbers and extract the section content
    $count = 0
    foreach ($entry in $linetable) {

        # Get the section name
        $name = $filecontent[$entry]
        # Remove the brackets from the section name
        $name = $name.replace('[', '').replace(']', '')

        # Get the start and stop line numbers for the section content
        # If the section is the last one, set the stop line number to the end of the file
        $start = $entry + 1

        if ($linetable.count -eq $count + 1) {
            $stop = $filecontent.count - 1
        }
        else {
            $stop = $linetable[$count + 1] - 1
        }

        # Extract the section content
        $configsection = $filecontent[$start..$stop]

        # Add the section content to the object array
        $objarray += @{$name = $configsection }

        # Increment the count
        $count++
    }

    # Return the object array and cast to PSCustomObject
    return [PSCustomObject]$objarray
}

<#
.SYNOPSIS
    Connects to an Azure tenant.

.DESCRIPTION
    The Connect-WAFAzure function connects to an Azure tenant using the provided Tenant ID and Subscription IDs.

.PARAMETER TenantID
    The Tenant ID to connect to.

.PARAMETER SubscriptionIds
    An array of Subscription IDs to scope the connection.

.PARAMETER AzureEnvironment
    The Azure environment to connect to. Defaults to 'AzureCloud'.

.OUTPUTS
    None.

.EXAMPLE
    PS> Connect-WAFAzure -TenantID "your-tenant-id" -SubscriptionIds @("sub1", "sub2") -AzureEnvironment "AzureCloud"
#>
function Connect-WAFAzure {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [GUID] $TenantID,

        [Parameter(Mandatory = $false)]
        [string] $AzureEnvironment = 'AzureCloud'
    )

    # Connect To Azure Tenant
    if ((Get-AzContext).Tenant.Id -ne $TenantID -or (Get-AzContext).Environment.Name -ne $AzureEnvironment) {
        Write-Debug "Connecting to Azure Tenant with Tenant ID: $TenantID and Azure Environment: $AzureEnvironment"
        Connect-AzAccount -Tenant $TenantID -WarningAction SilentlyContinue -Environment $AzureEnvironment | Out-Null
    }
}

<#
.SYNOPSIS
    Validates an array of tag patterns.

.DESCRIPTION
    The `Test-WAFTagPattern` function checks if each tag pattern in the input array follows the required format. Tags should be specified in the format 'Key!~Value||Key2!~Value2'.

.PARAMETER InputValue
    An array of tag patterns to validate.

.INPUTS
    System.String[]. The function accepts an array of tag pattern strings.

.OUTPUTS
    None. Throws an error if validation fails.

.EXAMPLE
    PS> Test-WAFTagPattern -InputValue @("Env!~Prod||Test", "Owner!~JohnDoe")

    This example validates valid tag patterns.

.EXAMPLE
    PS> Test-WAFTagPattern -InputValue @("InvalidTagPattern")

    Error:
    The tag pattern 'InvalidTagPattern' is invalid.

    This example demonstrates validation failure for an invalid tag pattern.

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
function Test-WAFTagPattern {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $InputValue
    )

    $pattern = '^[^<>&%\\?/]+=~[^<>&%\\?/]+$|[^<>&%\\?/]+!~[^<>&%\\?/]+$'

    $allMatch = $true
    foreach ($value in $InputValue) {
        if ($value -notmatch $pattern) {
            $allMatch = $false
            throw "Tag pattern [$value] is not valid."
            break
        }
    }
    return $allMatch
}

<#
.SYNOPSIS
    Validates an array of resource group IDs.

.DESCRIPTION
    The `Test-WAFResourceGroupId` function checks if each resource group ID in the input array follows the correct Azure resource group ID format.

.PARAMETER InputValue
    An array of resource group IDs to validate.

.INPUTS
    System.String[]. The function accepts an array of resource group ID strings.

.OUTPUTS
    None. Throws an error if validation fails.

.EXAMPLE
    PS> Test-WAFResourceGroupId -InputValue @("/subscriptions/59f6f1ab-6d68-4c90-b4e5-ad2d71cefc57/resourceGroups/MyResourceGroup")

    This example validates a valid resource group ID.

.EXAMPLE
    PS> Test-WAFResourceGroupId -InputValue @("invalid-resource-group-id")

    Error:
    The resource group ID 'invalid-resource-group-id' is invalid.

    This example demonstrates validation failure when an invalid resource group ID is provided.

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
function Test-WAFResourceGroupId {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $InputValue
    )

    $pattern = '\/subscriptions\/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\/resourceGroups\/[a-zA-Z0-9._-]+'

    $allMatch = $true
    foreach ($value in $InputValue) {
        if ($value -notmatch $pattern) {
            $allMatch = $false
            throw "Resource Group ID [$value] is not valid."
            break
        }
    }
    return $allMatch
}

<#
.SYNOPSIS
    Validates an array of subscription IDs.

.DESCRIPTION
    The `Test-WAFSubscriptionId` function checks if each subscription ID in the input array is a valid GUID format. It throws an error if any subscription ID is invalid.

.PARAMETER InputValue
    An array of subscription IDs to validate.

.INPUTS
    System.String[]. The function accepts an array of subscription ID strings.

.OUTPUTS
    None. Throws an error if validation fails.

.EXAMPLE
    PS> Test-WAFSubscriptionId -InputValue @("59f6f1ab-6d68-4c90-b4e5-ad2d71cefc57", "invalid-guid")

    Error:
    The subscription ID 'invalid-guid' is not a valid GUID.

    This example demonstrates validation failure when an invalid subscription ID is provided.

.EXAMPLE
    PS> Test-WAFSubscriptionId -InputValue @("59f6f1ab-6d68-4c90-b4e5-ad2d71cefc57")

    This example validates a valid subscription ID without any error.

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
function Test-WAFSubscriptionId {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $InputValue
    )

    $pattern = '^(\/subscriptions\/)?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\/?$'

    $allMatch = $true
    foreach ($value in $InputValue) {
        if ($value -notmatch $pattern) {
            $allMatch = $false
            throw "Subscription ID [$value] is not valid."
            break
        }
    }
    return $allMatch
}

<#
.SYNOPSIS
    Validates whether a string is a valid GUID.

.DESCRIPTION
    The `Test-WAFIsGuid` function checks if the input string is a valid GUID format.

.PARAMETER StringGuid
    The string to validate as a GUID.

.INPUTS
    System.String. The function accepts a string representing the GUID to validate.

.OUTPUTS
    System.Boolean. Returns `$true` if the input is a valid GUID, `$false` otherwise.

.EXAMPLE
    Test-WAFIsGuid -StringGuid "59f6f1ab-6d68-4c90-b4e5-ad2d71cefc57"

    Output:
    True

    This example checks if the provided string is a valid GUID.

.EXAMPLE
    Test-WAFIsGuid -StringGuid "invalid-guid"

    Output:
    False

    This example demonstrates that an invalid GUID returns `$false`.

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
function Test-WAFIsGuid {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $StringGuid
    )

    $ObjectGuid = [System.Guid]::Empty
    if (-not [System.Guid]::TryParse($StringGuid, [ref]$ObjectGuid)) {
        throw "The provided string [$StringGuid] is not a valid GUID."
    }
    return $true
}

<#
.SYNOPSIS
    Validates that the specified file exists.

.DESCRIPTION
    The `Test-FileExists` function checks if the specified file exists. If the file does not exist, the function throws an error.

.PARAMETER Path
    The path to the file to validate.

.OUTPUTS
    System.Boolean. Returns `$true` if the file exists, otherwise throws an error.

.EXAMPLE
    Test-FileExists -Path ".\this_file_exists.txt"

    Output:
    True

    This example demonstrates that the function returns `$true` when the specified file exists.

.EXAMPLE
    Test-FileExists -Path ".\this_file_does_not_exist.txt"

    Error:
    File [.\this_file_does_not_exist.txt] not found.

    This example demonstrates that the function throws an error when the specified file does not exist.

.NOTES
    Author: Casey Watson
    Date: 2025-02-04
#>
function Test-FileExists {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    if (-not (Test-Path -Path $Path -PathType Leaf -ErrorAction SilentlyContinue)) {
        throw "File [$Path] not found."
    }

    return $true
}

<#
    .SYNOPSIS
        Ensures that subscription IDs are in the correct ARM resource ID format by adding "/subscriptions/" prefix if missing.

    .DESCRIPTION
        The `Repair-WAFSubscriptionId` function accepts an array of subscription IDs and checks each one to ensure it follows the Azure Resource Manager (ARM) resource ID format. If a subscription ID does not start with "/subscriptions/", the function prefixes it with "/subscriptions/". This standardizes the subscription IDs for consistent use in ARM queries and operations.

    .PARAMETER SubscriptionIds
        An array of subscription ID strings to validate and correct if necessary.

    .INPUTS
        System.String[]. You can pipe an array of subscription ID strings to this function.

    .OUTPUTS
        System.String[]. Returns an array of subscription IDs, each starting with "/subscriptions/".

    .EXAMPLE
        PS> $subs = @("59f6f1ab-6d68-4c90-b4e5-ad2d71cefc57", "/subscriptions/abcd1234-5678-90ab-cdef-1234567890ab")
        PS> $fixedSubs = Repair-WAFSubscriptionId -SubscriptionIds $subs
        PS> $fixedSubs

        Output:
        /subscriptions/59f6f1ab-6d68-4c90-b4e5-ad2d71cefc57
        /subscriptions/abcd1234-5678-90ab-cdef-1234567890ab

        This example demonstrates that the function adds the "/subscriptions/" prefix to a subscription ID that lacks it and leaves properly formatted IDs unchanged.

    .EXAMPLE
        PS> $subs = @()
        PS> $fixedSubs = Repair-WAFSubscriptionId -SubscriptionIds $subs

        This example shows that the function correctly handles an empty array without errors, returning an empty array.

    .EXAMPLE
        PS> $subs = @("invalid-guid", "12345678-1234-1234-1234-1234567890ab")
        PS> $fixedSubs = Repair-WAFSubscriptionId -SubscriptionIds $subs
        PS> $fixedSubs

        Output:
        /subscriptions/invalid-guid
        /subscriptions/12345678-1234-1234-1234-1234567890ab

        This example illustrates that the function does not validate the format of the GUID itself; it only ensures the prefix is present.

    .NOTES
        Author: Kyle Poineal
        Date: 2024-12-12
    #>
function Repair-WAFSubscriptionId {
    [CmdletBinding()]
    [Parameter(Mandatory = $true)]
    [AllowEmptyCollection()]
    param (
        [string[]] $SubscriptionIds
    )

    $fixedSubscriptionIds = @()
    foreach ($subscriptionId in $SubscriptionIds) {
        if ($subscriptionId -notmatch '\/subscriptions\/') {
            $fixedSubscriptionIds += "/subscriptions/$subscriptionId"
        }
        else {
            $fixedSubscriptionIds += $subscriptionId
        }
    }
    return $fixedSubscriptionIds
}

# SIG # Begin signature block
# MIIoSAYJKoZIhvcNAQcCoIIoOTCCKDUCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCaIiwOYVFg9QmP
# EiAbVtXnPEeFXG0IJREEs1tbUN9dqqCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
# Bv9XKydyAAAAAAQEMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjQwOTEyMjAxMTE0WhcNMjUwOTExMjAxMTE0WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC0KDfaY50MDqsEGdlIzDHBd6CqIMRQWW9Af1LHDDTuFjfDsvna0nEuDSYJmNyz
# NB10jpbg0lhvkT1AzfX2TLITSXwS8D+mBzGCWMM/wTpciWBV/pbjSazbzoKvRrNo
# DV/u9omOM2Eawyo5JJJdNkM2d8qzkQ0bRuRd4HarmGunSouyb9NY7egWN5E5lUc3
# a2AROzAdHdYpObpCOdeAY2P5XqtJkk79aROpzw16wCjdSn8qMzCBzR7rvH2WVkvF
# HLIxZQET1yhPb6lRmpgBQNnzidHV2Ocxjc8wNiIDzgbDkmlx54QPfw7RwQi8p1fy
# 4byhBrTjv568x8NGv3gwb0RbAgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQU8huhNbETDU+ZWllL4DNMPCijEU4w
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzUwMjkyMzAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAIjmD9IpQVvfB1QehvpC
# Ge7QeTQkKQ7j3bmDMjwSqFL4ri6ae9IFTdpywn5smmtSIyKYDn3/nHtaEn0X1NBj
# L5oP0BjAy1sqxD+uy35B+V8wv5GrxhMDJP8l2QjLtH/UglSTIhLqyt8bUAqVfyfp
# h4COMRvwwjTvChtCnUXXACuCXYHWalOoc0OU2oGN+mPJIJJxaNQc1sjBsMbGIWv3
# cmgSHkCEmrMv7yaidpePt6V+yPMik+eXw3IfZ5eNOiNgL1rZzgSJfTnvUqiaEQ0X
# dG1HbkDv9fv6CTq6m4Ty3IzLiwGSXYxRIXTxT4TYs5VxHy2uFjFXWVSL0J2ARTYL
# E4Oyl1wXDF1PX4bxg1yDMfKPHcE1Ijic5lx1KdK1SkaEJdto4hd++05J9Bf9TAmi
# u6EK6C9Oe5vRadroJCK26uCUI4zIjL/qG7mswW+qT0CW0gnR9JHkXCWNbo8ccMk1
# sJatmRoSAifbgzaYbUz8+lv+IXy5GFuAmLnNbGjacB3IMGpa+lbFgih57/fIhamq
# 5VhxgaEmn/UjWyr+cPiAFWuTVIpfsOjbEAww75wURNM1Imp9NJKye1O24EspEHmb
# DmqCUcq7NqkOKIG4PVm3hDDED/WQpzJDkvu4FrIbvyTGVU01vKsg4UfcdiZ0fQ+/
# V0hf8yrtq9CkB8iIuk5bBxuPMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCGigwghokAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIPuTsBEuIB5T6XGXQMmlDuMc
# fxK4xFfrT1wFSo1xacCMMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQCbeKZbkJmRQ7anpkfJvmpcgycM9Xz4Sj9yYlvb9jx56tvHOdexjbd4
# fQfSNbdkqEBJiUxgd3GSgbzpjXUURew2Fq3OE/8UPBf6UKHvCZB+A3Ki31wVmDiD
# 7QN0+Jvq8hjQu5yCgpsIOhZ10kVxjZAf0RhJ7rXXp2MluZJVAE7MfHpZ6ZH7BMXq
# 8JRLnEYBUq9pD4lyPFLVBUmq3OcFy9umIA444GGSORRe3mz7j9faUPYFBUCxgk+I
# cFEcFXnHUJcg3DD/SKn0Go/DsS4JIkXybNJZvnTqmUWp4QJlUFWwQml3djR7hH5j
# c2QSRdSaeMj40zBKl8IH6SO/EKZ9p55HoYIXsDCCF6wGCisGAQQBgjcDAwExghec
# MIIXmAYJKoZIhvcNAQcCoIIXiTCCF4UCAQMxDzANBglghkgBZQMEAgEFADCCAVoG
# CyqGSIb3DQEJEAEEoIIBSQSCAUUwggFBAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIIhGsfPAADivGyH7n9EzObyHT5USf0jTEK034DuFkVHVAgZoLgjj
# dX0YEzIwMjUwNTIyMTQ1OTU3Ljc1MVowBIACAfSggdmkgdYwgdMxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1Mg
# RVNOOjRDMUEtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloIIR/jCCBygwggUQoAMCAQICEzMAAAH/Ejh898Fl1qEAAQAAAf8w
# DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcN
# MjQwNzI1MTgzMTE5WhcNMjUxMDIyMTgzMTE5WjCB0zELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQg
# T3BlcmF0aW9ucyBMaW1pdGVkMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046NEMx
# QS0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZp
# Y2UwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDJ6JXSkHtuDz+pz+aS
# IN0lefMlY9iCT2ZMZ4jenNCmzKtElERZwpgd3/11v6DfPh1ThUKQBkReq+TE/lA1
# O0Ebkil7GmmHg+FuIkrC9f5RLgqRIWF/XB+UMBjW270JCqGHF8cVXu+G2aocsIKY
# PGFk+YIGH39d8UlAhTBVlHxG1SSDOY31uZaJiB9fRH5sMCedxR22nXGMaYKl0EzK
# CT8rSHdtRNTNAdviQ9/bKWQo+hYVifYY1iBbDw8YFQ7S9MwqNgPqkt4E/SFkOHk/
# d/jGEYubrH3zG4hCn9EWfMFuC2HJJcaX41PVxkCobISFPsvRJ1HupCW/mnAM16ts
# rdhIQMqTewOH1LrSEsk2o/vWIcqQbXvkcDKDrOYTmnd842v398gSk8CULxiKzFdo
# ZfhGkMFhUqkaPQUJnCKyJmzGbRf3DplKTw45d/wnFNhYip9G5bN1SKvRneOI461o
# Ortd3KkHiBmuGv3Qpw9MNHC/LrTOtBxr/UPUns9AkAk5tuJpuiLXa6xXxrG2VP90
# J48Lid1wVxqvW/5+cKWGz27cWfouQcNFl83OFeAsMTBvp0DjLezob6BDfmj3SPaL
# pqZprwmxX9wIX6INIbMDFljWxDWat0ybPF9bNc3qw8kzLj212xZMiBlZU5JL25Qe
# FJiRuAzGct6Ipd4HkwH1Axw5JwIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFMP6leT+
# tP93sT/RATuEfTDP7pRhMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1Gely
# MF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lv
# cHMvY3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNy
# bDBsBggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBD
# QSUyMDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQA5I03kykuL
# K6ebzrp+tYiLSF1rMo0uBGndZk9+FiA8Lcr8M0zMuWJhBQCnpa2CiUitq2K9eM4b
# WUiNrIb2vp7DgfWfldl0N8nXYMuOilqnl7WJT9iTR660/J86J699uwjNOT8bnX66
# JQmTvvadXNq7qEjYobIYEk68BsBUVHSDymlnAuCFPjPeaQZmOr87hn89yZUa2Mam
# zZMK0jitmM81bw7hz/holGZhD811b3UlGs5dGnJetMpQ97eQ3w3nqOmX2Si0uF29
# 3z1Fs6wk1/ZfOpsBXteNXhxoKCUDZu3MPFzJ9/BeEu70cxTd0thMAj3WBM1QXsED
# 2rUS9KUIoqU3w3XRjiJTSfIiR+lHFjIBtHKrlA9g8kcYDRPLQ8PzdoK3v1FrQh0M
# gxK7BeWlSfIjLHCsPKWB84bLKxYHBD+Ozbj1upA5g92nI52BF7y1d0auAOgF65U4
# r5xEKVemKY1jCvrWhnb+Q8zNWvNFRgyQFd71ap1J7OHy3K266VhhxEr3mqKEXSKt
# Czr9Y5AmW1Bfv2XMVcT0UWWf0yLHRqz4Lgc/N35LRsE3cDddFE7AC/TXogK5PyFj
# UifJbuPBWY346RDXN6LroutTlG0DPSdPHHk54/KOdNoi1NJjg4a4ZTVJdofj0lI/
# e3zIZgD++ittbhWd54PvbUWDBolOgcWQ4jCCB3EwggVZoAMCAQICEzMAAAAVxedr
# ngKbSZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRp
# ZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4
# MzIyNVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqGSIb3
# DQEBAQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5osvXJHm9DtWC0/3unAcH0qls
# TnXIyjVX9gF/bErg4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLA
# EBjoYH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AKOG6N7dcP2CZTfDlhAnrE
# qv1yaa8dq6z2Nr41JmTamDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyF
# Vk3v3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJj361VI/c+gVVmG1o
# O5pGve2krnopN6zL64NF50ZuyjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg
# 3viSkR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9X4C626p+Nuw2
# TPYrbqgSUei/BQOj0XOmTTd0lBw0gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07B
# MzlMjgK8QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoSCtdjbwzJ
# NmSLW6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6
# r1AFemzFER1y7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57t7c+
# auIurQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUBBAUCAwEAATAjBgkrBgEEAYI3
# FQIEFgQUKqdS/mTEmr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl
# 0mWnG1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yDfQEBMEEwPwYIKwYBBQUH
# AgEWM2h0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0
# b3J5Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEEAYI3FAIEDB4KAFMA
# dQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAW
# gBTV9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8v
# Y3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRf
# MjAxMC0wNi0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRw
# Oi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEw
# LTA2LTIzLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvhnnJL
# /Klv6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx80HU5bbsPMeTCj/ts0aGUGCLu
# 6WZnOlNN3Zi6th542DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5t
# ggz1bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn0CS9QKC/GbYSEhFdPSfg
# QJY4rPf5KYnDvBewVIVCs/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8s
# CXgU6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aRAfbOxnT99kxybxCr
# dTDFNLB62FD+CljdQDzHVG2dY3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZ
# c9d/HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I6IleT53S0Ex2
# tVdUCbFpAUR+fKFhbHP+CrvsQWY9af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8C
# wYKiexcdFYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbhIurwJ0I9
# JZTmdHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDB
# cQZqELQdVTNYs6FwZvKhggNZMIICQQIBATCCAQGhgdmkgdYwgdMxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1Mg
# RVNOOjRDMUEtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQCpE4xsxLwlxSVyc+TBEsVE9cWymaCB
# gzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEB
# CwUAAgUA69kwEjAiGA8yMDI1MDUyMjA1MDkzOFoYDzIwMjUwNTIzMDUwOTM4WjB3
# MD0GCisGAQQBhFkKBAExLzAtMAoCBQDr2TASAgEAMAoCAQACAg2lAgH/MAcCAQAC
# AhK6MAoCBQDr2oGSAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKg
# CjAIAgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBABCn82l3
# hQSEYu05WstghY50GeqxiYiGBLDGYP3Y+IN9oqI6g0TmwYX2UOgMfxR5Nk8Sc34Q
# MokQELtcsuhhkls3qORzKYDzeo3GlCy9OCi3MYrPRwYRXymMONjObosm5HX+8tcr
# 2vqtJ69diiSDo9eSFX26JzTS/wr/FyouCCen8EwvBiiDF2K++goEgNZyokkeFjxS
# y4D8TKry1i1MtVPceqD78XFyKfuVZkiEJuYeHH/qo62qwzX2SNCBM6FB+O04aIVB
# cs/0QFWdsVRCdjbVe7JeGx/TLmu/pwRo74QGvvKP/2ZUS5c+SFjxBhiMrI6bIZYx
# PGRgTVzpNhipgrwxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0EgMjAxMAITMwAAAf8SOHz3wWXWoQABAAAB/zANBglghkgBZQMEAgEFAKCCAUow
# GgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCDsWfZH
# WlCxPa2EGeyd29h8z60x6bIHbmRp1ZmTLJMxEzCB+gYLKoZIhvcNAQkQAi8xgeow
# gecwgeQwgb0EIOQy777JAndprJwi4xPq8Dsk24xpU4jeoONIRXy6nKf9MIGYMIGA
# pH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcT
# B1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UE
# AxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAH/Ejh898Fl1qEA
# AQAAAf8wIgQg8pryN44Uv6hyev1m5bFdcE4jCyGmzNXVQdYyBUdfoQUwDQYJKoZI
# hvcNAQELBQAEggIAobS+apCW+w9yeHJrCYaLuVxMllilsDqQQBn8h0CUGPMlOjzu
# niZcC6ub4ZBBKPC43ppgqoMX/WJQZjYEVH0KYExqPgMc+wTYJsBO4ZfbpNVHr3e3
# 3tDckYFXUhdMU2lup3nyLyvPQ4j8zfp7bA17ZfRkn1I3Mu0ek+3Xuf57N/oxV9+8
# tpTLFqJC5JrSRv9tKieNMdl1HiHnzemMMyOFeGyfkh2FyO0EKVALMAoN4mOpXPJ4
# pTDqv7z2ocwuDQobJqrvPWdst3ZpHfnq66FnGm0yoGdQm70gik5zOQ3uRlmkp/tl
# EfO0s8GGBjbPyGV3G3mHqQPP6i4QqtfI6lZyij3xtsF+mVsSE4eGODN/ZUg5ZyCv
# /S0iD88uHoiHjS30/8fgAkRGjOLVCK+lSbJHJVCFUxltw0vQ3XFYBO11mTM6oHRk
# +TCc54kTNRT3jagnduTIvFexgYuocU5rgOxJOMzpQdNhoC48i3soND/xjuqiKMjm
# Gpk4gtQdLdGir2D2FIQICYLsIc/e/ENHtjacdtf66TocvMJZ2Mu96NxcqVTsSkue
# wbKDTkNjL+wIAYqoc9d+iVXrtVwVrOxnaMy+p+H3G7idQ9Jn2hyTflnyKhu1W296
# 7y9qMD8qCwN3hZVC3I2dlrUul/S7/8weJvS+w/+gYDopY3kGi7vHl9/4e5I=
# SIG # End signature block
