using module ../utils/utils.psd1

<#
.FUNCTION
    Get-RunbookSchema

.SYNOPSIS
    Retrieves the JSON schema for a runbook.

.DESCRIPTION
    The `Get-RunbookSchema` function returns the JSON schema that defines the structure
    of a Well-Architected Reliability Assessment (WARA) runbook. This schema ensures consistency
    in the configuration and validation of runbook content.

.OUTPUTS
    [string]
    The JSON schema as a string.

.EXAMPLE
    $schema = Get-RunbookSchema

    Retrieves the JSON schema for a runbook, ensuring it adheres to the expected structure.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
function Get-RunbookSchema {
    @"
{
  "title": "Runbook",
  "description": "A well-architected reliability assessment (WARA) runbook",
  "type": "object",
  "properties": {
    "parameters": {
      "type": "object"
    },
    "variables": {
      "type": "object"
    },
    "selectors": {
      "type": "object",
      "additionalProperties": {
        "type": "string"
      }
    },
    "checks": {
      "type": "object",
      "additionalProperties": {
        "type": "object",
        "additionalProperties": {
          "oneOf": [
            {
              "type": "string"
            },
            {
              "type": "object",
              "properties": {
                "selector": {
                  "type": "string"
                },
                "parameters": {
                  "type": "object"
                },
                "tags": {
                  "type": "array",
                  "items": {
                    "type": "string"
                  }
                }
              },
              "required": [
                "selector"
              ]
            }
          ]
        }
      }
    }
  },
  "required": [
    "selectors",
    "checks"
  ]
}
"@
}

<#
.FUNCTION
    New-RunbookFactory

.SYNOPSIS
    Instantiates a new RunbookFactory object.

.DESCRIPTION
    The `New-RunbookFactory` function creates an instance of the `RunbookFactory` class,
    which is responsible for parsing runbook files and generating `Runbook` instances.

.OUTPUTS
    [RunbookFactory]
    A new instance of the RunbookFactory class.

.EXAMPLE
    $factory = New-RunbookFactory

    Creates a new instance of the RunbookFactory class, which can then be used
    to parse runbook content.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
function New-RunbookFactory {
    return [RunbookFactory]::new()
}

<#
.FUNCTION
    New-Runbook

.SYNOPSIS
    Creates a new Runbook instance.

.DESCRIPTION
    The `New-Runbook` function returns a new instance of the `Runbook` class.
    It can optionally be initialized from a JSON string (-FromJson) or a JSON file (-FromJsonFile),
    but not both.

.PARAMETER FromJson
    JSON content to initialize the runbook. Cannot be used with -FromJsonFile.

.PARAMETER FromJsonFile
    Path to a JSON file to initialize the runbook. Cannot be used with -FromJson.

.OUTPUTS
    [Runbook]
    A new `Runbook` instance.

.EXAMPLE
    $runbook = New-Runbook

    Creates an empty `Runbook` instance.

.EXAMPLE
    $runbook = New-Runbook -FromJson $jsonContent

    Initializes a `Runbook` instance from JSON content.

.EXAMPLE
    $runbook = New-Runbook -FromJsonFile "C:\runbook.json"

    Initializes a `Runbook` instance from a JSON file.

.NOTES
    - If neither `-FromJson` nor `-FromJsonFile` is specified, an empty `Runbook` instance is returned.
    - If both `-FromJson` and `-FromJsonFile` are specified, the function throws an error.
    - The provided JSON must be valid and match the expected `Runbook` schema.

    Author: Casey Watson
    Date: 2025-02-27
#>
function New-Runbook {
    param(
        [Parameter(Mandatory = $false)]
        [string] $FromJson,

        [Parameter(Mandatory = $false)]
        [ValidateScript({ Test-FileExists -Path $_ })]
        [string] $FromJsonFile
    )

    if ($FromJson -or $FromJsonFile) {
        if ($FromJson -and $FromJsonFile) {
            throw "Cannot specify both -FromJson and -FromJsonFile."
        }
        else {
            $runbookFactory = New-RunbookFactory

            if ($FromJson) {
                return $runbookFactory.ParseRunbookContent($FromJson)
            }
            elseif ($(Test-RunbookFile -Path $FromJsonFile)) {
                return $runbookFactory.ParseRunbookFile($FromJsonFile)
            }
        }
    }
    else {
        return [Runbook]::new()
    }
}

<#
.FUNCTION
    New-Recommendation

.SYNOPSIS
    Creates a new `Recommendation` instance.

.DESCRIPTION
    The `New-Recommendation` function initializes and returns a new instance of the `Recommendation` class.
    This object contains metadata and evaluation logic for a specific runbook check.

.OUTPUTS
    [Recommendation]
    A new `Recommendation` object.

.EXAMPLE
    $recommendation = New-Recommendation

    Creates a new `Recommendation` instance.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
function New-Recommendation {
    return [Recommendation]::new()
}

<#
.FUNCTION
    New-RunbookRecommendation

.SYNOPSIS
    Creates a new `RunbookRecommendation` instance.

.DESCRIPTION
    The `New-RunbookRecommendation` function initializes and returns a new instance of
    the `RunbookRecommendation` class, which encapsulates metadata for a specific runbook
    check, including its associated recommendation.

.OUTPUTS
    [RunbookRecommendation]
    A new `RunbookRecommendation` object.

.EXAMPLE
    $runbookRecommendation = New-RunbookRecommendation

    Creates a new `RunbookRecommendation` instance.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
function New-RunbookRecommendation {
    return [RunbookRecommendation]::new()
}

<#
.FUNCTION
    New-RunbookCheckSet

.SYNOPSIS
    Creates a new `RunbookCheckSet` instance.

.DESCRIPTION
    The `New-RunbookCheckSet` function returns a new instance of the `RunbookCheckSet` class,
    which represents a logical grouping of related runbook checks.

.OUTPUTS
    [RunbookCheckSet]
    A new `RunbookCheckSet` instance.

.EXAMPLE
    $checkSet = New-RunbookCheckSet

    Creates a new `RunbookCheckSet` instance.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
function New-RunbookCheckSet {
    return [RunbookCheckSet]::new()
}

<#
.FUNCTION
    New-RunbookCheck

.SYNOPSIS
    Creates a new `RunbookCheck` instance.

.DESCRIPTION
    The `New-RunbookCheck` function returns a new instance of the `RunbookCheck` class,
    representing an individual check within a runbook. Each check is associated with a selector
    and contains parameterized logic for evaluating resource compliance.

.OUTPUTS
    [RunbookCheck]
    A new `RunbookCheck` instance.

.EXAMPLE
    $check = New-RunbookCheck

    Creates a new `RunbookCheck` instance.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
function New-RunbookCheck {
    return [RunbookCheck]::new()
}

<#
.FUNCTION
    Build-RunbookQueries

.SYNOPSIS
    Constructs queries for runbook checks.

.DESCRIPTION
    The `Build-RunbookQueries` function generates a list of queries based on the check sets
    and associated recommendations within a runbook. It dynamically merges parameters, variables,
    and selectors to construct accurate queries for evaluation.

.PARAMETER Runbook
    The `Runbook` object containing check sets, parameters, and selectors.

.PARAMETER Recommendations
    An array of `RunbookRecommendation` objects defining the checks to be executed.

.PARAMETER ProgressId
    (Optional) A progress indicator ID for `Write-Progress`.

.OUTPUTS
    [RunbookQuery[]]
    An array of `RunbookQuery` objects, each containing a check name, query, tags, and recommendation.

.EXAMPLE
    $queries = Build-RunbookQueries -Runbook $runbook -Recommendations $recommendations

    Constructs queries for the given runbook and recommendations.

.NOTES
    - Queries are created by combining parameters, variables, and selectors with recommendation queries.
    - Throws an error if a check references a missing selector.
    - Ensures that only recommendations matching the runbook's check sets are processed.

    Author: Casey Watson
    Date: 2025-02-27
#>
function Build-RunbookQueries {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Runbook] $Runbook,

        [Parameter(Mandatory = $true)]
        [RunbookRecommendation[]] $Recommendations,

        [Parameter(Mandatory = $false)]
        [int] $ProgressId = 30
    )

    $checkCount = 0
    $checkIndex = 0

    foreach ($checkSetKey in $Runbook.CheckSets.Keys) {
        $checkSet = $Runbook.CheckSets[$checkSetKey]
        $checkCount += $checkSet.Checks.Count
    }

    $queries = @()
    $globalParameters = @{}

    if ($Runbook.Parameters) {
        foreach ($globalParameterKey in $Runbook.Parameters.Keys) {
            $globalParameters[$globalParameterKey] = $Runbook.Parameters[$globalParameterKey].ToString()
        }
    }

    if ($Runbook.Variables) {
        foreach ($variableKey in $Runbook.Variables.Keys) {
            $variableValue = $Runbook.Variables[$variableKey].ToString()
            $globalParameters[$variableKey] = Merge-ParametersIntoString -Parameters $globalParameters -Into $variableValue
        }
    }

    $recommendationsMap = @{}

    foreach ($recommendation in $Recommendations) {
        if (-not ($recommendationsMap.ContainsKey($recommendation.CheckSetName))) {
            $recommendationsMap[$recommendation.CheckSetName] = @{}
        }

        $recommendationsMap[$recommendation.CheckSetName][$recommendation.CheckName] = $recommendation.Recommendation
    }

    foreach ($checkSetKey in $Runbook.CheckSets.Keys) {
        if (-not ($recommendationsMap.ContainsKey($checkSetKey))) {
            throw "No recommendations found for check set [$checkSetKey]."
        }

        $checkSet = $Runbook.CheckSets[$checkSetKey]

        foreach ($checkKey in $checkSet.Checks.Keys) {
            $checkIndex++
            $pctComplete = (($checkIndex / $checkCount) * 100)

            Write-Progress `
                -Activity "Building runbook queries" `
                -Status "$checkKey" `
                -PercentComplete $pctComplete `
                -Id $ProgressId

            if (-not ($recommendationsMap[$checkSetKey].ContainsKey($checkKey))) {
                throw "No recommendation found for check [$checkSetKey]:[$checkKey]."
            }

            $check = $checkSet.Checks[$checkKey]
            $recommendation = $recommendationsMap[$checkSetKey][$checkKey]

            $checkParameters = @{}
            $checkParameters += $globalParameters

            foreach ($checkParameterKey in $check.Parameters.Keys) {
                $checkParameterValue = $check.Parameters[$checkParameterKey].ToString()
                $checkParameters[$checkParameterKey] = Merge-ParametersIntoString -Parameters $globalParameters -Into $checkParameterValue
            }

            if ($Runbook.Selectors.ContainsKey($check.SelectorName)) {
                $query = Merge-ParametersIntoString -Parameters $checkParameters -Into $recommendation.Query

                foreach ($selectorKey in $Runbook.Selectors.Keys) {
                    $selector = $Runbook.Selectors[$selectorKey]
                    $selector = Merge-ParametersIntoString -Parameters $checkParameters -Into $selector
                    $query = $($query -replace "//\s*selector:$($selectorKey)", "| where $selector")
                }

                $selector = Merge-ParametersIntoString -Parameters $checkParameters -Into $Runbook.Selectors[$check.SelectorName]
                $query = $($query -replace "//\s*selector", "| where $selector")

                $queries += [RunbookQuery]@{
                    CheckSetName   = $checkSetKey
                    CheckName      = $checkKey
                    SelectorName   = $check.SelectorName
                    Query          = $query
                    Tags           = $check.Tags
                    Recommendation = $recommendation
                }
            }
        }
    }

    Write-Progress -Id $ProgressId -Completed

    return $queries
}

<#
.FUNCTION
    Merge-ParametersIntoString

.SYNOPSIS
    Replaces placeholders in a string with parameter values.

.DESCRIPTION
    The `Merge-ParametersIntoString` function iterates through a hashtable of parameters
    and replaces placeholders in the input string with corresponding values.
    Placeholders must follow the `{{Key}}` format.

.PARAMETER Parameters
    A hashtable containing key-value pairs for replacement.

.PARAMETER Into
    The string containing placeholders to be replaced.

.OUTPUTS
    [string]
    A string with placeholders replaced by their corresponding values.

.EXAMPLE
    $params = @{ "Region" = "eastus"; "Env" = "Production" }
    $result = Merge-ParametersIntoString -Parameters $params -Into "Deploying to {{Region}} in {{Env}}."

    Returns: "Deploying to eastus in Production."

.NOTES
    - Only placeholders matching keys in the hashtable are replaced.
    - Uses simple string replacement logic.

    Author: Casey Watson
    Date: 2025-02-27
#>
function Merge-ParametersIntoString {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [hashtable] $Parameters,

        [Parameter(Mandatory = $true)]
        [string] $Into
    )

    foreach ($parameterKey in $Parameters.Keys) {
        $Into = $Into.Replace("{{$parameterKey}}", $Parameters[$parameterKey])
    }

    return $Into
}

<#
.FUNCTION
    Read-RunbookFile

.SYNOPSIS
    Reads and parses a runbook file.

.DESCRIPTION
    The `Read-RunbookFile` function validates and loads a runbook from a JSON file.
    If the file is valid, it returns a parsed `Runbook` instance.

.PARAMETER Path
    The file path to the runbook JSON file.

.OUTPUTS
    [Runbook]
    A parsed `Runbook` object.

.EXAMPLE
    $runbook = Read-RunbookFile -Path "C:\runbook.json"

    Reads and parses the specified runbook file.

.NOTES
    - Uses `Test-RunbookFile` to validate the file before parsing.
    - If validation fails, an error is thrown.
    - The runbook is parsed using `RunbookFactory`.

    Author: Casey Watson
    Date: 2025-02-27
#>
function Read-RunbookFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-FileExists $_ })]
        [string] $Path
    )

    if (Test-RunbookFile -Path $Path) {
        $runbookFactory = New-RunbookFactory
        return $runbookFactory.ParseRunbookFile($Path)
    }
}

<#
.FUNCTION
    Write-RunbookFile

.SYNOPSIS
    Saves a `Runbook` object to a JSON file.

.DESCRIPTION
    The `Write-RunbookFile` function validates the provided `Runbook` object
    and serializes it into a JSON file at the specified path.

.PARAMETER Runbook
    The `Runbook` object to be saved.

.PARAMETER Path
    The file path where the `Runbook` JSON should be written.

.OUTPUTS
    None

.EXAMPLE
    Write-RunbookFile -Runbook $myRunbook -Path "C:\runbook.json"

    Saves the provided `Runbook` object to "C:\runbook.json".

.NOTES
    - Ensures the `Runbook` is valid before saving.
    - The output file is formatted in JSON.

    Author: Casey Watson
    Date: 2025-02-27
#>
function Write-RunbookFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Runbook] $Runbook,

        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $Runbook.Validate()

    $runbookFileContents = @{
        query_paths = $Runbook.QueryPaths
        parameters  = $Runbook.Parameters
        variables   = $Runbook.Variables
        selectors   = $Runbook.Selectors
        checks      = @{}
    }

    foreach ($checkSetKey in $Runbook.CheckSets.Keys) {
        $checkSet = $Runbook.CheckSets[$checkSetKey]
        $checkSetContents = @{}

        foreach ($checkKey in $checkSet.Checks.Keys) {
            $check = $checkSet.Checks[$checkKey]

            $checkContents = @{
                parameters = ($check.Parameters ?? @{})
                selector   = $check.SelectorName
                tags       = ($check.Tags ?? @())
            }

            $checkSetContents[$checkKey] = $checkContents
        }

        $runbookFileContents.checks[$checkSetKey] = $checkSetContents
    }

    $runbookFileJson = $runbookFileContents | ConvertTo-Json -Depth 15
    $runbookFileJson | Out-File -FilePath $Path -Force
}

<#
.FUNCTION
    Test-RunbookFile

.SYNOPSIS
    Validates a runbook file.

.DESCRIPTION
    The `Test-RunbookFile` function checks whether a specified runbook JSON file is
    valid according to the runbook schema. It ensures the JSON structure is correct
    and adheres to expected schema requirements.

.PARAMETER Path
    The full file path to the runbook JSON file.

.OUTPUTS
    [bool]
    Returns `$true` if the file is valid; otherwise, an error is thrown.

.EXAMPLE
    $isValid = Test-RunbookFile -Path "C:\runbook.json"

    Returns `$true` if the runbook is valid.

.NOTES
    - Uses `Test-Json` to validate JSON structure.
    - If validation fails, an error is thrown.

    Author: Casey Watson
    Date: 2025-02-27
#>
function Test-RunbookFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-FileExists $_ })]
        [string] $Path
    )

    $fileContent = Get-Content -Path $Path -Raw

    if (-not ($fileContent | Test-Json -ErrorAction SilentlyContinue)) {
        throw "[$Path] is not a valid JSON file."
    }

    if (-not ($fileContent | Test-Json -ErrorAction SilentlyContinue -Schema $(Get-RunbookSchema))) {
        throw "[$Path] does not adhere to the runbook JSON schema. Run [Get-RunbookSchema] to get the schema."
    }

    $runbookFactory = New-RunbookFactory
    $runbookFactory.ParseRunbookContent($fileContent).Validate()

    return $true
}

<#
.FUNCTION
    Build-RunbookSelectorReview

.SYNOPSIS
    Builds a selector review for a runbook.

.DESCRIPTION
    The `Build-RunbookSelectorReview` function evaluates each selector in a runbook,
    resolves parameters, and executes queries to identify matching resources across
    specified subscriptions.

.PARAMETER Runbook
    The `Runbook` object containing selectors, parameters, and variables.

.PARAMETER SubscriptionIds
    (Optional) An array of subscription IDs to scope the queries.

.OUTPUTS
    [SelectorReview]
    A `SelectorReview` object mapping each selector to its resolved query and matched resources.

.EXAMPLE
    $review = Build-RunbookSelectorReview -Runbook $runbook -SubscriptionIds @("sub1", "sub2")

    Generates a selector review to verify correct resource scoping.

.NOTES
    - Selectors define which resources are included in a runbook.
    - Misconfigured selectors may cause missing or incorrect results.
    - Uses `Invoke-WAFQuery` to fetch matching resources.

    Author: Casey Watson
    Date: 2025-02-27
#>
function Build-RunbookSelectorReview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Runbook] $Runbook,

        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]] $SubscriptionIds
    )

    $selectorReview = [SelectorReview]::new()

    $globalParameters = @{}

    if ($Runbook.Parameters) {
        foreach ($globalParameterKey in $Runbook.Parameters.Keys) {
            $globalParameters[$globalParameterKey] = $Runbook.Parameters[$globalParameterKey].ToString()
        }
    }

    if ($Runbook.Variables) {
        foreach ($variableKey in $Runbook.Variables.Keys) {
            $variableValue = $Runbook.Variables[$variableKey].ToString()
            $globalParameters[$variableKey] = Merge-ParametersIntoString -Parameters $globalParameters -Into $variableValue
        }
    }

    for ($i = 0; $i -lt $Runbook.Selectors.Keys.Count; $i++) {
        $selectorKey = $Runbook.Selectors.Keys[$i]
        $selector = Merge-ParametersIntoString -Parameters $globalParameters -Into $Runbook.Selectors[$selectorKey]
        $pctComplete = ((($i + 1) / $Runbook.Selectors.Keys.Count) * 100)

        Write-Progress `
            -Activity "Building selector review..." `
            -Status "$pctComplete% - Processing selector [$selectorKey]" `
            -PercentComplete $pctComplete

        $selectedResourceSet = [SelectedResourceSet]@{
            Selector           = $selector
            ResourceGraphQuery = Build-SelectorResourceGraphQuery -Selector $selector
        }

        $selectedResources = Invoke-WAFQuery `
            -Query $selectedResourceSet.SelectorResourceGraphQuery `
            -SubscriptionIds $SubscriptionIds

        foreach ($selectedResource in $selectedResources) {
            $selectedResourceSet.Resources += [SelectedResource]@{
                ResourceId        = $selectedResource.id
                ResourceType      = $selectedResource.type
                ResourceName      = $selectedResource.name
                ResourceLocation  = $selectedResource.location
                ResourceGroupName = $selectedResource.resourceGroup
                ResourceTags      = $(ConvertTo-Json $selectedResource.tags | ConvertFrom-Json -AsHashtable)
            }
        }

        $selectorReview.Selectors[$selectorKey] = $selectedResourceSet
    }

    Write-Progress -Activity "Selector review built." -Completed

    return $selectorReview
}

<#
.FUNCTION
    Build-SelectorResourceGraphQuery

.SYNOPSIS
    Constructs an Azure Resource Graph query from a selector.

.DESCRIPTION
    The `Build-SelectorResourceGraphQuery` function creates a Resource Graph query
    that filters resources based on the provided selector expression.

.PARAMETER Selector
    The filter expression used to scope resources in the query.

.OUTPUTS
    [string]
    The formatted Azure Resource Graph query.

.EXAMPLE
    $query = Build-SelectorResourceGraphQuery -Selector "type == 'Microsoft.Compute/virtualMachines'"

    Generates a query to filter virtual machines.

.NOTES
    - The selector should be a valid KQL (Kusto Query Language) expression.
    - The output query can be executed using Azure Resource Graph API.

    Author: Casey Watson
    Date: 2025-02-27
#>
function Build-SelectorResourceGraphQuery {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Selector
    )

    return @"
resources
| where $Selector
| project id, type, location, name, resourceGroup, tags
"@
}


# SIG # Begin signature block
# MIIoLAYJKoZIhvcNAQcCoIIoHTCCKBkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD8w0fHvQkIvTEQ
# SeZXXCMRCs6xwu0Ia9YZpQni+R9jbKCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGgwwghoIAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIKF0idhUio/IRtpT57uoZkU1
# S10z5meqlWliXgtnLvD5MEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQCB1NcZO2/p5dwn0QD6wcXlaYUbvOAvs65YGUWgv9Lp0lrUBMgMfWyl
# cp7gGblRBwkveKHEEVfPy24B3xk8SumaEmsqRXhzNbPUbw8w5hMYjqXQBcsBsBa/
# 2m6hsY3WIQ5aB5BoSUd1cYh/GdcjCWDGUFaGWmkSKDtfgj2DFdTH916L7CNLwb2M
# YPiE4CNN5oe77jbEFyQTs5N4aJnuFQsUlQpjbWcTiC+rwzHE/l91LhZ3jjGdjfHy
# xbcw9Jl2M1MWkR/cOF4tK8IG2b6xcQFADatjIHC768vEPxPDHzTEXzbsPOAiJJtz
# UdZE1xG2Ejf+VtWEsrrKDhuKicdWqF30oYIXlDCCF5AGCisGAQQBgjcDAwExgheA
# MIIXfAYJKoZIhvcNAQcCoIIXbTCCF2kCAQMxDzANBglghkgBZQMEAgEFADCCAVIG
# CyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIAqoTvwEabjzJ8K3kp5q3AK2ZsMk8tkCp9DgU46Jcsi9AgZoLSNQ
# 7MIYEzIwMjUwNTIyMTUwMDIyLjg0OFowBIACAfSggdGkgc4wgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo4RDAw
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZaCCEeowggcgMIIFCKADAgECAhMzAAACDQ13vns2j3/jAAEAAAINMA0GCSqGSIb3
# DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDEzMDE5
# NDMwMVoXDTI2MDQyMjE5NDMwMVowgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo4RDAwLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBALF/qAfd8ffCCYU3lNXzNEX83RMmC5ZRgKtBk9Sb
# AD5CFLJI2nuSIaYVl3hvnOwB83RaH4/Mgm9ixUyBUJplxtWv8MkF97HCdB9z5+Ai
# M6IDNhKODsS1I51aguTzVv+YaIYgEL758txIDNtk7cq1bg9Eo5Kx0P/l5F201xr8
# R3gQBxRPfbMAZL1Kr//iQNP3YoTNE1cmPzPc0MvigYbMJAy8Dze1Abmaud6nEXQ6
# w18oodgcyqPI4Q/mlOAp9Plcx/PSwQboBECOVcnv1Ib+Oh64eHxOoL5wkyvL3AHt
# kut0wAXLysc3dbX6SSnDjsqj72LN7ZuV74/mMxqq58rTfcG77YVdVIqmGKXaPSoF
# Ramfm8G1/Zb6CxWrJ0D56x4Ed9jFwNnj7mhGjIMCDS1b1/v284C/3U/htPMb4Fk8
# FsMzNIPCaPwBIoOVXCu5N4XV6bs1aJ51YAH7b2AUf9z803W5NAu/AYMg7DUNwucI
# PRVc0vuY6+J9i81a6BMAdwwUtWUDWN44ffkT8hb2pwoqqtqHsvIlrH/ozlQezBC6
# AaAuuKIw0S/Pcipb4CIgnT3uJjMwbOwOpstJTO0iMMS33F8gcdSz+neP0GSvZs8/
# W/jPy+n/jkKTeZ8VtbYXQS2M4KcY8ysk2OMaEQb3MArXXTkGDaQntGfxZNY7khsP
# +sO1AgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUFZeOkOYOA2tOwhKPbpKNTX9gw4Aw
# HwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
# UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAw
# XjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# ZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBABwfyBPjbBVDwZBn4wchc0n53h1CfT68
# 8b8B0W9PGT+2846QDaLuGegQ4QVxJvNLkR0hUIyLljkSJZPyEQmTO116rmFe1Raw
# 9j7aVhJT0d2EoN9EN+PBVn2R8K2GstELePD0VqHkwAYj0wdIROj2vv5y7HwBSpDe
# GZozaOZM1a+sU4ts7TUHVyI0Zu8TbEr6tvMEGH+5Z3eTfSlq7ounRrODusNgYwa/
# yNai0gt+kowF4JCi58ywj+DPGjj6T12pgGQDAyTuHWF/bfJxvlaeBIxBX9TGXdLF
# uD7rgoroYKxNIbrPvM1OBsvr08wZ/BXZydrBj3kjLZyAnpxsw5FFRxjO+y5R3ygR
# DztxbheqoEuuF57AmNE5PpjJatDahD6PMYrZYmgY2d9qeY1+pCdUqmLfbSidqD3k
# swWOPwHCt+Viw72QQ6LLsjlSeA8GYX5EdIK/aEVKvycruC26LL1JQ4o/oVtA7PIx
# G8ndTGyasffzjvtfMRNv+1ksCPXRN1gV+SOQV4WOVveIxNKz9WslIXSvS843+jFm
# ooJjxZl7P65A/ROvLRBD3BGnkPYVgnwTW4hf8u/C4WQF1Z/tfuqVv6nN28+lIBxo
# KAEhZOUyfZtaCRTm4/PFE8eH/S74Ujv95EJkZsszaqeffJdC87+H3BCbykgAFEC+
# h1qJXRtWYdsRMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMwMTgzMjI1WjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAOThpkzntHIhC3miy9ckeb0O1YLT/e6cBwfSqWxOdcjKNVf2AX9sSuDivbk+
# F2Az/1xPx2b3lVNxWuJ+Slr+uDZnhUYjDLWNE893MsAQGOhgfWpSg0S3po5GawcU
# 88V29YZQ3MFEyHFcUTE3oAo4bo3t1w/YJlN8OWECesSq/XJprx2rrPY2vjUmZNqY
# O7oaezOtgFt+jBAcnVL+tuhiJdxqD89d9P6OU8/W7IVWTe/dvI2k45GPsjksUZzp
# cGkNyjYtcI4xyDUoveO0hyTD4MmPfrVUj9z6BVWYbWg7mka97aSueik3rMvrg0Xn
# Rm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75xqRdbZ2De+JKRHh09/SDPc31BmkZ1
# zcRfNN0Sidb9pSB9fvzZnkXftnIv231fgLrbqn427DZM9ituqBJR6L8FA6PRc6ZN
# N3SUHDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1DTsEzOUyOArxCaC4Q6oRRRuLR
# vWoYWmEBc8pnol7XKHYC4jMYctenIPDC+hIK12NvDMk2ZItboKaDIV1fMHSRlJTY
# uVD5C4lh8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqvUAV6bMURHXLvjflSxIUX
# k8A8FdsaN8cIFRg/eKtFtvUeh17aj54WcmnGrnu3tz5q4i6tAgMBAAGjggHdMIIB
# 2TASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsGAQQBgjcVAgQWBBQqp1L+ZMSavoKR
# PEY1Kc8Q/y8E7jAdBgNVHQ4EFgQUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXAYDVR0g
# BFUwUzBRBgwrBgEEAYI3TIN9AQEwQTA/BggrBgEFBQcCARYzaHR0cDovL3d3dy5t
# aWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9zaXRvcnkuaHRtMBMGA1UdJQQM
# MAoGCCsGAQUFBwMIMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQE
# AwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQ
# W9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNv
# bS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNybDBa
# BggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0
# LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3J0MA0GCSqG
# SIb3DQEBCwUAA4ICAQCdVX38Kq3hLB9nATEkW+Geckv8qW/qXBS2Pk5HZHixBpOX
# PTEztTnXwnE2P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7pZmc6U03dmLq2HnjYNi6c
# qYJWAAOwBb6J6Gngugnue99qb74py27YP0h1AdkY3m2CDPVtI1TkeFN1JFe53Z/z
# jj3G82jfZfakVqr3lbYoVSfQJL1AoL8ZthISEV09J+BAljis9/kpicO8F7BUhUKz
# /AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNhcy4sa3tuPywJeBTpkbKpW99Jo3QMvOyR
# gNI95ko+ZjtPu4b6MhrZlvSP9pEB9s7GdP32THJvEKt1MMU0sHrYUP4KWN1APMdU
# bZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMBV0lUZNlz138eW0QBjloZkWsNn6Qo
# 3GcZKCS6OEuabvshVGtqRRFHqfG3rsjoiV5PndLQTHa1V1QJsWkBRH58oWFsc/4K
# u+xBZj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8vwLBgqJ7Fx0ViY1w/ue10Cga
# iQuPNtq6TPmb/wrpNPgkNWcr4A245oyZ1uEi6vAnQj0llOZ0dFtq0Z4+7X6gMTN9
# vMvpe784cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFxBmoQtB1VM1izoXBm8qGC
# A00wggI1AgEBMIH5oYHRpIHOMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OEQwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMV
# AHssLCiPobedDs7GGX+l+d7jIBM9oIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDr2ZvvMCIYDzIwMjUwNTIyMTI0
# OTUxWhgPMjAyNTA1MjMxMjQ5NTFaMHQwOgYKKwYBBAGEWQoEATEsMCowCgIFAOvZ
# m+8CAQAwBwIBAAICI0IwBwIBAAICFEcwCgIFAOva7W8CAQAwNgYKKwYBBAGEWQoE
# AjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDANBgkq
# hkiG9w0BAQsFAAOCAQEAisb5Yn1bmYS/2f2vsuJs8RkRxRYdTqnxX3All2f5vMVF
# dlG45CADeHEc+btQtXchGBsp1AsiDjn8HEoNQoqi5C8CV2nSHt4MZU7msektj9ok
# QgFxWz6b01Y3WvGxDvBlk9TRUlT5csRlLwmfbDFTYPzmavRKiXrUPmOSpJMQMd2O
# kO5P03uWSrlLI5BVZVNlUPxlsUjxTC4gaDbzMX703FutMl//II6gOYXn3RYHwIix
# NhMeCwdQYjJ2q9vjly8aEq6HGL9tpti5YfRUDzrvlJdcFuDR1p3heySpphrkIF7t
# hiiB+jjzDpsX6pgzYKROZqg671FkMYGkTbV5DQ+5JDGCBA0wggQJAgEBMIGTMHwx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAACDQ13vns2j3/jAAEAAAIN
# MA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQw
# LwYJKoZIhvcNAQkEMSIEIPCR6pzTzW0GvlBDDNrQgjs/KjtS1yxDKNPj5fLq8PKL
# MIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQgY+oHlOwkaojw66ScEq1K9vQV
# +rrDk1Kzm95NXCRr+EAwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMAITMwAAAg0Nd757No9/4wABAAACDTAiBCAe2RiPrhaRZyFoO/pEYVOJDcZW
# 2WqHSb0B1zDlwqZpzzANBgkqhkiG9w0BAQsFAASCAgAQEyJfyA01QC14oe2n74UL
# NI42q6/FL/XDAF7WR1tN+6pnPdco8EinPQj+atropZyPrA3nFD8lZqiLkJgf7VkF
# ljQFTs5yyDBNoPhwm4GzWpfSdgXDQQ+hYNkZNhJUKABfP74Ivnu9Vv2iLipYspH0
# YJSmDHv/MSlRMEVMbaK+HSJHMbQh7GEwQOiuSBY62Dw8/7k//TLxQDe7tMpjATeH
# 7TWLWHe8+X3m4V9Q2KdCLUqmiz+fNG72hGWT3stnsHlTGIZ06jnJqPbel0GT8CtZ
# 0IXEvOHFdpSNMwNxaPH2fRowtIvQfyMbw7m3uGN8fDz5UnJPX5j5zeSoMDGoLfJz
# zZeLHKwkwAB6433ELej/pQZLz9qUBEHKazVXuBBUFW/RnrqrkabGCFgFpfCbQJap
# H2o1QfEzEBZYHqy21tON7xxLE1JpumG3eOrV0efYGTuvBKwxYbQFx9HEBp73nUvJ
# ysFV8M5uHg0ou7uEE77D+65lCaTmrafiQsDt1p2zRqJUTr9CMdVluqGohOWSPlFd
# b0+yWMKtyPyZZF6I6uL3NSotnHMWOOqescXxEGI2YAJFElgH2PPctMQKB3mqXPBG
# VnJGtrKM0WFPmjYKcN9rC/rF9xXmtUx6A3YXUwR9PWZVX9lo+jgGtS+JfPW7McFD
# 1/YEfBr9h6h4ZHtM6HxnZg==
# SIG # End signature block
