using module ../utils/utils.psd1

<#
.SYNOPSIS
    Retrieves high availability recommendations from Azure Advisor.

.DESCRIPTION
    The Get-WAFAdvisorRecommendations function queries Azure Advisor for recommendations related to high availability.
    It uses Azure Resource Graph to fetch and join relevant resource data.

.PARAMETER SubscriptionIds
    The subscription IDs for which to retrieve recommendations.

.PARAMETER AdditionalRecommendationIds
    Additional recommendation IDs to include in the query. In the WARA we use this to include Advisor recommendations that are not categorized as high availability but are still relevant.

.PARAMETER HighAvailability
    Switch to filter recommendations related to high availability.

.PARAMETER Security
    Switch to filter recommendations related to security.

.INPUTS
    None. You cannot pipe objects to this function.

.OUTPUTS
    System.Object. The function returns a list of recommendations.

.EXAMPLE
    $subId = "22222222-2222-2222-2222-222222222222"
    Get-WAFAdvisorRecommendation -SubscriptionIds $subId -HighAvailability

.EXAMPLE
$subId = "22222222-2222-2222-2222-222222222222"
$AddtionalRecommendationIds = @()"82219546-1110-4f5d-a1c2-7defb204663c","693e2dbf-cdec-47a2-8e54-79752cd7e3fc")
Get-WAFAdvisorRecommendation -SubscriptionIds $subId -HighAvailability -AdditionalRecommendationIds $AddtionalRecommendationIds

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
function Get-WAFAdvisorRecommendation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array] $SubscriptionIds,

        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [array] $AdditionalRecommendationIds,

        [switch] $HighAvailability,

        [switch] $Security,

        [switch] $Cost,

        [switch] $Performance,

        [switch] $OperationalExcellence
    )

    # Initialize an array to hold the selected categories
    $categories = @()

    # Add categories based on the selected switches
    switch ($PSBoundParameters.Keys) {
        'HighAvailability' { $categories += 'HighAvailability' }
        'Security' { $categories += 'Security' }
        'Cost' { $categories += 'Cost' }
        'Performance' { $categories += 'Performance' }
        'OperationalExcellence' { $categories += 'OperationalExcellence' }
    }

    # Convert the categories array to a comma-separated string
    $categoriesString = $categories -join "','"
    $AdditionalRecommendationIdsString = $AdditionalRecommendationIds -join "','"

    $advquery = `
        "advisorresources
| where type == 'microsoft.advisor/recommendations'
| where tostring(properties.category) in ('$categoriesString') or properties.recommendationTypeId in ('$additionalRecommendationIdsString')
| where properties.tracked !~ 'true' and properties.extendedProperties.recommendationControl != 'ServiceUpgradeAndRetirement'
| extend resId = tolower(tostring(properties.resourceMetadata.resourceId))
| join kind=leftouter (resources | project ['resId']=tolower(id), subscriptionId, resourceGroup, location, type) on resId
| extend id = iff(properties.impactedField =~ 'microsoft.subscriptions/subscriptions', strcat('/subscriptions/', subscriptionId), resId1)
| extend subscriptionId = coalesce(subscriptionId,subscriptionId1)
| extend resourceGroup = iff(properties.impactedField =~ 'microsoft.subscriptions/subscriptions', 'N/A', resourceGroup)
| extend location = iff(properties.impactedField =~ 'microsoft.subscriptions/subscriptions', 'global', coalesce(location,location1))
| extend type = iff(properties.impactedField =~ 'microsoft.subscriptions/subscriptions', 'microsoft.subscription/subscriptions', tolower(properties.impactedField))
| project recommendationId = properties.recommendationTypeId, type, name = properties.impactedValue, id, subscriptionId, resourceGroup, location, category = properties.category, impact = properties.impact, description = properties.shortDescription.solution
| order by ['id']"

    <#  $advquery = `
"advisorresources
| where type == 'microsoft.advisor/recommendations' and tostring(properties.category) in ('$categoriesString')
| extend resId = tolower(tostring(properties.resourceMetadata.resourceId))
| join kind=leftouter (resources
| project ['resId']=tolower(id), subscriptionId, resourceGroup ,location) on resId
| project recommendationId = properties.recommendationTypeId, type = tolower(properties.impactedField), name = properties.impactedValue, id = resId1, subscriptionId = subscriptionId1,resourceGroup = resourceGroup, location = location1, category = properties.category, impact = properties.impact, description = properties.shortDescription.solution
| order by ['id']" #>

    $queryResults = Invoke-WAFQuery -Query $advquery -SubscriptionId $SubscriptionIds

    $return = Build-WAFAdvisorObject -AdvQueryResult $queryResults

    return $return
}

<#
.SYNOPSIS
    Builds a list of advisory objects from Azure Advisor query results.

.DESCRIPTION
    The Build-WAFAdvisorObject function processes the results of an Azure Advisor query and constructs a list of advisory objects.
    Each advisory object contains details such as recommendation ID, type, name, resource ID, subscription ID, resource group, location, category, impact, and description.

.PARAMETER AdvQueryResult
    An array of query results from Azure Advisor.

.EXAMPLE
    $advQueryResult = Get-WAFAdvisorRecommendations -Subid "12345"

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>

<#
.SYNOPSIS
    Builds a list of Advisor resource objects from Azure Advisor query results.

.DESCRIPTION
    The `Build-WAFAdvisorObject` function processes the results of an Azure Advisor query and constructs a list of `advisorResourceObj` objects. Each object contains detailed information about an Advisor recommendation, including IDs, resource details, category, impact, and descriptions.

.PARAMETER AdvQueryResult
    An array of query results from Azure Advisor.

.INPUTS
    System.Object[]. You can pipe an array of Advisor query results to this function.

.OUTPUTS
    advisorResourceObj[]. The function returns an array of `advisorResourceObj` instances representing Advisor recommendations.

.EXAMPLE
    $advQueryResult = Get-WAFAdvisorRecommendation -SubscriptionIds "12345" -HighAvailability
    $advisorObjects = Build-WAFAdvisorObject -AdvQueryResult $advQueryResult

    This example builds Advisor resource objects from the query results.

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
function Build-WAFAdvisorObject {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [PSCustomObject[]] $AdvQueryResult
    )

    # Initialize an array to hold the processed objects
    $return = $AdvQueryResult.ForEach({ [advisorResourceObj]::new($_) })

    # Return the processed objects
    return $return
}

<#
.SYNOPSIS
    Represents an Azure Advisor recommendation resource object.

.DESCRIPTION
    The `advisorResourceObj` class encapsulates the details of an Azure Advisor recommendation resource. It contains properties such as recommendation ID, type, name, resource ID, subscription ID, resource group, location, category, impact, and description.

.PARAMETER Recommendation
    A recommendation object returned from Azure Advisor.

    The attributes of the object are used to populate the properties of the `advisorResourceObj` instance.

    [string] $recommendationId
    [string] $type
    [string] $name
    [string] $id
    [string] $subscriptionId
    [string] $resourceGroup
    [string] $location
    [string] $category
    [string] $impact
    [string] $description

.INPUTS
    None. You cannot pipe input to this class.

.OUTPUTS
    advisorResourceObj. An instance representing an Advisor recommendation.

.EXAMPLE
    $advisorRecommendation = [advisorResourceObj]::new($recommendation)

    This example creates a new instance of `advisorResourceObj` using a recommendation object.

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
class advisorResourceObj : IComparable, IEquatable[object] {
    <# Define the class. Try constructors, properties, or methods. #>
    [string] $recommendationId
    [string] $type
    [string] $name
    [string] $id
    [string] $subscriptionId
    [string] $resourceGroup
    [string] $location
    [string] $category
    [string] $impact
    [string] $description

    # Default Constructor that takes a PSObject as input
    # Right now this is just a simple assignment of properties, but can be expanded to include more complex logic in the future.
    advisorResourceObj([PSObject]$psObject) {
        $this.RecommendationId = $psObject.recommendationId
        $this.Type = $psObject.type
        $this.Name = $psObject.name
        $this.Id = $psObject.id
        $this.SubscriptionId = $psObject.subscriptionId
        $this.ResourceGroup = $psObject.resourceGroup
        $this.Location = $psObject.location
        $this.Category = $psObject.category
        $this.Impact = $psObject.impact
        $this.Description = $psObject.description
    }

    # Implement the Equals method to equate two advisorResourceObj objects
    [bool] Equals([object] $other) {
        if ($other -isnot [advisorResourceObj]) {
            throw "Expected an advisorResourceObj object"
        }
        foreach ($property in $this.PSObject.Properties.Name) {
            if ($this.$property -ne $other.$property) {
                return $false
            }
        }
        return $true
    }

    # Implement the CompareTo method to compare two advisorResourceObj objects
    [int] CompareTo([object] $other) {
        if ($other -isnot [advisorResourceObj]) {
            throw "Expected an advisorResourceObj object"
        }
        foreach ($property in $this.PSObject.Properties.Name) {
            if ($this.$property -ne $other.$property) {
                return $this.$property.CompareTo($other.$property)
            }
        }
        return 0
    }

    # Implements the GetHashCode method to generate a hash code for the advisorResourceObj object
    #[int] GetHashCode() {
    #    return [System.HashCode]::Combine($this.RecommendationId, $this.Type, $this.Name, $this.Id, $this.SubscriptionId, $this.ResourceGroup, $this.Location, $this.Category, $this.Impact, $this.Description)
    #}

}


<#
.SYNOPSIS
    Retrieves metadata from Azure Advisor.

.DESCRIPTION
    The Get-WAFAdvisorMetadata function retrieves metadata from Azure Advisor using the Azure REST API.
    It uses an access token to authenticate and fetch the metadata.

.INPUTS
    None. You cannot pipe objects to this function.

.OUTPUTS
    System.Object. The function returns the supported values from the Advisor metadata.

.EXAMPLE
    $AdvisorMetadata = Get-WAFAdvisorMetadata

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
Function Get-WAFAdvisorMetadata {
    param($ResourceURL = "https://management.azure.com/" )

    # Get an access token for the Azure REST API
    $securetoken = Get-AzAccessToken -AsSecureString -ResourceUrl $ResourceURL -WarningAction SilentlyContinue

    # Convert the secure token to a plain text token
    $token = ConvertFrom-SecureString -SecureString $securetoken.token -AsPlainText

    # Create the authorization headers
    $authHeaders = @{
        'Authorization' = 'Bearer ' + $token
    }

    # Define the URI for the Advisor metadata
    $AdvisorMetadataURI = $ResourceUrl+"providers/Microsoft.Advisor/metadata?api-version=2023-01-01" #&%24expand=ibiza will return the full metadata. Use this soon to pull in the long description.

    # Invoke the REST API to get the metadata
    $r = Invoke-RestMethod -Uri $AdvisorMetadataURI -Headers $authHeaders -Method Get

    # Return the supported values from the metadata
    return $r.value.properties[0].supportedValues
}

# SIG # Begin signature block
# MIIoLwYJKoZIhvcNAQcCoIIoIDCCKBwCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBhQnFtoEaUkeKU
# hUTmsZf7vghUT7DhEnyFAx7X3SoydqCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGg8wghoLAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIE812jNJXNjOQHNNQDpaUZN/
# +Brf8L9ABuW6YdN6gWrlMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQCjWnM1I06i2Gfb/ezWyJmKMWSt3jcXGpLFpxOyFKrD2VDpaewvI5hn
# sPCgI0xqKZz8YdeWe9fUV33qMXFnFIe9hyHx5Das8TAoe7jlniFOj28U9WPEiTjm
# c0qYvrBwA6CN4ouB1pNQTF/AFQT+OetxJXUaBjS+7c2uNxg5uBpDzuzRCI21twZh
# DmoWMSUXxpKfb3+wYWGOA4vw7fjZorrpXNqZY4CDveeuXYarpWzvYcKQLdZeoX6p
# u78mJi3WaRphms6JQJuiAD/NQTBWHyyKLFW/BKRA8vzOSzAve3/iO+Vnb26J7a3n
# 9+8Fgh2qWebaFYdtqClt8ViNiqyLBSnboYIXlzCCF5MGCisGAQQBgjcDAwExgheD
# MIIXfwYJKoZIhvcNAQcCoIIXcDCCF2wCAQMxDzANBglghkgBZQMEAgEFADCCAVIG
# CyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEID/pzJGfrwen9CKt7obDPzxsFfe1hWnoGZQFTHSzFz+mAgZoJisf
# 9Y8YEzIwMjUwNTIyMTUwMDE4LjIwNFowBIACAfSggdGkgc4wgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo3RjAw
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZaCCEe0wggcgMIIFCKADAgECAhMzAAACBte8UTiYI+wsAAEAAAIGMA0GCSqGSIb3
# DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDEzMDE5
# NDI1MFoXDTI2MDQyMjE5NDI1MFowgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo3RjAwLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBAOlEhZsgzdGWvf3tyMdpjHzmXsj5lVYYwIEIz3XU
# GlTr4gZYKqSyqCp59kUSMrM1UNgL1hyAhMDPbvo0aC8QKbhl82/8U/BxpIPPvFsN
# uw6jFvBCgdQ1Guj7Hm5tmFPpYl5T3sXTr68OMDD9i3W9Y6BFOqY/902v2iohsTmg
# Ith0ffAj+ehiawlVzv3rqf4HtQAYBZTax7cvP7F3Gc2w1fgJHrMgxUlNJ7M//ZJM
# 1zElO72TayXv+/M6HEmEJDfyt1oSiqEYeteuZWQSFK/5LTQMwlzU4hfGp9vA+Myo
# RWnsreSZzMKRu6bUE4gnbC4MBsq4l6Wm141mP9Lnw1JDDqSF+4kCW6ocreKCRL86
# 7Hj2pM/6tT49B424P4a2sKikW5xGZqdC/EhIY2jGcGrdR4NOqmGbpojsYwe0UPoM
# 6MmWWUfWBVZc9PKK9/7i03xOY7rIiAHi4/TRsf2Of93LLFKPE9Daca9m2C2qe+re
# HdNGNGeRz57VcHW5q0NrXNRxLuveKh1OnIBN7aGCRVfebgOFHMjoDhInp9skz2Kw
# sfwAYpzKaKwrNi6kB4VJMnXQkQVroyMdBhiiGgIXvtHQILAw2O8Thd8se76oo9jw
# ZB+xl2KBD1yVQCLJ0WZW3rWHK2jFk/suZdvOMPRV5zLNmgvgSq7VezMGy6UCvkt3
# YrBzAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQU7TCwsp0MalP3tzHcjKbKj9IGbhIw
# HwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
# UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAw
# XjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# ZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAHbcZk5971OFNS8Pb2Li3qUOnEmGlVEy
# Z75RvJmEEUJmGgZO2MN2mEACtTZDrVZiDdhVyXZF0mbk9RtnZsDvvOT6q0vEL7d0
# 3FWxNx23E8NJJaDAEfFOPqkKagM1eiUBixam8dAUIcOoR8CIHFfV2ZpduJM/V3Rd
# 9++BHp2yFRypof+YV+MNkDEtTWzodxWAK8FAmUnvEQbmMUp22pqkpZxtQfBNWpdA
# ZsiUdUKU0nfKpbpndQkf8IVxiItX97ry6tOYa2JnEZJhvhIFI8CtOtNh4c6VAiP/
# uWhVaZ9ZfbLgAZX8P4zPJkzK8XDhXIvRWCr3oTNArK16JV4FpUSPFAqjcBw9QtEX
# hTPP3w/a0IzldsVndCiP08uDeuAVevSgkSF+Ha2pSuFMl3Xf6Lj996T3NaJyiyGX
# BeAW7TTZlYFXMBIQW6oQPjyrK6Vn/aMYkFy1r4V2TaWg/YrehKPg9BB7UzPNVk7n
# YBc7jYweWGbdIejf9GFD4jUDQ3L724B6GRAfouvGStU29kbh/Q8AoxupRxcbvHOc
# onTHQdivlrJYZscplFw5tT7/fhmkv02tc551UNeZJ3bKUpKX+++LVDA0mpcmX/6A
# mRAR62qYcBQVCQW16aLwxRdAbbD9EMddfBYCMT6ogNktD+TjPZnbXq1ZpHpEMoca
# TB4KgO1C3OQdMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkq
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
# A1AwggI4AgEBMIH5oYHRpIHOMIHLMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046N0YwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMV
# AARrR/XXxccz9U12ooGzhBfE2c33oIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDr2TrbMCIYDzIwMjUwNTIyMDU1
# NTM5WhgPMjAyNTA1MjMwNTU1MzlaMHcwPQYKKwYBBAGEWQoEATEvMC0wCgIFAOvZ
# OtsCAQAwCgIBAAICERACAf8wBwIBAAICElswCgIFAOvajFsCAQAwNgYKKwYBBAGE
# WQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDAN
# BgkqhkiG9w0BAQsFAAOCAQEA0Spxv7e13YRDjRkicWov22Y6v56qaETg+q7jYArD
# AVoHhHUj8yCF+iD4/Hq9cWoeeLR7pDthZJavnuhIx1M+f/w4o4bgDROSvIRy71iO
# LZaRiB9hcPYdrqWdQz+i05iI1pFRI40VsyaYeL49WxXVW9dBl+iGXLaMH1jEI6Fs
# WVUEhKhpL8uzyOjgiCv4ueaLFqSgMGkmBVpQN2T2josWDJOkIPQXfewZWwLlvBY5
# BBWIEmakddUR1vygLoUDGwZyidlBO7BkKqlRq3kFtBLoFfKO2zND9HcO7UV32YS4
# BcwTnnwy8nS6i3yAYq3vYBL9bAM5puRy8QbB5DCaaF7yrTGCBA0wggQJAgEBMIGT
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAACBte8UTiYI+wsAAEA
# AAIGMA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQ
# AQQwLwYJKoZIhvcNAQkEMSIEID5MLs2E8G2ruGs9LumdnWqdtrlMSiU3Z0b+pi9H
# nzF1MIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQg4Oj1lIiRnp1W0pP4T+5n
# HZYDLsqJczlHUkg6E0l/S9IwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0EgMjAxMAITMwAAAgbXvFE4mCPsLAABAAACBjAiBCBBvylFNg5kwTZpdlwWrNC/
# EmkgrCtNFpmgmu86EI0bhzANBgkqhkiG9w0BAQsFAASCAgAP8I0Kjj/WW65qSPvC
# 1YqAxF8rAu5cy9kKhgetM1aqVfozi7ngapbDbNVr2bJ9AQWos/80B43fmsL0atDa
# TTUqPbFAgRNSxh3tP6uLKfdsXJkb6vXvYGc60wS6ouHfL9V2OSjCLZshqPEHgXD3
# m87YIVMVRI6BDdzQyFWE+gvYAmp/fWUpXVnikNaTbpUbLXz5ulqdmGPZvScafigb
# pQS7qMzQwpUpv81PifBKl6P3bhMJneDzMlwvEDjiiPmzs7THr2BIBZCov6aSJueJ
# p4fJ+E4ZB8K2uuUyTeMh1c1ObrGks8BVY3cYwUbq4DxkiCHii/84frFLA9Sp5epl
# cdWnz46oPA8ZSObHjXMYV83sNmD+kRw9QUISAAFC0IfBrAxeXxThmHG6bRhhb84O
# OP3IbgaMk/9dazxr4apFoxSmmt4I+V4yOpzxmIoeSKhhYMEm8BVhBMwQQHdvD0mj
# kDnSqClHV0XmlLw2CK63jyiN53io0fnEybkGd+LspkVLRag0HbLdrw1vWrPbrD/l
# OEnFUvSyjWe5q5NQPrTW6IA5acUc1UoiJ7yRY++uhw1G2Z6ri7q5n9DM4GB8THZD
# dNgRD2dGWzwva3kaSy3g3v5AQrdBLRPbK8PftNbrIdF0hMCIlS9rk+d1do+6sGal
# amFYv+ew8xjbzPEck0Yv6ee0mA==
# SIG # End signature block
