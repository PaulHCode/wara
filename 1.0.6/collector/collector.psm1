using module ../utils/utils.psd1

<#
.SYNOPSIS
    Retrieves all resources with matching tags.

.DESCRIPTION
    The Get-WAFTaggedResources function queries Azure Resource Graph to retrieve all resources that have matching tags.

.PARAMETER tagArray
    An array of tags to filter resources by. Each tag should be in the format 'key==value'.

.PARAMETER SubscriptionIds
    An array of subscription IDs to scope the query.

.OUTPUTS
    Returns an array of resources with matching tags.

.EXAMPLE
    $taggedResources = Get-WAFTaggedResources -tagArray @('env==prod', 'app==myapp') -SubscriptionIds @('sub1', 'sub2')

.NOTES
    This function uses the Invoke-WAFQuery function to perform the query.
#>
function Get-WAFTaggedResource {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $TagArray,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]] $SubscriptionIds
    )

    $return = @()

    foreach ($tag in $TagArray) {
        switch -Wildcard ($tag) {
            "*=~*" {
                $tagKeys = $tag.Split("=~")[0].split("||") -join ("','")
                $tagValues = $tag.Split("=~")[1].split("||") -join ("','")
                $in = "in~"
            }
            "*!~*" {
                $tagKeys = $tag.Split("!~")[0].split("||") -join ("','")
                $tagValues = $tag.Split("!~")[1].split("||") -join ("','")
                $in = "!in~"
            }
        }

        $tagquery = "resources
| mv-expand bagexpansion=array tags
| where isnotempty(tags)
| where tolower(tags[0]) in~ ('$tagkeys')  // Specify your tag names here
| where tolower(tags[1]) $in ('$tagvalues')  // Specify your tag values here
| summarize by id
| order by ['id']"

        $result = Invoke-WAFQuery -Query $tagquery -SubscriptionIds $SubscriptionIds

        $return += $result
    }

    $return = ($return | Group-Object id | Where-Object { $_.count -eq $TagArray.Count } | Select-Object Name).Name

    return $return
}

<#
.SYNOPSIS
    Retrieves all resources in resource groups with matching tags.

.DESCRIPTION
    The Get-WAFTaggedRGResources function queries Azure Resource Graph to retrieve all resources in resource groups that have matching tags.

.PARAMETER tagKeys
    An array of tag keys to filter resource groups by.

.PARAMETER tagValues
    An array of tag values to filter resource groups by.

.PARAMETER SubscriptionIds
    An array of subscription IDs to scope the query.

.OUTPUTS
    Returns an array of resources in resource groups with matching tags.

.EXAMPLE
    $taggedRGResources = Get-WAFTaggedRGResources -tagKeys @('env') -tagValues @('prod') -SubscriptionIds @('sub1', 'sub2')

.NOTES
    This function uses the Invoke-WAFQuery function to perform the query.
#>
function Get-WAFTaggedResourceGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]] $TagArray,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]] $SubscriptionIds
    )

    $return = @()

    foreach ($tag in $TagArray) {
        switch -Wildcard ($tag) {
            "*=~*" {
                $tagKeys = $tag.Split("=~")[0].split("||") -join ("','")
                $tagValues = $tag.Split("=~")[1].split("||") -join ("','")
                $in = "in~"
            }
            "*!~*" {
                $tagKeys = $tag.Split("!~")[0].split("||") -join ("','")
                $tagValues = $tag.Split("!~")[1].split("||") -join ("','")
                $in = "!in~"
            }
        }

        $tagquery = `
            "resourcecontainers
| where type == 'microsoft.resources/subscriptions/resourcegroups'
| mv-expand bagexpansion=array tags
| where isnotempty(tags)
| where tolower(tags[0]) in~ ('$tagKeys')  // Specify your tag names here
| where tolower(tags[1]) $in ('$tagValues')  // Specify your tag values here
| summarize by id
| order by ['id']"

        $result = Invoke-WAFQuery -Query $tagquery -SubscriptionIds $SubscriptionIds

        $return += $result
    }

    $return = ($return | Group-Object id | Where-Object { $_.count -eq $TagArray.Count } | Select-Object Name).Name

    return $return
}

<#
.SYNOPSIS
    Invokes a loop to run queries for each recommendation object.

.DESCRIPTION
    The Invoke-WAFQueryLoop function runs queries for each recommendation object and retrieves the resources.

.PARAMETER RecommendationObject
    An array of recommendation objects to query.

.PARAMETER subscriptionIds
    An array of subscription IDs to scope the query.

.OUTPUTS
    Returns an array of resources for each recommendation object.

.EXAMPLE
    $resources = Invoke-WAFQueryLoop -RecommendationObject $recommendations -subscriptionIds @('sub1', 'sub2')

.NOTES
    This function uses the Invoke-WAFQuery function to perform the queries.
#>
function Invoke-WAFQueryLoop {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]] $RecommendationObject,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]] $SubscriptionIds,

        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [string[]] $AddedTypes,

        [Parameter(Mandatory = $false)]
        [int] $ProgressId = 1
    )

    $Types = Get-WAFResourceType -SubscriptionIds $SubscriptionIds

    $QueryObject = Get-WAFQueryByResourceType -ObjectList $RecommendationObject -FilterList $Types.type -KeyColumn 'recommendationResourceType'

    # Add additional types to query based on specialized workloads (This works even if it's empty.)
    $QueryObject += $AddedTypes.Foreach({
        $type = $_
        $RecommendationObject.where({$_.tags -contains $type})
    }) | Sort-Object -Property "APRLGuid" | Get-Unique -AsString

    $return = $QueryObject.Where({ $_.automationAvailable -eq $true -and $_.recommendationMetadataState -eq "Active" -and [string]::IsNullOrEmpty($_.recommendationTypeId) }) | ForEach-Object {
        Write-Progress -Activity 'Running Queries' -Status "Running Query for $($_.recommendationResourceType) - $($_.aprlGuid)" -PercentComplete (($QueryObject.IndexOf($_) / $QueryObject.Count) * 100) -Id $ProgressId
        try {
            $recommendation = $_
            (Invoke-WAFQuery -Query $recommendation.query -SubscriptionIds $subscriptionIds -ErrorAction Stop)
        }
        catch {
            Write-Error "Error running query for - $($recommendation.recommendationResourceType) - $($recommendation.aprlGuid)"
        }
    }
    Write-Progress -Activity 'Running Queries' -Status 'Completed' -Completed -Id $ProgressId

    return $return
}

<#
.SYNOPSIS
    Retrieves all resource types in the specified subscriptions.

.DESCRIPTION
    The Get-WAFResourceType function queries Azure Resource Graph to retrieve all resource types in the specified subscriptions.

.PARAMETER SubscriptionIds
    An array of subscription IDs to scope the query.

.OUTPUTS
    Returns an array of resource types.

.EXAMPLE
    $resourceTypes = Get-WAFResourceType -SubscriptionIds @('sub1', 'sub2')

.NOTES
    This function uses the Invoke-WAFQuery function to perform the query.
#>
function Get-WAFResourceType {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]] $SubscriptionIds
    )

    $q = "Resources
| summarize count() by type
| project type"

    $r = $SubscriptionIds ? (Invoke-WAFQuery -Query $q -SubscriptionIds $SubscriptionIds) : (Invoke-WAFQuery -Query $q)

    return $r
}

<#
.SYNOPSIS
    Filters objects by resource type.

.DESCRIPTION
    The Get-WAFQueryByResourceType function filters a list of objects by resource type.

.PARAMETER ObjectList
    An array of objects to filter.

.PARAMETER FilterList
    An array of resource types to filter by.

.PARAMETER KeyColumn
    The key column to use for filtering.

.OUTPUTS
    Returns an array of objects that match the specified resource types.

.EXAMPLE
    $filteredObjects = Get-WAFQueryByResourceType -ObjectList $objects -FilterList $types -KeyColumn "type"
#>
function Get-WAFQueryByResourceType {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCustomObject[]] $ObjectList,

        [Parameter(Mandatory = $true)]
        [string[]] $FilterList,

        [Parameter(Mandatory = $true)]
        [string] $KeyColumn
    )

    $matchingObjects = foreach ($obj in $ObjectList) {
        if ($obj.$KeyColumn -in $FilterList) {
            $obj
        }
    }

    return $matchingObjects
}

# SIG # Begin signature block
# MIIoOwYJKoZIhvcNAQcCoIIoLDCCKCgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBsrKdM1BMNiPM5
# iZvbfeo5szkgb1kyY/YEamvMrq0Fq6CCDYUwggYDMIID66ADAgECAhMzAAAEA73V
# lV0POxitAAAAAAQDMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjQwOTEyMjAxMTEzWhcNMjUwOTExMjAxMTEzWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCfdGddwIOnbRYUyg03O3iz19XXZPmuhEmW/5uyEN+8mgxl+HJGeLGBR8YButGV
# LVK38RxcVcPYyFGQXcKcxgih4w4y4zJi3GvawLYHlsNExQwz+v0jgY/aejBS2EJY
# oUhLVE+UzRihV8ooxoftsmKLb2xb7BoFS6UAo3Zz4afnOdqI7FGoi7g4vx/0MIdi
# kwTn5N56TdIv3mwfkZCFmrsKpN0zR8HD8WYsvH3xKkG7u/xdqmhPPqMmnI2jOFw/
# /n2aL8W7i1Pasja8PnRXH/QaVH0M1nanL+LI9TsMb/enWfXOW65Gne5cqMN9Uofv
# ENtdwwEmJ3bZrcI9u4LZAkujAgMBAAGjggGCMIIBfjAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQU6m4qAkpz4641iK2irF8eWsSBcBkw
# VAYDVR0RBE0wS6RJMEcxLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJh
# dGlvbnMgTGltaXRlZDEWMBQGA1UEBRMNMjMwMDEyKzUwMjkyNjAfBgNVHSMEGDAW
# gBRIbmTlUAXTgqoXNzcitW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
# d3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIw
# MTEtMDctMDguY3JsMGEGCCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDov
# L3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDEx
# XzIwMTEtMDctMDguY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIB
# AFFo/6E4LX51IqFuoKvUsi80QytGI5ASQ9zsPpBa0z78hutiJd6w154JkcIx/f7r
# EBK4NhD4DIFNfRiVdI7EacEs7OAS6QHF7Nt+eFRNOTtgHb9PExRy4EI/jnMwzQJV
# NokTxu2WgHr/fBsWs6G9AcIgvHjWNN3qRSrhsgEdqHc0bRDUf8UILAdEZOMBvKLC
# rmf+kJPEvPldgK7hFO/L9kmcVe67BnKejDKO73Sa56AJOhM7CkeATrJFxO9GLXos
# oKvrwBvynxAg18W+pagTAkJefzneuWSmniTurPCUE2JnvW7DalvONDOtG01sIVAB
# +ahO2wcUPa2Zm9AiDVBWTMz9XUoKMcvngi2oqbsDLhbK+pYrRUgRpNt0y1sxZsXO
# raGRF8lM2cWvtEkV5UL+TQM1ppv5unDHkW8JS+QnfPbB8dZVRyRmMQ4aY/tx5x5+
# sX6semJ//FbiclSMxSI+zINu1jYerdUwuCi+P6p7SmQmClhDM+6Q+btE2FtpsU0W
# +r6RdYFf/P+nK6j2otl9Nvr3tWLu+WXmz8MGM+18ynJ+lYbSmFWcAj7SYziAfT0s
# IwlQRFkyC71tsIZUhBHtxPliGUu362lIO0Lpe0DOrg8lspnEWOkHnCT5JEnWCbzu
# iVt8RX1IV07uIveNZuOBWLVCzWJjEGa+HhaEtavjy6i7MIIHejCCBWKgAwIBAgIK
# YQ6Q0gAAAAAAAzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEw
# OTA5WjB+MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYD
# VQQDEx9NaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+la
# UKq4BjgaBEm6f8MMHt03a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc
# 6Whe0t+bU7IKLMOv2akrrnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4D
# dato88tt8zpcoRb0RrrgOGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+
# lD3v++MrWhAfTVYoonpy4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nk
# kDstrjNYxbc+/jLTswM9sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6
# A4aN91/w0FK/jJSHvMAhdCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmd
# X4jiJV3TIUs+UsS1Vz8kA/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL
# 5zmhD+kjSbwYuER8ReTBw3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zd
# sGbiwZeBe+3W7UvnSSmnEyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3
# T8HhhUSJxAlMxdSlQy90lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS
# 4NaIjAsCAwEAAaOCAe0wggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRI
# bmTlUAXTgqoXNzcitW2oynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAL
# BgNVHQ8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBD
# uRQFTuHqp8cx0SOJNDBaBgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jv
# c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3JsMF4GCCsGAQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFf
# MDNfMjIuY3J0MIGfBgNVHSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEF
# BQcCARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1h
# cnljcHMuaHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkA
# YwB5AF8AcwB0AGEAdABlAG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn
# 8oalmOBUeRou09h0ZyKbC5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7
# v0epo/Np22O/IjWll11lhJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0b
# pdS1HXeUOeLpZMlEPXh6I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/
# KmtYSWMfCWluWpiW5IP0wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvy
# CInWH8MyGOLwxS3OW560STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBp
# mLJZiWhub6e3dMNABQamASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJi
# hsMdYzaXht/a8/jyFqGaJ+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYb
# BL7fQccOKO7eZS/sl/ahXJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbS
# oqKfenoi+kiVH6v7RyOA9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sL
# gOppO6/8MO0ETI7f33VtY5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtX
# cVZOSEXAQsmbdlsKgEhr/Xmfwb1tbWrJUnMTDXpQzTGCGgwwghoIAgEBMIGVMH4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01p
# Y3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTECEzMAAAQDvdWVXQ87GK0AAAAA
# BAMwDQYJYIZIAWUDBAIBBQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEINQu
# vswNFEPGKlFDTDlql4ygw4Qcyo4itjqkYgrLmTZXMEQGCisGAQQBgjcCAQwxNjA0
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQu
# Y29tIDANBgkqhkiG9w0BAQEFAASCAQCHhHMQW5/PqxiROvyix4lPr7KOXbY5fZwA
# dkglm2DoFeL1aeFEL0gViznT6ph8fS8lG1e8m2ofXYSboKo/8EJ9IEUlZXVPPJtX
# GcbgN7L/hAvtWvUseNJ9LuTfj0bZo6qB/x4QItfZKqGxrPMmafwSzVwqX7LnlNAb
# t1+7EsA1Z5jn5gNuFWFERkSel6alODYnOOnkSrzCGGc9hP+VOZJlW+oehiH5IKHL
# MdzsihpBgiW8exJNkd4pFKJJP5FcHRvhOa+gu6JibQ6f03bET+1h3FL1kxrnH13N
# a1E/6YPdyCCCygnN5G9GKuD4Et3Sv6u4jNoxBZsgMzHCnBBF3sFioYIXlDCCF5AG
# CisGAQQBgjcDAwExgheAMIIXfAYJKoZIhvcNAQcCoIIXbTCCF2kCAQMxDzANBglg
# hkgBZQMEAgEFADCCAVIGCyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEE
# AYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIHDLM2r7Vo7ljoelRCGzqhhPHF0JPo87
# 0KQnk19yxX/pAgZoJmwIGScYEzIwMjUwNTIyMTUwMDQ5LjU3MlowBIACAfSggdGk
# gc4wgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNV
# BAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGll
# bGQgVFNTIEVTTjpFMDAyLTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgU2VydmljZaCCEeowggcgMIIFCKADAgECAhMzAAACCxGdVimS+b+F
# AAEAAAILMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAy
# MDEwMB4XDTI1MDEzMDE5NDI1OFoXDTI2MDQyMjE5NDI1OFowgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpFMDAy
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKqs+K1GMBeoWHYgfNBg
# Pe7KQM/IkX7v17kKRjhpixv3dBW6fh9ncm7/RYeu0AXcmzBKIBWXunUfbMUqu/KF
# pW5FhxJzti7I/QYcRWrthVLm4XdguUt6jIY48pJcazpqBbrQKpUGE9D7afj17pgP
# 3Aa/pjQeYEqadCpj8AomKOBSJG6VvWi80RxNmZM1G7WZj7QmOIS6z54XrpU2gG/V
# rYxbQk9hsUGE/MUeV10VdyN/aoiRvkFhawUvtCUGNcCl/YeY2s/MbfqDXJuj8qdG
# NKAhGkd5hfEEhoIQa1HUSakn3Q5IUfFS7t4iNV96eqdW1qDTIBBjvMRZJcw3r4Ie
# O/dIE22blkhHLIKxRkbzjngr7zrlvDsXC4fy+TlD7TsrmUVjTZ4EPDdQrNaNa+pO
# nakrwxNriSjO+UqCtIanmyHnsaeSOW82+3vw1dtlNUivBLvFwgaNu9L/avmVENP5
# Dc/a1P6DooM90ue2VxsV0e67PInLHRZ8KvzMe/zAYIHeo85tFcriikAidJDiPmJL
# bwoBMkbBBptG55m+G93053TBwVg5viitS/V0PlfAXIqUYM8xxM85CvSDGYRBWsuI
# kSgvLJbwRX9oMVRdOjHKF/NZnoNzSX1562HUS3Y1DcD+oFHnHfemMp0clUnnGmnY
# AjdnM1gsMKo7WW2HPUZpLP49AgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUZlwsiN+v
# 6XpZ4y6ET2CzF73yP/4wHwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIw
# XwYDVR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9w
# cy9jcmwvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3Js
# MGwGCCsGAQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENB
# JTIwMjAxMCgxKS5jcnQwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEF
# BQcDCDAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAJoo9WfoYtye
# Rs8R0zRH2rhmbT7sl8bTn6mao/qcDhLBPbqT0eUlCOrD/5aSJAsPijAnGigiv0wN
# luJlfrKk+EX6DU/m7wE4Rc9hiQ5anMUYtEu2Vke9/GmoAhruaMsYcPMqv/XTYBdo
# PiYbJ0TBosaFSf0MEZ0ybl2GKKxYCpv8s3iWbg65fhguopxb/TJJwcqCmsaAtj7q
# mq/TM9TviPt4Kv48NM1gzy0Un9u1ScE0vI4ThZnAEiri2e9eN3ZwI62BSehCMVTV
# ljtlZ2fLcElt/eEGtvc0HYhTie7rba1WtvV82TB29PnTVZFQEx32gV/jGJ+PaPYK
# Vu/VcaWNeZ3rImO+35sG68ktS5z2b68wo5bJHBR286X8TlrzNXcyuMKdjp8istA3
# ME5nqSUgI8KL7YaptWIzQIjoJtnKl6LlGv2ElsfytkTONphYuev9c+xTTCvZOfov
# 10nHar8sCgr6i+IAiePRo/iJGszJRI7oovYrBjrMlzpIPFseTBgd1BGs80/dK60y
# DouAFK+2Z+rCBpcgeQKiJQUQDP1wYiq+CMakamkpNxO6ijcQSCn/NF3TDFbRCqZX
# fKBgaJx6ffQoBYesoULxJvBzqoLimh1VlglmaKscE/+6juxCd7kfTG128OZcNR/L
# mA2MuWNAhSwmFBLLJdtKNJvZrmOE5R/HMIIHcTCCBVmgAwIBAgITMwAAABXF52ue
# AptJmQAAAAAAFTANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlm
# aWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMjEwOTMwMTgyMjI1WhcNMzAwOTMwMTgz
# MjI1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
# BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYD
# VQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBAOThpkzntHIhC3miy9ckeb0O1YLT/e6cBwfSqWxO
# dcjKNVf2AX9sSuDivbk+F2Az/1xPx2b3lVNxWuJ+Slr+uDZnhUYjDLWNE893MsAQ
# GOhgfWpSg0S3po5GawcU88V29YZQ3MFEyHFcUTE3oAo4bo3t1w/YJlN8OWECesSq
# /XJprx2rrPY2vjUmZNqYO7oaezOtgFt+jBAcnVL+tuhiJdxqD89d9P6OU8/W7IVW
# Te/dvI2k45GPsjksUZzpcGkNyjYtcI4xyDUoveO0hyTD4MmPfrVUj9z6BVWYbWg7
# mka97aSueik3rMvrg0XnRm7KMtXAhjBcTyziYrLNueKNiOSWrAFKu75xqRdbZ2De
# +JKRHh09/SDPc31BmkZ1zcRfNN0Sidb9pSB9fvzZnkXftnIv231fgLrbqn427DZM
# 9ituqBJR6L8FA6PRc6ZNN3SUHDSCD/AQ8rdHGO2n6Jl8P0zbr17C89XYcz1DTsEz
# OUyOArxCaC4Q6oRRRuLRvWoYWmEBc8pnol7XKHYC4jMYctenIPDC+hIK12NvDMk2
# ZItboKaDIV1fMHSRlJTYuVD5C4lh8zYGNRiER9vcG9H9stQcxWv2XFJRXRLbJbqv
# UAV6bMURHXLvjflSxIUXk8A8FdsaN8cIFRg/eKtFtvUeh17aj54WcmnGrnu3tz5q
# 4i6tAgMBAAGjggHdMIIB2TASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsGAQQBgjcV
# AgQWBBQqp1L+ZMSavoKRPEY1Kc8Q/y8E7jAdBgNVHQ4EFgQUn6cVXQBeYl2D9OXS
# ZacbUzUZ6XIwXAYDVR0gBFUwUzBRBgwrBgEEAYI3TIN9AQEwQTA/BggrBgEFBQcC
# ARYzaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9Eb2NzL1JlcG9zaXRv
# cnkuaHRtMBMGA1UdJQQMMAoGCCsGAQUFBwMIMBkGCSsGAQQBgjcUAgQMHgoAUwB1
# AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaA
# FNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dF8y
# MDEwLTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIwMTAt
# MDYtMjMuY3J0MA0GCSqGSIb3DQEBCwUAA4ICAQCdVX38Kq3hLB9nATEkW+Geckv8
# qW/qXBS2Pk5HZHixBpOXPTEztTnXwnE2P9pkbHzQdTltuw8x5MKP+2zRoZQYIu7p
# Zmc6U03dmLq2HnjYNi6cqYJWAAOwBb6J6Gngugnue99qb74py27YP0h1AdkY3m2C
# DPVtI1TkeFN1JFe53Z/zjj3G82jfZfakVqr3lbYoVSfQJL1AoL8ZthISEV09J+BA
# ljis9/kpicO8F7BUhUKz/AyeixmJ5/ALaoHCgRlCGVJ1ijbCHcNhcy4sa3tuPywJ
# eBTpkbKpW99Jo3QMvOyRgNI95ko+ZjtPu4b6MhrZlvSP9pEB9s7GdP32THJvEKt1
# MMU0sHrYUP4KWN1APMdUbZ1jdEgssU5HLcEUBHG/ZPkkvnNtyo4JvbMBV0lUZNlz
# 138eW0QBjloZkWsNn6Qo3GcZKCS6OEuabvshVGtqRRFHqfG3rsjoiV5PndLQTHa1
# V1QJsWkBRH58oWFsc/4Ku+xBZj1p/cvBQUl+fpO+y/g75LcVv7TOPqUxUYS8vwLB
# gqJ7Fx0ViY1w/ue10CgaiQuPNtq6TPmb/wrpNPgkNWcr4A245oyZ1uEi6vAnQj0l
# lOZ0dFtq0Z4+7X6gMTN9vMvpe784cETRkPHIqzqKOghif9lwY1NNje6CbaUFEMFx
# BmoQtB1VM1izoXBm8qGCA00wggI1AgEBMIH5oYHRpIHOMIHLMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046RTAwMi0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wi
# IwoBATAHBgUrDgMCGgMVAKhCd1Qk3c+mrbGHxGG2xHfk0JnBoIGDMIGApH4wfDEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
# cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDr2XvJ
# MCIYDzIwMjUwNTIyMTAzMjQxWhgPMjAyNTA1MjMxMDMyNDFaMHQwOgYKKwYBBAGE
# WQoEATEsMCowCgIFAOvZe8kCAQAwBwIBAAICOt8wBwIBAAICEiUwCgIFAOvazUkC
# AQAwNgYKKwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEK
# MAgCAQACAwGGoDANBgkqhkiG9w0BAQsFAAOCAQEAjLBRDG7uJ+9VEuvhSBw+DxuQ
# ZSrOhkpFjpzVXZ9aYVwuLyqcQXnT52lWM8wqaAcBH3AVS1T0l9gPtv8ZO+F+l644
# NSsUizTvn93xYouEBs6ZJqd3EbKMKmp7RYKkcThJ1omn1kXKyAhaVTk4gc7IVyHP
# obcsBXpK33Jl/WF0MrlPSzFB7J0XusHzHsMWyH0rr7NL4yDsuG8yRNQv7O7sCFnm
# GVp+ePSHKmWkLUoFPhFVMlOz4uQxypQFYbVBdocsz7/ufW4SAj1pgZ12hT8v5LHp
# E4SHQb1uK6rUMZtTBlZLYFIFQ3cNoXXAPwkEvY6f7z+O2gwPVD4WWKyRJBJdczGC
# BA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAC
# CxGdVimS+b+FAAEAAAILMA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMx
# DQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIIUwvguWWDDwanQw4wcUwIu0
# C8XtAIqmWuDWdIElegBsMIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQgNNV0
# q7o3vtuHTB07IX6iBE3y8olzDmOd/b8S6yxiSjMwgZgwgYCkfjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAgsRnVYpkvm/hQABAAACCzAiBCAEP1Gi
# UtIVayJzIES1PXhTo9uDRpXC0PropG2gldwFZzANBgkqhkiG9w0BAQsFAASCAgA9
# pQsBPXDPWHv5iElqF90ydmk83YnhC8F5ehe4xFR5GsAwg67yuC2hfJUI27fCNg/4
# CfDIoLslFnoc9MJa4KLPm8aJrMJZ+O9TV5Ag+3ChPLvtblyPghuaP/mabjGcA9z2
# WdyayGNZNYdBXLPCkT7KAvdH/rGw9RsnqjgCTNA+YjxQVS+p3M5VWnBbqKqmdNRu
# lDB1oY3VI9znJcYMk54buiV6o/7xsVfeYUZPzKbyNC6/YVaVYBbOcbALpeIa4LGC
# qpkGwCtfyfQ/59gVmmWO5CcVn2bzWjp5lBUjWo8bR4HkwfVsBzzC/Ql8ec7PWGeC
# QxmA0NdfppEDOWz3hxO9vLb9HgmA8EadGCF6D9Ke1nnvxtPcq4k/OWrbXZhRZjwJ
# xEMAC6L95PqJJbw1TVqLRXektEEzSuNT/B7qQcgu11iOs2uuVeEabkS2+adUUth7
# 7UPZw3C0XnWk/G4AX/pwc+JMhrGqmw7ddpNWylWJvMCFBVc6sQy14dBRN4iBzb8n
# Ab/TG8n4VEVbE40SHYTovzpfN09+5k9tYFzR5sFRUBlfb8U94dIaHQjdVEJQK5RR
# FzjlXsiR5im4yWJ6MRgvQsGc29zosp2obqjSp0aL+FPkVDu18A+UThMxPTrlzNQR
# QEVT3/pG7RmWvaQ454sRYlLzuPHPMQwnnpEJUoxwlg==
# SIG # End signature block
