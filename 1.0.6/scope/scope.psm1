<#
.SYNOPSIS
    Retrieves filtered lists of Azure resources based on provided subscription, resource group, and resource filters.

.DESCRIPTION
    This module contains functions to filter Azure resources by subscription, resource group, and resource IDs.
    It includes the following functions:
    - Get-WAFResourceGroupsByList
    - Get-WAFSubscriptionsByList
    - Get-WAFResourcesByList
    - Get-WAFFilteredResourceList

.EXAMPLE
    $subscriptionFilters = @("/subscriptions/12345")
    $resourceGroupFilters = @("/subscriptions/12345/resourceGroups/myResourceGroup")
    $resourceFilters = @("/subscriptions/12345/resourceGroups/myResourceGroup/providers/Microsoft.Compute/virtualMachines/myVM")

    $filteredResources = Get-WAFFilteredResourceList -SubscriptionFilters $subscriptionFilters -ResourceGroupFilters $resourceGroupFilters -ResourceFilters $resourceFilters

.NOTES
    Author: Kyle Poineal
    Date: 2024-08-07
#>

<#
.SYNOPSIS
    Filters a list of objects based on resource group identifiers.

.DESCRIPTION
    The Get-WAFResourceGroupsByList function takes a list of objects and filters them based on the specified resource group identifiers.
    It compares the first five segments of the KeyColumn property of each object with the provided filter list.

.PARAMETER ObjectList
    An array of objects to be filtered.

.PARAMETER FilterList
    An array of resource group identifiers to filter the objects.

.PARAMETER KeyColumn
    The name of the property in the objects that contains the resource group identifier.

.OUTPUTS
    Returns an array of filtered resources.

.EXAMPLE
    $objectList = @(
        @{ Id = "/subscriptions/12345/resourceGroups/myResourceGroup/providers/Microsoft.Compute/virtualMachines/myVM" },
        @{ Id = "/subscriptions/12345/resourceGroups/anotherResourceGroup/providers/Microsoft.Compute/virtualMachines/anotherVM" }
    )
    $filterList = @("/subscriptions/12345/resourceGroups/myResourceGroup")

    $filteredObjects = Get-WAFResourceGroupsByList -ObjectList $objectList -FilterList $filterList -KeyColumn "Id"

.NOTES
    Author: Kyle Poineal
    Date: 2024-08-07
#>
function Get-WAFResourceGroupsByList {
    param (
        [Parameter(Mandatory = $true)]
        [array] $ObjectList,

        [Parameter(Mandatory = $true)]
        [array] $FilterList,

        [Parameter(Mandatory = $true)]
        [string] $KeyColumn
    )

    $matchingObjects = @()
    $matchingObjects += foreach ($obj in $ObjectList) {
        if (($obj.$KeyColumn.split('/')[0..4] -join '/') -in $FilterList) {
            $obj
        }
    }

    return ,$matchingObjects
}

<#
.SYNOPSIS
    Filters a list of objects based on subscription identifiers.

.DESCRIPTION
    The Get-WAFSubscriptionsByList function takes a list of objects and filters them based on the specified subscription identifiers.
    It compares the first three segments of the KeyColumn property of each object with the provided filter list.

.PARAMETER ObjectList
    An array of objects to be filtered.

.PARAMETER FilterList
    An array of subscription identifiers to filter the objects.

.PARAMETER KeyColumn
    The name of the property in the objects that contains the subscription identifier.

.OUTPUTS
    Returns an array of filtered resources.

.EXAMPLE
    $objectList = @(
        @{ Id = "/subscriptions/12345/resourceGroups/myResourceGroup/providers/Microsoft.Compute/virtualMachines/myVM" },
        @{ Id = "/subscriptions/67890/resourceGroups/anotherResourceGroup/providers/Microsoft.Compute/virtualMachines/anotherVM" }
    )
    $filterList = @("/subscriptions/12345")

    $filteredObjects = Get-WAFSubscriptionsByList -ObjectList $objectList -FilterList $filterList -KeyColumn "Id"

.NOTES
    Author: Kyle Poineal
    Date: 2024-08-07
#>
function Get-WAFSubscriptionsByList {
    param (
        [Parameter(Mandatory = $true)]
        [array] $ObjectList,

        [Parameter(Mandatory = $true)]
        [array] $FilterList,

        [Parameter(Mandatory = $true)]
        [string] $KeyColumn
    )

    $matchingObjects = @()
    $matchingObjects += foreach ($obj in $ObjectList) {
        if (($obj.$KeyColumn.split('/')[0..2] -join '/') -in $FilterList) {
            $obj
        }
    }

    return ,$matchingObjects
}

<#
.SYNOPSIS
    Filters a list of objects based on resource identifiers.

.DESCRIPTION
    The Get-WAFResourcesByList function takes a list of objects and filters them based on the specified resource identifiers.
    It compares the KeyColumn property of each object with the provided filter list.

.PARAMETER ObjectList
    An array of objects to be filtered.

.PARAMETER FilterList
    An array of resource identifiers to filter the objects.

.PARAMETER KeyColumn
    The name of the property in the objects that contains the resource identifier.

.OUTPUTS
    Returns an array of filtered resources.

.EXAMPLE
    $objectList = @(
        @{ Id = "/subscriptions/12345/resourceGroups/myResourceGroup/providers/Microsoft.Compute/virtualMachines/myVM" },
        @{ Id = "/subscriptions/12345/resourceGroups/myResourceGroup/providers/Microsoft.Compute/virtualMachines/anotherVM" }
    )
    $filterList = @("/subscriptions/12345/resourceGroups/myResourceGroup/providers/Microsoft.Compute/virtualMachines/myVM")

    $filteredObjects = Get-WAFResourcesByList -ObjectList $objectList -FilterList $filterList -KeyColumn "Id"

.NOTES
    Author: Your Name
    Date: 2024-08-07
#>
function Get-WAFResourcesByList {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array] $ObjectList,

        [Parameter(Mandatory = $true)]
        [array] $FilterList,

        [Parameter(Mandatory = $true)]
        [string] $KeyColumn,

        [parameter(Mandatory = $false)]
        [switch] $NotIn
    )

    $matchingObjects = @()

    if($NotIn.IsPresent){
        Write-Debug 'Filtering for objects not in the list'
        $matchingObjects += foreach ($obj in $ObjectList) {
            if ($obj.$KeyColumn -notin $FilterList) {
                $obj
            }
        }
    } else {
        Write-Debug 'Filtering for objects in the list'
        $matchingObjects += foreach ($obj in $ObjectList) {
            if ($obj.$KeyColumn -in $FilterList) {
                $obj
            }
        }
    }

    return ,$matchingObjects
}

<#
.SYNOPSIS
    Creates a list of unique subscription IDs based on provided subscription, resource group, and resource filters.

.DESCRIPTION
    The Get-WAFImplicitSubscriptionId function takes arrays of subscription filters, resource group filters, and resource filters.
    It creates a list of unique subscription IDs based on these filters by combining them, splitting them into subscription IDs, and removing duplicates.

.PARAMETER SubscriptionFilters
    An array of strings representing the subscription filters. Each string should be a subscription ID or a part of a subscription ID.

.PARAMETER ResourceGroupFilters
    An array of strings representing the resource group filters. Each string should be a resource group ID or a part of a resource group ID.

.PARAMETER ResourceFilters
    An array of strings representing the resource filters. Each string should be a resource ID or a part of a resource ID.

.OUTPUTS
    Returns an array of unique subscription IDs.

.EXAMPLE
    $subscriptionFilters = @('/subscriptions/11111111-1111-1111-1111-111111111111')
    $resourceGroupFilters = @('/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/test1')
    $resourceFilters = @('/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/test1/providers/Microsoft.Compute/virtualMachines/TestVM1')
    $implicitSubscriptionIds = Get-WAFImplicitSubscriptionId -SubscriptionFilters $subscriptionFilters -ResourceGroupFilters $resourceGroupFilters -ResourceFilters $resourceFilters

.NOTES
    This function assumes that the input filters are valid and properly formatted.
#>
function Get-WAFImplicitSubscriptionId {
    param (
        [array] $SubscriptionFilters = @(),

        [array] $ResourceGroupFilters = @(),

        [array] $ResourceFilters = @()
    )

    # Create a list of subscription ids based on the filters. Adds all the filters together then splits them into subscription Ids. Groups them to remove duplicates and returns a string array.
    $ImplicitSubscriptionIds = (($SubscriptionFilters + $ResourceGroupFilters + $ResourceFilters) | ForEach-Object { $_.split("/")[0..2] -join "/" } | Group-Object | Select-Object Name).Name
    return $ImplicitSubscriptionIds
}

<#
.SYNOPSIS
    Retrieves a filtered list of Azure resources based on subscription, resource group, and resource filters.

.DESCRIPTION
    The Get-WAFFilteredResourceList function filters Azure resources by combining subscription, resource group, and resource filters.
    It generates a list of implicit subscription IDs from the provided filters, retrieves unfiltered resources, and then applies the filters to return the matching resources.

.PARAMETER SubscriptionFilters
    An array of subscription identifiers to filter the resources.

.PARAMETER ResourceGroupFilters
    An array of resource group identifiers to filter the resources.

.PARAMETER ResourceFilters
    An array of resource identifiers to filter the resources.

.PARAMETER UnfilteredResources
    An array of unfiltered resources to be filtered.

.OUTPUTS
    Returns an array of filtered resources from Azure.

.EXAMPLE
    $subscriptionFilters = @("/subscriptions/12345")
    $resourceGroupFilters = @("/subscriptions/12345/resourceGroups/myResourceGroup")
    $resourceFilters = @("/subscriptions/12345/resourceGroups/myResourceGroup/providers/Microsoft.Compute/virtualMachines/myVM")
    $unfilteredResources = Get-WAFUnfilteredResourceList -ImplicitSubscriptionId (Get-WAFImplicitSubscriptionId -SubscriptionFilters $subscriptionFilters -ResourceGroupFilters $resourceGroupFilters -ResourceFilters $resourceFilters)
    $filteredResources = Get-WAFFilteredResourceList -SubscriptionFilters $subscriptionFilters -ResourceGroupFilters $resourceGroupFilters -ResourceFilters $resourceFilters -UnfilteredResources $unfilteredResources

.NOTES
    Author: Your Name
    Date: 2024-08-07
#>
function Get-WAFFilteredResourceList {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string[]] $SubscriptionFilters = @(),

        [Parameter(Mandatory = $false)]
        [string[]] $ResourceGroupFilters = @(),

        [Parameter(Mandatory = $false)]
        [string[]] $ResourceFilters = @(),

        [Parameter(Mandatory = $true)]
        [array] $UnfilteredResources = @(),

        [Parameter(Mandatory = $false)]
        [string] $KeyColumn = 'Id'
    )

    ##TODO: Replace all of this with something like $scopes.foreach({$allresources -match "$_*"})

    $SubscriptionFilteredResources = @()
    $SubscriptionFilters ? $($SubscriptionFilteredResources = Get-WAFSubscriptionsByList -ObjectList $UnfilteredResources -FilterList $SubscriptionFilters -KeyColumn $KeyColumn) : $(Write-Debug 'Subscription Filters not provided.')

    $ResourceGroupFilteredResources = @()
    $ResourceGroupFilters ? $($ResourceGroupFilteredResources = Get-WAFResourceGroupsByList -ObjectList $UnfilteredResources -FilterList $ResourceGroupFilters -KeyColumn $KeyColumn) : $(Write-Debug 'Resource Group Filters not provided.')

    $ResourceFilteredResources = @()
    $ResourceFilters ? $($ResourceFilteredResources = Get-WAFResourcesByList -ObjectList $UnfilteredResources -FilterList $ResourceFilters -KeyColumn $KeyColumn) : $(Write-Debug 'Resource Filters not provided.')

    $FilteredResources = @()

    #$FilteredResources += $SubscriptionFilteredResources + $ResourceGroupFilteredResources + $ResourceFilteredResources | Sort-Object | Get-Unique -CaseInsensitive -AsString
    $FilteredResources += $SubscriptionFilteredResources + $ResourceGroupFilteredResources + $ResourceFilteredResources

    if($FilteredResources.Count -eq 0){
        Write-Debug 'No resources found matching the provided filters.'
        return ,$FilteredResources
    }

    $FilteredResources = [System.Linq.Enumerable]::Distinct([object[]]$FilteredResources).toArray()
    return ,$FilteredResources
}

# SIG # Begin signature block
# MIIoOwYJKoZIhvcNAQcCoIIoLDCCKCgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBQ3lt7BQw4eWPg
# lB1rAHHRHBbgCrTPHr7kfl4bNeFE3qCCDYUwggYDMIID66ADAgECAhMzAAAEA73V
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
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIJ02
# FPj3TdhNeo4lEfJER3Y+V1BUSrkSqdRGR2krDu+QMEQGCisGAQQBgjcCAQwxNjA0
# oBSAEgBNAGkAYwByAG8AcwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQu
# Y29tIDANBgkqhkiG9w0BAQEFAASCAQAm5/nRpZdMYHvcMQXXslFgE2fAlxDaJAFG
# df54z/VmVlql7CixsIwPBuGhcjIK9Vmgu3Yc/jiqogdmeoX+5Xr6IpZJEHuDff30
# 66eZryp/7I9/JvHfUix1Gf8igA34xjKzV1VlhT6xWtyQkdt37xZcC1AX/cXuMG86
# 5H95IEi6yP2J1H8NTCrrfWippc+135SFfKH9gVxUabCl9LZfw3ffm0ILhOjbX2q9
# NmeU0jW+yONfV2HYSNOAx1yyuKqalXPdcoSgoXAqFP1lHAvBFartUPjJWDOuN2mM
# CeyVyrVYWDOqWwK1hodZTlJNkDVzejkN57uRzXzCsEnX0hZrWacPoYIXlDCCF5AG
# CisGAQQBgjcDAwExgheAMIIXfAYJKoZIhvcNAQcCoIIXbTCCF2kCAQMxDzANBglg
# hkgBZQMEAgEFADCCAVIGCyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEE
# AYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIBF9K536ay3Vnzqjy7PGMCXomcXElpLw
# fdj864az5ya3AgZoJdN0wGAYEzIwMjUwNTIyMTUwMDM5LjI5NVowBIACAfSggdGk
# gc4wgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNV
# BAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGll
# bGQgVFNTIEVTTjpBNDAwLTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgU2VydmljZaCCEeowggcgMIIFCKADAgECAhMzAAACAnlQdCEUfbih
# AAEAAAICMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAy
# MDEwMB4XDTI1MDEzMDE5NDI0NFoXDTI2MDQyMjE5NDI0NFowgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpBNDAw
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALd5Knpy5xQY6Rw+Di8p
# Yol8RB6yErZkGxhTW0Na9C7ov2Wn52eqtqMh014fUc3ejPeKIagla43YdU1mRw63
# fxpYZ5szSBRQ60+O4uG47l3rtilCwcEkBaFy978xV2hA+PWeOICNKI6svzEVqsUs
# jjpEfw14OEA9dwmlafsAjMLIiNk5onYNYD7pDA3PCqMGAil/WFYXCoe88R53LSei
# 1du1Z9P28JIv2x0Mror8cf0expjnAuZRQHtJ+4sajU5YSbownIbaOLGqL03JGjKl
# 0Xx1HKNbEpGXYnHC9t62UNOKjrpeWJM5ySrZGAz5mhxkRvoSg5213RcqHcvPHb0C
# EfGWT7p4jBq+Udi44tkMqh085U3qPUgn1uuiVjqZluhDnU6p7mcQzmH9YlfbwYtm
# KgSQk3yo57k/k/ZjH0eg6ou6BfTSoLPGrgEObzEfzkcrG8oI7kqKSilpEYa1CVeM
# PK6wxaWsdzJK3noOEvh1xWeft0W8vnTO9CUVkyFWh6FZJCSRa5SUIKog6tN7tFua
# dt0miwf7uUL6fneCcrLg6hnO5R6rMKdIHUk1c8qcmiM/cN7nHCymLm1S9AU1+V8Z
# OyNmBACAMF2D8M7RMaAtEMq9lAJnmoi5elBHKDfvJznV73nPxTabKxTRedKlZ6KA
# eqTI4C0N9wimrka/sdX51rZHAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQU2ga5tQ+M
# /Z/yJ+Qgq/DLWuVIdNkwHwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIw
# XwYDVR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9w
# cy9jcmwvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3Js
# MGwGCCsGAQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3Nv
# ZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENB
# JTIwMjAxMCgxKS5jcnQwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEF
# BQcDCDAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAIPzdoVBTE3f
# seQ6gkMzWZocVlVQZypNBw+c4PpShhEyYMq/QZpseUTzYBiAs+5WW6Sfse0k8XbP
# SOdOAB9EyfbokUs8bs79dsorbmGsE8nfSUG7CMBNW3nxQDUFajuWyafKu6v/qHwA
# XOtfKte2W/NBippFhj2TRQVjkYz6f1hoQQrYPbrx75r4cOZZ761gvYf707hDUxAt
# qD5yI3AuSP/5CXGleJai70q8A/S0iT58fwXfDDlU5OL1pn36o+OzPDfUfid22K8F
# lofmzlugmYfYlu0y5/bLuFJ0l0TRRbYHQURk8siZ6aUqGyUk1WoQ7tE+CXtzzVC5
# VI7nx9+mZvC1LGFisRLdWw+CVef04MXsOqY8wb8bKwHij9CSk1Sr7BLts5FM3Ooc
# y0f6it3ZhKZr7VvJYGv+LMgqCA4J0TNpkN/KbXYYzprhL4jLoBQinv8oikCZ9Z9e
# twwrtXsQHPGh7OQtEQRYjhe0/CkQGe05rWgMfdn/51HGzEvS+DJruM1+s7uiLNMC
# Wf/ZkFgH2KhR6huPkAYvjmbaZwpKTscTnNRF5WQgulgoFDn5f/yMU7X+lnKrNB4j
# X+gn9EuiJzVKJ4td8RP0RZkgGNkxnzjqYNunXKcr1Rs2IKNLCZMXnT1if0zjtVCz
# Gy/WiVC7nWtVUeRI2b6tOsvArW2+G/SZMIIHcTCCBVmgAwIBAgITMwAAABXF52ue
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
# cmljYSBPcGVyYXRpb25zMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTQwMC0w
# NUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wi
# IwoBATAHBgUrDgMCGgMVAEmJSGkJYD/df+NnIjLTJ7pEnAvOoIGDMIGApH4wfDEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
# cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDr2Yvm
# MCIYDzIwMjUwNTIyMTE0MTI2WhgPMjAyNTA1MjMxMTQxMjZaMHQwOgYKKwYBBAGE
# WQoEATEsMCowCgIFAOvZi+YCAQAwBwIBAAICEnUwBwIBAAICEgUwCgIFAOva3WYC
# AQAwNgYKKwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEK
# MAgCAQACAwGGoDANBgkqhkiG9w0BAQsFAAOCAQEAPiQDeAj9RyjBe1/umOAIlmNV
# wVjr1cvQwYRW52nbQT7iYslPtC+V4RLLx3B0HYe1mc0VB/tYFI14UWYIjX3RiilY
# 3ve6bzzEhLozq+kplR3N3wrER/BbTuKOM9PR4ssCD7EgmWd0NTG9xoHA3Pp18ErK
# 3qvZCvdy1Z6lMKqH/80XRUJef3lNoyuQytGqcAFXAS4fTprkHFaDLqzz/TVFlLBh
# 2Ko47mYb9/brtZYa5Zw0puI7j/4XiYJQiFJaOTBi7wTPnGt5FnKdTR/VbtECk9wF
# FXQdqThQ/uBQSEISsOAmoB/NPcHc+mAgS4o5Rmu1SGraPe7aDSMyuc3qnetqUTGC
# BA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAC
# AnlQdCEUfbihAAEAAAICMA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMx
# DQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIHGe5kwKo2zsPQf1If3aqedQ
# 6Jlo+5FLGnoRB0JjXuYQMIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQg843q
# ARgHlsvNcta5SYvxl3zFcCypeSx50XKiV8yUX+wwgZgwgYCkfjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAgJ5UHQhFH24oQABAAACAjAiBCBD9YPU
# 8GhIOek0O5630cyOjyXjzvtI7QNZyd569tN+OzANBgkqhkiG9w0BAQsFAASCAgBn
# TL/7mcB1lqjjfE1DogcmuN1mlGF2cnsGWfaiE1XOOUELpsfDDmmebrv5+EgqFy5j
# NKJPMh8YbchXb5nRbC4IT5DPP3hDtDh6vk5xsubI9eNzvDnmksFAuZviaDsmMyLr
# Y7PSa9BdNVJgvoOl1Hx4OKdS1cMz9tw0DgxciJBiiYHiTozijfTjKVzlQAZKHIJr
# BONnLOaDXqDGvp4kVN3x+7A26pAv38zsoLXNTJth2CFySJhgL2VtlXwxgBzHN8Az
# WFnlihHA97d7PdPuWDQx/HKR96SUqIZZS9o3eZMnkXHS9r4axqETWh6y+7sh0amZ
# tddlRYRT4rFfFg+PgusPZIwU+2B0KSWH6lgrach1hbgJSJHx+Z3HGF7CrJLm+oSN
# ccwfemvhZqvmW6PCYMaIFEs1n6R/CnbVykUYD301C2CDBgoxU0ARUw3O72UnWzms
# HQY7/k/LEv8UokNo2/OAKFPStgrrUnje/7lT2eta5yTZW3RwbYK9e/w1zcmFB//7
# KoUq3SRgtyjmwBv/bJiy32MJN7mwrRA6PxrCHp+mZ/cPPdcyrpD/jDThbr9MyApM
# JhVHHMfGaEgvbkSkJwB9AEhNYNoZIqfvzNMZv6t1bATFkop8E/+RES0xKuG5nb6r
# a2hmroDvGAoaiXh2F/uPuMDS5PLx+QKazcjzrdg6Ew==
# SIG # End signature block
