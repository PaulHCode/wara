using module ../utils/utils.psd1

<#
.SYNOPSIS
    Retrieves recent outage service issue events.

.DESCRIPTION
    This module contains functions related to the capturing and collecting to recent outage service issue events.
    It includes the following functions:
    - Get-WAFOutage
    - New-WAFOutageObject

.EXAMPLE
    PS> $outageObjects = Get-WAFOutage -SubscriptionIds '11111111-1111-1111-1111-111111111111'

.NOTES
    Author: Takeshi Katano
    Date: 2024-10-23
#>

<#
.SYNOPSIS
    Retrieves recent outage events for a given Azure subscription.

.DESCRIPTION
    The Get-WAFOldOutage function queries the Microsoft Resource Health API to retrieve recent outage events for a specified Azure subscription. It filters the events to include only those that have updated in the last three months.
    This function is used for backwards compatibility with current versions of the Analyzer script.

.PARAMETER SubscriptionIds
    The subscription ID for the Azure subscription to retrieve outage events.

.OUTPUTS
    Returns a list of outage events, including the name and properties of each event.

.EXAMPLE
    PS> $outageObjects = Get-WAFOldOutage -SubscriptionIds '11111111-1111-1111-1111-111111111111'

    This example retrieves the recent outage events for the specified Azure subscription.

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-16
#>
function Get-WAFOldOutage {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-WAFIsGuid -StringGuid $_ })]
        [string[]] $SubscriptionIds
    )

   $serviceIssueEvents = foreach ($subscriptionId in $SubscriptionIds) {
        # NOTE:
        # ARG query with ServiceHealthResources returns last 3 months of events.
        # Azure portal shows last 3 months of events maximum.
        $cmdletParams = @{
            Method               = 'GET'
            SubscriptionId       = $subscriptionId
            ResourceProviderName = 'Microsoft.ResourceHealth'
            ResourceType         = 'events'
            ApiVersion           = '2024-02-01'
            QueryString          = @(
                ('queryStartTime={0}' -f (Get-Date).AddMonths(-3).ToString('yyyy-MM-ddT00:00:00')),
                '$filter=(properties/eventType eq ''ServiceIssue'')'
            ) -join '&'
        }
        $response = Invoke-AzureRestApi @cmdletParams
        ($response.Content | ConvertFrom-Json).value | Select-Object name, properties
    }

    return $serviceIssueEvents
}

<#
.SYNOPSIS
    Retrieves recent outage service issue events.

.DESCRIPTION
    This module contains functions related to the capturing and collecting to recent outage service issue events.
    It includes the following functions:
    - Get-WAFOutage
    - New-WAFOutageObject

.EXAMPLE
    PS> $outageObjects = Get-WAFOutage -SubscriptionIds '11111111-1111-1111-1111-111111111111'

.NOTES
    Author: Takeshi Katano
    Date: 2024-10-23
#>

<#
.SYNOPSIS
    Retrieves recent outage events for a given Azure subscription.

.DESCRIPTION
    The Get-WAFOutage function queries the Microsoft Resource Health API to retrieve recent outage events for a specified Azure subscription. It filters the events to include only those that have updated in the last three months.

.PARAMETER SubscriptionIds
    The subscription ID for the Azure subscription to retrieve outage events.

.OUTPUTS
    Returns a list of outage events, including the name and properties of each event.

.EXAMPLE
    PS> $outageObjects = Get-WAFOutage -SubscriptionIds '11111111-1111-1111-1111-111111111111'

    This example retrieves the recent outage events for the specified Azure subscription.

.NOTES
    Author: Takeshi Katano
    Date: 2024-10-23
#>
function Get-WAFOutage {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-WAFIsGuid -StringGuid $_ })]
        [string[]] $SubscriptionIds
    )

    $outageObjects = @()

    foreach ($subscriptionId in $SubscriptionIds) {
        # NOTE:
        # ARG query with ServiceHealthResources returns last 3 months of events.
        # Azure portal shows last 3 months of events maximum.
        $cmdletParams = @{
            Method               = 'GET'
            SubscriptionId       = $subscriptionId
            ResourceProviderName = 'Microsoft.ResourceHealth'
            ResourceType         = 'events'
            ApiVersion           = '2024-02-01'
            QueryString          = @(
                ('queryStartTime={0}' -f (Get-Date).AddMonths(-3).ToString('yyyy-MM-ddT00:00:00')),
                '$filter=(properties/eventType eq ''ServiceIssue'')'
            ) -join '&'
        }
        $response = Invoke-AzureRestApi @cmdletParams
        $serviceIssueEvents = ($response.Content | ConvertFrom-Json).value

        $return = foreach ($serviceIssueEvent in $serviceIssueEvents) {
            $cmdletParams = @{
                SubscriptionId  = $subscriptionId
                TrackingId      = $serviceIssueEvent.name
                Status          = $serviceIssueEvent.properties.status
                LastUpdateTime  = $serviceIssueEvent.properties.lastUpdateTime
                StartTime       = $serviceIssueEvent.properties.impactStartTime
                EndTime         = $serviceIssueEvent.properties.impactMitigationTime
                Level           = $serviceIssueEvent.properties.level
                Title           = $serviceIssueEvent.properties.title
                Summary         = $serviceIssueEvent.properties.summary
                Header          = $serviceIssueEvent.properties.header
                ImpactedService = $serviceIssueEvent.properties.impact.impactedService
                Description     = $serviceIssueEvent.properties.description
            }
            New-WAFOutageObject @cmdletParams
        }
        $outageObjects += $return
    }

    return $outageObjects
}

<#
.SYNOPSIS
    Creates an outage object.

.DESCRIPTION
    The New-WAFOutageObject function creates an outage object based on the specified parameters.

.PARAMETER SubscriptionId
    The subscription ID of the outage event.

.PARAMETER TrackingId
    The tracking ID of the outage event. It's usually as the XXXX-XXX format.

.PARAMETER Status
    The status of the outage event. It's usually Active or Resolved.

.PARAMETER LastUpdateTime
    The last update time of the outage event.

.PARAMETER StartTime
    The impact start time of the outage event.

.PARAMETER EndTime
    The impact mitigation time of the outage event.

.PARAMETER Level
    The level of the outage event such as Warning, etc.

.PARAMETER Title
    The title of the outage event.

.PARAMETER Summary
    The summary of the outage event.

.PARAMETER Header
    The header of the outage event.

.PARAMETER ImpactedService
    The impacted services of the outage event.

.PARAMETER Description
    The description of the outage event.

.OUTPUTS
    Returns an OutageObject as a PSCustomObject.

.EXAMPLE
    PS> $outageObject = New-WAFOutageObject -SubscriptionId $subscriptionId -TrackingId 'XXXX-XXX' -Status 'Active' -LastUpdateTime $lastUpdateTime -StartTime $startTime -EndTime $endTime -Level 'Warning' -Title $title -Summary $summary -Header $header -ImpactedService $impactedServices -Description $description

.NOTES
    Author: Takeshi Katano
    Date: 2024-10-23
#>
function New-WAFOutageObject {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-WAFIsGuid -StringGuid $_ })]
        [string] $SubscriptionId,

        [Parameter(Mandatory = $true)]
        [string] $TrackingId,

        [Parameter(Mandatory = $true)]
        [string] $Status,

        [Parameter(Mandatory = $true)]
        [datetime] $LastUpdateTime,

        [Parameter(Mandatory = $true)]
        [datetime] $StartTime,

        [Parameter(Mandatory = $true)]
        [datetime] $EndTime,

        [Parameter(Mandatory = $true)]
        [string] $Level,

        [Parameter(Mandatory = $true)]
        [string] $Title,

        [Parameter(Mandatory = $true)]
        [string] $Summary,

        [Parameter(Mandatory = $true)]
        [string] $Header,

        [Parameter(Mandatory = $true)]
        [string[]] $ImpactedService,

        [Parameter(Mandatory = $true)]
        [string] $Description
    )

    return [PSCustomObject] @{
        Subscription    = $SubscriptionId
        TrackingId      = $TrackingId
        Status          = $Status
        LastUpdateTime  = $LastUpdateTime.ToString('yyyy-MM-dd HH:mm:ss')
        StartTime       = $StartTime.ToString('yyyy-MM-dd HH:mm:ss')
        EndTime         = $EndTime.ToString('yyyy-MM-dd HH:mm:ss')
        Level           = $Level
        Title           = $Title
        Summary         = $Summary
        Header          = $Header
        ImpactedService = $ImpactedService -join ', '
        Description     = $Description
    }
}

# SIG # Begin signature block
# MIIoLAYJKoZIhvcNAQcCoIIoHTCCKBkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBaPEKQ5GIc6eJH
# QkMM8CbLgt+zDtV0hTGrS13rPiHsp6CCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIJkYzeL87tmV8Lwrr5V7LHfq
# N+4Kr9M2QSJtawzhw5BnMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQCBEg+1IRbDY/pbFRG9UuTtJ4az4TRXjGaEY31l07rGKvyWdkBfElJ4
# biysFp5/CJ31dLyQkMUgqjkSJ7Q0mo1rgndZT1k3qQ+zxAODx0TFfCuEzkJ8IXg8
# eVSyLddFQYgU/KDD9oKaHi1b5AlamtYnQwkXNpLAVIhkHePQdxV/3ttaqVznBlov
# YMXiFP4EqqPGdtsD5rUXDefuAcdUsws7CDCd0lIaiwPvX0urS1L2A9xTVrNPj7MH
# zNllf0lKmq2/zmfhGLfu1VpSeRUW874kbMZBpwVtmUAVCpAYbCZBWB58JnLa9+er
# VC6ISGSJvEwULIgLCXfN/xPwlWdcUXoxoYIXlDCCF5AGCisGAQQBgjcDAwExgheA
# MIIXfAYJKoZIhvcNAQcCoIIXbTCCF2kCAQMxDzANBglghkgBZQMEAgEFADCCAVIG
# CyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIJTa613JqK2q+TXQNMJNc79W1E+7b0m+v+evH6XfZhp/AgZoJfbw
# kGMYEzIwMjUwNTIyMTUwMTA0LjExN1owBIACAfSggdGkgc4wgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpBMDAw
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZaCCEeowggcgMIIFCKADAgECAhMzAAACCHidWF2Sx9lSAAEAAAIIMA0GCSqGSIb3
# DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDEzMDE5
# NDI1M1oXDTI2MDQyMjE5NDI1M1owgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpBMDAwLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBALXLcAjmUjPcinWcrkExRRsZGyNKcLP9UazffO9j
# Q16yGw+B+V3f9nf/d1LxYoEuUOWiyZck5mUVKI1dR3oNMnED2KT9lsJ1YnvwsqNs
# 3e0WRfZzpFGlEykDyyr+gFtGvI/dzxD+DGkkAocfPxy5Kft7B8IvOc2bGqWJOTdD
# kserPY+N5goP91sowFPZMABYC+6bjP8dcgnq0V0ag1XhZRFmzAJK3pE7BqpDWgBg
# a8Sd0f4NdmrX5seyC9w80J1NIilahtCIlL9QouJHTYo0KoHgj3JqMVNKWcwgQzP8
# 2LnfygYjrimFy82lR7b6popYdnx3hPUqCG9GZJIXhgXkM0QlvFoJTCzLudQuawWd
# No6NU6hMVZZ9Ze8G44qQFxApYYq+uSL3vqPjH7l7MA/fp+re7p0dElMtkC7h0S46
# ihTf6Qxmv5EFhaNMdAIpX7JnVJPR4aRsdegDaXLJEOU2MFByh5kjFYJm2z93f6d/
# WOJIs3p/rB0dpTPQAPA5ND9oSjUgLzl4V4+/IgprEUmQZTYyprpfOreoKrm2iQge
# 2OGiCysSB8MpN4VdO12GXUg+0twJ8xxY4YYBeixVRTsb3jwXpb3rjbh/ZUcPwvcW
# jIAj36vjPIhBSaqIRLO5BZ5alNOMVjAKaBdoY65INXxw05VAHog/M+d5mFVOPDFs
# BmVpAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUozHi986pxROBg5UH1/Xz+aF6AU0w
# HwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
# UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAw
# XjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# ZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAFGzMokrAB+zCp1pA8WJpgH9k0BhYNrT
# jRdNxCwJoK5rmewbUiyVhpKkfuaJMuvp5ASpdzNmip45r/G8OcwaJ8Y11rQIdtDC
# 2mciGy62so7aOGMRobCUmA4yqXbWvsiTpecHNrR7eEE67hQGQyX8sRf4BRG3uLv5
# FM2wY3Rxc/A9JtMUT73PtZqAZtj2nBSj83GQYmx6oYJD/0rZUxTvhvDl7v7wgZSE
# zbGykk+qdJ4c+FiZwHRZyU7FxUh+P9m5C/Cis9tMQRgNULNI3ftzTIKE8xjsUn4c
# YUE0nHB3mUoivsZh+rxrSA6ILaWMZiVziu3hwJ53VcqDzd/SX1pRWKZYFhe1815u
# Gl+votzeMPw2CysOHO3RaCch2dNkKLuPGOwgGKUf32ljn+HptBwsor8TooI/0TVg
# 3vx8to5eRczI7rEuu9Bn64JKLWF1O58ULuhIH8JTlFt8hUdcbSPWjafW2d7h4Js1
# 8qpQ9MTfW01tYFHbdiLLSveRCYd5gTUYtsvinCSepqKnUFGfpYhQwm2CdxrAQ3fd
# /wBgZnhrc2ceinMZVXqd598ZVqDhN27L6jLVgX6yEKGhd0yp+E9YWkd7e4kZPgYk
# SI2zj7bxr/AdS4X5pFpHRw3k/teU7BTXfrSJQIm1B28pBo0DAYjb0o7BLdAauJH0
# XaM4Y9QCWl4TMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkq
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
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046QTAwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMV
# AI2S+7Q0pxKN1grKuEllyzJc5RM0oIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDr2a9kMCIYDzIwMjUwNTIyMTQx
# MjUyWhgPMjAyNTA1MjMxNDEyNTJaMHQwOgYKKwYBBAGEWQoEATEsMCowCgIFAOvZ
# r2QCAQAwBwIBAAICAsowBwIBAAICEjQwCgIFAOvbAOQCAQAwNgYKKwYBBAGEWQoE
# AjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDANBgkq
# hkiG9w0BAQsFAAOCAQEAQ0+Pnr1fGPf98ffUX1N3uACz4X/nc2TKR98P9erWLBkJ
# dEONXMSmlVuyocOxTCfpYkEQUtQfNGWKXG3HeMHhSs0Sr/OzFr3q6o4ughQXfGsf
# 60UYLrktbgPB67YrttK+HkwwMpXCeqseOpH6cMJPTj9mCJlSv6UsHumvfhqbuh9Y
# BEDRhIyE0emE3aTW+9VqdZf16GYDLRQ3K7IT4k95nwnVchK2a3HYpVuwDN9KxbdK
# caqyu2ZTFVJd5WHrF/QCV1un7DggdkVwRzPW5cR+9fEQL/VkXnXUPlziHxeVfOyt
# h+jLhlxbykXwUMw4eYfXnbgjnA1M40dSji0+TrwbFTGCBA0wggQJAgEBMIGTMHwx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAACCHidWF2Sx9lSAAEAAAII
# MA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQw
# LwYJKoZIhvcNAQkEMSIEIDPM01PvMA4USxSulOdj3Ej10e7BvK7EIKukBxVy2I0z
# MIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQgj/+ObwkdrZbU73vvy334W3Zn
# k2Yqq20+TpD71FGmZ6kwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMAITMwAAAgh4nVhdksfZUgABAAACCDAiBCD+X1WUnA/YqKxnI4J9RriIERgT
# 53MtOMGhx5FkgjZqpzANBgkqhkiG9w0BAQsFAASCAgA35I6oE5z0LW734EgLQBlk
# L2SvzbqqIC44F+xEmirP2esrxA68zBzA1SZw0zizyF3Iyeqs0A/2frETEhBEQa7M
# czmjiB6pabYdxluFRXi0gSFbpLM6bhZMKzvemZpy8zl2PJ+kSBKOwy3qtx3BqF2H
# c5JTPaArzwFsobXoW8Xgwe+zTzEA9yMjPGeVCxVa0lgDB5oS4Dz3n6gOI1QuYrZG
# X6NHnSkRtJ9Y8XS1Sgiw0H1/GTW/wzg4eBlQQPopbace1JK/ObqF7b/H9O/k/Akh
# YfepEqzhPJwqWdr/Kov74Gn28I8drwY90zVyMVRlsRpKo3qQWqluG1HM+uwFUqyc
# C00jIhJjOJA+wYLjk9Qhegbx4AK9H9ElXFVIok0fV48ELXHuEoeltbcdRLu0/cFa
# vjYAWU/khlI1BY6Cdyw3/yiV6+PalL/KZ7z3tgTFZSUyBACxhTSBsFwkvX3Kvh9t
# bZbm6mXCOSb72fk3MVHjdHbIR6w+DRMO00aVYmL8zMknez8LpXThvkkcg/v2PVOq
# gsgzGBvr3xP5cgJV/HVt29rGf5vyY3aKCfHdG+M87+3rrFdhG/8JsD96MeDJ8875
# 84Rlnt4kkqmIYNio8MddvGrbSTSx+TsZDhjQyCMXfSnkVlkJ94BeIOQsPx85EyUR
# F+uaOfNHJ2orcMVg6AMeYw==
# SIG # End signature block
