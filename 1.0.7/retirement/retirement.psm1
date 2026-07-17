using module ../utils/utils.psm1

<#
.SYNOPSIS
    Retrieves active retirement health advisory events.

.DESCRIPTION
    This module contains functions related to the capturing and collecting to active retirement health advisory events.
    It includes the following functions:
    - Get-WAFResourceRetirement
    - New-WAFResourceRetirementObject

.EXAMPLE
    PS> $retirementObjects = Get-WAFResourceRetirement -SubscriptionIds '11111111-1111-1111-1111-111111111111'

.NOTES
    Author: Takeshi Katano
    Date: 2024-10-02
#>

<#
.SYNOPSIS
    Retrieves active retirement health advisory events based on the specified subscription ID.

.DESCRIPTION
    The Get-WAFResourceRetirement function takes a subscription ID and retrieves active retirement health advisory events.

.PARAMETER SubscriptionIds
    A subscription ID to retrieves active retirement health advisory events.

.OUTPUTS
    Returns a list of retirement events, including the name and properties of each event.

.EXAMPLE
    PS> $retirementObjects = Get-WAFResourceRetirement -SubscriptionIds '11111111-1111-1111-1111-111111111111'

    This example retrieves the recent retirement events for the specified Azure subscription.

.NOTES
    Author: Takeshi Katano
    Date: 2024-10-02
#>
function Get-WAFResourceRetirement {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-WAFIsGuid -StringGuid $_ })]
        [string[]] $SubscriptionIds
    )

    $retirementObjects = @()

    foreach ($subscriptionId in $SubscriptionIds) {
        # NOTE:
        # ARG query with ServiceHealthResources returns last 3 months of events.
        # Azure portal shows last 1 months of events.
        $cmdletParams = @{
            Method               = 'GET'
            SubscriptionId       = $subscriptionId
            ResourceProviderName = 'Microsoft.ResourceHealth'
            ResourceType         = 'events'
            ApiVersion           = '2024-02-01'
            QueryString          = @(
            ('queryStartTime={0}' -f (Get-Date).AddMonths(-3).ToString('yyyy-MM-ddT00:00:00')),
                '$filter=(properties/eventType eq ''HealthAdvisory'') and (properties/eventSubType eq ''Retirement'') and (Properties/Status eq ''Active'')'
            ) -join '&'
        }
        $response = Invoke-AzureRestApi @cmdletParams
        $retirementEvents = ($response.Content | ConvertFrom-Json).value

        $return = foreach ($retirementEvent in $retirementEvents) {
            $cmdletParams = @{
                SubscriptionId  = $subscriptionId
                TrackingId      = $retirementEvent.name
                Status          = $retirementEvent.properties.status
                LastUpdateTime  = $retirementEvent.properties.lastUpdateTime
                StartTime       = $retirementEvent.properties.impactStartTime
                EndTime         = $retirementEvent.properties.impactMitigationTime
                Level           = $retirementEvent.properties.level
                Title           = $retirementEvent.properties.title
                Summary         = $retirementEvent.properties.summary
                Header          = $retirementEvent.properties.header
                ImpactedService = $retirementEvent.properties.impact.impactedService
                Description     = $retirementEvent.properties.description
            }
            New-WAFResourceRetirementObject @cmdletParams
        }
        $retirementObjects += $return
    }
    return $retirementObjects
}

<#
.SYNOPSIS
    Creates a retirement object.

.DESCRIPTION
    The New-WAFResourceRetirementObject function creates a retirement object based on the specified parameters.

.PARAMETER SubscriptionId
    The subscription ID of the retirement event.

.PARAMETER TrackingId
    The tracking ID of the retirement event. It's usually as the XXXX-XXX format.

.PARAMETER Status
    The status of the retirement event. It's usually Active or Resolved.

.PARAMETER LastUpdateTime
    The last update time of the retirement event.

.PARAMETER StartTime
    The impact start time of the retirement event.

.PARAMETER EndTime
    The impact mitigation time of the retirement event.

.PARAMETER Level
    The level of the retirement event such as Warning, etc.

.PARAMETER Title
    The title of the retirement event.

.PARAMETER Summary
    The summary of the retirement event.

.PARAMETER Header
    The header of the retirement event.

.PARAMETER ImpactedService
    The impacted services of the retirement event.

.PARAMETER Description
    The description of the retirement event.

.OUTPUTS
    Returns a ResourceRetirementObject as a PSCustomObject.

.EXAMPLE
    PS> $retirementObject = New-WAFResourceRetirementObject -SubscriptionId $subscriptionId -TrackingId 'XXXX-XXX' -Status 'Active' -LastUpdateTime $lastUpdateTime -StartTime $startTime -EndTime $endTime -Level 'Warning' -Title $title -Summary $summary -Header $header -ImpactedService $impactedServices -Description $description

.NOTES
    Author: Takeshi Katano
    Date: 2024-10-02
#>
function New-WAFResourceRetirementObject {
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
# MIIoRQYJKoZIhvcNAQcCoIIoNjCCKDICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAGurIIf+3UfLOF
# VEyFyMRNBZlfv8vUVrDD4CDO8mhWBaCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGiUwghohAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIMP180D6elmmfIbfoE35Fk+y
# jsYPK/ESCPKMvCk0lZ6tMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQArNIviUro64sunJ1UXk682IyLDDRJ0w5GwZ/8DDsgf2B3AtbqYgYes
# e5t7c+R9jfDlgyO/7ELIDFFbgknv3Ivm8z6J39qa86Mro4WC2X25Q2YxZtInOrhS
# vZqK45bo5QOdv2iOQFpWELvmLP9bpYOj3tl2+/IhOOZRAhLaVOwPfnQ51gxw9PLq
# 5YeZUJwiDadoRUrEMhDghnNH0vxgsbx4GDptT880ZnK1mIXQcRZo3BgZ7IF3Qq0c
# n2agEJQveQ1zZgW1faGGM22t5n4abrI4kpq0+F8EpI1Ab8Vc4fCuAG/akkqx/Ial
# sFl57kdnE9TrUClYnNg/8E8WWhOxc3LtoYIXrTCCF6kGCisGAQQBgjcDAwExgheZ
# MIIXlQYJKoZIhvcNAQcCoIIXhjCCF4ICAQMxDzANBglghkgBZQMEAgEFADCCAVoG
# CyqGSIb3DQEJEAEEoIIBSQSCAUUwggFBAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEILc1qVJiMJTMcl1j6zUOpxAetZuogzQ01Z9NG2Oc+peYAgZoLlXz
# zwsYEzIwMjUwNTIyMTQ1OTQ4LjUzOFowBIACAfSggdmkgdYwgdMxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1Mg
# RVNOOjY1MUEtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloIIR+zCCBygwggUQoAMCAQICEzMAAAH1mQmUvPHGUIwAAQAAAfUw
# DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcN
# MjQwNzI1MTgzMTAxWhcNMjUxMDIyMTgzMTAxWjCB0zELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQg
# T3BlcmF0aW9ucyBMaW1pdGVkMScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046NjUx
# QS0wNUUwLUQ5NDcxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZp
# Y2UwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDM73RwVBNZ39Y/zghP
# skwhbV9AvrWx1+CaGV9PSe9gRvaS+Q0XTvdnCO965Jai3fzsuMTRMKIb3d7ojQfM
# gVAGdvEY/9Y8FSKsWrtYTlECy6E19hQv48hv2MmrcLBbEgJ/Dm3+lPIg4eMq+jWV
# A5NZnKmKv+mxnAQTvLa5YA9tklMWsp6flHvfHYdvHLh5bUNyZePKbbAVa/XSwEfj
# RqMl9746TBxN2hitjcqSk39FBKN7JwrRuGOjQIZghhr5kwBqjRI1H9HUnVjqwuSI
# Ik7dpCttLVLPuX7+omDLx/IRkw0PkyzsLSwRo6+gEeJZKlMzi9zTEMsKZzo8a/Tc
# K1a7YqLKqsvwEAHURjI5KEjchPv1qsMgOsv5173UV+OZJsFjmP4e9LSXd1eSM/ce
# ifxvviVbCKQXSvMSsXSfeSFUC6zHtKbWgYb7TqHP1cDLdai75PpJ7qhrksOJCA9N
# 9ZH+P0U2Twm7TqhJ9OHpzTdXS6WVrQjDL4fNSX5aZjEUtTQ+JpeyaC503BWqfnXO
# v4GLdc8nznBa7LoYZPucEOZc3NM2TMr3wMFCNM5ptBdRnzzhhv0MU1yKCZ5VNiTJ
# RdnqGxx3w3KrjkDcPduT6deeyiArVnvNmPpdsZ+3vGA5i+95TqnT5+u2FsXsxe/6
# LmpwP0d8WY6rhVgd69V6xhvo/QIDAQABo4IBSTCCAUUwHQYDVR0OBBYEFFXgfFv1
# SjSgcPAlkl7baLF7YHUBMB8GA1UdIwQYMBaAFJ+nFV0AXmJdg/Tl0mWnG1M1Gely
# MF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lv
# cHMvY3JsL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNy
# bDBsBggrBgEFBQcBAQRgMF4wXAYIKwYBBQUHMAKGUGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2lvcHMvY2VydHMvTWljcm9zb2Z0JTIwVGltZS1TdGFtcCUyMFBD
# QSUyMDIwMTAoMSkuY3J0MAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwgwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBzDGZU9oD3
# Ed9+6ibRM5KnaGym5/UbRwdb8pccC6Gelbz9K+WrmP1ooj/z8bp8YhAyvTOWlq7y
# PLzcjjNyUZ0mOXlLTZLEOVQprC1a+B/uJ1rTo+CN5AzV5fgu63hts99PQUSnvsbv
# qGHKxfFMk0e/nL5/BOFR6KJyWKFnCpxkylrjqb6hXqKBNTojQSed6i0yoWzRDfMe
# BVWvhZeOcbYFyeSKnjZ53KD/2JdzOpMGsSS9PPRSWW2kUZpCcvOr42jxUSCrRQbt
# bUQtkaGabEWYcAHBNPqw8kVXrwN8ugLSIH1Btv1Vnya9tNXkm0hIGSVO5UCUSTNe
# L0siM2bH6Sd0F8o/x3Eb/FtFem1ANANoKxLqiAuTAuAfrNKz66X1abMjQXzMiZuG
# dmFTOIgeF4Wjgf5miiM9hyBMrr/duRJs5puZAV/3kHwHp7lapdtLmz050x1SVbWB
# MjWvm75YDAfYobt3Gd6hNt/+NiXdNS0/sAenJyTZzSe6f9DQLJylr1BQf8PLTWTq
# 1CiY1caOK+Db8EZyBknQfDwLopV6UQnfXEugTbWb340SBIoJGgUTUuSZrfVLIhrK
# dt1gRvyw6VnKcx2bzI+V0PC4Xni8mIQCuOtwM1d7oGhtlSJNZIq+/UMlp1HVJQI7
# 853bUaBT6Fmq750qCMoBh15Mi+L1Hau0tjCCB3EwggVZoAMCAQICEzMAAAAVxedr
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
# cQZqELQdVTNYs6FwZvKhggNWMIICPgIBATCCAQGhgdmkgdYwgdMxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJ
# cmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEnMCUGA1UECxMeblNoaWVsZCBUU1Mg
# RVNOOjY1MUEtMDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloiMKAQEwBwYFKw4DAhoDFQAmwAq7jw1tHlhGDdZIFALKPN2S9qCB
# gzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEB
# CwUAAgUA69l9KDAiGA8yMDI1MDUyMjEwMzgzMloYDzIwMjUwNTIzMTAzODMyWjB0
# MDoGCisGAQQBhFkKBAExLDAqMAoCBQDr2X0oAgEAMAcCAQACAiufMAcCAQACAhLN
# MAoCBQDr2s6oAgEAMDYGCisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAI
# AgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQELBQADggEBAJtZFAig35Ju
# Uqkwt+pHTl81URS3UuqwnNxrXhYULPAZZ2IQzoNL4Mq/IK9A1KLtMMwseXw2H0Xb
# 3tins9LSu94egLHjx006BRkkl6rw7dURQgzcJ7xpJoQgiPK2KsnNIKJC+xiryI9c
# VmXEqVLBMMAOHmCW8ZoJEu6qPymix/+z2MPUeClCSSwNMQ2fxdduAft6NcI5E5/s
# h1OYT3BXfK0AYnuvJ1L4o783vBvh+RYbFf7ZUW/HzDiEIpNzVQHXZ45KKF7fRGJk
# UvVMKZ3ryPfDKx3Glw00EFRnnNq28kFyt1YV+0QIjP8h5N31KmFJGZN9kNj5rxFh
# Q5T7gi5gPmkxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMAITMwAAAfWZCZS88cZQjAABAAAB9TANBglghkgBZQMEAgEFAKCCAUowGgYJ
# KoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCAH/VMlmt5u
# mHDvAFr+iBievNukGuC7FljlgJpyQ+yg4zCB+gYLKoZIhvcNAQkQAi8xgeowgecw
# geQwgb0EIMHW8tIXCHT0hK7iR0S/j+2D15HViTzDnHuPkZOGpo81MIGYMIGApH4w
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAH1mQmUvPHGUIwAAQAA
# AfUwIgQgq0GKwV7spW5cgNo4QygF0craMcqWeaF2lG+KVtpmfx4wDQYJKoZIhvcN
# AQELBQAEggIAeu4AuqcY8ckvPM0BCPOPcZprDjRdwqBysnrMMBkM4RRin+pbIrhv
# O2io143jW4OB/LTFPMn4FFJKlzy9nMSeSLdRdJ4Fl2izsUv5WPD5kP79hvqXh42r
# 8hZRd/3SdFwR2sBN4muQAV1FRitvs8r9T7+zEldIwRAiIMZQBY9pcD3syGy6+A6z
# kEivIKZxbGIQHiWGRmI1E+PImTVb/8vuuxpCzlS34+lHGPUnmvP61ALXZvCnQJUo
# lsQ2AarL42h7joyiHR6SeT5IBXYr/7wDX8jcbQtzQb1S9955uoux7UtlANjmvLi9
# pnQPGaLYpxQKQh/vHplTtSmA6cEGtm24Z5+InjuawfZYphvG/bS+gxlKH2w6WlN0
# IXiq1opiVwVfNMYdrOa37zuTSLPedhRGLEPqcRTcjvD6FwZFIFQ03Dd+aDH8Iw7M
# /yK0vESgWDP2O9nfHsnBGLI+xpDm7XRufLI5IgWYr6t2U3yV0w2T02Bvi3g76CbO
# No5Z7iB+7F7IXZT7Obq1GZBNKf577ZWHEJFioKp4zcvFoyr2fkJvu1IUPsHRz0/p
# J6LnkC1PXQHyzhsiSzorZ2qUnRUZvTW9qOMK+klx0fpyC4rcqSbEmNhJcf9hieM4
# zT2wgkxGjG+U9BVFH3dEHaCSK1AaIVN6n1fbYE6NPRC3G7vtzO0bFj4=
# SIG # End signature block
