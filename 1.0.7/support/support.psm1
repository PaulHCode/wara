using module ../utils/utils.psd1

<#
.SYNOPSIS
    Retrieves recent support tickets.

.DESCRIPTION
    This module contains functions related to the capturing and collecting to recent service tickets.
    It includes the following functions:
    - Get-WAFSupportTicket
    - New-WAFSupportTicketObject

.EXAMPLE
    PS> $serviceTiketObjects = Get-WAFSupportTicket -SubscriptionIds '11111111-1111-1111-1111-111111111111'

.NOTES
    Author: Takeshi Katano
    Date: 2024-11-08
#>

<#
.SYNOPSIS
    Retrieves recent service tickets for a given Azure subscription.

.DESCRIPTION
    The Get-WAFSupportTicket function queries the Azure Resource Graph to retrieve recent service tickets for a specified Azure subscription. It filters the service tickets to include only those that have created in the last three months.

.PARAMETER SubscriptionIds
    The subscription ID for the Azure subscription to retrieve service tickets.

.OUTPUTS
    Returns a list of service tickets, including the name and properties of each tickets.

.EXAMPLE
    PS> $serviceTiketObjects = Get-WAFSupportTicket -SubscriptionIds '11111111-1111-1111-1111-111111111111'

    This example retrieves the recent service tickets for the specified Azure subscription.

.NOTES
    Author: Takeshi Katano
    Date: 2024-11-08
#>
function Get-WAFSupportTicket {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-WAFIsGuid -StringGuid $_ })]
        [string[]] $SubscriptionIds
    )

    $argQuery = @'
SupportResources
| where type =~ "Microsoft.Support/supportTickets"
| where properties.Severity !~ "Minimal"  // Exclude severity C
| extend
    createdDateAsDateTime  = todatetime(properties.CreatedDate),   // UTC
    modifiedDateAsDateTime = todatetime(properties.ModifiedDate)   // UTC
| where createdDateAsDateTime >= datetime_add("Month", -3, now())  // Last 3 months (not 90 days)
| project
    supportTicketId                  = properties.SupportTicketId,
    severity                         = properties.Severity,
    status                           = properties.Status,
    supportPlanType                  = properties.SupportPlanType,
    createdDate                      = createdDateAsDateTime,
    modifiedDate                     = modifiedDateAsDateTime,
    title                            = properties.Title,
    technicalTicketDetailsResourceId = iif(isnull(properties.TechnicalTicketDetails.ResourceId), "", properties.TechnicalTicketDetails.ResourceId)
| order by createdDate desc
'@

    $supportTickets = Invoke-WAFQuery -SubscriptionIds $SubscriptionIds -Query $argQuery

    $supportTicketObjects = foreach ($supportTicket in $supportTickets) {
        $cmdletParams = @{
            SupportTicketId                  = $supportTicket.supportTicketId
            Severity                         = $supportTicket.severity
            Status                           = $supportTicket.status
            SupportPlanType                  = $supportTicket.supportPlanType
            CreatedDate                      = $supportTicket.createdDate
            ModifiedDate                     = $supportTicket.modifiedDate
            Title                            = $supportTicket.title
            TechnicalTicketDetailsResourceId = $supportTicket.technicalTicketDetailsResourceId
        }
        New-WAFSupportTicketObject @cmdletParams
    }

    return $supportTicketObjects
}

<#
.SYNOPSIS
    Creates a service ticket object.

.DESCRIPTION
    The New-WAFSupportTicketObject function creates a service ticket based on the specified parameters.

.PARAMETER SupportTicketId
    The ID of the support ticket. It's usually sixteen digits of number.

.PARAMETER Severity
    The severity of the support ticket such as Minimal, Moderate, etc.

.PARAMETER Status
    The status of the support ticket. It's usually Open or Closed.

.PARAMETER SupportPlanType
    The support plan type of the support ticket such as Unified Enterprise, etc.

.PARAMETER CreatedDate
    The created date of the support ticket.

.PARAMETER ModifiedDate
    The modified date of the support ticket.

.PARAMETER Title
    The title of the support ticket.

.PARAMETER TechnicalTicketDetailsResourceId
    The resource ID of the related Azure resource to the support ticket if it's available.

.OUTPUTS
    Returns a SupportTicketObject as a PSCustomObject.

.EXAMPLE
    PS> $serviceTiketObject = New-WAFSupportTicketObject -SupportTicketId '0123456789012345' -Severity 'Moderate' -Status 'Open' -SupportPlanType 'Unified Enterprise' -CreatedDate $createdDate -ModifiedDate $modifiedDate -Title $title -TechnicalTicketDetailsResourceId $resourceId

.NOTES
    Author: Takeshi Katano
    Date: 2024-11-08
#>
function New-WAFSupportTicketObject {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $true)]
        [string] $SupportTicketId,

        [Parameter(Mandatory = $true)]
        [string] $Severity,

        [Parameter(Mandatory = $true)]
        [string] $Status,

        [Parameter(Mandatory = $true)]
        [string] $SupportPlanType,

        [Parameter(Mandatory = $true)]
        [datetime] $CreatedDate,

        [Parameter(Mandatory = $true)]
        [datetime] $ModifiedDate,

        [Parameter(Mandatory = $true)]
        [string] $Title,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string] $TechnicalTicketDetailsResourceId
    )

    return [PSCustomObject] @{
        'Ticket ID'         = $SupportTicketId
        'Severity'          = $Severity
        'Status'            = $Status
        'Support Plan Type' = $SupportPlanType
        'Creation Date'     = $CreatedDate.ToString('yyyy-MM-dd HH:mm:ss')
        'Modified Date'     = $ModifiedDate.ToString('yyyy-MM-dd HH:mm:ss')
        'Title'             = $Title
        'Related Resource'  = $TechnicalTicketDetailsResourceId
    }
}

# SIG # Begin signature block
# MIIoLAYJKoZIhvcNAQcCoIIoHTCCKBkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD4BdaSFemt3BxU
# EKzFNhTsH0BTi6+UJ4sxdTp3JTqcaaCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIHEYaLVNcRtVoOI3Gz5ium+v
# SFEB1E2S5sie0R5pQ9SWMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQAQQ3xdXKuh59OmhTGgpjWT72nTsR2FemGgCjItWr8q92msqdoTi5Bu
# sRFNtQtHZBuM/HnhkTcF73Bf+FwVHgp+Trymmh8vOWhqBPzVuhRhu+sH5YP2h4u6
# gu0nKDn4BYh7ol8m24vyZhhgiKVne/h3KsyUC8v/d6YDZDOPQKVvCpeB9Iz/+Wn4
# 7OznknZ0WY/u0fnFtpm+iUrgCI0rzIQqVtwClMiBo8ixBjwCxM78ccsCf37XsUqx
# ktbB80NW6ezkvBorf3bHrVeu0nk8ZGojcK9ernkVMvNnVVVjyQ9PBOR4LmRquQXV
# 8Nn7nu17sG8eOuwIxbu9sGT77DzzCzc6oYIXlDCCF5AGCisGAQQBgjcDAwExgheA
# MIIXfAYJKoZIhvcNAQcCoIIXbTCCF2kCAQMxDzANBglghkgBZQMEAgEFADCCAVIG
# CyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIGVHV7eNAMKG0GqnZiMLQpdf490UVYhxBwBF5pDlEeorAgZoJewX
# v/wYEzIwMjUwNTIyMTUwMDE3Ljc3NFowBIACAfSggdGkgc4wgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo5MjAw
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZaCCEeowggcgMIIFCKADAgECAhMzAAACCQgH4PlcjOZVAAEAAAIJMA0GCSqGSIb3
# DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDEzMDE5
# NDI1NVoXDTI2MDQyMjE5NDI1NVowgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo5MjAwLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBAMKUSjD3Lgzd/VL3PXG00QRPBYvW8SKLDSgPtJcR
# 2/ix0/TGxXKJ2/ojauYSXw9iz0txmPOxY4cjt1CREvbwY/cJdy9jRmrqdawdjZBq
# YJkUsXYiVEfoEfHZGQ3tlEMqazsE6jggYFGUIyRS/033+3A7MCSlY2wzdv8FDFzC
# FWCxCq1Dw0Q9S6JH4ZXmt1AdRPimOKFlOQnCtqWLPRltilRMfk6SLd3cGnH2qI+u
# IHqGE18Y+OXQ8inbcPnv2ulbpmY+o9PyPXYpfvJJnA27Gzc9i8X/DXcaxFeTMhsj
# IsoQ/OP2XOaasXbCO+9SvH0BnDsYtJeTbwOfVdJ/raFuQW5QbA8UuncRtGohWYFn
# jbBzPmZIggLLdCz+HCERiFSd2cAGA2kPlq8As5XuxR8mscNldfp/2CBuMgDqPaeF
# IBIiqXwXkuwoHDRE+0O7LePYI/G1OZmjNssrxMy3EOIwKDFOl+DmJhS/KFXhqpoM
# vBEGygFGE7/6HDJsqdjBfEp546uw7BAudo4TkGYUlhYE4XPd3zwsEr1BEGB0QfkI
# tWHvCSAwh6H3pwfn4fTES+aDq3u7O2VdfZJXvF1Rg/EDe+ONXcSRXtptIcPkcdBl
# Ot3cWqwP9U5gAJRUE+vEX6RStkZfFgidlOmtgxgSrpQgbUNPikJU/0NxoIsYg5gQ
# nWDTAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUSYvo0cRdOOW98C9AzbV3MxaTytIw
# HwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
# UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAw
# XjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# ZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAFxefG84PCTiH+NtQGycWUW2tK4EFlvv
# BJl9rmUpExM182WZoALht3tajQjmEzGwQlTK6kfCHiQPmqRFlzMhzSMgAFBDXENQ
# Fr5ZPGun9QCoLXuKMUJ49kphWM2sd/8GaPPsVo4jjWTG55GHAs0hxDaCYGoNHlbh
# NLaG1EljJkCzuN8mZsO1NxQ4yESXU5aXH8We9xBui3lU/NpTCJPo2J7yXo9mOhCy
# 7GJqy5ICbEohB2wecnlCiSrB3KpeLUVkO0RNW9td8Oyh/NO1rh6fap/jyHMRnBS9
# uTPmya3z3SdUAruTPZyuvM3eGmd8W5+2n+tctZO/E9Bx9ZeIS4hR3YaDt5HxC3Iq
# 0kNTz48PAQKTOhomNsYIqrH0RKAUnPOtc3CGFfpFzyDYRT/7reaapZ4IX+Qk4WDZ
# 4nDtq79psRKCrcRrPIPVWUv4dpf4wEcbNCYe286bdCXjBVM3darxfxsJHryqIXms
# VqybhHEXrNqNl5IcL+pLnffr/howOqxXo7zpGU88JgYk4+1/Yxso7tckl4v9RA3R
# ze6LHlExOjrp1sBPE9QUQbk+Hg8fMaNRsQ7sPfku4QGKIbxiuUxE6QaXd8FCX1tZ
# uDD0IhRBvCrlxNoTGV8Skx1KjJ0miVRNAPkQsobPVMlqFOJ13bTCXCLkGTfpcibO
# wfhizXmJdF8CMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkq
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
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OTIwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMV
# AHzvras9NB3sicMJB1vWSAUpCQJEoIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDr2aSNMCIYDzIwMjUwNTIyMTMy
# NjM3WhgPMjAyNTA1MjMxMzI2MzdaMHQwOgYKKwYBBAGEWQoEATEsMCowCgIFAOvZ
# pI0CAQAwBwIBAAICEkEwBwIBAAICEnswCgIFAOva9g0CAQAwNgYKKwYBBAGEWQoE
# AjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDANBgkq
# hkiG9w0BAQsFAAOCAQEAN9800B5BNZXEjDmC2lQogJnHsMzSUarsOfLGUIoXINSW
# 6RiHgawNevR8N+xsyq+YnoOh80RvIDvUyiDDMN0ao/cM64MPxq0b3KlHc13OgHxN
# tY29rx+PoyVteBdICaVvr3bAtzETMfFnBhZgjZksjtli1dGLjkF4xSyFg3DRJFqi
# A8IrA4YZF3vhYHWCjFha4FluASwMBOnRw0EvtHCnlGPicFvh6k8ZaPHPIhpKH1DN
# RWblOaFFYs8RqdqCHt/znubWHBYwYb+Jm6y6K8+4Ikw01N1zRLj53cyRYso1g89Z
# vcyIqVyKVtIkeC07Mf3go7LbOG6VnbywNkRdcDzE5TGCBA0wggQJAgEBMIGTMHwx
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
# Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAACCQgH4PlcjOZVAAEAAAIJ
# MA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQw
# LwYJKoZIhvcNAQkEMSIEICs4/Wmv3E6mfZGw4oFbR1fSyRYfjqXT4V4aHaunBAOJ
# MIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQgaBssHsi99AIuZQ5RmGN1Sorx
# uKR8HplVV2hOM3CFEz4wgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0Eg
# MjAxMAITMwAAAgkIB+D5XIzmVQABAAACCTAiBCCYi4O3xoqTg3us1d+nxLz1lIMv
# UvsKLBl2o3vVHOZzsDANBgkqhkiG9w0BAQsFAASCAgC9+yyfL5rcv5cAvO3Stwt3
# uEnWIUBLWDUSdZSQqeY+vvVNUdVQajzyTXXgy7A2TYjgJ9QdgJ66+lncYfJTyPkm
# ONSXG//reYaeTaig0yJzAFIXEW/S4FOnt0L9TCZ4gwrsjWB8HNMbLYskgOUxhMTR
# plXU4Uav7xxou6sWbHwxunNRxD6Hl+O4cbmYb+Fa/Q8wqrkZQ18nDwfeS9+1q/0v
# NVv/6bFzVDs9pn1lmG3zC2DNTjZcINUeK39P0IL465N8jLULAKeGiFa/aTkx4MgS
# pye+0WvIa87hhGedt0fe1ZJy5f2oQh9FpqsPDNZ7BXMzFD1Bk5BuB6q7DCAmoemN
# 60X7vnpwSUR5ywydYzocIM+BGoo20fpCwcIAOnkQhiMM1PvdR6riTQFzTq0V/Ju3
# o/a+A1k3mzIdZu5S3qYuUToyDSmMS+IwSqPEvh1GeX5qBzkm2uJS0/vm8wp9XV8Q
# 0C1o3gMYxZEd75BFJkMCVPTqZX3X02DfDPSx/8fon377sfzRs6XG4Gn70WORgxjt
# EIFn3RgpuWqC7HYabgh+2GRidktoqgGu5lC0T8Iu/F+9dy3sL9MbgkXTNnIhGN+w
# 7D0F0Kv9+wV/9UdX1VuxwYzx6pRCM6YWwZcJW20qNxYGaKnHhGLcMH5UNYi2gxkm
# u6DdrXAvii4rLje4mpEi+A==
# SIG # End signature block
