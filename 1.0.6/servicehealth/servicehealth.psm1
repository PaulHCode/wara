using module ../utils/utils.psd1
<#
.SYNOPSIS
    Retrieves service health alerts from specified Azure subscriptions.

.DESCRIPTION
    The `Get-WAFServiceHealth` function queries Azure Resource Graph to retrieve service health alerts from the provided subscription IDs. It searches for activity log alerts related to 'ServiceHealth' and joins them with subscription information.

.PARAMETER SubscriptionIds
    An array of Azure subscription IDs for which to retrieve service health alerts.

.INPUTS
    None. You cannot pipe objects to this function.

.OUTPUTS
    System.Object[]. Returns an array of service health alert objects.

.EXAMPLE
    # Retrieve service health alerts for multiple subscriptions
    $subscriptions = @('59f6f1ab-6d68-4c90-b4e5-ad2d71cefc57', 'abcd1234-5678-90ab-cdef-1234567890ab')
    $serviceHealthAlerts = Get-WAFServiceHealth -SubscriptionIds $subscriptions

.EXAMPLE
    # Retrieve service health alerts for a single subscription
    $serviceHealthAlerts = Get-WAFServiceHealth -SubscriptionIds '59f6f1ab-6d68-4c90-b4e5-ad2d71cefc57'

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
function Get-WAFServiceHealth {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]] $SubscriptionIds
    )

    $Servicequery = `
        "resources
| where type == 'microsoft.insights/activitylogalerts' and properties.condition has 'ServiceHealth'
| join kind=inner (
    resourcecontainers
    | where type == 'microsoft.resources/subscriptions'
    | project subscriptionId, subscriptionName = name
) on subscriptionId
| project subscriptionId, subscriptionName, eventName = name, type, location, resourceGroup, properties"

    $queryResults = Invoke-WAFQuery -Query $Servicequery -SubscriptionIds $SubscriptionIds

    $AllServiceHealth = Build-WAFServiceHealthObject -AdvQueryResult $queryResults
    
    return $AllServiceHealth
}

<#
.SYNOPSIS
    Builds service health alert objects from query results.

.DESCRIPTION
    The `Build-WAFServiceHealthObject` function processes the results obtained from the Azure Resource Graph query and constructs custom objects representing service health alerts with relevant details.

.PARAMETER AdvQueryResult
    The results from the Azure Resource Graph query.

.INPUTS
    System.Object[]. Accepts an array of query result objects.

.OUTPUTS
    System.Object[]. Returns an array of service health alert objects with detailed properties.

.EXAMPLE
    # Process query results to get service health alert objects
    $queryResults = Invoke-WAFQuery -Query $Servicequery -SubscriptionIds $SubscriptionIds
    $serviceHealthAlerts = Build-WAFServiceHealthObject -AdvQueryResult $queryResults

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
function Build-WAFServiceHealthObject {
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [PSCustomObject[]] $AdvQueryResult
    )

    $return = $AdvQueryResult.ForEach({ [ServiceHealthAlert]::new($_) })

    return $return
}


<#
.SYNOPSIS
    Represents a Service Health Alert retrieved from Azure.

.DESCRIPTION
    The `ServiceHealthAlert` class encapsulates the details of a service health alert from Azure. It provides properties to access alert information such as the name, subscription, status, event type, affected services, regions, and associated action groups. The class includes methods to parse and extract specific details from the alert data.

.PROPERTIES
    [string] Name
        The name of the service health alert.

    [string] Subscription
        The name of the Azure subscription where the alert is configured.

    [string] Enabled
        Indicates whether the alert is enabled or disabled.

    [string] EventType
        The type of event the alert is configured to monitor (e.g., 'Service Health Incident', 'Planned Maintenance').

    [string] Services
        A comma-separated list of Azure services that the alert monitors.

    [string] Regions
        A comma-separated list of Azure regions that the alert monitors.

    [string] ActionGroup
        The name of the action group associated with the alert for notifications.

.CONSTRUCTORS
    ServiceHealthAlert([PSCustomObject] $Row)
        Initializes a new instance of the `ServiceHealthAlert` class using the provided alert data.

.METHODS
    static [string] GetEventType([PSCustomObject] $Row)
        Parses and returns the event type from the alert data.

    static [string] GetServices([PSCustomObject] $Row)
        Extracts and returns the services monitored by the alert.

    static [string] GetRegions([PSCustomObject] $Row)
        Extracts and returns the regions monitored by the alert.

    static [string] GetActionGroupName([PSCustomObject] $Row)
        Retrieves and returns the name of the action group associated with the alert.

.EXAMPLE
    # Example of creating ServiceHealthAlert objects from query results
    $queryResults = Invoke-WAFQuery -Query $ServiceQuery -SubscriptionIds $SubscriptionIds
    $serviceHealthAlerts = $queryResults.ForEach({ [ServiceHealthAlert]::new($_) })

.NOTES
    Author: Kyle Poineal
    Date: 2024-12-12
#>
class ServiceHealthAlert {
    [string] $Name
    [string] $Subscription
    [string] $Enabled
    [string] $EventType
    [string] $Services
    [string] $Regions
    [string] $ActionGroup

    ServiceHealthAlert([PSCustomObject]$row) {
        $this.Name = $Row.eventName
        $this.Subscription = $Row.subscriptionName
        $this.Enabled = $Row.properties.enabled
        $this.EventType = [ServiceHealthAlert]::GetEventType($Row)
        $this.Services = [ServiceHealthAlert]::GetServices($Row)
        $this.Regions = [ServiceHealthAlert]::GetRegions($Row)
        $this.ActionGroup = [ServiceHealthAlert]::GetActionGroupName($Row)
    }

    static [string] GetEventType($Row) {
        $equals = ($Row.Properties.condition.allOf | Where-Object { $_.field -eq 'properties.incidentType' } | Select-Object -Property equals).equals
        $return = switch ($equals) {
            'Incident' { 'Service Issues' }
            'Informational' { 'Health Advisories' }
            'ActionRequired' { 'Security Advisory' }
            'Maintenance' { 'Planned Maintenance' }
            default { 'All' } 
        }

        return $return
    }

    static [string] GetServices($Row) {
        if ($Row.Properties.condition.allOf | Where-Object { $_.field -eq 'properties.impactedServices[*].ServiceName' }) {
            return ($Row.Properties.condition.allOf | Where-Object { $_.field -eq 'properties.impactedServices[*].ServiceName' } | Select-Object -Property containsAny | ForEach-Object { $_.containsAny }) -join ', '
        }
        else {
            return 'All'
        }
    }

    static [string] GetRegions($Row) {
        if ($Row.Properties.condition.allOf | Where-Object { $_.field -eq 'properties.impactedServices[*].ImpactedRegions[*].RegionName' }) {
            return ($Row.Properties.condition.allOf | Where-Object { $_.field -eq 'properties.impactedServices[*].ImpactedRegions[*].RegionName' } | Select-Object -Property containsAny | ForEach-Object { $_.containsAny }) -join ', '
        }
        else {
            return 'All'
        }
    }

    static [string] GetActionGroupName($Row) {
        if ($Row.Properties.actions.actionGroups.actionGroupId) {
            return $Row.Properties.actions.actionGroups.actionGroupId.split('/')[8]
        }
        else {
            return ''
        }
    }
}

# SIG # Begin signature block
# MIIoLgYJKoZIhvcNAQcCoIIoHzCCKBsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB2ep70/XH+BlA1
# 9RlOWQ5XDtHVtdkqOFxjP9B69fqbw6CCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
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
# /Xmfwb1tbWrJUnMTDXpQzTGCGg4wghoKAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAQEbHQG/1crJ3IAAAAABAQwDQYJYIZIAWUDBAIB
# BQCggbAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIE9ogH0HwCYP3FfRl2Zs62KR
# u5HBazfwo5tJadxgDFVGMEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQAe4sO93D1qBfiAXnOsiloypnKfUAebjtbkIrpjizu7kQhbzUU6gW1D
# 49YoZpchedQjohjcfgEdDrGRhU+8KE4KCUXny6naC+kdLHuLuBfgHYMLyxh/8LN5
# X91JM+RA4Z9cEFqYxm5LUJgksJeFXbnGm33fx9BWhHUWnfyds70M1jwm++JDSNyA
# UsQmYpQ2hf1uV3QSC90uBK6nsTBgTU32xOy5Mwe3PEkSOC7u+jz+z0GxbbhCozEE
# 5Jef4ug8MRu082F8rbdAXyrdmnC+G9zag7sCglidN8oskYHxOv9YlPrOKB8IsjhA
# nzYLIYhZTwY/PSaWMek9oyDe7UkyzJCnoYIXljCCF5IGCisGAQQBgjcDAwExgheC
# MIIXfgYJKoZIhvcNAQcCoIIXbzCCF2sCAQMxDzANBglghkgBZQMEAgEFADCCAVEG
# CyqGSIb3DQEJEAEEoIIBQASCATwwggE4AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIIAGpH/k01Vp+2M55ZAB7GcTd92kVPD/htR/TYaKeAEmAgZoJmAf
# yJYYEjIwMjUwNTIyMTUwMDQwLjQxWjAEgAIB9KCB0aSBzjCByzELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFt
# ZXJpY2EgT3BlcmF0aW9uczEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkYwMDIt
# MDVFMC1EOTQ3MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
# oIIR7TCCByAwggUIoAMCAQICEzMAAAIFPHVsgkSHzf4AAQAAAgUwDQYJKoZIhvcN
# AQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcNMjUwMTMwMTk0
# MjQ5WhcNMjYwNDIyMTk0MjQ5WjCByzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldh
# c2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9u
# czEnMCUGA1UECxMeblNoaWVsZCBUU1MgRVNOOkYwMDItMDVFMC1EOTQ3MSUwIwYD
# VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEAkpLy33e4Bda9sBncvOQhWFx1AvMsBMg+C0S79FmB
# F3nmdLuWLiu6dnF1c0JmTzh0zfE1qhtkj5VG/uz5XcxQwwJUd71PKYjo5obvax1u
# NzNnW6K/Y5fYJboc8FHdknIlRmu3/beu7TNyhSkUjFxbRyhdysAQe2laPm9asuaf
# Q1paNjeRRqwahzBFZTcs63h2KAyy/pvH0rKjLv4mFKscyuReEuyGOTXpgAfAfgN0
# IMFSIuuCiSH3imVHoligk3+KHVID9wEIpcYePD+s+wE+CANHTBLSoWCxbOFvyjQz
# LGK+yqUDylQnAuRPLgx3SnsLm8s3p5E8cuH39Td4PMoaOT4vQX40dFcra5JqQ33q
# fCT8HG+ATTiFzqNaX3R2fBL50eyRWRUIqqTGRZTuQgLk2B/Lo3OT1B5WjACfDRGv
# UxSUzkgawez0YHof+jSdsbvcsT4f5FTfQRrLPdzAulI6aMXjOMe9G8G8IivEjRyD
# vA/HKpe1Unr1GG4zeDaIBRcIQQpYaHRP83hj6usuosQ+M+uSB2N88BUGwVV/8Pi/
# 1RzZ/wTBrNjxh55UYzvypPDSKTeLIZBUKgNXzNPH66w0jRGPVSg7abFKQBedWNaE
# OrSYVjNXd53gl4em/+jfl3hzkQsJ2PNyvqRTDIYPIrF0G+ikZeuZIPF2AXeCcJGy
# qFUCAwEAAaOCAUkwggFFMB0GA1UdDgQWBBR0elq7Nu2+vsid2xGfaOTXS9Wy8DAf
# BgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQ
# hk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQl
# MjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBe
# MFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2Nl
# cnRzL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAM
# BgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQE
# AwIHgDANBgkqhkiG9w0BAQsFAAOCAgEADrsZOO29Yu+VfNU8esaNdMTSK+M2cWFX
# 5BeUxatpJ3Tx4M1ci57LMPxypBGUQoGVaZChCemOI7xubboDIvlo7e4VDEoqZPka
# QeYBUL4dcZgBC9n5XoM01hLJ49MKxEqZSOWd74H9nhlwK/0XKho0qaLh2w9h2PWN
# xdDpehUQwlfxxBikR859jOa0KRRko2nE+A5KlWJnpvwKzn0r1aI5yhCFvdeFMRrb
# oSUq/YzqOUak1+xiKm7bze84VpXfot18XYXTXH5UM/WIaBakHsQXp6CEYADwLcB+
# vMXM6/SzAt5fQCxKZ7LztEYij1xeJdtvzn3BX32qYZ5f0w8JIiX8TsgDH1Bd8SPf
# t4s09Vl9ghbNkWjgKt3XKIcicPsURtBPMJAh6pFeewW1ARMy1/C/ZRidQ6MWDaaA
# 1+4kMyfUHZMqYuX7++9xNwofAPraMXhaehYn0GcgnPCHCAZR8mpOjG0+mE1UDYEP
# 4fBRfkuTqj+whAhbyB9irdj9BpTrvQtAX2rIZ046HZrWRWbKbVL4q5P9hziy4wYj
# Iw8CbEABQMybs+GbU8qK67xEddBpf5m5lYh6obzQAn08z4i34w4Mr6fbO/2x7vwm
# pSpnoiVCxo4f5cAI+d9faYILBiam4SeBWxXPqFOc3325v6yo1WfJMTQ94ptdEKeN
# Z9rf6qcj+hEwggdxMIIFWaADAgECAhMzAAAAFcXna54Cm0mZAAAAAAAVMA0GCSqG
# SIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkg
# MjAxMDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMyMjVaMHwxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK4OK9uT4X
# YDP/XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY6GB9alKDRLemjkZrBxTz
# xXb1hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk2pg7
# uhp7M62AW36MEBydUv626GIl3GoPz130/o5Tz9bshVZN7928jaTjkY+yOSxRnOlw
# aQ3KNi1wjjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDuaRr3tpK56KTesy+uDRedG
# bsoy1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74kpEeHT39IM9zfUGaRnXN
# xF803RKJ1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2K26oElHovwUDo9Fzpk03
# dJQcNIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5TI4CvEJoLhDqhFFG4tG9
# ahhaYQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZki1ugpoMhXV8wdJGUlNi5
# UPkLiWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9QBXpsxREdcu+N+VLEhReT
# wDwV2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3PmriLq0CAwEAAaOCAd0wggHZ
# MBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5kxJq+gpE8
# RjUpzxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJlpxtTNRnpcjBcBgNVHSAE
# VTBTMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wEwYDVR0lBAww
# CgYIKwYBBQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQD
# AgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb
# 186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29t
# L3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoG
# CCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwDQYJKoZI
# hvcNAQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/ypb+pcFLY+TkdkeLEGk5c9
# MTO1OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYeeNg2Lpyp
# glYAA7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM9W0jVOR4U3UkV7ndn/OO
# PcbzaN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSFQrP8
# DJ6LGYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4FOmRsqlb30mjdAy87JGA
# 0j3mSj5mO0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3UwxTSwethQ/gpY3UA8x1Rt
# nWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPXfx5bRAGOWhmRaw2fpCjc
# ZxkoJLo4S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyhYWxz/gq7
# 7EFmPWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGConsXHRWJjXD+57XQKBqJ
# C4822rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAxM328
# y+l7vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGbyoYID
# UDCCAjgCAQEwgfmhgdGkgc4wgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlvbnMx
# JzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjpGMDAyLTA1RTAtRDk0NzElMCMGA1UE
# AxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcGBSsOAwIaAxUA
# 1bB/adbSZ/pK8AjL6joVb1623rSggYMwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEG
# A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
# cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
# cCBQQ0EgMjAxMDANBgkqhkiG9w0BAQsFAAIFAOvZb98wIhgPMjAyNTA1MjIwOTQx
# NTFaGA8yMDI1MDUyMzA5NDE1MVowdzA9BgorBgEEAYRZCgQBMS8wLTAKAgUA69lv
# 3wIBADAKAgEAAgIPdQIB/zAHAgEAAgITfDAKAgUA69rBXwIBADA2BgorBgEEAYRZ
# CgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIBAAIDAYagMA0G
# CSqGSIb3DQEBCwUAA4IBAQCvcst8tyR1nQ7EA08cELKD3yMMBimIRG4AJVqgcAXT
# MP3IjpWqZCeaAtRk3lZsoM9S+MHLNStP/i3QG0/Ov04U4TUOFyvxWy4n15ro4OKz
# 9XQldPT+cEP/Ydr9ZW3yMxF7b6BNkVYlbRT2WMNo5qgvnfj+g9TOKLFn9H0SGEpp
# 0Blpyx0ph1uI6DWUUYHgKfbOgX88EoSziOWyvDyiCcZco0SgTw3jTEQBAxAzkZA3
# KxfrLzcbwUocZu888Qc6hckbXsqs5BH0AmIcoNQnvEKAaTqcNfuwHxwHAsqAfBJE
# sC5P8chLHg9WSZCy+63CnLdlbvgNI/Cd5BLtA6F1qcQEMYIEDTCCBAkCAQEwgZMw
# fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
# ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMd
# TWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAIFPHVsgkSHzf4AAQAA
# AgUwDQYJYIZIAWUDBAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRAB
# BDAvBgkqhkiG9w0BCQQxIgQgZ8xUU9+AIAZpv1MXicEjCcu27A1v3ZNQXSxAJs23
# +iYwgfoGCyqGSIb3DQEJEAIvMYHqMIHnMIHkMIG9BCCADQM93HmNLpoXVi0drCaa
# tDj6rSQ0wGEZox1ZMBFvSDCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
# EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
# ZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
# QSAyMDEwAhMzAAACBTx1bIJEh83+AAEAAAIFMCIEIMyuVamg0O+v1QeEtPpdYrKB
# /NAkErVicdfTZHi6AZM9MA0GCSqGSIb3DQEBCwUABIICAEW1S1XKlas5bfSciNw5
# yoVmar5GKYfjtCPrpf8bicQwdHFjFJTXxL6H2dOBmJwMfZuBhLhqD957BtjiDeCE
# Bqq4uUBbqxykfN7oZYqBHyxXWALJ19iPu/zLCTt5JtAHs7g/4XVXHnjFs38EW8z1
# tQ4D2gNmKaeefl7jt0Zc5m+067/NOIJf9zhRjp1z+0SjIJb15eW0rUgGiC1z7LmE
# WJe79t0ppJHBuddJiGBjZAKZcwS+EwHegvP1pmzfs/PyH8Qxu8eAR7/mDF4j2HzI
# p6a+l00yDyaqfISN1yKWU507ft4n/8Ps4s1D5ifjNrcMsRnfaR6HR9BdrS1kuzwj
# iQOW5flle/JoMPaPu3RC9CUo/Q6kPnOaPzX/uiuxovZllqu6nv+oO9CzO4vQewZX
# cmCTQklvya7zKV7dhQVV7KdwXTXJjLv3vVEUHi/N5SztZ6YsWz/8KSh4/jqhmPHS
# IBgzrpImgur6RkQcZkc3/R6w+F9WBMrzLc/e0TpXcan8r7U8GxiEIdolQ6rV6YaR
# mQBcKXw9WVRagQ1QOauOz0mz6yHWjWWwUE8I9r9pNVtFu7ytkSlCKWBeCANKtx1W
# wLGqiKKLva9HUuzNuKAOkA80Johvw/hQsFFeMwm8dH9fZ42MYCZbCEXBUofz+L6i
# n1S6aqbZQVKP5b5XZRL0y3Ld
# SIG # End signature block
