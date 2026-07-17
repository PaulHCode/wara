<#
.CLASS
    Recommendation

.SYNOPSIS
    Represents a recommendation associated with a runbook check.

.DESCRIPTION
    The `Recommendation` class defines a recommendation, including metadata,
    impact assessment, query logic, and automation availability.
    This class is used within the runbook module to define checks
    and their associated guidance.

.PROPERTY AprlGuid
    A unique identifier for the recommendation, if applicable.

.PROPERTY CheckName
    The name of the check associated with this recommendation.

.PROPERTY RecommendationTypeId
    The type identifier for this recommendation.

.PROPERTY RecommendationMetadataState
    The current state of the recommendation (e.g., Active, Deprecated).

.PROPERTY RecommendationControl
    The control category associated with the recommendation.

.PROPERTY LongDescription
    A detailed explanation of the recommendation, its purpose, and implementation guidance.

.PROPERTY Description
    A brief summary of the recommendation.

.PROPERTY PotentialBenefits
    A concise statement highlighting the benefits of implementing the recommendation.

.PROPERTY RecommendationResourceType
    The resource type that the recommendation applies to.

.PROPERTY RecommendationImpact
    The expected impact of following or ignoring the recommendation.

.PROPERTY Query
    The query logic used to evaluate compliance with the recommendation.

.PROPERTY Links
    A hashtable containing additional reference links for further details.

.PROPERTY Tags
    An array of tags categorizing or classifying the recommendation.

.PROPERTY PgVerified
    Indicates whether the recommendation has been verified by the product group.

.PROPERTY AutomationAvailable
    Indicates whether automation is available to apply the recommendation.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
class Recommendation {
    [string] $AprlGuid
    [string] $CheckName
    [string] $RecommendationTypeId
    [string] $RecommendationMetadataState
    [string] $RecommendationControl
    [string] $LongDescription
    [string] $Description
    [string] $PotentialBenefits
    [string] $RecommendationResourceType
    [string] $RecommendationImpact
    [string] $Query

    [hashtable] $Links = @{}

    [string[]] $Tags = @()

    [bool] $PgVerified
    [bool] $AutomationAvailable
}

<#
.CLASS
    RunbookRecommendation

.SYNOPSIS
    Represents a recommendation within a runbook.

.DESCRIPTION
    The `RunbookRecommendation` class encapsulates a specific recommendation
    as part of a runbook. It includes metadata such as the check set name,
    the individual check name, and an associated `Recommendation` object.

.PROPERTY CheckSetName
    The name of the check set that this recommendation belongs to.

.PROPERTY CheckName
    The name of the specific check within the check set.

.PROPERTY Recommendation
    The `Recommendation` object associated with this runbook entry.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
class RunbookRecommendation {
    [string] $CheckSetName
    [string] $CheckName

    [Recommendation] $Recommendation
}

<#
.CLASS
    RunbookQuery

.SYNOPSIS
    Represents a query associated with a runbook check.

.DESCRIPTION
    The `RunbookQuery` class stores query details for a specific check within
    a runbook check set. It includes metadata such as the check set and check name,
    the query string, selector name, associated tags, and the linked recommendation.

.PROPERTY CheckSetName
    The name of the check set containing this query.

.PROPERTY CheckName
    The name of the specific check associated with the query.

.PROPERTY SelectorName
    The selector used for filtering resources in the query.

.PROPERTY Query
    The query string to evaluate the check.

.PROPERTY Tags
    An array of tags associated with this query.

.PROPERTY Recommendation
    The `Recommendation` object related to this query.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>

class RunbookQuery {
    [string] $CheckSetName
    [string] $CheckName
    [string] $SelectorName
    [string] $Query

    [string[]] $Tags

    [Recommendation] $Recommendation
}

<#
.CLASS
    RunbookCheck

.SYNOPSIS
    Represents a check within a runbook.

.DESCRIPTION
    The `RunbookCheck` class defines an individual check, including its selector,
    parameters, and associated tags. It is part of a `RunbookCheckSet` and is used
    to evaluate specific conditions within a runbook.

.PROPERTY SelectorName
    The name of the selector applied to this check.

.PROPERTY Parameters
    A hashtable of parameters specific to this check.

.PROPERTY Tags
    An array of tags categorizing this check.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
class RunbookCheck {
    [string] $SelectorName

    [hashtable] $Parameters = @{}

    [string[]] $Tags = @()
}

<#
.CLASS
    RunbookCheckSet

.SYNOPSIS
    Represents a set of runbook checks.

.DESCRIPTION
    The `RunbookCheckSet` class groups multiple `RunbookCheck` objects, allowing them
    to be organized and evaluated together as part of a runbook.

.PROPERTY Checks
    A hashtable containing `RunbookCheck` objects.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
class RunbookCheckSet {
    [hashtable] $Checks = @{}
}

<#
.CLASS
    Runbook

.SYNOPSIS
    Represents a runbook containing check sets, selectors, and parameters.

.DESCRIPTION
    The `Runbook` class defines the structure of a runbook, including parameters,
    variables, selectors, and check sets. It provides a `Validate` method to
    ensure the runbook is correctly configured before execution.

.PROPERTY Parameters
    A hashtable of parameters defined within the runbook.

.PROPERTY Variables
    A hashtable of variables used in the runbook.

.PROPERTY Selectors
    A hashtable of selectors that define resource filters.

.PROPERTY CheckSets
    A hashtable of `RunbookCheckSet` objects.

.METHOD Validate
    Ensures the runbook is properly configured by verifying that all required elements exist.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
class Runbook {
    [hashtable] $Parameters = @{}
    [hashtable] $Variables = @{}
    [hashtable] $Selectors = @{}
    [hashtable] $CheckSets = @{}

    [void] Validate() {
        $errors = @()

        if ($this.Selectors.Count -eq 0) {
            $errors += "- [selectors]: At least one (1) selector is required."
        }

        if ($this.CheckSets.Count -eq 0) {
            $errors += "- [checks]: At least one (1) check set is required."
        }

        foreach ($checkSetKey in $this.CheckSets.Keys) {
            $checkSet = $this.CheckSets[$checkSetKey]

            foreach ($checkKey in $checkSet.Checks.Keys) {
                $check = $checkSet.Checks[$checkKey]
                $checkTitle = "[$checkSetKey]:[$checkKey]"

                if (-not $this.Selectors.ContainsKey($check.SelectorName)) {
                    $errors += "- [checks]: $checkTitle references a selector that does not exist: [$($check.SelectorName)]."
                }
            }
        }

        if ($errors.Count -gt 0) {
            throw "Runbook is invalid:`n$($errors -join "`n")"
        }
    }
}

<#
.CLASS
    SelectorReview

.SYNOPSIS
    Stores the results of a selector review.

.DESCRIPTION
    The `SelectorReview` class contains resolved selectors and their associated
    resources after evaluation. It helps users verify that selectors are correctly
    configured and returning the expected results.

.PROPERTY Selectors
    A hashtable mapping selector names to selected resource sets.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
class SelectorReview {
    [hashtable] $Selectors = @{}
}

<#
.CLASS
    SelectedResourceSet

.SYNOPSIS
    Represents a set of resources selected by a query.

.DESCRIPTION
    The `SelectedResourceSet` class stores the results of a resource selection
    based on a selector and an Azure Resource Graph query.

.PROPERTY Selector
    The selector that determined the resource set.

.PROPERTY ResourceGraphQuery
    The query used to retrieve resources.

.PROPERTY Resources
    An array of `SelectedResource` objects.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
class SelectedResourceSet {
    [string] $Selector
    [string] $ResourceGraphQuery

    [SelectedResource[]] $Resources = @()
}

<#
.CLASS
    SelectedResource

.SYNOPSIS
    Represents a selected resource.

.DESCRIPTION
    The `SelectedResource` class stores details about a resource, including
    its ID, type, name, location, resource group, and tags.

.PROPERTY ResourceId
    The unique identifier of the resource.

.PROPERTY ResourceType
    The type of the resource.

.PROPERTY ResourceName
    The name of the resource.

.PROPERTY ResourceLocation
    The geographical location of the resource.

.PROPERTY ResourceGroupName
    The resource group that contains the resource.

.PROPERTY ResourceTags
    A hashtable of key-value pairs representing resource tags.

.NOTES
    Author: Casey Watson
    Date: 2025-02-27
#>
class SelectedResource {
    [string] $ResourceId
    [string] $ResourceType
    [string] $ResourceName
    [string] $ResourceLocation
    [string] $ResourceGroupName

    [hashtable] $ResourceTags = @{}
}

<#
.CLASS
    RunbookFactory

.SYNOPSIS
    Parses and creates Runbook objects.

.DESCRIPTION
    The `RunbookFactory` class provides methods to parse runbooks from a JSON file
    or raw JSON content, returning a `Runbook` instance.

.METHOD ParseRunbookFile
    Reads and parses a runbook from a JSON file.

.METHOD ParseRunbookContent
    Parses a runbook from raw JSON content.

.NOTES
    - If the file does not exist, `ParseRunbookFile` returns `$null`.
    - Supports both simple and structured check definitions.
    - Automatically initializes missing properties as empty collections.

    Author: Casey Watson
    Date: 2025-02-27
#>
class RunbookFactory {
    [Runbook] ParseRunbookFile([string] $path) {
        if (Test-FileExists -Path $path) {
            $fileContent = Get-Content -Path $path -Raw
            return $this.ParseRunbookContent($fileContent)
        }

        return $null
    }

    [Runbook] ParseRunbookContent([string] $content) {
        $runbookHash = ($content | ConvertFrom-Json -AsHashtable)

        $runbook = [Runbook]@{
            Parameters = ($runbookHash.parameters ?? @{})
            Variables  = ($runbookHash.variables ?? @{})
            Selectors  = ($runbookHash.selectors ?? @{})
        }

        foreach ($checkSetKey in $runbookHash.checks.Keys) {
            $checkSet = [RunbookCheckSet]::new()
            $checkSetHash = $runbookHash.checks[$checkSetKey]

            foreach ($checkKey in $checkSetHash.Keys) {
                $check = [RunbookCheck]::new()
                $checkValue = $checkSetHash[$checkKey]

                switch ($checkValue.GetType().Name.ToLower()) {
                    "string" {
                        $check.SelectorName = $checkValue
                    }
                    "orderedhashtable" {
                        $check.SelectorName = $checkValue.selector
                        $check.Parameters = ($checkValue.parameters ?? @{})
                        $check.Tags = ($checkValue.tags ?? @())
                    }
                }

                $checkSet.Checks[$checkKey] = $check
            }

            $runbook.CheckSets[$checkSetKey] = $checkSet
        }

        return $runbook
    }
}

# SIG # Begin signature block
# MIIoLwYJKoZIhvcNAQcCoIIoIDCCKBwCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCKza/aZJ8QFRVz
# p1U7gk+Ldpz7iMvFtO7L+vd1ALufmaCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEINA7t5Oa3sJfgfjLwrmon/Q5
# 389MOTlgTjYlrOocFlR0MEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQCoML1OAkkin7Tkv/2V3smQtSctGYmj542nML6kZNIcU9ee9SzX4PyS
# w6Fz6BRjQRoV1Xcyp38vXzVhOPMGyoWRdXoZUOGtEj0kt/jAkc92YUTyBjtFGA5t
# uHwomjJUgKpcoTMDaGHD7R0+ESjfvorn0fV3v02KBbdQ/96DKN1GgmOShKaRCk5N
# 7E4F7d+XM03wxI3bVHt4us2SaoRfy2JuT/OA6eNmtp5gzjUQIqp9VqeWZAw/rjTe
# UJJQDK65Dy2H3LsqrHILnkqigzRiHhVQoi84v7MBq7EpGh66ojnWSSayNTUKiFmh
# Fc4rKWMAmhjcdoHUondILuQkTnNFMF0roYIXlzCCF5MGCisGAQQBgjcDAwExgheD
# MIIXfwYJKoZIhvcNAQcCoIIXcDCCF2wCAQMxDzANBglghkgBZQMEAgEFADCCAVIG
# CyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIHOlB2P+YnA9FdWBEXtOEdeJQJIZXmivHHiRoPaNlOKjAgZoJhQ/
# EggYEzIwMjUwNTIyMTUwMDUzLjY1MVowBIACAfSggdGkgc4wgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo4NjAz
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZaCCEe0wggcgMIIFCKADAgECAhMzAAACBywROYnNhfvFAAEAAAIHMA0GCSqGSIb3
# DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDEzMDE5
# NDI1MloXDTI2MDQyMjE5NDI1MlowgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo4NjAzLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBAMU//3p0+Zx+A4N7f+e4W964Gy38mZLFKQ6fz1kX
# K0dCbfjiIug+qRXCz4KJR6NBpsp/79zspTWerACaa2I+cbzObhKX35EllpDgPHeq
# 0D2Z1B1LsKF/phRs/hn77yVo1tNCKAmhcKbOVXfi+YLjOkWsRPgoABONdI8rSxC4
# WEqvuW01owUZyVdKciFydJyP1BQNUtCkCwm2wofIc3tw3vhoRcukUZzUj5ZgVHFp
# OCpI+oZF8R+5DbIasBtaMlg5e555MDUxUqFbzPNISl+Mp4r+3Ze4rKSkJRoqfmzy
# yo1sjdse3+sT+k3PBacArP484FFsnEiSYv6f8QxWKvm7y7JY+XW3zwwrnnUAZWH7
# YfjOJHXhgPHPIIb3biBqicqOJxidZQE61euc8roBL8s3pj7wrGHbprq8psVvNqpZ
# cCPMSJDwRj0r2lgj8oLKCLGMPAd9SBVJYLJPwrDuYYHJRmZE8/Fc42W4x78/wK0E
# kym6HwIFbKO8V8WY5I1ErwRORSaVNQBHUIg5p4GosbCxxKEV/K8NCtsKGaFeJvid
# ExflT1iv13tVxgefp5kmyDLOHlAqUhsJAL9i+EUrjZx4IEMxtz463lHpP8zBx7mN
# XJUKapdXFY5pBzisDadXuicw5kLpS8IbwsYVJkGePWWgMMtaj8j5G5GiTaP9DjNw
# yfCRAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUcrVSYsK9etAK9H3wkGrXz/jOjR4w
# HwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
# UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAw
# XjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# ZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAOO7Sq49ueLHSyUSMPuPbbbilg48ZOZ0
# O87T5s1EI2RmpS/Ts/Tid/Uh/dj+IkSZRpTvDXYWbnzYiakP8rDYKVes0os9ME7q
# d/G848a1qWkCXjCqgaBnG+nFvbIS6cbjJlDoRA6mDV0T245ejN7eAPgeO1xzvmRx
# rzKK+jAQj6uFe5VRYHu+iDhMZTEp2cO+mTkZIZec6E8OF0h36DqFHJd1mLCARr6r
# 0z1dy3PhMaEOA4oWxjEWFc0lmj0pG4arp6+G3I125iuTOMO1ZLqBbxqRHn1SG4sa
# xWr7gCCoRjxaVeNAYzY5OTIGeVAukHyoPvH2NGljYKrQ5ZaUrTB8f/XN5+tY3n5t
# 7ztLDZM9wi50gmff1tsMbtrAoxVgMd+w8nxm/GBaRm5/plkCSmHR5gaHchXzjm1o
# uR0s4K/Dj1bGqFrkOaLY6OHwaNrm/2TJjcpMXJfdPgLaxzF+Cn/rFF34MY6E1U+9
# U9r/fJFSpjmzlRinLwOdumlXudA7ax7ce8JJutv7I/J6hvWRR8xhr18TZsSygxs5
# odGAaOLxk+38l3Zs991CgEdxQ6o/CMcFQhxJzvF0lliNFvibzWrGOZrcMuO44WWM
# xlNii9GIa8Qwv3FmPakdFTK/6zm/tUbBwzquM1gzirNlAzoDZEZgkZTvzQZAbRA7
# 3zD6y5y5NWt9MIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkq
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
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046ODYwMy0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMV
# ANO9VT9iP2VRLJ4MJqInYNrmFSJLoIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDr2SP3MCIYDzIwMjUwNTIyMDQx
# NzU5WhgPMjAyNTA1MjMwNDE3NTlaMHcwPQYKKwYBBAGEWQoEATEvMC0wCgIFAOvZ
# I/cCAQAwCgIBAAICGMkCAf8wBwIBAAICElIwCgIFAOvadXcCAQAwNgYKKwYBBAGE
# WQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDAN
# BgkqhkiG9w0BAQsFAAOCAQEAh2lqcvxuNa8Sr6b536cbvcwGHm5a1VLMgvXESDbd
# 6dj1lnuuskKHQxrO1uSN0sTXqwX3rEvmvu9YI/bUNT/LLFKzNvSqUHu7xcLEKQtB
# RUk2J3yKWdYHMIc/fKr215oO96fV/tD/EqF2AdiglFjAwRbJT2f0Hxu7WK6ObcSg
# 3Xp2gP5uPRroA11H1ZVnMbohOccTWSP6ufM6S6vFnO+kP0uZgJpQjZ461gL9+Whb
# +uQJ8/VvSebBXlI/ujuYLhAdFWSqbxrJ8pq4+byw6h7FbssPTZXSz4vqBrrKwxNA
# qXukXJtQ/CLO9f7a/ps+qb5STmZSEQvbxhJ4xKJG3PNk1zGCBA0wggQJAgEBMIGT
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAACBywROYnNhfvFAAEA
# AAIHMA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQ
# AQQwLwYJKoZIhvcNAQkEMSIEIC8GLB2SVGaZwPYwmiM+1A/pHI4GSGoF8UWnTiTf
# vFtNMIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQgL/fU0dB2zUhnmw+e/n2n
# 6/oGMEOCgM8jCICafQ8ep7MwgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0EgMjAxMAITMwAAAgcsETmJzYX7xQABAAACBzAiBCDso1xoAIKuELuQZDyoRoia
# 3NgIwh8KyIl1T7MdMiX8IzANBgkqhkiG9w0BAQsFAASCAgAbAaxGUmfC/CXAe0dH
# 3OliHACaxSDkhJwasICdGmpdBDcnxHPweGoxbHNWgKHH956BJx11hxwGdh2PxYsS
# Vc52BXJffNFqXqHnIp9mMhAw9zlYNcVEXvhDmui2ijb465z13XlhHSWV05Z5WIHQ
# RFJ9l5L2//yjHN65unaslPF0oslrFl1mtSo7WWnrNzL6MkurLeJkDd/oGqdiEUIa
# hnSp5DZpT9FMwLSBaHDPqMWyRsDSLLc2w7Z/EpLzNg3CBWntIykYdEIWMknL5ctT
# 2eo/47w8ZY17Zy4UUcMFeO5R9RohG12zxqSw8SsKS/6qbn8G5KJ4V4kWh9MucTA+
# t4Pt6NuLVDzusS5pWsarOCRYpdt7Q3ffjkAqAvVTlz9QYyjNqRinx0EKcJrJuMIr
# aucZkJM4g7xn8/zU7b9jmapLun5qr42BLZ7DWgIQmYXW1razIQfMddiKQDnQUfU2
# +mfblHqtK5HKuHfP4jcxv2c1M9RSMfYYbk/gn3EywA/xG9xlVoxRZiF3haZ39rHY
# xh7S+f0RiAVNWFwFuXfDZQGHOn9RTJgTpHF0GuPXJzA2NguhuKjd4Cr6wQI8FDtg
# 5XPK8OXuIVaE+Csn82HTX5DnJxauMNL/jemnMP0pXoaOi8e4mv6TOq79vfgAM7qZ
# 8z5PTBpN8qx4GYW5QHzfz1MsYg==
# SIG # End signature block
