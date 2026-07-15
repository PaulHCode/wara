#Requires -Version 7

<#
.SYNOPSIS
Well-Architected Reliability Assessment Script

.DESCRIPTION
The script "2_wara_data_analyzer" will process the JSON file created by the "1_wara_collector" script and will edit the Expert-Analysis Excel file.

.PARAMETER Help
Switch to display help information.

.PARAMETER RepositoryUrl
Specifies the git repository URL that contains APRL contents if you want to use custom APRL repository.

.PARAMETER JSONFile
Path to the JSON file created by the "1_wara_collector" script.

.PARAMETER ExpertAnalysisFile
Path to the Expert-Analysis file to be customized by script.

.EXAMPLE
.\2_wara_data_analyzer.ps1 -JSONFile 'C:\Temp\WARA_File_2024-04-01_10_01.json' -ExpertAnalysisFile 'C:\Temp\Expert-Analysis-v1.xlsx'

.LINK
https://github.com/Azure/Azure-Proactive-Resiliency-Library-v2
#>

Param(
[ValidatePattern('^https:\/\/.+$')]
[string] $RecommendationDataUri = 'https://azure.github.io/WARA-Build/objects/recommendations.json',
[string] $CustomRecommendationObject,
[Parameter(mandatory = $true)]
[string] $JSONFile,
[string] $ExpertAnalysisFile
)

# WARA In Scope Resource Types CSV File
$RecommendationResourceTypesUri = 'https://azure.github.io/WARA-Build/objects/WARAinScopeResTypes.csv'

# Check if the Expert-Analysis file exists
$ExpertAnalysisPath = $PSScriptRoot + '\Expert-Analysis-v1.xlsx'

if (!$ExpertAnalysisFile)
	{
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + (' - Testing: ' + './Expert-Analysis-v1.xlsx'))
		if ((Test-Path -Path ($ExpertAnalysisPath) -PathType Leaf) -eq $true) {
			$ExpertAnalysisFile = $ExpertAnalysisPath
		}
	}
else
{
	Throw "Error locating the Expert-Analysis file. Please provide a valid path to the Expert-Analysis file or reinstall the WARA Module."
	Exit
}

if ((Test-Path -Path $ExpertAnalysisFile -PathType Leaf) -eq $true) {
	$ExpertAnalysisFile = (Resolve-Path -Path $ExpertAnalysisFile).Path
}
else
{
	Throw "The Expert-Analysis file does not exist. Please provide a valid path to the Expert-Analysis file."
	Exit
}


# Check if the JSON file exists
if ((Test-Path -Path $JSONFile -PathType Leaf) -eq $true) {
	$JSONFile = (Resolve-Path -Path $JSONFile).Path
}
else
{
	Throw "JSON file not found. Please provide a valid path to the JSON file."
	Exit
}

$TableStyle = 'Light19'


# Classes
Class ImpactedResourceObj {
    [string] $ValidationMSG
    [string] $ValidationCategory
    [string] $ResourceType
    [string] $SubscriptionId
    [string] $ResourceGroup
    [string] $Location
    [string] $Name
    [string] $Id
    [string] $Custom1
    [string] $Custom2
    [string] $Custom3
    [string] $Custom4
    [string] $Custom5
    [string] $RecommendationTitle
    [string] $Impact
    [string] $RecommendationControl
    [string] $PotentialBenefit
    [string] $LearnMoreLink
    [string] $LongDescription
    [string] $Guid
    [string] $Category
    [string] $Source
    [string] $WAFPillar
    [string] $PlatformIssueTrackingId
    [string] $RetirementTrackingId
    [string] $SupportRequestNumber
    [string] $Notes
    [string] $CheckName
}

Class AnalysisPlanningObj {
    [string] $Category
    [string] $ResourceType = 'N/A'
    [string] $NumberOfResources = 'n/a'
    [string] $ImpactedResources = 'n/a'
    [string] $HasRecommendationsInAPRLAdvisor = 'Yes'
    [string] $AssessmentStatus = 'Pending'
}

Class OutagesObj {
    [string] $OutageMSG
    [string] $TrackingID
    [string] $EventType
    [string] $EventSource
    [string] $Status
    [string] $Title
    [string] $Level
    [string] $EventLevel
    [string] $StartTime
    [string] $MitigationTime
    [string] $ImpactedService
    [string] $WhatHappened
    [string] $WhatWentWrongAndWhy
    [string] $HowDidWeRespond
    [string] $HowAreWeMakingIncidentsLessLikely
    [string] $HowCanCustomersMakeIncidentsLessImpactful
}

Class WorkLoadInvObj {
    [string] $subscriptionId
    [string] $resourceGroup
    [string] $type
    [string] $location
    [string] $name
    [string] $id
    [string] $tenantId
    [string] $kind
    [string] $managedBy
    [string] $sku
    [string] $plan
    [string] $zones
}

# function validate if the required modules are installed
function Test-Requirement {
	# Install required modules
	Write-Host 'Validating ' -NoNewline
	Write-Host 'ImportExcel' -ForegroundColor Cyan -NoNewline
	Write-Host ' Module..'
	$ImportExcel = Get-Module -Name ImportExcel -ListAvailable -ErrorAction silentlycontinue
	if ($null -eq $ImportExcel) {
		Write-Host 'Installing ImportExcel Module' -ForegroundColor Yellow
		Install-Module -Name ImportExcel -Force -SkipPublisherCheck
	}
<# 	Write-Host 'Validating ' -NoNewline
	Write-Host 'Powershell-YAML' -ForegroundColor Cyan -NoNewline
	Write-Host ' Module..'
	$AzModules = Get-Module -Name powershell-yaml -ListAvailable -ErrorAction silentlycontinue
	if ($null -eq $AzModules) {
		Write-Host 'Installing Az Modules' -ForegroundColor Yellow
		Install-Module -Name powershell-yaml -SkipPublisherCheck -InformationAction SilentlyContinue
	} #>
<# 	Write-Host 'Validating ' -NoNewline
	Write-Host 'Git' -ForegroundColor Cyan -NoNewline
	Write-Host ' Installation..'
	$GitVersion = git --version
	if ($null -eq $GitVersion) {
		Write-Host 'Missing Git' -ForegroundColor Red
		Exit
	} #>
}

# function to read the JSON file
function Read-JSONFile
{
	Param
	(
		[Parameter(Mandatory = $true)]
		[string]$JSONFile
	)

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Starting to Read the JSON file')
	$JSONResources = Get-Item -Path $JSONFile
	$JSONResources = $JSONResources.FullName
	$JSONContent = Get-Content -Path $JSONResources | ConvertFrom-Json
	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - JSON File Created with version: ' + $JSONContent.ScriptDetails.Version)
	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Raw ImpactedResources found: ' + $JSONContent.ImpactedResources.Count)
	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Raw PlatformIssues found: ' + $JSONContent.Outages.Count)
	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Raw SupportTickets found: ' + $JSONContent.SupportTickets.Count)
	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Raw Workload Inventory found: ' + $JSONContent.InScopeResources.Count)
	return $JSONContent
}

function Save-WARAExcelFile
{
	Param(
		$ExcelPackage
		)

	$workingFolderPath = Get-Location
	$workingFolderPath = $workingFolderPath.Path
	$NewExpertAnalysisFile = ($workingFolderPath + '\Expert-Analysis-v1-' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + '.xlsx')
	Close-ExcelPackage -ExcelPackage $ExcelPackage -SaveAs $NewExpertAnalysisFile

	return $NewExpertAnalysisFile
}

# function responsible to import recommendations from the JSON file
function Get-WARARecommendationList
{
    Param(
        [Parameter(Mandatory = $true)]
        [string]$RecommendationDataUri
    )
	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Processing Recommendations from JSON file.')
	# Get Recommendation Objects
    $RecommendationObject = Invoke-RestMethod $RecommendationDataUri

	return $RecommendationObject

}

# function to process the standard WARA message in the column A
function Get-WARAMessage
{
	Param($Message)

if ($Message -eq 'ImpactedResources_Type')
{
$WARAMessage = @"
REQUIRED ACTIONS: ResourceType not available in APRL/Advisor.
This Resource Type does not have recommendations in APRL or Advisor, then follow these steps:
1. Manually validate this resource and create your own resiliency and/or reliability-related recommendations, if applicable.
2. When creating a new recommendation, ensure the following fields are populated: Title, Impact, Potential Benefit, Description, ResourceType, and Learn More Link.
3. Duplicate this row if you need to create more than one recommendation for the same resource.
4. If the resource is not compliant with your recommendation,  update this cell to "Reviewed".
or
5. Delete this row if this resource is irrelevant or is already compliant with your recommendation.
"@
}

if ($Message -eq 'ImpactedResources_Unavailable')
{
$WARAMessage = @"
REQUIRED ACTIONS: Recommendation does not have automated validation.
This recommendation does not have automated validation. Since it cannot be validated automatically, follow these steps:
1. Review the Recommendation Title and Description to understand the recommendation. Use the "Read More" links for additional information.
2. Open the Azure Portal and locate the potentially impacted resource using its Resource Name or ResourceId.
3. Manually validate the resource.
4. If the resource is not compliant with the recommendation, update this cell to "Reviewed".
or
5. Delete this row if the resource is irrelevant or already compliant with the recommendation.
"@
}

if ($Message -eq 'ImpactedResources_ServiceRetirement')
{
$WARAMessage = @"
REQUIRED ACTIONS: Azure Service Health - Service Retirements.
This Service Health Retirement Notification was automatically imported to your workload review according to the Subscriptions being assessed and Services being used.
1. Review and if necessary summarize the cell "Recommendation Title."
2. Retrieve the Resource Name and Resource ID of the impacted resources from the Azure Portal.
3. If more resources are associated with the same Service Retirement Notification, duplicate this row and update it with the Resource Name and Resource ID for all applicable resources.
4. Once completed, update this cell to "Reviewed".
or
5. If this Service Retirement Notification is not relevant, delete this row.
"@
}

if ($Message -eq 'ImpactedResources_Architecture')
{
$WARAMessage = @"
REQUIRED ACTIONS: Architectural and Reliability Design Patterns Recommendations
This row is just an example. Modify it and create your own personalized recommendations for architecture and/or reliability design patterns.
1. Create new custom recommendations based on the Discovery Workshop Questionnaire.
2. Associate it with the SubscriptionId and "Microsoft.Subscription/Subscriptions" as the associated "resourceType", or any resource and ResourceType that you consider applicable.
3. Once completed, update this cell to "Reviewed".
or
4. Delete this row if no personalized/custom recommendations will be provided.
"@
}

if ($Message -eq 'ImpactedResources_WAF')
{
$WARAMessage = @"
REQUIRED ACTIONS: Well-Architected Framework.
This is a generic recommendation from the Well-Architected Framework - Reliability pillar. It cannot be validated automatically and need information from the Discovery Workshop session.
1. Review and update this row based on the Discovery Workshop Questionnaire.
2. If this recommendation is applicable, update the "Impact" column according to the importance for the Workload and then update this cell to "Reviewed".
or
3. If this recommendation is not relevant, delete this row.
"@
}

if ($Message -eq 'PlatformIssues_Standard')
{
$WARAMessage = @"
REQUIRED ACTIONS: Review Platform Issue and create recommendations.
This Platform Issue may have affected the workload, since it cannot be validated automatically, follow these steps:
1. Review all information about the Platform Issue.
2. Discuss with the Account Team/Workload Owner if this Platform Issue affected the workload and how, you can also check if there are associated Support Requests.
3. If this issue affected the workload, create recommendations in the "4.ImpactedResourcesAnalysis" worksheet based on the "How can customers make incidents like this less impactful" field. You can associate the recommendation(s) with the Workload or individual resources.
4. If the recommendation(s) already exists, simply add the TrackingID to the respective column of the associated Recommendation(s).
or
5. Delete this row if the workload was not affected by this Platform Issue.
"@
}

if ($Message -eq 'SupportTickets_Standard')
{
$WARAMessage = @"
REQUIRED ACTIONS: Review Customer Support Requests and create recommendations.
This Customer Support Request is associated with the Subscription of the workload, since it cannot be validated automatically, follow these steps:
1. Review all information about the Support Request in the Azure Portal.
2. Discuss with Workload Owner if this Support Request affected the workload and how.
3. If this Support Request was relevant for the workload, create recommendations in the "4.ImpactedResourcesAnalysis" worksheet based on it was resolved. You can associate the recommendation(s) with the Workload or individual resources. The goal is to make sure other resources or services follow the Microsoft recommendation/solution and prevent this incident from happening again.
4. If the recommendation(s) already exists, simply add the TicketID to the respective column of the associated Recommendation(s).
or
5. Delete this row if the workload was not affected, associated with this Support Request.
"@
}

return $WARAMessage
}

<############################## Impacted Resources #########################################>

function Initialize-WARAImpactedResources
{
	Param(
		[Parameter(mandatory = $true)]
        [AllowEmptyCollection()]
		$ImpactedResources,
		[Parameter(mandatory = $false)]
        [AllowEmptyCollection()]
		$Advisory,
		[Parameter(mandatory = $false)]
        [AllowEmptyCollection()]
		$Retirements,
		[Parameter(mandatory = $true)]
        [AllowEmptyCollection()]
		$ScriptDetails,
        [Parameter(mandatory = $true)]
        [AllowEmptyCollection()]
        $RecommendationDataUri
	)

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Getting All Recommendations')
	$RecommendationObject = Get-WARARecommendationList -RecommendationDataUri $RecommendationDataUri
    Write-Debug "Count of Recommendations: $($RecommendationObject.Count)"


    if($CustomRecommendationObject) {
        Write-Debug "Adding Custom Recommendations"
        $CustomRecommendationObject = Get-Content $CustomRecommendationObject -raw | ConvertFrom-Json -depth 20
        Write-Debug "Count of Custom Recommendations: $($CustomRecommendationObject.Count)"
        $RecommendationObject += $CustomRecommendationObject
        Write-Debug "Count of Recommendations after adding Custom Recommendations: $($RecommendationObject.Count)"
    }

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Getting Recommendations from the standard Resources')
    $ResourceRecommendations = $RecommendationObject | Where-Object {[string]::IsNullOrEmpty($_.tags)}

    Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Getting Recommendations from the standard Azure-WAF')
    $WAFRecommendations = $RecommendationObject | Where-Object { $_.tags -like 'WAF'}

	if ($ScriptDetails.SAP -eq 'True') {
		Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Getting Recommendations from SAP')
		$ResourceRecommendations += $RecommendationObject | Where-Object {$_.tags -like 'SAP'}
	}
	if ($ScriptDetails.AVD -eq 'True') {
		Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Getting Recommendations from AVD')
		$ResourceRecommendations += $RecommendationObject | Where-Object {$_.tags -like 'AVD'}
	}
	if ($ScriptDetails.AVS -eq 'True') {
		Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Getting Recommendations from AVS')
		$ResourceRecommendations += $RecommendationObject | Where-Object {$_.tags -like 'AVS'}
	}
	if ($ScriptDetails.HPC -eq 'True') {
		Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Getting Recommendations from HPC')
		$ResourceRecommendations += $RecommendationObject | Where-Object {$_.tags -like 'HPC'}
	}
    if ($ScriptDetails.AI_GPT_RAG -eq 'True') {
		Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Getting Recommendations from AI-GPT-RAG')
		$ResourceRecommendations += $RecommendationObject | Where-Object {$_.tags -like 'AI-GPT-RAG'}
	}
    if ($ScriptDetails.ORACLE -eq 'True') {
        Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Getting Recommendations from ORACLE')
        $ResourceRecommendations += $RecommendationObject | Where-Object {$_.tags -like 'ORACLE'}
    }

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Overall Recommendations found: ' + [string]$ResourceRecommendations.Count)

	# Filtering the recommendations to get only the active ones and the ones that are not already in the advisories list
	#Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Filtering Active Recommendations and Recommendations not in Advisories')
	#$RecommendationContent = $ResourceRecommendations | Where-Object {($_.recommendationMetadataState -eq 'Active' -and $_.recommendationTypeId -notin $JSONContent.Advisory.recommendationId) }

	#Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Recommendations After Filtering: ' + [string]$RecommendationContent.Count)

	$tmp = @()

	# First loop through the recommendations to get the impacted resources
	foreach ($Recom in $ResourceRecommendations)
		{
			# Getting the impacted resources for the recommendation and validating if the recommendation is a Custom Recommendation
			$Resources = $ImpactedResources | Where-Object {($_.recommendationId -eq $Recom.aprlGuid) -and ($_.checkName -eq $Recom.checkName) }

			# If the recommendation is not a Custom Recommendation, we need to validate if the resources are not already in the tmp array (from a previous loop of a Custom Recommendation)
			if ([string]::IsNullOrEmpty($Resources) -and $Recom.aprlGuid -notin $tmp.Guid -and -not $Recom.checkName)
			{
				$Resources = $ImpactedResources | Where-Object {($_.recommendationId -eq $Recom.aprlGuid) }
			}

			foreach ($Resource in $Resources)
			{
                $ValidationMSG = switch ($Resource.validationAction) {
                    'IMPORTANT - Query under development - Validate Resources manually' {
                        Get-WARAMessage -Message 'ImpactedResources_Unavailable'

                    }
                    'IMPORTANT - Recommendation cannot be validated with ARGs - Validate Resources manually' {
                        Get-WARAMessage -Message 'ImpactedResources_Unavailable'

                    }
                    'IMPORTANT - Resource Type is not available in either APRL or Advisor - Validate Resources manually if Applicable, if not Delete this line' {
                        Get-WARAMessage -Message 'ImpactedResources_Type'

                    }
                    'APRL - Queries' {
                        'Reviewed'

                    }
                    default {
                        'Error'
                    }
                }

                $ResObj = [ImpactedResourceObj]::new()
                $ResObj.ValidationMSG = $ValidationMSG
                $ResObj.ValidationCategory = 'Resource'
                $ResObj.ResourceType = $Resource.type
                $ResObj.SubscriptionId = $Resource.subscriptionId
                $ResObj.ResourceGroup = $Resource.resourceGroup
                $ResObj.Location = $Resource.location
                $ResObj.Name = $Resource.name
                $ResObj.Id = $Resource.id
                $ResObj.Custom1 = $Resource.param1
                $ResObj.Custom2 = $Resource.param2
                $ResObj.Custom3 = $Resource.param3
                $ResObj.Custom4 = $Resource.param4
                $ResObj.Custom5 = $Resource.param5
                $ResObj.RecommendationTitle = $Recom.description
                $ResObj.Impact = $Recom.recommendationImpact
                $ResObj.RecommendationControl = ($Recom.recommendationControl -csplit '(?=[A-Z])' -ne '' -join ' ')
                $ResObj.PotentialBenefit = $Recom.potentialBenefits
                $ResObj.LearnMoreLink = ($Recom.learnMoreLink.url -join " `n")
                $ResObj.LongDescription = $Recom.longDescription
                $ResObj.Guid = $Recom.aprlGuid
                $ResObj.Category = 'Azure Service'
                $ResObj.Source = $Resource.selector
                $ResObj.WAFPillar = 'Reliability'
                $ResObj.CheckName = $Resource.checkName

				$tmp += $ResObj
			}
		}

    # Second loop through the resources without GUID
    $Resources = $ImpactedResources | Where-Object {[string]::IsNullOrEmpty($_.recommendationId)}
    foreach ($Resource in $Resources)
    {
        $ValidationMSG = switch ($Resource.validationAction) {
            'IMPORTANT - Resource Type is not available in either APRL or Advisor - Validate Resources manually if Applicable, if not Delete this line' {
                Get-WARAMessage -Message 'ImpactedResources_Type'

            }
            default {
                'Custom'
            }
        }

        $ResObj = [ImpactedResourceObj]::new()
        $ResObj.ValidationMSG = $ValidationMSG
        $ResObj.ValidationCategory = 'Resource'
        $ResObj.ResourceType = $Resource.type
        $ResObj.SubscriptionId = $Resource.subscriptionId
        $ResObj.ResourceGroup = $Resource.resourceGroup
        $ResObj.Location = $Resource.location
        $ResObj.Name = $Resource.name
        $ResObj.Id = $Resource.id
        $ResObj.RecommendationControl = 'Other Best Practices'
        $ResObj.Impact = 'Low'
        $ResObj.Category = 'Azure Service'
        $ResObj.Source = $Resource.selector
        $ResObj.WAFPillar = 'Reliability'
        $ResObj.CheckName = $Resource.checkName

        $tmp += $ResObj
    }

	# Third loop through the Advisories
	$ADVMessage = "Reviewed"
	foreach ($adv in $Advisory)
		{
			if (![string]::IsNullOrEmpty($adv))
				{
                    $ADVobj = [impactedResourceObj]::new()
                    $ADVobj.ValidationMSG = $ADVMessage
                    $ADVobj.ValidationCategory = 'Resource'
                    $ADVobj.ResourceType = $adv.type
                    $ADVobj.SubscriptionId = $adv.subscriptionId
                    $ADVobj.ResourceGroup = $adv.resourceGroup
                    $ADVobj.Location = $adv.location
                    $ADVobj.Name = $adv.name
                    $ADVobj.Id = $adv.id
                    $ADVobj.RecommendationTitle = $adv.description
                    $ADVobj.Impact = $adv.impact
                    $ADVobj.RecommendationControl = ($adv.category -csplit '(?=[A-Z])' -ne '' -join ' ')
                    $ADVobj.Guid = $adv.recommendationId
                    $ADVobj.Category = "Azure Service"
                    $ADVobj.Source = 'ADVISOR'
                    $ADVobj.WAFPillar = 'Reliability'

                    $tmp += $ADVobj
				}
		}

	# Fourth loop through the Service Retirements
	$ServiceRetirementMSG = Get-WARAMessage -Message 'ImpactedResources_ServiceRetirement'
	foreach ($Retirement in $Retirements)
		{
			if (![string]::IsNullOrEmpty($Retirement)) {
				$RetirementType = $RootTypes | Where-Object {$_.FriendlyName -eq $Retirement.ImpactedService}
				$RetirementType = if(![string]::IsNullOrEmpty($RetirementType)) { $RetirementType.ResourceType } else { $Retirement.ImpactedService }

				try {
					$HTML = New-Object -Com 'HTMLFile'
					$HTML.write([ref]$Retirement.Description)
					$RetirementDescriptionFull = $Html.body.innerText
					$SplitDescription = $RetirementDescriptionFull.split('Help and support').split('Required action')
				} catch {
					$SplitDescription = ' ', ' '
				}

                $RetObj = [ImpactedResourceObj]::new()
                $RetObj.ValidationMSG = $ServiceRetirementMSG
                $RetObj.ValidationCategory = 'Retirements'
                $RetObj.ResourceType = $RetirementType
                $RetObj.SubscriptionId = $Retirement.Subscription
                $RetObj.ResourceGroup = 'RG name not needed'
                $RetObj.Location = 'Location not needed'
                $RetObj.Name = 'Get ResourceName from Azure Portal'
                $RetObj.Id = 'Get ResourceID from Azure Portal'
                $RetObj.RecommendationTitle = $Retirement.Title
                $RetObj.Impact = 'Medium'
                $RetObj.LongDescription = [string]$SplitDescription[0]
                $RetObj.Guid = 'GUID not needed'
                $RetObj.Source = 'Azure Service Health - Service Retirements'
                $RetObj.WAFPillar = 'Reliability'
                $RetObj.RetirementTrackingId = $Retirement.TrackingId

				$tmp += $RetObj
			}
		}

	# Fifth loop through the WAF Recommendations
	$WAFMSG = Get-WARAMessage -Message 'ImpactedResources_WAF'
	foreach ($waf in $WAFRecommendations)
		{
			if (![string]::IsNullOrEmpty($waf))
				{
                    $WAFObj = [ImpactedResourceObj]::new()
                    $WAFObj.ValidationMSG = $WAFMSG
                    $WAFObj.ValidationCategory = 'WAF'
                    $WAFObj.ResourceType = $waf.recommendationResourceType
                    $WAFObj.Name = 'Entire Workload'
                    $WAFObj.RecommendationTitle = $waf.description
                    $WAFObj.Impact = $waf.recommendationImpact
                    $WAFObj.RecommendationControl = ($waf.recommendationControl -csplit '(?=[A-Z])' -ne '' -join ' ')
                    $WAFObj.PotentialBenefit = $waf.potentialBenefits
                    $WAFObj.LearnMoreLink = ($waf.learnMoreLink.url -join " `n")
                    $WAFObj.LongDescription = $waf.longDescription
                    $WAFObj.Guid = $waf.aprlGuid
                    $WAFObj.Category = 'Well Architected'

					$tmp += $WAFObj
				}
		}

	# Standard Architecture and Reliability Design Patterns Recommendations
	$ArchtectureMSG = Get-WARAMessage -Message 'ImpactedResources_Architecture'

    $ARCHObj = [ImpactedResourceObj]::new()
    $ARCHObj.ValidationMSG = $ArchtectureMSG
    $ARCHObj.ValidationCategory = 'Architectural'
    $ARCHObj.ResourceType = 'Microsoft.Subscription/Subscriptions'
    $ARCHObj.Impact = 'Low'
    $ARCHObj.RecommendationControl = 'Governance'

    $tmp += $ARCHObj

    $ImpactedResourcesFormatted = foreach ($line in $tmp)
        {
            [PSCustomObject]@{
                'REQUIRED ACTIONS / REVIEW STATUS' = $line.ValidationMSG
                'ValidationCategory' = $line.ValidationCategory
                'Resource Type' = $line.ResourceType
                'subscriptionId' = $line.SubscriptionId
                'resourceGroup' = $line.ResourceGroup
                'location' = $line.Location
                'name' = $line.Name
                'id' = $line.Id
                'custom1' = $line.Custom1
                'custom2' = $line.Custom2
                'custom3' = $line.Custom3
                'custom4' = $line.Custom4
                'custom5' = $line.Custom5
                'Recommendation Title' = $line.RecommendationTitle
                'Impact' = $line.Impact
                'Recommendation Control' = $line.RecommendationControl
                'Potential Benefit' = $line.PotentialBenefit
                'Learn More Link' = $line.LearnMoreLink
                'Long Description' = $line.LongDescription
                'Guid' = $line.Guid
                'Category' = $line.Category
                'Source' = $line.Source
                'WAF Pillar' = $line.WAFPillar
                'Platform Issue TrackingId' = $line.PlatformIssueTrackingId
                'Retirement TrackingId' = $line.RetirementTrackingId
                'Support Request Number' = $line.SupportRequestNumber
                'Notes' = $line.Notes
                'checkName' = $line.CheckName
            }
        }

	# Returns the array with all the recommendations already formatted to be exported to Excel
	return $ImpactedResourcesFormatted
}

function Export-WARAImpactedResources
{
	Param($ImpactedResourcesFormatted,$ExcelPackage)

	$Style = @()
	$Style += New-ExcelStyle -HorizontalAlignment Left -WrapText -Range A:A
	$Style += New-ExcelStyle -HorizontalAlignment Center -Range A12:A12
	$Style += New-ExcelStyle -HorizontalAlignment Center -Range B:F
	$Style += New-ExcelStyle -HorizontalAlignment Left -Range G:L
	$Style += New-ExcelStyle -HorizontalAlignment Center -Range G12:L12
	$Style += New-ExcelStyle -HorizontalAlignment Center -WrapText -Range M:P
	$Style += New-ExcelStyle -HorizontalAlignment Center -WrapText -Range R:R
	$Style += New-ExcelStyle -HorizontalAlignment Center -Range S:Z
	$Style += New-ExcelStyle -VerticalAlignment Center -Range A:Z

	$ImpactedResourcesSheet = New-Object System.Collections.Generic.List[System.Object]
	$ImpactedResourcesSheet.Add('REQUIRED ACTIONS / REVIEW STATUS')
	$ImpactedResourcesSheet.Add('ValidationCategory')
	$ImpactedResourcesSheet.Add('Resource Type')
	$ImpactedResourcesSheet.Add('subscriptionId')
	$ImpactedResourcesSheet.Add('resourceGroup')
	$ImpactedResourcesSheet.Add('location')
	$ImpactedResourcesSheet.Add('name')
	$ImpactedResourcesSheet.Add('id')
	$ImpactedResourcesSheet.Add('custom1')
	$ImpactedResourcesSheet.Add('custom2')
	$ImpactedResourcesSheet.Add('custom3')
	$ImpactedResourcesSheet.Add('custom4')
	$ImpactedResourcesSheet.Add('custom5')
	$ImpactedResourcesSheet.Add('Recommendation Title')
	$ImpactedResourcesSheet.Add('Impact')
	$ImpactedResourcesSheet.Add('Recommendation Control')
	$ImpactedResourcesSheet.Add('Potential Benefit')
	$ImpactedResourcesSheet.Add('Learn More Link')
	$ImpactedResourcesSheet.Add('Long Description')
	$ImpactedResourcesSheet.Add('Guid')
	$ImpactedResourcesSheet.Add('Category')
	$ImpactedResourcesSheet.Add('Source')
	$ImpactedResourcesSheet.Add('WAF Pillar')
	$ImpactedResourcesSheet.Add('Platform Issue TrackingId')
	$ImpactedResourcesSheet.Add('Retirement TrackingId')
	$ImpactedResourcesSheet.Add('Support Request Number')
	$ImpactedResourcesSheet.Add('Notes')
	$ImpactedResourcesSheet.Add('checkName')



	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Exporting Impacted Resources to Excel')




    Add-ExcelDataValidationRule -Worksheet $excelPackage.$ImpactedResourcesSheetRef -Range "A13:A1048576" -ValidationType List -ValueSet @('Pending','Reviewed') -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Invalid Entry' -ErrorBody 'Please enter a valid value (Pending or Reviewed)' -NoBlank $true
    Add-ExcelDataValidationRule -Worksheet $excelPackage.$ImpactedResourcesSheetRef -Range "O13:O1048576" -ValidationType List -ValueSet @('High','Medium','Low') -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Invalid Entry' -ErrorBody 'Please enter a valid value (High, Medium, or Low)' -NoBlank $true
    Add-ExcelDataValidationRule -Worksheet $ExcelPackage.$ImpactedResourcesSheetRef -Range "G13:G1048576" -ValidationType TextLength -Operator greaterThan -Value 1 -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Invalid Entry' -ErrorBody 'Please enter a valid value (more than 1 character)' -NoBlank $true


    $null = $ImpactedResourcesFormatted | ForEach-Object { [PSCustomObject]$_ } | Select-Object $ImpactedResourcesSheet |
    Export-Excel -ExcelPackage $excelPackage -WorksheetName $ImpactedResourcesSheetRef -TableName 'impactedresources' -TableStyle $TableStyle -Style $Style -StartRow 12 -PassThru




    #Close-ExcelPackage $excelPackage
}

<############################## Analysis Planning #########################################>

function Initialize-WARAAnalysisPlanning
{
	Param(
		$InScopeResources
	)

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Grouping InScope Resources by Resource Type')
	$ResourceTypes = $InScopeResources | Group-Object -Property type

# Formula has to be extracted from the Excel file to be in this specific standard used by Excel, otherwise it will not work
$InventoryFormula = @"
=COUNTA(_xlfn.UNIQUE(_xlfn.VSTACK(_xlfn._xlws.FILTER($WorkloadInventorySheetRef!A:A, $WorkloadInventorySheetRef!C:C = TableTypes8[[#This Row],[Resource Type]]), _xlfn._xlws.FILTER($WorkloadInventorySheetRef!A:A, $WorkloadInventorySheetRef!C:C = TableTypes8[[#This Row],[Resource Type]]))))
"@

$ImpactedResourcesFormula = @"
=IF(COUNTIF($ImpactedResourcesSheetRef!C:C, TableTypes8[[#This Row],[Resource Type]])=0, 0, COUNTA(_xlfn.UNIQUE(_xlfn._xlws.FILTER($ImpactedResourcesSheetRef!H:H, ($ImpactedResourcesSheetRef!C:C=TableTypes8[[#This Row],[Resource Type]]) * ($ImpactedResourcesSheetRef!H:H<>"Get ResourceID from Azure Portal")))))
"@

$ReviewedFormula = @"
=IF(OR(AND(TableTypes8[[#This Row],[Category]]="Support Requests", COUNTIFS($SupportRequestsSheetRef!A:A, "<>Reviewed")=0),AND(TableTypes8[[#This Row],[Category]]="Platform Issues", COUNTIFS($PlatformIssuesSheetRef!A:A,
"<>Reviewed")=0),AND(TableTypes8[[#This Row],[Category]]="Impacted Resources", COUNTIFS($ImpactedResourcesSheetRef!A:A, "<>Reviewed", $ImpactedResourcesSheetRef!C:C, TableTypes8[[#This Row],[Resource Type]], $ImpactedResourcesSheetRef!O:O,
"<>Low")=0)), "Reviewed", "Pending")
"@

	$tmp = @()
	foreach ($ResourceType in $ResourceTypes)
		{
			$RootType = ""
			$RootType = $RootTypes | Where-Object {$_.ResourceType -eq $ResourceType.Name}
			$APRLOrAdv = if($RootType.WARAinScope -eq 'yes' -and $RootType.InAprlAndOrAdvisor -eq 'yes') { 'Yes' } else { 'No' }

            $ResTypeObj = [AnalysisPlanningObj]::new()
            $ResTypeObj.Category = 'Impacted Resources'
            $ResTypeObj.ResourceType = $ResourceType.Name
            $ResTypeObj.NumberOfResources = $InventoryFormula
            $ResTypeObj.ImpactedResources = $ImpactedResourcesFormula
            $ResTypeObj.HasRecommendationsInAPRLAdvisor = $APRLOrAdv
            $ResTypeObj.AssessmentStatus = $ReviewedFormula

			$tmp += $ResTypeObj
		}

    $SupObj = [AnalysisPlanningObj]::new()
    $SupObj.Category = 'Support Requests'
	$tmp += $SupObj

    $PlatObj = [AnalysisPlanningObj]::new()
    $PlatObj.Category = 'Platform Issues'
	$tmp += $PlatObj


    $AnalysisPlanningFormatted = foreach ($line in $tmp)
        {
            [PSCustomObject]@{
                'Category' = $line.Category
                'Resource Type' = $line.ResourceType
                'Number of Resources' = $line.NumberOfResources
                'Impacted Resources' = $line.ImpactedResources
                'Has Recommendations_x000a_in APRL/Advisor' = $line.HasRecommendationsInAPRLAdvisor
                'Assessment Status' = $line.AssessmentStatus
            }
        }

	return $AnalysisPlanningFormatted

}

function Export-WARAAnalysisPlanning
{
	Param($AnalysisPlanningFormatted, $ExcelPackage)

	$Style = @()
	$Style += New-ExcelStyle -HorizontalAlignment Center -Range A:H


	$AnalysesPlanning = New-Object System.Collections.Generic.List[System.Object]
	$AnalysesPlanning.Add('Category')
	$AnalysesPlanning.Add('Resource Type')
	$AnalysesPlanning.Add('Number of Resources')
	$AnalysesPlanning.Add('Impacted Resources')
	$AnalysesPlanning.Add('Has Recommendations_x000a_in APRL/Advisor')
	$AnalysesPlanning.Add('Assessment Owner')
	$AnalysesPlanning.Add('Assessment Status')
	$AnalysesPlanning.Add('Notes')


	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Exporting Analysis Planning to Excel')
	$null = $AnalysisPlanningFormatted | ForEach-Object { [PSCustomObject]$_ } | Select-Object $AnalysesPlanning |
		Export-Excel -ExcelPackage $ExcelPackage -WorksheetName $AnalysisPlanningSheetRef -TableName 'TableTypes8' -TableStyle $TableStyle -Style $Style -StartRow 10 -PassThru
}

<############################## Platform Issues #########################################>

function Initialize-WARAPlatformIssues
{
	Param(
		$PlatformIssues
	)

	$OutagesMSG = Get-WARAMessage -Message 'PlatformIssues_Standard'

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Formatting Platform Issues')
	# Just getting a sum of the total outages that will be included in the Excel file, in case there are no outages, a row will be added with empty values later
	$TotalOutages = ($PlatformIssues | Where-Object {$_.properties.description -like '*How can customers make incidents like this less impactful?*'}).count

	$tmp = @()
	foreach ($Outage in $PlatformIssues)
		{
			if ($Outage.properties.description -like '*How can customers make incidents like this less impactful?*') {
				try {
					$HTML = New-Object -Com 'HTMLFile'
					$HTML.write([ref]$Outage.properties.description)
					$OutageDescription = $Html.body.innerText
					$SplitDescription = $OutageDescription.split('How can we make our incident communications more useful?').split('How can customers make incidents like this less impactful?').split('How are we making incidents like this less likely or less impactful?').split('How did we respond?').split('What went wrong and why?').split('What happened?')
					$whathap = ($SplitDescription[1]).Split([Environment]::NewLine)[1]
					$whatwent = ($SplitDescription[2]).Split([Environment]::NewLine)[1]
					$howdid = ($SplitDescription[3]).Split([Environment]::NewLine)[1]
					$howarewe = ($SplitDescription[4]).Split([Environment]::NewLine)[1]
					$howcan = ($SplitDescription[5]).Split([Environment]::NewLine)[1]
				}
				catch {
					$whathap = ""
					$whatwent = ""
					$howdid = ""
					$howarewe = ""
					$howcan = ""
				}

				$ImpactedSvc = if ($Outage.properties.impact.impactedService.count -gt 1) { $Outage.properties.impact.impactedService | ForEach-Object { $_ + ' ,' } }else { $Outage.properties.impact.impactedService}
				$ImpactedSvc = [string]$ImpactedSvc
				$ImpactedSvc = if ($ImpactedSvc -like '* ,*') { $ImpactedSvc -replace ".$" }else { $ImpactedSvc }

                $OutageObj = [OutagesObj]::new()
                $OutageObj.OutageMSG = $OutagesMSG
                $OutageObj.TrackingID = $Outage.name
                $OutageObj.EventType = $Outage.properties.eventType
                $OutageObj.EventSource = $Outage.properties.eventSource
                $OutageObj.Status = $Outage.properties.status
                $OutageObj.Title = $Outage.properties.title
                $OutageObj.Level = $Outage.properties.level
                $OutageObj.EventLevel = $Outage.properties.eventLevel
                $OutageObj.StartTime = $Outage.properties.impactStartTime
                $OutageObj.MitigationTime = $Outage.properties.impactMitigationTime
                $OutageObj.ImpactedService = $ImpactedSvc
                $OutageObj.WhatHappened = $whathap
                $OutageObj.WhatWentWrongAndWhy = $whatwent
                $OutageObj.HowDidWeRespond = $howdid
                $OutageObj.HowAreWeMakingIncidentsLessLikely = $howarewe
                $OutageObj.HowCanCustomersMakeIncidentsLessImpactful = $howcan

                $tmp += $OutageObj
			}
		}

	# If there are no outages, a row will be added with empty values
	if ($TotalOutages -eq 0) {
        $ZeroOutageObj = [OutagesObj]::new()
		$ZeroOutageObj.OutageMSG = $OutagesMSG
		$tmp += $ZeroOutageObj
	}

    $PlatformIssuesFormatted = foreach ($line in $tmp)
        {
            [PSCustomObject]@{
                'REQUIRED ACTIONS / REVIEW STATUS'	 									= $line.OutageMSG
                'Tracking ID' 															= $line.TrackingID
                'Event Type' 															= $line.EventType
                'Event Source' 															= $line.EventSource
                'Status' 																= $line.Status
                'Title' 																= $line.Title
                'Level' 																= $line.Level
                'Event Level' 															= $line.EventLevel
                'Start Time' 															= $line.StartTime
                'Mitigation Time'														= $line.MitigationTime
                'Impacted Service'														= $line.ImpactedService
                'What happened' 														= $line.WhatHappened
                'What went wrong and why' 												= $line.WhatWentWrongAndWhy
                'How did we respond' 													= $line.HowDidWeRespond
                'How are we making incidents like this less likely or less impactful' 	= $line.HowAreWeMakingIncidentsLessLikely
                'How can customers make incidents like this less impactful' 			= $line.HowCanCustomersMakeIncidentsLessImpactful
            }
            $obj
        }

	return $PlatformIssuesFormatted
}

function Export-WARAPlatformIssues
{
	Param($PlatformIssuesFormatted, $excelPackage)

	$Style = @()
	$Style += New-ExcelStyle -HorizontalAlignment Left -WrapText -Range A:A
	$Style += New-ExcelStyle -HorizontalAlignment Center -Range A12:A12
	$Style += New-ExcelStyle -HorizontalAlignment Center -Range B:K
	$Style += New-ExcelStyle -HorizontalAlignment Center -WrapText -Range L:P
	$Style += New-ExcelStyle -VerticalAlignment Center -Range B:P


	$PlatformIssuesSheet = New-Object System.Collections.Generic.List[System.Object]
	$PlatformIssuesSheet.Add('REQUIRED ACTIONS / REVIEW STATUS')
	$PlatformIssuesSheet.Add('Tracking ID')
	$PlatformIssuesSheet.Add('Event Type')
	$PlatformIssuesSheet.Add('Event Source')
	$PlatformIssuesSheet.Add('Status')
	$PlatformIssuesSheet.Add('Title')
	$PlatformIssuesSheet.Add('Level')
	$PlatformIssuesSheet.Add('Event Level')
	$PlatformIssuesSheet.Add('Start Time')
	$PlatformIssuesSheet.Add('Mitigation Time')
	$PlatformIssuesSheet.Add('Impacted Service')
	$PlatformIssuesSheet.Add('What happened')
	$PlatformIssuesSheet.Add('What went wrong and why')
	$PlatformIssuesSheet.Add('How did we respond')
	$PlatformIssuesSheet.Add('How are we making incidents like this less likely or less impactful')
	$PlatformIssuesSheet.Add('How can customers make incidents like this less impactful')

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Exporting Platform Issues to Excel')
	$null = $PlatformIssuesFormatted | ForEach-Object { [PSCustomObject]$_ } | Select-Object $PlatformIssuesSheet |
    Export-Excel -ExcelPackage $excelPackage -WorksheetName $PlatformIssuesSheetRef -TableName 'platformIssues' -TableStyle $TableStyle -Style $Style -StartRow 12 -PassThru

}

<############################## Support Requests #########################################>

function Initialize-WARASupportTicket
{
	Param(
		$SupportTickets
	)

	$SupportTicketsMSG = Get-WARAMessage -Message 'SupportTickets_Standard'

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Formatting Support Tickets')
	# Just getting a sum of the total support tickets that will be included in the Excel file, in case there are no support tickets, a row will be added with empty values later
	$TotalSupportTickets = ($SupportTickets | Where-Object {![string]::IsNullOrEmpty($_.title)}).count

	$tmp = @()
	foreach ($Ticket in $SupportTickets)
		{
			if (![string]::IsNullOrEmpty($Ticket.title)) {
				$obj = @{
					'REQUIRED ACTIONS / REVIEW STATUS' 			= $SupportTicketsMSG;
					'Ticket ID' 								= $Ticket.'Ticket ID';
					'Severity' 									= $Ticket.Severity;
					'Status' 									= $Ticket.Status;
					'Support Plan Type' 						= $Ticket.'Support Plan Type';
					'Creation Date' 							= $Ticket.'Creation Date';
					'Modified Date' 							= $Ticket.'Modified Date';
					'Title' 									= $Ticket.Title;
					'Related Resource' 							= $Ticket.'Related Resource'
				}
				$tmp += $obj
			}
		}

	# If there are no support tickets, a row will be added with empty values
	if ($TotalSupportTickets -eq 0) {
		$obj = @{
			'REQUIRED ACTIONS / REVIEW STATUS'	 		= $SupportTicketsMSG;
			'Ticket ID' 								= '';
			'Severity' 									= '';
			'Status' 									= '';
			'Support Plan Type' 						= '';
			'Creation Date' 							= '';
			'Modified Date' 							= '';
			'Title' 									= '';
			'Related Resource' 							= ''
		}
		$tmp += $obj
	}

	return $tmp

}

function Export-WARASupportTicket
{
	Param($SupportTicketsFormatted,$ExcelPackage)

	$Style = @()
	$Style += New-ExcelStyle -HorizontalAlignment Left -WrapText -Range A:A
	$Style += New-ExcelStyle -HorizontalAlignment Center -Range A12:A12
	$Style += New-ExcelStyle -HorizontalAlignment Center -Range B:I
	$Style += New-ExcelStyle -VerticalAlignment Center -Range B:I



	$SupportTicketsSheet = New-Object System.Collections.Generic.List[System.Object]
	$SupportTicketsSheet.Add('REQUIRED ACTIONS / REVIEW STATUS')
	$SupportTicketsSheet.Add('Ticket ID')
	$SupportTicketsSheet.Add('Severity')
	$SupportTicketsSheet.Add('Status')
	$SupportTicketsSheet.Add('Support Plan Type')
	$SupportTicketsSheet.Add('Creation Date')
	$SupportTicketsSheet.Add('Modified Date')
	$SupportTicketsSheet.Add('Title')
	$SupportTicketsSheet.Add('Related Resource')


	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Exporting Support Tickets to Excel')
	$null = $SupportTicketsFormatted | ForEach-Object { [PSCustomObject]$_ } | Select-Object $SupportTicketsSheet |
		Export-Excel -ExcelPackage $ExcelPackage -WorksheetName $SupportRequestsSheetRef -TableName 'supportRequests' -TableStyle $TableStyle -Style $Style -StartRow 12 -NoNumberConversion "Ticket ID" -PassThru
}

<############################## Workload Inventory #########################################>

function Initialize-WARAWorkloadInventory
{
	Param(
		$InScopeResources,
		$TenantID
	)

	# Just getting a sum of the total in scope resources that will be included in the Excel file, in case there are no in scope resources, a row will be added with empty values later
	$TotalInScope = ($InScopeResources | Where-Object {![string]::IsNullOrEmpty($_.id)}).count

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Formatting Workload Inventory')

	$tmp = @()
	foreach ($resource in $InScopeResources)
		{
			if (![string]::IsNullOrEmpty($resource.id)) {
                $ResourceObj = [WorkLoadInvObj]::new()
                $ResourceObj.subscriptionId = $resource.subscriptionId
                $ResourceObj.resourceGroup = $resource.resourceGroup
                $ResourceObj.type = $resource.type
                $ResourceObj.location = $resource.location
                $ResourceObj.name = $resource.name
                $ResourceObj.id = $resource.id
                $ResourceObj.tenantId = $TenantID
                $ResourceObj.kind = $resource.kind
                $ResourceObj.managedBy = $resource.managedBy
                $ResourceObj.sku = [string]$resource.sku
                $ResourceObj.plan = $resource.plan
                $ResourceObj.zones = [string]$resource.zones

				$tmp += $ResourceObj
			}
		}

	if ($TotalInScope -eq 0) {
		$ResourceObj = [WorkLoadInvObj]::new()
		$tmp += $ResourceObj
	}

	return $tmp
}

function Export-WARAWorkloadInventory
{
	Param($WorkloadInventoryFormatted, $excelPackage)

	$Style = @()
	$Style += New-ExcelStyle -HorizontalAlignment Center


	$InScopeSheet = New-Object System.Collections.Generic.List[System.Object]
	$InScopeSheet.Add('id')
	$InScopeSheet.Add('name')
	$InScopeSheet.Add('type')
	$InScopeSheet.Add('tenantId')
	$InScopeSheet.Add('kind')
	$InScopeSheet.Add('location')
	$InScopeSheet.Add('resourceGroup')
	$InScopeSheet.Add('subscriptionId')
	$InScopeSheet.Add('managedBy')
	$InScopeSheet.Add('sku')
	$InScopeSheet.Add('plan')
	$InScopeSheet.Add('zones')

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Exporting Workload Inventory to Excel')
	$null = $WorkloadInventoryFormatted | ForEach-Object { [PSCustomObject]$_ } | Select-Object $InScopeSheet |
		Export-Excel -ExcelPackage $excelPackage -WorksheetName $WorkloadInventorySheetRef -TableName 'InScopeResources' -TableStyle $TableStyle -Style $Style -StartRow 12 -PassThru
}

<############################## Extra Configurations #########################################>

function Set-ExpertAnalysisFile
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		$ExcelPackage
	)

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Openning Excel File for Last Customization')
	$Excel = $ExcelPackage #Open-ExcelPackage -path $NewExpertAnalysisFile

	$sheet = $Excel.Workbook.Worksheets[$ImpactedResourcesSheetRef]

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Adding Conditional Formatting to Impacted Resources Sheet')

	Add-ConditionalFormatting -WorkSheet $sheet -RuleType ContainsText -ConditionValue "REQUIRED ACTIONS: Azure Service Health - Service Retirements" -Address A:A -BackgroundColor ([System.Drawing.Color]::fromArgb(0xd87406))
	Add-ConditionalFormatting -WorkSheet $sheet -RuleType ContainsText -ConditionValue "REQUIRED ACTIONS: Azure Service Health - Platform Issues" -Address A:A -BackgroundColor ([System.Drawing.Color]::fromArgb(0xee9432))
	Add-ConditionalFormatting -WorkSheet $sheet -RuleType ContainsText -ConditionValue "REQUIRED ACTIONS: Well-Architected Framework" -Address A:A -BackgroundColor ([System.Drawing.Color]::fromArgb(0xfbe757))
	Add-ConditionalFormatting -WorkSheet $sheet -RuleType ContainsText -ConditionValue "REQUIRED ACTIONS: ResourceType not available in APRL/Advisor" -Address A:A -BackgroundColor ([System.Drawing.Color]::fromArgb(0xfa7a06))
	Add-ConditionalFormatting -WorkSheet $sheet -RuleType ContainsText -ConditionValue "REQUIRED ACTIONS: Recommendation does not have automated validation." -Address A:A -BackgroundColor ([System.Drawing.Color]::fromArgb(0xffa500))
	Add-ConditionalFormatting -WorkSheet $sheet -RuleType ContainsText -ConditionValue "REQUIRED ACTIONS: Architectural and Reliability Design Patterns Recommendations" -Address A:A -BackgroundColor ([System.Drawing.Color]::fromArgb(0x92d050))

	$sheet = $Excel.Workbook.Worksheets[$PlatformIssuesSheetRef]

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Adding Conditional Formatting to Platform Issues Sheet')

	Add-ConditionalFormatting -WorkSheet $sheet -RuleType ContainsText -ConditionValue "REQUIRED ACTIONS: Review Platform Issue and create recommendations" -Address A:A -BackgroundColor ([System.Drawing.Color]::fromArgb(0xee9432))

	$sheet = $Excel.Workbook.Worksheets[$SupportRequestsSheetRef]

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Adding Conditional Formatting to Support Requests Sheet')

	Add-ConditionalFormatting -WorkSheet $sheet -RuleType ContainsText -ConditionValue "REQUIRED ACTIONS: Review Customer Support Requests and create recommendations" -Address A:A -BackgroundColor ([System.Drawing.Color]::fromArgb(0xFA7A06))

	Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Closing Excel File')
	#Close-ExcelPackage -ExcelPackage $Excel -SaveAs #-Calculate
}


# Start the stopwatch
$Runtime = [System.Diagnostics.Stopwatch]::StartNew()

# Excel Sheet Reference

$ImpactedResourcesSheetRef = '4.ImpactedResourcesAnalysis'
$AnalysisPlanningSheetRef = '3.AnalysisPlanning'
$PlatformIssuesSheetRef = '5.PlatformIssuesAnalysis'
$SupportRequestsSheetRef = '6.SupportRequestsAnalysis'
$WorkloadInventorySheetRef = '2.WorkloadInventory'

Write-Debug (' ---------------------------------- STARTING DATA ANALYZER SCRIPT --------------------------------------- ')
#Call the functions
$Version = '2.2.0'
Write-Host 'Version: ' -NoNewline
Write-Host $Version -ForegroundColor DarkBlue

Write-Host "Starting the " -NoNewline
Write-Host "WARA Analyzer" -ForegroundColor Magenta

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Read-JSONFile')
Test-Requirement

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Read-JSONFile')
$JSONContent = Read-JSONFile -JSONFile $JSONFile

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Importing Supported Types')
# Importing the CSV files to get the supported types and the friendly names for the resource types in the Retirements
$RootTypes = Invoke-RestMethod $RecommendationResourceTypesUri | ConvertFrom-Csv
$RootTypes = $RootTypes | Where-Object {$_.InAprlAndOrAdvisor -eq 'yes'}

Write-Host 'Analysing Excel File Template'

#$NewExpertAnalysisFile = Save-WARAExcelFile -ExpertAnalysisFile $ExpertAnalysisFile

$ExpertAnalysisTemplate = Open-ExcelPackage -Path $ExpertAnalysisFile

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Initialize-WARAImpactedResources')
# Creating the Array with the Impacted Resources to be added to the Excel file
$ImpactedResources 	= Initialize-WARAImpactedResources -ImpactedResources $JSONContent.ImpactedResources -Advisory $JSONContent.Advisory -Retirements $JSONContent.Retirements -ScriptDetails $JSONContent.ScriptDetails -RecommendationDataUri $RecommendationDataUri

Write-Host $ImpactedResourcesSheetRef -NoNewline -ForegroundColor Green
Write-Host ': ' -NoNewline
$ImpactResCount = $ImpactedResources | Measure-Object
Write-Host ([string]$ImpactResCount.count) -NoNewline -ForegroundColor Cyan
Write-Host ' (Lines to be added to the new Excel file)'

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Export-WARAImpactedResources')
# Adding the Impacted Resources to the Excel file
Export-WARAImpactedResources -ImpactedResourcesFormatted $ImpactedResources -ExcelPackage $ExpertAnalysisTemplate

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Initialize-WARAPlatformIssues')
# Creating the Array with the Platform Issues to be added to the Excel file
$PlatformIssues 	= Initialize-WARAPlatformIssues -PlatformIssues $JSONContent.Outages

Write-Host $PlatformIssuesSheetRef -NoNewline -ForegroundColor Green
Write-Host ': ' -NoNewline
$PlatissuesCount = $PlatformIssues | Measure-Object
Write-Host ([string]$PlatissuesCount.Count) -NoNewline -ForegroundColor Cyan
Write-Host ' (Lines to be added to the new Excel file)'

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Export-WaraPlatformIssues')
# Adding the Platform Issues to the Excel file
Export-WARAPlatformIssues -PlatformIssuesFormatted $PlatformIssues -excelPackage $ExpertAnalysisTemplate

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Initialize-WARASupportTicket')
# Creating the Array with the Support Tickets to be added to the Excel file
$SupportTickets 	= Initialize-WARASupportTicket -SupportTickets $JSONContent.SupportTickets

Write-Host $SupportRequestsSheetRef -NoNewline -ForegroundColor Green
Write-Host ': ' -NoNewline
$SuppTicketsCount = $SupportTickets | Measure-Object
Write-Host ([string]$SuppTicketsCount.count) -NoNewline -ForegroundColor Cyan
Write-Host ' (Lines to be added to the new Excel file)'

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Export-WaraSupportTicket')
# Adding the Support Tickets to the Excel file
Export-WARASupportTicket -SupportTicketsFormatted $SupportTickets -ExcelPackage $ExpertAnalysisTemplate

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Initialize-WARAAnalysisPlanning')
# Creating the Array with the Analysis Planning to be added to the Excel file
$AnalysisPlanning 	= Initialize-WARAAnalysisPlanning -InScopeResources $JSONContent.impactedResources

Write-Host $AnalysisPlanningSheetRef -NoNewline -ForegroundColor Green
Write-Host ': ' -NoNewline
$AnalysisPlanningCount = $AnalysisPlanning | Measure-Object
Write-Host ([string]$AnalysisPlanningCount.count) -NoNewline -ForegroundColor Cyan
Write-Host ' (Lines to be added to the new Excel file)'

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Export-WaraAnalysisPlanning')
# Adding the Analysis Planning to the Excel file
Export-WARAAnalysisPlanning -AnalysisPlanningFormatted $AnalysisPlanning -ExcelPackage $ExpertAnalysisTemplate

Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Initialize-WARAWorkloadInventory')
# Creating the Array with the Workload Inventory to be added to the Excel file
$WorkloadInventory 	= Initialize-WARAWorkloadInventory -InScopeResources $JSONContent.resourceInventory -TenantID $JSONContent.ScriptDetails.TenantId

Write-Host $WorkloadInventorySheetRef -NoNewline -ForegroundColor Green
Write-Host ': ' -NoNewline
$WorkloadInvCount = $WorkloadInventory | Measure-Object
Write-Host ([string]$WorkloadInvCount.count) -NoNewline -ForegroundColor Cyan
Write-Host ' (Lines to be added to the new Excel file)'

# Adding the Workload Inventory to the Excel file
Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Export-WaraWorkloadInventory')
Export-WARAWorkloadInventory -WorkloadInventoryFormatted $WorkloadInventory -excelPackage $ExpertAnalysisTemplate

Write-Host 'Overall Excel File' -NoNewline -ForegroundColor Green
Write-Host ': ' -NoNewline
Write-Host 'Extra Excel Customization' -ForegroundColor Cyan

# Setting the Excel file with the extra configurations like the conditional formatting
Write-Debug ((get-date -Format 'yyyy-MM-dd HH:mm:ss') + ' - Invoking Function: Set-ExpertAnalysisFile')
#Set-ExpertAnalysisFile -ExcelPackage $ExpertAnalysisTemplate

$NewExpertAnalysisFile = Save-WARAExcelFile -ExcelPackage $ExpertAnalysisTemplate

$Runtime.Stop()
$TotalTime = $Runtime.Elapsed.toString('hh\:mm\:ss')

Write-Host '---------------------------------------------------------------------'
Write-Host ('Execution Complete. Total Runtime was: ') -NoNewline
Write-Host $TotalTime -NoNewline -ForegroundColor Cyan
Write-Host (' Minutes')
Write-Host 'Excel File: ' -NoNewline
Write-Host $NewExpertAnalysisFile -ForegroundColor Blue
Write-Host '---------------------------------------------------------------------'

# SIG # Begin signature block
# MIIoLwYJKoZIhvcNAQcCoIIoIDCCKBwCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDHTFKIvgztBVsi
# Fcg3KlHmvMR6xVVU+Y0HEpoQMovpXqCCDXYwggX0MIID3KADAgECAhMzAAAEBGx0
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
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIDN/NICvga1S7NZ8m8osJ3vV
# hEr0pglh+XBxCnHuJVa6MEQGCisGAQQBgjcCAQwxNjA0oBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEcgBpodHRwczovL3d3dy5taWNyb3NvZnQuY29tIDANBgkqhkiG9w0B
# AQEFAASCAQAmqNUqLc6oyUcUyq6SxUF1ftoelGZHPVjQCNb8QovbWsH0UzDSzpeX
# xARBAypwFRCyVI2wmziVyqgWNDnUdhAzKxqmydr0vIidITKm+zpFojBdTp3dxb+B
# n6prkmiK01IrpOU2rMvmx3+RhbaQz48FJlQMzYuBSpn+jvWv0UEY7dmnm/fP9vt1
# aRYBXy1Dx75HRtW1UmoM+Q0GWvfH9kh7F30gVDPEUymkl1mgkjj6X4zUgRfZMJIj
# z6FWQtyAi++RHqr10kG4vkS3bc1nmJxv1YOHIq7fildCun2aeGrlMmZPwBeGHbx0
# hMIy90CWRxXZOg7HgxTHrlJbEe8hfWkwoYIXlzCCF5MGCisGAQQBgjcDAwExgheD
# MIIXfwYJKoZIhvcNAQcCoIIXcDCCF2wCAQMxDzANBglghkgBZQMEAgEFADCCAVIG
# CyqGSIb3DQEJEAEEoIIBQQSCAT0wggE5AgEBBgorBgEEAYRZCgMBMDEwDQYJYIZI
# AWUDBAIBBQAEIFraOxWVdxpx73Yhmv5dSuvQcNYzqxDAtWu1HAev1+M0AgZoK7oE
# TE8YEzIwMjUwNTIyMTUwMDM0LjIyOFowBIACAfSggdGkgc4wgcsxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
# bWVyaWNhIE9wZXJhdGlvbnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo5NjAw
# LTA1RTAtRDk0NzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZaCCEe0wggcgMIIFCKADAgECAhMzAAACBNjgDgeXMliYAAEAAAIEMA0GCSqGSIb3
# DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
# BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTI1MDEzMDE5
# NDI0N1oXDTI2MDQyMjE5NDI0N1owgcsxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlv
# bnMxJzAlBgNVBAsTHm5TaGllbGQgVFNTIEVTTjo5NjAwLTA1RTAtRDk0NzElMCMG
# A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZTCCAiIwDQYJKoZIhvcN
# AQEBBQADggIPADCCAgoCggIBAPDdJtx57Z3rq+RYZMheF8aqqBAbFBdOerjheVS8
# 3MVK3sQu07gH3f2PBkVfsOtG3/h+nMY2QV0alzsQvlLzqopi/frR5eNb58i/WUCo
# MPfV3+nwCL38BnPwz3nOjSsOkrZyzP1YDJH0W1QPHnZU6z2o/f+mCke+BS8Pyzr/
# co0hPOazxALW0ndMzDVxGf0JmBUhjPDaIP9m85bSxsX8NF2AzxR23GMUgpNdNoj9
# smGxCB7dPBrIpDaPzlFp8UVUJHn8KFqmSsFBYbA0Vo/OmZg3jqY+I69TGuIhIL2d
# D8asNdQlbMsOZyGuavZtoAEl6+/DfVRiVOUtljrNSaOSBpF+mjN34aWr1NjYTcOC
# Wvo+1MQqA+7aEzq/w2JTmdO/GEOfF2Zx/xQ3uCh5WUQtds6buPzLDXEz0jLJC5Qx
# aSisFo3/mv2DiW9iQyiFFcRgHS0xo4+3QWZmZAwsEWk1FWdcFNriFpe+fVp0qu9P
# PxWV+cfGQfquID+HYCWphaG/RhQuwRwedoNaCoDb2vL6MfT3sykn8UcYfGT532Qf
# Yvlok+kBi42Yw08HsUNM9YDHsCmOv8nkyFTHSLTuBXZusBn0n1EeL58w9tL5CbgC
# icLmI5OP50oK21VGz6Moq47rcIvCqWWO+dQKa5Jq85fnghc60pwVmR8N05ntwTgO
# Kg/VAgMBAAGjggFJMIIBRTAdBgNVHQ4EFgQUGnV2S0Bwalb8qbqqb6+7gzUZol8w
# HwYDVR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYDVR0fBFgwVjBUoFKg
# UIZOaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0
# JTIwVGltZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsGAQUFBwEBBGAw
# XjBcBggrBgEFBQcwAoZQaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# ZXJ0cy9NaWNyb3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcnQw
# DAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAF5y/qxHDYdMszJQLVYkn4VH4OAD0mS/
# SUawi3jLr0KY6PxHregVuFKZx2lqTGo1uvy/13JNvhEPI2q2iGKJdu2teZArlfvL
# 9D74XTMyi1O1OlM+8bd6W3JX8u87Xmasug1DtbhUfnxou3TfS05HGzxWcBBAXkGZ
# BAw65r4RCAfh/UXi4XquXcQLXskFInTCMdJ5r+fRZiIc9HSqTP81EB/yVJRRXSBs
# gxrAYiOfv5ErIKv7yXXF02Qr8XRRi5feEbScT71ZzQvgD96eW5Q3s9r285XpWLcE
# 4lJPRFj9rHuJnjmV4zySoLDsEU9xMiRbPGmOvacK2KueTDs4FDoU2DAi4C9g1NTu
# vrRbjbVgU4vmlOwxlw0M46wDTXG/vKYIXrOScwalEe7DRFvYEAkL2q5TsJdZsxsA
# kt1npcg0pquJKYJff8wt3Nxblc7JwrRCGhE1F/hapdGyEQFpjbKYm8c7jyhJJj+S
# m5i8FLeWMAC4s3tGnyNZLu33XqloZ4Tumuas/0UmyjLUsUqYWdb6+DjcA2EHK4AR
# er0JrLmjsrYfk0WdHnCP9ItErArWLJRf3bqLVMS+ISICH89XIlsAPiSiKmKDbyn/
# ocO6Jg5nTBSSb9rlbyisiOg51TdewniLTwJ82nkjvcKy8HlA9gxwukX007/Uu+hA
# DDdQ90vnkzkdMIIHcTCCBVmgAwIBAgITMwAAABXF52ueAptJmQAAAAAAFTANBgkq
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
# MScwJQYDVQQLEx5uU2hpZWxkIFRTUyBFU046OTYwMC0wNUUwLUQ5NDcxJTAjBgNV
# BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMV
# ALo9gdHD371If7WnDLqrNUbeT2VuoIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzAR
# BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
# Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3Rh
# bXAgUENBIDIwMTAwDQYJKoZIhvcNAQELBQACBQDr2YQKMCIYDzIwMjUwNTIyMTEw
# NzU0WhgPMjAyNTA1MjMxMTA3NTRaMHcwPQYKKwYBBAGEWQoEATEvMC0wCgIFAOvZ
# hAoCAQAwCgIBAAICCboCAf8wBwIBAAICEj0wCgIFAOva1YoCAQAwNgYKKwYBBAGE
# WQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDAN
# BgkqhkiG9w0BAQsFAAOCAQEAgKgp+Z4z4aLFO3KyquXlgCcJ4jFV0x7OezP8JZhI
# Fb5elFs78vtbmVQZBHYLDo/jhErOxTuVK85tVV4YDFbqvJpcgc2OliKg/7YBKHV6
# fk27y5g73AJ5BZaOz3hJMiotH9WBl34uTYOBSOYSrN39VUCmv9j3OPd4d1JtnNpO
# o/favC2Cm5mnERZeU5uozO+5SYmMOibc4wNM2D5XUwNI5BORGS361EP68Wqdkbjm
# C1Hc2JFNY4KBoMwNORblINx4wd9w84Ib4Zj1slMdoOf/QJjLUDedDqsPlC97LSX4
# 12w03y3g7gVAUAL1MyHoMXTiFUKK5AYPlNly1n5aR/lObTGCBA0wggQJAgEBMIGT
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAACBNjgDgeXMliYAAEA
# AAIEMA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQ
# AQQwLwYJKoZIhvcNAQkEMSIEIDdkHddW5bILA6C12DytjNnSs0+KOeJAxPmenYH4
# bsyiMIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQg+e14Zf1bCrxV0kzqaN/H
# UYQmy7v/qRTqXRJLmtx5uf4wgZgwgYCkfjB8MQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
# b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQ
# Q0EgMjAxMAITMwAAAgTY4A4HlzJYmAABAAACBDAiBCAD3okBinK54U4dOkNJMIkg
# 0OXYaBOTDIWg/jl/zc38djANBgkqhkiG9w0BAQsFAASCAgBVyybCCAS8wsDrYid3
# nOsrosRIH/+dqL+bP5KNlriRBmPUli1YG8j4B5r8mNBmclIY91xp8bYFuJKsOMj0
# ZOHrPcZwaaFDR51pwuqD6ssNgaSAiZ/EJH+ptf7IZrX9dsDR+lisXoqTIElJ0rSW
# HMLfJHxsz864eQsbmG97GV4zRKpr7ZMNsQNk0L/nfDKkhun3eZj4Y/6tjVvo/PR5
# cN82umTFbRG+kobs65Dxm45CcR7VSOahNxoy05x3CqJ7dE4DYIzFOMfhwvSzBQOT
# /HLAw/P8npf8msjQvj2M2aabPuHDSDrI6er9otmdRsh+9MbxYPAQTy/Y1xe8URLI
# BLNRz7nk7q1/D+v9qSBB77zGuHkEMMMxDPj1K3z5tV1z+xj/GO6L4hiGzcnk/ieS
# alUXAHFpGZby2JSlhvH5T5HbJa9TGMbPunXqWxBMb4CA8ymUlryMGPaLjX+QTceC
# n5qtQaXoyWf2Kb2fog375tektO3DO50dWCYh+uGUBLwMZ/yvMmRCSaTUG4XINne8
# Ux9VxwtvAEZuPOvgJhfVisFsua3b4m6/1u7SN4J6HHBA1rp7l/rrza8B2BPEXEyc
# 1h/NSvNgRlNSUtAQInRnQZm0Yirq5fukTqMrrXZe/46yR/glw6cjeZjp21iEvnJp
# a1x4Gga9mG+qNmsrNKVBzoffrA==
# SIG # End signature block
