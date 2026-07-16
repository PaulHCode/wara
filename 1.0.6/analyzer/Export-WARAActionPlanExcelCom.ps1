#Requires -Version 7
<#
.SYNOPSIS
    Builds the WARA "Expert Analysis" action-plan Excel workbook from a collector JSON file using
    local Microsoft Excel (COM automation) instead of the third-party ImportExcel module.

.DESCRIPTION
    This is a Microsoft-only replacement for the ImportExcel-based logic in 2_wara_data_analyzer.ps1.
    It reuses the exact same data-shaping the analyzer performs, then writes the workbook by driving a
    local Excel instance via COM. It requires desktop Excel to be installed. No PowerShell Gallery /
    internet access is needed (recommendation data is read from the bundled offline-data folder by default).

    Sheets produced (starting from the shipped Expert-Analysis-v1.xlsx template):
      2.WorkloadInventory        (table InScopeResources)
      3.AnalysisPlanning         (table TableTypes8, live count/status formulas)
      4.ImpactedResourcesAnalysis(table impactedresources, data validation + conditional formatting)
      5.PlatformIssuesAnalysis   (table platformIssues, conditional formatting)
      6.SupportRequestsAnalysis  (table supportRequests, conditional formatting)

.PARAMETER JSONFile
    Path to the JSON produced by Start-WARACollector.

.PARAMETER TemplatePath
    Path to the Expert-Analysis-v1.xlsx template. Defaults to the copy shipped beside this script.

.PARAMETER RecommendationDataUri
    Local path or http(s) URI to recommendations.json. Defaults to the bundled offline copy.

.PARAMETER RecommendationResourceTypesUri
    Local path or http(s) URI to WARAinScopeResTypes.csv. Defaults to the bundled offline copy.

.PARAMETER CustomRecommendationObject
    Optional path to a custom recommendations JSON to merge in.

.PARAMETER OutputPath
    Full path for the generated .xlsx. Defaults to .\Expert-Analysis-v1-<timestamp>.xlsx in the current dir.

.PARAMETER ShowExcel
    Show the Excel window while generating (useful for debugging). Default is hidden.

.EXAMPLE
    .\Export-WARAActionPlanExcelCom.ps1 -JSONFile 'C:\Temp\WARA-File-2026-07-15-18-28.json'

.NOTES
    Microsoft-only (Excel COM) alternative to ImportExcel. Iterative build.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string] $JSONFile,

    [string] $TemplatePath = (Join-Path $PSScriptRoot 'Expert-Analysis-v1.xlsx'),

    [ValidateScript({ ($_ -match '^https?://') -or (Test-Path -LiteralPath $_ -PathType Leaf) })]
    [string] $RecommendationDataUri = (Join-Path $PSScriptRoot '../offline-data/recommendations.json'),

    [ValidateScript({ ($_ -match '^https?://') -or (Test-Path -LiteralPath $_ -PathType Leaf) })]
    [string] $RecommendationResourceTypesUri = (Join-Path $PSScriptRoot '../offline-data/WARAinScopeResTypes.csv'),

    [string] $CustomRecommendationObject,

    [string] $OutputPath,

    [switch] $ShowExcel
)

$ErrorActionPreference = 'Stop'

# Excel sheet references (must match the template sheet names).
$ImpactedResourcesSheetRef = '4.ImpactedResourcesAnalysis'
$AnalysisPlanningSheetRef  = '3.AnalysisPlanning'
$PlatformIssuesSheetRef    = '5.PlatformIssuesAnalysis'
$SupportRequestsSheetRef   = '6.SupportRequestsAnalysis'
$WorkloadInventorySheetRef = '2.WorkloadInventory'

#region ------------------------------------------------------------ Classes
class ImpactedResourceObj {
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

class AnalysisPlanningObj {
    [string] $Category
    [string] $ResourceType = 'N/A'
    [string] $NumberOfResources = 'n/a'
    [string] $ImpactedResources = 'n/a'
    [string] $HasRecommendationsInAPRLAdvisor = 'Yes'
    [string] $AssessmentStatus = 'Pending'
}

class OutagesObj {
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

class WorkLoadInvObj {
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
#endregion

#region ------------------------------------------------------------ Data prep (ported from analyzer)
function Get-WARAMessage {
    param($Message)

    if ($Message -eq 'ImpactedResources_Type') {
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
    if ($Message -eq 'ImpactedResources_Unavailable') {
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
    if ($Message -eq 'ImpactedResources_ServiceRetirement') {
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
    if ($Message -eq 'ImpactedResources_Architecture') {
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
    if ($Message -eq 'ImpactedResources_WAF') {
        $WARAMessage = @"
REQUIRED ACTIONS: Well-Architected Framework.
This is a generic recommendation from the Well-Architected Framework - Reliability pillar. It cannot be validated automatically and need information from the Discovery Workshop session.
1. Review and update this row based on the Discovery Workshop Questionnaire.
2. If this recommendation is applicable, update the "Impact" column according to the importance for the Workload and then update this cell to "Reviewed".
or
3. If this recommendation is not relevant, delete this row.
"@
    }
    if ($Message -eq 'PlatformIssues_Standard') {
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
    if ($Message -eq 'SupportTickets_Standard') {
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

function Get-WARARecommendationList {
    param([Parameter(Mandatory = $true)][string]$RecommendationDataUri)
    if (Test-Path -LiteralPath $RecommendationDataUri -PathType Leaf) {
        return (Get-Content -LiteralPath $RecommendationDataUri -Raw | ConvertFrom-Json)
    }
    return (Invoke-RestMethod $RecommendationDataUri)
}

function Initialize-WARAImpactedResources {
    param($ImpactedResources, $Advisory, $Retirements, $ScriptDetails, $RecommendationDataUri, $RootTypes)

    $RecommendationObject = Get-WARARecommendationList -RecommendationDataUri $RecommendationDataUri

    if ($CustomRecommendationObject) {
        $custom = Get-Content $CustomRecommendationObject -Raw | ConvertFrom-Json -Depth 20
        $RecommendationObject += $custom
    }

    $ResourceRecommendations = $RecommendationObject | Where-Object { [string]::IsNullOrEmpty($_.tags) }
    $WAFRecommendations = $RecommendationObject | Where-Object { $_.tags -like 'WAF' }

    if ($ScriptDetails.SAP -eq 'True') { $ResourceRecommendations += $RecommendationObject | Where-Object { $_.tags -like 'SAP' } }
    if ($ScriptDetails.AVD -eq 'True') { $ResourceRecommendations += $RecommendationObject | Where-Object { $_.tags -like 'AVD' } }
    if ($ScriptDetails.AVS -eq 'True') { $ResourceRecommendations += $RecommendationObject | Where-Object { $_.tags -like 'AVS' } }
    if ($ScriptDetails.HPC -eq 'True') { $ResourceRecommendations += $RecommendationObject | Where-Object { $_.tags -like 'HPC' } }
    if ($ScriptDetails.AI_GPT_RAG -eq 'True') { $ResourceRecommendations += $RecommendationObject | Where-Object { $_.tags -like 'AI-GPT-RAG' } }
    if ($ScriptDetails.ORACLE -eq 'True') { $ResourceRecommendations += $RecommendationObject | Where-Object { $_.tags -like 'ORACLE' } }

    $tmp = @()

    foreach ($Recom in $ResourceRecommendations) {
        $Resources = $ImpactedResources | Where-Object { ($_.recommendationId -eq $Recom.aprlGuid) -and ($_.checkName -eq $Recom.checkName) }
        if ([string]::IsNullOrEmpty($Resources) -and $Recom.aprlGuid -notin $tmp.Guid -and -not $Recom.checkName) {
            $Resources = $ImpactedResources | Where-Object { ($_.recommendationId -eq $Recom.aprlGuid) }
        }
        foreach ($Resource in $Resources) {
            $ValidationMSG = switch ($Resource.validationAction) {
                'IMPORTANT - Query under development - Validate Resources manually' { Get-WARAMessage -Message 'ImpactedResources_Unavailable' }
                'IMPORTANT - Recommendation cannot be validated with ARGs - Validate Resources manually' { Get-WARAMessage -Message 'ImpactedResources_Unavailable' }
                'IMPORTANT - Resource Type is not available in either APRL or Advisor - Validate Resources manually if Applicable, if not Delete this line' { Get-WARAMessage -Message 'ImpactedResources_Type' }
                'APRL - Queries' { 'Reviewed' }
                default { 'Error' }
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

    $Resources = $ImpactedResources | Where-Object { [string]::IsNullOrEmpty($_.recommendationId) }
    foreach ($Resource in $Resources) {
        $ValidationMSG = switch ($Resource.validationAction) {
            'IMPORTANT - Resource Type is not available in either APRL or Advisor - Validate Resources manually if Applicable, if not Delete this line' { Get-WARAMessage -Message 'ImpactedResources_Type' }
            default { 'Custom' }
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

    $ADVMessage = 'Reviewed'
    foreach ($adv in $Advisory) {
        if (![string]::IsNullOrEmpty($adv)) {
            $ADVobj = [ImpactedResourceObj]::new()
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
            $ADVobj.Category = 'Azure Service'
            $ADVobj.Source = 'ADVISOR'
            $ADVobj.WAFPillar = 'Reliability'
            $tmp += $ADVobj
        }
    }

    $ServiceRetirementMSG = Get-WARAMessage -Message 'ImpactedResources_ServiceRetirement'
    foreach ($Retirement in $Retirements) {
        if (![string]::IsNullOrEmpty($Retirement)) {
            $RetirementType = $RootTypes | Where-Object { $_.FriendlyName -eq $Retirement.ImpactedService }
            $RetirementType = if (![string]::IsNullOrEmpty($RetirementType)) { $RetirementType.ResourceType } else { $Retirement.ImpactedService }
            try {
                $HTML = New-Object -Com 'HTMLFile'
                $HTML.write([ref]$Retirement.Description)
                $RetirementDescriptionFull = $Html.body.innerText
                $SplitDescription = $RetirementDescriptionFull.split('Help and support').split('Required action')
            }
            catch { $SplitDescription = ' ', ' ' }
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

    $WAFMSG = Get-WARAMessage -Message 'ImpactedResources_WAF'
    foreach ($waf in $WAFRecommendations) {
        if (![string]::IsNullOrEmpty($waf)) {
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

    $ArchtectureMSG = Get-WARAMessage -Message 'ImpactedResources_Architecture'
    $ARCHObj = [ImpactedResourceObj]::new()
    $ARCHObj.ValidationMSG = $ArchtectureMSG
    $ARCHObj.ValidationCategory = 'Architectural'
    $ARCHObj.ResourceType = 'Microsoft.Subscription/Subscriptions'
    $ARCHObj.Impact = 'Low'
    $ARCHObj.RecommendationControl = 'Governance'
    $tmp += $ARCHObj

    $ImpactedResourcesFormatted = foreach ($line in $tmp) {
        [PSCustomObject][ordered]@{
            'REQUIRED ACTIONS / REVIEW STATUS' = $line.ValidationMSG
            'ValidationCategory'               = $line.ValidationCategory
            'Resource Type'                    = $line.ResourceType
            'subscriptionId'                   = $line.SubscriptionId
            'resourceGroup'                    = $line.ResourceGroup
            'location'                         = $line.Location
            'name'                             = $line.Name
            'id'                               = $line.Id
            'custom1'                          = $line.Custom1
            'custom2'                          = $line.Custom2
            'custom3'                          = $line.Custom3
            'custom4'                          = $line.Custom4
            'custom5'                          = $line.Custom5
            'Recommendation Title'             = $line.RecommendationTitle
            'Impact'                           = $line.Impact
            'Recommendation Control'           = $line.RecommendationControl
            'Potential Benefit'                = $line.PotentialBenefit
            'Learn More Link'                  = $line.LearnMoreLink
            'Long Description'                 = $line.LongDescription
            'Guid'                             = $line.Guid
            'Category'                         = $line.Category
            'Source'                           = $line.Source
            'WAF Pillar'                       = $line.WAFPillar
            'Platform Issue TrackingId'        = $line.PlatformIssueTrackingId
            'Retirement TrackingId'            = $line.RetirementTrackingId
            'Support Request Number'           = $line.SupportRequestNumber
            'Notes'                            = $line.Notes
            'checkName'                        = $line.CheckName
        }
    }
    return , $ImpactedResourcesFormatted
}

function Initialize-WARAAnalysisPlanning {
    param($InScopeResources, $RootTypes)

    $ResourceTypes = $InScopeResources | Group-Object -Property type

    $InventoryFormula = "=COUNTA(UNIQUE(VSTACK(FILTER('$WorkloadInventorySheetRef'!A:A, '$WorkloadInventorySheetRef'!C:C = TableTypes8[[#This Row],[Resource Type]]), FILTER('$WorkloadInventorySheetRef'!A:A, '$WorkloadInventorySheetRef'!C:C = TableTypes8[[#This Row],[Resource Type]]))))"
    $ImpactedResourcesFormula = "=IF(COUNTIF('$ImpactedResourcesSheetRef'!C:C, TableTypes8[[#This Row],[Resource Type]])=0, 0, COUNTA(UNIQUE(FILTER('$ImpactedResourcesSheetRef'!H:H, ('$ImpactedResourcesSheetRef'!C:C=TableTypes8[[#This Row],[Resource Type]]) * ('$ImpactedResourcesSheetRef'!H:H<>""Get ResourceID from Azure Portal"")))))"
    $ReviewedFormula = "=IF(OR(AND(TableTypes8[[#This Row],[Category]]=""Support Requests"", COUNTIFS('$SupportRequestsSheetRef'!A:A, ""<>Reviewed"")=0),AND(TableTypes8[[#This Row],[Category]]=""Platform Issues"", COUNTIFS('$PlatformIssuesSheetRef'!A:A, ""<>Reviewed"")=0),AND(TableTypes8[[#This Row],[Category]]=""Impacted Resources"", COUNTIFS('$ImpactedResourcesSheetRef'!A:A, ""<>Reviewed"", '$ImpactedResourcesSheetRef'!C:C, TableTypes8[[#This Row],[Resource Type]], '$ImpactedResourcesSheetRef'!O:O, ""<>Low"")=0)), ""Reviewed"", ""Pending"")"

    $tmp = @()
    foreach ($ResourceType in $ResourceTypes) {
        $RootType = $RootTypes | Where-Object { $_.ResourceType -eq $ResourceType.Name }
        $APRLOrAdv = if ($RootType.WARAinScope -eq 'yes' -and $RootType.InAprlAndOrAdvisor -eq 'yes') { 'Yes' } else { 'No' }
        $ResTypeObj = [AnalysisPlanningObj]::new()
        $ResTypeObj.Category = 'Impacted Resources'
        $ResTypeObj.ResourceType = $ResourceType.Name
        $ResTypeObj.NumberOfResources = $InventoryFormula
        $ResTypeObj.ImpactedResources = $ImpactedResourcesFormula
        $ResTypeObj.HasRecommendationsInAPRLAdvisor = $APRLOrAdv
        $ResTypeObj.AssessmentStatus = $ReviewedFormula
        $tmp += $ResTypeObj
    }

    $SupObj = [AnalysisPlanningObj]::new(); $SupObj.Category = 'Support Requests'; $tmp += $SupObj
    $PlatObj = [AnalysisPlanningObj]::new(); $PlatObj.Category = 'Platform Issues'; $tmp += $PlatObj

    $AnalysisPlanningFormatted = foreach ($line in $tmp) {
        [PSCustomObject][ordered]@{
            'Category'                                   = $line.Category
            'Resource Type'                              = $line.ResourceType
            'Number of Resources'                        = $line.NumberOfResources
            'Impacted Resources'                         = $line.ImpactedResources
            'Has Recommendations_x000a_in APRL/Advisor'  = $line.HasRecommendationsInAPRLAdvisor
            'Assessment Owner'                           = ''
            'Assessment Status'                          = $line.AssessmentStatus
            'Notes'                                      = ''
        }
    }
    return , $AnalysisPlanningFormatted
}

function Initialize-WARAPlatformIssues {
    param($PlatformIssues)

    $OutagesMSG = Get-WARAMessage -Message 'PlatformIssues_Standard'
    $TotalOutages = ($PlatformIssues | Where-Object { $_.properties.description -like '*How can customers make incidents like this less impactful?*' }).count

    $tmp = @()
    foreach ($Outage in $PlatformIssues) {
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
            catch { $whathap = ''; $whatwent = ''; $howdid = ''; $howarewe = ''; $howcan = '' }

            $ImpactedSvc = if ($Outage.properties.impact.impactedService.count -gt 1) { $Outage.properties.impact.impactedService | ForEach-Object { $_ + ' ,' } } else { $Outage.properties.impact.impactedService }
            $ImpactedSvc = [string]$ImpactedSvc
            $ImpactedSvc = if ($ImpactedSvc -like '* ,*') { $ImpactedSvc -replace ".$" } else { $ImpactedSvc }

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

    if ($TotalOutages -eq 0) {
        $ZeroOutageObj = [OutagesObj]::new()
        $ZeroOutageObj.OutageMSG = $OutagesMSG
        $tmp += $ZeroOutageObj
    }

    $PlatformIssuesFormatted = foreach ($line in $tmp) {
        [PSCustomObject][ordered]@{
            'REQUIRED ACTIONS / REVIEW STATUS'                                     = $line.OutageMSG
            'Tracking ID'                                                          = $line.TrackingID
            'Event Type'                                                           = $line.EventType
            'Event Source'                                                         = $line.EventSource
            'Status'                                                               = $line.Status
            'Title'                                                                = $line.Title
            'Level'                                                                = $line.Level
            'Event Level'                                                          = $line.EventLevel
            'Start Time'                                                           = $line.StartTime
            'Mitigation Time'                                                      = $line.MitigationTime
            'Impacted Service'                                                     = $line.ImpactedService
            'What happened'                                                        = $line.WhatHappened
            'What went wrong and why'                                              = $line.WhatWentWrongAndWhy
            'How did we respond'                                                   = $line.HowDidWeRespond
            'How are we making incidents like this less likely or less impactful'  = $line.HowAreWeMakingIncidentsLessLikely
            'How can customers make incidents like this less impactful'            = $line.HowCanCustomersMakeIncidentsLessImpactful
        }
    }
    return , $PlatformIssuesFormatted
}

function Initialize-WARASupportTicket {
    param($SupportTickets)

    $SupportTicketsMSG = Get-WARAMessage -Message 'SupportTickets_Standard'
    $TotalSupportTickets = ($SupportTickets | Where-Object { ![string]::IsNullOrEmpty($_.title) }).count

    $tmp = @()
    foreach ($Ticket in $SupportTickets) {
        if (![string]::IsNullOrEmpty($Ticket.title)) {
            $tmp += [PSCustomObject][ordered]@{
                'REQUIRED ACTIONS / REVIEW STATUS' = $SupportTicketsMSG
                'Ticket ID'                        = $Ticket.'Ticket ID'
                'Severity'                         = $Ticket.Severity
                'Status'                           = $Ticket.Status
                'Support Plan Type'                = $Ticket.'Support Plan Type'
                'Creation Date'                    = $Ticket.'Creation Date'
                'Modified Date'                    = $Ticket.'Modified Date'
                'Title'                            = $Ticket.Title
                'Related Resource'                 = $Ticket.'Related Resource'
            }
        }
    }
    if ($TotalSupportTickets -eq 0) {
        $tmp += [PSCustomObject][ordered]@{
            'REQUIRED ACTIONS / REVIEW STATUS' = $SupportTicketsMSG
            'Ticket ID'                        = ''
            'Severity'                         = ''
            'Status'                           = ''
            'Support Plan Type'                = ''
            'Creation Date'                    = ''
            'Modified Date'                    = ''
            'Title'                            = ''
            'Related Resource'                 = ''
        }
    }
    return , $tmp
}

function Initialize-WARAWorkloadInventory {
    param($InScopeResources, $TenantID)

    $TotalInScope = ($InScopeResources | Where-Object { ![string]::IsNullOrEmpty($_.id) }).count
    $tmp = @()
    foreach ($resource in $InScopeResources) {
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
    if ($TotalInScope -eq 0) { $tmp += [WorkLoadInvObj]::new() }

    $WorkloadInventoryFormatted = foreach ($line in $tmp) {
        [PSCustomObject][ordered]@{
            'id'             = $line.id
            'name'           = $line.name
            'type'           = $line.type
            'tenantId'       = $line.tenantId
            'kind'           = $line.kind
            'location'       = $line.location
            'resourceGroup'  = $line.resourceGroup
            'subscriptionId' = $line.subscriptionId
            'managedBy'      = $line.managedBy
            'sku'            = $line.sku
            'plan'           = $line.plan
            'zones'          = $line.zones
        }
    }
    return , $WorkloadInventoryFormatted
}
#endregion

#region ------------------------------------------------------------ Main
Write-Host 'WARA Action Plan (Excel COM) generator' -ForegroundColor Magenta

if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) { throw "Template not found: $TemplatePath" }

$JSONFile = (Resolve-Path -LiteralPath $JSONFile).Path
$JSONContent = Get-Content -LiteralPath $JSONFile -Raw | ConvertFrom-Json

# In-scope resource types friendly-name map (offline-aware).
$RootTypes = if (Test-Path -LiteralPath $RecommendationResourceTypesUri -PathType Leaf) {
    Get-Content -LiteralPath $RecommendationResourceTypesUri -Raw
}
else { Invoke-RestMethod $RecommendationResourceTypesUri }
$RootTypes = $RootTypes | ConvertFrom-Csv | Where-Object { $_.InAprlAndOrAdvisor -eq 'yes' }

Write-Host 'Preparing data...' -ForegroundColor Cyan
$ImpactedResources = Initialize-WARAImpactedResources -ImpactedResources $JSONContent.impactedResources -Advisory $JSONContent.advisory -Retirements $JSONContent.retirements -ScriptDetails $JSONContent.scriptDetails -RecommendationDataUri $RecommendationDataUri -RootTypes $RootTypes
$PlatformIssues    = Initialize-WARAPlatformIssues -PlatformIssues $JSONContent.outages
$SupportTickets    = Initialize-WARASupportTicket -SupportTickets $JSONContent.supportTickets
$AnalysisPlanning  = Initialize-WARAAnalysisPlanning -InScopeResources $JSONContent.impactedResources -RootTypes $RootTypes
$WorkloadInventory = Initialize-WARAWorkloadInventory -InScopeResources $JSONContent.resourceInventory -TenantID $JSONContent.scriptDetails.TenantId

Write-Host ("  ImpactedResources : {0}" -f @($ImpactedResources).Count)
Write-Host ("  PlatformIssues    : {0}" -f @($PlatformIssues).Count)
Write-Host ("  SupportRequests   : {0}" -f @($SupportTickets).Count)
Write-Host ("  AnalysisPlanning  : {0}" -f @($AnalysisPlanning).Count)
Write-Host ("  WorkloadInventory : {0}" -f @($WorkloadInventory).Count)

# Resolve the final output path (may live on OneDrive / a synced folder).
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path (Get-Location).Path ('Expert-Analysis-v1-' + (Get-Date -Format 'yyyy-MM-dd-HH-mm') + '.xlsx')
}

# Excel's Protected View blocks COM automation on files under OneDrive / synced or
# mark-of-the-web'd folders, so we always build the workbook in a plain local temp
# file and copy the finished result to $OutputPath at the very end.
$workPath = Join-Path ([System.IO.Path]::GetTempPath()) ('wara-com-work-' + [guid]::NewGuid().ToString('N') + '.xlsx')
Copy-Item -LiteralPath $TemplatePath -Destination $workPath -Force
Unblock-File -LiteralPath $workPath -ErrorAction SilentlyContinue

# Column orders per sheet (must match the ImportExcel output).
$colsImpacted = @('REQUIRED ACTIONS / REVIEW STATUS','ValidationCategory','Resource Type','subscriptionId','resourceGroup','location','name','id','custom1','custom2','custom3','custom4','custom5','Recommendation Title','Impact','Recommendation Control','Potential Benefit','Learn More Link','Long Description','Guid','Category','Source','WAF Pillar','Platform Issue TrackingId','Retirement TrackingId','Support Request Number','Notes','checkName')
$colsAnalysis = @('Category','Resource Type','Number of Resources','Impacted Resources','Has Recommendations_x000a_in APRL/Advisor','Assessment Owner','Assessment Status','Notes')
$colsPlatform = @('REQUIRED ACTIONS / REVIEW STATUS','Tracking ID','Event Type','Event Source','Status','Title','Level','Event Level','Start Time','Mitigation Time','Impacted Service','What happened','What went wrong and why','How did we respond','How are we making incidents like this less likely or less impactful','How can customers make incidents like this less impactful')
$colsSupport = @('REQUIRED ACTIONS / REVIEW STATUS','Ticket ID','Severity','Status','Support Plan Type','Creation Date','Modified Date','Title','Related Resource')
$colsInventory = @('id','name','type','tenantId','kind','location','resourceGroup','subscriptionId','managedBy','sku','plan','zones')

# ---------------------------------------------------------------------------------------------
# The Excel COM work runs inside this scriptblock. In normal (hidden) mode it executes in a
# short-lived background job -- its own PowerShell process -- which isolates the COM automation
# from the caller's session. The worker records the PID of the Excel instance it spawns and, as
# its final step, terminates ONLY that exact /automation process (a graceful Quit is attempted
# first but does not reliably close Excel on every machine). Because the target is identified by
# the specific spawned PID and re-confirmed to be an /automation instance, a user's manually
# opened Excel window is never touched.
# For -ShowExcel we run the same scriptblock in-process so the window stays open for inspection.
# ---------------------------------------------------------------------------------------------
$comWork = {
    param($P)
    $ErrorActionPreference = 'Stop'

    # Writes header row + data block for one sheet, creates the table, and sets any formula cells
    # (values that start with '=') AFTER the table exists so structured references resolve.
    function Write-WARASheet {
        param($Worksheet, [int]$StartRow, [string[]]$Columns, [object[]]$Rows, [string]$TableName, [string]$TableStyle = 'TableStyleLight19')

        $colCount = $Columns.Count
        $dataCount = @($Rows).Count
        if ($dataCount -lt 1) { $dataCount = 1 }   # always at least one (blank) data row so the table is valid

        # Clear any stale example content in the data region (values only; keep template formatting).
        $clearLast = $StartRow + [Math]::Max($dataCount, 3000)
        $Worksheet.Range($Worksheet.Cells.Item($StartRow, 1), $Worksheet.Cells.Item($clearLast, $colCount)).ClearContents() | Out-Null

        # Header row (convert ImportExcel line-break token to a real newline for display).
        for ($c = 0; $c -lt $colCount; $c++) {
            $Worksheet.Cells.Item($StartRow, $c + 1).Value2 = ($Columns[$c] -replace '_x000a_', "`n")
        }

        # Build a 2D value block; remember formula cells to set after the table exists.
        $formulaCells = New-Object System.Collections.Generic.List[object]
        $block = New-Object 'object[,]' $dataCount, $colCount
        for ($r = 0; $r -lt @($Rows).Count; $r++) {
            $row = $Rows[$r]
            for ($c = 0; $c -lt $colCount; $c++) {
                $val = $row.$($Columns[$c])
                if ($val -is [string] -and $val.StartsWith('=')) {
                    $formulaCells.Add([PSCustomObject]@{ R = $StartRow + 1 + $r; C = $c + 1; F = $val })
                    $block[$r, $c] = $null
                }
                else {
                    if ($null -ne $val -and $val -isnot [string]) { $val = [string]$val }
                    $block[$r, $c] = $val
                }
            }
        }
        if (@($Rows).Count -ge 1) {
            $topLeft = $Worksheet.Cells.Item($StartRow + 1, 1)
            $botRight = $Worksheet.Cells.Item($StartRow + @($Rows).Count, $colCount)
            # Use .Value (not .Value2): the Value2 property-put fails to marshal a 2D
            # object[,] array in PowerShell COM ("Unable to cast object[,] to String").
            $dataRange = $Worksheet.Range($topLeft, $botRight)
            $dataRange.Value = $block
        }

        # Create (or replace) the table over header + data rows.
        $lastRow = $StartRow + $dataCount
        $tblRange = $Worksheet.Range($Worksheet.Cells.Item($StartRow, 1), $Worksheet.Cells.Item($lastRow, $colCount))
        foreach ($existing in @($Worksheet.ListObjects)) {
            if ($existing.Name -eq $TableName) { $existing.Unlist() }
        }
        $lo = $Worksheet.ListObjects.Add(1, $tblRange, $null, 1)   # xlSrcRange, headers = xlYes
        $lo.Name = $TableName
        try { $lo.TableStyle = $TableStyle } catch { Write-Verbose "TableStyle '$TableStyle' not applied: $($_.Exception.Message)" }

        # Now set formula cells (table exists, so TableTypes8[[#This Row],...] structured refs resolve).
        foreach ($fc in $formulaCells) {
            try { $Worksheet.Cells.Item($fc.R, $fc.C).Formula2 = $fc.F }
            catch {
                try { $Worksheet.Cells.Item($fc.R, $fc.C).Formula = $fc.F }
                catch { Write-Verbose "Formula not set at R$($fc.R)C$($fc.C): $($_.Exception.Message)" }
            }
        }
    }

    function Set-WARAListValidation {
        param($Worksheet, [string]$ColumnLetter, [int]$FirstRow, [string[]]$Values)
        $rng = $Worksheet.Range("$ColumnLetter$FirstRow`:$ColumnLetter`1048576")
        $rng.Validation.Delete()
        $rng.Validation.Add(3, 1, 1, ($Values -join ',')) | Out-Null   # 3 = xlValidateList, 1 = xlValidAlertStop
    }

    function Set-WARALengthValidation {
        param($Worksheet, [string]$ColumnLetter, [int]$FirstRow)
        $rng = $Worksheet.Range("$ColumnLetter$FirstRow`:$ColumnLetter`1048576")
        $rng.Validation.Delete()
        $rng.Validation.Add(6, 1, 5, '1') | Out-Null   # 6 = xlValidateTextLength, 1 = xlValidAlertStop, 5 = xlGreater
    }

    # Record the exact PID of the Excel instance we are about to spawn so cleanup can target ONLY
    # that process. A COM-activated Excel launches as a new EXCEL.EXE; the PID that appears after
    # New-Object (and is absent before) is unambiguously ours.
    $excelPidsBefore = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $excel = New-Object -ComObject Excel.Application
    $ownExcelPid = @(Get-Process EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $_.Id -notin $excelPidsBefore } | Select-Object -ExpandProperty Id) | Select-Object -First 1
    $excel.Visible = [bool]$P.ShowExcel
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    # Suppress every interactive prompt (update-links, AutoRecover, overwrite, events) so a hidden
    # modal dialog can never block COM automation and hang the worker.
    try { $excel.AutomationSecurity = 1 } catch {}
    try { $excel.AskToUpdateLinks = $false } catch {}
    try { $excel.AlertBeforeOverwriting = $false } catch {}
    try { $excel.EnableEvents = $false } catch {}
    try { $excel.Interactive = $false } catch {}
    try { $excel.FeatureInstall = 0 } catch {}

    try {
        $wb = $excel.Workbooks.Open($P.WorkPath, 0, $false)

        $ws = $wb.Worksheets.Item($P.WorkloadInventorySheetRef)
        Write-WARASheet -Worksheet $ws -StartRow 12 -Columns $P.ColsInventory -Rows @($P.WorkloadInventory) -TableName 'InScopeResources'

        $ws = $wb.Worksheets.Item($P.ImpactedResourcesSheetRef)
        Write-WARASheet -Worksheet $ws -StartRow 12 -Columns $P.ColsImpacted -Rows @($P.ImpactedResources) -TableName 'impactedresources'
        Set-WARAListValidation   -Worksheet $ws -ColumnLetter 'A' -FirstRow 13 -Values @('Pending','Reviewed')
        Set-WARAListValidation   -Worksheet $ws -ColumnLetter 'O' -FirstRow 13 -Values @('High','Medium','Low')
        Set-WARALengthValidation -Worksheet $ws -ColumnLetter 'G' -FirstRow 13
        # Conditional formatting is intentionally NOT added here: the shipped template (which we copy
        # verbatim) already carries the 6 REQUIRED-ACTIONS colour rules on column A (A1:A1048576),
        # so they apply to the newly written rows automatically.

        $ws = $wb.Worksheets.Item($P.PlatformIssuesSheetRef)
        Write-WARASheet -Worksheet $ws -StartRow 12 -Columns $P.ColsPlatform -Rows @($P.PlatformIssues) -TableName 'platformIssues'

        $ws = $wb.Worksheets.Item($P.SupportRequestsSheetRef)
        Write-WARASheet -Worksheet $ws -StartRow 12 -Columns $P.ColsSupport -Rows @($P.SupportTickets) -TableName 'supportRequests'

        $ws = $wb.Worksheets.Item($P.AnalysisPlanningSheetRef)
        Write-WARASheet -Worksheet $ws -StartRow 10 -Columns $P.ColsAnalysis -Rows @($P.AnalysisPlanning) -TableName 'TableTypes8'

        $excel.Calculate()
        $wb.SaveAs($P.WorkPath, 51)   # 51 = xlOpenXMLWorkbook
        if (-not $P.ShowExcel) {
            $wb.Close($true)
            $excel.Quit()
        }
    }
    finally {
        if (-not $P.ShowExcel) {
            # First try to close Excel gracefully (release refs -> GC -> Quit already issued above).
            try { $excel.ScreenUpdating = $true } catch {}
            foreach ($ref in ,$ws + ,$wb) {
                if ($ref) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ref) } catch {} }
            }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
            $ws = $wb = $excel = $null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()

            # Surgical safety net: on this machine a graceful Quit does not always terminate the
            # automation instance. If OUR instance is still alive, terminate ONLY that exact PID --
            # and only after re-confirming it is still an EXCEL.EXE launched with /automation. A
            # user's manually-opened Excel is never an /automation instance and never has this PID,
            # so it can never be affected.
            if ($ownExcelPid) {
                $stillOurs = Get-CimInstance Win32_Process -Filter "ProcessId=$ownExcelPid" -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -eq 'EXCEL.EXE' -and $_.CommandLine -like '*/automation*' }
                if ($stillOurs) { Stop-Process -Id $ownExcelPid -Force -ErrorAction SilentlyContinue }
            }
        }
    }
}

$payload = @{
    WorkPath                  = $workPath
    ShowExcel                 = [bool]$ShowExcel
    WorkloadInventorySheetRef = $WorkloadInventorySheetRef
    ImpactedResourcesSheetRef = $ImpactedResourcesSheetRef
    PlatformIssuesSheetRef    = $PlatformIssuesSheetRef
    SupportRequestsSheetRef   = $SupportRequestsSheetRef
    AnalysisPlanningSheetRef  = $AnalysisPlanningSheetRef
    ColsInventory             = $colsInventory
    ColsImpacted              = $colsImpacted
    ColsPlatform              = $colsPlatform
    ColsSupport               = $colsSupport
    ColsAnalysis              = $colsAnalysis
    WorkloadInventory         = $WorkloadInventory
    ImpactedResources         = $ImpactedResources
    PlatformIssues            = $PlatformIssues
    SupportTickets            = $SupportTickets
    AnalysisPlanning          = $AnalysisPlanning
}

if ($ShowExcel) {
    Write-Host 'Writing workbook (Excel COM, visible window)...' -ForegroundColor Cyan
    & $comWork $payload
    Write-Host "Workbook is open in Excel (temporary file): $workPath" -ForegroundColor Green
    return $workPath
}

Write-Host 'Writing workbook in a short-lived Excel COM worker process...' -ForegroundColor Cyan
$job = Start-Job -ScriptBlock $comWork -ArgumentList $payload
$completed = Wait-Job -Job $job -Timeout 600
Receive-Job -Job $job    # surface worker output/errors (rethrows a terminating worker error here)
$workerState = $job.State
# The worker terminates its own /automation Excel before returning; removing the job just disposes it.
Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
if (-not $completed) { throw "The Excel COM worker did not finish within the timeout and was stopped." }
if ($workerState -eq 'Failed') { throw "The Excel COM worker process failed. See the errors above." }

# Copy the finished workbook from the local temp file to the requested destination.
Copy-Item -LiteralPath $workPath -Destination $OutputPath -Force
Remove-Item -LiteralPath $workPath -Force -ErrorAction SilentlyContinue
Write-Host "Saved: $OutputPath" -ForegroundColor Green

return $OutputPath
#endregion
